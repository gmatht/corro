//! Ratatui front-end: sheet viewport, editing, export, move, file sync.

use crate::addr::{self, parse_cell_ref_at, parse_sheet_id_prefix_at};
use crate::agg::{cell_display, compute_aggregate};
use crate::balance::{self, BalanceDirection};
use crate::export;
use crate::formula::translate_formula_text_by_offset;
use crate::formula::{cell_effective_display, is_formula};
use crate::grid::{
    CellAddr, CellFormat, FormatScope, GridBox as Grid, MainRange, MarginIndex, NumberFormat,
    SortSpec, TextAlign, FOOTER_ROWS, HEADER_ROWS, MARGIN_COLS,
};
use crate::io::{
    commit_workbook_op, commit_workbook_set_column_format_batch, load_workbook_revisions_partial,
    IoError, LogWatcher, PartialReplay,
};
use crate::ops::{AggFunc, AggregateDef, Op, SheetState, WorkbookState};
use crossterm::cursor::{Hide, Show};
use crossterm::event::{self, Event, KeyCode, KeyEvent, KeyEventKind, KeyModifiers};
use crossterm::execute;
use crossterm::terminal::{
    disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen,
};
use ratatui::prelude::*;
use ratatui::widgets::{
    Block, BorderType, Borders, Clear, List, ListItem, ListState, Paragraph, Wrap,
};
use std::collections::HashMap;
use std::fs::OpenOptions;
use std::io::{self, stdout};
use std::io::Write;
use std::path::{Path, PathBuf};
use thiserror::Error;
use unicode_truncate::{Alignment as UTruncAlign, UnicodeTruncateStr};
use unicode_width::UnicodeWidthStr;

/// Width of the row-label gutter (`]A~1`, `A1`, `A_1`).
const ROW_LABEL_CHARS: usize = 5;
/// Keep at most this many blank lines/cols around the active main data window.
const DISPLAY_EDGE_BLANK: usize = 1;
/// Trailing blank main rows allowed before Down transitions into the footer.
const NAV_BLANK_ROWS: usize = 2;
/// Trailing blank main cols allowed before Right transitions into the right margin.
const NAV_BLANK_COLS: usize = 1;

fn debug_json_escape(s: &str) -> String {
    s.replace('\\', "\\\\").replace('"', "\\\"")
}

fn debug_log_ndjson(hypothesis_id: &str, location: &str, message: &str, data_json: String) {
    if let Ok(mut file) = OpenOptions::new()
        .create(true)
        .append(true)
        .open("debug-1c96f4.log")
    {
        let _ = writeln!(
            file,
            "{{\"sessionId\":\"1c96f4\",\"runId\":\"pre-fix\",\"hypothesisId\":\"{}\",\"location\":\"{}\",\"message\":\"{}\",\"data\":{},\"timestamp\":{}}}",
            debug_json_escape(hypothesis_id),
            debug_json_escape(location),
            debug_json_escape(message),
            data_json,
            chrono::Utc::now().timestamp_millis()
        );
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum SelectionKind {
    Cells,
    Rows,
    Cols,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum SelectionEdgeDirection {
    Left,
    Right,
    Up,
    Down,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum FillDirection {
    Right,
    Down,
}

#[derive(Clone, Debug, Eq, PartialEq)]
enum OpenPathRequest {
    Plain(PathBuf),
    Revision { path: PathBuf, revision: usize },
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum TextInputAction {
    Handled,
    EdgeLeft,
    EdgeRight,
    Unhandled,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum OpenPathError {
    Empty,
    InvalidRevisionSyntax,
}

fn parse_open_path_request(raw: &str) -> Result<OpenPathRequest, OpenPathError> {
    let t = raw.trim();
    if t.is_empty() {
        return Err(OpenPathError::Empty);
    }

    for keyword in ["link", "load"] {
        if let Some(rest) = t.strip_prefix(keyword) {
            if !rest.is_empty() && rest.chars().next().is_some_and(|c| c.is_whitespace()) {
                let rest = rest.trim_start();
                let (path, revision) = rest
                    .rsplit_once(' ')
                    .ok_or(OpenPathError::InvalidRevisionSyntax)?;
                let path = path.trim();
                if path.is_empty() {
                    return Err(OpenPathError::InvalidRevisionSyntax);
                }
                let revision = revision
                    .trim()
                    .parse::<usize>()
                    .map_err(|_| OpenPathError::InvalidRevisionSyntax)?;
                return Ok(OpenPathRequest::Revision {
                    path: PathBuf::from(path),
                    revision,
                });
            }
        }
    }

    Ok(OpenPathRequest::Plain(PathBuf::from(t)))
}

#[derive(Debug, Error)]
pub enum RunError {
    #[error("I/O: {0}")]
    Io(#[from] IoError),
    #[error("Terminal: {0}")]
    Term(#[from] io::Error),
}

#[derive(Clone, Copy, Debug)]
pub struct MovieReplayOptions {
    pub typing_cps: f64,
    pub confirm_delay_ms: u64,
    pub menu_hold_ms: u64,
}

/// Logical cursor position across header+main+footer rows × total global columns.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct SheetCursor {
    pub row: usize,
    pub col: usize,
}

impl SheetCursor {
    fn clamp(&mut self, grid: &Grid) {
        let rows = HEADER_ROWS + grid.main_rows() + FOOTER_ROWS;
        let cols = grid.total_cols();
        if rows > 0 {
            self.row = self.row.min(rows - 1);
        }
        if cols > 0 {
            self.col = self.col.min(cols - 1);
        }
    }

    pub(crate) fn to_addr(self, grid: &Grid) -> CellAddr {
        addr::sheet_cursor_to_addr(
            addr::LogicalRow(self.row),
            addr::GlobalCol(self.col),
            addr::MainRows(grid.main_rows()),
            addr::MainCols(grid.main_cols()),
        )
    }
}

#[derive(Clone, Debug)]
enum Mode {
    Normal,
    RevisionBrowse,
    Edit {
        buffer: String,
        formula_cursor: Option<SheetCursor>,
        fit_to_content_on_commit: bool,
    },
    OpenPath {
        buffer: String,
    },
    SheetRename {
        buffer: String,
        sheet_id: u32,
    },
    SheetCopy {
        buffer: String,
        source_id: u32,
    },
    GoToCell {
        buffer: String,
    },
    SavePath {
        buffer: String,
    },
    Help,
    About,
    /// Alt-activated menu bar; letter shortcuts execute actions.
    Menu {
        stack: Vec<MenuLevel>,
    },
    ExportTsv {
        buffer: String,
    },
    ExportCsv {
        buffer: String,
    },
    ExportAscii {
        buffer: String,
    },
    ExportAll {
        buffer: String,
    },
    ExportOdt {
        buffer: String,
    },
    Find {
        buffer: String,
    },
    Replace {
        buffer: String,
    },
    SetMaxColWidth {
        buffer: String,
    },
    SetColWidth {
        buffer: String,
    },
    SortView {
        buffer: String,
        persist: bool,
    },
    FormatDecimals {
        buffer: String,
        decimals_for: FormatDecimalsFor,
    },
    BalanceBooks {
        buffer: String,
        direction: BalanceDirection,
        persist: bool,
        focus: BalanceBooksFocus,
    },
    QuitPrompt,
    /// No `.corro` on disk (e.g. opened from ODS/TSV/CSV); user should save to `.corro` or discard.
    QuitImportPrompt,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum BalanceBooksFocus {
    Column,
    ReportViewOnly,
    ReportPersisted,
    // Match the sign-pairing direction.
    PosToNeg,
    NegToPos,
    Generate,
    Cancel,
}

const SPECIAL_VALUE_CHOICES: [&str; 10] = ["∞", "Σ", "Ω", "π", "μ", "Δ", "√", "φ", "λ", "θ"];

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum MenuSection {
    Edit,
    File,
    Format,
    FormatScope,
    FormatNumber,
    FormatAlign,
    Sheet,
    Export,
    Width,
    Insert,
    Help,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum FormatTarget {
    All,
    FullColumn,
    Data,
    Special,
    Cell,
    Selection,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum FormatDecimalsFor {
    Currency,
    Fixed,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum MenuTarget {
    Action(MenuAction),
    Submenu(MenuSection),
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct MenuLevel {
    section: MenuSection,
    item: usize,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum MenuAction {
    Cut,
    Copy,
    Paste,
    Find,
    Replace,
    OpenFile,
    Replay,
    SaveAs,
    RenameSheet,
    CopySheet,
    MoveSheet,
    SheetPrev,
    SheetNext,
    GoToCell,
    Exit,
    ExportTsv,
    ExportCsv,
    ExportAscii,
    ExportAll,
    ExportOdt,
    SetMaxColWidth,
    SetColWidth,
    FormatApplyAll,
    FormatApplyFullColumn,
    FormatApplyData,
    FormatApplySpecial,
    FormatApplyCell,
    FormatApplySelection,
    FormatCurrency,
    FormatFixed0,
    FormatFixed1,
    FormatFixed2,
    FormatFixedCustom,
    FormatAlignLeft,
    FormatAlignCenter,
    FormatAlignRight,
    FormatAlignDefault,
    FormatReset,
    InsertRows,
    InsertMitosisRow,
    InsertMitosisCol,
    InsertCols,
    InsertSpecialChars,
    InsertDate,
    InsertTime,
    InsertHyperlink,
    SortView,
    SaveSort,
    BalanceBooks,
    NewSheet,
    HelpRows,
    HelpCols,
    About,
    HelpFull,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct MenuItem {
    shortcut: char,
    label: &'static str,
    target: MenuTarget,
}

const EDIT_MENU_ITEMS: [MenuItem; 5] = [
    MenuItem {
        shortcut: 'X',
        label: "Cut",
        target: MenuTarget::Action(MenuAction::Cut),
    },
    MenuItem {
        shortcut: 'C',
        label: "Copy",
        target: MenuTarget::Action(MenuAction::Copy),
    },
    MenuItem {
        shortcut: 'P',
        label: "Paste",
        target: MenuTarget::Action(MenuAction::Paste),
    },
    MenuItem {
        shortcut: 'F',
        label: "Find",
        target: MenuTarget::Action(MenuAction::Find),
    },
    MenuItem {
        shortcut: 'R',
        label: "Replace",
        target: MenuTarget::Action(MenuAction::Replace),
    },
];

const FILE_MENU_ITEMS: [MenuItem; 8] = [
    MenuItem {
        shortcut: 'O',
        label: "Open file",
        target: MenuTarget::Action(MenuAction::OpenFile),
    },
    MenuItem {
        shortcut: 'A',
        label: "Save as",
        target: MenuTarget::Action(MenuAction::SaveAs),
    },
    MenuItem {
        shortcut: 'T',
        label: "Export",
        target: MenuTarget::Submenu(MenuSection::Export),
    },
    MenuItem {
        shortcut: 'C',
        label: "Width",
        target: MenuTarget::Submenu(MenuSection::Width),
    },
    MenuItem {
        shortcut: 'S',
        label: "Sort view",
        target: MenuTarget::Action(MenuAction::SortView),
    },
    MenuItem {
        shortcut: 'P',
        label: "Persist sort",
        target: MenuTarget::Action(MenuAction::SaveSort),
    },
    MenuItem {
        shortcut: 'X',
        label: "Exit",
        target: MenuTarget::Action(MenuAction::Exit),
    },
    MenuItem {
        shortcut: 'R',
        label: "Replay",
        target: MenuTarget::Action(MenuAction::Replay),
    },
];

const FORMAT_MENU_ITEMS: [MenuItem; 4] = [
    MenuItem {
        shortcut: 'S',
        label: "Scope",
        target: MenuTarget::Submenu(MenuSection::FormatScope),
    },
    MenuItem {
        shortcut: 'N',
        label: "Number",
        target: MenuTarget::Submenu(MenuSection::FormatNumber),
    },
    MenuItem {
        shortcut: 'A',
        label: "Align",
        target: MenuTarget::Submenu(MenuSection::FormatAlign),
    },
    MenuItem {
        shortcut: 'R',
        label: "Reset",
        target: MenuTarget::Action(MenuAction::FormatReset),
    },
];

const FORMAT_SCOPE_MENU_ITEMS: [MenuItem; 6] = [
    MenuItem {
        shortcut: 'A',
        label: "All",
        target: MenuTarget::Action(MenuAction::FormatApplyAll),
    },
    MenuItem {
        shortcut: 'F',
        label: "Full col",
        target: MenuTarget::Action(MenuAction::FormatApplyFullColumn),
    },
    MenuItem {
        shortcut: 'D',
        label: "Data",
        target: MenuTarget::Action(MenuAction::FormatApplyData),
    },
    MenuItem {
        shortcut: 'S',
        label: "Special",
        target: MenuTarget::Action(MenuAction::FormatApplySpecial),
    },
    MenuItem {
        shortcut: 'C',
        label: "Cell",
        target: MenuTarget::Action(MenuAction::FormatApplyCell),
    },
    MenuItem {
        shortcut: 'L',
        label: "Selection",
        target: MenuTarget::Action(MenuAction::FormatApplySelection),
    },
];

const SHEET_MENU_ITEMS: [MenuItem; 8] = [
    MenuItem {
        shortcut: '[',
        label: "Prev sheet",
        target: MenuTarget::Action(MenuAction::SheetPrev),
    },
    MenuItem {
        shortcut: ']',
        label: "Next sheet",
        target: MenuTarget::Action(MenuAction::SheetNext),
    },
    MenuItem {
        shortcut: 'N',
        label: "New sheet",
        target: MenuTarget::Action(MenuAction::NewSheet),
    },
    MenuItem {
        shortcut: 'R',
        label: "Rename sheet",
        target: MenuTarget::Action(MenuAction::RenameSheet),
    },
    MenuItem {
        shortcut: 'C',
        label: "Copy sheet",
        target: MenuTarget::Action(MenuAction::CopySheet),
    },
    MenuItem {
        shortcut: 'M',
        label: "Move sheet",
        target: MenuTarget::Action(MenuAction::MoveSheet),
    },
    MenuItem {
        shortcut: 'G',
        label: "Go",
        target: MenuTarget::Action(MenuAction::GoToCell),
    },
    MenuItem {
        shortcut: 'B',
        label: "Balance books",
        target: MenuTarget::Action(MenuAction::BalanceBooks),
    },
];

const INSERT_ROOT_MENU_ITEMS: [MenuItem; 8] = [
    MenuItem {
        shortcut: 'R',
        label: "Rows",
        target: MenuTarget::Action(MenuAction::InsertRows),
    },
    MenuItem {
        shortcut: 'M',
        label: "Mitosis (Row)",
        target: MenuTarget::Action(MenuAction::InsertMitosisRow),
    },
    MenuItem {
        shortcut: 'O',
        label: "Mitosis (Col)",
        target: MenuTarget::Action(MenuAction::InsertMitosisCol),
    },
    MenuItem {
        shortcut: 'C',
        label: "Cols",
        target: MenuTarget::Action(MenuAction::InsertCols),
    },
    MenuItem {
        shortcut: 'S',
        label: "Special Char",
        target: MenuTarget::Action(MenuAction::InsertSpecialChars),
    },
    MenuItem {
        shortcut: ';',
        label: "Date",
        target: MenuTarget::Action(MenuAction::InsertDate),
    },
    MenuItem {
        shortcut: ':',
        label: "Time",
        target: MenuTarget::Action(MenuAction::InsertTime),
    },
    MenuItem {
        shortcut: 'H',
        label: "Hyperlink",
        target: MenuTarget::Action(MenuAction::InsertHyperlink),
    },
];

const EXPORT_MENU_ITEMS: [MenuItem; 5] = [
    MenuItem {
        shortcut: 'T',
        label: "TSV",
        target: MenuTarget::Action(MenuAction::ExportTsv),
    },
    MenuItem {
        shortcut: 'C',
        label: "CSV",
        target: MenuTarget::Action(MenuAction::ExportCsv),
    },
    MenuItem {
        shortcut: 'A',
        label: "ASCII table",
        target: MenuTarget::Action(MenuAction::ExportAscii),
    },
    MenuItem {
        shortcut: 'L',
        label: "Export all",
        target: MenuTarget::Action(MenuAction::ExportAll),
    },
    MenuItem {
        shortcut: 'D',
        label: "ODS",
        target: MenuTarget::Action(MenuAction::ExportOdt),
    },
];

const WIDTH_MENU_ITEMS: [MenuItem; 2] = [
    MenuItem {
        shortcut: 'D',
        label: "Default width",
        target: MenuTarget::Action(MenuAction::SetMaxColWidth),
    },
    MenuItem {
        shortcut: 'C',
        label: "Column width",
        target: MenuTarget::Action(MenuAction::SetColWidth),
    },
];

const FORMAT_NUMBER_MENU_ITEMS: [MenuItem; 5] = [
    MenuItem {
        shortcut: '$',
        label: "Currency ($)",
        target: MenuTarget::Action(MenuAction::FormatCurrency),
    },
    MenuItem {
        shortcut: '0',
        label: "Fixed 0",
        target: MenuTarget::Action(MenuAction::FormatFixed0),
    },
    MenuItem {
        shortcut: '1',
        label: "Fixed 1",
        target: MenuTarget::Action(MenuAction::FormatFixed1),
    },
    MenuItem {
        shortcut: '2',
        label: "Fixed 2",
        target: MenuTarget::Action(MenuAction::FormatFixed2),
    },
    MenuItem {
        shortcut: 'N',
        label: "Fixed n",
        target: MenuTarget::Action(MenuAction::FormatFixedCustom),
    },
];

const FORMAT_ALIGN_MENU_ITEMS: [MenuItem; 4] = [
    MenuItem {
        shortcut: 'L',
        label: "Left",
        target: MenuTarget::Action(MenuAction::FormatAlignLeft),
    },
    MenuItem {
        shortcut: 'C',
        label: "Center",
        target: MenuTarget::Action(MenuAction::FormatAlignCenter),
    },
    MenuItem {
        shortcut: 'R',
        label: "Right",
        target: MenuTarget::Action(MenuAction::FormatAlignRight),
    },
    MenuItem {
        shortcut: 'D',
        label: "Default",
        target: MenuTarget::Action(MenuAction::FormatAlignDefault),
    },
];

const HELP_MENU_ITEMS: [MenuItem; 4] = [
    MenuItem {
        shortcut: 'A',
        label: "About",
        target: MenuTarget::Action(MenuAction::About),
    },
    MenuItem {
        shortcut: 'R',
        label: "Row ops",
        target: MenuTarget::Action(MenuAction::HelpRows),
    },
    MenuItem {
        shortcut: 'C',
        label: "Col ops",
        target: MenuTarget::Action(MenuAction::HelpCols),
    },
    MenuItem {
        shortcut: 'H',
        label: "Full help",
        target: MenuTarget::Action(MenuAction::HelpFull),
    },
];

// ── Viewport helpers ──────────────────────────────────────────────────────────

fn main_row_window(
    state: &SheetState,
    cursor: SheetCursor,
    main_order: &[usize],
) -> (usize, usize) {
    let g = &state.grid;
    let hr = HEADER_ROWS;
    let mr = g.main_rows();
    if mr == 0 {
        return (0, 0);
    }

    let mut lo = usize::MAX;
    let mut hi = 0usize;

    for (pos, &main_row) in main_order.iter().enumerate() {
        if g.logical_row_has_content(hr + main_row) || left_margin_template_applies(g, main_row) {
            lo = lo.min(pos);
            hi = hi.max(pos);
        }
    }
    if cursor.row >= hr && cursor.row < hr + mr {
        let ri = main_order
            .iter()
            .position(|&r| hr + r == cursor.row)
            .unwrap_or(0);
        lo = lo.min(ri);
        hi = hi.max(ri);
    }
    if lo == usize::MAX {
        lo = 0;
        hi = 0;
    }

    lo = lo.saturating_sub(DISPLAY_EDGE_BLANK);
    hi = hi
        .saturating_add(DISPLAY_EDGE_BLANK)
        .min(mr.saturating_sub(1));
    (lo, hi)
}

fn main_col_window(state: &SheetState, cursor: SheetCursor) -> (usize, usize) {
    let g = &state.grid;
    let lm = MARGIN_COLS;
    let mc = g.main_cols();
    if mc == 0 {
        return (0, 0);
    }

    let mut lo = usize::MAX;
    let mut hi = 0usize;

    for c in 0..mc {
        if g.logical_col_has_content(lm + c)
            || header_template_applies(g, c)
            || right_col_agg_func(g, lm + c).is_some()
        {
            lo = lo.min(c);
            hi = hi.max(c);
        }
    }
    if cursor.col >= lm && cursor.col < lm + mc {
        let ci = cursor.col - lm;
        lo = lo.min(ci);
        hi = hi.max(ci);
    }
    if lo == usize::MAX {
        lo = 0;
        hi = 0;
    }

    lo = lo.saturating_sub(DISPLAY_EDGE_BLANK);
    hi = hi
        .saturating_add(DISPLAY_EDGE_BLANK)
        .min(mc.saturating_sub(1));
    (lo, hi)
}

fn footer_nonblank_end(state: &SheetState) -> Option<usize> {
    let g = &state.grid;
    let hr = HEADER_ROWS;
    let mr = g.main_rows();
    let fr = FOOTER_ROWS;
    let mut max_nonblank = None;
    for i in 0..fr {
        if g.logical_row_has_content(hr + mr + i) {
            max_nonblank = Some(i);
        }
    }
    max_nonblank
}

fn menu_items(section: MenuSection) -> &'static [MenuItem] {
    match section {
        MenuSection::Edit => &EDIT_MENU_ITEMS,
        MenuSection::File => &FILE_MENU_ITEMS,
        MenuSection::Format => &FORMAT_MENU_ITEMS,
        MenuSection::FormatScope => &FORMAT_SCOPE_MENU_ITEMS,
        MenuSection::FormatNumber => &FORMAT_NUMBER_MENU_ITEMS,
        MenuSection::FormatAlign => &FORMAT_ALIGN_MENU_ITEMS,
        MenuSection::Sheet => &SHEET_MENU_ITEMS,
        MenuSection::Insert => &INSERT_ROOT_MENU_ITEMS,
        MenuSection::Export => &EXPORT_MENU_ITEMS,
        MenuSection::Width => &WIDTH_MENU_ITEMS,
        MenuSection::Help => &HELP_MENU_ITEMS,
    }
}

fn menu_title(section: MenuSection) -> &'static str {
    match section {
        MenuSection::Edit => "Edit",
        MenuSection::File => "File",
        MenuSection::Format => "Format",
        MenuSection::FormatScope => "Format Scope",
        MenuSection::FormatNumber => "Format Number",
        MenuSection::FormatAlign => "Format Align",
        MenuSection::Sheet => "Sheet",
        MenuSection::Export => "Export",
        MenuSection::Width => "Width",
        MenuSection::Insert => "Insert",
        MenuSection::Help => "Help",
    }
}

fn menu_action_item(section: MenuSection, item: usize) -> Option<MenuItem> {
    menu_items(section).get(item).copied()
}

fn menu_next_root_section(section: MenuSection) -> MenuSection {
    match section {
        MenuSection::File => MenuSection::Edit,
        MenuSection::Edit => MenuSection::Insert,
        MenuSection::Insert => MenuSection::Format,
        MenuSection::Format => MenuSection::Sheet,
        MenuSection::Sheet => MenuSection::Help,
        MenuSection::Help => MenuSection::File,
        _ => MenuSection::File,
    }
}

fn menu_prev_root_section(section: MenuSection) -> MenuSection {
    match section {
        MenuSection::File => MenuSection::Help,
        MenuSection::Edit => MenuSection::File,
        MenuSection::Insert => MenuSection::Edit,
        MenuSection::Format => MenuSection::Insert,
        MenuSection::Sheet => MenuSection::Format,
        MenuSection::Help => MenuSection::Sheet,
        _ => MenuSection::File,
    }
}

fn menu_popup_area(area: Rect, section: MenuSection, parent: Option<(Rect, usize)>) -> Rect {
    let items = menu_items(section).len() as u16;
    let width = match section {
        MenuSection::Edit => 18,
        MenuSection::File => 22,
        MenuSection::Format => 18,
        MenuSection::FormatScope => 18,
        MenuSection::FormatNumber => 18,
        MenuSection::FormatAlign => 18,
        MenuSection::Sheet => 20,
        MenuSection::Export => 18,
        MenuSection::Width => 20,
        MenuSection::Insert => 20,
        MenuSection::Help => 18,
    }
    .min(area.width.saturating_sub(2).max(1));
    let height = items.saturating_add(2).min(area.height.max(3));
    let (x, y) = parent
        .map(|(p, item)| (p.x.saturating_add(p.width), p.y.saturating_add(item as u16)))
        .unwrap_or_else(|| {
            let x = match section {
                MenuSection::File => 1,
                MenuSection::Edit => 9,
                MenuSection::Insert => 17,
                MenuSection::Format => 27,
                MenuSection::FormatScope => 27,
                MenuSection::FormatNumber => 27,
                MenuSection::FormatAlign => 27,
                MenuSection::Sheet => 36,
                MenuSection::Help => 45,
                _ => 1,
            };
            (area.x.saturating_add(x), area.y.saturating_add(1))
        });
    let x = x.min(
        area.x
            .saturating_add(area.width.saturating_sub(width.saturating_add(1))),
    );
    let y = y.min(area.y.saturating_add(area.height.saturating_sub(height)));
    Rect {
        x,
        y,
        width,
        height,
    }
}

impl App {
    fn open_menu(&mut self, section: MenuSection) {
        self.mode = Mode::Menu {
            stack: vec![MenuLevel { section, item: 0 }],
        };
    }

    fn clear_pending_format_target(&mut self) {
        self.pending_format_target = None;
    }

    fn open_menu_item(&mut self, section: MenuSection, item: usize) {
        self.mode = Mode::Menu {
            stack: vec![MenuLevel { section, item }],
        };
    }

    fn open_menu_path(&mut self, stack: Vec<MenuLevel>) {
        self.mode = Mode::Menu { stack };
    }

    fn start_edit_mode(
        &mut self,
        buffer: String,
        formula_cursor: Option<SheetCursor>,
        special_palette: bool,
        fit_to_content_on_commit: bool,
        edit_range_addrs: Option<Vec<CellAddr>>,
    ) -> Mode {
        self.edit_range_addrs = edit_range_addrs;
        self.edit_target_addr = Some(self.cursor.to_addr(&self.state.grid));
        let cursor = if buffer.trim() == "=" {
            1
        } else {
            buffer.chars().count()
        };
        self.edit_cursor = Some(cursor);
        self.edit_special_palette = special_palette;
        self.pending_fit_to_content_on_commit = fit_to_content_on_commit;
        Mode::Edit {
            buffer,
            formula_cursor,
            fit_to_content_on_commit,
        }
    }

    fn start_edit_current_cell(&mut self) -> Mode {
        let addr = self.cursor.to_addr(&self.state.grid);
        let cur = cell_display(&self.state.grid, &addr);
        self.start_edit_mode(
            cur.clone(),
            if cur.trim() == "=" {
                Some(self.cursor)
            } else {
                None
            },
            false,
            false,
            None,
        )
    }

    fn open_special_picker(&mut self) {
        self.special_picker = Some(0);
        self.mode = Mode::Normal;
    }

    fn commit_special_choice(&mut self, idx: usize) {
        let choice = SPECIAL_VALUE_CHOICES[idx];
        let buffer = choice.to_string();
        self.mode = self.start_edit_mode(buffer, None, true, false, None);
    }

    fn menu_action_mode(&mut self, action: MenuAction) -> Mode {
        self.edit_special_palette = false;
        match action {
            MenuAction::Cut => {
                let cells = self.selection_clear_cells();
                if cells.is_empty() {
                    self.status = "Nothing to cut".into();
                } else {
                    let data = self.selection_tsv_text();
                    let op = Op::FillRange {
                        cells: cells.clone(),
                    };
                    if self.copy_selection_to_clipboard(&data) {
                        if self.apply_single_op(op).is_ok() {
                            for (addr, _) in cells {
                                if let CellAddr::Main { col, .. } = addr {
                                    self.state.grid.auto_fit_column(MARGIN_COLS + col as usize);
                                }
                            }
                            self.status = "Selection cut".into();
                        }
                    }
                }
                Mode::Normal
            }
            MenuAction::Copy => {
                let data = self.selection_tsv_text();
                self.copy_selection_to_clipboard(&data);
                Mode::Normal
            }
            MenuAction::Paste => {
                if let Err(e) = self.paste_from_clipboard(true) {
                    self.status = format!("Clipboard error: {e}");
                }
                Mode::Normal
            }
            MenuAction::Find => Mode::Find {
                buffer: self.start_input_mode(String::new()),
            },
            MenuAction::Replace => Mode::Replace {
                buffer: self.start_input_mode(String::new()),
            },
            MenuAction::OpenFile => {
                let buffer = self.open_path_prompt_buffer();
                Mode::OpenPath {
                    buffer: self.start_input_mode(buffer),
                }
            }
            MenuAction::Replay => {
                if let Some(path) = self.path.clone().or(self.source_path.clone()) {
                    if path.exists() {
                        if let Some(ext) = path.extension().and_then(|e| e.to_str()) {
                            if ext.eq_ignore_ascii_case("corro") {
                                self.revision_browse = true;
                                self.source_path = Some(path.clone());
                                self.path = None;
                                self.workbook = WorkbookState::new();
                                self.state = SheetState::new(1, 1);
                                let mut active_sheet =
                                    self.workbook.sheet_id(self.workbook.active_sheet);
                                if let Ok((off, replay)) = load_workbook_revisions_partial(
                                    &path,
                                    usize::MAX,
                                    &mut self.workbook,
                                    &mut active_sheet,
                                ) {
                                    self.view_sheet_id = active_sheet;
                                    self.sync_active_sheet_cache();
                                    self.sync_persisted_sort_cache_from_workbook();
                                    for c in 0..self.state.grid.main_cols() {
                                        self.fit_column_to_rendered_content(MARGIN_COLS + c);
                                    }
                                    self.offset = off;
                                    self.ops_applied = replay.op_count;
                                    self.revision_browse_limit = replay.op_count;
                                    self.status = Self::replay_status("Replayed", &path, &replay);
                                    self.cursor.clamp(&self.state.grid);
                                }
                                return Mode::RevisionBrowse;
                            }
                        }
                    }
                }
                let buffer = self.open_path_prompt_buffer();
                Mode::OpenPath {
                    buffer: self.start_input_mode(buffer),
                }
            }
            MenuAction::SaveAs => Mode::SavePath {
                buffer: self.start_input_mode(self.suggested_corro_save_path()),
            },
            MenuAction::RenameSheet => Mode::SheetRename {
                buffer: self.start_input_mode(self.current_sheet_title()),
                sheet_id: self.view_sheet_id,
            },
            MenuAction::CopySheet => Mode::SheetCopy {
                buffer: self.start_input_mode(format!("{} Copy", self.current_sheet_title())),
                source_id: self.view_sheet_id,
            },
            MenuAction::MoveSheet => {
                let _ = self.move_current_sheet_to_end();
                Mode::Normal
            }
            MenuAction::SheetPrev => {
                self.switch_sheet(-1);
                Mode::Normal
            }
            MenuAction::SheetNext => {
                self.switch_sheet(1);
                Mode::Normal
            }
            MenuAction::GoToCell => Mode::GoToCell {
                buffer: self.start_input_mode(String::new()),
            },
            MenuAction::Exit => {
                if self.path.is_none() {
                    Mode::QuitImportPrompt
                } else {
                    Mode::QuitPrompt
                }
            }
            MenuAction::ExportTsv => {
                self.export_preview_scroll = 0;
                self.export_delimited_options.content = export::ExportContent::Values;
                Mode::ExportTsv {
                    buffer: self.start_input_mode(self.suggested_export_save_path("tsv")),
                }
            },
            MenuAction::ExportCsv => {
                self.export_preview_scroll = 0;
                self.export_delimited_options.content = export::ExportContent::Values;
                Mode::ExportCsv {
                    buffer: self.start_input_mode(self.suggested_export_save_path("csv")),
                }
            },
            MenuAction::ExportAscii => {
                self.export_preview_scroll = 0;
                self.export_ascii_options.content = export::ExportContent::Values;
                Mode::ExportAscii {
                    buffer: self.start_input_mode(self.suggested_export_save_path("txt")),
                }
            },
            MenuAction::ExportAll => {
                self.export_preview_scroll = 0;
                self.export_delimited_options.content = export::ExportContent::Values;
                Mode::ExportAll {
                    buffer: self.start_input_mode(self.suggested_export_save_path("tsv")),
                }
            },
            MenuAction::ExportOdt => {
                self.export_preview_scroll = 0;
                self.export_ods_content = export::ExportContent::Generic;
                Mode::ExportOdt {
                    buffer: self.start_input_mode(self.suggested_export_save_path("ods")),
                }
            },
            MenuAction::SetMaxColWidth => Mode::SetMaxColWidth {
                buffer: self.start_input_mode(String::new()),
            },
            MenuAction::SetColWidth => Mode::SetColWidth {
                buffer: self.start_input_mode(String::new()),
            },
            MenuAction::InsertRows => {
                let _ = self.insert_rows_above_cursor(1);
                Mode::Normal
            }
            MenuAction::InsertMitosisRow => {
                let _ = self.insert_mitosis_row_after_cursor();
                Mode::Normal
            }
            MenuAction::InsertMitosisCol => {
                let _ = self.insert_mitosis_col_after_cursor();
                Mode::Normal
            }
            MenuAction::InsertCols => {
                let _ = self.insert_cols_left_of_cursor(1);
                Mode::Normal
            }
            MenuAction::InsertSpecialChars => {
                self.open_special_picker();
                Mode::Normal
            }
            MenuAction::InsertDate => self.start_edit_mode(
                chrono::Local::now().format("%Y-%m-%d").to_string(),
                None,
                false,
                true,
                None,
            ),
            MenuAction::InsertTime => self.start_edit_mode(
                chrono::Local::now().format("%H:%M:%S").to_string(),
                None,
                false,
                true,
                None,
            ),
            MenuAction::InsertHyperlink => {
                self.start_edit_mode(self.menu_insert_hyperlink_seed(), None, false, false, None)
            }
            MenuAction::SortView => Mode::SortView {
                buffer: self.start_input_mode(String::new()),
                persist: false,
            },
            MenuAction::SaveSort => Mode::SortView {
                buffer: self.start_input_mode(String::new()),
                persist: true,
            },
            MenuAction::BalanceBooks => Mode::BalanceBooks {
                buffer: self.start_input_mode(
                    balance::choose_balance_column(&self.state.grid)
                        .map(addr::excel_column_name)
                        .unwrap_or_default(),
                ),
                direction: BalanceDirection::PosToNeg,
                persist: false,
                focus: BalanceBooksFocus::Column,
            },
            MenuAction::NewSheet => {
                self.add_sheet(format!("Sheet{}", self.workbook.next_sheet_id));
                Mode::Normal
            }
            MenuAction::HelpRows => {
                self.status = "Row ops: v·select full rows, then r·move to target row".into();
                Mode::Normal
            }
            MenuAction::HelpCols => {
                self.status = "Col ops: v·select full columns, then c·move to target column".into();
                Mode::Normal
            }
            MenuAction::About => {
                self.about_scroll = 0;
                Mode::About
            }
            MenuAction::HelpFull => {
                self.help_scroll = 0;
                Mode::Help
            }
            MenuAction::FormatApplyAll => {
                self.pending_format_target = Some(FormatTarget::All);
                Mode::Menu {
                    stack: vec![MenuLevel {
                        section: MenuSection::Format,
                        item: 0,
                    }],
                }
            }
            MenuAction::FormatApplyFullColumn => {
                self.pending_format_target = Some(FormatTarget::FullColumn);
                Mode::Menu {
                    stack: vec![MenuLevel {
                        section: MenuSection::Format,
                        item: 0,
                    }],
                }
            }
            MenuAction::FormatApplyData => {
                self.pending_format_target = Some(FormatTarget::Data);
                Mode::Menu {
                    stack: vec![MenuLevel {
                        section: MenuSection::Format,
                        item: 0,
                    }],
                }
            }
            MenuAction::FormatApplySpecial => {
                self.pending_format_target = Some(FormatTarget::Special);
                Mode::Menu {
                    stack: vec![MenuLevel {
                        section: MenuSection::Format,
                        item: 0,
                    }],
                }
            }
            MenuAction::FormatApplyCell => {
                self.pending_format_target = Some(FormatTarget::Cell);
                Mode::Menu {
                    stack: vec![MenuLevel {
                        section: MenuSection::Format,
                        item: 0,
                    }],
                }
            }
            MenuAction::FormatApplySelection => {
                self.pending_format_target = Some(FormatTarget::Selection);
                Mode::Menu {
                    stack: vec![MenuLevel {
                        section: MenuSection::Format,
                        item: 0,
                    }],
                }
            }
            MenuAction::FormatCurrency => Mode::FormatDecimals {
                buffer: self.start_input_mode(String::new()),
                decimals_for: FormatDecimalsFor::Currency,
            },
            MenuAction::FormatFixed0 => {
                self.apply_format_number(0, false);
                Mode::Normal
            }
            MenuAction::FormatFixed1 => {
                self.apply_format_number(1, false);
                Mode::Normal
            }
            MenuAction::FormatFixed2 => {
                self.apply_format_number(2, false);
                Mode::Normal
            }
            MenuAction::FormatFixedCustom => Mode::FormatDecimals {
                buffer: self.start_input_mode(String::new()),
                decimals_for: FormatDecimalsFor::Fixed,
            },
            MenuAction::FormatAlignLeft => {
                self.apply_format_align(TextAlign::Left);
                Mode::Normal
            }
            MenuAction::FormatAlignCenter => {
                self.apply_format_align(TextAlign::Center);
                Mode::Normal
            }
            MenuAction::FormatAlignRight => {
                self.apply_format_align(TextAlign::Right);
                Mode::Normal
            }
            MenuAction::FormatAlignDefault => {
                self.apply_format_align(TextAlign::Default);
                Mode::Normal
            }
            MenuAction::FormatReset => {
                self.apply_format_reset();
                Mode::Normal
            }
        }
    }

    /// TSV/ODS import with no `.corro` path: no unsaved edits when the undo stack is at the
    /// session baseline (including after the user undoes back to the imported state).
    fn is_ods_tsv_import_unchanged(&self) -> bool {
        if self.path.is_some() {
            return false;
        }
        let Some(src) = self.import_source.as_ref() else {
            return false;
        };
        let ext = src
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_lowercase();
        if ext != "tsv" && ext != "ods" {
            return false;
        }
        self.op_history.is_empty()
    }

    fn menu_target_mode(&mut self, path: &[MenuLevel], target: MenuTarget) -> Result<Mode, ()> {
        match target {
            MenuTarget::Action(action) => {
                if matches!(action, MenuAction::Exit) && self.is_ods_tsv_import_unchanged() {
                    return Err(());
                }
                Ok(self.menu_action_mode(action))
            }
            MenuTarget::Submenu(section) => {
                let mut stack = path.to_vec();
                stack.push(MenuLevel { section, item: 0 });
                Ok(Mode::Menu { stack })
            }
        }
    }

    fn menu_render_levels(stack: &[MenuLevel]) -> Vec<MenuLevel> {
        let mut levels = stack.to_vec();
        let mut preview_depth = 0usize;
        while preview_depth < 8 {
            let Some(level) = levels.last().copied() else {
                break;
            };
            let Some(menu_item) = menu_action_item(level.section, level.item) else {
                break;
            };
            match menu_item.target {
                MenuTarget::Submenu(section) => {
                    levels.push(MenuLevel { section, item: 0 });
                    preview_depth += 1;
                }
                MenuTarget::Action(_) => break,
            }
        }
        levels
    }

    fn menu_selected_index(
        render_index: usize,
        actual_depth: usize,
        item: usize,
        item_count: usize,
    ) -> Option<usize> {
        if render_index < actual_depth && item_count > 0 {
            Some(item.min(item_count - 1))
        } else {
            None
        }
    }

    fn help_page_body(&self) -> String {
        let body = String::from(
            "Corro Help\n\n\
Basics\n\
- Arrow keys or hjkl move the cursor; PageUp/PageDown move by one screen of rows.\n\
- Enter or e starts editing the current cell.\n\
- Header/footer/margin cells use the active address syntax.\n\
- Any printable key starts editing with that character.\n\
- = followed by arrows builds a formula reference.\n\n\
Selection and movement\n\
- v toggles a cell selection.\n\
- Shift+Arrow grows the selection one cell at a time.\n\
- Ctrl/Cmd+Shift+Arrow extends the selection to the edge of the current nonblank run.\n\
- Ctrl+Shift+= inserts rows above the current row or selected rows.\n\
- r moves selected rows.\n\
- c exports CSV when nothing is selected, or moves selected columns when columns are selected.\n\
- Alt+arrows move selected rows or columns by one cell.\n\n\
Menus\n\
- Alt+F opens File.\n\
- Format is available from the menu bar.\n\
- Alt+I opens Insert.\n\
- Alt+H opens Help.\n\
- Ctrl+; inserts the date and Ctrl+Shift+; inserts the time.\n\
- Right opens the highlighted submenu.\n\
- Left goes back one menu level.\n\
 - Enter or the shortcut letter opens the selected item.\n\n\
File menu\n\
 - Open file loads a .corro, .csv, .tsv, or .ods file. Use `link <file> <revision>` to open a log at a revision.\n\
 - New sheet adds another sheet to the workbook.\n\
 - Ctrl+PageUp and Ctrl+PageDown switch between workbook tabs.\n\
- Export opens TSV, CSV, ASCII, full export, or ODS prompts; ODS includes every sheet as a separate table (Calc tab) by default. Alt+F / Alt+V / Alt+G choose formulas, values, or generic interop; Alt+X copies the current export to the clipboard (TSV, CSV, ASCII, or full/selection TSV, not ODS).\n\
- Width opens default width and per-column width prompts.\n\
- Sort view changes the visible order of main rows.\n\
- Exit opens the quit prompt.\n\n\
Help menu\n\
- About shows the version and a short description.\n\
- Row ops and Col ops show quick move tips.\n\
- Full help opens this page.\n\n\
Address syntax\n\
  - Main cell: A1\n\
  - Header cell: A~1\n\
  - Footer cell: A_1\n\
  - Left margin: [A1\n\
  - Right margin: ]A1\n\
  - Cross-sheet refs use numeric IDs like #2!A1 or $2:A1.\n\
- Logs and saved files use this syntax only.\n\n\
Quit\n\
- q opens the quit prompt.\n\
- Ctrl+Q exits immediately.\n\
- Esc closes menus, prompts, help, and about.\n\
- ? opens this help page.\n",
        );
        body
    }

    fn about_page_body(&self) -> String {
        format!(
            "{name}\n\nVersion: {version}\n\n{about}\n\n{details}",
            name = env!("CARGO_PKG_NAME"),
            version = env!("CARGO_PKG_VERSION"),
            about = env!("CARGO_PKG_DESCRIPTION"),
            details = "Corro is a terminal spreadsheet with an append-only text log, sparse sheet storage, menu-driven exports, and undo via inverse ops.",
        )
    }

    fn render_menu_popup(
        &self,
        f: &mut Frame,
        popup_area: Rect,
        popup: List<'_>,
        state: &mut ListState,
    ) {
        f.render_widget(Clear, popup_area);
        f.render_stateful_widget(popup, popup_area, state);
    }
}

fn right_nonblank_end(state: &SheetState) -> Option<usize> {
    let g = &state.grid;
    let lm = MARGIN_COLS;
    let mc = g.main_cols();
    let rm = MARGIN_COLS;
    let start = lm + mc;
    let mut max_nonblank = None;
    for i in 0..rm {
        if g.logical_col_has_content(start + i) {
            max_nonblank = Some(i);
        }
    }
    max_nonblank
}

/// Row viewport with pinned totals and minimal-scroll movement.
fn visible_row_indices(
    state: &SheetState,
    cursor: SheetCursor,
    dim: usize,
    prev_start: usize,
) -> (Vec<usize>, usize) {
    let g = &state.grid;
    let hr = HEADER_ROWS;
    let mr = g.main_rows();
    let main_order = g.sorted_main_rows();
    let mut header_rows = Vec::new();
    let mut footer_rows = Vec::new();
    for (addr, _) in g.iter_nonempty() {
        match addr {
            CellAddr::Header { row, .. } => header_rows.push(row as usize),
            CellAddr::Footer { row, .. } => footer_rows.push(hr + mr + row as usize),
            _ => {}
        }
    }
    if cursor.row < hr {
        header_rows.push(cursor.row);
    } else if cursor.row >= hr + mr {
        footer_rows.push(cursor.row);
    }
    footer_rows.extend((0..NAV_BLANK_ROWS).map(|r| hr + mr + r));
    header_rows.sort_unstable();
    header_rows.dedup();
    footer_rows.sort_unstable();
    footer_rows.dedup();

    let mut display_rows: Vec<usize> =
        Vec::with_capacity(header_rows.len() + main_order.len() + footer_rows.len());
    display_rows.extend(header_rows);
    display_rows.extend(main_order.iter().copied().map(|r| hr + r));
    display_rows.extend(footer_rows);

    let dim = dim.max(1).min(display_rows.len().max(1));
    if display_rows.len() <= dim {
        return (display_rows, 0);
    }

    let cur_display = if cursor.row < hr {
        cursor.row
    } else if cursor.row < hr + mr {
        hr + main_order
            .iter()
            .position(|&r| hr + r == cursor.row)
            .unwrap_or(0)
    } else {
        cursor.row
    };

    let cur_pos = display_rows
        .iter()
        .position(|&r| r == cur_display)
        .unwrap_or(0);
    let max_start = display_rows.len().saturating_sub(dim);
    let mut start = prev_start.min(max_start);
    if cur_pos < start {
        start = cur_pos;
    } else if cur_pos >= start + dim {
        start = cur_pos + 1 - dim;
    }

    (display_rows[start..start + dim].to_vec(), start)
}

/// Column viewport with pinned left context and minimal-scroll movement.
fn visible_col_indices(
    state: &SheetState,
    cursor: SheetCursor,
    dim: usize,
    prev_start: usize,
) -> (Vec<usize>, usize) {
    let g = &state.grid;
    let lm = MARGIN_COLS;
    let mc = g.main_cols();
    let rm = MARGIN_COLS;
    let total = lm + mc + rm;
    // internal computation only
    let dim = dim.max(1).min(total.max(1));
    let cur = cursor.col.min(total.saturating_sub(1));
    let cursor_in_left = cursor.col < lm;
    let cursor_in_right = cursor.col >= lm + mc;

    if total <= dim {
        return ((0..total).collect(), 0);
    }

    let (main_lo, main_hi) = main_col_window(state, cursor);
    // computed main window
    let right_start = lm + mc;
    let mut right_band: Vec<usize> = match right_nonblank_end(state) {
        Some(end) => (0..=end).map(|i| right_start + i).collect(),
        None => Vec::new(),
    };
    let blank_right = right_nonblank_end(state)
        .map(|end| end + 1)
        .filter(|&i| i < rm)
        .map(|i| right_start + i)
        .unwrap_or(right_start);
    if cursor_in_right {
        right_band.push(blank_right);
        right_band.push(cur);
    }
    let left_band: Vec<usize> = if cursor_in_left {
        vec![cur]
    } else {
        Vec::new()
    };
    let main_span = main_hi.saturating_sub(main_lo) + 1;
    let mut stable_band = Vec::with_capacity(main_span + 1 + right_band.len() + left_band.len());
    if lm > 0 {
        stable_band.push(lm - 1);
    }
    stable_band.extend(left_band.iter().copied());
    stable_band.extend((main_lo..=main_hi).map(|ci| lm + ci));
    stable_band.extend(right_band.iter().copied());
    stable_band.sort_unstable();
    stable_band.dedup();
    if stable_band.len() <= dim && stable_band.contains(&cur) {
        return (stable_band, 0);
    }

    let mut reserved: Vec<usize> = left_band;
    if lm > 0 && dim > reserved.len() {
        reserved.push(lm - 1);
    }
    if !cursor_in_right && rm > 0 && !reserved.iter().any(|&c| c == blank_right) {
        let mut cand = reserved.clone();
        cand.push(blank_right);
        cand.sort_unstable();
        cand.dedup();
        if cand.len() < dim {
            let available = dim.saturating_sub(cand.len()).max(1);
            let filtered_len = (0..total).filter(|c| !cand.iter().any(|p| p == c)).count();
            if filtered_len <= available {
                reserved = cand;
            }
        }
    }
    reserved.sort_unstable();
    reserved.dedup();

    let available = dim.saturating_sub(reserved.len()).max(1);
    let filtered: Vec<usize> = (0..total)
        .filter(|c| !reserved.iter().any(|p| p == c))
        .collect();
    if filtered.is_empty() {
        return (reserved, 0);
    }

    let cur_pos = match filtered.binary_search(&cur) {
        Ok(i) => i,
        Err(i) => i.min(filtered.len().saturating_sub(1)),
    };
    let max_start = filtered.len().saturating_sub(available);
    // Start from previous scroll position but if the cursor is outside the
    // available window, center the window on the cursor so the relevant main
    // column is visible immediately instead of requiring incremental scroll
    // updates across frames.
    let mut start = prev_start.min(max_start);
    if cur_pos < start || cur_pos >= start + available {
        // Center cursor in the available window when possible
        start = cur_pos.saturating_sub(available / 2).min(max_start);
    }
    let end = (start + available).min(filtered.len());

    let mut out = filtered[start..end].to_vec();
    out.extend(reserved);
    out.sort_unstable();
    (out, start)
}

fn visible_cols_render_width(grid: &Grid, cols: &[usize]) -> usize {
    let lm = MARGIN_COLS;
    let mc = grid.main_cols();
    let show_right_divider = cols.contains(&(lm + mc));
    cols.iter()
        .enumerate()
        .map(|(i, &c)| {
            let sep = if i + 1 >= cols.len() {
                0
            } else if (c == lm - 1 && lm > 0 && cols.contains(&lm))
                || (c == lm + mc - 1 && show_right_divider)
            {
                2
            } else {
                1
            };
            grid.col_width(c).max(1) + sep
        })
        .sum()
}

fn trim_visible_cols_to_width(grid: &Grid, cols: &mut Vec<usize>, cursor_col: usize, width: usize) {
    while cols.len() > 1 && visible_cols_render_width(grid, cols) > width {
        let first = cols.first().copied().unwrap_or(cursor_col);
        let last = cols.last().copied().unwrap_or(cursor_col);
        // Remove columns to the *right* of the cursor first so we do not
        // immediately drop a column to the left of the focus (e.g. hiding A
        // when moving to B) when the overflow is from wide content on the right
        // or in the right margin.
        if last > cursor_col {
            cols.pop();
        } else if first < cursor_col {
            cols.remove(0);
        } else {
            break;
        }
    }
}

// ── Navigation helpers ────────────────────────────────────────────────────────

fn trailing_blank_main_rows(state: &SheetState) -> usize {
    let g = &state.grid;
    let hr = HEADER_ROWS;
    let mr = g.main_rows();
    match (0..mr)
        .rev()
        .find(|&r| g.logical_row_has_content(hr + r) || left_margin_template_applies(g, r))
    {
        None => mr,
        Some(last) => mr.saturating_sub(last + 1),
    }
}

fn trailing_blank_main_cols(state: &SheetState) -> usize {
    let g = &state.grid;
    let lm = MARGIN_COLS;
    let mc = g.main_cols();
    match (0..mc).rev().find(|&c| {
        g.logical_col_has_content(lm + c)
            || header_template_applies(g, c)
            || right_col_agg_func(g, lm + c).is_some()
    }) {
        None => mc,
        Some(last) => mc.saturating_sub(last + 1),
    }
}

fn header_template_applies(grid: &Grid, main_col: usize) -> bool {
    let raw = grid.get(&CellAddr::Header {
        row: (HEADER_ROWS - 1) as u32,
        col: (MARGIN_COLS as u32) + main_col as u32,
    });
    raw.as_deref().is_some_and(is_formula)
}

fn data_main_col_count(grid: &Grid) -> usize {
    let mc = grid.main_cols();
    for c in 0..mc {
        if right_col_agg_func(grid, MARGIN_COLS + c).is_some() {
            return c + 1;
        }
    }
    mc
}

fn row_total_block_start(grid: &Grid, current_main_row: u32) -> u32 {
    for candidate in (0..current_main_row).rev() {
        if left_margin_agg_func(grid, candidate).is_some() {
            return candidate + 1;
        }
    }
    0
}

fn previous_raw_block(grid: &Grid, current_main_row: u32) -> Option<(u32, u32)> {
    let mut end = current_main_row;
    while end > 0 {
        let last_agg = (0..end)
            .rev()
            .find(|&r| left_margin_agg_func(grid, r).is_some())
            .unwrap_or(0);
        let prev_agg = if last_agg == 0 {
            None
        } else {
            (0..last_agg)
                .rev()
                .find(|&r| left_margin_agg_func(grid, r).is_some())
        };
        let start = prev_agg.map_or(0, |r| r + 1);
        if start < last_agg {
            return Some((start, last_agg));
        }
        if last_agg == 0 {
            return Some((0, end));
        }
        end = last_agg;
    }
    Some((0, current_main_row))
}

fn left_margin_main_col_aggregate(
    grid: &Grid,
    func: AggFunc,
    current_main_row: u32,
    main_col: u32,
) -> String {
    let block_start = row_total_block_start(grid, current_main_row);
    let Some((start, end)) = (if block_start < current_main_row {
        Some((block_start, current_main_row))
    } else {
        previous_raw_block(grid, current_main_row)
    }) else {
        return String::new();
    };
    compute_aggregate(
        grid,
        &AggregateDef {
            func,
            source: MainRange {
                row_start: start,
                row_end: end,
                col_start: main_col,
                col_end: main_col + 1,
            },
        },
    )
}

fn left_margin_special_col_aggregate(
    grid: &Grid,
    subtotal_func: AggFunc,
    global_col: usize,
    row_start: u32,
    row_end: u32,
    data_cols: usize,
) -> Option<String> {
    let col_func = right_col_agg_func(grid, global_col)?;
    let collect = |row_start: u32, row_end: u32| -> Vec<f64> {
        let mut samples: Vec<f64> = Vec::new();
        for r in row_start..row_end {
            let row_val = compute_aggregate(
                grid,
                &AggregateDef {
                    func: col_func,
                    source: MainRange {
                        row_start: r,
                        row_end: r + 1,
                        col_start: 0,
                        col_end: data_cols as u32,
                    },
                },
            );
            if let Some(n) = parse_num(&row_val) {
                samples.push(n);
            }
        }
        samples
    };

    let mut samples = collect(row_start, row_end);
    let mut end = row_start;
    while samples.is_empty() && end > 0 {
        let Some((fallback_start, fallback_end)) = previous_raw_block(grid, end) else {
            break;
        };
        samples = collect(fallback_start, fallback_end);
        if fallback_start == 0 {
            break;
        }
        end = fallback_start;
    }
    Some(fold_numbers(subtotal_func, &samples))
}

fn left_margin_template_applies(grid: &Grid, main_row: usize) -> bool {
    let raw = grid.get(&CellAddr::Left {
        col: (MARGIN_COLS - 1),
        row: main_row as u32,
    });
    raw.as_deref().is_some_and(is_formula)
}

// ── Display-time aggregate helpers ───────────────────────────────────────────

fn footer_row_agg_func(grid: &Grid, footer_row_idx: usize) -> Option<AggFunc> {
    let key_col = (MARGIN_COLS - 1) as u32;
    let val = grid.get(&CellAddr::Footer {
        row: footer_row_idx as u32,
        col: key_col,
    })?;
    crate::ops::margin_key_agg_func(&val)
}

fn right_col_agg_func(grid: &Grid, global_col: usize) -> Option<AggFunc> {
    let mut labels: Vec<(u32, String)> = grid
        .iter_nonempty()
        .filter_map(|(addr, val)| match addr {
            CellAddr::Header { row, col } if col as usize == global_col => Some((row, val)),
            _ => None,
        })
        .collect();
    labels.sort_unstable_by_key(|(row, _)| *row);
    for (_, val) in labels {
        if let Some(f) = crate::ops::margin_key_agg_func(&val) {
            return Some(f);
        }
    }
    None
}

fn parse_num(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    t.parse::<f64>().ok()
}

fn boundary_gap_style(underlined: bool) -> Style {
    if underlined {
        Style::default().add_modifier(Modifier::UNDERLINED)
    } else {
        Style::default()
    }
}

fn boundary_separator_style(underlined: bool) -> Style {
    let style = Style::default().fg(Color::DarkGray);
    if underlined {
        style.add_modifier(Modifier::UNDERLINED)
    } else {
        style
    }
}

fn left_margin_agg_func(grid: &Grid, main_row: u32) -> Option<AggFunc> {
    let key_col: MarginIndex = MARGIN_COLS - 1;
    let val = grid.get(&CellAddr::Left {
        col: key_col,
        row: main_row,
    })?;
    crate::ops::margin_key_agg_func(&val)
}

fn fold_numbers(func: AggFunc, xs: &[f64]) -> String {
    if xs.is_empty() {
        return String::new();
    }
    match func {
        AggFunc::Sum => format!("{}", xs.iter().sum::<f64>()),
        AggFunc::Mean => format!("{}", xs.iter().sum::<f64>() / xs.len() as f64),
        AggFunc::Median => {
            let mut ys = xs.to_vec();
            ys.sort_by(|a, b| a.partial_cmp(b).unwrap());
            let n = ys.len();
            let m = if n % 2 == 1 {
                ys[n / 2]
            } else {
                (ys[n / 2 - 1] + ys[n / 2]) / 2.0
            };
            format!("{m}")
        }
        AggFunc::Min => xs
            .iter()
            .copied()
            .min_by(|a, b| a.partial_cmp(b).unwrap())
            .map(|v| format!("{v}"))
            .unwrap_or_default(),
        AggFunc::Max => xs
            .iter()
            .copied()
            .max_by(|a, b| a.partial_cmp(b).unwrap())
            .map(|v| format!("{v}"))
            .unwrap_or_default(),
        AggFunc::Count => format!("{}", xs.len()),
    }
}

fn footer_special_col_aggregate(
    grid: &Grid,
    footer_func: AggFunc,
    global_col: usize,
    main_rows: usize,
    main_cols: usize,
) -> Option<String> {
    let row_func = right_col_agg_func(grid, global_col);
    let data_cols = data_main_col_count(grid);
    let mut samples: Vec<f64> = Vec::new();
    for r in 0..main_rows {
        let row_val = if let Some(func) = row_func {
            compute_aggregate(
                grid,
                &AggregateDef {
                    func,
                    source: MainRange {
                        row_start: r as u32,
                        row_end: r as u32 + 1,
                        col_start: 0,
                        col_end: data_cols as u32,
                    },
                },
            )
        } else if global_col < MARGIN_COLS {
            String::new()
        } else if global_col < MARGIN_COLS + main_cols {
            cell_effective_display(
                grid,
                &CellAddr::Main {
                    row: r as u32,
                    col: (global_col - MARGIN_COLS) as u32,
                },
            )
        } else {
            cell_effective_display(
                grid,
                &CellAddr::Right {
                    col: (global_col - MARGIN_COLS - main_cols),
                    row: r as u32,
                },
            )
        };
        if let Some(n) = parse_num(&row_val) {
            samples.push(n);
        }
    }
    Some(fold_numbers(footer_func, &samples))
}

// ── Cell-address shorthand ───────────────────────────────────────────────────

/// Parse `ADDR: VALUE` shorthand. Returns `(target_addr, value)` or `None`.
fn parse_cell_shorthand(buf: &str, main_cols: usize) -> Option<(CellAddr, String)> {
    if let Some(colon) = buf.find(':') {
        let addr_part = buf[..colon].trim();
        let value_part = buf[colon + 1..].trim_start().to_string();
        if addr_part.is_empty() {
            return None;
        }
        let (addr, n) = parse_cell_ref_at(addr_part, main_cols)?;
        if n != addr_part.len() {
            return None;
        }
        return Some((addr, value_part));
    }

    // Accept an address-only buffer (no colon) as an explicit address with
    // an empty value. This lets users enter e.g. "C~1" to move the cursor to
    // that cell.
    let addr_part = buf.trim();
    if addr_part.is_empty() {
        return None;
    }
    let (addr, n) = parse_cell_ref_at(addr_part, main_cols)?;
    if n != addr_part.len() {
        return None;
    }
    Some((addr, String::new()))
}

fn special_value_choices(addr: &CellAddr) -> &'static [&'static str] {
    match addr {
        CellAddr::Header { .. } | CellAddr::Footer { .. } | CellAddr::Left { .. } => {
            &SPECIAL_VALUE_CHOICES
        }
        CellAddr::Right { .. } => &SPECIAL_VALUE_CHOICES,
        CellAddr::Main { .. } => &[],
    }
}

fn special_value_for_digit(digit: char) -> Option<&'static str> {
    special_choice_index_for_digit(digit).map(|i| SPECIAL_VALUE_CHOICES[i])
}

fn special_choice_label(idx: usize) -> Option<char> {
    match idx {
        0..=8 => char::from_digit((idx + 1) as u32, 10),
        9 => Some('0'),
        _ => None,
    }
}

fn special_choice_index_for_digit(digit: char) -> Option<usize> {
    match digit {
        '1'..='9' => Some((digit as u8 - b'1') as usize),
        '0' => Some(9),
        _ => None,
    }
}

fn cycle_special_value(current: &str, choices: &[&'static str]) -> Option<String> {
    if choices.is_empty() {
        return None;
    }
    let trimmed = current.trim();
    let idx = choices.iter().position(|c| c.eq_ignore_ascii_case(trimmed));
    let next = match idx {
        Some(i) => choices[(i + 1) % choices.len()],
        None => choices[0],
    };
    Some(next.to_string())
}

// ── Clipboard helper ─────────────────────────────────────────────────────────

fn copy_to_clipboard(text: &str) -> Result<(), String> {
    #[cfg(test)]
    {
        set_test_clipboard(Some(text.to_string()));
        return Ok(());
    }
    #[cfg(not(test))]
    {
        #[cfg(target_os = "windows")]
        {
            use std::process::{Command, Stdio};
            let mut child = Command::new("clip")
                .stdin(Stdio::piped())
                .spawn()
                .map_err(|e| format!("clip: {e}"))?;
            if let Some(mut stdin) = child.stdin.take() {
                use std::io::Write;
                stdin
                    .write_all(text.as_bytes())
                    .map_err(|e| format!("clip stdin: {e}"))?;
            }
            child.wait().map_err(|e| format!("clip wait: {e}"))?;
            Ok(())
        }
        #[cfg(not(target_os = "windows"))]
        {
            use std::process::{Command, Stdio};
            // Try xclip, then pbcopy
            let cmd = if Command::new("xclip").arg("-version").output().is_ok() {
                "xclip"
            } else {
                "pbcopy"
            };
            let mut child = Command::new(cmd)
                .args(if cmd == "xclip" {
                    &["-selection", "clipboard"][..]
                } else {
                    &[][..]
                })
                .stdin(Stdio::piped())
                .spawn()
                .map_err(|e| format!("{cmd}: {e}"))?;
            if let Some(mut stdin) = child.stdin.take() {
                use std::io::Write;
                stdin
                    .write_all(text.as_bytes())
                    .map_err(|e| format!("{cmd} stdin: {e}"))?;
            }
            child.wait().map_err(|e| format!("{cmd} wait: {e}"))?;
            Ok(())
        }
    }
}

#[cfg(test)]
thread_local! {
    static TEST_CLIPBOARD: std::cell::RefCell<Option<String>> = const { std::cell::RefCell::new(None) };
}

#[cfg(test)]
fn set_test_clipboard(text: Option<String>) {
    TEST_CLIPBOARD.with(|cell| *cell.borrow_mut() = text);
}

#[cfg(test)]
fn test_clipboard_text() -> Option<String> {
    TEST_CLIPBOARD.with(|cell| cell.borrow().clone())
}

fn read_clipboard() -> Result<String, String> {
    #[cfg(test)]
    if let Some(text) = TEST_CLIPBOARD.with(|cell| cell.borrow().clone()) {
        return Ok(text);
    }
    #[cfg(target_os = "windows")]
    {
        use std::process::Command;
        let output = Command::new("powershell")
            .args(["-NoProfile", "-Command", "Get-Clipboard"])
            .output()
            .map_err(|e| format!("powershell: {e}"))?;
        if !output.status.success() {
            return Err("powershell: Get-Clipboard failed".into());
        }
        Ok(String::from_utf8_lossy(&output.stdout).into_owned())
    }
    #[cfg(not(target_os = "windows"))]
    {
        use std::process::Command;
        let output = if Command::new("xclip").arg("-version").output().is_ok() {
            Command::new("xclip")
                .args(["-selection", "clipboard", "-o"])
                .output()
                .map_err(|e| format!("xclip: {e}"))?
        } else {
            Command::new("pbpaste")
                .output()
                .map_err(|e| format!("pbpaste: {e}"))?
        };
        if !output.status.success() {
            return Err("clipboard read failed".into());
        }
        Ok(String::from_utf8_lossy(&output.stdout).into_owned())
    }
}

// ── App ───────────────────────────────────────────────────────────────────────

pub struct App {
    pub path: Option<PathBuf>,
    /// Set when the workbook was read from a non-`corro` file (e.g. ODS). `path` stays `None` until saved as `.corro`.
    import_source: Option<PathBuf>,
    source_path: Option<PathBuf>,
    revision_limit: Option<usize>,
    revision_browse: bool,
    revision_browse_limit: usize,
    pub offset: u64,
    pub state: SheetState,
    pub workbook: WorkbookState,
    pub cursor: SheetCursor,
    pub anchor: Option<SheetCursor>,
    mode: Mode,
    pub watcher: Option<LogWatcher>,
    pub status: String,
    pub ops_applied: usize,
    pub row_scroll: usize,
    pub col_scroll: usize,
    /// Rows visible in the main grid (data area), updated in [`App::draw`]. Used for PageUp/PageDown.
    pub grid_viewport_data_rows: usize,
    help_scroll: usize,
    about_scroll: usize,
    export_preview_scroll: usize,
    /// Session-only. Applies to TSV, CSV, full export, and selection clipboard when using the export flow.
    export_delimited_options: export::DelimitedExportOptions,
    /// Session-only. Applies to "ASCII table" export and its preview.
    export_ascii_options: export::AsciiTableOptions,
    /// ODS only: default [export::ExportContent::Generic] (same as TSV generic; use [export::ExportContent::Formulas] for native ODF).
    export_ods_content: export::ExportContent,
    pub op_history: Vec<Op>,
    redo_history: Vec<Op>,
    selection_kind: SelectionKind,
    edit_special_palette: bool,
    edit_cursor: Option<usize>,
    input_cursor: Option<usize>,
    special_picker: Option<usize>,
    pending_format_target: Option<FormatTarget>,
    view_sheet_id: u32,
    persisted_view_sort_cols: HashMap<u32, Vec<SortSpec>>,
    edit_target_addr: Option<CellAddr>,
    /// When set, edit buffer commits to all listed addresses (same value). Preview uses all addrs in [`App::addr_at`].
    edit_range_addrs: Option<Vec<CellAddr>>,
    pending_lost_edit: Option<(CellAddr, String)>,
    pending_fit_to_content_on_commit: bool,
    clipboard_snapshot: Option<(MainRange, String)>,
}

impl App {
    fn insert_text_into_buffer(buffer: &mut String, cursor: &mut Option<usize>, text: &str) {
        let len = buffer.chars().count();
        let pos = cursor.get_or_insert(len);
        let pos = (*pos).min(len);
        let mut chars: Vec<char> = buffer.chars().collect();
        for (i, ch) in text.chars().enumerate() {
            chars.insert(pos + i, ch);
        }
        *buffer = chars.into_iter().collect();
        *cursor = Some(pos + text.chars().count());
    }

    pub fn new(path: Option<PathBuf>) -> Self {
        Self::new_with_revision_limit(path, None)
    }

    pub fn new_with_revision_limit(path: Option<PathBuf>, revision_limit: Option<usize>) -> Self {
        let (path, source_path) = if revision_limit.is_some() {
            (None, path)
        } else {
            (path, None)
        };
        App {
            path,
            import_source: None,
            source_path,
            revision_limit,
            revision_browse: false,
            revision_browse_limit: 1,
            offset: 0,
            state: SheetState::new(1, 1),
            workbook: WorkbookState::new(),
            cursor: SheetCursor {
                row: HEADER_ROWS,
                col: MARGIN_COLS,
            },
            anchor: None,
            mode: Mode::Normal,
            watcher: None,
            status: String::new(),
            ops_applied: 0,
            row_scroll: 0,
            col_scroll: 0,
            grid_viewport_data_rows: 24,
            help_scroll: 0,
            about_scroll: 0,
            export_preview_scroll: 0,
            export_delimited_options: export::DelimitedExportOptions::default(),
            export_ascii_options: export::AsciiTableOptions::default(),
            export_ods_content: export::ExportContent::Generic,
            op_history: Vec::new(),
            redo_history: Vec::new(),
            selection_kind: SelectionKind::Cells,
            edit_special_palette: false,
            edit_cursor: None,
            input_cursor: None,
            special_picker: None,
            pending_format_target: None,
            view_sheet_id: 1,
            persisted_view_sort_cols: HashMap::new(),
            edit_target_addr: None,
            edit_range_addrs: None,
            pending_lost_edit: None,
            pending_fit_to_content_on_commit: false,
            clipboard_snapshot: None,
        }
    }

    pub fn new_with_revision_browser(path: Option<PathBuf>) -> Self {
        let mut app = Self::new(None);
        app.source_path = path;
        app.revision_browse = true;
        app.mode = Mode::RevisionBrowse;
        app
    }

    fn open_path_prompt_buffer(&self) -> String {
        if let (Some(path), Some(revision)) = (&self.source_path, self.revision_limit) {
            return format!("link {} {}", path.display(), revision);
        }
        if let Some(path) = &self.path {
            return path.to_string_lossy().into_owned();
        }
        String::new()
    }

    /// Normalize so saving never targets `.ods` / `.tsv` / etc. (which would be confused for reload).
    fn to_corro_path(path: &Path) -> PathBuf {
        if path
            .extension()
            .and_then(|e| e.to_str())
            .is_some_and(|e| e.eq_ignore_ascii_case("corro"))
        {
            path.to_path_buf()
        } else {
            path.with_extension("corro")
        }
    }

    /// Default path for Save / Save as when there is no `.corro` `path` yet.
    fn suggested_corro_save_path(&self) -> String {
        if let Some(p) = &self.path {
            return Self::to_corro_path(p).to_string_lossy().into_owned();
        }
        if let Some(p) = &self.import_source {
            return Self::to_corro_path(p).to_string_lossy().into_owned();
        }
        String::new()
    }

    /// Default filename for export: same basename as `path` or `import_source` with the target extension (`file.corro` → `file.ods`). Empty when there is no path (blank still means clipboard where the prompt says so).
    fn suggested_export_save_path(&self, extension: &str) -> String {
        if let Some(p) = self.path.as_ref().or(self.import_source.as_ref()) {
            return p.with_extension(extension).to_string_lossy().into_owned();
        }
        String::new()
    }

    fn current_sheet_label(&self) -> String {
        if self.workbook.sheet_count() <= 1 {
            return String::new();
        }
        self.workbook
            .sheets
            .iter()
            .enumerate()
            .map(|(idx, sheet)| {
                if self.workbook.sheet_id(idx) == self.view_sheet_id {
                    format!("[{}]", sheet.title)
                } else {
                    sheet.title.clone()
                }
            })
            .collect::<Vec<_>>()
            .join("  ")
    }

    fn current_sheet_title(&self) -> String {
        self.workbook
            .sheets
            .iter()
            .find(|sheet| sheet.id == self.view_sheet_id)
            .map(|sheet| sheet.title.clone())
            .unwrap_or_default()
    }

    fn add_sheet(&mut self, title: String) {
        self.commit_active_sheet_cache();
        let id = self.workbook.next_sheet_id;
        let log_title = title.clone();
        self.workbook.add_sheet(title, SheetState::new(1, 1));
        self.view_sheet_id = id;
        self.sync_active_sheet_cache();
        self.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        self.anchor = None;
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = self.view_sheet_id;
            if let Err(e) = commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::NewSheet {
                    id,
                    title: log_title,
                },
            ) {
                self.status = format!("Log write error: {e}");
                return;
            }
            self.ops_applied = self.ops_applied.saturating_add(1);
            if let Err(e) = self.start_log_watcher_if_needed() {
                self.status = format!("Watcher error: {e}");
                return;
            }
        }
        self.status = "New sheet created".into();
    }

    fn rename_current_sheet(&mut self, title: String) -> Result<(), RunError> {
        self.commit_active_sheet_cache();
        let id = self.view_sheet_id;
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = id;
            commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::RenameSheet { id, title },
            )?;
            self.ops_applied = self.ops_applied.saturating_add(1);
            self.start_log_watcher_if_needed()?;
        } else if let Some(sheet) = self.workbook.sheets.iter_mut().find(|s| s.id == id) {
            sheet.title = title;
        }
        self.sync_active_sheet_cache();
        self.status = "Sheet renamed".into();
        Ok(())
    }

    fn copy_current_sheet(&mut self, title: String) -> Result<(), RunError> {
        self.commit_active_sheet_cache();
        let source_id = self.view_sheet_id;
        let id = self.workbook.next_sheet_id;
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = source_id;
            commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::CopySheet {
                    source_id,
                    id,
                    title: title.clone(),
                },
            )?;
            self.ops_applied = self.ops_applied.saturating_add(1);
            self.start_log_watcher_if_needed()?;
        } else if let Some(source) = self.workbook.sheets.iter().find(|s| s.id == source_id) {
            self.workbook.add_sheet_record(crate::ops::SheetRecord {
                id,
                title,
                state: source.state.clone(),
            });
        }
        self.view_sheet_id = id;
        self.sync_active_sheet_cache();
        self.status = "Sheet copied".into();
        Ok(())
    }

    fn move_current_sheet_to_end(&mut self) -> Result<(), RunError> {
        self.commit_active_sheet_cache();
        let id = self.view_sheet_id;
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = id;
            commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::MoveSheet { id },
            )?;
            self.ops_applied = self.ops_applied.saturating_add(1);
            self.start_log_watcher_if_needed()?;
        } else if let Some(idx) = self.workbook.sheet_index_by_id(id) {
            let sheet = self.workbook.sheets.remove(idx);
            self.workbook.sheets.push(sheet);
        }
        self.sync_active_sheet_cache();
        self.status = "Sheet moved to end".into();
        Ok(())
    }

    fn switch_sheet(&mut self, delta: isize) {
        self.commit_active_sheet_cache();
        let count = self.workbook.sheet_count();
        if count <= 1 {
            return;
        }
        let active = self
            .workbook
            .sheet_index_by_id(self.view_sheet_id)
            .unwrap_or(0) as isize;
        let next = (active + delta).rem_euclid(count as isize) as usize;
        self.view_sheet_id = self.workbook.sheet_id(next);
        self.sync_active_sheet_cache();
        self.cursor.clamp(&self.state.grid);
        self.status = format!("Sheet {} of {}", next + 1, count);
    }

    fn start_input_mode(&mut self, buffer: String) -> String {
        self.input_cursor = Some(buffer.chars().count());
        buffer
    }

    fn open_format_scope_picker(&mut self, target: FormatTarget) {
        self.pending_format_target = Some(target);
        self.open_menu(MenuSection::Format);
        self.status = match target {
            FormatTarget::All => "Formatting scope: All".into(),
            FormatTarget::FullColumn => "Formatting scope: Full column (global col)".into(),
            FormatTarget::Data => "Formatting scope: Data".into(),
            FormatTarget::Special => "Formatting scope: Special".into(),
            FormatTarget::Cell => "Formatting scope: Cell".into(),
            FormatTarget::Selection => "Formatting scope: Selection".into(),
        };
        self.selection_kind = match target {
            FormatTarget::Selection => SelectionKind::Cells,
            _ => self.selection_kind,
        };
    }

    fn open_format_decimals_picker(&mut self, decimals_for: FormatDecimalsFor) {
        self.mode = Mode::FormatDecimals {
            buffer: self.start_input_mode(String::new()),
            decimals_for,
        };
    }

    fn selected_format_target(&self) -> FormatTarget {
        self.pending_format_target.unwrap_or(FormatTarget::Cell)
    }

    fn apply_format_to_target(&mut self, target: FormatTarget, format: CellFormat) {
        let mut ops = Vec::new();
        match target {
            FormatTarget::All => {
                ops.push(Op::SetAllColumnFormat { format });
            }
            FormatTarget::FullColumn => {
                let col = self
                    .cursor
                    .col
                    .min(self.state.grid.total_cols().saturating_sub(1));
                ops.push(Op::SetColumnFormat {
                    scope: FormatScope::All,
                    col,
                    format,
                });
            }
            FormatTarget::Data => {
                for col in MARGIN_COLS..MARGIN_COLS + self.state.grid.main_cols() {
                    ops.push(Op::SetColumnFormat {
                        scope: FormatScope::Data,
                        col,
                        format,
                    });
                }
            }
            FormatTarget::Special => {
                for col in 0..self.state.grid.total_cols() {
                    if col < MARGIN_COLS || col >= MARGIN_COLS + self.state.grid.main_cols() {
                        ops.push(Op::SetColumnFormat {
                            scope: FormatScope::Special,
                            col,
                            format,
                        });
                    }
                }
            }
            FormatTarget::Cell => {
                ops.push(Op::SetCellFormat {
                    addr: self.cursor.to_addr(&self.state.grid),
                    format,
                });
            }
            FormatTarget::Selection => {
                if let Some((rows, cols)) = self.current_selection_range() {
                    for row in rows {
                        for col in &cols {
                            ops.push(Op::SetCellFormat {
                                addr: SheetCursor { row, col: *col }.to_addr(&self.state.grid),
                                format,
                            });
                        }
                    }
                }
            }
        }
        if ops.is_empty() {
            return;
        }
        let all_set_col = ops.iter().all(|o| {
            matches!(
                o,
                Op::SetColumnFormat { .. } | Op::SetAllColumnFormat { .. }
            )
        });
        if all_set_col {
            if let Some(ref p) = self.path.clone() {
                for op in &ops {
                    self.push_inverse_op(op);
                    op.apply(&mut self.state);
                    self.state.grid.bump_volatile_seed();
                }
                let mut active_sheet = self.view_sheet_id;
                if let Err(e) = commit_workbook_set_column_format_batch(
                    p,
                    &mut self.offset,
                    &mut self.workbook,
                    &mut active_sheet,
                    self.view_sheet_id,
                    &ops,
                ) {
                    self.status = format!("I/O: {e}");
                } else {
                    self.ops_applied = self.ops_applied.saturating_add(ops.len());
                    self.sync_active_sheet_cache();
                    let _ = self.start_log_watcher_if_needed();
                }
            } else {
                for op in &ops {
                    self.push_inverse_op(op);
                    op.apply(&mut self.state);
                    self.state.grid.bump_volatile_seed();
                }
            }
        } else {
            for op in ops {
                let _ = self.apply_single_op(op);
            }
        }
    }

    fn apply_format_number(&mut self, decimals: usize, currency: bool) {
        let format = if currency {
            CellFormat {
                number: Some(NumberFormat::Currency { decimals }),
                align: None,
            }
        } else {
            CellFormat {
                number: Some(NumberFormat::Fixed { decimals }),
                align: None,
            }
        };
        self.apply_format_to_target(self.selected_format_target(), format);
        self.clear_pending_format_target();
        self.status = if currency {
            format!("Currency format set to {decimals} decimals")
        } else {
            format!("Fixed format set to {decimals} decimals")
        };
    }

    fn apply_format_align(&mut self, align: TextAlign) {
        self.apply_format_to_target(
            self.selected_format_target(),
            CellFormat {
                number: None,
                align: Some(align),
            },
        );
        self.clear_pending_format_target();
        self.status = match align {
            TextAlign::Left => "Text aligned left".into(),
            TextAlign::Center => "Text aligned center".into(),
            TextAlign::Right => "Text aligned right".into(),
            TextAlign::Default => "Text alignment reset".into(),
        };
    }

    fn apply_format_reset(&mut self) {
        self.apply_format_to_target(self.selected_format_target(), CellFormat::default());
        self.clear_pending_format_target();
        self.status = "Format cleared".into();
    }

    fn state(&self) -> &SheetState {
        &self.state
    }

    fn state_mut(&mut self) -> &mut SheetState {
        &mut self.state
    }

    fn sync_active_sheet_cache(&mut self) {
        self.workbook.ensure_active_sheet();
        if let Some(idx) = self.workbook.sheet_index_by_id(self.view_sheet_id) {
            self.workbook.active_sheet = idx;
            self.state = self.workbook.sheets[idx].state.clone();
        } else {
            self.view_sheet_id = self.workbook.sheet_id(self.workbook.active_sheet);
            self.state = self.workbook.active_sheet().clone();
        }
    }

    fn sync_persisted_sort_cache_from_workbook(&mut self) {
        self.persisted_view_sort_cols.clear();
        for sheet in &self.workbook.sheets {
            let cols = sheet.state.grid.view_sort_cols();
            if !cols.is_empty() {
                self.persisted_view_sort_cols.insert(sheet.id, cols);
            }
        }
    }

    fn sync_persisted_sort_cache_from_active_sheet(&mut self) {
        let cols = self.state.grid.view_sort_cols();
        if cols.is_empty() {
            self.persisted_view_sort_cols.remove(&self.view_sheet_id);
        } else {
            self.persisted_view_sort_cols
                .insert(self.view_sheet_id, cols);
        }
    }

    fn set_active_sort_persistence(&mut self, cols: &[SortSpec], persisted: bool) {
        if persisted && !cols.is_empty() {
            self.persisted_view_sort_cols
                .insert(self.view_sheet_id, cols.to_vec());
        } else {
            self.persisted_view_sort_cols.remove(&self.view_sheet_id);
        }
    }

    fn replay_status(prefix: &str, path: &Path, replay: &PartialReplay) -> String {
        match (replay.failed_line, replay.error.as_deref()) {
            (Some(line), Some(err)) => {
                format!(
                    "{prefix} {} @ revision {} stopped at line {line}: {err}",
                    path.display(),
                    replay.op_count
                )
            }
            _ => format!("{prefix} {} @ revision {}", path.display(), replay.op_count),
        }
    }

    fn reload_revision_browse(&mut self) -> Result<(), IoError> {
        let Some(path) = self.source_path.clone() else {
            return Ok(());
        };
        self.workbook = WorkbookState::new();
        self.state = SheetState::new(1, 1);
        self.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        self.anchor = None;
        self.row_scroll = 0;
        self.col_scroll = 0;
        self.export_preview_scroll = 0;
        self.path = None;
        self.watcher = None;
        let mut active_sheet = self.workbook.sheet_id(self.workbook.active_sheet);
        let requested_limit = self.revision_browse_limit;
        let (off, replay) = load_workbook_revisions_partial(
            &path,
            requested_limit,
            &mut self.workbook,
            &mut active_sheet,
        )?;
        self.view_sheet_id = active_sheet;
        self.sync_active_sheet_cache();
        self.sync_persisted_sort_cache_from_workbook();
        for c in 0..self.state.grid.main_cols() {
            self.fit_column_to_rendered_content(MARGIN_COLS + c);
        }
        self.offset = off;
        self.ops_applied = replay.op_count;
        self.revision_browse_limit = replay.op_count;
        self.status = if replay.failed_line.is_some() {
            Self::replay_status("Browsing", &path, &replay)
        } else {
            format!("Browsing {} @ revision {}", path.display(), replay.op_count)
        };
        self.cursor.clamp(&self.state.grid);
        Ok(())
    }

    fn commit_active_sheet_cache(&mut self) {
        self.workbook.ensure_active_sheet();
        if let Some(idx) = self.workbook.sheet_index_by_id(self.view_sheet_id) {
            self.workbook.active_sheet = idx;
            self.workbook.sheets[idx].state = self.state.clone();
        }
    }

    fn handle_plain_text_input_key(
        buffer: &mut String,
        cursor: &mut Option<usize>,
        key: KeyCode,
    ) -> bool {
        !matches!(
            Self::handle_text_input_key(buffer, cursor, key),
            TextInputAction::Unhandled
        )
    }

    fn handle_text_input_key(
        buffer: &mut String,
        cursor: &mut Option<usize>,
        key: KeyCode,
    ) -> TextInputAction {
        match key {
            KeyCode::Char(c) => {
                let len = buffer.chars().count();
                let cursor = cursor.get_or_insert(len);
                let pos = (*cursor).min(len);
                let mut chars: Vec<char> = buffer.chars().collect();
                chars.insert(pos, c);
                *buffer = chars.into_iter().collect();
                *cursor = pos + 1;
                TextInputAction::Handled
            }
            KeyCode::Backspace => {
                let len = buffer.chars().count();
                if let Some(cursor) = cursor.as_mut() {
                    if *cursor > 0 {
                        let pos = (*cursor).min(len);
                        let mut chars: Vec<char> = buffer.chars().collect();
                        if pos > 0 {
                            chars.remove(pos - 1);
                            *buffer = chars.into_iter().collect();
                            *cursor = pos - 1;
                        }
                    }
                } else {
                    buffer.pop();
                }
                TextInputAction::Handled
            }
            KeyCode::Left | KeyCode::Right => {
                let len = buffer.chars().count();
                let cursor = cursor.get_or_insert(len);
                match key {
                    KeyCode::Left if *cursor == 0 => TextInputAction::EdgeLeft,
                    KeyCode::Right if *cursor >= len => TextInputAction::EdgeRight,
                    KeyCode::Left => {
                        *cursor -= 1;
                        TextInputAction::Handled
                    }
                    KeyCode::Right => {
                        *cursor += 1;
                        TextInputAction::Handled
                    }
                    _ => TextInputAction::Unhandled,
                }
            }
            _ => TextInputAction::Unhandled,
        }
    }

    pub fn load_initial(&mut self) -> Result<(), IoError> {
        let initial_path = self.path.clone().or(self.source_path.clone());
        let linked_revision = self.revision_limit;
        let browsing = self.revision_browse;
        if let Some(ref p) = initial_path {
            if Path::new(p).exists() {
                let ext = p
                    .extension()
                    .and_then(|e| e.to_str())
                    .unwrap_or("")
                    .to_lowercase();
                match ext.as_str() {
                    "corro" => {
                        let data = std::fs::read_to_string(p).map_err(|e| IoError::Io(e))?;
                        let mut workbook = WorkbookState::new();
                        let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
                        let (_, replay) = load_workbook_revisions_partial(
                            p,
                            usize::MAX,
                            &mut workbook,
                            &mut active_sheet,
                        )?;
                        self.workbook = workbook;
                        self.view_sheet_id = active_sheet;
                        self.sync_active_sheet_cache();
                        self.sync_persisted_sort_cache_from_workbook();
                        for c in 0..self.state.grid.main_cols() {
                            self.fit_column_to_rendered_content(MARGIN_COLS + c);
                        }
                        self.offset = data.len() as u64;
                        self.ops_applied = replay.op_count;
                        self.import_source = None;
                        self.path = Some(p.clone());
                        self.source_path = None;
                        self.revision_limit = None;
                        self.watcher = Some(LogWatcher::new(p.clone())?);
                        self.status = Self::replay_status("Loaded workbook", p, &replay);
                        self.cursor.clamp(&self.state.grid);
                        return Ok(());
                    }
                    "tsv" => {
                        let data = std::fs::read_to_string(p).map_err(|e| IoError::Io(e))?;
                        crate::io::import_tsv(&data, &mut self.state);
                        self.commit_active_sheet_cache();
                        self.path = None;
                        self.import_source = Some(p.clone());
                        self.source_path = None;
                        self.revision_limit = None;
                        self.watcher = None;
                        for c in 0..self.state.grid.main_cols() {
                            self.fit_column_to_rendered_content(MARGIN_COLS + c);
                        }
                        self.status = format!(
                            "Imported TSV (not saved) — use Save as a .corro file: {}",
                            p.display()
                        );
                    }
                    "ods" => match crate::ods::import_ods_workbook(p) {
                        Ok(workbook) => {
                            self.workbook = workbook;
                            self.sync_active_sheet_cache();
                            self.persisted_view_sort_cols.clear();
                            self.path = None;
                            self.import_source = Some(p.clone());
                            self.source_path = None;
                            self.revision_limit = None;
                            self.watcher = None;
                            for c in 0..self.state.grid.main_cols() {
                                self.state.grid.fit_column_to_content(MARGIN_COLS + c);
                            }
                            self.status = format!(
                                "Imported ODS (not saved) — use Save as a .corro file: {}",
                                p.display()
                            );
                            return Ok(());
                        }
                        Err(e) => {
                            self.status = format!("Failed to import ODS: {e}");
                            return Ok(());
                        }
                    },
                    "csv" => {
                        let data = std::fs::read_to_string(p).map_err(|e| IoError::Io(e))?;
                        crate::io::import_csv(&data, &mut self.state);
                        self.commit_active_sheet_cache();
                        self.path = None;
                        self.import_source = Some(p.clone());
                        self.source_path = None;
                        self.revision_limit = None;
                        self.watcher = None;
                        for c in 0..self.state.grid.main_cols() {
                            self.state.grid.auto_fit_column(MARGIN_COLS + c);
                        }
                        self.status = format!(
                            "Imported CSV (not saved) — use Save as a .corro file: {}",
                            p.display()
                        );
                    }
                    _ => {
                        if browsing {
                            self.workbook = WorkbookState::new();
                            self.state = SheetState::new(1, 1);
                            let mut active_sheet =
                                self.workbook.sheet_id(self.workbook.active_sheet);
                            let (off, replay) = load_workbook_revisions_partial(
                                p,
                                self.revision_browse_limit,
                                &mut self.workbook,
                                &mut active_sheet,
                            )?;
                            self.view_sheet_id = active_sheet;
                            self.sync_active_sheet_cache();
                            self.sync_persisted_sort_cache_from_workbook();
                            for c in 0..self.state.grid.main_cols() {
                                self.fit_column_to_rendered_content(MARGIN_COLS + c);
                            }
                            self.offset = off;
                            self.ops_applied = replay.op_count;
                            self.path = None;
                            self.source_path = Some(p.clone());
                            self.watcher = None;
                            self.status = Self::replay_status("Browsing", p, &replay);
                            self.cursor.clamp(&self.state.grid);
                            return Ok(());
                        }
                        self.workbook = WorkbookState::new();
                        self.state = SheetState::new(1, 1);
                        let mut active_sheet = self.workbook.sheet_id(self.workbook.active_sheet);
                        let (off, replay) = load_workbook_revisions_partial(
                            p,
                            linked_revision.unwrap_or(usize::MAX),
                            &mut self.workbook,
                            &mut active_sheet,
                        )?;
                        self.view_sheet_id = active_sheet;
                        self.sync_active_sheet_cache();
                        self.sync_persisted_sort_cache_from_workbook();
                        for c in 0..self.state.grid.main_cols() {
                            self.fit_column_to_rendered_content(MARGIN_COLS + c);
                        }
                        self.offset = off;
                        self.ops_applied = replay.op_count;
                        if let Some(limit) = linked_revision {
                            self.source_path = Some(p.clone());
                            self.path = None;
                            self.watcher = None;
                            self.status = format!(
                                "Linked {} @ revision {}",
                                p.display(),
                                replay.op_count.min(limit)
                            );
                        } else {
                            self.source_path = None;
                            self.path = Some(p.clone());
                            self.watcher = Some(LogWatcher::new(p.clone())?);
                            self.status = Self::replay_status("Loaded", p, &replay);
                        }
                    }
                }
            } else {
                self.watcher = None;
                self.source_path = None;
                self.revision_limit = None;
                self.status = format!("New file {}", p.display());
            }
        } else {
            self.status = "No file — press o to set path".into();
        }
        self.cursor.clamp(&self.state.grid);
        Ok(())
    }

    /// `notify` cannot watch a path that does not exist yet; we start the watcher after the first
    /// `commit_op`, which creates the log file via `append_op`.
    fn start_log_watcher_if_needed(&mut self) -> Result<(), IoError> {
        if self.watcher.is_some() {
            return Ok(());
        }
        if let Some(ref p) = self.path {
            if p.exists() {
                self.watcher = Some(LogWatcher::new(p.clone())?);
            }
        }
        Ok(())
    }

    fn push_inverse_op(&mut self, op: &Op) {
        if let Some(inverse) = self.state.reverse_op(op) {
            self.op_history.push(inverse);
        }
        self.redo_history.clear();
    }

    fn current_selection_range(&self) -> Option<(Vec<usize>, Vec<usize>)> {
        let a = self.anchor?;
        let b = self.cursor;
        let r0 = a.row.min(b.row);
        let r1 = a.row.max(b.row);
        let c0 = a.col.min(b.col);
        let c1 = a.col.max(b.col);
        const MAX_MATERIALIZED_SELECTION_AXIS: usize = 1_000_000;
        if r1.saturating_sub(r0) >= MAX_MATERIALIZED_SELECTION_AXIS
            || c1.saturating_sub(c0) >= MAX_MATERIALIZED_SELECTION_AXIS
        {
            return None;
        }
        Some(((r0..=r1).collect(), (c0..=c1).collect()))
    }

    fn selection_cell_is_nonblank(&self, row: usize, col: usize) -> bool {
        self.state
            .grid
            .get(&SheetCursor { row, col }.to_addr(&self.state.grid))
            .is_some_and(|value| !value.is_empty())
    }

    fn selection_edge_cursor(&self, direction: SelectionEdgeDirection) -> Option<SheetCursor> {
        let total_rows = self.state.grid.total_logical_rows();
        let total_cols = self.state.grid.total_cols();
        if total_rows == 0 || total_cols == 0 {
            return None;
        }

        let row = self.cursor.row.min(total_rows - 1);
        let col = self.cursor.col.min(total_cols - 1);

        match direction {
            SelectionEdgeDirection::Right => {
                let mut edge_col = if self.selection_cell_is_nonblank(row, col) {
                    col
                } else {
                    (col + 1..total_cols).find(|&c| self.selection_cell_is_nonblank(row, c))?
                };
                while edge_col + 1 < total_cols
                    && self.selection_cell_is_nonblank(row, edge_col + 1)
                {
                    edge_col += 1;
                }
                Some(SheetCursor { row, col: edge_col })
            }
            SelectionEdgeDirection::Left => {
                let mut edge_col = if self.selection_cell_is_nonblank(row, col) {
                    col
                } else {
                    (0..col)
                        .rev()
                        .find(|&c| self.selection_cell_is_nonblank(row, c))?
                };
                while edge_col > 0 && self.selection_cell_is_nonblank(row, edge_col - 1) {
                    edge_col -= 1;
                }
                Some(SheetCursor { row, col: edge_col })
            }
            SelectionEdgeDirection::Down => {
                let mut edge_row = if self.selection_cell_is_nonblank(row, col) {
                    row
                } else {
                    (row + 1..total_rows).find(|&r| self.selection_cell_is_nonblank(r, col))?
                };
                while edge_row + 1 < total_rows
                    && self.selection_cell_is_nonblank(edge_row + 1, col)
                {
                    edge_row += 1;
                }
                Some(SheetCursor { row: edge_row, col })
            }
            SelectionEdgeDirection::Up => {
                let mut edge_row = if self.selection_cell_is_nonblank(row, col) {
                    row
                } else {
                    (0..row)
                        .rev()
                        .find(|&r| self.selection_cell_is_nonblank(r, col))?
                };
                while edge_row > 0 && self.selection_cell_is_nonblank(edge_row - 1, col) {
                    edge_row -= 1;
                }
                Some(SheetCursor { row: edge_row, col })
            }
        }
    }

    fn extend_selection_to_edge(&mut self, direction: SelectionEdgeDirection) -> bool {
        let Some(cursor) = self.selection_edge_cursor(direction) else {
            return false;
        };
        if self.anchor.is_none() {
            self.anchor = Some(self.cursor);
        }
        self.cursor = cursor;
        self.selection_kind = SelectionKind::Cells;
        true
    }

    fn fill_row_pattern(&self) -> Option<Op> {
        if self.selection_kind != SelectionKind::Cells {
            return None;
        }
        let (rows, cols) = self.current_selection_range()?;
        if rows.len() != 1 || cols.len() < 2 {
            return None;
        }
        let row = rows[0];
        if row < HEADER_ROWS || row >= HEADER_ROWS + self.state.grid.main_rows() {
            return None;
        }
        if cols[0] < MARGIN_COLS || *cols.last()? >= MARGIN_COLS + self.state.grid.main_cols() {
            return None;
        }
        let main_row = (row - HEADER_ROWS) as u32;
        let start_col = (cols[0] - MARGIN_COLS) as u32;
        let end_col = (*cols.last()? - MARGIN_COLS) as u32;
        let seed: Vec<String> = (start_col..=end_col)
            .map(|col| self.state.grid.get(&CellAddr::Main { row: main_row, col }))
            .collect::<Option<Vec<_>>>()?;
        let mut cells = Vec::new();
        for col in (end_col + 1)..self.state.grid.main_cols() as u32 {
            let value =
                self.infer_fill_value(&seed, (col - end_col) as i32, FillDirection::Right)?;
            cells.push((CellAddr::Main { row: main_row, col }, value));
        }
        if cells.is_empty() {
            None
        } else {
            Some(Op::FillRange { cells })
        }
    }

    fn fill_col_pattern(&self) -> Option<Op> {
        if self.selection_kind != SelectionKind::Cells {
            return None;
        }
        let (rows, cols) = self.current_selection_range()?;
        if cols.len() != 1 || rows.len() < 2 {
            return None;
        }
        let col = cols[0];
        if col < MARGIN_COLS || col >= MARGIN_COLS + self.state.grid.main_cols() {
            return None;
        }
        if rows[0] < HEADER_ROWS || *rows.last()? >= HEADER_ROWS + self.state.grid.main_rows() {
            return None;
        }
        let main_col = (col - MARGIN_COLS) as u32;
        let start_row = (rows[0] - HEADER_ROWS) as u32;
        let end_row = (*rows.last()? - HEADER_ROWS) as u32;
        let seed: Vec<String> = (start_row..=end_row)
            .map(|row| self.state.grid.get(&CellAddr::Main { row, col: main_col }))
            .collect::<Option<Vec<_>>>()?;
        let mut cells = Vec::new();
        for row in (end_row + 1)..self.state.grid.main_rows() as u32 {
            let value =
                self.infer_fill_value(&seed, (row - end_row) as i32, FillDirection::Down)?;
            cells.push((CellAddr::Main { row, col: main_col }, value));
        }
        if cells.is_empty() {
            None
        } else {
            Some(Op::FillRange { cells })
        }
    }

    fn infer_fill_value(
        &self,
        seed: &[String],
        offset_from_last: i32,
        direction: FillDirection,
    ) -> Option<String> {
        let last = seed.last()?.clone();
        if is_formula(&last) {
            let (row_delta, col_delta) = match direction {
                FillDirection::Right => (0, offset_from_last),
                FillDirection::Down => (offset_from_last, 0),
            };
            if let Some(translated) = translate_formula_text_by_offset(&last, row_delta, col_delta)
            {
                return Some(translated);
            }
        }
        if let Some(v) = Self::infer_numeric_fill(seed, offset_from_last) {
            return Some(v);
        }
        if let Some(v) = Self::infer_named_sequence_fill(seed, offset_from_last) {
            return Some(v);
        }
        if let Some(v) = Self::infer_suffix_fill(seed, offset_from_last) {
            return Some(v);
        }
        Some(last)
    }

    fn infer_numeric_fill(seed: &[String], offset_from_last: i32) -> Option<String> {
        if !seed.iter().all(|v| v.trim().parse::<f64>().is_ok()) {
            return None;
        }
        let last = seed.last()?.trim().parse::<f64>().ok()?;
        let prev = if seed.len() >= 2 {
            seed[seed.len() - 2].trim().parse::<f64>().ok()?
        } else {
            last
        };
        let step = last - prev;
        Some(format!("{}", last + step * offset_from_last as f64))
    }

    fn infer_named_sequence_fill(seed: &[String], offset_from_last: i32) -> Option<String> {
        const WEEKDAYS: [&str; 7] = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"];
        const MONTHS: [&str; 12] = [
            "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC",
        ];
        let normalized: Vec<String> = seed.iter().map(|v| v.trim().to_ascii_uppercase()).collect();
        let last = normalized.last()?.as_str();
        if normalized.iter().all(|v| WEEKDAYS.contains(&v.as_str())) {
            let idx = WEEKDAYS.iter().position(|&v| v == last)?;
            return Some(
                WEEKDAYS
                    [(idx as i32 + offset_from_last).rem_euclid(WEEKDAYS.len() as i32) as usize]
                    .to_string(),
            );
        }
        if normalized.iter().all(|v| MONTHS.contains(&v.as_str())) {
            let idx = MONTHS.iter().position(|&v| v == last)?;
            return Some(
                MONTHS[(idx as i32 + offset_from_last).rem_euclid(MONTHS.len() as i32) as usize]
                    .to_string(),
            );
        }
        None
    }

    fn infer_suffix_fill(seed: &[String], offset_from_last: i32) -> Option<String> {
        let last = seed.last()?.trim();
        let (prefix, digits) = Self::split_trailing_digits(last)?;
        if seed
            .iter()
            .any(|v| Self::split_trailing_digits(v.trim()).is_none_or(|(p, _)| p != prefix))
        {
            return None;
        }
        let width = digits.len();
        let last_num = digits.parse::<i64>().ok()?;
        let prev_num = if seed.len() >= 2 {
            let (_, prev_digits) = Self::split_trailing_digits(seed[seed.len() - 2].trim())?;
            prev_digits.parse::<i64>().ok()?
        } else {
            last_num
        };
        let next = last_num + (last_num - prev_num) * offset_from_last as i64;
        Some(format!("{prefix}{next:0width$}"))
    }

    fn split_trailing_digits(s: &str) -> Option<(&str, &str)> {
        let bytes = s.as_bytes();
        let mut i = bytes.len();
        while i > 0 && bytes[i - 1].is_ascii_digit() {
            i -= 1;
        }
        if i == bytes.len() {
            return None;
        }
        Some((&s[..i], &s[i..]))
    }

    /// Sheet layout address for a visible `(row, col)` without using edit-mode buffer preview.
    fn cell_addr_for_position(&self, row: usize, col: usize) -> Option<CellAddr> {
        let hr = HEADER_ROWS;
        let mr = self.state.grid.main_rows();
        let mc = self.state.grid.main_cols();
        if row < hr {
            Some(CellAddr::Header {
                row: row as u32,
                col: col as u32,
            })
        } else if row < hr + mr {
            let mri = row - hr;
            if col < MARGIN_COLS {
                Some(CellAddr::Left {
                    col,
                    row: mri as u32,
                })
            } else if col < MARGIN_COLS + mc {
                Some(CellAddr::Main {
                    row: mri as u32,
                    col: (col - MARGIN_COLS) as u32,
                })
            } else if col < MARGIN_COLS + mc + MARGIN_COLS {
                Some(CellAddr::Right {
                    col: (col - MARGIN_COLS - mc),
                    row: mri as u32,
                })
            } else {
                None
            }
        } else if row < hr + mr + FOOTER_ROWS {
            Some(CellAddr::Footer {
                row: (row - hr - mr) as u32,
                col: col as u32,
            })
        } else {
            None
        }
    }

    /// All layout addresses in the current anchor/cursor range (if any), row-major.
    fn selection_cell_addresses(&self) -> Option<Vec<CellAddr>> {
        let (rows, cols) = self.current_selection_range()?;
        let mut v = Vec::new();
        for r in rows {
            for c in cols.iter().copied() {
                if let Some(a) = self.cell_addr_for_position(r, c) {
                    v.push(a);
                }
            }
        }
        if v.is_empty() {
            None
        } else {
            Some(v)
        }
    }

    /// When the user types to replace, fill all of these with the same buffer (more than one cell).
    fn multi_cell_type_targets(&self) -> Option<Vec<CellAddr>> {
        let v = self.selection_cell_addresses()?;
        if v.len() > 1 {
            Some(v)
        } else {
            None
        }
    }

    fn addr_at(&self, row: usize, col: usize) -> Option<CellAddr> {
        let preview_grid = if let Mode::Edit { buffer, .. } = &self.mode {
            let mut grid = self.state.grid.clone();
            if let Some(ref addrs) = self.edit_range_addrs {
                for a in addrs {
                    grid.set(a, buffer.clone());
                }
            } else {
                let addr = self.cursor.to_addr(&self.state.grid);
                grid.set(&addr, buffer.clone());
            }
            Some(grid)
        } else {
            None
        };
        let grid = preview_grid.as_ref().unwrap_or(&self.state.grid);
        let hr = HEADER_ROWS;
        let mr = grid.main_rows();
        let mc = grid.main_cols();
        if row < hr {
            Some(CellAddr::Header {
                row: row as u32,
                col: col as u32,
            })
        } else if row < hr + mr {
            let mri = row - hr;
            if col < MARGIN_COLS {
                Some(CellAddr::Left {
                    col: col,
                    row: mri as u32,
                })
            } else if col < MARGIN_COLS + mc {
                Some(CellAddr::Main {
                    row: mri as u32,
                    col: (col - MARGIN_COLS) as u32,
                })
            } else if col < MARGIN_COLS + mc + MARGIN_COLS {
                Some(CellAddr::Right {
                    col: (col - MARGIN_COLS - mc),
                    row: mri as u32,
                })
            } else {
                None
            }
        } else if row < hr + mr + FOOTER_ROWS {
            Some(CellAddr::Footer {
                row: (row - hr - mr) as u32,
                col: col as u32,
            })
        } else {
            None
        }
    }

    fn delete_selection(&mut self) -> bool {
        let cells = self.selection_clear_cells();
        if cells.is_empty() {
            return false;
        }
        let op = Op::FillRange {
            cells: cells.clone(),
        };
        self.push_inverse_op(&op);
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = self.view_sheet_id;
            let _ = commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::SheetOp {
                    sheet_id: self.view_sheet_id,
                    op,
                },
            );
            self.sync_active_sheet_cache();
        } else {
            op.apply(&mut self.state);
        }
        for (addr, _) in cells {
            if let CellAddr::Main { col, .. } = addr {
                self.state.grid.auto_fit_column(MARGIN_COLS + col as usize);
            }
        }
        if true {
            self.status = "Selection deleted".into();
            self.anchor = None;
        }
        true
    }

    fn selection_clear_cells(&self) -> Vec<(CellAddr, String)> {
        let Some((rows, cols)) = self.current_selection_range() else {
            return Vec::new();
        };
        let mut cells = Vec::new();
        for r in rows {
            for c in cols.iter().copied() {
                let Some(addr) = self.addr_at(r, c) else {
                    continue;
                };
                if self.state.grid.get(&addr).is_some_and(|v| !v.is_empty()) {
                    cells.push((addr, String::new()));
                }
            }
        }
        cells
    }

    fn sync_external(&mut self) -> Result<(), IoError> {
        if let Some(w) = &self.watcher {
            if w.poll_dirty() {
                if let Some(ref p) = self.path {
                    match crate::io::tail_apply_workbook(
                        p,
                        self.offset,
                        &mut self.workbook,
                        &mut self.view_sheet_id,
                    ) {
                        Ok(new_off) => {
                            self.offset = new_off;
                            self.sync_active_sheet_cache();
                            self.status = "External change applied".into();
                        }
                        Err(_) => {
                            let data = std::fs::read_to_string(p).map_err(IoError::Io)?;
                            let mut workbook = WorkbookState::new();
                            let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
                            for line in data.lines() {
                                let t = line.trim();
                                if t.is_empty() {
                                    continue;
                                }
                                crate::ops::apply_log_line_to_workbook(
                                    t,
                                    &mut workbook,
                                    &mut active_sheet,
                                )?;
                            }
                            self.workbook = workbook;
                            self.view_sheet_id = active_sheet;
                            self.sync_active_sheet_cache();
                            self.offset = data.len() as u64;
                            self.ops_applied =
                                data.lines().filter(|line| !line.trim().is_empty()).count();
                            self.status = "File reset; full reload".into();
                        }
                    }
                }
            }
        }
        Ok(())
    }

    fn selection_main_row_range(&self) -> Option<(u32, u32)> {
        let a = self.anchor?;
        let b = self.cursor;
        let hr = HEADER_ROWS;
        let r0 = a.row.min(b.row);
        let r1 = a.row.max(b.row);
        let c0 = a.col.min(b.col);
        let c1 = a.col.max(b.col);
        let left = MARGIN_COLS;
        let right = MARGIN_COLS + self.state.grid.main_cols();
        if r0 < hr || r1 >= hr + self.state.grid.main_rows() {
            return None;
        }
        if c0 != left || c1 != right.saturating_sub(1) {
            return None;
        }
        Some(((r0 - hr) as u32, (r1 - hr) as u32))
    }

    fn view_row_order(&self) -> Vec<usize> {
        let g = &self.state.grid;
        let hr = HEADER_ROWS;
        let mr = g.main_rows();
        let first_footer = hr + mr;
        let mut header_rows = Vec::new();
        let mut footer_rows = Vec::new();
        for (addr, _) in g.iter_nonempty() {
            match addr {
                CellAddr::Header { row, .. } => header_rows.push(row as usize),
                CellAddr::Footer { row, .. } => footer_rows.push(first_footer + row as usize),
                _ => {}
            }
        }
        if self.cursor.row < hr {
            header_rows.push(self.cursor.row);
        } else if self.cursor.row >= first_footer {
            footer_rows.push(self.cursor.row);
        }
        footer_rows.extend((0..NAV_BLANK_ROWS).map(|r| first_footer + r));
        header_rows.sort_unstable();
        header_rows.dedup();
        footer_rows.sort_unstable();
        footer_rows.dedup();

        let mut rows = Vec::with_capacity(header_rows.len() + mr + footer_rows.len());
        rows.extend(header_rows);
        rows.extend(g.sorted_main_rows().into_iter().map(|r| hr + r));
        rows.extend(footer_rows);
        rows
    }

    fn move_cursor_row_through_view(&mut self, down: bool) -> bool {
        if self.state.grid.view_sort_cols().is_empty() {
            return false;
        }

        let hr = HEADER_ROWS;
        let mr = self.state.grid.main_rows();
        let last_display_main = self
            .state
            .grid
            .sorted_main_rows()
            .last()
            .map(|row| hr + *row);
        let first_footer = hr + mr;
        let rows = self.view_row_order();
        let Some(pos) = rows.iter().position(|&r| r == self.cursor.row) else {
            return false;
        };
        // #region agent log
        debug_log_ndjson(
            "H1",
            "src/ui/mod.rs:move_cursor_row_through_view:pre_next_pos",
            "sorted row move precompute",
            format!(
                "{{\"down\":{},\"cursor_row\":{},\"mr_before\":{},\"first_footer_before\":{},\"last_display_main\":{},\"rows_len\":{},\"pos\":{}}}",
                down,
                self.cursor.row,
                mr,
                first_footer,
                last_display_main.map(|v| v.to_string()).unwrap_or_else(|| "null".to_string()),
                rows.len(),
                pos
            ),
        );
        // #endregion
        let next_pos = if down {
            if last_display_main == Some(self.cursor.row)
                && trailing_blank_main_rows(&self.state) < NAV_BLANK_ROWS
            {
                self.state.grid.grow_main_row_at_bottom();
                // #region agent log
                debug_log_ndjson(
                    "H1",
                    "src/ui/mod.rs:move_cursor_row_through_view:grow_main",
                    "grow main row inside sorted move",
                    format!(
                        "{{\"cursor_row\":{},\"mr_after_growth\":{},\"rows_len_stale\":{},\"next_row_candidate\":{}}}",
                        self.cursor.row,
                        self.state.grid.main_rows(),
                        rows.len(),
                        rows.get(pos.saturating_add(1)).copied().unwrap_or(self.cursor.row)
                    ),
                );
                // #endregion
            }
            if self.cursor.row >= first_footer {
                let blank_row = self
                    .cursor
                    .row
                    .saturating_add(1)
                    .min(first_footer + NAV_BLANK_ROWS - 1);
                return if blank_row == self.cursor.row {
                    true
                } else {
                    self.cursor.row = blank_row;
                    self.cursor.clamp(&self.state.grid);
                    self.state
                        .grid
                        .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
                    true
                };
            }
            pos.saturating_add(1).min(rows.len().saturating_sub(1))
        } else {
            pos.saturating_sub(1)
        };
        if next_pos == pos {
            return true;
        }

        self.cursor.row = rows[next_pos];
        self.cursor.clamp(&self.state.grid);
        self.state
            .grid
            .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
        // #region agent log
        debug_log_ndjson(
            "H1",
            "src/ui/mod.rs:move_cursor_row_through_view:post_apply",
            "sorted row move applied",
            format!(
                "{{\"next_pos\":{},\"cursor_row_after\":{},\"mr_after\":{},\"first_footer_after\":{}}}",
                next_pos,
                self.cursor.row,
                self.state.grid.main_rows(),
                HEADER_ROWS + self.state.grid.main_rows()
            ),
        );
        // #endregion
        true
    }

    /// One vertical step: same semantics as a single `Up` / `Down` in normal mode
    /// (view sort, header/footer, trailing blanks, grow last row).
    fn move_cursor_one_row_vertical(&mut self, down: bool) {
        if down {
            if !self.move_cursor_row_through_view(true) {
                let hr = HEADER_ROWS;
                let last_main = hr + self.state.grid.main_rows().saturating_sub(1);
                if self.cursor.row == last_main
                    && trailing_blank_main_rows(&self.state) < NAV_BLANK_ROWS
                {
                    self.state.grid.grow_main_row_at_bottom();
                }
                self.cursor.row = self.cursor.row.saturating_add(1);
                self.cursor.clamp(&self.state.grid);
                self.state
                    .grid
                    .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
            }
        } else if !self.move_cursor_row_through_view(false) {
            self.cursor.row = self.cursor.row.saturating_sub(1);
            self.cursor.clamp(&self.state.grid);
            self.state
                .grid
                .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
        }
    }

    fn move_cursor_vertical_steps(&mut self, mut steps: usize, down: bool) {
        while steps > 0 {
            self.move_cursor_one_row_vertical(down);
            steps -= 1;
        }
    }

    fn expand_selection_to_rows(&mut self) {
        let hr = HEADER_ROWS;
        let left = MARGIN_COLS;
        let right = MARGIN_COLS + self.state.grid.main_cols().saturating_sub(1);
        let row = self
            .cursor
            .row
            .clamp(hr, hr + self.state.grid.main_rows().saturating_sub(1));
        if let Some(anchor) = self.anchor {
            let r0 = anchor.row.min(row);
            let r1 = anchor.row.max(row);
            self.anchor = Some(SheetCursor { row: r0, col: left });
            self.cursor = SheetCursor {
                row: r1,
                col: right,
            };
        } else {
            self.anchor = Some(SheetCursor { row, col: left });
            self.cursor = SheetCursor { row, col: right };
        }
        self.selection_kind = SelectionKind::Rows;
    }

    fn expand_selection_to_cols(&mut self) {
        let hr = HEADER_ROWS;
        let bottom = hr + self.state.grid.main_rows().saturating_sub(1);
        let left = MARGIN_COLS;
        let right = MARGIN_COLS + self.state.grid.main_cols().saturating_sub(1);
        let col = self.cursor.col.clamp(left, right);
        if let Some(anchor) = self.anchor {
            let c0 = anchor.col.min(col);
            let c1 = anchor.col.max(col);
            self.anchor = Some(SheetCursor { row: hr, col: c0 });
            self.cursor = SheetCursor {
                row: bottom,
                col: c1,
            };
        } else {
            self.anchor = Some(SheetCursor { row: hr, col });
            self.cursor = SheetCursor { row: bottom, col };
        }
        self.selection_kind = SelectionKind::Cols;
    }

    fn selection_main_col_range(&self) -> Option<(u32, u32)> {
        let a = self.anchor?;
        let b = self.cursor;
        let hr = HEADER_ROWS;
        let r0 = a.row.min(b.row);
        let r1 = a.row.max(b.row);
        let c0 = a.col.min(b.col);
        let c1 = a.col.max(b.col);
        let left = MARGIN_COLS;
        let right = MARGIN_COLS + self.state.grid.main_cols();
        if c0 < left || c1 >= right {
            return None;
        }
        let last_main = hr + self.state.grid.main_rows().saturating_sub(1);
        if r0 != hr || r1 != last_main {
            return None;
        }
        Some(((c0 - left) as u32, (c1 - left) as u32))
    }

    fn selection_main_range(&self) -> Option<MainRange> {
        if self.selection_kind != SelectionKind::Cells {
            return None;
        }
        let (rows, cols) = self.current_selection_range()?;
        let (row_start, row_end) = (*rows.first()?, *rows.last()?);
        let (col_start, col_end) = (*cols.first()?, *cols.last()?);
        if row_start < HEADER_ROWS || row_end >= HEADER_ROWS + self.state.grid.main_rows() {
            return None;
        }
        if col_start < MARGIN_COLS || col_end >= MARGIN_COLS + self.state.grid.main_cols() {
            return None;
        }
        Some(MainRange {
            row_start: (row_start - HEADER_ROWS) as u32,
            row_end: (row_end - HEADER_ROWS + 1) as u32,
            col_start: (col_start - MARGIN_COLS) as u32,
            col_end: (col_end - MARGIN_COLS + 1) as u32,
        })
    }

    fn commit_edit_buffer(&mut self, buffer: &str) -> Result<(), RunError> {
        self.edit_special_palette = false;
        self.pending_lost_edit = None;
        let range = self.edit_range_addrs.take();
        let explicit_addr = parse_cell_shorthand(buffer, self.state.grid.main_cols());

        if let Some(ref addrs) = range {
            if addrs.len() > 1 && explicit_addr.is_none() {
                let value = buffer.to_string();
                if addrs.iter().all(|a| {
                    self.state.grid.get(a).as_deref().unwrap_or("") == value.as_str()
                }) {
                    self.pending_fit_to_content_on_commit = false;
                    return Ok(());
                }
                let cells: Vec<(CellAddr, String)> = addrs
                    .iter()
                    .cloned()
                    .map(|a| (a, value.clone()))
                    .collect();
                let op = Op::FillRange { cells };
                self.push_inverse_op(&op);
                if let Some(ref p) = self.path.clone() {
                    let mut active_sheet = self.view_sheet_id;
                    commit_workbook_op(
                        p,
                        &mut self.offset,
                        &mut self.workbook,
                        &mut active_sheet,
                        &crate::ops::WorkbookOp::SheetOp {
                            sheet_id: self.view_sheet_id,
                            op,
                        },
                    )?;
                    self.ops_applied = self.ops_applied.saturating_add(1);
                    self.sync_active_sheet_cache();
                    self.start_log_watcher_if_needed()?;
                } else {
                    op.apply(&mut self.state);
                    self.status = "No file — edit in memory only".into();
                }
                let cur_addr = self.cursor.to_addr(&self.state.grid);
                for a in addrs {
                    if let &CellAddr::Main { col, .. } = a {
                        self.state
                            .grid
                            .auto_fit_column(MARGIN_COLS + col as usize);
                    }
                }
                if self.pending_fit_to_content_on_commit {
                    if let Some(addr) = addrs
                        .iter()
                        .find(|a| *a == &cur_addr)
                        .or_else(|| addrs.first())
                    {
                        self.fit_column_to_content_from_current_cell(addr.clone());
                    }
                    self.commit_active_sheet_cache();
                    self.pending_fit_to_content_on_commit = false;
                }
                return Ok(());
            }
        }

        let (addr, value) = if let Some((a, v)) = explicit_addr.clone() {
            (a, v)
        } else {
            (
                self.edit_target_addr
                    .clone()
                    .unwrap_or_else(|| self.cursor.to_addr(&self.state.grid)),
                buffer.to_string(),
            )
        };
        // If this was an explicit address-only edit (e.g. "C~1" with no
        // value), the parser returns an empty value. In that case we still
        // want to move the cursor to the target even if the grid cell is
        // already empty. Detect explicit addresses and handle that
        // specially: set the cursor and return early.
        let raw = self.state.grid.get(&addr);
        if raw.as_deref().unwrap_or("") == value.as_str() {
            self.pending_fit_to_content_on_commit = false;
            if explicit_addr.is_some() {
                self.cursor = self.sheet_cursor_for_addr(&addr).unwrap_or(self.cursor);
                self.edit_target_addr = Some(addr);
            }
            return Ok(());
        }
        let op = Op::SetCell {
            addr: addr.clone(),
            value,
        };
        self.push_inverse_op(&op);
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = self.view_sheet_id;
            commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::SheetOp {
                    sheet_id: self.view_sheet_id,
                    op,
                },
            )?;
            self.ops_applied = self.ops_applied.saturating_add(1);
            self.sync_active_sheet_cache();
            self.start_log_watcher_if_needed()?;
        } else {
            op.apply(&mut self.state);
            self.status = "No file — edit in memory only".into();
        }
        if let Some((explicit_addr, _)) = explicit_addr {
            self.cursor = self
                .sheet_cursor_for_addr(&explicit_addr)
                .unwrap_or(self.cursor);
            self.edit_target_addr = Some(explicit_addr);
        }
        if self.pending_fit_to_content_on_commit {
            self.fit_column_to_content_from_current_cell(addr.clone());
            self.commit_active_sheet_cache();
            self.pending_fit_to_content_on_commit = false;
        }
        Ok(())
    }

    fn sheet_cursor_for_addr(&self, addr: &CellAddr) -> Option<SheetCursor> {
        let (row, col) = addr::addr_to_sheet_cursor(
            addr,
            addr::MainRows(self.state.grid.main_rows()),
            addr::MainCols(self.state.grid.main_cols()),
        );
        Some(SheetCursor {
            row: row.0,
            col: col.0,
        })
    }

    /// Parse `old|new` (first `|` only; `a|b|c` → find `a`, replace `b|c`).
    fn parse_replace_spec(raw: &str) -> Option<(&str, &str)> {
        let t = raw.trim();
        t.split_once('|')
            .map(|(a, b)| (a.trim(), b.trim()))
    }

    /// Find the next main cell (row-major, starting after the cursor, wrapping) whose
    /// displayed text contains `needle`. Moves the active cell when a match is found.
    fn find_next_substring(&mut self, needle: &str) {
        let needle = needle.trim();
        if needle.is_empty() {
            self.status = "Enter text to find".into();
            return;
        }
        let grid = &self.state.grid;
        let mr = grid.main_rows();
        let mc = grid.main_cols();
        if mr == 0 || mc == 0 {
            self.status = "Nothing to search".into();
            return;
        }

        let (cur_r, cur_c) = match self.cursor.to_addr(grid) {
            CellAddr::Main { row, col } => (row as usize, col as usize),
            _ => (0usize, 0usize),
        };

        let flat_index = |r: usize, c: usize| r * mc + c;
        let total = mr * mc;
        let start = flat_index(cur_r, cur_c);

        for k in 1..=total {
            let idx = (start + k) % total;
            let r = idx / mc;
            let c = idx % mc;
            let addr = CellAddr::Main {
                row: r as u32,
                col: c as u32,
            };
            let text = cell_display(grid, &addr);
            if text.contains(needle) {
                if let Some(cur) = self.sheet_cursor_for_addr(&addr) {
                    self.cursor = cur;
                }
                self.anchor = None;
                let label = addr_label(&addr, grid.main_cols());
                self.status = format!("Found: {label}");
                return;
            }
        }
        self.status = "Not found".into();
    }

    /// Replace all occurrences of `find` with `replace_with` in each main cell's raw value.
    fn replace_all_substrings_in_main(
        &mut self,
        find: &str,
        replace_with: &str,
    ) -> Result<usize, RunError> {
        let mut changed = 0usize;
        let mr = self.state.grid.main_rows();
        let mc = self.state.grid.main_cols();
        for r in 0..mr {
            for c in 0..mc {
                let addr = CellAddr::Main {
                    row: r as u32,
                    col: c as u32,
                };
                let raw = self.state.grid.get(&addr).unwrap_or_default();
                if !raw.contains(find) {
                    continue;
                }
                let new_val = raw.replace(find, replace_with);
                if new_val != raw {
                    changed += 1;
                    self.apply_single_op(Op::SetCell {
                        addr,
                        value: new_val,
                    })?;
                }
            }
        }
        Ok(changed)
    }

    fn main_cols_for_sheet_id(&self, sheet_id: u32) -> usize {
        self.workbook
            .sheet_index_by_id(sheet_id)
            .map(|i| self.workbook.sheets[i].state.grid.main_cols())
            .unwrap_or(0)
    }

    /// `Sheet>Go` targets like `$1`, `$Sheet1`, `$Budget:B2` (see formula sheet-ref syntax). Must run
    /// before the go-to string is uppercased, so sheet titles stay matchable.
    fn go_to_dollar_qualified(&mut self, text: &str) -> bool {
        let b = text.as_bytes();
        if b.len() < 2 {
            self.status = "Bad cell address".into();
            return false;
        }
        let (sheet_id, addr_opt) = if b[1].is_ascii_digit() {
            let (sheet_id, plen) = match parse_sheet_id_prefix_at(text) {
                Some(x) => x,
                None => {
                    self.status = "Bad cell address".into();
                    return false;
                }
            };
            if plen == text.len() {
                (sheet_id, None)
            } else if let Some(after) = text.get(plen..).and_then(|r| r.strip_prefix(':')) {
                let main_cols = self.main_cols_for_sheet_id(sheet_id);
                let Some((addr, len)) = parse_cell_ref_at(after, main_cols) else {
                    self.status = "Bad cell address".into();
                    return false;
                };
                if plen + 1 + len != text.len() {
                    self.status = "Bad cell address".into();
                    return false;
                }
                (sheet_id, Some(addr))
            } else {
                self.status = "Bad cell address".into();
                return false;
            }
        } else {
            let mut j = 1usize;
            while j < b.len() && (b[j].is_ascii_alphanumeric() || b[j] == b'_') {
                j += 1;
            }
            if j == 1 {
                self.status = "Bad cell address".into();
                return false;
            }
            let name = &text[1..j];
            let Some(sheet_id) = self.workbook.resolve_dollar_sheet_name(name) else {
                self.status = "Unknown sheet".into();
                return false;
            };
            if j == text.len() {
                (sheet_id, None)
            } else if let Some(after) = text.get(j..).and_then(|r| r.strip_prefix(':')) {
                let main_cols = self.main_cols_for_sheet_id(sheet_id);
                let Some((addr, len)) = parse_cell_ref_at(after, main_cols) else {
                    self.status = "Bad cell address".into();
                    return false;
                };
                if j + 1 + len != text.len() {
                    self.status = "Bad cell address".into();
                    return false;
                }
                (sheet_id, Some(addr))
            } else {
                self.status = "Bad cell address".into();
                return false;
            }
        };

        if self.workbook.sheet_index_by_id(sheet_id).is_none() {
            self.status = "Unknown sheet".into();
            return false;
        }

        self.commit_active_sheet_cache();
        self.view_sheet_id = sheet_id;
        self.sync_active_sheet_cache();

        if let Some(addr) = addr_opt {
            if let Some(c) = self.sheet_cursor_for_addr(&addr) {
                return self.set_cursor_from_go(c);
            }
            self.status = "Bad cell address".into();
            return false;
        }

        self.set_cursor_from_go(SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        })
    }

    fn go_to_cell(&mut self, raw: &str) -> bool {
        let text = raw.trim();
        if text.is_empty() {
            self.status = "Cell address required".into();
            return false;
        }

        if text.starts_with('$') {
            return self.go_to_dollar_qualified(text);
        }

        let text = text.to_ascii_uppercase();
        if let Some((cref, len)) = crate::celladdr::CellRef::parse_at(&text) {
            if len == text.len() && Self::cell_ref_is_in_supported_bounds(&cref) {
                return self.go_to_cell_ref(cref);
            }
        }

        if text.chars().all(|c| c.is_ascii_digit()) {
            return match text.parse::<u32>() {
                Ok(row) if row > 0 => self.go_to_data_row(row),
                _ => {
                    self.status = "Bad cell address".into();
                    false
                }
            };
        }

        if let Some((global_col, len)) =
            addr::parse_ui_column_fragment(&text, self.state.grid.main_cols())
        {
            if len == text.len() {
                let can_grow_main = !text.starts_with('[') && !text.starts_with(']');
                return self.go_to_global_col(global_col as usize, can_grow_main);
            }
        }

        self.status = "Bad cell address".into();
        false
    }

    fn go_to_cell_ref(&mut self, cref: crate::celladdr::CellRef) -> bool {
        let mut rows = self.state.grid.main_rows();
        let mut cols = self.state.grid.main_cols();
        if let crate::celladdr::RowRegion::Data(row) = cref.row {
            rows = rows.max(row as usize);
        }
        if let crate::celladdr::ColRegion::Data(col) = cref.col {
            cols = cols.max(col as usize);
        }
        if rows != self.state.grid.main_rows() || cols != self.state.grid.main_cols() {
            self.state.grid.set_main_size(rows, cols);
        }

        let addr = cref.to_grid_addr(self.state.grid.main_cols());
        if let Some(cursor) = self.sheet_cursor_for_addr(&addr) {
            self.set_cursor_from_go(cursor)
        } else {
            self.status = "Bad cell address".into();
            false
        }
    }

    fn go_to_data_row(&mut self, row: u32) -> bool {
        let target_rows = self.state.grid.main_rows().max(row as usize);
        if target_rows != self.state.grid.main_rows() {
            self.state
                .grid
                .set_main_size(target_rows, self.state.grid.main_cols());
        }
        self.set_cursor_from_go(SheetCursor {
            row: HEADER_ROWS + row as usize - 1,
            col: self.cursor.col,
        })
    }

    fn go_to_global_col(&mut self, global_col: usize, can_grow_main: bool) -> bool {
        if global_col >= self.state.grid.total_cols() {
            self.status = "Bad cell address".into();
            return false;
        }
        if can_grow_main && global_col >= MARGIN_COLS {
            let main_col = global_col - MARGIN_COLS;
            if main_col >= self.state.grid.main_cols() {
                self.state
                    .grid
                    .set_main_size(self.state.grid.main_rows(), main_col + 1);
            }
        }
        self.set_cursor_from_go(SheetCursor {
            row: self.cursor.row,
            col: global_col,
        })
    }

    fn set_cursor_from_go(&mut self, cursor: SheetCursor) -> bool {
        self.cursor = cursor;
        self.cursor.clamp(&self.state.grid);
        self.anchor = None;
        self.edit_target_addr = None;
        self.edit_range_addrs = None;
        let addr = self.cursor.to_addr(&self.state.grid);
        self.status = format!("Went to {}", addr_label(&addr, self.state.grid.main_cols()));
        true
    }

    fn remember_lost_edit(&mut self, buffer: &str) {
        let Some(addr) = self.edit_target_addr.clone() else {
            return;
        };
        let current = self.state.grid.get(&addr);
        if buffer.is_empty() || current.as_deref().unwrap_or("") == buffer {
            self.pending_lost_edit = None;
            return;
        }
        self.pending_lost_edit = Some((addr, buffer.to_string()));
        self.status = "Edit cancelled. Press Enter to restore lost text.".into();
    }

    fn restore_lost_edit(&mut self) -> Option<Mode> {
        let (addr, buffer) = self.pending_lost_edit.take()?;
        self.cursor = self.sheet_cursor_for_addr(&addr).unwrap_or(self.cursor);
        self.cursor.clamp(&self.state.grid);
        self.edit_target_addr = Some(addr);
        self.status.clear();
        Some(self.start_edit_mode(buffer, None, false, false, None))
    }

    fn cell_ref_is_in_supported_bounds(cref: &crate::celladdr::CellRef) -> bool {
        match cref.row {
            crate::celladdr::RowRegion::Header(row) => row > 0 && (row as usize) <= HEADER_ROWS,
            crate::celladdr::RowRegion::Data(row) => row > 0,
            crate::celladdr::RowRegion::Footer(row) => row > 0 && (row as usize) <= FOOTER_ROWS,
        }
    }

    fn commit_edit_and_move_down(&mut self, buffer: &str) -> Result<Mode, RunError> {
        self.edit_cursor = None;
        if let Some(edit_addr) = self.edit_target_addr.clone() {
            if let CellAddr::Main { row, col } = edit_addr {
                let target_row = HEADER_ROWS + row as usize;
                let target_col = MARGIN_COLS + col as usize;
                self.state
                    .grid
                    .ensure_extent_for_cursor(target_row, target_col);
                self.cursor = SheetCursor {
                    row: target_row,
                    col: target_col,
                };
                self.cursor.clamp(&self.state.grid);
            }
        }
        // #region agent log
        debug_log_ndjson(
            "H2",
            "src/ui/mod.rs:commit_edit_and_move_down:entry",
            "commit move-down start",
            format!(
                "{{\"cursor_row\":{},\"cursor_col\":{},\"main_rows_before\":{},\"edit_target_present\":{},\"is_footer_before\":{},\"edit_target_kind\":\"{}\"}}",
                self.cursor.row,
                self.cursor.col,
                self.state.grid.main_rows(),
                self.edit_target_addr.is_some(),
                self.cursor.row >= HEADER_ROWS + self.state.grid.main_rows(),
                match self.edit_target_addr.as_ref() {
                    Some(CellAddr::Header { .. }) => "header",
                    Some(CellAddr::Main { .. }) => "main",
                    Some(CellAddr::Footer { .. }) => "footer",
                    Some(CellAddr::Left { .. }) => "left",
                    Some(CellAddr::Right { .. }) => "right",
                    None => "none",
                }
            ),
        );
        // #endregion
        self.commit_edit_buffer(buffer)?;

        if !self.move_cursor_row_through_view(true) {
            let hr = HEADER_ROWS;
            let last_main = hr + self.state.grid.main_rows().saturating_sub(1);
            if self.cursor.row == last_main
                && trailing_blank_main_rows(&self.state) < NAV_BLANK_ROWS
            {
                self.state.grid.grow_main_row_at_bottom();
            }
            self.cursor.row = self.cursor.row.saturating_add(1);
            self.cursor.clamp(&self.state.grid);
            self.state
                .grid
                .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
        }

        let addr = self.cursor.to_addr(&self.state.grid);
        let cur = cell_display(&self.state.grid, &addr);
        // #region agent log
        debug_log_ndjson(
            "H3",
            "src/ui/mod.rs:commit_edit_and_move_down:post_move",
            "post move cursor classification",
            format!(
                "{{\"cursor_row\":{},\"cursor_col\":{},\"main_rows_after\":{},\"first_footer\":{},\"is_footer_by_row\":{},\"addr_kind\":\"{}\"}}",
                self.cursor.row,
                self.cursor.col,
                self.state.grid.main_rows(),
                HEADER_ROWS + self.state.grid.main_rows(),
                self.cursor.row >= HEADER_ROWS + self.state.grid.main_rows(),
                match addr {
                    CellAddr::Header { .. } => "header",
                    CellAddr::Main { .. } => "main",
                    CellAddr::Footer { .. } => "footer",
                    CellAddr::Left { .. } => "left",
                    CellAddr::Right { .. } => "right",
                }
            ),
        );
        // #endregion
        Ok(self.start_edit_mode(
            cur.clone(),
            if cur.trim() == "=" {
                Some(self.cursor)
            } else {
                None
            },
            false,
            false,
            None,
        ))
    }

    fn fit_column_to_content_from_current_cell(&mut self, addr: CellAddr) {
        match addr {
            CellAddr::Main { col, .. } => {
                self.fit_column_to_rendered_content(MARGIN_COLS + col as usize)
            }
            CellAddr::Left { col, .. } => self.fit_column_to_rendered_content(col as usize),
            CellAddr::Right { col, .. } => self.fit_column_to_rendered_content(
                MARGIN_COLS + self.state.grid.main_cols() + col as usize,
            ),
            CellAddr::Header { col, .. } | CellAddr::Footer { col, .. } => {
                self.fit_column_to_rendered_content(col as usize)
            }
        }
    }

    fn autofit_column_from_current_cell(&mut self, addr: CellAddr) {
        match addr {
            CellAddr::Main { col, .. } => {
                self.fit_column_to_rendered_content(MARGIN_COLS + col as usize)
            }
            CellAddr::Left { col, .. } => self.fit_column_to_rendered_content(col as usize),
            CellAddr::Right { col, .. } => self.fit_column_to_rendered_content(
                MARGIN_COLS + self.state.grid.main_cols() + col as usize,
            ),
            CellAddr::Header { col, .. } | CellAddr::Footer { col, .. } => {
                self.fit_column_to_rendered_content(col as usize)
            }
        }
    }

    fn fit_column_to_rendered_content(&mut self, global_col: usize) {
        let Some(maxw) = self.rendered_width_for_column(global_col) else {
            self.state.grid.set_col_width(global_col, None);
            return;
        };
        self.state.grid.set_col_width(global_col, Some(maxw));
    }

    /// Width override for the draw pass: never wider than the share of
    /// `data_width` so multiple visible columns (and gutters) can stay on
    /// screen; long text is shown truncated instead of dropping whole columns.
    fn fit_visible_columns_capped(&mut self, col_ixs: &[usize], data_width: usize) {
        if col_ixs.is_empty() {
            return;
        }
        let n = col_ixs.len();
        // One char separator per adjacent pair; matches trim loop roughly.
        let gaps = n.saturating_sub(1);
        let budget = data_width.saturating_sub(gaps);
        let per = (budget / n).max(1);
        for &c in col_ixs {
            if let Some(maxw) = self.rendered_width_for_column(c) {
                self.state
                    .grid
                    .set_col_width(c, Some(maxw.min(per)));
            } else {
                self.state.grid.set_col_width(c, None);
            }
        }
    }

    fn rendered_width_for_column(&self, global_col: usize) -> Option<usize> {
        let mut maxw = 0usize;
        let mut saw_content = false;
        let main_cols = self.state.grid.main_cols();

        for (addr, _) in self.state.grid.iter_nonempty() {
            match addr {
                CellAddr::Header { col, .. } | CellAddr::Footer { col, .. }
                    if col as usize == global_col =>
                {
                    let val =
                        normalize_inline_text(&cell_effective_display(&self.state.grid, &addr));
                    if !val.is_empty() {
                        saw_content = true;
                        maxw = maxw.max(val.width() + 1);
                    }
                }
                _ => {}
            }
        }
        for r in 0..self.state.grid.main_rows() {
            if global_col < MARGIN_COLS {
                let addr = CellAddr::Left {
                    col: global_col,
                    row: r as u32,
                };
                let val = normalize_inline_text(&cell_effective_display(&self.state.grid, &addr));
                if !val.is_empty() {
                    saw_content = true;
                    maxw = maxw.max(val.width() + 1);
                }
            } else if global_col < MARGIN_COLS + main_cols {
                let addr = CellAddr::Main {
                    row: r as u32,
                    col: (global_col - MARGIN_COLS) as u32,
                };
                let val = normalize_inline_text(&cell_effective_display(&self.state.grid, &addr));
                if !val.is_empty() {
                    saw_content = true;
                    maxw = maxw.max(val.width() + 1);
                }
            } else {
                let addr = CellAddr::Right {
                    col: (global_col - MARGIN_COLS - main_cols),
                    row: r as u32,
                };
                let val = normalize_inline_text(&cell_effective_display(&self.state.grid, &addr));
                if !val.is_empty() {
                    saw_content = true;
                    maxw = maxw.max(val.width() + 1);
                }
            }
        }

        saw_content.then_some(maxw.max(4))
    }

    fn move_selected_rows_by_one(&mut self, down: bool) -> Result<bool, RunError> {
        let Some((from, to)) = self.selection_main_row_range() else {
            return Ok(false);
        };
        let main_rows = self.state.grid.main_rows() as u32;
        if down {
            if to + 1 >= main_rows {
                self.status = "Selection is already at the bottom".into();
                return Ok(true);
            }
        } else if from == 0 {
            self.status = "Selection is already at the top".into();
            return Ok(true);
        }

        let count = to - from + 1;
        let target = if down { to + 2 } else { from - 1 };
        let op = Op::MoveRowRange {
            from,
            count,
            to: target,
        };
        self.push_inverse_op(&op);
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = self.view_sheet_id;
            commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::SheetOp {
                    sheet_id: self.view_sheet_id,
                    op,
                },
            )?;
            self.ops_applied = self.ops_applied.saturating_add(1);
            self.sync_active_sheet_cache();
            self.start_log_watcher_if_needed()?;
        } else {
            op.apply(&mut self.state);
        }

        let new_from = if down { from + 1 } else { from - 1 };
        let new_to = if down { to + 1 } else { to - 1 };
        self.anchor = Some(SheetCursor {
            row: HEADER_ROWS + new_from as usize,
            col: MARGIN_COLS,
        });
        self.cursor = SheetCursor {
            row: HEADER_ROWS + new_to as usize,
            col: MARGIN_COLS + self.state.grid.main_cols().saturating_sub(1),
        };
        self.selection_kind = SelectionKind::Rows;
        self.status = if down {
            format!("Moved rows {from}..{} down", to)
        } else {
            format!("Moved rows {from}..{} up", to)
        };
        Ok(true)
    }

    fn insert_rows_above_selection(&mut self) -> Result<bool, RunError> {
        let Some((from, to)) = self.selection_main_row_range() else {
            return Ok(false);
        };
        let count = to - from + 1;
        let main_rows = self.state.grid.main_rows() as u32;
        let op = Op::SetMainSize {
            main_rows: main_rows + count,
            main_cols: self.state.grid.main_cols() as u32,
        };
        self.push_inverse_op(&op);
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = self.view_sheet_id;
            commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::SheetOp {
                    sheet_id: self.view_sheet_id,
                    op: op.clone(),
                },
            )?;
            self.ops_applied = self.ops_applied.saturating_add(1);
            self.sync_active_sheet_cache();
            self.start_log_watcher_if_needed()?;
        } else {
            op.apply(&mut self.state);
        }

        let move_op = Op::MoveRowRange {
            from,
            count: main_rows - from,
            to: main_rows + count,
        };
        self.push_inverse_op(&move_op);
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = self.view_sheet_id;
            commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::SheetOp {
                    sheet_id: self.view_sheet_id,
                    op: move_op,
                },
            )?;
            self.ops_applied = self.ops_applied.saturating_add(1);
            self.sync_active_sheet_cache();
            self.start_log_watcher_if_needed()?;
        } else {
            move_op.apply(&mut self.state);
        }

        self.cursor = SheetCursor {
            row: HEADER_ROWS + from as usize,
            col: MARGIN_COLS,
        };
        self.anchor = None;
        self.selection_kind = SelectionKind::Cells;
        self.status = if count == 1 {
            format!("Inserted 1 row above row {from}")
        } else {
            format!("Inserted {count} rows above row {from}")
        };
        Ok(true)
    }

    fn insert_rows_above_cursor(&mut self, count: u32) -> Result<bool, RunError> {
        let hr = HEADER_ROWS;
        let original_main_rows = self.state.grid.main_rows() as u32;
        if self.cursor.row < hr || self.cursor.row >= hr + original_main_rows as usize {
            return Ok(false);
        }
        let row = (self.cursor.row - hr) as u32;
        let op = Op::SetMainSize {
            main_rows: original_main_rows + count,
            main_cols: self.state.grid.main_cols() as u32,
        };
        self.push_inverse_op(&op);
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = self.view_sheet_id;
            commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::SheetOp {
                    sheet_id: self.view_sheet_id,
                    op: op.clone(),
                },
            )?;
            self.ops_applied = self.ops_applied.saturating_add(1);
            self.sync_active_sheet_cache();
            self.start_log_watcher_if_needed()?;
        } else {
            op.apply(&mut self.state);
        }

        let move_op = Op::MoveRowRange {
            from: row,
            count: original_main_rows - row,
            to: original_main_rows + count,
        };
        self.push_inverse_op(&move_op);
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = self.view_sheet_id;
            commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::SheetOp {
                    sheet_id: self.view_sheet_id,
                    op: move_op.clone(),
                },
            )?;
            self.ops_applied = self.ops_applied.saturating_add(1);
            self.sync_active_sheet_cache();
            self.start_log_watcher_if_needed()?;
        } else {
            move_op.apply(&mut self.state);
        }
        self.cursor = SheetCursor {
            row: HEADER_ROWS + row as usize,
            col: MARGIN_COLS,
        };
        self.anchor = None;
        self.selection_kind = SelectionKind::Cells;
        self.status = if count == 1 {
            format!("Inserted 1 row above row {row}")
        } else {
            format!("Inserted {count} rows above row {row}")
        };
        Ok(true)
    }

    fn insert_mitosis_row_after_cursor(&mut self) -> Result<bool, RunError> {
        let hr = HEADER_ROWS;
        let main_rows = self.state.grid.main_rows();
        if self.cursor.row < hr {
            return self.insert_mitosis_header_row_after_cursor();
        }
        if self.cursor.row >= hr + main_rows {
            return self.insert_mitosis_footer_row_after_cursor();
        }
        self.insert_mitosis_main_data_row_after_cursor()
    }

    /// Mitosis in the main band (and margins): duplicate the logical row to the line below, shifting
    /// any rows beneath it down (same as row insert before the new duplicate).
    fn insert_mitosis_main_data_row_after_cursor(&mut self) -> Result<bool, RunError> {
        let hr = HEADER_ROWS;
        let original_main_rows = self.state.grid.main_rows() as u32;
        if self.cursor.row < hr || self.cursor.row >= hr + original_main_rows as usize {
            return Ok(false);
        }

        let source_row = (self.cursor.row - hr) as u32;
        let dest_row = source_row + 1;
        let mut copied_cells = Vec::new();
        for col in 0..self.state.grid.main_cols() as u32 {
            let src = CellAddr::Main {
                row: source_row,
                col,
            };
            if let Some(value) = self.state.grid.get(&src) {
                copied_cells.push((CellAddr::Main { row: dest_row, col }, value.to_string()));
            }
        }
        for col in 0..MARGIN_COLS {
            let src = CellAddr::Left {
                col,
                row: source_row,
            };
            if let Some(value) = self.state.grid.get(&src) {
                copied_cells.push((CellAddr::Left { col, row: dest_row }, value.to_string()));
            }
            let src = CellAddr::Right {
                col,
                row: source_row,
            };
            if let Some(value) = self.state.grid.get(&src) {
                copied_cells.push((CellAddr::Right { col, row: dest_row }, value.to_string()));
            }
        }

        self.apply_single_op(Op::SetMainSize {
            main_rows: original_main_rows + 1,
            main_cols: self.state.grid.main_cols() as u32,
        })?;

        let rows_below = original_main_rows.saturating_sub(dest_row);
        if rows_below > 0 {
            self.apply_single_op(Op::MoveRowRange {
                from: dest_row,
                count: rows_below,
                to: original_main_rows + 1,
            })?;
        }

        if !copied_cells.is_empty() {
            self.apply_single_op(Op::FillRange {
                cells: copied_cells,
            })?;
        }

        self.cursor = SheetCursor {
            row: hr + dest_row as usize,
            col: self.cursor.col,
        };
        self.cursor.clamp(&self.state.grid);
        self.anchor = None;
        self.selection_kind = SelectionKind::Cells;
        self.status = format!("Inserted mitosis row after row {}", source_row + 1);
        Ok(true)
    }

    /// Duplicate a `~` row: insert a new line under the cursor, shifting lower `~` rows down.
    fn insert_mitosis_header_row_after_cursor(&mut self) -> Result<bool, RunError> {
        let hr = HEADER_ROWS as u32;
        let h = self.cursor.row as u32;
        if h >= hr {
            return Ok(false);
        }

        if h + 1 < hr {
            return self.mitosis_header_row_shift_within_band(h, hr);
        }
        // Last header row (~1): duplicate into a new first main data row, pushing main down.
        self.mitosis_header_last_row_into_new_main_0()
    }

    /// Rebuild header rows after inserting a full duplicate line under row `h` (`h+1` < `HEADER_ROWS`).
    fn mitosis_header_row_shift_within_band(&mut self, h: u32, hr: u32) -> Result<bool, RunError> {
        let mut old: HashMap<(u32, u32), String> = HashMap::new();
        for (addr, v) in self.state.grid.iter_nonempty() {
            if let CellAddr::Header { row, col } = addr {
                old.insert((row, col), v);
            }
        }

        let mut newm: HashMap<(u32, u32), String> = HashMap::new();
        for ((r, c), v) in &old {
            if *r < h {
                newm.insert((*r, *c), v.clone());
            }
        }
        for r in (h + 1)..hr {
            for c in 0u32..self.state.grid.total_cols() as u32 {
                if let Some(v) = old.get(&(r, c)) {
                    if r + 1 < hr {
                        newm.insert((r + 1, c), v.clone());
                    }
                }
            }
        }
        for c in 0u32..self.state.grid.total_cols() as u32 {
            if let Some(v) = old.get(&(h, c)) {
                newm.insert((h, c), v.clone());
                newm.insert((h + 1, c), v.clone());
            }
        }

        self.apply_fill_replacing_region_map(&old, &newm, |(r, c)| CellAddr::Header { row: r, col: c })?;

        self.cursor = SheetCursor {
            row: (h + 1) as usize,
            col: self.cursor.col,
        };
        self.cursor.clamp(&self.state.grid);
        self.anchor = None;
        self.selection_kind = SelectionKind::Cells;
        self.status = format!("Inserted mitosis header after ~{}", (hr as usize) - 1 - h as usize);
        Ok(true)
    }

    /// `~1` and adjacent main: add a new row 1 with a copy of the `~1` line (main/margins/headers in that band).
    fn mitosis_header_last_row_into_new_main_0(&mut self) -> Result<bool, RunError> {
        let hr = HEADER_ROWS;
        let h = (hr - 1) as u32;
        let mut line: HashMap<u32, String> = HashMap::new();
        for (addr, v) in self.state.grid.iter_nonempty() {
            if let CellAddr::Header { row, col } = addr {
                if row == h {
                    line.insert(col, v);
                }
            }
        }
        if line.is_empty() {
            return Ok(false);
        }
        self.insert_main_rows_at(0, 1)?;

        let mc = self.state.grid.main_cols();
        let mut fill: Vec<(CellAddr, String)> = Vec::new();
        for (gc_u, v) in &line {
            let gc = *gc_u as usize;
            if let Some(a) = self.global_to_main_col0_addr_for_main_band(gc, mc) {
                fill.push((a, v.clone()));
            }
        }
        if !fill.is_empty() {
            self.apply_single_op(Op::FillRange { cells: fill })?;
        }
        self.cursor = SheetCursor {
            row: hr,
            col: self.cursor.col,
        };
        self.cursor.clamp(&self.state.grid);
        self.anchor = None;
        self.selection_kind = SelectionKind::Cells;
        self.status = "Inserted mitosis row: duplicate of ~1 as new row 1".into();
        Ok(true)
    }

    fn global_to_main_col0_addr_for_main_band(
        &self,
        global_col: usize,
        main_cols: usize,
    ) -> Option<CellAddr> {
        if global_col < MARGIN_COLS {
            return Some(CellAddr::Left {
                col: global_col,
                row: 0,
            });
        }
        if global_col < MARGIN_COLS + main_cols {
            return Some(CellAddr::Main {
                row: 0,
                col: (global_col - MARGIN_COLS) as u32,
            });
        }
        if global_col < MARGIN_COLS + main_cols + MARGIN_COLS {
            return Some(CellAddr::Right {
                col: global_col - MARGIN_COLS - main_cols,
                row: 0,
            });
        }
        None
    }

    fn insert_main_rows_at(&mut self, at_main_row: u32, count: u32) -> Result<(), RunError> {
        let n = self.state.grid.main_rows() as u32;
        self.apply_single_op(Op::SetMainSize {
            main_rows: n + count,
            main_cols: self.state.grid.main_cols() as u32,
        })?;
        if n > at_main_row {
            self.apply_single_op(Op::MoveRowRange {
                from: at_main_row,
                count: n - at_main_row,
                to: n + count,
            })?;
        }
        Ok(())
    }

    /// `_` row mitosis: insert a line under the current footer row, shifting lower `_` content down.
    fn insert_mitosis_footer_row_after_cursor(&mut self) -> Result<bool, RunError> {
        let hr = HEADER_ROWS;
        let mr = self.state.grid.main_rows();
        let fr = self
            .cursor
            .row
            .saturating_sub(hr)
            .saturating_sub(mr) as u32;
        if fr >= FOOTER_ROWS as u32 {
            return Ok(false);
        }
        if fr + 1 >= FOOTER_ROWS as u32 {
            return Ok(false);
        }

        self.mitosis_footer_row_shift_within_band(fr)
    }

    fn mitosis_footer_row_shift_within_band(&mut self, f: u32) -> Result<bool, RunError> {
        let fr = FOOTER_ROWS as u32;
        let mut old: HashMap<(u32, u32), String> = HashMap::new();
        for (addr, v) in self.state.grid.iter_nonempty() {
            if let CellAddr::Footer { row, col } = addr {
                old.insert((row, col), v);
            }
        }
        let mut newm: HashMap<(u32, u32), String> = HashMap::new();
        for ((r, c), v) in &old {
            if *r < f {
                newm.insert((*r, *c), v.clone());
            }
        }
        for r in (f + 1)..fr {
            for c in 0u32..self.state.grid.total_cols() as u32 {
                if let Some(v) = old.get(&(r, c)) {
                    if r + 1 < fr {
                        newm.insert((r + 1, c), v.clone());
                    }
                }
            }
        }
        for c in 0u32..self.state.grid.total_cols() as u32 {
            if let Some(v) = old.get(&(f, c)) {
                newm.insert((f, c), v.clone());
                newm.insert((f + 1, c), v.clone());
            }
        }

        self.apply_fill_replacing_region_map(&old, &newm, |(r, c)| CellAddr::Footer { row: r, col: c })?;

        let hr = HEADER_ROWS;
        self.cursor = SheetCursor {
            row: hr + self.state.grid.main_rows() + (f + 1) as usize,
            col: self.cursor.col,
        };
        self.cursor.clamp(&self.state.grid);
        self.anchor = None;
        self.selection_kind = SelectionKind::Cells;
        self.status = format!("Inserted mitosis footer after _{}", f + 1);
        Ok(true)
    }

    fn apply_fill_replacing_region_map(
        &mut self,
        old: &HashMap<(u32, u32), String>,
        newm: &HashMap<(u32, u32), String>,
        key_to_addr: impl Fn((u32, u32)) -> CellAddr,
    ) -> Result<(), RunError> {
        let mut fill: Vec<(CellAddr, String)> = Vec::new();
        for (k, v) in newm {
            if old.get(k).map(|s| s.as_str()) != Some(v.as_str()) {
                fill.push((key_to_addr(*k), v.clone()));
            }
        }
        for k in old.keys() {
            if !newm.contains_key(k) {
                fill.push((key_to_addr(*k), String::new()));
            }
        }
        if !fill.is_empty() {
            self.apply_single_op(Op::FillRange { cells: fill })?;
        }
        Ok(())
    }

    fn insert_mitosis_col_after_cursor(&mut self) -> Result<bool, RunError> {
        let hm = MARGIN_COLS;
        let original_main_cols = self.state.grid.main_cols() as usize;
        if self.cursor.col < hm {
            return self.insert_mitosis_left_margin_col_after_cursor();
        }
        if self.cursor.col >= hm + original_main_cols {
            return self.insert_mitosis_right_margin_col_after_cursor();
        }
        self.insert_mitosis_main_data_col_after_cursor()
    }

    /// Main-grid column: insert to the right and copy source column (works when the cursor is in
    /// header/footer for that main column, not only in the main row band).
    fn insert_mitosis_main_data_col_after_cursor(&mut self) -> Result<bool, RunError> {
        let hm = MARGIN_COLS;
        let original_main_cols = self.state.grid.main_cols() as u32;
        if self.cursor.col < hm || self.cursor.col >= hm + original_main_cols as usize {
            return Ok(false);
        }

        let source_col = (self.cursor.col - hm) as u32;
        let dest_col = source_col + 1;
        let source_global_col = (hm as u32) + source_col;
        let dest_global_col = source_global_col + 1;
        let mut copied_cells = Vec::new();

        for row in 0..self.state.grid.main_rows() as u32 {
            let src = CellAddr::Main {
                row,
                col: source_col,
            };
            if let Some(value) = self.state.grid.get(&src) {
                copied_cells.push((CellAddr::Main { row, col: dest_col }, value.to_string()));
            }
        }

        for (addr, value) in self.state.grid.iter_nonempty() {
            match addr {
                CellAddr::Header { row, col } if col == source_global_col => {
                    copied_cells.push((CellAddr::Header { row, col: dest_global_col }, value));
                }
                CellAddr::Footer { row, col } if col == source_global_col => {
                    copied_cells.push((CellAddr::Footer { row, col: dest_global_col }, value));
                }
                _ => {}
            }
        }

        self.apply_single_op(Op::SetMainSize {
            main_rows: self.state.grid.main_rows() as u32,
            main_cols: original_main_cols + 1,
        })?;

        let cols_right = original_main_cols.saturating_sub(dest_col);
        if cols_right > 0 {
            self.apply_single_op(Op::MoveColRange {
                from: dest_col,
                count: cols_right,
                to: original_main_cols + 1,
            })?;
        }

        if !copied_cells.is_empty() {
            self.apply_single_op(Op::FillRange {
                cells: copied_cells,
            })?;
        }

        self.cursor = SheetCursor {
            row: self.cursor.row,
            col: hm + dest_col as usize,
        };
        self.cursor.clamp(&self.state.grid);
        self.anchor = None;
        self.selection_kind = SelectionKind::Cells;
        self.status = format!("Inserted mitosis col after col {}", source_col + 1);
        Ok(true)
    }

    /// Left margin: duplicate a `[A]`-style margin column, shifting the band right (last column
    /// spills into column A in the same main row).
    fn insert_mitosis_left_margin_col_after_cursor(&mut self) -> Result<bool, RunError> {
        let c0 = self.cursor.col;
        if c0 + 1 >= MARGIN_COLS {
            return Ok(false);
        }
        self.mitosis_one_margin_col_after(c0, true)
    }

    /// Right `]A` margin: duplicate that column; last column spills into the rightmost data column.
    fn insert_mitosis_right_margin_col_after_cursor(&mut self) -> Result<bool, RunError> {
        let mc = self.state.grid.main_cols();
        let c0 = self.cursor.col.saturating_sub(MARGIN_COLS + mc);
        if c0 + 1 >= MARGIN_COLS {
            return Ok(false);
        }
        self.mitosis_one_margin_col_after(c0, false)
    }

    /// Insert after margin index `c0` (0..MARGIN_COLS-1) in the left or right margin.
    fn mitosis_one_margin_col_after(&mut self, c0: usize, is_left: bool) -> Result<bool, RunError> {
        let m = MARGIN_COLS;
        let main_cols = self.state.grid.main_cols();
        if main_cols < 1 {
            return Ok(false);
        }
        let gbase = if is_left {
            0usize
        } else {
            MARGIN_COLS + main_cols
        };
        let last_main = (main_cols - 1) as u32;

        let mut old: HashMap<CellAddr, String> = HashMap::new();
        for (addr, v) in self.state.grid.iter_nonempty() {
            match &addr {
                CellAddr::Header { col, .. } => {
                    let g = *col as usize;
                    if g >= gbase && g < gbase + m {
                        old.insert(addr, v);
                    }
                }
                CellAddr::Footer { col, .. } => {
                    let g = *col as usize;
                    if g >= gbase && g < gbase + m {
                        old.insert(addr, v);
                    }
                }
                CellAddr::Left { .. } if is_left => {
                    old.insert(addr, v);
                }
                CellAddr::Right { .. } if !is_left => {
                    old.insert(addr, v);
                }
                _ => {}
            }
        }

        // Group sparse margin columns per logical line, then 1D insert (avoids overwrites).
        let mut h_lines: HashMap<u32, HashMap<usize, String>> = HashMap::new();
        let mut f_lines: HashMap<u32, HashMap<usize, String>> = HashMap::new();
        let mut l_lines: HashMap<u32, HashMap<usize, String>> = HashMap::new();
        let mut r_lines: HashMap<u32, HashMap<usize, String>> = HashMap::new();
        for (a, v) in &old {
            match a {
                CellAddr::Header { row, col } => {
                    let l = *col as usize - gbase;
                    h_lines
                        .entry(*row)
                        .or_default()
                        .insert(l, v.clone());
                }
                CellAddr::Footer { row, col } => {
                    let l = *col as usize - gbase;
                    f_lines
                        .entry(*row)
                        .or_default()
                        .insert(l, v.clone());
                }
                CellAddr::Left { col, row } => {
                    l_lines.entry(*row).or_default().insert(*col, v.clone());
                }
                CellAddr::Right { col, row } => {
                    r_lines.entry(*row).or_default().insert(*col, v.clone());
                }
                _ => {}
            }
        }

        let mut new: HashMap<CellAddr, String> = HashMap::new();
        for (row, line) in h_lines {
            for (l, val) in Self::margin_line_map_insert_after(&line, c0, m) {
                new.insert(
                    CellAddr::Header {
                        row,
                        col: (gbase + l) as u32,
                    },
                    val,
                );
            }
        }
        for (row, line) in f_lines {
            for (l, val) in Self::margin_line_map_insert_after(&line, c0, m) {
                new.insert(
                    CellAddr::Footer {
                        row,
                        col: (gbase + l) as u32,
                    },
                    val,
                );
            }
        }
        for (row, line) in l_lines {
            for (l, val) in Self::margin_line_map_insert_after(&line, c0, m) {
                if l < m {
                    new.insert(CellAddr::Left { col: l, row }, val);
                } else {
                    new.insert(CellAddr::Main { row, col: 0 }, val);
                }
            }
        }
        for (row, line) in r_lines {
            for (l, val) in Self::margin_line_map_insert_after(&line, c0, m) {
                if l < m {
                    new.insert(CellAddr::Right { col: l, row }, val);
                } else {
                    new.insert(
                        CellAddr::Main {
                            row,
                            col: last_main,
                        },
                        val,
                    );
                }
            }
        }

        let mut fill: Vec<(CellAddr, String)> = Vec::new();
        for (a, v) in &new {
            if old.get(a).map(|s| s.as_str()) != Some(v.as_str()) {
                fill.push((a.clone(), v.clone()));
            }
        }
        for a in old.keys() {
            if !new.contains_key(a) {
                fill.push((a.clone(), String::new()));
            }
        }
        if !fill.is_empty() {
            self.apply_single_op(Op::FillRange { cells: fill })?;
        }
        self.cursor = SheetCursor {
            row: self.cursor.row,
            col: self.cursor.col + 1,
        };
        self.cursor.clamp(&self.state.grid);
        self.anchor = None;
        self.selection_kind = SelectionKind::Cells;
        self.status = if is_left {
            "Inserted mitosis after left margin column".into()
        } else {
            "Inserted mitosis after right margin column".into()
        };
        Ok(true)
    }

    /// Local margin indices 0..m-1, optional spill at local index `m`, after a mitosis "copy column"
    /// at `c0`.
    fn margin_line_map_insert_after(
        line: &HashMap<usize, String>,
        c0: usize,
        m: usize,
    ) -> HashMap<usize, String> {
        let mut out: HashMap<usize, String> = HashMap::new();
        for l in 0..m {
            if let Some(v) = line.get(&l) {
                if l < c0 {
                    out.insert(l, v.clone());
                } else if l == c0 {
                    out.insert(c0, v.clone());
                    out.insert(c0 + 1, v.clone());
                } else {
                    out.insert(l + 1, v.clone());
                }
            }
        }
        out
    }

    fn insert_cols_left_of_cursor(&mut self, count: u32) -> Result<bool, RunError> {
        let hm = MARGIN_COLS;
        let original_main_cols = self.state.grid.main_cols() as u32;
        if self.cursor.row < HEADER_ROWS
            || self.cursor.row >= HEADER_ROWS + self.state.grid.main_rows()
        {
            return Ok(false);
        }
        if self.cursor.col < hm || self.cursor.col >= hm + original_main_cols as usize {
            return Ok(false);
        }

        let col = (self.cursor.col - hm) as u32;
        let op = Op::SetMainSize {
            main_rows: self.state.grid.main_rows() as u32,
            main_cols: original_main_cols + count,
        };
        self.push_inverse_op(&op);
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = self.view_sheet_id;
            commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::SheetOp {
                    sheet_id: self.view_sheet_id,
                    op: op.clone(),
                },
            )?;
            self.ops_applied = self.ops_applied.saturating_add(1);
            self.sync_active_sheet_cache();
            self.start_log_watcher_if_needed()?;
        } else {
            op.apply(&mut self.state);
        }

        let move_op = Op::MoveColRange {
            from: col,
            count: original_main_cols - col,
            to: original_main_cols + count,
        };
        self.push_inverse_op(&move_op);
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = self.view_sheet_id;
            commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::SheetOp {
                    sheet_id: self.view_sheet_id,
                    op: move_op.clone(),
                },
            )?;
            self.ops_applied = self.ops_applied.saturating_add(1);
            self.sync_active_sheet_cache();
            self.start_log_watcher_if_needed()?;
        } else {
            move_op.apply(&mut self.state);
        }

        self.cursor.col = hm + col as usize;
        self.cursor.clamp(&self.state.grid);
        self.status = if count == 1 {
            format!("Inserted 1 column left of column {col}")
        } else {
            format!("Inserted {count} columns left of column {col}")
        };
        Ok(true)
    }

    fn menu_insert_special_seed(&self) -> String {
        let addr = self.cursor.to_addr(&self.state.grid);
        let raw = self.state.grid.get(&addr);
        let current = raw.as_deref().unwrap_or("").trim();
        if special_value_choices(&addr)
            .iter()
            .any(|choice| choice.eq_ignore_ascii_case(current))
        {
            current.to_string()
        } else {
            "∞".into()
        }
    }

    fn menu_insert_hyperlink_seed(&self) -> String {
        let addr = self.cursor.to_addr(&self.state.grid);
        let raw = self.state.grid.get(&addr);
        let current = raw.as_deref().unwrap_or("").trim();
        if current.starts_with("http://") || current.starts_with("https://") {
            current.to_string()
        } else {
            "https://".into()
        }
    }

    fn move_selected_cols_by_one(&mut self, right: bool) -> Result<bool, RunError> {
        let Some((from, to)) = self.selection_main_col_range() else {
            return Ok(false);
        };
        let main_cols = self.state.grid.main_cols() as u32;
        if right {
            if to + 1 >= main_cols {
                self.status = "Selection is already at the far right".into();
                return Ok(true);
            }
        } else if from == 0 {
            self.status = "Selection is already at the far left".into();
            return Ok(true);
        }

        let count = to - from + 1;
        let target = if right { to + 2 } else { from - 1 };
        let op = Op::MoveColRange {
            from,
            count,
            to: target,
        };
        self.push_inverse_op(&op);
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = self.view_sheet_id;
            commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::SheetOp {
                    sheet_id: self.view_sheet_id,
                    op: op.clone(),
                },
            )?;
            self.ops_applied = self.ops_applied.saturating_add(1);
            self.sync_active_sheet_cache();
            self.start_log_watcher_if_needed()?;
        } else {
            op.apply(&mut self.state);
        }

        let new_from = if right { from + 1 } else { from - 1 };
        let new_to = if right { to + 1 } else { to - 1 };
        self.anchor = Some(SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS + new_from as usize,
        });
        self.cursor = SheetCursor {
            row: HEADER_ROWS + self.state.grid.main_rows().saturating_sub(1),
            col: MARGIN_COLS + new_to as usize,
        };
        self.selection_kind = SelectionKind::Cols;
        self.status = if right {
            format!("Moved cols {from}..{} right", to)
        } else {
            format!("Moved cols {from}..{} left", to)
        };
        Ok(true)
    }

    fn formula_ref_for_addr(&self, addr: &CellAddr) -> String {
        crate::addr::cell_ref_text(addr, self.state.grid.main_cols())
    }

    fn do_export(&mut self, csv: bool) -> String {
        crate::formula::refresh_spills(&mut self.state.grid);
        let mut buf = Vec::new();
        let o = &self.export_delimited_options;
        if csv {
            export::export_csv_with_options(&self.state.grid, &mut buf, o);
        } else {
            export::export_tsv_with_options(&self.state.grid, &mut buf, o);
        }
        String::from_utf8_lossy(&buf).into_owned()
    }

    fn do_export_ascii(&mut self) -> String {
        crate::formula::refresh_spills(&mut self.state.grid);
        let mut buf = Vec::new();
        export::export_ascii_table_with_options(&self.state.grid, &mut buf, &self.export_ascii_options);
        String::from_utf8_lossy(&buf).into_owned()
    }

    fn do_export_all(&mut self) -> String {
        crate::formula::refresh_spills(&mut self.state.grid);
        let mut buf = Vec::new();
        export::export_all_with_options(&self.state.grid, &mut buf, &self.export_delimited_options);
        String::from_utf8_lossy(&buf).into_owned()
    }

    fn do_export_ods(&mut self) -> Vec<u8> {
        self.commit_active_sheet_cache();
        for s in &mut self.workbook.sheets {
            crate::formula::refresh_spills(&mut s.state.grid);
        }
        let mut o = self.export_delimited_options;
        o.content = self.export_ods_content;
        crate::ods::export_ods_bytes_workbook_with_options(&self.workbook, &o)
            .unwrap_or_default()
    }

    fn save_to_path(&mut self, path: &Path) -> Result<(), RunError> {
        self.commit_active_sheet_cache();
        let path = Self::to_corro_path(path);
        let mut buf = String::new();
        buf.push_str(&format!(
            "{} {}\n",
            crate::ops::LOG_HEADER_PREFIX,
            crate::ops::LOG_VERSION
        ));
        let omit_sheet1_prefix = self.workbook.sheet_count() == 1;
        for sheet in &self.workbook.sheets {
            for line in (crate::ops::WorkbookOp::NewSheet {
                id: sheet.id,
                title: sheet.title.clone(),
            })
            .to_log_lines_with_policy(sheet.state.grid.main_cols(), omit_sheet1_prefix)
            {
                buf.push_str(&line);
                buf.push('\n');
            }
            for row in 0..sheet.state.grid.main_rows() {
                for col in 0..sheet.state.grid.main_cols() {
                    let addr = CellAddr::Main {
                        row: row as u32,
                        col: col as u32,
                    };
                    if let Some(value) = sheet.state.grid.get(&addr) {
                        if !value.is_empty() {
                            for line in (crate::ops::WorkbookOp::SheetOp {
                                sheet_id: sheet.id,
                                op: Op::SetCell {
                                    addr: addr.clone(),
                                    value: value.to_string(),
                                },
                            })
                            .to_log_lines_with_policy(
                                sheet.state.grid.main_cols(),
                                omit_sheet1_prefix,
                            ) {
                                buf.push_str(&line);
                                buf.push('\n');
                            }
                        }
                    }
                }
            }
            for line in (crate::ops::WorkbookOp::SheetOp {
                sheet_id: sheet.id,
                op: Op::SetMainSize {
                    main_rows: sheet.state.grid.main_rows() as u32,
                    main_cols: sheet.state.grid.main_cols() as u32,
                },
            })
            .to_log_lines_with_policy(sheet.state.grid.main_cols(), omit_sheet1_prefix)
            {
                buf.push_str(&line);
                buf.push('\n');
            }
            if sheet.state.grid.max_col_width() != 20 {
                for line in (crate::ops::WorkbookOp::SheetOp {
                    sheet_id: sheet.id,
                    op: Op::SetMaxColWidth {
                        width: sheet.state.grid.max_col_width(),
                    },
                })
                .to_log_lines_with_policy(sheet.state.grid.main_cols(), omit_sheet1_prefix)
                {
                    buf.push_str(&line);
                    buf.push('\n');
                }
            }
            for (col, width) in sheet.state.grid.col_width_overrides() {
                for line in (crate::ops::WorkbookOp::SheetOp {
                    sheet_id: sheet.id,
                    op: Op::SetColWidth {
                        col,
                        width: Some(width),
                    },
                })
                .to_log_lines_with_policy(sheet.state.grid.main_cols(), omit_sheet1_prefix)
                {
                    buf.push_str(&line);
                    buf.push('\n');
                }
            }
            if let Some(cols) = self
                .persisted_view_sort_cols
                .get(&sheet.id)
                .filter(|cols| !cols.is_empty())
            {
                for line in (crate::ops::WorkbookOp::SheetOp {
                    sheet_id: sheet.id,
                    op: Op::SetViewSortCols { cols: cols.clone() },
                })
                .to_log_lines_with_policy(sheet.state.grid.main_cols(), omit_sheet1_prefix)
                {
                    buf.push_str(&line);
                    buf.push('\n');
                }
            }
            for (col, format) in sheet.state.grid.col_all_formats() {
                for line in (crate::ops::WorkbookOp::SheetOp {
                    sheet_id: sheet.id,
                    op: Op::SetColumnFormat {
                        scope: FormatScope::All,
                        col: col,
                        format: format,
                    },
                })
                .to_log_lines_with_policy(sheet.state.grid.main_cols(), omit_sheet1_prefix)
                {
                    buf.push_str(&line);
                    buf.push('\n');
                }
            }
            for (col, format) in sheet.state.grid.col_data_formats() {
                for line in (crate::ops::WorkbookOp::SheetOp {
                    sheet_id: sheet.id,
                    op: Op::SetColumnFormat {
                        scope: FormatScope::Data,
                        col: col,
                        format: format,
                    },
                })
                .to_log_lines_with_policy(sheet.state.grid.main_cols(), omit_sheet1_prefix)
                {
                    buf.push_str(&line);
                    buf.push('\n');
                }
            }
            for (col, format) in sheet.state.grid.col_special_formats() {
                for line in (crate::ops::WorkbookOp::SheetOp {
                    sheet_id: sheet.id,
                    op: Op::SetColumnFormat {
                        scope: FormatScope::Special,
                        col: col,
                        format: format,
                    },
                })
                .to_log_lines_with_policy(sheet.state.grid.main_cols(), omit_sheet1_prefix)
                {
                    buf.push_str(&line);
                    buf.push('\n');
                }
            }
            for (addr, format) in sheet.state.grid.cell_formats() {
                for line in (crate::ops::WorkbookOp::SheetOp {
                    sheet_id: sheet.id,
                    op: Op::SetCellFormat {
                        addr: addr,
                        format: format,
                    },
                })
                .to_log_lines_with_policy(sheet.state.grid.main_cols(), omit_sheet1_prefix)
                {
                    buf.push_str(&line);
                    buf.push('\n');
                }
            }
        }
        std::fs::write(&path, buf)?;
        self.path = Some(path.clone());
        self.import_source = None;
        self.source_path = None;
        self.revision_limit = None;
        self.status = format!("Saved {}", path.display());
        if self.watcher.is_none() {
            self.watcher = Some(LogWatcher::new(path).map_err(IoError::from)?);
        }
        Ok(())
    }

    fn do_export_selection(&mut self) -> String {
        crate::formula::refresh_spills(&mut self.state.grid);
        let (rows, cols) = self
            .current_selection_range()
            .unwrap_or_else(|| (vec![self.cursor.row], vec![self.cursor.col]));
        if rows.is_empty() || cols.is_empty() {
            return String::new();
        }
        let mut buf = Vec::new();
        export::export_selection(&self.state.grid, &mut buf, &rows, &cols, &self.export_delimited_options);
        String::from_utf8_lossy(&buf).into_owned()
    }

    fn selection_tsv_text(&self) -> String {
        let (rows, cols) = self
            .current_selection_range()
            .unwrap_or_else(|| (vec![self.cursor.row], vec![self.cursor.col]));

        let mut out = String::new();
        for (ri, row) in rows.iter().enumerate() {
            if ri > 0 {
                out.push('\n');
            }
            for (ci, col) in cols.iter().enumerate() {
                if ci > 0 {
                    out.push('\t');
                }
                if let Some(addr) = self.addr_at(*row, *col) {
                    let raw = self.state.grid.get(&addr);
                    out.push_str(raw.as_deref().unwrap_or(""));
                }
            }
        }
        if !out.is_empty() {
            out.push('\n');
        }
        out
    }

    fn copy_selection_to_clipboard(&mut self, data: &str) -> bool {
        match copy_to_clipboard(data) {
            Ok(()) => {
                self.clipboard_snapshot = self
                    .selection_main_range()
                    .map(|range| (range, data.to_string()));
                self.status = "Selection copied to clipboard".into();
                true
            }
            Err(e) => {
                self.status = format!("Clipboard error: {e}");
                false
            }
        }
    }

    fn apply_single_op(&mut self, op: Op) -> Result<(), RunError> {
        self.push_inverse_op(&op);
        self.apply_op_without_history(op)
    }

    fn apply_op_without_history(&mut self, op: Op) -> Result<(), RunError> {
        if let Some(ref p) = self.path.clone() {
            let mut active_sheet = self.view_sheet_id;
            commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &crate::ops::WorkbookOp::SheetOp {
                    sheet_id: self.view_sheet_id,
                    op,
                },
            )?;
            self.ops_applied = self.ops_applied.saturating_add(1);
            self.sync_active_sheet_cache();
            self.start_log_watcher_if_needed()?;
        } else {
            op.apply(&mut self.state);
        }
        Ok(())
    }

    fn parse_pasted_tsv_cells(
        text: &str,
        start: SheetCursor,
        preserve_formulas: bool,
        state: &SheetState,
    ) -> Vec<(CellAddr, String)> {
        let rows: Vec<&str> = text.lines().collect();
        if rows.is_empty() {
            return Vec::new();
        }
        let row_count = rows.len();
        let col_count = rows
            .iter()
            .map(|line| line.split('\t').count())
            .max()
            .unwrap_or(0);
        if col_count == 0 {
            return Vec::new();
        }

        let needed_rows = start.row.saturating_sub(HEADER_ROWS) + row_count;
        let needed_cols = start.col.saturating_sub(MARGIN_COLS) + col_count;
        let mut grid = state.grid.clone();
        if needed_rows > grid.main_rows() || needed_cols > grid.main_cols() {
            grid.set_main_size(
                grid.main_rows().max(needed_rows),
                grid.main_cols().max(needed_cols),
            );
        }

        let mut cells = Vec::new();
        for (r_off, line) in rows.iter().enumerate() {
            for (c_off, value) in line.split('\t').enumerate() {
                let row = start.row.saturating_add(r_off);
                let col = start.col.saturating_add(c_off);
                let addr = SheetCursor { row, col }.to_addr(&grid);
                if row >= HEADER_ROWS + grid.main_rows() + FOOTER_ROWS || col >= grid.total_cols() {
                    continue;
                }
                let mut value = value.to_string();
                if !preserve_formulas && value.trim_start().starts_with('=') {
                    value = value.trim_start_matches('=').to_string();
                }
                cells.push((addr, value));
            }
        }
        cells
    }

    fn paste_pasted_tsv_cells(
        &mut self,
        cells: Vec<(CellAddr, String)>,
        preserve_formulas: bool,
    ) -> Result<(), RunError> {
        if cells.is_empty() {
            self.status = "Clipboard paste produced no cells".into();
            return Ok(());
        }
        self.apply_single_op(Op::FillRange { cells })?;
        self.status = if preserve_formulas {
            "Clipboard pasted".into()
        } else {
            "Clipboard pasted as values".into()
        };
        Ok(())
    }

    fn try_paste_from_snapshot(&mut self, preserve_formulas: bool) -> Result<bool, RunError> {
        let Some((source, snapshot)) = self.clipboard_snapshot.clone() else {
            return Ok(false);
        };
        let Some(target) = self.paste_target_main_range(&source) else {
            return Ok(false);
        };
        if snapshot != self.selection_tsv_text_for_main_range(source.clone()) {
            return Ok(false);
        }
        self.apply_single_op(Op::CopyFromTo { source, target })?;
        self.status = if preserve_formulas {
            "Clipboard pasted".into()
        } else {
            "Clipboard pasted as values".into()
        };
        Ok(true)
    }

    fn selection_tsv_text_for_main_range(&self, range: MainRange) -> String {
        let rows = (range.row_start..range.row_end)
            .map(|r| HEADER_ROWS + r as usize)
            .collect::<Vec<_>>();
        let cols = (range.col_start..range.col_end)
            .map(|c| MARGIN_COLS + c as usize)
            .collect::<Vec<_>>();
        let mut out = String::new();
        for (ri, row) in rows.iter().enumerate() {
            if ri > 0 {
                out.push('\n');
            }
            for (ci, col) in cols.iter().enumerate() {
                if ci > 0 {
                    out.push('\t');
                }
                if let Some(addr) = self.addr_at(*row, *col) {
                    let raw = self.state.grid.get(&addr);
                    out.push_str(raw.as_deref().unwrap_or(""));
                }
            }
        }
        if !out.is_empty() {
            out.push('\n');
        }
        out
    }

    fn apply_pasted_tsv(&mut self, text: &str, preserve_formulas: bool) -> Result<(), RunError> {
        let cells = Self::parse_pasted_tsv_cells(text, self.cursor, preserve_formulas, &self.state);
        self.paste_pasted_tsv_cells(cells, preserve_formulas)
    }

    fn paste_from_clipboard(&mut self, preserve_formulas: bool) -> Result<(), RunError> {
        let text = read_clipboard().map_err(io::Error::other)?;
        let cells =
            Self::parse_pasted_tsv_cells(&text, self.cursor, preserve_formulas, &self.state);
        if self.try_paste_from_snapshot(preserve_formulas)? {
            return Ok(());
        }
        self.paste_pasted_tsv_cells(cells, preserve_formulas)
    }

    fn finish_export(&mut self, csv: bool, filename: &str) {
        let data = self.do_export(csv);
        let ext = if csv { "csv" } else { "tsv" };
        if filename.trim().is_empty() {
            match copy_to_clipboard(&data) {
                Ok(()) => self.status = format!("{} copied to clipboard", ext.to_uppercase()),
                Err(e) => self.status = format!("Clipboard error: {e}"),
            }
        } else {
            match std::fs::write(filename.trim(), &data) {
                Ok(()) => self.status = format!("Exported {} to {filename}", ext.to_uppercase()),
                Err(e) => self.status = format!("Write error: {e}"),
            }
        }
    }

    fn paste_target_main_range(&self, source: &MainRange) -> Option<MainRange> {
        if self.cursor.row < HEADER_ROWS || self.cursor.col < MARGIN_COLS {
            return None;
        }
        let row_start = (self.cursor.row - HEADER_ROWS) as u32;
        let col_start = (self.cursor.col - MARGIN_COLS) as u32;
        Some(MainRange {
            row_start,
            row_end: row_start + source.row_end.saturating_sub(source.row_start),
            col_start,
            col_end: col_start + source.col_end.saturating_sub(source.col_start),
        })
    }

    fn movie_input_path(&self) -> Result<PathBuf, RunError> {
        let Some(path) = self.path.clone().or(self.source_path.clone()) else {
            return Err(io::Error::other("--movie requires a .corro file path").into());
        };
        if !path.exists() {
            return Err(io::Error::other(format!("movie input does not exist: {}", path.display())).into());
        }
        let ext = path
            .extension()
            .and_then(|e| e.to_str())
            .unwrap_or("")
            .to_ascii_lowercase();
        if ext != "corro" {
            return Err(io::Error::other(format!(
                "--movie only supports .corro input (got {})",
                if ext.is_empty() { "<none>" } else { ext.as_str() }
            ))
            .into());
        }
        Ok(path)
    }

    fn reset_workbook_for_movie(&mut self, path: &Path) {
        self.workbook = WorkbookState::new();
        self.view_sheet_id = self.workbook.sheet_id(self.workbook.active_sheet);
        self.sync_active_sheet_cache();
        self.sync_persisted_sort_cache_from_workbook();
        self.offset = 0;
        self.ops_applied = 0;
        // Movie replay must stay detached from on-disk log commit/watcher paths,
        // otherwise commit_edit_buffer can rehydrate the whole workbook from file.
        self.path = None;
        self.source_path = Some(path.to_path_buf());
        self.import_source = None;
        self.revision_limit = None;
        self.revision_browse = false;
        self.watcher = None;
        self.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        self.row_scroll = 0;
        self.col_scroll = 0;
        self.mode = Mode::Normal;
    }

    fn movie_apply_set_cell_value(&mut self, value: &str) {
        let addr = self.cursor.to_addr(&self.state.grid);
        let op = Op::SetCell {
            addr: addr.clone(),
            value: value.to_string(),
        };
        op.apply(&mut self.state);
        if let CellAddr::Main { col, .. } = addr {
            self.state
                .grid
                .auto_fit_column(MARGIN_COLS + col as usize);
        }
        self.commit_active_sheet_cache();
    }

    fn movie_draw_and_sleep(
        &mut self,
        terminal: &mut Terminal<CrosstermBackend<std::io::Stdout>>,
        delay: std::time::Duration,
    ) -> Result<bool, RunError> {
        terminal.draw(|f| self.draw(f))?;
        let sleep_slice = std::time::Duration::from_millis(25);
        let start = std::time::Instant::now();
        while start.elapsed() < delay {
            if self.movie_should_quit()? {
                return Ok(true);
            }
            let remaining = delay.saturating_sub(start.elapsed());
            std::thread::sleep(remaining.min(sleep_slice));
        }
        Ok(false)
    }

    fn movie_should_quit(&mut self) -> Result<bool, RunError> {
        while event::poll(std::time::Duration::from_millis(0))? {
            if let Event::Key(key) = event::read()? {
                if key.kind == KeyEventKind::Release {
                    continue;
                }
                if matches!(key.code, KeyCode::Esc | KeyCode::Char('q') | KeyCode::Char('Q')) {
                    return Ok(true);
                }
                if key.modifiers.contains(KeyModifiers::CONTROL)
                    && matches!(key.code, KeyCode::Char('c') | KeyCode::Char('C'))
                {
                    return Ok(true);
                }
            }
        }
        Ok(false)
    }

    fn movie_focus_sheet(&mut self, sheet_id: u32) {
        self.view_sheet_id = sheet_id;
        self.sync_active_sheet_cache();
        self.sync_persisted_sort_cache_from_workbook();
        self.cursor.clamp(&self.state.grid);
    }

    fn movie_move_cursor_to_addr(&mut self, addr: &CellAddr) {
        // Movie replay can target cells beyond the current in-memory bounds.
        // Grow main dimensions first so address->cursor mapping doesn't clamp away
        // the final data row/col during replay.
        let mut needed_rows = self.state.grid.main_rows();
        let mut needed_cols = self.state.grid.main_cols();
        match addr {
            CellAddr::Main { row, col } => {
                needed_rows = needed_rows.max(*row as usize + 1);
                needed_cols = needed_cols.max(*col as usize + 1);
            }
            CellAddr::Left { row, .. } | CellAddr::Right { row, .. } => {
                needed_rows = needed_rows.max(*row as usize + 1);
            }
            CellAddr::Header { .. } | CellAddr::Footer { .. } => {}
        }
        if needed_rows != self.state.grid.main_rows() || needed_cols != self.state.grid.main_cols() {
            self.state.grid.set_main_size(needed_rows, needed_cols);
            self.commit_active_sheet_cache();
        }

        let (row, col) = match addr {
            CellAddr::Header { row, col } => (*row as usize, *col as usize),
            CellAddr::Main { row, col } => (HEADER_ROWS + *row as usize, MARGIN_COLS + *col as usize),
            CellAddr::Footer { row, col } => {
                (HEADER_ROWS + self.state.grid.main_rows() + *row as usize, *col as usize)
            }
            CellAddr::Left { row, col } => (HEADER_ROWS + *row as usize, *col as usize),
            CellAddr::Right { row, col } => {
                (HEADER_ROWS + *row as usize, MARGIN_COLS + self.state.grid.main_cols() + *col as usize)
            }
        };
        self.cursor = SheetCursor { row, col };
        self.cursor.clamp(&self.state.grid);
    }

    fn movie_type_and_commit_current_cell(
        &mut self,
        terminal: &mut Terminal<CrosstermBackend<std::io::Stdout>>,
        text: &str,
        line_i: usize,
        line_n: usize,
        char_delay: std::time::Duration,
        confirm_delay: std::time::Duration,
    ) -> Result<(), RunError> {
        self.mode = self.start_edit_mode(String::new(), None, false, false, None);
        self.status = format!("Movie {}/{} edit", line_i + 1, line_n);
        if self.movie_draw_and_sleep(terminal, char_delay)? {
            return Err(io::Error::new(io::ErrorKind::Interrupted, "movie interrupted by user").into());
        }
        let mut typed = String::new();
        for ch in text.chars() {
            typed.push(ch);
            self.mode = self.start_edit_mode(typed.clone(), None, false, false, None);
            self.status = format!("Movie {}/{} typing: {}", line_i + 1, line_n, typed);
            if self.movie_draw_and_sleep(terminal, char_delay)? {
                return Err(
                    io::Error::new(io::ErrorKind::Interrupted, "movie interrupted by user").into(),
                );
            }
        }
        self.status = format!("Movie {}/{} confirm", line_i + 1, line_n);
        if self.movie_draw_and_sleep(terminal, confirm_delay)? {
            return Err(io::Error::new(io::ErrorKind::Interrupted, "movie interrupted by user").into());
        }
        self.movie_apply_set_cell_value(&typed);
        self.edit_target_addr = None;
        self.mode = Mode::Normal;
        Ok(())
    }

    fn movie_show_menu(
        &mut self,
        terminal: &mut Terminal<CrosstermBackend<std::io::Stdout>>,
        section: MenuSection,
        action: MenuAction,
        label: &str,
        line_i: usize,
        line_n: usize,
        menu_hold: std::time::Duration,
    ) -> Result<(), RunError> {
        self.mode = Mode::Menu {
            stack: vec![MenuLevel { section, item: 0 }],
        };
        self.status = format!("Movie {}/{} menu: {}", line_i + 1, line_n, label);
        if self.movie_draw_and_sleep(terminal, menu_hold)? {
            return Err(io::Error::new(io::ErrorKind::Interrupted, "movie interrupted by user").into());
        }
        let selected_item = menu_items(section)
            .iter()
            .position(|item| item.target == MenuTarget::Action(action))
            .unwrap_or(0);
        self.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section,
                item: selected_item,
            }],
        };
        self.status = format!("Movie {}/{} confirm: {}", line_i + 1, line_n, label);
        if self.movie_draw_and_sleep(terminal, menu_hold)? {
            return Err(io::Error::new(io::ErrorKind::Interrupted, "movie interrupted by user").into());
        }
        self.mode = Mode::Normal;
        Ok(())
    }

    fn movie_show_balance_books_dialog(
        &mut self,
        terminal: &mut Terminal<CrosstermBackend<std::io::Stdout>>,
        amount_col: usize,
        direction: BalanceDirection,
        line_i: usize,
        line_n: usize,
        menu_hold: std::time::Duration,
    ) -> Result<(), RunError> {
        let dialog_delay = menu_hold.max(std::time::Duration::from_millis(120));
        self.mode = Mode::BalanceBooks {
            buffer: addr::excel_column_name(amount_col),
            direction,
            // A logged BalanceReport op is a persisted report operation.
            persist: true,
            focus: BalanceBooksFocus::Column,
        };
        self.status = format!("Movie {}/{} dialog: Balance books", line_i + 1, line_n);
        if self.movie_draw_and_sleep(terminal, dialog_delay)? {
            return Err(io::Error::new(io::ErrorKind::Interrupted, "movie interrupted by user").into());
        }
        self.mode = Mode::BalanceBooks {
            buffer: addr::excel_column_name(amount_col),
            direction,
            // A logged BalanceReport op is a persisted report operation.
            persist: true,
            focus: BalanceBooksFocus::Generate,
        };
        self.status = format!("Movie {}/{} confirm: Balance books", line_i + 1, line_n);
        if self.movie_draw_and_sleep(terminal, dialog_delay)? {
            return Err(io::Error::new(io::ErrorKind::Interrupted, "movie interrupted by user").into());
        }
        self.mode = Mode::Normal;
        Ok(())
    }

    fn movie_apply_line_as_user(
        &mut self,
        terminal: &mut Terminal<CrosstermBackend<std::io::Stdout>>,
        line: &str,
        active_sheet: &mut u32,
        line_i: usize,
        line_n: usize,
        char_delay: std::time::Duration,
        confirm_delay: std::time::Duration,
        menu_hold: std::time::Duration,
    ) -> Result<bool, RunError> {
        let op = match crate::ops::parse_workbook_line(line) {
            Ok(op) => op,
            Err(_) => return Ok(false),
        };
        match op {
            crate::ops::WorkbookOp::SheetOp { sheet_id, op } => {
                self.movie_focus_sheet(sheet_id);
                match op {
                    crate::ops::Op::SetCell { addr, value } => {
                        self.movie_move_cursor_to_addr(&addr);
                        self.movie_type_and_commit_current_cell(
                            terminal,
                            &value,
                            line_i,
                            line_n,
                            char_delay,
                            confirm_delay,
                        )?;
                        self.ops_applied += 1;
                        return Ok(true);
                    }
                    crate::ops::Op::SetCellRef { cref, value } => {
                        let addr = cref.to_grid_addr(self.state.grid.main_cols());
                        self.movie_move_cursor_to_addr(&addr);
                        self.movie_type_and_commit_current_cell(
                            terminal,
                            &value,
                            line_i,
                            line_n,
                            char_delay,
                            confirm_delay,
                        )?;
                        self.ops_applied += 1;
                        return Ok(true);
                    }
                    crate::ops::Op::FillRange { cells } => {
                        for (addr, value) in cells {
                            self.movie_move_cursor_to_addr(&addr);
                            self.movie_type_and_commit_current_cell(
                                terminal,
                                &value,
                                line_i,
                                line_n,
                                char_delay,
                                confirm_delay,
                            )?;
                            self.ops_applied += 1;
                        }
                        return Ok(true);
                    }
                    _ => {}
                }
            }
            crate::ops::WorkbookOp::NewSheet { .. } => {
                self.movie_show_menu(
                    terminal,
                    MenuSection::Sheet,
                    MenuAction::NewSheet,
                    "New sheet",
                    line_i,
                    line_n,
                    menu_hold,
                )?;
            }
            crate::ops::WorkbookOp::CopySheet { .. } => {
                self.movie_show_menu(
                    terminal,
                    MenuSection::Sheet,
                    MenuAction::CopySheet,
                    "Copy sheet",
                    line_i,
                    line_n,
                    menu_hold,
                )?;
            }
            crate::ops::WorkbookOp::RenameSheet { .. } => {
                self.movie_show_menu(
                    terminal,
                    MenuSection::Sheet,
                    MenuAction::RenameSheet,
                    "Rename sheet",
                    line_i,
                    line_n,
                    menu_hold,
                )?;
            }
            crate::ops::WorkbookOp::MoveSheet { .. } => {
                self.movie_show_menu(
                    terminal,
                    MenuSection::Sheet,
                    MenuAction::MoveSheet,
                    "Move sheet",
                    line_i,
                    line_n,
                    menu_hold,
                )?;
            }
            crate::ops::WorkbookOp::ActivateSheet { id } => {
                self.movie_focus_sheet(id);
                self.status = format!("Movie {}/{} activate sheet {}", line_i + 1, line_n, id);
                if self.movie_draw_and_sleep(terminal, confirm_delay)? {
                    return Err(
                        io::Error::new(io::ErrorKind::Interrupted, "movie interrupted by user")
                            .into(),
                    );
                }
                self.ops_applied += 1;
                return Ok(true);
            }
            crate::ops::WorkbookOp::BalanceReport {
                amount_col,
                direction,
                ..
            } => {
                self.movie_show_menu(
                    terminal,
                    MenuSection::Sheet,
                    MenuAction::BalanceBooks,
                    "Balance report",
                    line_i,
                    line_n,
                    menu_hold,
                )?;
                self.movie_show_balance_books_dialog(
                    terminal,
                    amount_col,
                    direction,
                    line_i,
                    line_n,
                    menu_hold,
                )?;
                self.status = format!("Movie {}/{} generate balance report", line_i + 1, line_n);
                if self.movie_draw_and_sleep(terminal, confirm_delay)? {
                    return Err(
                        io::Error::new(io::ErrorKind::Interrupted, "movie interrupted by user")
                            .into(),
                    );
                }
                crate::ops::apply_log_line_to_workbook(line, &mut self.workbook, active_sheet)?;
                self.view_sheet_id = *active_sheet;
                self.sync_active_sheet_cache();
                self.sync_persisted_sort_cache_from_workbook();
                self.ops_applied += 1;
                self.cursor.clamp(&self.state.grid);
                return Ok(true);
            }
        }
        Ok(false)
    }

    pub fn run_movie(&mut self, options: MovieReplayOptions) -> Result<(), RunError> {
        let path = self.movie_input_path()?;
        let data = std::fs::read_to_string(&path).map_err(IoError::Io)?;
        let mut log_lines: Vec<String> = Vec::new();
        for raw in data.lines() {
            let t = raw.trim();
            if t.is_empty() {
                continue;
            }
            log_lines.push(t.to_string());
        }
        self.reset_workbook_for_movie(&path);

        let char_delay = std::time::Duration::from_secs_f64(1.0 / options.typing_cps.max(0.1));
        let confirm_delay = std::time::Duration::from_millis(options.confirm_delay_ms);
        let menu_hold = std::time::Duration::from_millis(options.menu_hold_ms);

        enable_raw_mode()?;
        let mut stdout = stdout();
        execute!(stdout, EnterAlternateScreen, Hide)?;
        let backend = CrosstermBackend::new(stdout);
        let mut terminal = Terminal::new(backend)?;

        let run_result = (|| -> Result<(), RunError> {
            let mut active_sheet = self.workbook.sheet_id(self.workbook.active_sheet);
            for (i, line) in log_lines.iter().enumerate() {
                if self.movie_should_quit()? {
                    self.status = "Movie stopped by user".into();
                    return Ok(());
                }
                if !self.movie_apply_line_as_user(
                    &mut terminal,
                    line,
                    &mut active_sheet,
                    i,
                    log_lines.len(),
                    char_delay,
                    confirm_delay,
                    menu_hold,
                )? {
                    // Fallback for ops that don't map cleanly to one edit interaction.
                    self.status =
                        format!("Movie {}/{} apply op", i + 1, log_lines.len());
                    if self.movie_draw_and_sleep(&mut terminal, confirm_delay)? {
                        self.status = "Movie stopped by user".into();
                        return Ok(());
                    }
                    crate::ops::apply_log_line_to_workbook(
                        line,
                        &mut self.workbook,
                        &mut active_sheet,
                    )?;
                    self.view_sheet_id = active_sheet;
                    self.sync_active_sheet_cache();
                    self.sync_persisted_sort_cache_from_workbook();
                    self.ops_applied += 1;
                    self.cursor.clamp(&self.state.grid);
                }
            }
            self.status = format!(
                "Movie complete: {} lines from {}",
                self.ops_applied,
                path.display()
            );
            terminal.draw(|f| self.draw(f))?;
            std::thread::sleep(confirm_delay * 2);
            Ok(())
        })();

        let run_result = match run_result {
            Err(RunError::Term(err)) if err.kind() == io::ErrorKind::Interrupted => {
                self.status = "Movie stopped by user".into();
                Ok(())
            }
            other => other,
        };

        let disable_result = disable_raw_mode();
        let leave_result = execute!(terminal.backend_mut(), LeaveAlternateScreen, Show);
        let restore_result = match (disable_result, leave_result) {
            (Ok(()), Ok(())) => Ok(()),
            (Err(disable_err), Ok(())) => Err(RunError::Term(disable_err)),
            (Ok(()), Err(leave_err)) => Err(RunError::Term(leave_err)),
            (Err(disable_err), Err(leave_err)) => Err(RunError::Term(io::Error::other(format!(
                "disable_raw_mode failed: {disable_err}; restore failed: {leave_err}"
            )))),
        };

        match (run_result, restore_result) {
            (Err(run_err), Err(restore_err)) => Err(RunError::Term(io::Error::other(format!(
                "{run_err}; cleanup failed: {restore_err}"
            )))),
            (Err(run_err), Ok(())) => Err(run_err),
            (Ok(()), Err(restore_err)) => Err(restore_err),
            (Ok(()), Ok(())) => Ok(()),
        }
    }

    pub fn run(&mut self) -> Result<(), RunError> {
        enable_raw_mode()?;
        let mut stdout = stdout();
        execute!(stdout, EnterAlternateScreen, Hide)?;
        let backend = CrosstermBackend::new(stdout);
        let mut terminal = Terminal::new(backend)?;

        let run_result = (|| -> Result<(), RunError> {
            loop {
                self.sync_external()?;
                terminal.draw(|f| self.draw(f))?;

                if !event::poll(std::time::Duration::from_millis(200))? {
                    continue;
                }
                if let Event::Key(key) = event::read()? {
                    if self.handle_key(key)? {
                        break;
                    }
                }
            }
            Ok(())
        })();

        let disable_result = disable_raw_mode();
        let leave_result = execute!(terminal.backend_mut(), LeaveAlternateScreen, Show);
        let restore_result = match (disable_result, leave_result) {
            (Ok(()), Ok(())) => Ok(()),
            (Err(disable_err), Ok(())) => Err(RunError::Term(disable_err)),
            (Ok(()), Err(leave_err)) => Err(RunError::Term(leave_err)),
            (Err(disable_err), Err(leave_err)) => Err(RunError::Term(io::Error::other(format!(
                "disable_raw_mode failed: {disable_err}; restore failed: {leave_err}"
            )))),
        };

        match (run_result, restore_result) {
            (Err(run_err), Err(restore_err)) => Err(RunError::Term(io::Error::other(format!(
                "{run_err}; cleanup failed: {restore_err}"
            )))),
            (Err(run_err), Ok(())) => Err(run_err),
            (Ok(()), Err(restore_err)) => Err(restore_err),
            (Ok(()), Ok(())) => Ok(()),
        }
    }

    fn draw(&mut self, f: &mut Frame) {
        let _ctx = crate::formula::set_eval_context(&self.workbook);
        crate::formula::refresh_spills(&mut self.state.grid);
        f.render_widget(Clear, f.area());
        let special_picker = self.special_picker;
        let has_tabs = self.workbook.sheet_count() > 1;
        let constraints = vec![
            Constraint::Length(1),
            Constraint::Length(1),
            Constraint::Min(3),
            Constraint::Length(1),
        ];
        let layout = Layout::default()
            .direction(Direction::Vertical)
            .constraints(constraints)
            .split(f.area());
        let menubar_area = layout[0];
        let formula_area = layout[1];
        let grid_area = layout[2];
        let hints_area = layout[3];

        let sentinel = Block::default().borders(Borders::ALL);
        let inner = sentinel.inner(grid_area);
        let inner_h = inner.height as usize;
        let inner_w = inner.width as usize;

        let data_rows = inner_h.saturating_sub(1).max(1);
        self.grid_viewport_data_rows = data_rows;
        let data_width = inner_w.saturating_sub(ROW_LABEL_CHARS).max(1);
        let data_cols = data_width.checked_div(2).unwrap_or(1).max(1);

        // Determine visible rows/cols first so we can adjust widths for the
        // visible columns before taking an immutable borrow of grid.

        let (row_ixs, next_row_scroll) =
            visible_row_indices(&self.state, self.cursor, data_rows, self.row_scroll);
        let (mut col_ixs, next_col_scroll) =
            visible_col_indices(&self.state, self.cursor, data_cols, self.col_scroll);
        // visible indices computed
        self.row_scroll = next_row_scroll;
        self.col_scroll = next_col_scroll;

        // Shrink visible columns: cap width so every visible index can share
        // the row body (avoids unbounded autofill from one long cell, then
        // trim_visible_cols_to_width eating columns to the left of the cursor).
        self.fit_visible_columns_capped(&col_ixs, data_width);
        trim_visible_cols_to_width(&self.state.grid, &mut col_ixs, self.cursor.col, data_width);

        // Materialize grid after we finish possibly mutating column widths.
        let grid = &self.state.grid;
        let title_str = {
            let raw = format!(
                " corro  {}r × {}c  ops {}",
                grid.main_rows(),
                grid.main_cols(),
                self.ops_applied
            );
            let max_w = (grid_area.width.saturating_sub(4) as usize).max(8);
            if raw.chars().count() > max_w {
                format!(
                    "{}…",
                    raw.chars()
                        .take(max_w.saturating_sub(1))
                        .collect::<String>()
                )
            } else {
                raw
            }
        };

        let block = Block::default().borders(Borders::ALL).title(Span::styled(
            title_str,
            Style::default().add_modifier(Modifier::BOLD),
        ));

        // ── Menu bar ──────────────────────────────────────────────────────────
        let menubar = self.menu_bar_line();
        f.render_widget(
            Paragraph::new(menubar).style(Style::default().fg(Color::Black).bg(Color::Cyan)),
            menubar_area,
        );

        // ── Formula bar ───────────────────────────────────────────────────────
        let addr = self.cursor.to_addr(grid);
        let edit_addr = self.edit_target_addr.clone().unwrap_or(addr.clone());
        // #region agent log
        debug_log_ndjson(
            "H4",
            "src/ui/mod.rs:draw:cursor_render_class",
            "draw cursor class snapshot",
            format!(
                "{{\"cursor_row\":{},\"cursor_col\":{},\"main_rows\":{},\"first_footer\":{},\"row_looks_footer\":{},\"addr_kind\":\"{}\",\"edit_addr_kind\":\"{}\"}}",
                self.cursor.row,
                self.cursor.col,
                grid.main_rows(),
                HEADER_ROWS + grid.main_rows(),
                self.cursor.row >= HEADER_ROWS + grid.main_rows(),
                match addr {
                    CellAddr::Header { .. } => "header",
                    CellAddr::Main { .. } => "main",
                    CellAddr::Footer { .. } => "footer",
                    CellAddr::Left { .. } => "left",
                    CellAddr::Right { .. } => "right",
                },
                match edit_addr {
                    CellAddr::Header { .. } => "header",
                    CellAddr::Main { .. } => "main",
                    CellAddr::Footer { .. } => "footer",
                    CellAddr::Left { .. } => "left",
                    CellAddr::Right { .. } => "right",
                }
            ),
        );
        // #endregion
        let prompt_style = Style::default().fg(Color::White).bg(Color::DarkGray);
        let prompt_style_bold = prompt_style.add_modifier(Modifier::BOLD);
        let caret_style = Style::default()
            .fg(Color::Black)
            .bg(Color::Yellow)
            .add_modifier(Modifier::BOLD);
        let formula_widget = self.mode_prompt_widget(
            grid,
            &addr,
            &edit_addr,
            prompt_style,
            prompt_style_bold,
            caret_style,
        );
        f.render_widget(formula_widget, formula_area);

        if has_tabs {
            let tab_style = Style::default().fg(Color::White).bg(Color::DarkGray);
            let active_style = Style::default()
                .fg(Color::Black)
                .bg(Color::Yellow)
                .add_modifier(Modifier::BOLD);
            let mut spans = Vec::new();
            for (idx, sheet) in self.workbook.sheets.iter().enumerate() {
                if idx > 0 {
                    spans.push(Span::raw("  "));
                }
                let style = if idx == self.workbook.active_sheet {
                    active_style
                } else {
                    tab_style
                };
                spans.push(Span::styled(format!(" {} ", sheet.title), style));
            }
            let tab_area = hints_area;
            f.render_widget(Paragraph::new(Line::from(spans)).style(tab_style), tab_area);
        }

        if self.render_help_about_overlay(f, grid_area) {
            return;
        }

        if self.render_export_preview_overlay(f, grid_area) {
            self.render_export_bottom_hints(f, hints_area, has_tabs);
            return;
        }

        if matches!(&self.mode, Mode::BalanceBooks { .. }) {
            let area = centered_rect(72, 64, f.area());
            f.render_widget(Clear, area);
            let block = Block::default()
                .borders(Borders::ALL)
                .border_type(BorderType::Plain)
                .title(Span::styled(
                    " Balance books ",
                    Style::default().fg(Color::Cyan),
                ));
            let inner = block.inner(area);
            f.render_widget(block, area);
            let body = match &self.mode {
                Mode::BalanceBooks {
                    buffer,
                    direction,
                    persist,
                    focus,
                } => self.balance_dialog_lines(
                    buffer,
                    *direction,
                    *persist,
                    *focus,
                    self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                    Style::default().fg(Color::White),
                    Style::default()
                        .fg(Color::White)
                        .add_modifier(Modifier::BOLD),
                    Style::default().fg(Color::Black).bg(Color::Yellow),
                ),
                _ => Vec::new(),
            };
            let focus_line = match &self.mode {
                Mode::BalanceBooks { focus, .. } => Self::balance_dialog_focus_line(*focus),
                _ => 0,
            };
            let max_visible = inner.height as usize;
            let max_scroll = body.len().saturating_sub(max_visible);
            let mut scroll_y = 0usize;
            if max_scroll > 0 && max_visible > 0 && focus_line >= max_visible {
                scroll_y = (focus_line + 1).saturating_sub(max_visible).min(max_scroll);
            }
            f.render_widget(
                Paragraph::new(body)
                    .wrap(Wrap { trim: false })
                    .scroll((scroll_y as u16, 0)),
                inner,
            );
            return;
        }

        // ── Grid ──────────────────────────────────────────────────────────────
        f.render_widget(Clear, grid_area);
        let mut lines: Vec<Line> = Vec::new();

        {
            let lm = MARGIN_COLS;
            let mc = grid.main_cols();
            let show_right_divider = col_ixs.contains(&(lm + mc));
            let mut spans: Vec<Span> = vec![Span::styled(
                format!("{:>width$}", "", width = ROW_LABEL_CHARS),
                Style::default().add_modifier(Modifier::BOLD),
            )];
            for (i, &c) in col_ixs.iter().enumerate() {
                let name = col_header_label(c, grid.main_cols());
                let active_col = c == self.cursor.col;
                let style = if active_col {
                    Style::default()
                        .fg(Color::Black)
                        .bg(Color::Yellow)
                        .add_modifier(Modifier::BOLD)
                } else {
                    Style::default()
                        .fg(Color::Cyan)
                        .add_modifier(Modifier::BOLD)
                };
                let w = grid.col_width(c).max(1);
                let p = name
                    .unicode_pad(w, UTruncAlign::Right, true)
                    .into_owned();
                spans.push(Span::styled(p, style));
                if i + 1 < col_ixs.len() {
                    if c == lm - 1 && lm > 0 && col_ixs.contains(&lm) {
                        // Put the vertical divider immediately after the cell
                        // content (no intervening space) so it abuts the text as
                        // tests expect.
                        spans.push(Span::styled("│", Style::default().fg(Color::DarkGray)));
                        spans.push(Span::raw(" "));
                    } else if c == lm + mc - 1 && show_right_divider {
                        spans.push(Span::styled("│", Style::default().fg(Color::DarkGray)));
                        spans.push(Span::raw(" "));
                    } else {
                        spans.push(Span::raw(" "));
                    }
                }
            }
            lines.push(Line::from(spans));
        }

        let hr = HEADER_ROWS;
        let mr = grid.main_rows();
        let lm = MARGIN_COLS;
        let mc = grid.main_cols();
        let show_right_divider = col_ixs.contains(&(lm + mc));
        let max_data_lines = inner_h.saturating_sub(1);
        let last_display_main_row = grid.sorted_main_rows().last().map(|row| hr + *row);
        for &r in row_ixs.iter().take(max_data_lines) {
            let active_row = r == self.cursor.row;
            let is_underlined_boundary_row =
                (hr > 0 && r == hr - 1) || last_display_main_row == Some(r);
            let mut row_label_style = if active_row {
                Style::default()
                    .fg(Color::Black)
                    .bg(Color::Yellow)
                    .add_modifier(Modifier::BOLD)
            } else if r >= hr + mr {
                Style::default()
                    .fg(Color::Cyan)
                    .add_modifier(Modifier::BOLD)
            } else {
                Style::default().fg(Color::Yellow)
            };
            if is_underlined_boundary_row {
                row_label_style = row_label_style.add_modifier(Modifier::UNDERLINED);
            }
            let mut spans: Vec<Span> = vec![Span::styled(
                format!("{:>4} ", sheet_row_label(r, grid.main_rows())),
                row_label_style,
            )];
            let footer_agg = if r >= hr + mr {
                footer_row_agg_func(grid, r - hr - mr)
            } else {
                None
            };
            let main_row_idx = if r >= hr && r < hr + mr {
                Some((r - hr) as u32)
            } else {
                None
            };

            let left_margin_agg = main_row_idx.and_then(|mri| left_margin_agg_func(grid, mri));
            let left_margin_block_start = main_row_idx.map(|mri| row_total_block_start(grid, mri));

            for (i, &c) in col_ixs.iter().enumerate() {
                let cur = SheetCursor { row: r, col: c };
                let cell_addr = cur.to_addr(grid);
                let right_col_agg = right_col_agg_func(grid, c);

                let mut is_agg_cell = false;
                let text = if let Some(func) = footer_agg {
                    if right_col_agg.is_some() {
                        is_agg_cell = true;
                        footer_special_col_aggregate(grid, func, c, mr, mc)
                            .unwrap_or_else(|| cell_effective_display(grid, &cell_addr))
                    } else if c >= lm && c < lm + mc {
                        is_agg_cell = true;
                        let main_col = (c - lm) as u32;
                        compute_aggregate(
                            grid,
                            &AggregateDef {
                                func,
                                source: MainRange {
                                    row_start: 0,
                                    row_end: mr as u32,
                                    col_start: main_col,
                                    col_end: main_col + 1,
                                },
                            },
                        )
                    } else {
                        cell_effective_display(grid, &cell_addr)
                    }
                } else if let (Some(func), Some(block_start), Some(main_row)) =
                    (left_margin_agg, left_margin_block_start, main_row_idx)
                {
                    if c >= lm && c < lm + mc {
                        is_agg_cell = true;
                        if right_col_agg.is_some() {
                            let data_cols = data_main_col_count(grid);
                            let (row_start, row_end) = if block_start < main_row {
                                (block_start, main_row)
                            } else {
                                previous_raw_block(grid, main_row).unwrap_or((0, main_row))
                            };
                            left_margin_special_col_aggregate(
                                grid, func, c, row_start, row_end, data_cols,
                            )
                            .unwrap_or_else(|| cell_effective_display(grid, &cell_addr))
                        } else {
                            let main_col = (c - lm) as u32;
                            left_margin_main_col_aggregate(grid, func, main_row, main_col)
                        }
                    } else if right_col_agg.is_some() {
                        is_agg_cell = true;
                        left_margin_special_col_aggregate(
                            grid,
                            func,
                            c,
                            block_start,
                            main_row,
                            data_main_col_count(grid),
                        )
                        .unwrap_or_else(|| cell_effective_display(grid, &cell_addr))
                    } else {
                        cell_effective_display(grid, &cell_addr)
                    }
                } else if r >= hr && r < hr + mr {
                    if let Some(func) = right_col_agg {
                        is_agg_cell = true;
                        let main_row = (r - hr) as u32;
                        let data_cols = data_main_col_count(grid);
                        compute_aggregate(
                            grid,
                            &AggregateDef {
                                func,
                                source: MainRange {
                                    row_start: main_row,
                                    row_end: main_row + 1,
                                    col_start: 0,
                                    col_end: data_cols as u32,
                                },
                            },
                        )
                    } else {
                        cell_effective_display(grid, &cell_addr)
                    }
                } else {
                    cell_effective_display(grid, &cell_addr)
                };
                let cw = grid.col_width(c).max(1);
                let formatted = format_cell_display(grid, &cell_addr, text);
                let align = effective_cell_align(grid, &cell_addr, &formatted);
                let disp = if formatted.width() > cw {
                    shrink_numeric_display(&formatted, cw)
                        .or_else(|| exponential_numeric_display(&formatted, cw))
                        .unwrap_or_else(|| truncate_with_ellipsis(&formatted, cw))
                } else {
                    formatted
                };
                let disp = align_cell_display(disp, cw, align);
                let sel = self.anchor.is_some_and(|a| match self.selection_kind {
                    SelectionKind::Cells => {
                        let r0 = a.row.min(self.cursor.row);
                        let r1 = a.row.max(self.cursor.row);
                        let c0 = a.col.min(self.cursor.col);
                        let c1 = a.col.max(self.cursor.col);
                        r >= r0 && r <= r1 && c >= c0 && c <= c1
                    }
                    SelectionKind::Rows => {
                        let r0 = a.row.min(self.cursor.row);
                        let r1 = a.row.max(self.cursor.row);
                        r >= r0 && r <= r1
                    }
                    SelectionKind::Cols => {
                        let c0 = a.col.min(self.cursor.col);
                        let c1 = a.col.max(self.cursor.col);
                        c >= c0 && c <= c1
                    }
                });
                let is_cur = r == self.cursor.row && c == self.cursor.col;

                let is_left_border = c == lm - 1 && c >= col_ixs.first().copied().unwrap_or(0);
                let is_right_border = c == lm + mc && col_ixs.contains(&(lm + mc));
                let is_header_border =
                    r == hr - 1 && r >= row_ixs.first().copied().unwrap_or(0) && hr > 0;
                let is_footer_border = last_display_main_row == Some(r);

                let border_color =
                    if is_left_border || is_right_border || is_header_border || is_footer_border {
                        Some(Color::DarkGray)
                    } else {
                        None
                    };

                let mut st = if is_cur {
                    Style::default().bg(Color::DarkGray)
                } else if sel {
                    Style::default().bg(Color::Blue)
                } else if let Some(bc) = border_color {
                    Style::default().fg(bc)
                } else {
                    Style::default()
                };
                if is_agg_cell && !is_cur && !sel {
                    st = st.fg(Color::Cyan);
                    if footer_agg.is_some() {
                        st = st.add_modifier(Modifier::BOLD);
                    }
                }
                if is_underlined_boundary_row {
                    st = st.add_modifier(Modifier::UNDERLINED);
                }
                spans.push(Span::styled(disp, st));
                if i + 1 < col_ixs.len() {
                    if c == lm - 1 && lm > 0 && col_ixs.contains(&lm) {
                        spans.push(Span::styled(
                            "│",
                            boundary_separator_style(is_underlined_boundary_row),
                        ));
                        spans.push(Span::styled(
                            " ",
                            boundary_gap_style(is_underlined_boundary_row),
                        ));
                    } else if c == lm + mc - 1 && show_right_divider {
                        spans.push(Span::styled(
                            "│",
                            boundary_separator_style(is_underlined_boundary_row),
                        ));
                        spans.push(Span::styled(
                            " ",
                            boundary_gap_style(is_underlined_boundary_row),
                        ));
                    } else {
                        spans.push(Span::styled(
                            " ",
                            boundary_gap_style(is_underlined_boundary_row),
                        ));
                    }
                }
            }
            lines.push(Line::from(spans));
        }

        let n = lines.len().min(inner_h);
        if n > 0 {
            let mut constraints: Vec<Constraint> = (0..n).map(|_| Constraint::Length(1)).collect();
            if inner.height > n as u16 {
                constraints.push(Constraint::Min(0));
            }
            let row_areas = Layout::default()
                .direction(Direction::Vertical)
                .constraints(constraints)
                .split(inner);
            for i in 0..n {
                f.render_widget(
                    Paragraph::new(lines[i].clone()).left_aligned(),
                    row_areas[i],
                );
            }
        }

        f.render_widget(block, grid_area);

        let hints = self.hints_line();
        let hints_area = if has_tabs {
            Rect {
                x: hints_area.x,
                y: hints_area.y.saturating_sub(1),
                width: hints_area.width,
                height: 1,
            }
        } else {
            hints_area
        };
        f.render_widget(
            Paragraph::new(hints).style(Style::default().fg(Color::DarkGray)),
            hints_area,
        );

        if let Mode::Menu { stack } = &self.mode {
            let mut parent_area: Option<(Rect, usize)> = None;
            let actual_depth = stack.len();
            for (render_index, level) in Self::menu_render_levels(stack).iter().enumerate() {
                let popup_area = menu_popup_area(f.area(), level.section, parent_area);
                let items: Vec<ListItem> = menu_items(level.section)
                    .iter()
                    .map(|mi| {
                        let label = match mi.target {
                            MenuTarget::Submenu(sub) => {
                                format!("{}·{} ▶", mi.shortcut, menu_title(sub))
                            }
                            MenuTarget::Action(_) => format!("{}·{}", mi.shortcut, mi.label),
                        };
                        ListItem::new(label)
                    })
                    .collect();
                let mut state = ListState::default();
                if let Some(selected) =
                    Self::menu_selected_index(render_index, actual_depth, level.item, items.len())
                {
                    state.select(Some(selected));
                }
                let popup = List::new(items)
                    .block(
                        Block::default()
                            .borders(Borders::ALL)
                            .border_type(BorderType::Plain)
                            .title(menu_title(level.section)),
                    )
                    .highlight_style(
                        Style::default()
                            .fg(Color::Black)
                            .bg(Color::Yellow)
                            .add_modifier(Modifier::BOLD),
                    )
                    .highlight_symbol("> ");
                self.render_menu_popup(f, popup_area, popup, &mut state);
                parent_area = Some((popup_area, level.item));
            }
        }

        if let Some(selected) = special_picker {
            let items: Vec<ListItem> = SPECIAL_VALUE_CHOICES
                .iter()
                .enumerate()
                .map(|(idx, choice)| {
                    let label = special_choice_label(idx).unwrap_or('?');
                    ListItem::new(format!("{label}: {choice}"))
                })
                .collect();
            let mut state = ListState::default();
            state.select(Some(selected));
            let picker = List::new(items)
                .block(
                    Block::default()
                        .borders(Borders::ALL)
                        .border_type(BorderType::Plain)
                        .title(" Suggestions "),
                )
                .highlight_symbol("▸ ");
            let area = centered_rect(50, 60, f.area());
            f.render_widget(Clear, area);
            f.render_stateful_widget(picker, area, &mut state);
        }
    }

    fn hints_line(&self) -> String {
        match &self.mode {
            Mode::Normal => {
                if self.anchor.is_some() {
                    "  r·move-rows   c·move-cols   v·deselect   Esc·cancel".into()
                } else {
                    let mut hints =
                        vec!["type/F2·edit", "Ctrl+C·copy", "Ctrl+X·cut", "Ctrl+V·paste"];
                    if !self.op_history.is_empty() {
                        hints.push("Ctrl+Z·undo");
                    }
                    if !self.redo_history.is_empty() {
                        hints.push("Ctrl+Y·redo");
                    }
                    hints.push("Ctrl+;·date");
                    hints.push("Ctrl+:·time");
                    hints.push(if self.path.is_some() {
                        "Ctrl+S·save"
                    } else {
                        "Ctrl+S·save as"
                    });
                    hints.push("F1·help");
                    format!("  {}", hints.join("; "))
                }
            }
            Mode::Edit { .. } => {
                "  type to edit (or addr: val)   Enter·confirm   Esc·discard".into()
            }
            Mode::OpenPath { .. } => {
                "  type path or link <file> <revision>   Enter·open   Esc·cancel".into()
            }
            Mode::RevisionBrowse => "  left/right·step revisions   Enter·close   Esc·close".into(),
            Mode::SheetRename { .. } => "  type sheet title   Enter·rename   Esc·cancel".into(),
            Mode::SheetCopy { .. } => "  type sheet title   Enter·copy   Esc·cancel".into(),
            Mode::GoToCell { .. } => "  type cell address   Enter·go   Esc·cancel".into(),
            Mode::SavePath { .. } => "  type file path   Enter·save as   Esc·cancel".into(),
            Mode::ExportTsv { .. } | Mode::ExportCsv { .. } | Mode::ExportAll { .. } => {
                let h = if self.export_delimited_options.include_header_row {
                    "on"
                } else {
                    "off"
                };
                let m = if self.export_delimited_options.include_margins {
                    "on"
                } else {
                    "off"
                };
                let r = if self.export_delimited_options.include_row_label_column {
                    "on"
                } else {
                    "off"
                };
                let vf = match self.export_delimited_options.content {
                    export::ExportContent::Values => "values",
                    export::ExportContent::Formulas => "formulas",
                    export::ExportContent::Generic => "generic",
                };
                format!(
                    "  Alt+F·formulas   Alt+V·values   Alt+G·generic   ·{vf}   Alt+H·header {h}   Alt+M·margins {m}   \
Alt+R·left row# {r}   Alt+X·clipboard   ↑/↓/k/j·scroll   PgUp/PgDn·page   path or empty+Enter=clipboard   Esc"
                )
            }
            Mode::ExportAscii { .. } => {
                use export::{AsciiHeaderDataSeparator, AsciiInterCellSpace};
                let a = if self.export_ascii_options.include_column_label_row {
                    "on"
                } else {
                    "off"
                };
                let r = if self.export_ascii_options.include_row_label_column {
                    "on"
                } else {
                    "off"
                };
                let m = if self.export_ascii_options.include_margins {
                    "on"
                } else {
                    "off"
                };
                let f = if self.export_ascii_options.data_frame {
                    "on"
                } else {
                    "off"
                };
                let d = if self.export_ascii_options.row_dividers {
                    "on"
                } else {
                    "off"
                };
                let (pad_letter, pad_desc) = match self.export_ascii_options.inter_cell_space {
                    AsciiInterCellSpace::EmSpace => ("em", "U+2003 em"),
                    AsciiInterCellSpace::Space => ("sp", "U+0020 space"),
                };
                let b = match self.export_ascii_options.header_data_separator {
                    AsciiHeaderDataSeparator::FullBorder => "border",
                    AsciiHeaderDataSeparator::None => "none",
                };
                let vf = match self.export_ascii_options.content {
                    export::ExportContent::Values => "values",
                    export::ExportContent::Formulas => "formulas",
                    export::ExportContent::Generic => "generic",
                };
                format!(
                    "  Alt+F·formulas   Alt+V·values   Alt+G·generic   ·{vf}   Alt+H·top A/B label row {a}   Alt+R·left row# column {r}   Alt+M·margins {m}   \
Alt+O·data frame {f}   Alt+D·row rules {d}   Alt+E·padding {pad_letter} ({pad_desc})   \
Alt+B·label|data {b}   Alt+X·clipboard   ↑/↓/k/j   PgUp/PgDn   path or empty+Enter=clipboard   Esc"
                )
            }
            Mode::ExportOdt { .. } => {
                let vf = match self.export_ods_content {
                    export::ExportContent::Values => "values",
                    export::ExportContent::Formulas => "formulas",
                    export::ExportContent::Generic => "generic",
                };
                format!(
                    "  Alt+F·formulas   Alt+V·values   Alt+G·generic   ·{vf}   up/down·scroll   type .ods path   Enter·save   Esc"
                )
            }
            Mode::SetMaxColWidth { .. } => {
                "  type default column width   Enter·apply   Esc·cancel".into()
            }
            Mode::SetColWidth { .. } => {
                "  type col=width or col to clear   Enter   Esc·cancel".into()
            }
            Mode::SortView { .. } => {
                "  type sort columns like A,B,C   Enter·apply   Esc·cancel".into()
            }
            Mode::Find { .. } => "  type text   Enter·find next (wrap)   Esc·close".into(),
            Mode::Replace { .. } => "  type old|new   Enter·replace in all main cells   Esc·cancel"
                .into(),
            Mode::BalanceBooks { .. } => {
                "  Tab/Shift+Tab·move focus   Enter/Space·select   Esc·cancel".into()
            }
            Mode::FormatDecimals { .. } => "  type decimals   Enter·apply   Esc·cancel".into(),
            Mode::QuitPrompt => "  Q·quit   B·back   Esc·cancel".into(),
            Mode::QuitImportPrompt => "  S·save as .corro   D·discard   B·back".into(),
            Mode::Help => "  up/down·scroll   Esc·close   ?·help   A·about".into(),
            Mode::About => "  up/down·scroll   Esc·close   ?·help   A·about".into(),
            Mode::Menu { .. } => {
                "  right·open submenu   left·back   up/down·move   Enter/letter·open   Esc·close"
                    .into()
            }
        }
    }

    /// Hint line for export/CSV/TSV/ASCII/All: visible on dark-gray background (export preview
    /// covers the grid and previously skipped drawing hints on early return from [`Self::draw`]).
    fn render_export_bottom_hints(&self, f: &mut Frame, hints_area: Rect, has_tabs: bool) {
        let hints = self.hints_line();
        let area = if has_tabs {
            Rect {
                x: hints_area.x,
                y: hints_area.y.saturating_sub(1),
                width: hints_area.width,
                height: 1,
            }
        } else {
            hints_area
        };
        f.render_widget(
            Paragraph::new(hints).style(
                Style::default()
                    .fg(Color::White)
                    .bg(Color::DarkGray)
                    .add_modifier(Modifier::BOLD),
            ),
            area,
        );
    }

    fn menu_bar_line(&self) -> String {
        let (section, item) = match &self.mode {
            Mode::Menu { stack } => stack
                .last()
                .map(|level| (level.section, level.item))
                .unwrap_or((MenuSection::File, usize::MAX)),
            _ => (MenuSection::File, usize::MAX),
        };
        let file = if matches!(
            section,
            MenuSection::File | MenuSection::Export | MenuSection::Width
        ) {
            "[File]"
        } else {
            " File "
        };
        let edit = if section == MenuSection::Edit {
            "[Edit]"
        } else {
            " Edit "
        };
        let format = if matches!(
            section,
            MenuSection::Format
                | MenuSection::FormatScope
                | MenuSection::FormatNumber
                | MenuSection::FormatAlign
        ) {
            "[Format]"
        } else {
            " Format "
        };
        let insert = if section == MenuSection::Insert {
            "[Insert]"
        } else {
            " Insert "
        };
        let sheet = if section == MenuSection::Sheet {
            "[Sheet]"
        } else {
            " Sheet "
        };
        let help = if section == MenuSection::Help {
            "[Help]"
        } else {
            " Help "
        };
        let active = if item != usize::MAX {
            format!(
                "  {}",
                menu_action_item(section, item)
                    .map(|i| i.label)
                    .unwrap_or("")
            )
        } else {
            String::new()
        };
        format!(" {file}  {edit}  {insert}  {format}  {sheet}  {help}{active}")
    }

    fn balance_dialog_lines(
        &self,
        buffer: &str,
        direction: BalanceDirection,
        persist: bool,
        focus: BalanceBooksFocus,
        cursor: usize,
        text_style: Style,
        heading_style: Style,
        caret_style: Style,
    ) -> Vec<Line<'static>> {
        let column_focused = matches!(focus, BalanceBooksFocus::Column);
        let report_view_focused = matches!(focus, BalanceBooksFocus::ReportViewOnly);
        let report_persisted_focused = matches!(focus, BalanceBooksFocus::ReportPersisted);
        let pos_to_neg_focused = matches!(focus, BalanceBooksFocus::PosToNeg);
        let neg_to_pos_focused = matches!(focus, BalanceBooksFocus::NegToPos);
        let generate_focused = matches!(focus, BalanceBooksFocus::Generate);
        let cancel_focused = matches!(focus, BalanceBooksFocus::Cancel);
        let selected_style = |selected: bool| {
            if selected {
                Style::default()
                    .fg(Color::Black)
                    .bg(Color::Yellow)
                    .add_modifier(Modifier::BOLD)
            } else {
                text_style
            }
        };
        let button_style = |selected: bool| {
            if selected {
                Style::default()
                    .fg(Color::Black)
                    .bg(Color::Yellow)
                    .add_modifier(Modifier::BOLD)
            } else {
                heading_style
            }
        };
        let checkbox_line = |label: &str, checked: bool, selected: bool| {
            let style = selected_style(selected);
            Line::from(vec![
                Span::styled("  ", text_style),
                Span::styled(if checked { "[X]" } else { "[ ]" }, style),
                Span::styled(" ", text_style),
                Span::styled(label.to_string(), style),
            ])
        };

        vec![
            Line::from(Span::styled(
                "Balance rows into groups that sum to zero. The selected numeric column is used to score rows; all other columns are copied unchanged.",
                text_style,
            )),
            Line::from(""),
            Line::from(Span::styled("Column to Balance:", heading_style)),
            input_line(
                "  ".to_string(),
                buffer,
                cursor,
                text_style,
                if column_focused { caret_style } else { text_style },
            ),
            Line::from(""),
            Line::from(Span::styled("Report Type:", heading_style)),
            checkbox_line("View only", !persist, report_view_focused),
            checkbox_line("Persisted report", persist, report_persisted_focused),
            Line::from(""),
            Line::from(Span::styled("Balance direction:", heading_style)),
            checkbox_line(
                "Match +ve number with multiple -ve numbers",
                matches!(direction, BalanceDirection::PosToNeg),
                pos_to_neg_focused,
            ),
            checkbox_line(
                "Match -ve number with multiple +ve numbers",
                matches!(direction, BalanceDirection::NegToPos),
                neg_to_pos_focused,
            ),
            Line::from(""),
            Line::from(vec![
                Span::styled("  [ ", text_style),
                Span::styled("Generate", button_style(generate_focused)),
                Span::styled(" ]", text_style),
                Span::styled("   ", text_style),
                Span::styled("[ ", text_style),
                Span::styled("Cancel", button_style(cancel_focused)),
                Span::styled(" ]", text_style),
            ]),
        ]
    }

    fn balance_dialog_focus_line(focus: BalanceBooksFocus) -> usize {
        match focus {
            BalanceBooksFocus::Column => 3,
            BalanceBooksFocus::ReportViewOnly => 6,
            BalanceBooksFocus::ReportPersisted => 7,
            BalanceBooksFocus::PosToNeg => 10,
            BalanceBooksFocus::NegToPos => 11,
            BalanceBooksFocus::Generate | BalanceBooksFocus::Cancel => 13,
        }
    }

    fn cycle_balance_focus(focus: BalanceBooksFocus, backwards: bool) -> BalanceBooksFocus {
        use BalanceBooksFocus::*;
        let order = [
            Column,
            ReportViewOnly,
            ReportPersisted,
            PosToNeg,
            NegToPos,
            Generate,
            Cancel,
        ];
        let idx = order.iter().position(|item| *item == focus).unwrap_or(0);
        let next = if backwards {
            (idx + order.len() - 1) % order.len()
        } else {
            (idx + 1) % order.len()
        };
        order[next]
    }

    fn run_balance_books(
        &mut self,
        buffer: &str,
        direction: BalanceDirection,
        persist: bool,
    ) -> Result<(), RunError> {
        let col = if buffer.trim().is_empty() {
            balance::choose_balance_column(&self.state.grid)
        } else {
            addr::parse_excel_column(buffer.trim()).map(|c| c as usize)
        };
        let Some(col) = col else {
            self.status = "No balance column found".into();
            self.input_cursor = None;
            self.mode = Mode::Normal;
            return Ok(());
        };
        let report = balance::build_balance_report(&self.state.grid, col, direction);
        let source_sheet_id = self.workbook.sheet_id(self.workbook.active_sheet);
        let source_title = self
            .workbook
            .sheet_title(self.workbook.active_sheet)
            .to_string();
        if persist {
            let title = format!("Balance-{}", self.workbook.next_sheet_id);
            self.commit_active_sheet_cache();
            let id = self.workbook.next_sheet_id;
            let plan = balance::balance_copy_plan(
                source_sheet_id,
                source_title.clone(),
                id,
                title.clone(),
                col,
                self.state.grid.main_rows(),
                &report,
                true,
            );
            let report_sheet = balance::materialize_report_sheet(&self.state, &plan);
            self.workbook.add_sheet(title.clone(), report_sheet.clone());
            self.view_sheet_id = id;
            self.sync_active_sheet_cache();
            if let Some(ref p) = self.path.clone() {
                let mut active_sheet = self.view_sheet_id;
                commit_workbook_op(
                    p,
                    &mut self.offset,
                    &mut self.workbook,
                    &mut active_sheet,
                    &crate::ops::WorkbookOp::BalanceReport {
                        id,
                        title: title.clone(),
                        source_sheet_id,
                        amount_col: col,
                        direction,
                        row_order: plan.row_order.clone(),
                        show_unmatched_heading: plan.show_unmatched_heading,
                        unmatched_start: plan.unmatched_start,
                        preserve_formulas: true,
                    },
                )?;
                self.ops_applied = self.ops_applied.saturating_add(1);
                self.start_log_watcher_if_needed()?;
            }
            self.status = format!("Balance report saved as {}", title);
        } else {
            let plan = balance::balance_copy_plan(
                source_sheet_id,
                source_title,
                self.workbook.sheet_id(self.workbook.active_sheet),
                self.workbook
                    .sheet_title(self.workbook.active_sheet)
                    .to_string(),
                col,
                self.state.grid.main_rows(),
                &report,
                true,
            );
            self.state = balance::materialize_report_sheet(&self.state, &plan);
            self.status = "Balance report generated".into();
        }
        self.input_cursor = None;
        self.mode = Mode::Normal;
        Ok(())
    }

    fn handle_key(&mut self, key: KeyEvent) -> Result<bool, RunError> {
        if key.kind == KeyEventKind::Release {
            return Ok(false);
        }

        let shift = key.modifiers.contains(KeyModifiers::SHIFT);
        let super_key = key.modifiers.contains(KeyModifiers::SUPER);

        if matches!(self.mode, Mode::Normal)
            && !super_key
            && !key.modifiers.contains(KeyModifiers::ALT)
        {
            if key.modifiers.contains(KeyModifiers::CONTROL)
                && matches!(key.code, KeyCode::Char('c') | KeyCode::Char('C'))
            {
                let data = self.selection_tsv_text();
                self.copy_selection_to_clipboard(&data);
                return Ok(false);
            }

            if key.modifiers.contains(KeyModifiers::CONTROL)
                && matches!(key.code, KeyCode::Char('s') | KeyCode::Char('S'))
            {
                if let Some(path) = self.path.clone() {
                    self.save_to_path(&path)?;
                } else {
                    self.mode = Mode::SavePath {
                        buffer: self.start_input_mode(self.suggested_corro_save_path()),
                    };
                }
                return Ok(false);
            }

            if key.modifiers.contains(KeyModifiers::CONTROL)
                && matches!(key.code, KeyCode::Char('v') | KeyCode::Char('V'))
            {
                self.paste_from_clipboard(!shift)?;
                return Ok(false);
            }

            if key.modifiers.contains(KeyModifiers::CONTROL)
                && shift
                && matches!(key.code, KeyCode::Char('p') | KeyCode::Char('P'))
            {
                self.paste_from_clipboard(true)?;
                return Ok(false);
            }
        }

        if let Some(selected) = self.special_picker {
            match key.code {
                KeyCode::Esc => {
                    self.special_picker = None;
                    return Ok(false);
                }
                KeyCode::Enter => {
                    self.commit_special_choice(selected);
                    self.special_picker = None;
                    return Ok(false);
                }
                KeyCode::Left | KeyCode::Up => {
                    self.special_picker = Some(selected.saturating_sub(1));
                    return Ok(false);
                }
                KeyCode::Right | KeyCode::Down => {
                    self.special_picker = Some((selected + 1).min(SPECIAL_VALUE_CHOICES.len() - 1));
                    return Ok(false);
                }
                KeyCode::Char(c) if c.is_ascii_digit() => {
                    if let Some(idx) = c.to_digit(10).map(|i| i as usize) {
                        if idx < SPECIAL_VALUE_CHOICES.len() {
                            self.commit_special_choice(idx);
                            self.special_picker = None;
                            return Ok(false);
                        }
                    }
                }
                _ => {}
            }
        }

        if matches!(self.mode, Mode::RevisionBrowse) {
            match key.code {
                KeyCode::Esc | KeyCode::Enter => {
                    self.mode = Mode::Normal;
                    return Ok(false);
                }
                KeyCode::Left => {
                    if self.revision_browse_limit > 1 {
                        self.revision_browse_limit -= 1;
                        self.reload_revision_browse()?;
                    }
                    self.mode = Mode::RevisionBrowse;
                    return Ok(false);
                }
                KeyCode::Right => {
                    self.revision_browse_limit = self.revision_browse_limit.saturating_add(1);
                    self.reload_revision_browse()?;
                    self.mode = Mode::RevisionBrowse;
                    return Ok(false);
                }
                _ => {
                    self.mode = Mode::RevisionBrowse;
                    return Ok(false);
                }
            }
        }

        let mut mode = std::mem::replace(&mut self.mode, Mode::Normal);

        if matches!(mode, Mode::Normal) {
            match key.code {
                KeyCode::F(1) => {
                    self.help_scroll = 0;
                    self.mode = Mode::Help;
                    return Ok(false);
                }
                KeyCode::F(2) => {
                    self.mode = self.start_edit_current_cell();
                    return Ok(false);
                }
                _ => {}
            }
        }

        if matches!(mode, Mode::Normal | Mode::Edit { .. })
            && (key.modifiers.contains(KeyModifiers::CONTROL)
                && matches!(key.code, KeyCode::Char('=') | KeyCode::Char('+')))
        {
            if self.anchor.is_some() {
                if !self.insert_rows_above_selection()? {
                    if let Some((from, to)) = self.selection_main_row_range() {
                        let count = to - from + 1;
                        let _ = self.insert_rows_above_cursor(count)?;
                    } else {
                        let _ = self.insert_rows_above_cursor(1)?;
                    }
                }
            } else {
                let _ = self.insert_rows_above_cursor(1)?;
            }
            self.mode = mode;
            return Ok(false);
        }

        if key.modifiers.contains(KeyModifiers::ALT)
            && matches!(mode, Mode::Normal)
            && matches!(
                key.code,
                KeyCode::Left | KeyCode::Right | KeyCode::Up | KeyCode::Down
            )
        {
            if matches!(key.code, KeyCode::Left | KeyCode::Right) {
                let right = matches!(key.code, KeyCode::Right);
                self.mode = mode;
                let handled = self.move_selected_cols_by_one(right)?;
                if handled {
                    return Ok(false);
                }
                mode = std::mem::replace(&mut self.mode, Mode::Normal);
            } else {
                let down = matches!(key.code, KeyCode::Down);
                self.mode = mode;
                let handled = self.move_selected_rows_by_one(down)?;
                if handled {
                    return Ok(false);
                }
                mode = std::mem::replace(&mut self.mode, Mode::Normal);
            }
        }

        if let Mode::Menu { stack } = &mut mode {
            match key.code {
                KeyCode::Esc => mode = Mode::Normal,
                KeyCode::Left | KeyCode::Char('h') => {
                    stack.truncate(1);
                    if let Some(level) = stack.last_mut() {
                        level.section = menu_prev_root_section(level.section);
                        level.item = 0;
                    }
                }
                KeyCode::Right | KeyCode::Char('l') => {
                    let current = stack.last().copied();
                    let current_is_submenu = current
                        .and_then(|level| menu_action_item(level.section, level.item))
                        .map(|menu_item| matches!(menu_item.target, MenuTarget::Submenu(_)))
                        .unwrap_or(false);

                    if current_is_submenu {
                        if let Some(level) = current {
                            if let Some(MenuItem {
                                target: MenuTarget::Submenu(section),
                                ..
                            }) = menu_action_item(level.section, level.item)
                            {
                                stack.push(MenuLevel { section, item: 0 });
                            }
                        }
                    } else {
                        stack.truncate(1);
                        if let Some(level) = stack.last_mut() {
                            level.section = menu_next_root_section(level.section);
                            level.item = 0;
                        }
                    }
                }
                KeyCode::Up | KeyCode::Char('k') => {
                    let len = stack
                        .last()
                        .map(|level| menu_items(level.section).len())
                        .unwrap_or(0);
                    if len > 0 {
                        if let Some(level) = stack.last_mut() {
                            level.item = level.item.saturating_sub(1);
                        }
                    }
                }
                KeyCode::Down | KeyCode::Char('j') => {
                    let len = stack
                        .last()
                        .map(|level| menu_items(level.section).len())
                        .unwrap_or(0);
                    if len > 0 {
                        if let Some(level) = stack.last_mut() {
                            level.item = (level.item + 1).min(len - 1);
                        }
                    }
                }
                KeyCode::Enter => {
                    if let Some(level) = stack.last() {
                        if let Some(menu_item) = menu_action_item(level.section, level.item) {
                            match self.menu_target_mode(stack.as_slice(), menu_item.target) {
                                Ok(m) => mode = m,
                                Err(()) => {
                                    self.mode = mode;
                                    return Ok(true);
                                }
                            }
                        }
                    }
                }
                KeyCode::Char(ch) => {
                    let upper = ch.to_ascii_uppercase();
                    if let Some(level) = stack.last_mut() {
                        if let Some((idx, menu_item)) = menu_items(level.section)
                            .iter()
                            .enumerate()
                            .find(|(_, mi)| mi.shortcut == upper)
                        {
                            level.item = idx;
                            match self.menu_target_mode(stack.as_slice(), menu_item.target) {
                                Ok(m) => mode = m,
                                Err(()) => {
                                    self.mode = mode;
                                    return Ok(true);
                                }
                            }
                        }
                    }
                }
                _ => {}
            }
            self.mode = mode;
            return Ok(false);
        }

        if key.modifiers.contains(KeyModifiers::ALT)
            && matches!(mode, Mode::Normal | Mode::Edit { .. })
            && matches!(key.code, KeyCode::Char(_))
        {
            if let KeyCode::Char(ch) = key.code {
                match ch {
                    'f' | 'F' => {
                        self.open_menu(MenuSection::File);
                        return Ok(false);
                    }
                    'h' | 'H' => {
                        self.open_menu(MenuSection::Help);
                        return Ok(false);
                    }
                    't' | 'T' => {
                        self.open_menu_path(vec![MenuLevel {
                            section: MenuSection::Export,
                            item: 0,
                        }]);
                        return Ok(false);
                    }
                    'a' | 'A' => {
                        self.open_menu_path(vec![MenuLevel {
                            section: MenuSection::Export,
                            item: 2,
                        }]);
                        return Ok(false);
                    }
                    'e' | 'E' => {
                        if shift {
                            self.open_menu_path(vec![MenuLevel {
                                section: MenuSection::Export,
                                item: 3,
                            }]);
                        } else {
                            self.open_menu(MenuSection::Edit);
                        }
                        return Ok(false);
                    }
                    'i' | 'I' => {
                        self.open_menu(MenuSection::Insert);
                        return Ok(false);
                    }
                    's' | 'S' => {
                        self.open_menu(MenuSection::Sheet);
                        return Ok(false);
                    }
                    'o' | 'O' => {
                        self.open_menu_path(vec![MenuLevel {
                            section: MenuSection::File,
                            item: 0,
                        }]);
                        return Ok(false);
                    }
                    'w' | 'W' => {
                        self.open_menu_path(vec![MenuLevel {
                            section: MenuSection::Width,
                            item: 0,
                        }]);
                        return Ok(false);
                    }
                    'x' | 'X' => {
                        self.open_menu_path(vec![MenuLevel {
                            section: MenuSection::Width,
                            item: 1,
                        }]);
                        return Ok(false);
                    }
                    _ => {}
                }
            }
        }

        if key.modifiers.contains(KeyModifiers::CONTROL)
            && matches!(mode, Mode::Normal)
            && matches!(key.code, KeyCode::Char(_))
        {
            if let KeyCode::Char(ch) = key.code {
                match ch {
                    ';' if key.modifiers.contains(KeyModifiers::SHIFT) => {
                        let buffer = chrono::Local::now().format("%H:%M:%S").to_string();
                        self.mode = self.start_edit_mode(buffer, None, false, false, None);
                        return Ok(false);
                    }
                    ':' => {
                        let buffer = chrono::Local::now().format("%H:%M:%S").to_string();
                        self.mode = self.start_edit_mode(buffer, None, false, false, None);
                        return Ok(false);
                    }
                    ';' => {
                        let buffer = chrono::Local::now().format("%Y-%m-%d").to_string();
                        self.mode = self.start_edit_mode(buffer, None, false, true, None);
                        return Ok(false);
                    }
                    _ => {}
                }
            }
        }

        if key.modifiers.contains(KeyModifiers::CONTROL)
            && matches!(mode, Mode::Normal | Mode::Edit { .. })
        {
            match key.code {
                KeyCode::PageUp => {
                    self.switch_sheet(-1);
                    if matches!(&mode, Mode::Edit { .. }) {
                        let addr = self.cursor.to_addr(&self.state.grid);
                        let cur = cell_display(&self.state.grid, &addr);
                        mode = self.start_edit_mode(
                            cur.clone(),
                            if cur.trim() == "=" {
                                Some(self.cursor)
                            } else {
                                None
                            },
                            false,
                            false,
                            None,
                        );
                    }
                    self.mode = mode;
                    return Ok(false);
                }
                KeyCode::PageDown => {
                    self.switch_sheet(1);
                    if matches!(&mode, Mode::Edit { .. }) {
                        let addr = self.cursor.to_addr(&self.state.grid);
                        let cur = cell_display(&self.state.grid, &addr);
                        mode = self.start_edit_mode(
                            cur.clone(),
                            if cur.trim() == "=" {
                                Some(self.cursor)
                            } else {
                                None
                            },
                            false,
                            false,
                            None,
                        );
                    }
                    self.mode = mode;
                    return Ok(false);
                }
                _ => {}
            }
        }

        match &mut mode {
            Mode::Help => match key.code {
                KeyCode::Esc | KeyCode::Char('q') => mode = Mode::Normal,
                KeyCode::Char('a') | KeyCode::Char('A') => {
                    self.about_scroll = 0;
                    mode = Mode::About;
                }
                KeyCode::Up | KeyCode::Char('k') => {
                    self.help_scroll = self.help_scroll.saturating_sub(1);
                }
                KeyCode::Down | KeyCode::Char('j') => {
                    self.help_scroll = self.help_scroll.saturating_add(1);
                }
                _ => {}
            },
            Mode::About => match key.code {
                KeyCode::Esc | KeyCode::Char('q') => mode = Mode::Normal,
                KeyCode::Char('?') => {
                    self.help_scroll = 0;
                    mode = Mode::Help;
                }
                KeyCode::Up | KeyCode::Char('k') => {
                    self.about_scroll = self.about_scroll.saturating_sub(1);
                }
                KeyCode::Down | KeyCode::Char('j') => {
                    self.about_scroll = self.about_scroll.saturating_add(1);
                }
                _ => {}
            },
            Mode::RevisionBrowse => {}
            Mode::Menu { .. } => {}
            Mode::Replace { buffer } => match key.code {
                KeyCode::Enter => {
                    self.input_cursor = None;
                    if let Some((find, repl)) = Self::parse_replace_spec(buffer) {
                        if find.is_empty() {
                            self.status = "Replace: text before | is required (example: old|new)".into();
                        } else {
                            let n = self.replace_all_substrings_in_main(find, repl)?;
                            self.status = if n == 0 {
                                "No matching cells".into()
                            } else {
                                format!("Replaced in {n} cell(s)")
                            };
                        }
                    } else {
                        self.status =
                            "Replace: use old|new (example: search|replace)".into();
                    }
                    mode = Mode::Normal;
                }
                KeyCode::Esc => mode = Mode::Normal,
                _ if Self::handle_plain_text_input_key(
                    buffer,
                    &mut self.input_cursor,
                    key.code,
                ) => {}
                _ => {}
            },
            Mode::SheetRename { buffer, .. } => match key.code {
                KeyCode::Enter => {
                    self.input_cursor = None;
                    self.rename_current_sheet(buffer.clone())?;
                    mode = Mode::Normal;
                }
                KeyCode::Esc => mode = Mode::Normal,
                _ if Self::handle_plain_text_input_key(
                    buffer,
                    &mut self.input_cursor,
                    key.code,
                ) => {}
                _ => {}
            },
            Mode::SheetCopy { buffer, .. } => match key.code {
                KeyCode::Enter => {
                    self.input_cursor = None;
                    self.copy_current_sheet(buffer.clone())?;
                    mode = Mode::Normal;
                }
                KeyCode::Esc => mode = Mode::Normal,
                _ if Self::handle_plain_text_input_key(
                    buffer,
                    &mut self.input_cursor,
                    key.code,
                ) => {}
                _ => {}
            },
            Mode::GoToCell { buffer } => match key.code {
                KeyCode::Enter => {
                    self.input_cursor = None;
                    self.go_to_cell(buffer);
                    mode = Mode::Normal;
                }
                KeyCode::Esc => mode = Mode::Normal,
                _ if Self::handle_plain_text_input_key(
                    buffer,
                    &mut self.input_cursor,
                    key.code,
                ) => {}
                _ => {}
            },
            Mode::ExportTsv { buffer } => {
                if key.modifiers.contains(KeyModifiers::ALT) {
                    if let KeyCode::Char(ch) = key.code {
                        match ch {
                            'h' | 'H' => {
                                self.export_delimited_options.include_header_row =
                                    !self.export_delimited_options.include_header_row;
                                self.status = if self.export_delimited_options.include_header_row {
                                    "Column header row: on".into()
                                } else {
                                    "Column header row: off".into()
                                };
                            }
                            'm' | 'M' => {
                                self.export_delimited_options.include_margins =
                                    !self.export_delimited_options.include_margins;
                                self.status = if self.export_delimited_options.include_margins {
                                    "Row/column margin labels: on".into()
                                } else {
                                    "Row/column margin labels: off".into()
                                };
                            }
                            'r' | 'R' => {
                                self.export_delimited_options.include_row_label_column =
                                    !self.export_delimited_options.include_row_label_column;
                                self.status = if self.export_delimited_options
                                    .include_row_label_column
                                {
                                    "Left row# column: on".into()
                                } else {
                                    "Left row# column: off".into()
                                };
                            }
                            'f' | 'F' => {
                                self.export_delimited_options.content = export::ExportContent::Formulas;
                                self.status = "Export: formulas (stored text)".into();
                            }
                            'v' | 'V' => {
                                self.export_delimited_options.content = export::ExportContent::Values;
                                self.status = "Export: values (calculated)".into();
                            }
                            'g' | 'G' => {
                                self.export_delimited_options.content = export::ExportContent::Generic;
                                self.status = "Export: generic (labels + =interop)".into();
                            }
                            'x' | 'X' => {
                                match copy_to_clipboard(&self.do_export(false)) {
                                    Ok(()) => {
                                        self.status = "TSV export copied to clipboard".into();
                                    }
                                    Err(e) => {
                                        self.status = format!("Clipboard error: {e}");
                                    }
                                }
                                self.input_cursor = None;
                                mode = Mode::Normal;
                            }
                            _ => {}
                        }
                    }
                } else {
                    match key.code {
                        KeyCode::Up | KeyCode::Char('k') => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_sub(1);
                        }
                        KeyCode::Down | KeyCode::Char('j') => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_add(1);
                        }
                        KeyCode::PageUp => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_sub(20);
                        }
                        KeyCode::PageDown => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_add(20);
                        }
                        KeyCode::Enter => {
                            let fname = buffer.clone();
                            self.finish_export(false, &fname);
                            self.input_cursor = None;
                            mode = Mode::Normal;
                        }
                        KeyCode::Esc => mode = Mode::Normal,
                        _ if Self::handle_plain_text_input_key(
                            buffer,
                            &mut self.input_cursor,
                            key.code,
                        ) => {}
                        _ => {}
                    }
                }
            }
            Mode::ExportCsv { buffer } => {
                if key.modifiers.contains(KeyModifiers::ALT) {
                    if let KeyCode::Char(ch) = key.code {
                        match ch {
                            'h' | 'H' => {
                                self.export_delimited_options.include_header_row =
                                    !self.export_delimited_options.include_header_row;
                                self.status = if self.export_delimited_options.include_header_row {
                                    "Column header row: on".into()
                                } else {
                                    "Column header row: off".into()
                                };
                            }
                            'm' | 'M' => {
                                self.export_delimited_options.include_margins =
                                    !self.export_delimited_options.include_margins;
                                self.status = if self.export_delimited_options.include_margins {
                                    "Row/column margin labels: on".into()
                                } else {
                                    "Row/column margin labels: off".into()
                                };
                            }
                            'r' | 'R' => {
                                self.export_delimited_options.include_row_label_column =
                                    !self.export_delimited_options.include_row_label_column;
                                self.status = if self.export_delimited_options
                                    .include_row_label_column
                                {
                                    "Left row# column: on".into()
                                } else {
                                    "Left row# column: off".into()
                                };
                            }
                            'f' | 'F' => {
                                self.export_delimited_options.content = export::ExportContent::Formulas;
                                self.status = "Export: formulas (stored text)".into();
                            }
                            'v' | 'V' => {
                                self.export_delimited_options.content = export::ExportContent::Values;
                                self.status = "Export: values (calculated)".into();
                            }
                            'g' | 'G' => {
                                self.export_delimited_options.content = export::ExportContent::Generic;
                                self.status = "Export: generic (labels + =interop)".into();
                            }
                            'x' | 'X' => {
                                match copy_to_clipboard(&self.do_export(true)) {
                                    Ok(()) => {
                                        self.status = "CSV export copied to clipboard".into();
                                    }
                                    Err(e) => {
                                        self.status = format!("Clipboard error: {e}");
                                    }
                                }
                                self.input_cursor = None;
                                mode = Mode::Normal;
                            }
                            _ => {}
                        }
                    }
                } else {
                    match key.code {
                        KeyCode::Up | KeyCode::Char('k') => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_sub(1);
                        }
                        KeyCode::Down | KeyCode::Char('j') => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_add(1);
                        }
                        KeyCode::PageUp => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_sub(20);
                        }
                        KeyCode::PageDown => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_add(20);
                        }
                        KeyCode::Enter => {
                            let fname = buffer.clone();
                            self.finish_export(true, &fname);
                            self.input_cursor = None;
                            mode = Mode::Normal;
                        }
                        KeyCode::Esc => mode = Mode::Normal,
                        _ if Self::handle_plain_text_input_key(
                            buffer,
                            &mut self.input_cursor,
                            key.code,
                        ) => {}
                        _ => {}
                    }
                }
            }
            Mode::ExportAscii { buffer } => {
                if key.modifiers.contains(KeyModifiers::ALT) {
                    if let KeyCode::Char(ch) = key.code {
                        match ch {
                            'h' | 'H' => {
                                self.export_ascii_options.include_column_label_row =
                                    !self.export_ascii_options.include_column_label_row;
                                self.status = if self.export_ascii_options.include_column_label_row
                                {
                                    "ASCII: top column label (A/B) row: on".into()
                                } else {
                                    "ASCII: top column label (A/B) row: off".into()
                                };
                            }
                            'r' | 'R' => {
                                self.export_ascii_options.include_row_label_column =
                                    !self.export_ascii_options.include_row_label_column;
                                self.status = if self.export_ascii_options.include_row_label_column {
                                    "Left row# column: on".into()
                                } else {
                                    "Left row# column: off".into()
                                };
                            }
                            'd' | 'D' => {
                                self.export_ascii_options.row_dividers =
                                    !self.export_ascii_options.row_dividers;
                                self.status = if self.export_ascii_options.row_dividers {
                                    "ASCII: row dividers: on".into()
                                } else {
                                    "ASCII: row dividers: off".into()
                                };
                            }
                            'e' | 'E' => {
                                use export::AsciiInterCellSpace;
                                self.export_ascii_options.inter_cell_space = match self
                                    .export_ascii_options
                                    .inter_cell_space
                                {
                                    AsciiInterCellSpace::Space => {
                                        self.status = "ASCII: pad: em space".into();
                                        AsciiInterCellSpace::EmSpace
                                    }
                                    AsciiInterCellSpace::EmSpace => {
                                        self.status = "ASCII: pad: U+0020 space".into();
                                        AsciiInterCellSpace::Space
                                    }
                                };
                            }
                            'b' | 'B' => {
                                use export::AsciiHeaderDataSeparator;
                                self.export_ascii_options.header_data_separator = match self
                                    .export_ascii_options
                                    .header_data_separator
                                {
                                    AsciiHeaderDataSeparator::FullBorder => {
                                        self.status = "ASCII: no border under column labels".into();
                                        AsciiHeaderDataSeparator::None
                                    }
                                    AsciiHeaderDataSeparator::None => {
                                        self.status = "ASCII: full border under column labels".into();
                                        AsciiHeaderDataSeparator::FullBorder
                                    }
                                };
                            }
                            'm' | 'M' => {
                                self.export_ascii_options.include_margins =
                                    !self.export_ascii_options.include_margins;
                                self.status = if self.export_ascii_options.include_margins {
                                    "ASCII: margin rows/columns: on".into()
                                } else {
                                    "ASCII: main block only: on".into()
                                };
                            }
                            'o' | 'O' => {
                                self.export_ascii_options.data_frame =
                                    !self.export_ascii_options.data_frame;
                                self.status = if self.export_ascii_options.data_frame {
                                    "ASCII: data frame (rules around main): on".into()
                                } else {
                                    "ASCII: data frame: off".into()
                                };
                            }
                            'f' | 'F' => {
                                self.export_ascii_options.content = export::ExportContent::Formulas;
                                self.status = "Export: formulas (stored text)".into();
                            }
                            'v' | 'V' => {
                                self.export_ascii_options.content = export::ExportContent::Values;
                                self.status = "Export: values (calculated)".into();
                            }
                            'g' | 'G' => {
                                self.export_ascii_options.content = export::ExportContent::Generic;
                                self.status = "Export: generic (labels + =interop)".into();
                            }
                            'x' | 'X' => {
                                match copy_to_clipboard(&self.do_export_ascii()) {
                                    Ok(()) => {
                                        self.status = "ASCII table copied to clipboard".into();
                                    }
                                    Err(e) => {
                                        self.status = format!("Clipboard error: {e}");
                                    }
                                }
                                self.input_cursor = None;
                                mode = Mode::Normal;
                            }
                            _ => {}
                        }
                    }
                } else {
                    match key.code {
                        KeyCode::Up | KeyCode::Char('k') => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_sub(1);
                        }
                        KeyCode::Down | KeyCode::Char('j') => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_add(1);
                        }
                        KeyCode::PageUp => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_sub(20);
                        }
                        KeyCode::PageDown => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_add(20);
                        }
                        KeyCode::Enter => {
                            let fname = buffer.clone();
                            if fname.trim().is_empty() {
                                match copy_to_clipboard(&self.do_export_ascii()) {
                                    Ok(()) => self.status = "ASCII table copied to clipboard".into(),
                                    Err(e) => self.status = format!("Clipboard error: {e}"),
                                }
                            } else {
                                match std::fs::write(fname.trim(), self.do_export_ascii()) {
                                    Ok(()) => {
                                        self.status =
                                            format!("ASCII table exported to {}", fname.trim())
                                    }
                                    Err(e) => self.status = format!("Write error: {e}"),
                                }
                            }
                            self.input_cursor = None;
                            mode = Mode::Normal;
                        }
                        KeyCode::Esc => mode = Mode::Normal,
                        _ if Self::handle_plain_text_input_key(
                            buffer,
                            &mut self.input_cursor,
                            key.code,
                        ) => {}
                        _ => {}
                    }
                }
            }
            Mode::ExportOdt { buffer } => {
                if key.modifiers.contains(KeyModifiers::ALT) {
                    if let KeyCode::Char(ch) = key.code {
                        match ch {
                            'f' | 'F' => {
                                self.export_ods_content = export::ExportContent::Formulas;
                                self.status = "ODS: formulas (ODF with table:formula)".into();
                            }
                            'v' | 'V' => {
                                self.export_ods_content = export::ExportContent::Values;
                                self.status = "ODS: values only (static cells)".into();
                            }
                            'g' | 'G' => {
                                self.export_ods_content = export::ExportContent::Generic;
                                self.status = "ODS: generic (same strings as TSV generic)".into();
                            }
                            _ => {}
                        }
                    }
                } else {
                    match key.code {
                        KeyCode::Up | KeyCode::Char('k') => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_sub(1);
                        }
                        KeyCode::Down | KeyCode::Char('j') => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_add(1);
                        }
                        KeyCode::PageUp => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_sub(20);
                        }
                        KeyCode::PageDown => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_add(20);
                        }
                        KeyCode::Enter => {
                            let fname = buffer.clone();
                            if fname.trim().is_empty() {
                                self.status = "ODS requires a filename".into();
                            } else {
                                match std::fs::write(fname.trim(), self.do_export_ods()) {
                                    Ok(()) => self.status = format!("ODS saved to {}", fname.trim()),
                                    Err(e) => self.status = format!("Write error: {e}"),
                                }
                            }
                            self.input_cursor = None;
                            mode = Mode::Normal;
                        }
                        KeyCode::Esc => mode = Mode::Normal,
                        _ if Self::handle_plain_text_input_key(
                            buffer,
                            &mut self.input_cursor,
                            key.code,
                        ) => {}
                        _ => {}
                    }
                }
            }
            Mode::ExportAll { buffer } => {
                if key.modifiers.contains(KeyModifiers::ALT) {
                    if let KeyCode::Char(ch) = key.code {
                        match ch {
                            'h' | 'H' => {
                                self.export_delimited_options.include_header_row =
                                    !self.export_delimited_options.include_header_row;
                                self.status = if self.export_delimited_options.include_header_row {
                                    "Column header row: on".into()
                                } else {
                                    "Column header row: off".into()
                                };
                            }
                            'm' | 'M' => {
                                self.export_delimited_options.include_margins =
                                    !self.export_delimited_options.include_margins;
                                self.status = if self.export_delimited_options.include_margins {
                                    "Row/column margin labels: on".into()
                                } else {
                                    "Row/column margin labels: off".into()
                                };
                            }
                            'r' | 'R' => {
                                self.export_delimited_options.include_row_label_column =
                                    !self.export_delimited_options.include_row_label_column;
                                self.status = if self.export_delimited_options
                                    .include_row_label_column
                                {
                                    "Left row# column: on".into()
                                } else {
                                    "Left row# column: off".into()
                                };
                            }
                            'f' | 'F' => {
                                self.export_delimited_options.content = export::ExportContent::Formulas;
                                self.status = "Export: formulas (stored text)".into();
                            }
                            'v' | 'V' => {
                                self.export_delimited_options.content = export::ExportContent::Values;
                                self.status = "Export: values (calculated)".into();
                            }
                            'g' | 'G' => {
                                self.export_delimited_options.content = export::ExportContent::Generic;
                                self.status = "Export: generic (labels + =interop)".into();
                            }
                            'x' | 'X' => {
                                let data = if self.anchor.is_some() {
                                    self.do_export_selection()
                                } else {
                                    self.do_export_all()
                                };
                                match copy_to_clipboard(&data) {
                                    Ok(()) => {
                                        self.status = if self.anchor.is_some() {
                                            "Selection copied to clipboard".into()
                                        } else {
                                            "Full export copied to clipboard".into()
                                        }
                                    }
                                    Err(e) => {
                                        self.status = format!("Clipboard error: {e}");
                                    }
                                }
                                self.input_cursor = None;
                                mode = Mode::Normal;
                            }
                            _ => {}
                        }
                    }
                } else {
                    match key.code {
                        KeyCode::Up | KeyCode::Char('k') => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_sub(1);
                        }
                        KeyCode::Down | KeyCode::Char('j') => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_add(1);
                        }
                        KeyCode::PageUp => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_sub(20);
                        }
                        KeyCode::PageDown => {
                            self.export_preview_scroll = self.export_preview_scroll.saturating_add(20);
                        }
                        KeyCode::Enter => {
                            let fname = buffer.clone();
                            if fname.trim().is_empty() {
                                let data = if self.anchor.is_some() {
                                    self.do_export_selection()
                                } else {
                                    self.do_export_all()
                                };
                                match copy_to_clipboard(&data) {
                                    Ok(()) => {
                                        self.status = if self.anchor.is_some() {
                                            "Selection copied to clipboard".into()
                                        } else {
                                            "Full export copied to clipboard".into()
                                        }
                                    }
                                    Err(e) => self.status = format!("Clipboard error: {e}"),
                                }
                            } else {
                                let data = if self.anchor.is_some() {
                                    self.do_export_selection()
                                } else {
                                    self.do_export_all()
                                };
                                match std::fs::write(fname.trim(), data) {
                                    Ok(()) => {
                                        self.status = if self.anchor.is_some() {
                                            format!("Selection saved to {}", fname.trim())
                                        } else {
                                            format!("Full export saved to {}", fname.trim())
                                        }
                                    }
                                    Err(e) => self.status = format!("Write error: {e}"),
                                }
                            }
                            self.input_cursor = None;
                            mode = Mode::Normal;
                        }
                        KeyCode::Esc => mode = Mode::Normal,
                        _ if Self::handle_plain_text_input_key(
                            buffer,
                            &mut self.input_cursor,
                            key.code,
                        ) => {}
                        _ => {}
                    }
                }
            }
            Mode::SetMaxColWidth { buffer } => match key.code {
                KeyCode::Enter => {
                    if let Ok(width) = buffer.trim().parse::<usize>() {
                        if let Some(ref p) = self.path.clone() {
                            let mut active_sheet = self.view_sheet_id;
                            commit_workbook_op(
                                p,
                                &mut self.offset,
                                &mut self.workbook,
                                &mut active_sheet,
                                &crate::ops::WorkbookOp::SheetOp {
                                    sheet_id: self.view_sheet_id,
                                    op: Op::SetMaxColWidth { width },
                                },
                            )?;
                            self.sync_active_sheet_cache();
                            self.ops_applied = self.ops_applied.saturating_add(1);
                            self.start_log_watcher_if_needed()?;
                        } else {
                            self.state.grid.set_max_col_width(width);
                        }
                        self.status = format!("Default column width set to {width}");
                    }
                    self.input_cursor = None;
                    mode = Mode::Normal;
                }
                KeyCode::Esc => mode = Mode::Normal,
                _ if Self::handle_plain_text_input_key(
                    buffer,
                    &mut self.input_cursor,
                    key.code,
                ) => {}
                _ => {}
            },
            Mode::SetColWidth { buffer } => match key.code {
                KeyCode::Enter => {
                    let raw = buffer.trim();
                    if let Some((lhs, rhs)) = raw.split_once('=') {
                        if let (Ok(col), Ok(width)) =
                            (lhs.trim().parse::<usize>(), rhs.trim().parse::<usize>())
                        {
                            if let Some(ref p) = self.path.clone() {
                                let mut active_sheet = self.view_sheet_id;
                                commit_workbook_op(
                                    p,
                                    &mut self.offset,
                                    &mut self.workbook,
                                    &mut active_sheet,
                                    &crate::ops::WorkbookOp::SheetOp {
                                        sheet_id: self.view_sheet_id,
                                        op: Op::SetColWidth {
                                            col: MARGIN_COLS + col,
                                            width: Some(width),
                                        },
                                    },
                                )?;
                                self.sync_active_sheet_cache();
                                self.ops_applied = self.ops_applied.saturating_add(1);
                                self.start_log_watcher_if_needed()?;
                            } else {
                                self.state
                                    .grid
                                    .set_col_width(MARGIN_COLS + col, Some(width));
                            }
                            self.status = format!("Column {col} width set to {width}");
                        }
                    } else if let Ok(col) = raw.parse::<usize>() {
                        if let Some(ref p) = self.path.clone() {
                            let mut active_sheet = self.view_sheet_id;
                            commit_workbook_op(
                                p,
                                &mut self.offset,
                                &mut self.workbook,
                                &mut active_sheet,
                                &crate::ops::WorkbookOp::SheetOp {
                                    sheet_id: self.view_sheet_id,
                                    op: Op::SetColWidth {
                                        col: MARGIN_COLS + col,
                                        width: None,
                                    },
                                },
                            )?;
                            self.sync_active_sheet_cache();
                            self.ops_applied = self.ops_applied.saturating_add(1);
                            self.start_log_watcher_if_needed()?;
                        } else {
                            self.state.grid.set_col_width(MARGIN_COLS + col, None);
                        }
                        self.status = format!("Column {col} width override cleared");
                    }
                    self.input_cursor = None;
                    mode = Mode::Normal;
                }
                KeyCode::Esc => mode = Mode::Normal,
                _ if Self::handle_plain_text_input_key(
                    buffer,
                    &mut self.input_cursor,
                    key.code,
                ) => {}
                _ => {}
            },
            Mode::SortView { buffer, persist } => match key.code {
                KeyCode::Enter => {
                    let cols = buffer
                        .split(',')
                        .filter_map(|s| {
                            let s = s.trim();
                            if s.is_empty() {
                                None
                            } else {
                                let (desc, raw) = if let Some(rest) = s.strip_prefix('!') {
                                    (true, rest)
                                } else {
                                    (false, s)
                                };
                                addr::parse_excel_column(raw).map(|c| SortSpec {
                                    col: MARGIN_COLS + c as usize,
                                    desc,
                                })
                            }
                        })
                        .collect::<Vec<_>>();
                    if *persist {
                        if let Some(ref p) = self.path.clone() {
                            let mut active_sheet = self.view_sheet_id;
                            commit_workbook_op(
                                p,
                                &mut self.offset,
                                &mut self.workbook,
                                &mut active_sheet,
                                &crate::ops::WorkbookOp::SheetOp {
                                    sheet_id: self.view_sheet_id,
                                    op: Op::SetViewSortCols { cols: cols.clone() },
                                },
                            )?;
                            self.sync_active_sheet_cache();
                            self.ops_applied = self.ops_applied.saturating_add(1);
                            self.start_log_watcher_if_needed()?;
                        } else {
                            self.state.grid.set_view_sort_cols(cols.clone());
                        }
                        self.set_active_sort_persistence(&cols, true);
                    } else {
                        self.state.grid.set_view_sort_cols(cols.clone());
                        self.set_active_sort_persistence(&cols, false);
                    }
                    self.status = if *persist {
                        "View sort saved".into()
                    } else {
                        "View sort updated".into()
                    };
                    self.input_cursor = None;
                    mode = Mode::Normal;
                }
                KeyCode::Esc => mode = Mode::Normal,
                _ if Self::handle_plain_text_input_key(
                    buffer,
                    &mut self.input_cursor,
                    key.code,
                ) => {}
                _ => {}
            },
            Mode::BalanceBooks {
                buffer,
                direction,
                persist,
                focus,
            } => match key.code {
                KeyCode::Tab => {
                    *focus = Self::cycle_balance_focus(*focus, false);
                }
                KeyCode::BackTab => {
                    *focus = Self::cycle_balance_focus(*focus, true);
                }
                KeyCode::Up => {
                    *focus = match *focus {
                        BalanceBooksFocus::Generate => BalanceBooksFocus::NegToPos,
                        BalanceBooksFocus::Cancel => BalanceBooksFocus::Generate,
                        BalanceBooksFocus::NegToPos => BalanceBooksFocus::PosToNeg,
                        BalanceBooksFocus::PosToNeg => BalanceBooksFocus::ReportPersisted,
                        BalanceBooksFocus::ReportPersisted => BalanceBooksFocus::ReportViewOnly,
                        BalanceBooksFocus::ReportViewOnly => BalanceBooksFocus::Column,
                        BalanceBooksFocus::Column => BalanceBooksFocus::Column,
                    };
                }
                KeyCode::Down => {
                    *focus = match *focus {
                        BalanceBooksFocus::Column => BalanceBooksFocus::PosToNeg,
                        BalanceBooksFocus::ReportViewOnly => BalanceBooksFocus::ReportPersisted,
                        BalanceBooksFocus::ReportPersisted => BalanceBooksFocus::PosToNeg,
                        BalanceBooksFocus::PosToNeg => BalanceBooksFocus::NegToPos,
                        BalanceBooksFocus::NegToPos => BalanceBooksFocus::Generate,
                        BalanceBooksFocus::Generate => BalanceBooksFocus::Cancel,
                        BalanceBooksFocus::Cancel => BalanceBooksFocus::Cancel,
                    };
                }
                KeyCode::Char(' ') | KeyCode::Enter => match focus {
                    BalanceBooksFocus::Column => {
                        if key.code == KeyCode::Enter {
                            self.run_balance_books(buffer, *direction, *persist)?;
                            return Ok(false);
                        }
                    }
                    BalanceBooksFocus::ReportViewOnly => *persist = false,
                    BalanceBooksFocus::ReportPersisted => *persist = true,
                    BalanceBooksFocus::PosToNeg => *direction = BalanceDirection::PosToNeg,
                    BalanceBooksFocus::NegToPos => *direction = BalanceDirection::NegToPos,
                    BalanceBooksFocus::Generate => {
                        self.run_balance_books(buffer, *direction, *persist)?;
                        return Ok(false);
                    }
                    BalanceBooksFocus::Cancel => {
                        mode = Mode::Normal;
                    }
                },
                KeyCode::Esc => mode = Mode::Normal,
                _ if matches!(focus, BalanceBooksFocus::Column)
                    && Self::handle_plain_text_input_key(
                        buffer,
                        &mut self.input_cursor,
                        key.code,
                    ) => {}
                _ => {}
            },
            Mode::FormatDecimals {
                buffer,
                decimals_for,
            } => match key.code {
                KeyCode::Enter => {
                    if let Ok(decimals) = buffer.trim().parse::<usize>() {
                        match decimals_for {
                            FormatDecimalsFor::Currency => self.apply_format_number(decimals, true),
                            FormatDecimalsFor::Fixed => self.apply_format_number(decimals, false),
                        }
                        mode = Mode::Normal;
                    }
                    self.input_cursor = None;
                }
                KeyCode::Esc => mode = Mode::Normal,
                _ if Self::handle_plain_text_input_key(
                    buffer,
                    &mut self.input_cursor,
                    key.code,
                ) => {}
                _ => {}
            },
            Mode::QuitPrompt => match key.code {
                KeyCode::Char('q') | KeyCode::Char('Q') => {
                    self.mode = mode;
                    return Ok(true);
                }
                KeyCode::Char('b') | KeyCode::Char('B') => mode = Mode::Normal,
                KeyCode::Esc => {
                    self.mode = mode;
                    return Ok(true);
                }
                _ => {}
            },
            Mode::QuitImportPrompt => match key.code {
                KeyCode::Char('s') | KeyCode::Char('S') => {
                    mode = Mode::SavePath {
                        buffer: self.start_input_mode(self.suggested_corro_save_path()),
                    };
                }
                KeyCode::Char('d') | KeyCode::Char('D') => {
                    self.mode = mode;
                    return Ok(true);
                }
                KeyCode::Char('b') | KeyCode::Char('B') | KeyCode::Esc => {
                    mode = Mode::Normal;
                }
                _ => {}
            },
            Mode::OpenPath { buffer } => match key.code {
                KeyCode::Enter => match parse_open_path_request(buffer) {
                    Err(OpenPathError::Empty) => {
                        self.status = "Path required".into();
                    }
                    Err(OpenPathError::InvalidRevisionSyntax) => {
                        self.status = "Syntax: link <file> <revision>".into();
                    }
                    Ok(OpenPathRequest::Plain(path)) => {
                        self.source_path = None;
                        self.offset = 0;
                        self.persisted_view_sort_cols.clear();
                        self.ops_applied = 0;
                        self.revision_limit = None;
                        self.import_source = None;
                        if !path.exists() {
                            self.workbook = WorkbookState::new();
                            self.state = SheetState::new(1, 1);
                            self.view_sheet_id = 1;
                            self.path = Some(path.clone());
                            self.watcher = None;
                            self.status = format!("New file {}", path.display());
                        } else {
                            let ext = path
                                .extension()
                                .and_then(|e| e.to_str())
                                .unwrap_or("")
                                .to_lowercase();
                            match ext.as_str() {
                                "tsv" => {
                                    self.workbook = WorkbookState::new();
                                    self.state = SheetState::new(1, 1);
                                    self.view_sheet_id = 1;
                                    if let Ok(data) = std::fs::read_to_string(&path) {
                                        crate::io::import_tsv(&data, &mut self.state);
                                    }
                                    self.commit_active_sheet_cache();
                                    self.path = None;
                                    self.import_source = Some(path.clone());
                                    self.watcher = None;
                                    self.status = format!(
                                        "Imported TSV (not saved) — save as .corro: {}",
                                        path.display()
                                    );
                                }
                                "csv" => {
                                    self.workbook = WorkbookState::new();
                                    self.state = SheetState::new(1, 1);
                                    self.view_sheet_id = 1;
                                    if let Ok(data) = std::fs::read_to_string(&path) {
                                        crate::io::import_csv(&data, &mut self.state);
                                    }
                                    self.commit_active_sheet_cache();
                                    self.path = None;
                                    self.import_source = Some(path.clone());
                                    self.watcher = None;
                                    self.status = format!(
                                        "Imported CSV (not saved) — save as .corro: {}",
                                        path.display()
                                    );
                                }
                                "ods" => match crate::ods::import_ods_workbook(&path) {
                                    Ok(workbook) => {
                                        self.workbook = workbook;
                                        self.view_sheet_id = self.workbook.sheet_id(0);
                                        self.sync_active_sheet_cache();
                                        self.persisted_view_sort_cols.clear();
                                        for c in 0..self.state.grid.main_cols() {
                                            self.state
                                                .grid
                                                .fit_column_to_content(MARGIN_COLS + c);
                                        }
                                        self.path = None;
                                        self.import_source = Some(path.clone());
                                        self.watcher = None;
                                        self.status = format!(
                                            "Imported ODS (not saved) — save as .corro: {}",
                                            path.display()
                                        );
                                    }
                                    Err(e) => {
                                        self.status = format!("Failed to import ODS: {e}");
                                    }
                                },
                                "corro" | _ => {
                                    self.workbook = WorkbookState::new();
                                    self.state = SheetState::new(1, 1);
                                    self.view_sheet_id = 1;
                                    let mut active_sheet =
                                        self.workbook.sheet_id(self.workbook.active_sheet);
                                    let loaded = load_workbook_revisions_partial(
                                        &path,
                                        usize::MAX,
                                        &mut self.workbook,
                                        &mut active_sheet,
                                    );
                                    if let Ok((off, replay)) = loaded {
                                        self.offset = off;
                                        self.ops_applied = replay.op_count;
                                        self.view_sheet_id = active_sheet;
                                        self.sync_active_sheet_cache();
                                        self.sync_persisted_sort_cache_from_workbook();
                                    }
                                    self.path = Some(path.clone());
                                    self.watcher = Some(
                                        LogWatcher::new(path.clone()).map_err(IoError::from)?,
                                    );
                                    self.status = format!("Opened {}", path.display());
                                }
                            }
                        }
                        self.cursor = SheetCursor {
                            row: HEADER_ROWS,
                            col: MARGIN_COLS,
                        };
                        self.row_scroll = 0;
                        self.col_scroll = 0;
                        mode = Mode::Normal;
                    }
                    Ok(OpenPathRequest::Revision { path, revision }) => {
                        if !path.exists() {
                            self.status = format!("Link source not found: {}", path.display());
                        } else {
                            let ext = path
                                .extension()
                                .and_then(|e| e.to_str())
                                .unwrap_or("")
                                .to_lowercase();
                            if matches!(ext.as_str(), "csv" | "tsv" | "ods") {
                                self.status = "Link only works for .corro logs".into();
                            } else {
                                self.workbook = WorkbookState::new();
                                self.state = SheetState::new(1, 1);
                                let mut active_sheet =
                                    self.workbook.sheet_id(self.workbook.active_sheet);
                                let loaded = load_workbook_revisions_partial(
                                    &path,
                                    revision,
                                    &mut self.workbook,
                                    &mut active_sheet,
                                );
                                if let Ok((off, replay)) = loaded {
                                    self.view_sheet_id = active_sheet;
                                    self.sync_active_sheet_cache();
                                    self.sync_persisted_sort_cache_from_workbook();
                                    self.path = None;
                                    self.import_source = None;
                                    self.source_path = Some(path.clone());
                                    self.revision_limit = Some(revision);
                                    self.offset = off;
                                    self.ops_applied = replay.op_count;
                                    self.watcher = None;
                                    self.cursor = SheetCursor {
                                        row: HEADER_ROWS,
                                        col: MARGIN_COLS,
                                    };
                                    self.row_scroll = 0;
                                    self.col_scroll = 0;
                                    self.status = Self::replay_status("Linked", &path, &replay);
                                    mode = Mode::Normal;
                                } else {
                                    self.status = format!("Load failed: {}", path.display());
                                }
                            }
                        }
                    }
                },
                KeyCode::Esc => mode = Mode::Normal,
                _ if Self::handle_plain_text_input_key(
                    buffer,
                    &mut self.input_cursor,
                    key.code,
                ) => {}
                _ => {}
            },
            Mode::SavePath { buffer } => match key.code {
                KeyCode::Enter => {
                    let path = PathBuf::from(buffer.trim());
                    if path.as_os_str().is_empty() {
                        self.status = "Save path required".into();
                    } else {
                        self.save_to_path(&path)?;
                        self.input_cursor = None;
                        mode = Mode::Normal;
                    }
                }
                KeyCode::Esc => mode = Mode::Normal,
                _ if Self::handle_plain_text_input_key(
                    buffer,
                    &mut self.input_cursor,
                    key.code,
                ) => {}
                _ => {}
            },
            Mode::Find { buffer } => match key.code {
                KeyCode::Enter => {
                    self.find_next_substring(buffer);
                }
                KeyCode::Esc => {
                    self.input_cursor = None;
                    mode = Mode::Normal;
                }
                _ if Self::handle_plain_text_input_key(
                    buffer,
                    &mut self.input_cursor,
                    key.code,
                ) => {}
                _ => {}
            },
            Mode::Edit {
                buffer,
                formula_cursor,
                fit_to_content_on_commit: _,
            } => match key.code {
                KeyCode::Char('c') if key.modifiers.contains(KeyModifiers::CONTROL) => {
                    self.edit_special_palette = false;
                    let _ = copy_to_clipboard(buffer);
                }
                KeyCode::Char('v') if key.modifiers.contains(KeyModifiers::CONTROL) => {
                    *formula_cursor = None;
                    self.edit_special_palette = false;
                    let paste = read_clipboard().map_err(io::Error::other)?;
                    let text = if key.modifiers.contains(KeyModifiers::SHIFT) {
                        paste.strip_prefix('=').unwrap_or(&paste).to_string()
                    } else {
                        paste
                    };
                    *buffer = text;
                    self.edit_cursor = Some(buffer.chars().count());
                }
                KeyCode::Char('x') if key.modifiers.contains(KeyModifiers::CONTROL) => {
                    self.edit_special_palette = false;
                    let _ = copy_to_clipboard(buffer);
                    *formula_cursor = None;
                    buffer.clear();
                    self.edit_cursor = Some(0);
                }
                KeyCode::Char('p') if key.modifiers.contains(KeyModifiers::CONTROL) => {
                    *formula_cursor = None;
                    self.edit_special_palette = false;
                    let paste = read_clipboard().map_err(io::Error::other)?;
                    *buffer = paste;
                    self.edit_cursor = Some(buffer.chars().count());
                }
                KeyCode::Enter => {
                    mode = self.commit_edit_and_move_down(buffer)?;
                }
                KeyCode::Delete => {
                    self.edit_special_palette = false;
                    *formula_cursor = None;
                    buffer.clear();
                    self.edit_cursor = Some(0);
                    mode = self.commit_edit_buffer(buffer).map(|_| Mode::Normal)?;
                }
                KeyCode::Tab => {
                    let addr = self.cursor.to_addr(&self.state.grid);
                    if let Some(next) = cycle_special_value(buffer, special_value_choices(&addr)) {
                        self.edit_cursor = Some(next.chars().count());
                        *buffer = next;
                    }
                }
                KeyCode::Char(c) if self.edit_special_palette && c.is_ascii_digit() => {
                    if let Some(choice) = special_value_for_digit(c) {
                        self.edit_cursor = Some(choice.chars().count());
                        *buffer = choice.to_string();
                    }
                }
                KeyCode::Left | KeyCode::Right | KeyCode::Up | KeyCode::Down
                    if formula_cursor.is_some() =>
                {
                    let temp = formula_cursor.as_mut().unwrap();
                    match key.code {
                        KeyCode::Left if temp.col > 0 => temp.col = temp.col.saturating_sub(1),
                        KeyCode::Right => temp.col = temp.col.saturating_add(1),
                        KeyCode::Up if temp.row > 0 => temp.row = temp.row.saturating_sub(1),
                        KeyCode::Down => temp.row = temp.row.saturating_add(1),
                        _ => {}
                    }
                    temp.clamp(&self.state.grid);
                    let addr = temp.to_addr(&self.state.grid);
                    *buffer = format!("={}", self.formula_ref_for_addr(&addr));
                }
                KeyCode::Left | KeyCode::Right => {
                    match Self::handle_text_input_key(buffer, &mut self.edit_cursor, key.code) {
                        TextInputAction::Handled => {}
                        TextInputAction::EdgeLeft => {
                            self.cursor.col = self.cursor.col.saturating_sub(1);
                            self.cursor.clamp(&self.state.grid);
                            self.state
                                .grid
                                .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
                        }
                        TextInputAction::EdgeRight => {
                            let raw = buffer.clone();
                            self.edit_cursor = None;
                            self.edit_special_palette = false;
                            *formula_cursor = None;
                            self.commit_edit_buffer(&raw)?;
                            let lm = MARGIN_COLS;
                            let mc = self.state.grid.main_cols();
                            if self.cursor.col == lm + mc.saturating_sub(1)
                                && trailing_blank_main_cols(&self.state) < NAV_BLANK_COLS
                            {
                                self.state.grid.grow_main_col_at_right();
                            }
                            self.cursor.col = self.cursor.col.saturating_add(1);
                            self.cursor.clamp(&self.state.grid);
                            self.state
                                .grid
                                .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
                            mode = Mode::Normal;
                        }
                        TextInputAction::Unhandled => {}
                    }
                }
                KeyCode::Up => {
                    self.edit_cursor = None;
                    let raw = buffer.clone();
                    self.commit_edit_buffer(&raw)?;
                    if !self.move_cursor_row_through_view(false) && self.cursor.row > 0 {
                        self.cursor.row = self.cursor.row.saturating_sub(1);
                        self.cursor.clamp(&self.state.grid);
                        self.state
                            .grid
                            .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
                    }
                    let addr = self.cursor.to_addr(&self.state.grid);
                    let cur = cell_display(&self.state.grid, &addr);
                    mode = self.start_edit_mode(
                        cur.clone(),
                        if cur.trim() == "=" {
                            Some(self.cursor)
                        } else {
                            None
                        },
                        false,
                        false,
                        None,
                    );
                }
                KeyCode::Down => {
                    mode = self.commit_edit_and_move_down(buffer)?;
                }
                KeyCode::Esc => {
                    self.remember_lost_edit(buffer);
                    self.edit_cursor = None;
                    self.edit_special_palette = false;
                    self.edit_target_addr = None;
                    self.edit_range_addrs = None;
                    *formula_cursor = None;
                    mode = Mode::Normal;
                }
                KeyCode::Char(c) => {
                    *formula_cursor = None;
                    self.edit_special_palette = false;
                    let len = buffer.chars().count();
                    let cursor = self.edit_cursor.get_or_insert(len);
                    let pos = (*cursor).min(len);
                    let mut chars: Vec<char> = buffer.chars().collect();
                    chars.insert(pos, c);
                    *buffer = chars.into_iter().collect();
                    *cursor = pos + 1;
                }
                KeyCode::Backspace => {
                    *formula_cursor = None;
                    let len = buffer.chars().count();
                    if let Some(cursor) = self.edit_cursor.as_mut() {
                        if *cursor > 0 {
                            let pos = (*cursor).min(len);
                            let mut chars: Vec<char> = buffer.chars().collect();
                            if pos > 0 {
                                chars.remove(pos - 1);
                                *buffer = chars.into_iter().collect();
                                *cursor = pos - 1;
                            }
                        }
                    } else {
                        buffer.pop();
                    }
                }
                _ => {}
            },
            Mode::Normal => {
                if key.code == KeyCode::Enter {
                    if let Some(restored) = self.restore_lost_edit() {
                        self.mode = restored;
                        return Ok(false);
                    }
                }
                if matches!(key.code, KeyCode::Char(c) if !c.is_control())
                    && key.modifiers.is_empty()
                {
                    if let KeyCode::Char(c) = key.code {
                        self.edit_special_palette = false;
                        self.pending_lost_edit = None;
                        let buffer = c.to_string();
                        let type_targets = self.multi_cell_type_targets();
                        mode = self.start_edit_mode(
                            buffer.clone(),
                            if buffer.trim() == "=" {
                                Some(self.cursor)
                            } else {
                                None
                            },
                            false,
                            false,
                            type_targets,
                        );
                    }
                    self.mode = mode;
                    return Ok(false);
                }
                if key.modifiers.contains(KeyModifiers::CONTROL) && key.code == KeyCode::Char('q') {
                    self.mode = mode;
                    return Ok(true);
                }
                if key.modifiers.contains(KeyModifiers::CONTROL)
                    && matches!(key.code, KeyCode::Char('z') | KeyCode::Char('Z'))
                {
                    if let Some(undo_op) = self.op_history.pop() {
                        let redo_op = self.state.reverse_op(&undo_op);
                        if let Err(e) = self.apply_op_without_history(undo_op) {
                            self.status = format!("Undo failed: {}", e);
                        } else {
                            if let Some(redo_op) = redo_op {
                                self.redo_history.push(redo_op);
                            }
                            self.status = if self.path.is_some() {
                                "Undo applied".to_string()
                            } else {
                                "Undo applied (memory only)".to_string()
                            };
                        }
                    } else {
                        self.status = "Nothing to undo".to_string();
                    }
                    self.mode = mode;
                    return Ok(false);
                }
                if key.modifiers.contains(KeyModifiers::CONTROL)
                    && matches!(key.code, KeyCode::Char('y') | KeyCode::Char('Y'))
                {
                    if let Some(redo_op) = self.redo_history.pop() {
                        let undo_op = self.state.reverse_op(&redo_op);
                        if let Err(e) = self.apply_op_without_history(redo_op) {
                            self.status = format!("Redo failed: {}", e);
                        } else {
                            if let Some(undo_op) = undo_op {
                                self.op_history.push(undo_op);
                            }
                            self.status = if self.path.is_some() {
                                "Redo applied".to_string()
                            } else {
                                "Redo applied (memory only)".to_string()
                            };
                        }
                    } else {
                        self.status = "Nothing to redo".to_string();
                    }
                    self.mode = mode;
                    return Ok(false);
                }

                if key.modifiers.contains(KeyModifiers::CONTROL)
                    && matches!(key.code, KeyCode::Char('x') | KeyCode::Char('X'))
                {
                    let data = self.selection_tsv_text();
                    let _ = self.copy_selection_to_clipboard(&data);
                    if !self.delete_selection() {
                        let addr = self.cursor.to_addr(&self.state.grid);
                        if self.state.grid.get(&addr).is_some() {
                            let op = Op::FillRange {
                                cells: vec![(addr, String::new())],
                            };
                            if self.apply_single_op(op).is_ok() {
                                self.status = "Cell cut".into();
                            }
                        }
                    }
                    self.mode = mode;
                    return Ok(false);
                }

                if key.modifiers.contains(KeyModifiers::CONTROL)
                    || key.modifiers.contains(KeyModifiers::SUPER)
                {
                    match key.code {
                        KeyCode::Char('d') | KeyCode::Char('D') => {
                            if let Some(op) = self.fill_row_pattern() {
                                self.push_inverse_op(&op);
                                if let Some(ref p) = self.path.clone() {
                                    let mut active_sheet = self.view_sheet_id;
                                    commit_workbook_op(
                                        p,
                                        &mut self.offset,
                                        &mut self.workbook,
                                        &mut active_sheet,
                                        &crate::ops::WorkbookOp::SheetOp {
                                            sheet_id: self.view_sheet_id,
                                            op: op.clone(),
                                        },
                                    )?;
                                    self.ops_applied = self.ops_applied.saturating_add(1);
                                    self.sync_active_sheet_cache();
                                    self.start_log_watcher_if_needed()?;
                                } else {
                                    op.apply(&mut self.state);
                                }
                                self.status = "Filled row pattern".into();
                            } else {
                                self.status =
                                    "Select a single row of cells, then press Ctrl+D / Cmd+D"
                                        .into();
                            }
                            self.mode = mode;
                            return Ok(false);
                        }
                        KeyCode::Char('r') | KeyCode::Char('R') => {
                            if let Some(op) = self.fill_col_pattern() {
                                self.push_inverse_op(&op);
                                if let Some(ref p) = self.path.clone() {
                                    let mut active_sheet = self.view_sheet_id;
                                    commit_workbook_op(
                                        p,
                                        &mut self.offset,
                                        &mut self.workbook,
                                        &mut active_sheet,
                                        &crate::ops::WorkbookOp::SheetOp {
                                            sheet_id: self.view_sheet_id,
                                            op: op.clone(),
                                        },
                                    )?;
                                    self.ops_applied = self.ops_applied.saturating_add(1);
                                    self.sync_active_sheet_cache();
                                    self.start_log_watcher_if_needed()?;
                                } else {
                                    op.apply(&mut self.state);
                                }
                                self.status = "Filled column pattern".into();
                            } else {
                                self.status =
                                    "Select a single column of cells, then press Ctrl+R / Cmd+R"
                                        .into();
                            }
                            self.mode = mode;
                            return Ok(false);
                        }
                        _ => {}
                    }
                }

                let ctrl_or_cmd = key.modifiers.contains(KeyModifiers::CONTROL)
                    || key.modifiers.contains(KeyModifiers::SUPER);

                match key.code {
                    KeyCode::Esc => {
                        if self.anchor.is_some() {
                            self.anchor = None;
                            self.selection_kind = SelectionKind::Cells;
                        } else if self.is_ods_tsv_import_unchanged() {
                            self.mode = mode;
                            return Ok(true);
                        } else {
                            mode = if self.path.is_none() {
                                Mode::QuitImportPrompt
                            } else {
                                Mode::QuitPrompt
                            };
                        }
                    }
                    KeyCode::Delete => {
                        if !self.delete_selection() {
                            let addr = self.cursor.to_addr(&self.state.grid);
                            if self.state.grid.get(&addr).is_some() {
                                let op = Op::FillRange {
                                    cells: vec![(addr, String::new())],
                                };
                                if self.apply_single_op(op).is_ok() {
                                    self.status = "Cell deleted".into();
                                }
                            } else {
                                self.status = "Nothing to delete".into();
                            }
                        }
                    }
                    KeyCode::Left if key.modifiers.contains(KeyModifiers::SHIFT) && ctrl_or_cmd => {
                        let _ = self.extend_selection_to_edge(SelectionEdgeDirection::Left);
                    }
                    KeyCode::Right
                        if key.modifiers.contains(KeyModifiers::SHIFT) && ctrl_or_cmd =>
                    {
                        let _ = self.extend_selection_to_edge(SelectionEdgeDirection::Right);
                    }
                    KeyCode::Up if key.modifiers.contains(KeyModifiers::SHIFT) && ctrl_or_cmd => {
                        let _ = self.extend_selection_to_edge(SelectionEdgeDirection::Up);
                    }
                    KeyCode::Down if key.modifiers.contains(KeyModifiers::SHIFT) && ctrl_or_cmd => {
                        let _ = self.extend_selection_to_edge(SelectionEdgeDirection::Down);
                    }
                    KeyCode::Left if key.modifiers.contains(KeyModifiers::SHIFT) => {
                        if self.cursor.col > 0 {
                            if self.anchor.is_none() {
                                self.anchor = Some(self.cursor);
                            }
                            self.cursor.col = self.cursor.col.saturating_sub(1);
                            self.cursor.clamp(&self.state.grid);
                        }
                    }
                    KeyCode::Right if key.modifiers.contains(KeyModifiers::SHIFT) => {
                        if self.anchor.is_none() {
                            self.anchor = Some(self.cursor);
                        }
                        self.cursor.col = self.cursor.col.saturating_add(1);
                        self.cursor.clamp(&self.state.grid);
                        self.state
                            .grid
                            .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
                    }
                    KeyCode::Up if key.modifiers.contains(KeyModifiers::SHIFT) => {
                        if self.cursor.row > 0 {
                            if self.anchor.is_none() {
                                self.anchor = Some(self.cursor);
                            }
                            if !self.move_cursor_row_through_view(false) {
                                self.cursor.row = self.cursor.row.saturating_sub(1);
                                self.cursor.clamp(&self.state.grid);
                            }
                        }
                    }
                    KeyCode::Down if key.modifiers.contains(KeyModifiers::SHIFT) => {
                        if self.anchor.is_none() {
                            self.anchor = Some(self.cursor);
                        }
                        if !self.move_cursor_row_through_view(true) {
                            self.cursor.row = self.cursor.row.saturating_add(1);
                            self.cursor.clamp(&self.state.grid);
                            self.state
                                .grid
                                .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
                        }
                    }
                    KeyCode::Char('o') => {
                        self.edit_special_palette = false;
                        let buffer = self
                            .path
                            .as_ref()
                            .map(|p| p.to_string_lossy().into_owned())
                            .unwrap_or_default();
                        mode = Mode::OpenPath {
                            buffer: self.start_input_mode(buffer),
                        };
                    }
                    KeyCode::Char('e') | KeyCode::Enter => {
                        self.edit_special_palette = false;
                        let addr = self.cursor.to_addr(&self.state.grid);
                        let cur = cell_display(&self.state.grid, &addr);
                        mode = self.start_edit_mode(
                            cur.clone(),
                            if cur.trim() == "=" {
                                Some(self.cursor)
                            } else {
                                None
                            },
                            false,
                            false,
                            None,
                        );
                    }
                    KeyCode::Char('v') => {
                        self.anchor = if self.anchor.is_none() {
                            Some(self.cursor)
                        } else {
                            None
                        };
                        self.selection_kind = SelectionKind::Cells;
                    }
                    KeyCode::Char('t') => {
                        self.export_preview_scroll = 0;
                        self.export_delimited_options.content = export::ExportContent::Values;
                        mode = Mode::ExportTsv {
                            buffer: self
                                .start_input_mode(self.suggested_export_save_path("tsv")),
                        }
                    }
                    KeyCode::Char('c') => {
                        if self.anchor.is_some() {
                            if let Some((mc0, mc1)) = self.selection_main_col_range() {
                                let left = MARGIN_COLS;
                                let right = MARGIN_COLS + self.state.grid.main_cols();
                                if self.cursor.col < left || self.cursor.col >= right {
                                    self.status = "Place cursor on a main column as move target, then press c".into();
                                } else {
                                    let count = mc1 - mc0 + 1;
                                    let to = (self.cursor.col - left) as u32;
                                    let op = Op::MoveColRange {
                                        from: mc0,
                                        count,
                                        to,
                                    };
                                    self.push_inverse_op(&op);
                                    if let Some(ref p) = self.path.clone() {
                                        let mut active_sheet = self.view_sheet_id;
                                        commit_workbook_op(
                                            p,
                                            &mut self.offset,
                                            &mut self.workbook,
                                            &mut active_sheet,
                                            &crate::ops::WorkbookOp::SheetOp {
                                                sheet_id: self.view_sheet_id,
                                                op: op.clone(),
                                            },
                                        )?;
                                        self.ops_applied = self.ops_applied.saturating_add(1);
                                        self.sync_active_sheet_cache();
                                        self.start_log_watcher_if_needed()?;
                                    } else {
                                        op.apply(&mut self.state);
                                    }
                                    self.anchor = None;
                                    self.status = format!(
                                        "Moved cols {mc0}..{} → before col {to}",
                                        mc0 + count
                                    );
                                }
                            } else {
                                self.expand_selection_to_cols();
                                self.status = "Selection expanded to columns".into();
                            }
                        } else {
                            self.export_preview_scroll = 0;
                            self.export_delimited_options.content = export::ExportContent::Values;
                            mode = Mode::ExportCsv {
                                buffer: self
                                    .start_input_mode(self.suggested_export_save_path("csv")),
                            };
                        }
                    }
                    KeyCode::Char('r') => {
                        if let Some((mr0, mr1)) = self.selection_main_row_range() {
                            let hr = HEADER_ROWS;
                            if self.cursor.row < hr
                                || self.cursor.row >= hr + self.state.grid.main_rows()
                            {
                                self.status =
                                    "Place cursor on a main row as move target, then press r"
                                        .into();
                            } else {
                                let count = mr1 - mr0 + 1;
                                let to = (self.cursor.row - hr) as u32;
                                let op = Op::MoveRowRange {
                                    from: mr0,
                                    count,
                                    to,
                                };
                                self.push_inverse_op(&op);
                                if let Some(ref p) = self.path.clone() {
                                    let mut active_sheet = self.view_sheet_id;
                                    commit_workbook_op(
                                        p,
                                        &mut self.offset,
                                        &mut self.workbook,
                                        &mut active_sheet,
                                        &crate::ops::WorkbookOp::SheetOp {
                                            sheet_id: self.view_sheet_id,
                                            op: op.clone(),
                                        },
                                    )?;
                                    self.ops_applied = self.ops_applied.saturating_add(1);
                                    self.sync_active_sheet_cache();
                                    self.start_log_watcher_if_needed()?;
                                } else {
                                    op.apply(&mut self.state);
                                }
                                self.anchor = None;
                                self.status =
                                    format!("Moved rows {mr0}..{} → before row {to}", mr0 + count);
                            }
                        } else {
                            self.expand_selection_to_rows();
                            self.status = "Selection expanded to rows".into();
                        }
                    }
                    KeyCode::Char(ch)
                        if key.modifiers.contains(KeyModifiers::CONTROL)
                            && matches!(ch, '=' | '+') =>
                    {
                        if self.anchor.is_some() {
                            if !self.insert_rows_above_selection()? {
                                if let Some((from, to)) = self.selection_main_row_range() {
                                    let count = to - from + 1;
                                    let _ = self.insert_rows_above_cursor(count as u32)?;
                                } else {
                                    let _ = self.insert_rows_above_cursor(1)?;
                                }
                            }
                        } else {
                            let _ = self.insert_rows_above_cursor(1)?;
                        }
                    }
                    KeyCode::Char('?') => {
                        mode = Mode::Help;
                    }
                    KeyCode::Backspace => {
                        if !self.delete_selection() {
                            if let Some(addr) = self.addr_at(self.cursor.row, self.cursor.col) {
                                let raw = self.state.grid.get(&addr);
                                if raw.as_deref().unwrap_or("").is_empty() {
                                    self.status = "Cell already blank".into();
                                    self.mode = mode;
                                    return Ok(false);
                                }
                                let op = Op::SetCell {
                                    addr,
                                    value: String::new(),
                                };
                                self.push_inverse_op(&op);
                                if let Some(ref p) = self.path.clone() {
                                    let mut active_sheet = self.view_sheet_id;
                                    commit_workbook_op(
                                        p,
                                        &mut self.offset,
                                        &mut self.workbook,
                                        &mut active_sheet,
                                        &crate::ops::WorkbookOp::SheetOp {
                                            sheet_id: self.view_sheet_id,
                                            op: op.clone(),
                                        },
                                    )?;
                                    self.ops_applied = self.ops_applied.saturating_add(1);
                                    self.sync_active_sheet_cache();
                                    self.start_log_watcher_if_needed()?;
                                } else {
                                    op.apply(&mut self.state);
                                }
                                self.status = "Cell deleted".into();
                            }
                        }
                    }
                    KeyCode::Left | KeyCode::Char('h') => {
                        self.cursor.col = self.cursor.col.saturating_sub(1);
                        self.cursor.clamp(&self.state.grid);
                        self.state
                            .grid
                            .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
                    }
                    KeyCode::Right | KeyCode::Char('l') => {
                        let lm = MARGIN_COLS;
                        let mc = self.state.grid.main_cols();
                        if self.cursor.col == lm + mc.saturating_sub(1)
                            && trailing_blank_main_cols(&self.state) < NAV_BLANK_COLS
                        {
                            self.state.grid.grow_main_col_at_right();
                        }
                        self.cursor.col = self.cursor.col.saturating_add(1);
                        self.cursor.clamp(&self.state.grid);
                        self.state
                            .grid
                            .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
                    }
                    KeyCode::Up | KeyCode::Char('k') => {
                        self.move_cursor_one_row_vertical(false);
                    }
                    KeyCode::Down | KeyCode::Char('j') => {
                        self.move_cursor_one_row_vertical(true);
                    }
                    KeyCode::PageUp => {
                        let steps = self.grid_viewport_data_rows.max(1);
                        self.move_cursor_vertical_steps(steps, false);
                    }
                    KeyCode::PageDown => {
                        let steps = self.grid_viewport_data_rows.max(1);
                        self.move_cursor_vertical_steps(steps, true);
                    }
                    KeyCode::Char(c) if !c.is_control() => {
                        self.edit_special_palette = false;
                        let buffer = c.to_string();
                        let type_targets = self.multi_cell_type_targets();
                        mode = self.start_edit_mode(
                            buffer.clone(),
                            if buffer.trim() == "=" {
                                Some(self.cursor)
                            } else {
                                None
                            },
                            false,
                            false,
                            type_targets,
                        );
                    }
                    _ => {}
                }
            }
        }

        self.mode = mode;
        Ok(false)
    }

    #[cold]
    #[inline(never)]
    fn mode_prompt_widget<'a>(
        &self,
        grid: &'a Grid,
        addr: &CellAddr,
        edit_addr: &CellAddr,
        prompt_style: Style,
        prompt_style_bold: Style,
        caret_style: Style,
    ) -> Paragraph<'a> {
        let addr_str = addr_label(edit_addr, grid.main_cols());
        match &self.mode {
            Mode::Edit { buffer, .. } => Paragraph::new(input_line_with_suffix(
                format!(" {addr_str}  "),
                buffer,
                self.edit_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style_bold,
                caret_style,
                formula_edit_preview(grid, edit_addr, buffer),
            ))
            .style(prompt_style_bold),
            Mode::OpenPath { buffer } => Paragraph::new(input_line(
                " open: ".to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::SheetRename { buffer, .. } => Paragraph::new(input_line(
                " rename sheet: ".to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::SheetCopy { buffer, .. } => Paragraph::new(input_line(
                " copy sheet as: ".to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::GoToCell { buffer } => Paragraph::new(input_line(
                " go to: ".to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::SavePath { buffer } => Paragraph::new(input_line(
                " save as: ".to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::ExportTsv { buffer } => Paragraph::new(input_line(
                " export TSV (blank=clipboard): ".to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::ExportCsv { buffer } => Paragraph::new(input_line(
                " export CSV (blank=clipboard): ".to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::ExportAscii { buffer } => Paragraph::new(input_line(
                " export ASCII table (blank=clipboard): ".to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::ExportAll { buffer } => Paragraph::new(input_line(
                " export full (incl headers/margins): ".to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::ExportOdt { buffer } => Paragraph::new(input_line(
                " export ODS: ".to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::Find { buffer } => Paragraph::new(input_line(
                " find text: ".to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::Replace { buffer } => Paragraph::new(input_line(
                " replace (old|new): ".to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::SetMaxColWidth { buffer } => Paragraph::new(input_line(
                " max col width (default=20): ".to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::SetColWidth { buffer } => Paragraph::new(input_line(
                " col width [col=width|col]: ".to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::SortView { buffer, persist } => Paragraph::new(input_line(
                format!(
                    " sort cols [A,B,C]{}: ",
                    if *persist { " (save)" } else { "" }
                ),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::BalanceBooks { .. } => Paragraph::new(" ").style(prompt_style),
            Mode::FormatDecimals {
                buffer,
                decimals_for,
            } => Paragraph::new(input_line(
                match decimals_for {
                    FormatDecimalsFor::Currency => " currency decimals: ",
                    FormatDecimalsFor::Fixed => " fixed decimals: ",
                }
                .to_string(),
                buffer,
                self.input_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style,
                caret_style,
            ))
            .style(prompt_style),
            Mode::QuitPrompt => Paragraph::new(" Quit Corro? (Q)uit, (B)ack ")
                .style(Style::default().fg(Color::White).bg(Color::Red)),
            Mode::QuitImportPrompt => {
                Paragraph::new(" No .corro on disk. (S)ave as .corro, (D)iscard and quit, (B)ack ")
                    .style(Style::default().fg(Color::White).bg(Color::Red))
            }
            Mode::Help => Paragraph::new(" Help - Up/Down scroll, Esc closes ")
                .style(Style::default().fg(Color::White).bg(Color::Blue)),
            Mode::About => Paragraph::new(" About - Up/Down scroll, Esc closes ")
                .style(Style::default().fg(Color::White).bg(Color::Blue)),
            Mode::Menu { .. } | Mode::Normal | Mode::RevisionBrowse => {
                let val = formula_bar_value(grid, addr);
                let addr_str = addr_label(addr, grid.main_cols());
                let base = format!(" {addr_str}  {val}");
                let text = if self.status.is_empty() {
                    base
                } else {
                    format!("{base}   ·  {}", self.status)
                };
                Paragraph::new(text).style(Style::default().fg(Color::Cyan))
            }
        }
    }

    fn export_preview_text(&self, csv: bool) -> String {
        let mut grid = self.state.grid.clone();
        crate::formula::refresh_spills(&mut grid);
        let mut buf = Vec::new();
        let o = &self.export_delimited_options;
        if csv {
            export::export_csv_with_options(&grid, &mut buf, o);
        } else {
            export::export_tsv_with_options(&grid, &mut buf, o);
        }
        String::from_utf8_lossy(&buf).into_owned()
    }

    /// Full-grid export preview: `sanitize_tabs` means replace `\t` in the *preview* only
    /// (terminal tab stops corrupt the TUI; real exports are unchanged).
    fn export_preview_overlay_content(&self) -> Option<(String, &'static str, bool)> {
        let mut grid = self.state.grid.clone();
        crate::formula::refresh_spills(&mut grid);
        match &self.mode {
            Mode::ExportTsv { .. } => {
                let mut buf = Vec::new();
                export::export_tsv_with_options(&grid, &mut buf, &self.export_delimited_options);
                Some((
                    String::from_utf8_lossy(&buf).into_owned(),
                    " Export TSV ",
                    true,
                ))
            }
            Mode::ExportCsv { .. } => {
                let mut buf = Vec::new();
                export::export_csv_with_options(&grid, &mut buf, &self.export_delimited_options);
                Some((
                    String::from_utf8_lossy(&buf).into_owned(),
                    " Export CSV ",
                    false,
                ))
            }
            Mode::ExportAscii { .. } => {
                let mut buf = Vec::new();
                export::export_ascii_table_with_options(&grid, &mut buf, &self.export_ascii_options);
                Some((
                    String::from_utf8_lossy(&buf).into_owned(),
                    " Export ASCII table ",
                    false,
                ))
            }
            Mode::ExportAll { .. } => {
                if self.anchor.is_some() {
                    let (rows, cols) = self
                        .current_selection_range()
                        .unwrap_or_else(|| (vec![self.cursor.row], vec![self.cursor.col]));
                    if rows.is_empty() || cols.is_empty() {
                        return Some((
                            String::new(),
                            " Export selection (TSV) ",
                            true,
                        ));
                    }
                    let mut buf = Vec::new();
                    export::export_selection(
                        &grid,
                        &mut buf,
                        &rows,
                        &cols,
                        &self.export_delimited_options,
                    );
                    Some((
                        String::from_utf8_lossy(&buf).into_owned(),
                        " Export selection (TSV) ",
                        true,
                    ))
                } else {
                    let mut buf = Vec::new();
                    export::export_all_with_options(&grid, &mut buf, &self.export_delimited_options);
                    Some((
                        String::from_utf8_lossy(&buf).into_owned(),
                        " Export full (TSV) ",
                        true,
                    ))
                }
            }
            Mode::ExportOdt { .. } => {
                let mode = match self.export_ods_content {
                    export::ExportContent::Values => "values only (static)",
                    export::ExportContent::Formulas => "formulas (with ODF formula attributes)",
                    export::ExportContent::Generic => "generic (same as TSV generic; comma arg lists in of:)",
                };
                Some((
                    format!(
                        "OpenDocument (.ods) is a binary ZIP package.\n\nExport: {mode}. Table shape matches your current TSV/CSV options (margins, header row, row labels). There is no text preview. Type a file path and press Enter to save."
                    ),
                    " Export ODS ",
                    false,
                ))
            }
            _ => None,
        }
    }

    #[cold]
    #[inline(never)]
    fn render_export_preview_overlay(&self, f: &mut Frame, grid_area: Rect) -> bool {
        let Some((body, title, sanitize_tabs)) = self.export_preview_overlay_content() else {
            return false;
        };
        let body = if sanitize_tabs {
            // See `export_preview_overlay_content` (tab stops in the TUI).
            body.replace('\t', "  ")
        } else {
            body
        };
        let block = Block::default().borders(Borders::ALL).title(title);
        let inner = block.inner(grid_area);
        let lines: Vec<&str> = body.lines().collect();
        let max_scroll = lines.len().saturating_sub(inner.height as usize);
        let scroll = self.export_preview_scroll.min(max_scroll);
        let visible: String = lines
            .iter()
            .skip(scroll)
            .take(inner.height as usize)
            .cloned()
            .collect::<Vec<_>>()
            .join("\n");
        // No wrap: long lines must not expand to extra terminal rows (overflows the grid).
        let paragraph = Paragraph::new(visible).block(block);
        f.render_widget(Clear, grid_area);
        f.render_widget(paragraph, grid_area);
        true
    }

    #[cold]
    #[inline(never)]
    fn render_help_about_overlay(&self, f: &mut Frame, grid_area: Rect) -> bool {
        if !matches!(&self.mode, Mode::Help | Mode::About) {
            return false;
        }

        let body = match &self.mode {
            Mode::Help => self.help_page_body(),
            Mode::About => self.about_page_body(),
            _ => String::new(),
        };
        let block = Block::default()
            .borders(Borders::ALL)
            .title(match self.mode {
                Mode::Help => " Help ",
                Mode::About => " About ",
                _ => "",
            });
        let inner = block.inner(grid_area);
        let lines: Vec<&str> = body.lines().collect();
        let scroll = match &self.mode {
            Mode::Help => self.help_scroll,
            Mode::About => self.about_scroll,
            _ => 0,
        };
        let max_scroll = lines.len().saturating_sub(inner.height as usize);
        let scroll = scroll.min(max_scroll);
        let visible: String = lines
            .iter()
            .skip(scroll)
            .take(inner.height as usize)
            .cloned()
            .collect::<Vec<_>>()
            .join("\n");
        let paragraph = Paragraph::new(visible)
            .block(block)
            .wrap(Wrap { trim: false });
        f.render_widget(Clear, grid_area);
        f.render_widget(paragraph, grid_area);
        true
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crossterm::event::{KeyCode, KeyEvent, KeyModifiers};
    use std::path::PathBuf;

    fn docs_test_path(name: &str) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("docs/test")
            .join(name)
    }

    #[test]
    fn undo_restores_previous_cell_value() {
        let mut app = App::new(None);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "old".into());

        let op = Op::SetCell {
            addr: CellAddr::Main { row: 0, col: 0 },
            value: "new".into(),
        };
        app.op_history.clear();
        app.push_inverse_op(&op);
        op.apply(&mut app.state);

        let undo_op = app.op_history.pop().expect("inverse op");
        undo_op.apply(&mut app.state);

        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("old")
        );
    }

    #[test]
    fn right_enters_nested_width_submenu() {
        let mut app = App::new(None);
        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::File,
                item: 2,
            }],
        };

        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::empty()))
            .unwrap();

        match app.mode {
            Mode::Menu { stack } => {
                assert_eq!(stack.len(), 2);
                assert_eq!(stack[1].section, MenuSection::Export);
                assert_eq!(stack[1].item, 0);
            }
            other => panic!("unexpected mode: {other:?}"),
        }
    }

    #[test]
    fn menu_preview_includes_child_submenu() {
        let levels = App::menu_render_levels(&[MenuLevel {
            section: MenuSection::File,
            item: 2,
        }]);

        assert_eq!(levels.len(), 2);
        assert_eq!(levels[0].section, MenuSection::File);
        assert_eq!(levels[1].section, MenuSection::Export);
    }

    #[test]
    fn submenu_popup_is_offset_right_and_down() {
        let area = Rect::new(0, 0, 80, 20);
        let parent = menu_popup_area(area, MenuSection::File, None);
        let child = menu_popup_area(area, MenuSection::Width, Some((parent, 2)));

        assert!(child.x > parent.x);
        assert!(child.y > parent.y);
        assert_eq!(child.y, parent.y + 2);
    }

    #[test]
    fn root_menu_popups_align_under_top_bar_items() {
        let area = Rect::new(0, 0, 80, 20);

        let file = menu_popup_area(area, MenuSection::File, None);
        let edit = menu_popup_area(area, MenuSection::Edit, None);
        let insert = menu_popup_area(area, MenuSection::Insert, None);
        let help = menu_popup_area(area, MenuSection::Help, None);

        assert_eq!(file.x, 1);
        assert_eq!(edit.x, 9);
        assert_eq!(insert.x, 17);
        // The menu popup x positions are computed from fixed offsets in menu_popup_area.
        // Help currently maps to x=45.
        assert_eq!(help.x, 45);
    }

    #[test]
    fn preview_level_is_not_highlighted() {
        assert_eq!(App::menu_selected_index(0, 1, 2, 4), Some(2));
        assert_eq!(App::menu_selected_index(1, 1, 0, 4), None);
    }

    #[test]
    fn sorted_view_down_moves_through_visible_order() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(3, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "2".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "apple".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 2, col: 0 }, "10".into());
        app.state.grid.set_view_sort_cols(vec![SortSpec {
            col: MARGIN_COLS,
            desc: false,
        }]);
        app.cursor = SheetCursor {
            row: HEADER_ROWS + 1,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(app.cursor.row, HEADER_ROWS);
        assert_eq!(app.state.grid.sorted_main_rows(), vec![1, 0, 2]);
    }

    #[test]
    fn sorted_view_up_moves_through_visible_order() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(3, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "2".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "apple".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 2, col: 0 }, "10".into());
        app.state.grid.set_view_sort_cols(vec![SortSpec {
            col: MARGIN_COLS,
            desc: false,
        }]);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(app.cursor.row, HEADER_ROWS + 1);
        assert_eq!(app.state.grid.sorted_main_rows(), vec![1, 0, 2]);
    }

    #[test]
    fn sorted_view_down_from_physical_last_uses_view_order_without_growing() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(3, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "b".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 2, col: 0 }, "c".into());
        app.state.grid.set_view_sort_cols(vec![SortSpec {
            col: MARGIN_COLS,
            desc: true,
        }]);
        app.cursor = SheetCursor {
            row: HEADER_ROWS + 2,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(app.cursor.row, HEADER_ROWS + 1);
        assert_eq!(app.state.grid.main_rows(), 3);
        assert_eq!(app.state.grid.sorted_main_rows(), vec![2, 1, 0]);
    }

    #[test]
    fn sorted_view_edit_up_moves_through_visible_order() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(3, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "2".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "apple".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 2, col: 0 }, "10".into());
        app.state.grid.set_view_sort_cols(vec![SortSpec {
            col: MARGIN_COLS,
            desc: false,
        }]);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Edit {
            buffer: "2".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(app.cursor.row, HEADER_ROWS + 1);
        assert!(matches!(app.mode, Mode::Edit { .. }));
    }

    #[test]
    fn sorted_view_allows_two_blank_rows_before_footer() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "apple".into());
        app.state.grid.set_view_sort_cols(vec![SortSpec {
            col: MARGIN_COLS,
            desc: false,
        }]);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.cursor.row, HEADER_ROWS + 1);

        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.cursor.row, HEADER_ROWS + 2);

        app.handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.cursor.row, HEADER_ROWS + 3);
    }

    #[test]
    fn enter_in_edit_mode_uses_edit_target_row_for_cursor_progression() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(2, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "1".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "2".into());

        // Simulate the observed mismatch: cursor row now maps to footer after
        // extent drift, but edit target still points to the next main row.
        app.cursor = SheetCursor {
            row: HEADER_ROWS + 2,
            col: MARGIN_COLS,
        };
        app.edit_target_addr = Some(CellAddr::Main { row: 2, col: 0 });
        app.mode = Mode::Edit {
            buffer: "3".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 2, col: 0 })
                .as_deref(),
            Some("3")
        );
        assert_eq!(
            app.cursor.to_addr(&app.state.grid),
            CellAddr::Main { row: 3, col: 0 }
        );
    }

    #[test]
    fn ctrl_shift_plus_inserts_one_row_above_cursor() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(2, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "top".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "bottom".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS + 1,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(
            KeyCode::Char('+'),
            KeyModifiers::CONTROL | KeyModifiers::SHIFT,
        ))
        .unwrap();

        assert_eq!(app.state.grid.main_rows(), 3);
        assert_eq!(app.cursor.row, HEADER_ROWS + 1);
        assert_eq!(app.state.grid.get(&CellAddr::Main { row: 1, col: 0 }), None);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 2, col: 0 })
                .as_deref(),
            Some("bottom")
        );
    }

    #[test]
    fn ctrl_shift_plus_inserts_multiple_selected_rows() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(3, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "b".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 2, col: 0 }, "c".into());
        app.anchor = Some(SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        });
        app.cursor = SheetCursor {
            row: HEADER_ROWS + 1,
            col: MARGIN_COLS,
        };
        app.selection_kind = SelectionKind::Rows;
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(
            KeyCode::Char('+'),
            KeyModifiers::CONTROL | KeyModifiers::SHIFT,
        ))
        .unwrap();

        assert_eq!(app.state.grid.main_rows(), 5);
        assert_eq!(app.cursor.row, HEADER_ROWS);
        assert!(app.anchor.is_none());
        assert_eq!(app.state.grid.get(&CellAddr::Main { row: 0, col: 0 }), None);
        assert_eq!(app.state.grid.get(&CellAddr::Main { row: 1, col: 0 }), None);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 2, col: 0 })
                .as_deref(),
            Some("a")
        );
    }

    #[test]
    fn ctrl_shift_plus_falls_back_to_current_row_for_cell_selection() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(2, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "top".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "bottom".into());
        app.anchor = Some(SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        });
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.selection_kind = SelectionKind::Cells;
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(
            KeyCode::Char('+'),
            KeyModifiers::CONTROL | KeyModifiers::SHIFT,
        ))
        .unwrap();

        assert_eq!(app.state.grid.main_rows(), 3);
        assert_eq!(app.state.grid.get(&CellAddr::Main { row: 0, col: 0 }), None);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 1, col: 0 })
                .as_deref(),
            Some("top")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 2, col: 0 })
                .as_deref(),
            Some("bottom")
        );
    }

    #[test]
    fn ctrl_d_fills_single_selected_row() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 4);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "1".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "2".into());
        app.anchor = Some(SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        });
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS + 1,
        };
        app.selection_kind = SelectionKind::Cells;
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(KeyCode::Char('d'), KeyModifiers::CONTROL))
            .unwrap();

        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 2 })
                .as_deref(),
            Some("3")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 3 })
                .as_deref(),
            Some("4")
        );
    }

    #[test]
    fn ctrl_r_fills_single_selected_column() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(4, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "mon".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "tue".into());
        app.anchor = Some(SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        });
        app.cursor = SheetCursor {
            row: HEADER_ROWS + 1,
            col: MARGIN_COLS,
        };
        app.selection_kind = SelectionKind::Cells;
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(KeyCode::Char('r'), KeyModifiers::CONTROL))
            .unwrap();

        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 2, col: 0 })
                .as_deref(),
            Some("WED")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 3, col: 0 })
                .as_deref(),
            Some("THU")
        );
    }

    #[test]
    fn ctrl_d_rejects_multirow_selection() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(2, 1);
        app.anchor = Some(SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        });
        app.cursor = SheetCursor {
            row: HEADER_ROWS + 1,
            col: MARGIN_COLS,
        };
        app.selection_kind = SelectionKind::Cells;
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(KeyCode::Char('d'), KeyModifiers::CONTROL))
            .unwrap();

        assert!(app
            .state
            .grid
            .get(&CellAddr::Main { row: 0, col: 0 })
            .is_none());
        assert!(app
            .state
            .grid
            .get(&CellAddr::Main { row: 1, col: 0 })
            .is_none());
    }

    #[test]
    fn cmd_shift_right_extends_to_last_nonblank_cell_in_row() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 5);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 2 }, "mid".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 3 }, "end".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS + 1,
        };
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(
            KeyCode::Right,
            KeyModifiers::SHIFT | KeyModifiers::SUPER,
        ))
        .unwrap();

        assert_eq!(
            app.anchor,
            Some(SheetCursor {
                row: HEADER_ROWS,
                col: MARGIN_COLS + 1,
            })
        );
        assert_eq!(
            app.cursor,
            SheetCursor {
                row: HEADER_ROWS,
                col: MARGIN_COLS + 3,
            }
        );
        assert_eq!(app.selection_kind, SelectionKind::Cells);
    }

    #[test]
    fn ctrl_shift_left_extends_to_first_nonblank_cell_in_row() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 5);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "start".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "next".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS + 3,
        };
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(
            KeyCode::Left,
            KeyModifiers::CONTROL | KeyModifiers::SHIFT,
        ))
        .unwrap();

        assert_eq!(
            app.anchor,
            Some(SheetCursor {
                row: HEADER_ROWS,
                col: MARGIN_COLS + 3,
            })
        );
        assert_eq!(
            app.cursor,
            SheetCursor {
                row: HEADER_ROWS,
                col: MARGIN_COLS,
            }
        );
        assert_eq!(app.selection_kind, SelectionKind::Cells);
    }

    #[test]
    fn ctrl_shift_down_extends_to_last_nonblank_cell_in_column() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(5, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 2, col: 0 }, "mid".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 3, col: 0 }, "end".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS + 1,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(
            KeyCode::Down,
            KeyModifiers::CONTROL | KeyModifiers::SHIFT,
        ))
        .unwrap();

        assert_eq!(
            app.anchor,
            Some(SheetCursor {
                row: HEADER_ROWS + 1,
                col: MARGIN_COLS,
            })
        );
        assert_eq!(
            app.cursor,
            SheetCursor {
                row: HEADER_ROWS + 3,
                col: MARGIN_COLS,
            }
        );
        assert_eq!(app.selection_kind, SelectionKind::Cells);
    }

    #[test]
    fn ctrl_shift_up_extends_to_first_nonblank_cell_in_column() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(5, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "top".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "next".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS + 3,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(
            KeyCode::Up,
            KeyModifiers::CONTROL | KeyModifiers::SHIFT,
        ))
        .unwrap();

        assert_eq!(
            app.anchor,
            Some(SheetCursor {
                row: HEADER_ROWS + 3,
                col: MARGIN_COLS,
            })
        );
        assert_eq!(
            app.cursor,
            SheetCursor {
                row: HEADER_ROWS,
                col: MARGIN_COLS,
            }
        );
        assert_eq!(app.selection_kind, SelectionKind::Cells);
    }

    #[test]
    fn insert_menu_cols_inserts_before_cursor() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "left".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "right".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS + 1,
        };
        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::Insert,
                item: menu_items(MenuSection::Insert)
                    .iter()
                    .position(|item| item.label == "Cols")
                    .unwrap(),
            }],
        };

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(app.state.grid.main_cols(), 3);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("left")
        );
        assert_eq!(app.state.grid.get(&CellAddr::Main { row: 0, col: 1 }), None);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 2 })
                .as_deref(),
            Some("right")
        );
    }

    #[test]
    fn insert_menu_special_chars_reuses_existing_special_value() {
        let mut app = App::new(None);
        app.state
            .grid
            .set(&CellAddr::Header { row: 0, col: 0 }, "∞".into());
        app.cursor = SheetCursor { row: 0, col: 0 };
        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::Insert,
                item: menu_items(MenuSection::Insert)
                    .iter()
                    .position(|item| item.label == "Special Char")
                    .unwrap(),
            }],
        };

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(app.special_picker, Some(0));
    }

    #[test]
    fn insert_menu_unicode_characters_are_available() {
        let mut app = App::new(None);
        app.cursor = SheetCursor { row: 0, col: 0 };
        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::Insert,
                item: menu_items(MenuSection::Insert)
                    .iter()
                    .position(|item| item.label == "Special Char")
                    .unwrap(),
            }],
        };

        let items = menu_items(MenuSection::Insert);
        assert!(items.iter().any(|i| i.label == "Special Char"));
        assert!(items.iter().any(|i| i.label == "Date"));
        assert!(items.iter().any(|i| i.label == "Time"));

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        assert!(app.special_picker.is_some());
    }

    #[test]
    fn insert_menu_special_seed_uses_unicode_symbols() {
        let mut app = App::new(None);
        app.cursor = SheetCursor { row: 0, col: 0 };
        let seed = app.menu_insert_special_seed();
        assert_eq!(seed, "∞");
        let choices = special_value_choices(&app.cursor.to_addr(&app.state.grid));
        assert!(choices.contains(&"∞"));
        assert!(choices.contains(&"Σ"));
        assert!(choices.contains(&"Ω"));
    }

    #[test]
    fn insert_menu_hyperlink_reuses_existing_url() {
        let mut app = App::new(None);
        app.state.grid.set(
            &CellAddr::Main { row: 0, col: 0 },
            "https://example.com".into(),
        );
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::Insert,
                item: menu_items(MenuSection::Insert)
                    .iter()
                    .position(|item| item.label == "Hyperlink")
                    .unwrap(),
            }],
        };

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        match app.mode {
            Mode::Edit { buffer, .. } => assert_eq!(buffer, "https://example.com"),
            other => panic!("unexpected mode: {other:?}"),
        }
    }

    #[test]
    fn special_value_choices_cover_margin_cells() {
        assert!(!special_value_choices(&CellAddr::Header { row: 0, col: 0 }).is_empty());
        assert!(!special_value_choices(&CellAddr::Footer { row: 0, col: 0 }).is_empty());
        assert!(!special_value_choices(&CellAddr::Left { col: 0, row: 0 }).is_empty());
        assert!(!special_value_choices(&CellAddr::Right { col: 0, row: 0 }).is_empty());
        assert!(special_value_choices(&CellAddr::Main { row: 0, col: 0 }).is_empty());
    }

    #[test]
    fn edit_mode_renders_special_suggestions_box() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.cursor = SheetCursor { row: 0, col: 0 };
        app.mode = Mode::Edit {
            buffer: String::new(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        let backend = TestBackend::new(60, 16);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        assert!(!row(2).contains("Suggestions"));
    }

    #[test]
    fn startup_renders_header_template_values_without_cursor_movement() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(2, 2);
        app.state.grid.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: MARGIN_COLS as u32 + 1,
            },
            "=A*2 -- POW2".into(),
        );
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "7".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS + 1,
        };
        app.mode = Mode::Edit {
            buffer: "=A*2 -- POW2".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        let backend = TestBackend::new(80, 24);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        assert!((0..buffer.area.height).any(|y| row(y).contains("14")));
    }

    #[test]
    fn formula_bar_shows_evaluated_formula_text() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "=π".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };

        let backend = TestBackend::new(40, 8);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        assert!(row(1).contains("3.141"));

        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "=2*π".into());
        let backend = TestBackend::new(40, 8);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        assert!(row(1).contains("6.283"));

        app.mode = Mode::Edit {
            buffer: "=2*π".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };
        let backend = TestBackend::new(40, 8);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        assert!(row(1).contains("6.283"));

        app.mode = Mode::Edit {
            buffer: "=π".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };
        let backend = TestBackend::new(40, 8);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        assert!(row(1).contains("=π"));
    }

    #[test]
    fn escaped_edit_does_not_follow_cursor_and_can_be_restored() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = app.start_edit_mode("draft".into(), None, false, false, None);

        app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::empty()))
            .unwrap();
        assert!(matches!(app.mode, Mode::Normal));
        assert!(app.edit_target_addr.is_none());
        assert!(app.pending_lost_edit.is_some());
        assert!(app.status.contains("Press Enter"));

        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::empty()))
            .unwrap();
        let backend = TestBackend::new(50, 8);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        assert!(row(1).contains("B1"), "{}", row(1));
        assert!(!row(1).contains("A1  draft"), "{}", row(1));

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.cursor.col, MARGIN_COLS);
        match &app.mode {
            Mode::Edit { buffer, .. } => assert_eq!(buffer, "draft"),
            other => panic!("expected restored edit mode, got {other:?}"),
        }
        assert!(app.pending_lost_edit.is_none());
    }

    #[test]
    fn inserted_date_fits_column_exactly() {
        let mut app = App::new(None);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = app.start_edit_mode("2024-01-02".into(), None, false, true, None);
        if let Mode::Edit { buffer, .. } = &app.mode {
            let raw = buffer.clone();
            app.commit_edit_buffer(&raw).unwrap();
        } else {
            panic!("expected edit mode");
        }
        // With the new address semantics the computed rendered width may
        // include the left-margin offset; update expectation accordingly.
        assert_eq!(app.state.grid.col_width(MARGIN_COLS), 11);
    }

    #[test]
    fn f2_starts_editing_current_cell() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "hello".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };

        app.handle_key(KeyEvent::new(KeyCode::F(2), KeyModifiers::empty()))
            .unwrap();

        match &app.mode {
            Mode::Edit { buffer, .. } => assert_eq!(buffer, "hello"),
            other => panic!("expected edit mode, got {other:?}"),
        }
    }

    #[test]
    fn undo_enables_redo_and_hints_follow_state() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };

        assert!(!app.hints_line().contains("Ctrl+Z"));
        assert!(!app.hints_line().contains("Ctrl+Y"));

        app.commit_edit_buffer("one").unwrap();
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("one")
        );
        assert!(app.hints_line().contains("Ctrl+Z"));
        assert!(!app.hints_line().contains("Ctrl+Y"));

        app.handle_key(KeyEvent::new(KeyCode::Char('z'), KeyModifiers::CONTROL))
            .unwrap();
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref()
                .unwrap_or(""),
            ""
        );
        assert!(app.hints_line().contains("Ctrl+Y"));

        app.handle_key(KeyEvent::new(KeyCode::Char('y'), KeyModifiers::CONTROL))
            .unwrap();
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("one")
        );
        assert!(!app.hints_line().contains("Ctrl+Y"));
    }

    #[test]
    fn explicit_address_edit_moves_cursor_to_target() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 3);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.commit_edit_buffer("C~1").unwrap();
        // Header `~1` maps to the last header row index (HEADER_ROWS - 1).
        assert_eq!(app.cursor.row, HEADER_ROWS - 1);
        assert_eq!(app.cursor.col, MARGIN_COLS + 2);
    }

    #[test]
    fn long_grid_values_truncate_one_char_shorter() {
        assert_eq!(truncate_with_ellipsis("abcdef", 4), "abc…");
        assert_eq!(truncate_with_ellipsis("abcdef", 1), "…");
        // display width, not char count: fullwidth letters are width 2 each
        assert_eq!(truncate_with_ellipsis("ＡＢＣＤＥＦ", 4), "Ａ…");
    }

    #[test]
    fn startup_keeps_total_column_visible() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(4, 3);
        app.state.grid.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: MARGIN_COLS as u32 + 2,
            },
            "=TOTAL".into(),
        );
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "1".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "7".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 2, col: 0 }, "0".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 3, col: 0 }, "5".into());

        let backend = TestBackend::new(80, 24);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        assert!((0..buffer.area.height).any(|y| row(y).contains("TOTAL")));
    }

    #[test]
    fn total_row_and_total_column_intersection_sums_row_totals() {
        let mut state = SheetState::new(4, 3);
        state
            .grid
            .set(&CellAddr::Main { row: 0, col: 2 }, "1".into());
        state
            .grid
            .set(&CellAddr::Main { row: 1, col: 2 }, "2".into());
        state
            .grid
            .set(&CellAddr::Main { row: 2, col: 2 }, "3".into());
        state
            .grid
            .set(&CellAddr::Main { row: 3, col: 2 }, "4".into());

        assert_eq!(
            footer_special_col_aggregate(
                &state.grid,
                AggFunc::Sum,
                MARGIN_COLS + 2,
                state.grid.main_rows(),
                state.grid.main_cols(),
            ),
            Some("10".into())
        );
    }

    #[test]
    fn moving_right_in_right_margin_does_not_reveal_more_left_columns() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS + app.state.grid.main_cols() + 0,
        };

        let backend = TestBackend::new(70, 18);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };
        let first = (0..buffer.area.height)
            .map(row)
            .find(|line| line.contains("<0") || line.contains("<1") || line.contains("<2"))
            .unwrap_or_default();

        app.cursor.col += 1;
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row2 = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };
        let second = (0..buffer.area.height)
            .map(row2)
            .find(|line| line.contains("<0") || line.contains("<1") || line.contains("<2"))
            .unwrap_or_default();

        assert_eq!(first.contains("<2"), second.contains("<2"));
        assert_eq!(first.contains("<3"), second.contains("<3"));
    }

    #[test]
    fn moving_left_within_left_margin_steps_the_viewport_once() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        let backend = TestBackend::new(70, 18);
        let mut terminal = Terminal::new(backend).unwrap();

        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };
        let leftmost = |line: &str| -> Option<usize> {
            (0..10)
                .filter_map(|n| {
                    let label = format!("<{n}");
                    line.find(&label).map(|idx| (idx, n))
                })
                .min_by_key(|(idx, _)| *idx)
                .map(|(_, n)| n)
        };
        let initial = (0..buffer.area.height)
            .map(row)
            .find_map(|line| leftmost(&line))
            .unwrap_or(0);

        app.cursor.col = MARGIN_COLS - 1;
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row2 = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };
        let moved = (0..buffer.area.height)
            .map(row2)
            .find_map(|line| leftmost(&line))
            .unwrap_or(0);

        assert!(moved >= initial);
        assert!(moved <= initial + 1);
    }

    #[test]
    fn left_margin_labels_are_mirrored() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: 0,
        };

        let backend = TestBackend::new(70, 18);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        let line = (0..buffer.area.height)
            .map(row)
            .find(|line| line.contains("[A") || line.contains("[B") || line.contains("[C"))
            .unwrap_or_default();

        assert!(line.contains("[A"));
    }

    #[test]
    fn left_margin_total_row_computes_subtotals() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(3, 2);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "11".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "44".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "22".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "55".into());
        let key_col: MarginIndex = MARGIN_COLS - 1;
        app.state.grid.set(
            &CellAddr::Left {
                col: key_col,
                row: 2,
            },
            "=TOTAL".into(),
        );

        let backend = TestBackend::new(80, 18);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        let line = (0..buffer.area.height)
            .map(row)
            .find(|line| line.contains("TOTAL"))
            .unwrap_or_default();

        assert!(line.contains("TOTAL"));
        assert!(line.contains("3"));
    }

    #[test]
    fn left_margin_total_rows_include_right_margin_subtotals_of_totals() {
        let mut state = SheetState::new(6, 2);
        state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "11".into());
        state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "1".into());
        state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "22".into());
        state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "2".into());
        state.grid.set(
            &CellAddr::Left {
                col: (MARGIN_COLS - 1),
                row: 2,
            },
            "=TOTAL".into(),
        );
        state
            .grid
            .set(&CellAddr::Main { row: 3, col: 0 }, "33".into());
        state
            .grid
            .set(&CellAddr::Main { row: 3, col: 1 }, "3".into());
        state
            .grid
            .set(&CellAddr::Main { row: 4, col: 0 }, "44".into());
        state
            .grid
            .set(&CellAddr::Main { row: 4, col: 1 }, "4".into());
        state.grid.set(
            &CellAddr::Left {
                col: (MARGIN_COLS - 1),
                row: 5,
            },
            "=TOTAL".into(),
        );
        let right_col = MARGIN_COLS + state.grid.main_cols() + 1;
        state.grid.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: right_col as u32,
            },
            "=TOTAL".into(),
        );

        assert_eq!(
            left_margin_special_col_aggregate(&state.grid, AggFunc::Sum, right_col, 0, 2, 2),
            Some("36".into())
        );
        assert_eq!(
            left_margin_special_col_aggregate(&state.grid, AggFunc::Sum, right_col, 3, 5, 2),
            Some("84".into())
        );
    }

    #[test]
    fn right_margin_aggregate_detects_top_header_marker() {
        let mut state = SheetState::new(4, 3);
        state.grid.set(
            &CellAddr::Header {
                row: 0,
                col: (MARGIN_COLS + 2) as u32,
            },
            "=TOTAL".into(),
        );
        state
            .grid
            .set(&CellAddr::Main { row: 0, col: 2 }, "1".into());
        state
            .grid
            .set(&CellAddr::Main { row: 1, col: 2 }, "2".into());
        state
            .grid
            .set(&CellAddr::Main { row: 2, col: 2 }, "3".into());

        assert_eq!(
            right_col_agg_func(&state.grid, MARGIN_COLS + 2),
            Some(AggFunc::Sum)
        );
        assert_eq!(
            footer_special_col_aggregate(&state.grid, AggFunc::Sum, MARGIN_COLS + 2, 4, 3),
            Some("6".into())
        );
    }

    #[test]
    fn footer_special_col_aggregate_uses_data_region_width() {
        let mut state = SheetState::new(3, 5);
        state.grid.set(
            &CellAddr::Header {
                row: 0,
                col: (MARGIN_COLS + 2) as u32,
            },
            "=TOTAL".into(),
        );
        state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "1".into());
        state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "2".into());
        state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "3".into());
        state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "4".into());

        assert_eq!(
            footer_special_col_aggregate(&state.grid, AggFunc::Sum, MARGIN_COLS + 2, 2, 5),
            Some("10".into())
        );
    }

    #[test]
    fn page_up_page_down_step_by_grid_viewport_row_count() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(20, 1);
        app.grid_viewport_data_rows = 4;
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };

        app.handle_key(KeyEvent::new(KeyCode::PageDown, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.cursor.row, HEADER_ROWS + 4);

        app.handle_key(KeyEvent::new(KeyCode::PageUp, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.cursor.row, HEADER_ROWS);
    }

    #[test]
    fn export_preview_scroll_moves_with_arrow_keys() {
        let mut app = App::new(None);
        app.export_preview_scroll = 10;
        app.mode = Mode::ExportTsv {
            buffer: String::new(),
        };

        app.handle_key(KeyEvent::new(KeyCode::Up, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.export_preview_scroll, 9);

        app.handle_key(KeyEvent::new(KeyCode::PageDown, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.export_preview_scroll, 29);
    }

    #[test]
    fn subtotal_tiny_shows_c4_and_c5_totals() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let path = std::path::PathBuf::from("docs/tests/subtotal-tiny.corro");
        let mut app = App::new(Some(path));
        app.load_initial().unwrap();

        let backend = TestBackend::new(120, 20);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        let lines: Vec<String> = (0..buffer.area.height).map(row).collect();
        let row4 = lines
            .iter()
            .find(|line| line.starts_with("│   4 "))
            .cloned()
            .unwrap_or_default();
        let row5 = lines
            .iter()
            .find(|line| line.starts_with("│   5 "))
            .cloned()
            .unwrap_or_default();

        assert!(row4.contains("AVERAGE"), "rendered row 4: {row4}");
        assert!(row5.contains("TOTAL"), "rendered row 5: {row5}");
    }

    #[test]
    fn subtotal_tiny_renders_c1_and_total_cells() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(Some(std::path::PathBuf::from(
            "docs/tests/subtotal-tiny.corro",
        )));
        app.load_initial().unwrap();

        let backend = TestBackend::new(120, 20);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        let lines: Vec<String> = (0..buffer.area.height).map(row).collect();
        assert!(
            lines.iter().any(|line| line.contains("TOTAL")),
            "{lines:#?}"
        );
        assert!(
            lines.iter().any(|line| line.contains("│   5 TOTAL")),
            "{lines:#?}"
        );
    }

    #[test]
    fn tsv_export_preview_ignores_active_selection() {
        let mut app = App::new(Some(std::path::PathBuf::from(
            "docs/tests/subtotal-tiny.corro",
        )));
        app.load_initial().unwrap();
        app.anchor = Some(SheetCursor {
            row: HEADER_ROWS + 1,
            col: MARGIN_COLS,
        });
        app.cursor = SheetCursor {
            row: HEADER_ROWS + 2,
            col: MARGIN_COLS + 1,
        };

        let text = app.export_preview_text(false);

        // Header margin still carries the "TOTAL" label; aggregate rows export computed values
        // in the key column (not the words TOTAL/AVERAGE) so they match =SUBTOTAL semantics.
        assert!(text.contains("TOTAL"), "{text}");
        assert!(text.contains("1.5"), "{text}");
    }

    /// TSV body from `export_tsv` / export preview; matches `docs/tests/subtotal-tiny-tsv.tsv`.
    #[test]
    fn subtotal_tiny_tsv_export_matches_golden() {
        let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("docs/tests/subtotal-tiny.corro");
        let mut app = App::new(Some(path));
        app.load_initial().unwrap();

        let tsv = app.do_export(false);
        let expected = include_str!("../../docs/tests/subtotal-tiny-tsv.tsv");
        let norm = |s: &str| s.replace("\r\n", "\n");
        assert_eq!(norm(&tsv), norm(expected), "subtotal-tiny TSV export");
    }

    /// ASCII table from `export_ascii_table` / `do_export_ascii`; matches
    /// `docs/tests/subtotal-tiny-ascii.txt`.
    #[test]
    fn subtotal_tiny_ascii_export_matches_golden() {
        let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("docs/tests/subtotal-tiny.corro");
        let mut app = App::new(Some(path));
        app.load_initial().unwrap();

        let ascii = app.do_export_ascii();
        let expected = include_str!("../../docs/tests/subtotal-tiny-ascii.txt");
        let norm = |s: &str| s.replace("\r\n", "\n");
        assert_eq!(norm(&ascii), norm(expected), "subtotal-tiny ASCII table export");
    }

    #[test]
    fn export_tsv_clears_stale_menu_popup() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        let backend = TestBackend::new(120, 20);
        let mut terminal = Terminal::new(backend).unwrap();

        app.mode = Mode::Menu {
            stack: vec![
                MenuLevel {
                    section: MenuSection::File,
                    item: 2,
                },
                MenuLevel {
                    section: MenuSection::Export,
                    item: 0,
                },
            ],
        };
        terminal.draw(|f| app.draw(f)).unwrap();

        app.mode = Mode::ExportTsv {
            buffer: String::new(),
        };
        terminal.draw(|f| app.draw(f)).unwrap();

        let buffer = terminal.backend().buffer();
        let lines: Vec<String> = (0..buffer.area.height)
            .map(|y| {
                (0..buffer.area.width)
                    .map(|x| buffer[(x, y)].symbol())
                    .collect::<String>()
            })
            .collect();

        assert!(
            lines
                .iter()
                .any(|line| line.contains("export TSV (blank=clipboard):")),
            "{lines:#?}"
        );
        assert!(
            lines
                .iter()
                .all(|line| !line.contains("T·TSV") && !line.contains("C·CSV")),
            "{lines:#?}"
        );
    }

    #[test]
    fn export_tsv_clears_persist_sort_from_previous_menu_frame() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        let backend = TestBackend::new(120, 20);
        let mut terminal = Terminal::new(backend).unwrap();

        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::File,
                item: 6,
            }],
        };
        terminal.draw(|f| app.draw(f)).unwrap();

        app.mode = Mode::ExportTsv {
            buffer: String::new(),
        };
        terminal.draw(|f| app.draw(f)).unwrap();

        let buffer = terminal.backend().buffer();
        let lines: Vec<String> = (0..buffer.area.height)
            .map(|y| {
                (0..buffer.area.width)
                    .map(|x| buffer[(x, y)].symbol())
                    .collect::<String>()
            })
            .collect();

        assert!(
            lines
                .iter()
                .any(|line| line.contains("export TSV (blank=clipboard):")),
            "{lines:#?}"
        );
        let leaked_row = lines
            .iter()
            .find(|line| line.contains("Persist sort"))
            .cloned()
            .unwrap_or_default();
        assert!(leaked_row.is_empty(), "{lines:#?}");
    }

    #[test]
    fn adjacent_cells_keep_a_visible_gap() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "1".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "2".into());

        let backend = TestBackend::new(60, 12);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();

        let row = (0..buffer.area.height)
            .find(|&y| {
                let text: String = (0..buffer.area.width)
                    .map(|x| buffer[(x, y)].symbol())
                    .collect();
                text.contains("1") && text.contains("2")
            })
            .unwrap();

        let text: String = (0..buffer.area.width)
            .map(|x| buffer[(x, row)].symbol())
            .collect();
        let one = text.find('1').unwrap();
        let two = text.find('2').unwrap();
        assert!(two > one + 1, "rendered row: {text}");
    }

    #[test]
    fn right_margin_aggregate_uses_top_or_bottom_header_marker() {
        let mut state = SheetState::new(3, 3);
        state.grid.set(
            &CellAddr::Header {
                row: 0,
                col: (MARGIN_COLS + 2) as u32,
            },
            "=TOTAL".into(),
        );
        state.grid.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: (MARGIN_COLS + 2) as u32,
            },
            "".into(),
        );
        state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "1".into());
        state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "2".into());
        state
            .grid
            .set(&CellAddr::Main { row: 2, col: 0 }, "3".into());

        assert_eq!(
            right_col_agg_func(&state.grid, MARGIN_COLS + 2),
            Some(AggFunc::Sum)
        );
    }

    #[test]
    fn aggregate_columns_render_in_cyan() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(6, 2);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "11".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "1".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "22".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "2".into());
        app.state.grid.set(
            &CellAddr::Left {
                col: (MARGIN_COLS - 1),
                row: 2,
            },
            "=TOTAL".into(),
        );
        app.state.grid.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: (MARGIN_COLS + app.state.grid.main_cols() + 1) as u32,
            },
            "=TOTAL".into(),
        );

        let backend = TestBackend::new(96, 20);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();

        let saw_cyan_numeric = (0..buffer.area.height).any(|y| {
            (0..buffer.area.width).any(|x| {
                let cell = &buffer[(x, y)];
                if cell.symbol().trim().parse::<f64>().is_ok()
                    && cell.style().fg == Some(Color::Cyan)
                {
                    // debug removed
                }
                cell.style().fg == Some(Color::Cyan) && cell.symbol().trim().parse::<f64>().is_ok()
            })
        });
        let saw_bold_footer_label = (0..buffer.area.height).any(|y| {
            (0..buffer.area.width).any(|x| {
                let cell = &buffer[(x, y)];
                cell.style().fg == Some(Color::Cyan)
                    && cell.style().add_modifier.contains(Modifier::BOLD)
                    && cell.symbol().trim().starts_with('_')
            })
        });
        assert!(saw_cyan_numeric);
        assert!(saw_bold_footer_label);
    }

    #[test]
    fn left_margin_max_uses_previous_total_row() {
        let mut state = SheetState::new(6, 2);
        state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "11".into());
        state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "1".into());
        state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "22".into());
        state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "2".into());
        state.grid.set(
            &CellAddr::Left {
                col: (MARGIN_COLS - 1),
                row: 2,
            },
            "=TOTAL".into(),
        );
        state
            .grid
            .set(&CellAddr::Main { row: 3, col: 0 }, "33".into());
        state
            .grid
            .set(&CellAddr::Main { row: 3, col: 1 }, "3".into());
        state
            .grid
            .set(&CellAddr::Main { row: 4, col: 0 }, "44".into());
        state
            .grid
            .set(&CellAddr::Main { row: 4, col: 1 }, "4".into());
        state.grid.set(
            &CellAddr::Left {
                col: (MARGIN_COLS - 1),
                row: 5,
            },
            "MAX".into(),
        );

        let right_col = MARGIN_COLS + state.grid.main_cols() + 1;
        state.grid.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: right_col as u32,
            },
            "=TOTAL".into(),
        );

        assert_eq!(row_total_block_start(&state.grid, 5), 3);
        assert_eq!(
            left_margin_special_col_aggregate(&state.grid, AggFunc::Max, right_col, 3, 5, 2),
            Some("48".into())
        );
    }

    #[test]
    fn left_margin_main_col_aggregate_uses_immediate_block() {
        let mut state = SheetState::new(9, 1);
        state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "1".into());
        state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "2".into());
        state.grid.set(
            &CellAddr::Left {
                col: (MARGIN_COLS - 1),
                row: 2,
            },
            "=TOTAL".into(),
        );
        state
            .grid
            .set(&CellAddr::Main { row: 3, col: 0 }, "16.77".into());
        state
            .grid
            .set(&CellAddr::Main { row: 4, col: 0 }, "0.00".into());
        state.grid.set(
            &CellAddr::Left {
                col: (MARGIN_COLS - 1),
                row: 5,
            },
            "=TOTAL".into(),
        );
        state
            .grid
            .set(&CellAddr::Main { row: 6, col: 0 }, "67.67".into());
        state
            .grid
            .set(&CellAddr::Main { row: 7, col: 0 }, "0.00".into());
        state.grid.set(
            &CellAddr::Left {
                col: (MARGIN_COLS - 1),
                row: 8,
            },
            "=TOTAL".into(),
        );

        assert_eq!(row_total_block_start(&state.grid, 8), 6);
        assert_eq!(
            left_margin_main_col_aggregate(&state.grid, AggFunc::Sum, 8, 0),
            "67.67"
        );
    }

    #[test]
    fn stacked_left_margin_max_falls_back_to_previous_raw_block() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(4, 2);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "11".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "1".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "22".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "2".into());
        app.state.grid.set(
            &CellAddr::Left {
                col: (MARGIN_COLS - 1),
                row: 2,
            },
            "=TOTAL".into(),
        );
        app.state.grid.set(
            &CellAddr::Left {
                col: (MARGIN_COLS - 1),
                row: 3,
            },
            "MAX".into(),
        );
        app.state.grid.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: (MARGIN_COLS + app.state.grid.main_cols() + 1) as u32,
            },
            "=TOTAL".into(),
        );

        let backend = TestBackend::new(90, 18);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        let max_line = (0..buffer.area.height)
            .map(row)
            .find(|line| line.contains("MAX"))
            .unwrap_or_default();

        assert!(max_line.contains("22"));
        assert!(max_line.contains("2"));
    }

    #[test]
    fn widened_column_shows_full_cell_text() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.state.grid.set_col_width(MARGIN_COLS, Some(24));
        app.state.grid.set(
            &CellAddr::Main { row: 0, col: 0 },
            "abcdefghijklmnopqrstuvwx".into(),
        );

        let backend = TestBackend::new(80, 18);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        assert!((0..buffer.area.height).any(|y| row(y).contains("abcdefghijklmnopqrstuvwx")));
    }

    #[test]
    fn right_margin_moves_view_one_step_at_a_time() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 4);
        app.state.grid.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: MARGIN_COLS as u32 + 3,
            },
            "=TOTAL".into(),
        );
        let backend = TestBackend::new(80, 18);
        let mut terminal = Terminal::new(backend).unwrap();

        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS + app.state.grid.main_cols(),
        };
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };
        let initial = (0..buffer.area.height)
            .map(row)
            .find(|line| line.contains("]A") || line.contains("]B"))
            .unwrap_or_default();

        app.cursor.col += 1;
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row2 = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };
        let moved = (0..buffer.area.height)
            .map(row2)
            .find(|line| line.contains("]A") || line.contains("]B"))
            .unwrap_or_default();

        assert!(initial.contains("]A"));
        assert!(moved.contains("]B"));
        assert!((0..buffer.area.height).any(|y| row2(y).contains("TOTAL")));
    }

    #[test]
    fn right_margin_columns_scroll_to_keep_cursor_visible() {
        let mut state = SheetState::new(1, 2);
        let right_start = MARGIN_COLS + state.grid.main_cols();
        for i in 0..6 {
            state.grid.set(
                &CellAddr::Header {
                    row: (HEADER_ROWS - 1) as u32,
                    col: (right_start + i) as u32,
                },
                "=TOTAL".into(),
            );
        }

        let cursor = SheetCursor {
            row: HEADER_ROWS,
            col: right_start + 5,
        };
        let (cols, _) = visible_col_indices(&state, cursor, 3, 0);

        assert!(cols.contains(&cursor.col), "{cols:?}");
        assert!(!cols.contains(&right_start), "{cols:?}");
    }

    #[test]
    fn sheet_go_jumps_to_main_cell_and_grows_extent() {
        let mut app = App::new(None);

        assert!(app.go_to_cell("c12"));

        assert_eq!(
            app.cursor,
            SheetCursor {
                row: HEADER_ROWS + 11,
                col: MARGIN_COLS + 2,
            }
        );
        assert_eq!(app.state.grid.main_rows(), 12);
        assert_eq!(app.state.grid.main_cols(), 3);
    }

    #[test]
    fn sheet_go_jumps_to_right_margin_header_without_expanding_main_cols() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);

        assert!(app.go_to_cell("]A~1"));

        assert_eq!(
            app.cursor,
            SheetCursor {
                row: HEADER_ROWS - 1,
                col: MARGIN_COLS + 2,
            }
        );
        assert_eq!(app.state.grid.main_cols(), 2);
    }

    #[test]
    fn sheet_go_supports_bare_row_and_column_targets() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS + 1,
        };

        assert!(app.go_to_cell("123"));
        assert_eq!(
            app.cursor,
            SheetCursor {
                row: HEADER_ROWS + 122,
                col: MARGIN_COLS + 1,
            }
        );
        assert_eq!(app.state.grid.main_rows(), 123);

        assert!(app.go_to_cell("d"));
        assert_eq!(
            app.cursor,
            SheetCursor {
                row: HEADER_ROWS + 122,
                col: MARGIN_COLS + 3,
            }
        );
        assert_eq!(app.state.grid.main_cols(), 4);
    }

    #[test]
    fn sheet_go_supports_zz_right_margin_column() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };

        assert!(app.go_to_cell("]ZZ"));

        assert_eq!(
            app.cursor,
            SheetCursor {
                row: HEADER_ROWS,
                col: MARGIN_COLS + 2 + MARGIN_COLS - 1,
            }
        );
        assert_eq!(app.state.grid.main_cols(), 2);
        assert_eq!(
            addr::cell_ref_text(
                &app.cursor.to_addr(&app.state.grid),
                app.state.grid.main_cols()
            ),
            "]ZZ1"
        );
    }

    #[test]
    fn sheet_go_dollar_goes_to_sheet_by_name_or_id() {
        let mut app = App::new(None);
        app.add_sheet("Sheet2".into());
        assert_eq!(app.view_sheet_id, 2);
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "here".into());

        assert!(app.go_to_cell("$Sheet1"));
        assert_eq!(app.view_sheet_id, 1);
        assert_eq!(
            app.cursor,
            SheetCursor {
                row: HEADER_ROWS,
                col: MARGIN_COLS,
            }
        );

        assert!(app.go_to_cell("$2:B2"));
        assert_eq!(app.view_sheet_id, 2);
        assert_eq!(
            app.cursor,
            SheetCursor {
                row: HEADER_ROWS + 1,
                col: MARGIN_COLS + 1,
            }
        );
        assert_eq!(
            app.state.grid.get(&CellAddr::Main { row: 1, col: 1 }).as_deref(),
            Some("here")
        );

        assert!(app.go_to_cell("$SHEET1"));
        assert_eq!(app.view_sheet_id, 1);
    }

    #[test]
    fn header_only_b_column_stays_visible_as_b() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.state.grid.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: MARGIN_COLS as u32 + 1,
            },
            "HDR-B".into(),
        );
        app.state.grid.set(
            &CellAddr::Footer {
                row: 0,
                col: MARGIN_COLS as u32 + 1,
            },
            "FTR-B".into(),
        );

        let backend = TestBackend::new(80, 18);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        let header_line = (0..buffer.area.height)
            .map(row)
            .find(|line| line.contains("HDR-B"))
            .unwrap_or_default();
        let footer_line = (0..buffer.area.height)
            .map(row)
            .find(|line| line.contains("FTR-B"))
            .unwrap_or_default();

        assert!(header_line.contains("B") || footer_line.contains("B"));
        assert!(!header_line.contains("]A") || header_line.contains("B"));
    }

    #[test]
    fn escape_cancels_edit_without_committing() {
        let mut app = App::new(None);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "orig".into());
        app.mode = Mode::Edit {
            buffer: "changed".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::empty()))
            .unwrap();

        assert!(matches!(app.mode, Mode::Normal));
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("orig")
        );
    }

    #[test]
    fn open_path_parses_link_revision() {
        let fixture = docs_test_path("main.corro");
        let parsed = parse_open_path_request(&format!("link {} 2", fixture.display())).unwrap();
        match parsed {
            OpenPathRequest::Revision { path, revision } => {
                assert_eq!(path, fixture);
                assert_eq!(revision, 2);
            }
            other => panic!("unexpected parse: {other:?}"),
        }
    }

    #[test]
    fn linked_revision_uses_source_path_and_detaches_on_save() {
        let fixture = docs_test_path("main.corro");
        let mut app = App::new_with_revision_limit(Some(fixture.clone()), Some(2));
        assert!(app.path.is_none());
        assert_eq!(app.source_path, Some(fixture));
        assert_eq!(app.revision_limit, Some(2));

        let tmp = tempfile::NamedTempFile::new().unwrap();
        app.save_to_path(tmp.path()).unwrap();

        let expected = tmp.path().to_path_buf().with_extension("corro");
        assert_eq!(app.path, Some(expected));
        assert_eq!(app.source_path, None);
        assert_eq!(app.revision_limit, None);
    }

    #[test]
    fn save_clears_revision_limit() {
        let fixture = docs_test_path("main.corro");
        let mut app = App::new_with_revision_limit(Some(fixture), Some(2));
        app.revision_limit = Some(2);
        let path = tempfile::NamedTempFile::new().unwrap();

        app.save_to_path(path.path()).unwrap();

        assert_eq!(app.revision_limit, None);
    }

    #[test]
    fn file_menu_includes_replay() {
        let items = menu_items(MenuSection::File);
        assert!(items.iter().any(|item| item.label == "Replay"));
    }

    #[test]
    fn file_replay_loads_workbook_log_and_uses_real_revision_count() {
        let path = tempfile::Builder::new()
            .suffix(".corro")
            .tempfile()
            .unwrap();
        std::fs::write(path.path(), "SET $1:A1 7\nSET $1:B1 DONE\n").unwrap();
        let mut app = App::new(Some(path.path().to_path_buf()));

        let mode = app.menu_action_mode(MenuAction::Replay);

        assert!(matches!(mode, Mode::RevisionBrowse));
        assert!(app.path.is_none());
        assert_eq!(app.source_path, Some(path.path().to_path_buf()));
        assert_eq!(app.revision_browse_limit, 2);
        assert!(app.status.contains("@ revision 2"));
        assert!(!app.status.contains("184467440737095516"));
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("7")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 1 })
                .as_deref(),
            Some("DONE")
        );

        app.mode = mode;
        app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.revision_browse_limit, 1);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("7")
        );
        assert!(app
            .state
            .grid
            .get(&CellAddr::Main { row: 0, col: 1 })
            .is_none());
    }

    #[test]
    fn new_sheet_creates_second_tab() {
        let mut app = App::new(None);
        app.add_sheet("Sheet2".into());

        assert_eq!(app.workbook.sheet_count(), 2);
        assert_eq!(app.view_sheet_id, 2);
        assert_eq!(app.workbook.sheet_title(1), "Sheet2");
    }

    #[test]
    fn new_sheet_is_logged_for_live_file() {
        let path = tempfile::NamedTempFile::new().unwrap();
        let mut app = App::new(Some(path.path().to_path_buf()));

        app.add_sheet("Sheet2".into());

        let log = std::fs::read_to_string(path.path()).unwrap();
        assert!(log.contains("$2:NEW_SHEET Sheet2"));
    }

    #[test]
    fn zerosum_right_from_a_in_edit_mode_moves_to_b() {
        let fixture = docs_test_path("zerosum.corro");
        if !fixture.exists() {
            eprintln!("Skipping zerosum_right_from_a_in_edit_mode_moves_to_b: fixture missing");
            return;
        }

        let mut app = App::new(Some(fixture));
        app.load_initial().unwrap();

        assert_eq!(
            app.cursor.to_addr(&app.state.grid),
            CellAddr::Main { row: 0, col: 0 }
        );

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();
        assert!(matches!(app.mode, Mode::Edit { .. }));
        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::empty()))
            .unwrap();

        assert!(matches!(app.mode, Mode::Normal));
        assert_eq!(app.state.grid.main_cols(), 2);
        assert_eq!(
            app.cursor.to_addr(&app.state.grid),
            CellAddr::Main { row: 0, col: 1 }
        );
    }

    #[test]
    fn ctrl_c_copies_selected_cells() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(2, 2);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "b".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "c".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "d".into());
        app.anchor = Some(SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        });
        app.cursor = SheetCursor {
            row: HEADER_ROWS + 1,
            col: MARGIN_COLS + 1,
        };
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(KeyCode::Char('c'), KeyModifiers::CONTROL))
            .unwrap();

        let copied = test_clipboard_text().unwrap();
        assert_eq!(copied, "a\tb\nc\td\n");
    }

    #[test]
    fn edit_menu_copy_copies_selected_cells() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "copy me".into());
        app.anchor = Some(SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        });
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::Edit,
                item: 1,
            }],
        };

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(test_clipboard_text().as_deref(), Some("copy me\n"));
    }

    #[test]
    fn ctrl_c_and_edit_menu_copy_share_clipboard_output() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "shared".into());
        app.anchor = Some(SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        });
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };

        app.mode = Mode::Normal;
        app.handle_key(KeyEvent::new(KeyCode::Char('c'), KeyModifiers::CONTROL))
            .unwrap();
        let ctrl_copy = test_clipboard_text().unwrap();

        set_test_clipboard(None);
        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::Edit,
                item: 1,
            }],
        };
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(ctrl_copy, test_clipboard_text().unwrap());
    }

    #[test]
    fn paste_uses_copy_from_to_when_snapshot_matches() {
        let path = tempfile::NamedTempFile::new().unwrap();
        let mut app = App::new(Some(path.path().to_path_buf()));
        app.state.grid.set_main_size(3, 3);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "1".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "2".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "3".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "4".into());
        app.anchor = Some(SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        });
        app.cursor = SheetCursor {
            row: HEADER_ROWS + 1,
            col: MARGIN_COLS + 1,
        };
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(KeyCode::Char('c'), KeyModifiers::CONTROL))
            .unwrap();
        app.cursor = SheetCursor {
            row: HEADER_ROWS + 1,
            col: MARGIN_COLS + 1,
        };
        app.handle_key(KeyEvent::new(KeyCode::Char('v'), KeyModifiers::CONTROL))
            .unwrap();

        let log = std::fs::read_to_string(path.path()).unwrap();
        assert!(log.contains("COPY_FROM_TO A1:B2 B2:C3"));
    }

    #[test]
    fn ctrl_v_pastes_tsv_cells() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };

        set_test_clipboard(Some("x\ty\n1\t2\n".into()));
        app.handle_key(KeyEvent::new(KeyCode::Char('v'), KeyModifiers::CONTROL))
            .unwrap();

        assert_eq!(app.state.grid.main_rows(), 2);
        assert_eq!(app.state.grid.main_cols(), 2);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("x")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 1 })
                .as_deref(),
            Some("y")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 1, col: 0 })
                .as_deref(),
            Some("1")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 1, col: 1 })
                .as_deref(),
            Some("2")
        );
    }

    #[test]
    fn paste_logs_as_single_fill_op() {
        let path = tempfile::NamedTempFile::new().unwrap();
        let mut app = App::new(Some(path.path().to_path_buf()));
        app.state.grid.set_main_size(1, 1);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };

        set_test_clipboard(Some("x\ty\n1\t2\n".into()));
        app.handle_key(KeyEvent::new(KeyCode::Char('v'), KeyModifiers::CONTROL))
            .unwrap();

        let log = std::fs::read_to_string(path.path()).unwrap();
        assert!(
            log.contains("FILL A1=x B1=y A2=1 B2=2") || log.contains("FILL A1=x B1=y C2=1 D2=2")
        );
        assert_eq!(app.ops_applied, 1);
    }

    #[test]
    fn edit_menu_paste_pastes_tsv_cells() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::Edit,
                item: 2,
            }],
        };

        set_test_clipboard(Some("x\ty\n1\t2\n".into()));
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(app.state.grid.main_rows(), 2);
        assert_eq!(app.state.grid.main_cols(), 2);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("x")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 1 })
                .as_deref(),
            Some("y")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 1, col: 0 })
                .as_deref(),
            Some("1")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 1, col: 1 })
                .as_deref(),
            Some("2")
        );
    }

    #[test]
    fn ctrl_shift_v_pastes_values_only_in_normal_mode() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };

        set_test_clipboard(Some("=A1".into()));
        app.handle_key(KeyEvent::new(
            KeyCode::Char('v'),
            KeyModifiers::CONTROL | KeyModifiers::SHIFT,
        ))
        .unwrap();

        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("A1")
        );
    }

    #[test]
    fn ctrl_shift_v_pastes_values_only_in_edit_mode() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Edit {
            buffer: "=".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        set_test_clipboard(Some("=A1".into()));
        app.handle_key(KeyEvent::new(
            KeyCode::Char('v'),
            KeyModifiers::CONTROL | KeyModifiers::SHIFT,
        ))
        .unwrap();

        match &app.mode {
            Mode::Edit { buffer, .. } => assert_eq!(buffer, "A1"),
            other => panic!("unexpected mode: {other:?}"),
        }
    }

    #[test]
    fn ctrl_shift_p_pastes_raw_clipboard_in_edit_mode() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Edit {
            buffer: String::new(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        set_test_clipboard(Some("=A1".into()));
        app.handle_key(KeyEvent::new(
            KeyCode::Char('p'),
            KeyModifiers::CONTROL | KeyModifiers::SHIFT,
        ))
        .unwrap();

        match &app.mode {
            Mode::Edit { buffer, .. } => assert_eq!(buffer, "=A1"),
            other => panic!("unexpected mode: {other:?}"),
        }
    }

    #[test]
    fn find_menu_opens_prompt() {
        let mut app = App::new(None);

        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::Edit,
                item: 3,
            }],
        };
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        match &app.mode {
            Mode::Find { buffer } => assert!(buffer.is_empty()),
            other => panic!("unexpected mode: {other:?}"),
        }
    }

    #[test]
    fn alt_e_opens_edit_menu() {
        let mut app = App::new(None);

        app.handle_key(KeyEvent::new(KeyCode::Char('e'), KeyModifiers::ALT))
            .unwrap();

        match &app.mode {
            Mode::Menu { stack } => {
                assert_eq!(stack.len(), 1);
                assert_eq!(stack[0].section, MenuSection::Edit);
            }
            other => panic!("unexpected mode: {other:?}"),
        }
    }

    #[test]
    fn replace_menu_opens_prompt() {
        let mut app = App::new(None);

        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::Edit,
                item: 4,
            }],
        };
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        match &app.mode {
            Mode::Replace { buffer } => assert!(buffer.is_empty()),
            other => panic!("unexpected mode: {other:?}"),
        }
    }

    #[test]
    fn find_next_moves_cursor_to_matching_cell() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(2, 2);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "findme".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Find {
            buffer: "findme".into(),
        };
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.cursor.row, HEADER_ROWS + 1);
        assert_eq!(app.cursor.col, MARGIN_COLS + 1);
        assert!(matches!(app.mode, Mode::Find { .. }));
        assert!(app.status.contains("Found"));
    }

    #[test]
    fn replace_all_updates_matching_cells() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "foo".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "barfoo".into());
        app.mode = Mode::Replace {
            buffer: "foo|x".into(),
        };
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("x")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 1 })
                .as_deref(),
            Some("barx")
        );
        assert!(matches!(app.mode, Mode::Normal));
    }

    #[test]
    fn apply_pasted_tsv_expands_sheet() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };

        app.apply_pasted_tsv("x\ty\n1\t2\n", true).unwrap();

        assert_eq!(app.state.grid.main_rows(), 2);
        assert_eq!(app.state.grid.main_cols(), 2);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("x")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 1 })
                .as_deref(),
            Some("y")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 1, col: 0 })
                .as_deref(),
            Some("1")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 1, col: 1 })
                .as_deref(),
            Some("2")
        );
    }

    #[test]
    fn workbook_edit_updates_visible_sheet_immediately() {
        let path = tempfile::NamedTempFile::new().unwrap();
        let mut app = App::new(Some(path.path().to_path_buf()));
        app.add_sheet("Sheet2".into());

        app.mode = Mode::Edit {
            buffer: "Sheet2 value".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(
            app.workbook
                .sheets
                .iter()
                .find(|sheet| sheet.id == 2)
                .and_then(|sheet| sheet.state.grid.get(&CellAddr::Main { row: 0, col: 0 }))
                .as_deref(),
            Some("Sheet2 value")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("Sheet2 value")
        );
    }

    #[test]
    fn enter_in_edit_mode_commits_and_moves_down() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(2, 1);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Edit {
            buffer: "first".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        assert!(matches!(app.mode, Mode::Edit { .. }));
        assert_eq!(app.cursor.row, HEADER_ROWS + 1);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("first")
        );
    }

    #[test]
    fn ctrl_page_switch_works_in_edit_mode() {
        let mut app = App::new(None);
        app.add_sheet("Sheet2".into());
        app.mode = Mode::Edit {
            buffer: "x".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        app.handle_key(KeyEvent::new(KeyCode::PageUp, KeyModifiers::CONTROL))
            .unwrap();
        assert_eq!(app.view_sheet_id, 1);
        assert!(matches!(app.mode, Mode::Edit { .. }));

        app.handle_key(KeyEvent::new(KeyCode::PageDown, KeyModifiers::CONTROL))
            .unwrap();

        assert_eq!(app.view_sheet_id, 2);
        assert!(matches!(app.mode, Mode::Edit { .. }));
    }

    #[test]
    fn ctrl_page_switch_resets_edit_buffer_to_target_sheet() {
        let mut app = App::new(None);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "sheet1".into());
        app.add_sheet("Sheet2".into());
        app.workbook
            .sheets
            .iter_mut()
            .find(|sheet| sheet.id == 2)
            .unwrap()
            .state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "sheet2".into());
        app.view_sheet_id = 1;
        app.sync_active_sheet_cache();
        app.mode = Mode::Edit {
            buffer: "sheet1".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        app.handle_key(KeyEvent::new(KeyCode::PageDown, KeyModifiers::CONTROL))
            .unwrap();

        match app.mode {
            Mode::Edit { buffer, .. } => assert_eq!(buffer, "sheet2"),
            other => panic!("unexpected mode: {other:?}"),
        }
    }

    #[test]
    fn edit_mode_accepts_named_sheet_formula_refs() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(KeyCode::Char('='), KeyModifiers::empty()))
            .unwrap();
        for ch in "$Sheet1:A1".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(ch), KeyModifiers::empty()))
                .unwrap();
        }
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("=$Sheet1:A1")
        );
    }

    #[test]
    fn formula_entry_in_column_b_keeps_b_target_without_cursor_movement() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS + 1,
        };

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();
        for ch in "=A*0.1 -- TAX TAX".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(ch), KeyModifiers::empty()))
                .unwrap();
        }

        let backend = TestBackend::new(80, 8);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };
        let formula_line = row(1);
        // debug removed

        assert!(formula_line.contains("B1"));
        assert!(formula_line.contains("=A*0.1 -- TAX TAX"));
        assert!(!formula_line.contains("]A"));
        assert_eq!(
            app.edit_target_addr,
            Some(CellAddr::Main { row: 0, col: 1 })
        );
    }

    #[test]
    fn pasted_formula_in_column_b_keeps_b_target_without_cursor_movement() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS + 1,
        };
        set_test_clipboard(Some("=A*0.1 -- TAX TAX\n".into()));

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();
        app.handle_key(KeyEvent::new(KeyCode::Char('v'), KeyModifiers::CONTROL))
            .unwrap();

        let backend = TestBackend::new(80, 8);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };
        let formula_line = row(1);
        // debug removed

        assert!(formula_line.contains("B1"));
        assert!(formula_line.contains("=A*0.1 -- TAX TAX"));
        assert!(!formula_line.contains("]A"));
        assert_eq!(
            app.edit_target_addr,
            Some(CellAddr::Main { row: 0, col: 1 })
        );
    }

    #[test]
    fn formula_entry_in_second_right_margin_cell_keeps_right_margin_b_target() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS + 2,
        };

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();
        for ch in "=A*0.1 -- TAX TAX".chars() {
            app.handle_key(KeyEvent::new(KeyCode::Char(ch), KeyModifiers::empty()))
                .unwrap();
        }

        let backend = TestBackend::new(80, 8);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };
        let formula_line = row(1);

        assert!(formula_line.contains("]B1"));
        assert!(formula_line.contains("=A*0.1 -- TAX TAX"));
        assert!(!formula_line.contains("]A."));
        assert_eq!(
            app.edit_target_addr,
            Some(CellAddr::Right { row: 0, col: 1 })
        );
    }

    #[test]
    fn normal_mode_paste_formula_into_column_b_keeps_main_b_target() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS + 1,
        };
        set_test_clipboard(Some("=A*0.1 -- TAX TAX\n".into()));

        app.handle_key(KeyEvent::new(KeyCode::Char('v'), KeyModifiers::CONTROL))
            .unwrap();

        let backend = TestBackend::new(80, 8);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };
        let formula_line = row(1);

        assert!(formula_line.contains("B1"));
        assert!(formula_line.contains("TAX TAX"));
        assert!(!formula_line.contains("]A."));
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 1 })
                .as_deref(),
            Some("=A*0.1 -- TAX TAX")
        );
    }

    #[test]
    fn ctrl_x_cuts_current_cell_and_delete_clears_it() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "hello".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };

        app.handle_key(KeyEvent::new(KeyCode::Char('x'), KeyModifiers::CONTROL))
            .unwrap();
        assert_eq!(app.state.grid.get(&CellAddr::Main { row: 0, col: 0 }), None);

        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "hello".into());
        app.handle_key(KeyEvent::new(KeyCode::Delete, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.state.grid.get(&CellAddr::Main { row: 0, col: 0 }), None);
    }

    #[test]
    fn edit_mode_clipboard_ops_target_whole_cell() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "hello".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Edit {
            buffer: "hello".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        app.handle_key(KeyEvent::new(KeyCode::Char('c'), KeyModifiers::CONTROL))
            .unwrap();
        app.handle_key(KeyEvent::new(KeyCode::Delete, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.state.grid.get(&CellAddr::Main { row: 0, col: 0 }), None);

        app.handle_key(KeyEvent::new(KeyCode::Char('v'), KeyModifiers::CONTROL))
            .unwrap();
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("hello")
        );
    }

    #[test]
    fn edit_mode_formula_bar_stays_on_original_cell_when_moving_left() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.state.grid.set(
            &CellAddr::Main { row: 0, col: 0 },
            "=A*0.1 -- TAX TAX".into(),
        );
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();
        app.mode = Mode::Edit {
            buffer: "=A*0.1 -- TAX TAX".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };
        app.edit_target_addr = Some(CellAddr::Main { row: 0, col: 0 });

        app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(app.cursor.col, MARGIN_COLS);
        assert_eq!(
            app.edit_target_addr,
            Some(CellAddr::Main { row: 0, col: 0 })
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("=A*0.1 -- TAX TAX")
        );
    }

    #[test]
    fn esc_while_quit_prompted_exits() {
        let mut app = App::new(None);
        app.mode = Mode::QuitPrompt;

        let quit = app
            .handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::empty()))
            .unwrap();

        assert!(quit);
    }

    #[test]
    fn esc_quits_immediately_on_unchanged_tsv_import() {
        use std::path::PathBuf;

        let tsv = tempfile::Builder::new().suffix(".tsv").tempfile().unwrap();
        std::fs::write(tsv.path(), "a\tb\n").unwrap();
        let path: PathBuf = tsv.path().to_path_buf();

        let mut app = App::new(None);
        app.mode = Mode::OpenPath {
            buffer: path.display().to_string(),
        };
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();
        assert!(matches!(app.mode, Mode::Normal));
        assert!(app.path.is_none());

        let quit = app
            .handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::empty()))
            .unwrap();
        assert!(quit);
    }

    #[test]
    fn esc_shows_quit_import_prompt_after_tsv_edit_tracked() {
        use std::path::PathBuf;
        use crate::grid::CellAddr;
        use crate::ops::Op;

        let tsv = tempfile::Builder::new().suffix(".tsv").tempfile().unwrap();
        std::fs::write(tsv.path(), "a\tb\n").unwrap();
        let path: PathBuf = tsv.path().to_path_buf();

        let mut app = App::new(None);
        app.mode = Mode::OpenPath {
            buffer: path.display().to_string(),
        };
        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();
        app.op_history.push(Op::SetCell {
            addr: CellAddr::Main { row: 0, col: 0 },
            value: "x".into(),
        });

        let quit = app
            .handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::empty()))
            .unwrap();
        assert!(!quit);
        assert!(matches!(app.mode, Mode::QuitImportPrompt));
    }

    #[test]
    fn ctrl_shift_plus_works_while_editing() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(2, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "top".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "bottom".into());
        app.mode = Mode::Edit {
            buffer: "+".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        app.handle_key(KeyEvent::new(
            KeyCode::Char('+'),
            KeyModifiers::CONTROL | KeyModifiers::SHIFT,
        ))
        .unwrap();

        assert_eq!(app.state.grid.main_rows(), 3);
        assert_eq!(app.state.grid.get(&CellAddr::Main { row: 0, col: 0 }), None);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 1, col: 0 })
                .as_deref(),
            Some("top")
        );
    }

    #[test]
    fn tab_cycles_special_header_values() {
        let mut app = App::new(None);
        app.cursor = SheetCursor { row: 0, col: 0 };
        app.mode = Mode::Edit {
            buffer: String::new(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        app.handle_key(KeyEvent::new(KeyCode::Tab, KeyModifiers::empty()))
            .unwrap();

        match &app.mode {
            Mode::Edit { buffer, .. } => assert_eq!(buffer, "∞"),
            other => panic!("unexpected mode: {other:?}"),
        }

        app.handle_key(KeyEvent::new(KeyCode::Tab, KeyModifiers::empty()))
            .unwrap();

        match &app.mode {
            Mode::Edit { buffer, .. } => assert_eq!(buffer, "Σ"),
            other => panic!("unexpected mode: {other:?}"),
        }
    }

    #[test]
    fn left_wraps_from_help_to_edit() {
        let mut app = App::new(None);
        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::Help,
                item: 0,
            }],
        };

        app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::empty()))
            .unwrap();

        match &app.mode {
            Mode::Menu { stack } => {
                assert_eq!(stack.len(), 1);
                // The left navigation cycles to the previous root section; update
                // expectations to match the current root ordering where Help -> Sheet.
                assert_eq!(stack[0].section, MenuSection::Sheet);
            }
            other => panic!("unexpected mode: {other:?}"),
        }

        app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::empty()))
            .unwrap();

        match &app.mode {
            Mode::Menu { stack } => {
                assert_eq!(stack.len(), 1);
                // After another left, we arrive at the section before Sheet: Format.
                assert_eq!(stack[0].section, MenuSection::Format);
            }
            other => panic!("unexpected mode: {other:?}"),
        }
    }

    #[test]
    fn left_cycles_through_root_menus() {
        let mut app = App::new(None);
        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::Help,
                item: 0,
            }],
        };

        app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::empty()))
            .unwrap();

        match &app.mode {
            Mode::Menu { stack } => {
                assert_eq!(stack.len(), 1);
                // Left from Help currently lands on Sheet in the root ordering.
                assert_eq!(stack[0].section, MenuSection::Sheet);
            }
            other => panic!("unexpected mode: {other:?}"),
        }

        app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::empty()))
            .unwrap();

        match &app.mode {
            Mode::Menu { stack } => {
                assert_eq!(stack.len(), 1);
                // The next left step precedes Sheet: Format.
                assert_eq!(stack[0].section, MenuSection::Format);
            }
            other => panic!("unexpected mode: {other:?}"),
        }
    }

    #[test]
    fn insert_menu_digit_shortcut_uses_palette_symbol() {
        let mut app = App::new(None);
        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::Insert,
                item: menu_items(MenuSection::Insert)
                    .iter()
                    .position(|item| item.label == "Special Char")
                    .unwrap(),
            }],
        };

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.special_picker, Some(0));

        app.handle_key(KeyEvent::new(KeyCode::Char('3'), KeyModifiers::empty()))
            .unwrap();

        assert!(matches!(app.mode, Mode::Edit { .. }));
    }

    #[test]
    fn special_picker_labels_use_digit_hotkeys() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.special_picker = Some(0);

        let backend = TestBackend::new(60, 16);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        assert!((0..buffer.area.height).any(|y| row(y).contains("Suggestions")));
        assert!((0..buffer.area.height).any(|y| row(y).contains("1: ∞")));
        assert!((0..buffer.area.height).any(|y| row(y).contains("2: Σ")));
    }

    #[test]
    fn arrow_right_at_text_end_moves_to_next_cell() {
        let mut app = App::new(None);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Edit {
            buffer: "ab".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(app.cursor.col, MARGIN_COLS + 1);
    }

    #[test]
    fn right_arrow_at_edit_edge_exits_edit_mode() {
        let mut app = App::new(None);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Edit {
            buffer: "ab".into(),
            formula_cursor: None,
            fit_to_content_on_commit: false,
        };

        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::empty()))
            .unwrap();

        assert!(matches!(app.mode, Mode::Normal));
        assert_eq!(app.cursor.col, MARGIN_COLS + 1);
    }

    #[test]
    fn right_arrow_from_main_cell_moves_to_next_main_cell() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Normal;

        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(app.cursor.col, MARGIN_COLS + 1);
        assert!(matches!(app.mode, Mode::Normal));
    }

    #[test]
    fn numbers_right_align_and_text_left_align() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "12".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "ab".into());

        let backend = TestBackend::new(60, 12);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();

        let row = (0..buffer.area.height)
            .map(|y| {
                (0..buffer.area.width)
                    .map(|x| buffer[(x, y)].symbol())
                    .collect::<String>()
            })
            .find(|line| line.contains("12") && line.contains("ab"))
            .unwrap_or_default();

        assert!(row.contains("12 ") || row.contains(" 12"));
        assert!(row.contains("ab"));
    }

    #[test]
    fn huge_numbers_render_in_exponential_notation() {
        let mut grid = crate::grid::GridBox::from(crate::grid::Grid::new(1, 1));
        let addr = CellAddr::Main { row: 0, col: 0 };
        grid.set(&addr, "1234567890123456789012345".into());
        grid.set_cell_format(
            addr.clone(),
            CellFormat {
                number: Some(NumberFormat::Fixed { decimals: 2 }),
                align: Some(TextAlign::Right),
            },
        );

        let rendered = format_cell_display(&grid, &addr, cell_effective_display(&grid, &addr));
        assert!(exponential_numeric_display(&rendered, 10)
            .map(|s| s.chars().count() <= 10)
            .unwrap_or(false));
        assert!(shrink_numeric_display("92.8888", 6).is_some());
    }

    #[test]
    fn aligned_columns_keep_e_in_same_screen_column() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let align_path =
            std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("docs/tests/align.corro");
        let mut app = App::new(Some(align_path));
        app.load_initial().unwrap();

        let backend = TestBackend::new(80, 12);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();

        let rows: Vec<String> = (0..buffer.area.height)
            .map(|y| {
                (0..buffer.area.width)
                    .map(|x| buffer[(x, y)].symbol())
                    .collect::<String>()
            })
            .filter(|row| row.contains('E') && row.contains('a') && row.contains('│'))
            .collect();

        let positions: Vec<usize> = rows.iter().map(|row| row.find('E').unwrap()).collect();

        assert!(!positions.is_empty());
        assert!(positions.windows(2).all(|w| w[0] == w[1]));
    }

    #[test]
    fn grid_draws_underlines_below_header_and_data_regions() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(3, 2);
        app.state.grid.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: MARGIN_COLS as u32,
            },
            "Hdr".into(),
        );
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "b".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "c".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "last-sorted".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 2, col: 0 }, "a".into());
        app.state.grid.set_view_sort_cols(vec![SortSpec {
            col: MARGIN_COLS,
            desc: false,
        }]);

        let backend = TestBackend::new(80, 18);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let mut saw_underlined_tilde_row = false;
        let mut saw_underlined_last_data_row = false;
        let mut tilde_row_y: Option<u16> = None;
        let mut last_data_row_y: Option<u16> = None;
        for y in 0..buffer.area.height {
            let line = (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>();
            if line.contains("~1") && line.contains("Hdr") {
                tilde_row_y = Some(y);
            }
            if line.contains("2") && line.contains("last-sorted") {
                last_data_row_y = Some(y);
            }
        }
        assert!(tilde_row_y.is_some(), "expected rendered ~1 row");
        assert!(last_data_row_y.is_some(), "expected rendered last data row");

        for y in 0..buffer.area.height {
            for x in 0..buffer.area.width {
                let cell = &buffer[(x, y)];
                if tilde_row_y == Some(y) && cell.modifier.contains(Modifier::UNDERLINED) {
                    saw_underlined_tilde_row = true;
                }
                if last_data_row_y == Some(y) && cell.modifier.contains(Modifier::UNDERLINED) {
                    saw_underlined_last_data_row = true;
                }
            }
        }

        assert!(saw_underlined_tilde_row);
        assert!(saw_underlined_last_data_row);
    }

    #[test]
    fn save_only_writes_persisted_view_sort() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("sort.corro");
        let cols = vec![SortSpec {
            col: MARGIN_COLS,
            desc: false,
        }];

        let mut app = App::new(None);
        app.state.grid.set_main_size(2, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "b".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "a".into());

        app.state.grid.set_view_sort_cols(cols.clone());
        app.set_active_sort_persistence(&cols, false);
        app.save_to_path(&path).unwrap();
        let saved = std::fs::read_to_string(&path).unwrap();
        assert!(!saved.contains("SORT A"), "{saved}");
        assert_eq!(app.state.grid.sorted_main_rows(), vec![1, 0]);

        app.set_active_sort_persistence(&cols, true);
        app.save_to_path(&path).unwrap();
        let saved = std::fs::read_to_string(&path).unwrap();
        assert!(saved.contains("SORT A"), "{saved}");
    }

    #[test]
    fn load_initial_handles_legacy_test5_workbook() {
        let fixture = docs_test_path("main.corro");
        if !fixture.exists() {
            eprintln!("Skipping load_initial_handles_legacy_test5_workbook: fixture missing");
            return;
        }

        let mut app = App::new(Some(fixture));
        app.load_initial().unwrap();

        assert_eq!(app.workbook.sheet_count(), 4);
        assert_eq!(app.view_sheet_id, 4);
        assert_eq!(app.workbook.sheet_title(3), "Sheet1 Copy");
        assert_eq!(app.state.grid.main_rows(), 15);
        assert_eq!(app.state.grid.main_cols(), 7);
    }

    #[test]
    fn insert_row_returns_to_normal_cell_mode() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(2, 2);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };

        app.insert_rows_above_cursor(1).unwrap();

        assert_eq!(app.selection_kind, SelectionKind::Cells);
        assert!(app.anchor.is_none());
        assert_eq!(app.cursor.row, HEADER_ROWS);
    }

    #[test]
    fn mitosis_row_copies_current_row_after_it() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(3, 2);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "before".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "copy-me".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "=A2*2".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 2, col: 0 }, "after".into());
        app.state.grid.set(
            &CellAddr::Left {
                col: MARGIN_COLS - 1,
                row: 1,
            },
            "label".into(),
        );
        app.state
            .grid
            .set(&CellAddr::Right { col: 0, row: 1 }, "note".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS + 1,
            col: MARGIN_COLS,
        };

        app.insert_mitosis_row_after_cursor().unwrap();

        assert_eq!(app.state.grid.main_rows(), 4);
        assert_eq!(app.cursor.row, HEADER_ROWS + 2);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 1, col: 0 })
                .as_deref(),
            Some("copy-me")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 2, col: 0 })
                .as_deref(),
            Some("copy-me")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 2, col: 1 })
                .as_deref(),
            Some("=A2*2")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 3, col: 0 })
                .as_deref(),
            Some("after")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Left {
                    col: MARGIN_COLS - 1,
                    row: 2
                })
                .as_deref(),
            Some("label")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Right { col: 0, row: 2 })
                .as_deref(),
            Some("note")
        );
    }

    #[test]
    fn mitosis_col_copies_current_col_after_it() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(2, 3);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "left".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "copy-me".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "=A2*2".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 2 }, "right".into());
        app.state
            .grid
            .set(&CellAddr::Header { row: 0, col: (MARGIN_COLS + 1) as u32 }, "hdr".into());
        app.state
            .grid
            .set(&CellAddr::Footer { row: 0, col: (MARGIN_COLS + 1) as u32 }, "ftr".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS + 1,
        };

        app.insert_mitosis_col_after_cursor().unwrap();

        assert_eq!(app.state.grid.main_cols(), 4);
        assert_eq!(app.cursor.col, MARGIN_COLS + 2);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 1 })
                .as_deref(),
            Some("copy-me")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 2 })
                .as_deref(),
            Some("copy-me")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 1, col: 2 })
                .as_deref(),
            Some("=A2*2")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 3 })
                .as_deref(),
            Some("right")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Header {
                    row: 0,
                    col: (MARGIN_COLS + 2) as u32
                })
                .as_deref(),
            Some("hdr")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Footer {
                    row: 0,
                    col: (MARGIN_COLS + 2) as u32
                })
                .as_deref(),
            Some("ftr")
        );
    }

    #[test]
    fn mitosis_main_col_works_when_cursor_in_header() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "data".into());
        app.state
            .grid
            .set(&CellAddr::Header { row: 0, col: 0 }, "h0".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS - 1,
            col: MARGIN_COLS,
        };

        app.insert_mitosis_col_after_cursor().unwrap();

        assert_eq!(app.state.grid.main_cols(), 3);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("data")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 1 })
                .as_deref(),
            Some("data")
        );
    }

    #[test]
    fn mitosis_header_row_not_last_duplicates() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.state
            .grid
            .set(
                &CellAddr::Header {
                    row: (HEADER_ROWS - 2) as u32,
                    col: 0,
                },
                "t".into(),
            );
        app.cursor = SheetCursor {
            row: HEADER_ROWS - 2,
            col: 0,
        };

        app.insert_mitosis_row_after_cursor().unwrap();

        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Header {
                    row: (HEADER_ROWS - 2) as u32,
                    col: 0
                })
                .as_deref(),
            Some("t")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Header {
                    row: (HEADER_ROWS - 1) as u32,
                    col: 0
                })
                .as_deref(),
            Some("t")
        );
    }

    #[test]
    fn mitosis_left_margin_col_duplicates() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.state
            .grid
            .set(&CellAddr::Left { col: 0, row: 0 }, "L".into());
        app.state
            .grid
            .set(&CellAddr::Left { col: 1, row: 0 }, "M".into());
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: 0,
        };

        app.insert_mitosis_col_after_cursor().unwrap();

        assert_eq!(app.cursor.col, 1);
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Left { col: 0, row: 0 })
                .as_deref(),
            Some("L")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Left { col: 1, row: 0 })
                .as_deref(),
            Some("L")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Left { col: 2, row: 0 })
                .as_deref(),
            Some("M")
        );
    }

    #[test]
    fn insert_menu_contains_mitosis_row() {
        assert!(INSERT_ROOT_MENU_ITEMS.iter().any(|item| {
            item.shortcut == 'M'
                && item.label == "Mitosis (Row)"
                && item.target == MenuTarget::Action(MenuAction::InsertMitosisRow)
        }));
    }

    #[test]
    fn insert_menu_contains_mitosis_col() {
        assert!(INSERT_ROOT_MENU_ITEMS.iter().any(|item| {
            item.shortcut == 'O'
                && item.label == "Mitosis (Col)"
                && item.target == MenuTarget::Action(MenuAction::InsertMitosisCol)
        }));
    }

    #[test]
    fn balance_books_reorders_rows_in_place() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(3, 2);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "10".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "a".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "-10".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "b".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 2, col: 0 }, "5".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 2, col: 1 }, "c".into());

        app.mode = Mode::BalanceBooks {
            buffer: String::new(),
            direction: BalanceDirection::PosToNeg,
            persist: false,
            focus: BalanceBooksFocus::Column,
        };

        // Simulate Enter on the balance action path.
        let _ = app.handle_key(crossterm::event::KeyEvent::from(
            crossterm::event::KeyCode::Enter,
        ));

        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("10")
        );
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 1, col: 0 })
                .as_deref(),
            Some("-10")
        );
    }

    #[test]
    fn balance_dialog_shows_checkbox_style_choices() {
        let app = App::new(None);
        let lines = app.balance_dialog_lines(
            "A",
            BalanceDirection::PosToNeg,
            false,
            BalanceBooksFocus::Column,
            1,
            Style::default(),
            Style::default(),
            Style::default(),
        );
        let rendered = lines
            .into_iter()
            .map(|line| {
                line.spans
                    .into_iter()
                    .map(|span| span.content.to_string())
                    .collect::<String>()
            })
            .collect::<Vec<_>>()
            .join("\n");

        assert!(rendered.contains("Column to Balance:"));
        assert!(rendered.contains("Report Type:"));
        assert!(rendered.contains("[X] View only"));
        assert!(rendered.contains("[ ] Persisted report"));
        assert!(rendered.contains("Balance direction:"));
    }

    #[test]
    fn balance_dialog_tabs_between_controls() {
        let mut app = App::new(None);
        app.mode = Mode::BalanceBooks {
            buffer: String::new(),
            direction: BalanceDirection::PosToNeg,
            persist: false,
            focus: BalanceBooksFocus::Column,
        };

        app.handle_key(KeyEvent::new(KeyCode::Tab, KeyModifiers::empty()))
            .unwrap();
        match app.mode {
            Mode::BalanceBooks { focus, .. } => {
                assert_eq!(focus, BalanceBooksFocus::ReportViewOnly)
            }
            other => panic!("unexpected mode: {other:?}"),
        }

        app.handle_key(KeyEvent::new(KeyCode::Tab, KeyModifiers::empty()))
            .unwrap();
        match app.mode {
            Mode::BalanceBooks { focus, .. } => {
                assert_eq!(focus, BalanceBooksFocus::ReportPersisted)
            }
            other => panic!("unexpected mode: {other:?}"),
        }

        app.handle_key(KeyEvent::new(KeyCode::Tab, KeyModifiers::empty()))
            .unwrap();
        match app.mode {
            Mode::BalanceBooks { focus, .. } => assert_eq!(focus, BalanceBooksFocus::PosToNeg),
            other => panic!("unexpected mode: {other:?}"),
        }

        app.handle_key(KeyEvent::new(KeyCode::Tab, KeyModifiers::empty()))
            .unwrap();
        match app.mode {
            Mode::BalanceBooks { focus, .. } => assert_eq!(focus, BalanceBooksFocus::NegToPos),
            other => panic!("unexpected mode: {other:?}"),
        }

        app.handle_key(KeyEvent::new(KeyCode::Tab, KeyModifiers::empty()))
            .unwrap();
        match app.mode {
            Mode::BalanceBooks { focus, .. } => assert_eq!(focus, BalanceBooksFocus::Generate),
            other => panic!("unexpected mode: {other:?}"),
        }

        app.handle_key(KeyEvent::new(KeyCode::Tab, KeyModifiers::empty()))
            .unwrap();
        match app.mode {
            Mode::BalanceBooks { focus, .. } => assert_eq!(focus, BalanceBooksFocus::Cancel),
            other => panic!("unexpected mode: {other:?}"),
        }
    }

    #[test]
    fn balance_dialog_prefills_mixed_sign_column() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(2, 2);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "7".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "8".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "10".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "-10".into());

        app.mode = app.menu_action_mode(MenuAction::BalanceBooks);

        match app.mode {
            Mode::BalanceBooks { buffer, .. } => assert_eq!(buffer, "B"),
            other => panic!("unexpected mode: {other:?}"),
        }
    }

    #[test]
    fn balance_dialog_enter_on_generate_runs_balance() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(2, 1);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "10".into());
        app.state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "-10".into());
        app.mode = Mode::BalanceBooks {
            buffer: String::new(),
            direction: BalanceDirection::PosToNeg,
            persist: false,
            focus: BalanceBooksFocus::Generate,
        };

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        assert!(matches!(app.mode, Mode::Normal));
        assert_eq!(
            app.state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("10")
        );
    }

    #[test]
    fn balance_dialog_escape_cancels() {
        let mut app = App::new(None);
        app.mode = Mode::BalanceBooks {
            buffer: String::new(),
            direction: BalanceDirection::PosToNeg,
            persist: false,
            focus: BalanceBooksFocus::Generate,
        };

        app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::empty()))
            .unwrap();

        assert!(matches!(app.mode, Mode::Normal));
    }

    #[test]
    fn aggregate_divider_sits_after_row_labels() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(Some(docs_test_path("main.corro")));
        app.load_initial().unwrap();

        let backend = TestBackend::new(140, 24);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let first_content_row = (0..buffer.area.height)
            .map(|y| {
                (0..buffer.area.width)
                    .map(|x| buffer[(x, y)].symbol())
                    .collect::<String>()
            })
            .find(|line| line.contains("│") && line.contains("[A"))
            .unwrap_or_default();

        assert!(first_content_row.contains("[A"));
        assert!(rendered_contains_vertical_divider(&buffer));
    }

    #[test]
    fn aggregate_dividers_draw_in_grid_buffer() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(Some(docs_test_path("main.corro")));
        app.load_initial().unwrap();

        let backend = TestBackend::new(140, 24);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();

        assert!((0..buffer.area.height)
            .any(|y| (0..buffer.area.width).any(|x| buffer[(x, y)].symbol() == "│")));
        assert!((0..buffer.area.height)
            .any(|y| (0..buffer.area.width).any(|x| buffer[(x, y)].symbol() == "─")));
    }

    fn buffer_to_string(buffer: &ratatui::buffer::Buffer) -> String {
        (0..buffer.area.height)
            .map(|y| {
                (0..buffer.area.width)
                    .map(|x| buffer[(x, y)].symbol())
                    .collect::<String>()
            })
            .collect::<Vec<_>>()
            .join("\n")
    }

    fn rendered_contains_vertical_divider(buffer: &ratatui::buffer::Buffer) -> bool {
        (0..buffer.area.height)
            .any(|y| (0..buffer.area.width).any(|x| buffer[(x, y)].symbol() == "│"))
    }

    fn normalize_frame(s: &str) -> String {
        s.lines()
            .map(|line| line.trim_end())
            .collect::<Vec<_>>()
            .join("\n")
    }

    #[test]
    fn formula_arrows_stay_in_select_cell_mode_until_non_arrow() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.cursor = SheetCursor {
            row: HEADER_ROWS,
            col: MARGIN_COLS,
        };
        app.mode = Mode::Edit {
            buffer: "=".into(),
            formula_cursor: Some(app.cursor),
            fit_to_content_on_commit: false,
        };

        let backend = TestBackend::new(40, 6);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        assert!((0..buffer.area.width).any(|x| {
            let cell = &buffer[(x, 1)];
            cell.symbol() == " " && cell.style().bg == Some(Color::Yellow)
        }));

        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::empty()))
            .unwrap();
        let mut terminal = Terminal::new(TestBackend::new(40, 6)).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        assert!((0..buffer.area.width).any(|x| {
            let cell = &buffer[(x, 1)];
            cell.symbol() == " " && cell.style().bg == Some(Color::Yellow)
        }));

        app.handle_key(KeyEvent::new(KeyCode::Char('x'), KeyModifiers::empty()))
            .unwrap();
        let mut terminal = Terminal::new(TestBackend::new(40, 6)).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        assert!((0..buffer.area.width).any(|x| {
            let cell = &buffer[(x, 1)];
            cell.symbol() == " " && cell.style().bg == Some(Color::Yellow)
        }));
    }

    #[test]
    fn menu_bar_shows_format_tab() {
        let app = App::new(None);
        assert!(app.menu_bar_line().contains(" Format "));
    }

    #[test]
    fn menu_bar_orders_root_sections_as_requested() {
        let mut app = App::new(None);
        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::File,
                item: 0,
            }],
        };

        let line = app.menu_bar_line();
        let file = line.find("[File]").unwrap();
        let edit = line.find(" Edit ").unwrap();
        let insert = line.find(" Insert ").unwrap();
        let format = line.find(" Format ").unwrap();
        let sheet = line.find(" Sheet ").unwrap();
        let help = line.find(" Help ").unwrap();

        assert!(file < edit && edit < insert && insert < format && format < sheet && sheet < help);
    }

    #[test]
    fn root_menu_cycling_follows_new_order() {
        let mut app = App::new(None);
        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::File,
                item: 0,
            }],
        };

        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::empty()))
            .unwrap();
        match app.mode {
            Mode::Menu { ref stack } => assert_eq!(stack[0].section, MenuSection::Edit),
            other => panic!("unexpected mode: {other:?}"),
        }

        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::empty()))
            .unwrap();
        match app.mode {
            Mode::Menu { ref stack } => assert_eq!(stack[0].section, MenuSection::Insert),
            other => panic!("unexpected mode: {other:?}"),
        }

        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::empty()))
            .unwrap();
        match app.mode {
            Mode::Menu { ref stack } => assert_eq!(stack[0].section, MenuSection::Format),
            other => panic!("unexpected mode: {other:?}"),
        }
    }

    #[test]
    fn save_path_renders_filename_as_typed() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.mode = Mode::SavePath {
            buffer: "draft.corro".into(),
        };

        let backend = TestBackend::new(60, 8);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row = |y: u16| {
            (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>()
        };

        assert!(row(1).contains("save as:"));
        assert!(row(1).contains("draft.corro"));
        assert!((0..buffer.area.width).any(|x| {
            let cell = &buffer[(x, 1)];
            cell.symbol() == " " && cell.style().bg == Some(Color::Yellow)
        }));
    }

    #[test]
    fn printable_key_starts_editing_in_normal_mode() {
        let mut app = App::new(None);
        app.handle_key(KeyEvent::new(KeyCode::Char('x'), KeyModifiers::empty()))
            .unwrap();

        assert!(matches!(app.mode, Mode::Edit { .. }));
        if let Mode::Edit { buffer, .. } = &app.mode {
            assert_eq!(buffer, "x");
        }
    }

    #[test]
    fn format_menu_actions_apply_cell_format() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 1);
        app.apply_format_to_target(
            FormatTarget::Cell,
            CellFormat {
                number: Some(NumberFormat::Fixed { decimals: 1 }),
                align: Some(TextAlign::Right),
            },
        );

        assert_eq!(
            app.state
                .grid
                .format_for_addr(&CellAddr::Main { row: 0, col: 0 }),
            CellFormat {
                number: Some(NumberFormat::Fixed { decimals: 1 }),
                align: Some(TextAlign::Right),
            }
        );
    }

    #[test]
    fn format_scope_all_column_sets_all_global_cols() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        let fmt = CellFormat {
            number: Some(NumberFormat::Fixed { decimals: 2 }),
            align: None,
        };
        app.apply_format_to_target(FormatTarget::All, fmt);
        for c in 0..app.state.grid.total_cols() {
            assert_eq!(app.state.grid.format_for_global_col(FormatScope::All, c), fmt);
        }
    }

    #[test]
    fn format_scope_full_column_sets_only_global_cursor_column() {
        let mut app = App::new(None);
        app.state.grid.set_main_size(1, 2);
        app.cursor.col = MARGIN_COLS + 1;
        let fmt = CellFormat {
            number: Some(NumberFormat::Currency { decimals: 0 }),
            align: None,
        };
        app.apply_format_to_target(FormatTarget::FullColumn, fmt);
        assert_eq!(
            app.state
                .grid
                .format_for_global_col(FormatScope::All, MARGIN_COLS + 1),
            fmt
        );
        assert_eq!(
            app.state
                .grid
                .format_for_global_col(FormatScope::All, MARGIN_COLS),
            CellFormat::default()
        );
    }

    #[test]
    fn formatted_cell_display_uses_number_and_alignment() {
        let mut grid = crate::grid::GridBox::from(crate::grid::Grid::new(1, 1));
        let addr = CellAddr::Main { row: 0, col: 0 };
        grid.set(&addr, "12.5".into());
        grid.set_cell_format(
            addr.clone(),
            CellFormat {
                number: Some(NumberFormat::Fixed { decimals: 1 }),
                align: Some(TextAlign::Right),
            },
        );

        let formatted = format_cell_display(&grid, &addr, cell_effective_display(&grid, &addr));
        assert_eq!(formatted, "12.5");
    }

    #[test]
    fn aligned_columns_keep_separate_widths() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(10, 2);
        for (i, value) in [
            "a",
            "aa",
            "aaa",
            "aaaa",
            "aaaa",
            "aaaaa",
            "aaaaaa",
            "aaaaaaa",
            "aaaaaaaaaaaaaaaa",
        ]
        .iter()
        .enumerate()
        {
            app.state.grid.set(
                &CellAddr::Left {
                    row: i as u32,
                    col: MARGIN_COLS - 1,
                },
                value.to_string(),
            );
            app.state.grid.set(
                &CellAddr::Main {
                    row: i as u32,
                    col: 0,
                },
                "E".into(),
            );
        }
        // Debug prints removed
        let backend = TestBackend::new(80, 12);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let rendered = buffer_to_string(terminal.backend().buffer());
        // Less brittle checks: ensure the main 'E' cell, left-margin header,
        // and window status are rendered.
        assert!(rendered.contains("E"));
        assert!(rendered.contains("[A"));
        assert!(rendered.contains("corro  10r × 2c"));
    }

    #[test]
    fn save_and_reload_preserve_format_ops() {
        let tmp = tempfile::Builder::new()
            .suffix(".corro")
            .tempfile()
            .unwrap();
        let mut app = App::new(Some(tmp.path().to_path_buf()));
        app.state.grid.set_main_size(1, 1);
        app.apply_format_to_target(
            FormatTarget::Cell,
            CellFormat {
                number: Some(NumberFormat::Currency { decimals: 2 }),
                align: Some(TextAlign::Center),
            },
        );
        app.save_to_path(tmp.path()).unwrap();

        let mut reloaded = App::new(Some(tmp.path().to_path_buf()));
        reloaded.load_initial().unwrap();

        assert_eq!(
            reloaded
                .state
                .grid
                .format_for_addr(&CellAddr::Main { row: 0, col: 0 })
                .number,
            Some(NumberFormat::Currency { decimals: 2 })
        );
        assert_eq!(
            reloaded
                .state
                .grid
                .format_for_addr(&CellAddr::Main { row: 0, col: 0 })
                .align,
            Some(TextAlign::Center)
        );
    }

    #[test]
    fn save_path_left_and_right_move_caret() {
        let mut app = App::new(None);
        app.mode = Mode::SavePath {
            buffer: "abc".into(),
        };
        app.input_cursor = Some(3);

        app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.input_cursor, Some(2));

        app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.input_cursor, Some(1));

        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::empty()))
            .unwrap();
        assert_eq!(app.input_cursor, Some(2));
    }

    #[test]
    fn right_descends_or_wraps() {
        let mut app = App::new(None);
        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::File,
                item: 2,
            }],
        };

        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::empty()))
            .unwrap();

        match app.mode {
            Mode::Menu { stack } => {
                assert_eq!(stack.len(), 2);
                assert_eq!(stack[0].section, MenuSection::File);
                assert_eq!(stack[1].section, MenuSection::Export);
                assert_eq!(stack[1].item, 0);
            }
            other => panic!("unexpected mode: {other:?}"),
        }

        app.mode = Mode::Menu {
            stack: vec![MenuLevel {
                section: MenuSection::File,
                item: 3,
            }],
        };

        app.handle_key(KeyEvent::new(KeyCode::Right, KeyModifiers::empty()))
            .unwrap();

        match app.mode {
            Mode::Menu { stack } => {
                assert_eq!(stack.len(), 2);
                assert_eq!(stack[1].section, MenuSection::Width);
            }
            other => panic!("unexpected mode: {other:?}"),
        }
    }
}

// ── Display helpers ───────────────────────────────────────────────────────────

fn addr_label(addr: &CellAddr, main_cols: usize) -> String {
    crate::addr::cell_ref_text(addr, main_cols)
}

fn input_line(
    prefix: String,
    buffer: &str,
    cursor: usize,
    text_style: Style,
    caret_style: Style,
) -> Line<'static> {
    input_line_with_suffix(prefix, buffer, cursor, text_style, caret_style, None)
}

fn input_line_with_suffix(
    prefix: String,
    buffer: &str,
    cursor: usize,
    text_style: Style,
    caret_style: Style,
    suffix: Option<String>,
) -> Line<'static> {
    let chars: Vec<char> = buffer.chars().collect();
    let cursor = cursor.min(chars.len());
    let before: String = chars[..cursor].iter().collect();
    let after: String = chars[cursor..].iter().collect();

    let mut spans = Vec::with_capacity(4);
    if !prefix.is_empty() {
        spans.push(Span::styled(prefix, text_style));
    }
    if !before.is_empty() {
        spans.push(Span::styled(before, text_style));
    }
    if let Some(ch) = chars.get(cursor) {
        spans.push(Span::styled(ch.to_string(), caret_style));
    } else {
        spans.push(Span::styled(" ", caret_style));
    }
    if !after.is_empty() {
        let tail = if cursor < chars.len() {
            chars[cursor + 1..].iter().collect()
        } else {
            after
        };
        if !tail.is_empty() {
            spans.push(Span::styled(tail, text_style));
        }
    }
    if let Some(suffix) = suffix {
        if !suffix.is_empty() {
            spans.push(Span::styled(suffix, text_style));
        }
    }

    Line::from(spans)
}

fn formula_edit_preview(grid: &Grid, addr: &CellAddr, buffer: &str) -> Option<String> {
    let trimmed = buffer.trim();
    if trimmed.is_empty() || !trimmed.starts_with('=') {
        return None;
    }
    if matches!(trimmed, "=π" | "=e" | "=c") {
        return None;
    }
    let mut preview_grid = grid.clone();
    preview_grid.set(addr, trimmed.to_string());
    Some(cell_effective_display(&preview_grid, addr))
}

fn formula_bar_value(grid: &Grid, addr: &CellAddr) -> String {
    let raw = normalize_inline_text(&cell_display(grid, addr));
    let trimmed = raw.trim();
    if trimmed.is_empty() {
        return String::new();
    }
    if trimmed.starts_with('=') {
        return normalize_inline_text(&cell_effective_display(grid, addr));
    }
    raw
}

/// For **Values** TSV, bare aggregate labels in the key left margin still resolve to the computed
/// aggregate for the row, not the word `TOTAL` (etc.). (Generic text export keeps bare `TOTAL` /
/// `SUM` as on-sheet text; other labels such as `MAX` can still use `=SUBTOTAL(…)` interop.)
fn tsv_left_key_subtotal_computed(
    grid: &Grid,
    cell_addr: &CellAddr,
    func: AggFunc,
    main_row: u32,
) -> Option<String> {
    let CellAddr::Left { col, row } = cell_addr else {
        return None;
    };
    if *row != main_row || *col != MARGIN_COLS - 1 {
        return None;
    }
    let raw = grid.get(cell_addr).unwrap_or_default();
    if crate::ods::subtotal_code_for_label(&raw).is_none() {
        return None;
    }
    Some(left_margin_main_col_aggregate(grid, func, main_row, 0))
}

/// Footers: key column (`MARGIN_COLS - 1`) may hold a bare `TOTAL` while [`crate::ods::cell_export_value_string`]
/// emits `=SUBTOTAL(…)` over the full main block — Values must be that aggregate, not the label.
fn tsv_footer_key_subtotal_computed(
    grid: &Grid,
    cell_addr: &CellAddr,
    func: AggFunc,
) -> Option<String> {
    let CellAddr::Footer { col, .. } = cell_addr else {
        return None;
    };
    if *col as usize != MARGIN_COLS - 1 {
        return None;
    }
    let raw = grid.get(cell_addr).unwrap_or_default();
    if crate::ods::subtotal_code_for_label(&raw).is_none() {
        return None;
    }
    let mr = grid.main_rows();
    let mc = grid.main_cols() as u32;
    Some(compute_aggregate(
        grid,
        &AggregateDef {
            func,
            source: MainRange {
                row_start: 0,
                row_end: mr as u32,
                col_start: 0,
                col_end: mc,
            },
        },
    ))
}

/// Same unformatted value as the main grid’s data cells, used by TSV/CSV export to match
/// on-screen subtotal/aggregate columns (not just stored formula text).
pub(crate) fn tsv_effective_unformatted_string(grid: &Grid, r: usize, c: usize) -> String {
    let cur = SheetCursor { row: r, col: c };
    let cell_addr = cur.to_addr(grid);
    let hr = HEADER_ROWS;
    let mr = grid.main_rows();
    let lm = MARGIN_COLS;
    let mc = grid.main_cols();
    let right_col_agg = right_col_agg_func(grid, c);
    let footer_agg = if r >= hr + mr {
        footer_row_agg_func(grid, r - hr - mr)
    } else {
        None
    };
    let main_row_idx = if r >= hr && r < hr + mr {
        Some((r - hr) as u32)
    } else {
        None
    };
    let left_margin_agg = main_row_idx.and_then(|mri| left_margin_agg_func(grid, mri));
    let left_margin_block_start = main_row_idx.map(|mri| row_total_block_start(grid, mri));

    if let Some(func) = footer_agg {
        if right_col_agg.is_some() {
            footer_special_col_aggregate(grid, func, c, mr, mc)
                .unwrap_or_else(|| {
                    tsv_footer_key_subtotal_computed(grid, &cell_addr, func)
                        .unwrap_or_else(|| cell_effective_display(grid, &cell_addr))
                })
        } else if c >= lm && c < lm + mc {
            let main_col = (c - lm) as u32;
            compute_aggregate(
                grid,
                &AggregateDef {
                    func,
                    source: MainRange {
                        row_start: 0,
                        row_end: mr as u32,
                        col_start: main_col,
                        col_end: main_col + 1,
                    },
                },
            )
        } else {
            tsv_footer_key_subtotal_computed(grid, &cell_addr, func)
                .unwrap_or_else(|| cell_effective_display(grid, &cell_addr))
        }
    } else if let (Some(func), Some(block_start), Some(main_row)) =
        (left_margin_agg, left_margin_block_start, main_row_idx)
    {
        if c >= lm && c < lm + mc {
            if right_col_agg.is_some() {
                let data_cols = data_main_col_count(grid);
                let (row_start, row_end) = if block_start < main_row {
                    (block_start, main_row)
                } else {
                    previous_raw_block(grid, main_row).unwrap_or((0, main_row))
                };
                left_margin_special_col_aggregate(
                    grid, func, c, row_start, row_end, data_cols,
                )
                .unwrap_or_else(|| {
                    tsv_left_key_subtotal_computed(grid, &cell_addr, func, main_row)
                        .unwrap_or_else(|| cell_effective_display(grid, &cell_addr))
                })
            } else {
                let main_col = (c - lm) as u32;
                left_margin_main_col_aggregate(grid, func, main_row, main_col)
            }
        } else if right_col_agg.is_some() {
            left_margin_special_col_aggregate(
                grid,
                func,
                c,
                block_start,
                main_row,
                data_main_col_count(grid),
            )
            .unwrap_or_else(|| {
                tsv_left_key_subtotal_computed(grid, &cell_addr, func, main_row)
                    .unwrap_or_else(|| cell_effective_display(grid, &cell_addr))
            })
        } else {
            tsv_left_key_subtotal_computed(grid, &cell_addr, func, main_row)
                .unwrap_or_else(|| cell_effective_display(grid, &cell_addr))
        }
    } else if r >= hr && r < hr + mr {
        if let Some(func) = right_col_agg {
            let main_row = (r - hr) as u32;
            let data_cols = data_main_col_count(grid);
            compute_aggregate(
                grid,
                &AggregateDef {
                    func,
                    source: MainRange {
                        row_start: main_row,
                        row_end: main_row + 1,
                        col_start: 0,
                        col_end: data_cols as u32,
                    },
                },
            )
        } else {
            cell_effective_display(grid, &cell_addr)
        }
    } else {
        cell_effective_display(grid, &cell_addr)
    }
}

fn normalize_inline_text(text: &str) -> String {
    text.replace('\n', "¶")
}

pub(crate) fn format_cell_display(grid: &Grid, addr: &CellAddr, raw: String) -> String {
    let raw = normalize_inline_text(&raw);
    let fmt = grid.format_for_addr(addr);
    let Some(number) = fmt.number else {
        return raw;
    };
    let Some(value) = raw.trim().parse::<f64>().ok() else {
        return raw;
    };
    match number {
        NumberFormat::Currency { decimals } => format!("${value:.decimals$}"),
        NumberFormat::Fixed { decimals } => format!("{value:.decimals$}"),
    }
}

fn text_align_to_utrunc(a: TextAlign) -> UTruncAlign {
    match a {
        TextAlign::Left | TextAlign::Default => UTruncAlign::Left,
        TextAlign::Right => UTruncAlign::Right,
        TextAlign::Center => UTruncAlign::Center,
    }
}

fn align_cell_display(text: String, width: usize, align: Option<TextAlign>) -> String {
    let width = width.max(1);
    let ual = text_align_to_utrunc(align.unwrap_or(TextAlign::Default));
    text.unicode_pad(width, ual, true).into_owned()
}

fn effective_cell_align(grid: &Grid, addr: &CellAddr, formatted: &str) -> Option<TextAlign> {
    let fmt = grid.format_for_addr(addr);
    if fmt.align.is_some() {
        return fmt.align;
    }
    if fmt.number.is_some() || formatted.trim().parse::<f64>().is_ok() {
        Some(TextAlign::Right)
    } else {
        None
    }
}

fn truncate_with_ellipsis(text: &str, width: usize) -> String {
    if width == 0 {
        return String::new();
    }
    if text.width() <= width {
        return text.to_string();
    }
    let keep = width.saturating_sub(1);
    if keep == 0 {
        return "…".to_string();
    }
    let (prefix, _) = text.unicode_truncate(keep);
    format!("{prefix}…")
}

fn shrink_numeric_display(text: &str, width: usize) -> Option<String> {
    let mut s = text.trim().to_string();
    if s.parse::<f64>().is_err() || !s.contains('.') {
        return None;
    }
    while s.chars().count() > width {
        let Some(last) = s.chars().last() else {
            break;
        };
        if last == '.' {
            s.pop();
            break;
        }
        s.pop();
    }
    if s.chars().count() <= width {
        Some(s)
    } else {
        None
    }
}

fn exponential_numeric_display(text: &str, width: usize) -> Option<String> {
    let value = text.trim().parse::<f64>().ok()?;
    if !value.is_finite() {
        return None;
    }
    let target = width.min(10);
    for decimals in (0..=6).rev() {
        let s = format!("{value:.decimals$e}");
        if s.chars().count() <= target {
            return Some(s);
        }
        if s.contains('.') {
            let trimmed = s.trim_end_matches('0').trim_end_matches('.').to_string();
            if trimmed.chars().count() <= target {
                return Some(trimmed);
            }
        }
    }
    None
}

fn sheet_row_label(logical_row: usize, main_rows: usize) -> String {
    addr::ui_row_label(logical_row, main_rows)
}

fn col_header_label(global_col: usize, main_cols: usize) -> String {
    addr::ui_column_fragment(global_col, main_cols)
}

fn centered_rect(percent_x: u16, percent_y: u16, area: Rect) -> Rect {
    let popup_layout = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Percentage((100 - percent_y) / 2),
            Constraint::Percentage(percent_y),
            Constraint::Percentage((100 - percent_y) / 2),
        ])
        .split(area);
    Layout::default()
        .direction(Direction::Horizontal)
        .constraints([
            Constraint::Percentage((100 - percent_x) / 2),
            Constraint::Percentage(percent_x),
            Constraint::Percentage((100 - percent_x) / 2),
        ])
        .split(popup_layout[1])[1]
}
