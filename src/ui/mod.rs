//! Ratatui front-end: sheet viewport, editing, export, move, file sync.

use crate::addr::{self, parse_cell_ref_at};
use crate::agg::{cell_display, compute_aggregate};
use crate::export;
use crate::formula::{cell_effective_display, is_formula};
use crate::grid::MainRange;
use crate::grid::{CellAddr, Grid, SortSpec, FOOTER_ROWS, HEADER_ROWS, MARGIN_COLS};
use crate::io::{commit_line, commit_workbook_op, load_full, load_revisions, IoError, LogWatcher};
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
use std::io::{self, stdout};
use std::path::{Path, PathBuf};
use thiserror::Error;

/// Width of the row-label gutter (`~1`, ` 1 `, `_1`).
const ROW_LABEL_CHARS: usize = 5;
/// Fixed cell display width in terminal columns.
const CELL_W: usize = 12;
/// Keep at most this many blank lines/cols around the active main data window.
const DISPLAY_EDGE_BLANK: usize = 1;
/// Trailing blank main rows allowed before Down transitions into the footer.
const NAV_BLANK_ROWS: usize = 2;
/// Trailing blank main cols allowed before Right transitions into the right margin.
const NAV_BLANK_COLS: usize = 1;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum SelectionKind {
    Cells,
    Rows,
    Cols,
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

    fn to_addr(self, grid: &Grid) -> CellAddr {
        let hr = HEADER_ROWS;
        let mr = grid.main_rows();
        let mc = grid.main_cols();
        if self.row < hr {
            CellAddr::Header {
                row: self.row as u8,
                col: self.col as u32,
            }
        } else if self.row < hr + mr {
            let mri = self.row - hr;
            let mcc = self.col;
            if mcc < MARGIN_COLS {
                CellAddr::Left {
                    col: mcc as u8,
                    row: mri as u32,
                }
            } else if mcc < MARGIN_COLS + mc {
                CellAddr::Main {
                    row: mri as u32,
                    col: (mcc - MARGIN_COLS) as u32,
                }
            } else {
                CellAddr::Right {
                    col: (mcc - MARGIN_COLS - mc) as u8,
                    row: mri as u32,
                }
            }
        } else {
            CellAddr::Footer {
                row: (self.row - hr - mr) as u8,
                col: self.col as u32,
            }
        }
    }
}

#[derive(Clone, Debug)]
enum Mode {
    Normal,
    Edit {
        buffer: String,
        formula_cursor: Option<SheetCursor>,
    },
    OpenPath {
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
    QuitPrompt,
}

const SPECIAL_VALUE_CHOICES: [&str; 10] = ["∞", "Σ", "Ω", "π", "μ", "Δ", "√", "φ", "λ", "θ"];

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum MenuSection {
    File,
    Export,
    Width,
    Insert,
    Help,
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
    OpenFile,
    SaveAs,
    Exit,
    ExportTsv,
    ExportCsv,
    ExportAscii,
    ExportAll,
    ExportOdt,
    SetMaxColWidth,
    SetColWidth,
    InsertRows,
    InsertCols,
    InsertSpecialChars,
    InsertDate,
    InsertTime,
    InsertHyperlink,
    SortView,
    SaveSort,
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
        shortcut: 'N',
        label: "New sheet",
        target: MenuTarget::Action(MenuAction::NewSheet),
    },
    MenuItem {
        shortcut: 'X',
        label: "Exit",
        target: MenuTarget::Action(MenuAction::Exit),
    },
];

const INSERT_ROOT_MENU_ITEMS: [MenuItem; 6] = [
    MenuItem {
        shortcut: 'R',
        label: "Rows",
        target: MenuTarget::Action(MenuAction::InsertRows),
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
        label: "ODT",
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
        MenuSection::File => &FILE_MENU_ITEMS,
        MenuSection::Insert => &INSERT_ROOT_MENU_ITEMS,
        MenuSection::Export => &EXPORT_MENU_ITEMS,
        MenuSection::Width => &WIDTH_MENU_ITEMS,
        MenuSection::Help => &HELP_MENU_ITEMS,
    }
}

fn menu_title(section: MenuSection) -> &'static str {
    match section {
        MenuSection::File => "File",
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
        MenuSection::File => MenuSection::Insert,
        MenuSection::Insert => MenuSection::Help,
        MenuSection::Help => MenuSection::File,
        _ => MenuSection::Help,
    }
}

fn menu_prev_root_section(section: MenuSection) -> MenuSection {
    match section {
        MenuSection::File => MenuSection::Help,
        MenuSection::Insert => MenuSection::File,
        MenuSection::Help => MenuSection::Insert,
        _ => MenuSection::File,
    }
}

fn menu_popup_area(area: Rect, section: MenuSection, parent: Option<(Rect, usize)>) -> Rect {
    let items = menu_items(section).len() as u16;
    let width = match section {
        MenuSection::File => 22,
        MenuSection::Export => 18,
        MenuSection::Width => 20,
        MenuSection::Insert => 20,
        MenuSection::Help => 18,
    }
    .min(area.width.saturating_sub(2).max(1));
    let height = items.saturating_add(2).min(area.height.max(3));
    let (x, y) = parent
        .map(|(p, item)| (p.x.saturating_add(p.width), p.y.saturating_add(item as u16)))
        .unwrap_or_else(|| match section {
            MenuSection::Help => (area.x.saturating_add(9), area.y.saturating_add(1)),
            _ => (area.x.saturating_add(1), area.y.saturating_add(1)),
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

    fn open_menu_path(&mut self, stack: Vec<MenuLevel>) {
        self.mode = Mode::Menu { stack };
    }

    fn start_edit_mode(
        &mut self,
        buffer: String,
        formula_cursor: Option<SheetCursor>,
        special_palette: bool,
    ) -> Mode {
        let cursor = if buffer.trim() == "=" {
            1
        } else {
            buffer.chars().count()
        };
        self.edit_cursor = Some(cursor);
        self.edit_special_palette = special_palette;
        Mode::Edit {
            buffer,
            formula_cursor,
        }
    }

    fn open_special_picker(&mut self) {
        self.special_picker = Some(0);
        self.mode = Mode::Normal;
    }

    fn commit_special_choice(&mut self, idx: usize) {
        let choice = SPECIAL_VALUE_CHOICES[idx];
        let buffer = choice.to_string();
        self.mode = self.start_edit_mode(buffer, None, true);
    }

    fn menu_action_mode(&mut self, action: MenuAction) -> Mode {
        self.edit_special_palette = false;
        match action {
            MenuAction::OpenFile => {
                let buffer = self.open_path_prompt_buffer();
                Mode::OpenPath {
                    buffer: self.start_input_mode(buffer),
                }
            }
            MenuAction::SaveAs => {
                let buffer = self
                    .path
                    .as_ref()
                    .map(|p| p.to_string_lossy().into_owned())
                    .unwrap_or_default();
                Mode::SavePath {
                    buffer: self.start_input_mode(buffer),
                }
            }
            MenuAction::Exit => Mode::QuitPrompt,
            MenuAction::ExportTsv => Mode::ExportTsv {
                buffer: self.start_input_mode(String::new()),
            },
            MenuAction::ExportCsv => Mode::ExportCsv {
                buffer: self.start_input_mode(String::new()),
            },
            MenuAction::ExportAscii => Mode::ExportAscii {
                buffer: self.start_input_mode(String::new()),
            },
            MenuAction::ExportAll => Mode::ExportAll {
                buffer: self.start_input_mode(String::new()),
            },
            MenuAction::ExportOdt => Mode::ExportOdt {
                buffer: self.start_input_mode(String::new()),
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
            ),
            MenuAction::InsertTime => self.start_edit_mode(
                chrono::Local::now().format("%H:%M:%S").to_string(),
                None,
                false,
            ),
            MenuAction::InsertHyperlink => {
                self.start_edit_mode(self.menu_insert_hyperlink_seed(), None, false)
            }
            MenuAction::SortView => Mode::SortView {
                buffer: self.start_input_mode(String::new()),
                persist: false,
            },
            MenuAction::SaveSort => Mode::SortView {
                buffer: self.start_input_mode(String::new()),
                persist: true,
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
        }
    }

    fn menu_target_mode(&mut self, path: &[MenuLevel], target: MenuTarget) -> Mode {
        match target {
            MenuTarget::Action(action) => self.menu_action_mode(action),
            MenuTarget::Submenu(section) => {
                let mut stack = path.to_vec();
                stack.push(MenuLevel { section, item: 0 });
                Mode::Menu { stack }
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
- Arrow keys or hjkl move the cursor.\n\
- Enter or e starts editing the current cell.\n\
- Header/footer/margin cells use the active address syntax.\n\
- Any printable key starts editing with that character.\n\
- = followed by arrows builds a formula reference.\n\n\
Selection and movement\n\
- v toggles a cell selection.\n\
- Ctrl+Shift+= inserts rows above the current row or selected rows.\n\
- r moves selected rows.\n\
- c exports CSV when nothing is selected, or moves selected columns when columns are selected.\n\
- Alt+arrows move selected rows or columns by one cell.\n\n\
Menus\n\
- Alt+F opens File.\n\
- Alt+I opens Insert.\n\
- Alt+H opens Help.\n\
- Ctrl+; inserts the date and Ctrl+Shift+; inserts the time.\n\
- Right opens the highlighted submenu.\n\
- Left goes back one menu level.\n\
 - Enter or the shortcut letter opens the selected item.\n\n\
File menu\n\
 - Open file loads a .corro, .csv, or .tsv file. Use `link <file> <revision>` to open a log at a revision.\n\
 - New sheet adds another sheet to the workbook.\n\
 - Ctrl+PageUp and Ctrl+PageDown switch between workbook tabs.\n\
- Export opens TSV, CSV, ASCII, full export, or ODT prompts.\n\
- Width opens default width and per-column width prompts.\n\
- Sort view changes the visible order of main rows.\n\
- Exit opens the quit prompt.\n\n\
Help menu\n\
- About shows the version and a short description.\n\
- Row ops and Col ops show quick move tips.\n\
- Full help opens this page.\n\n\
Address syntax\n\
 - Main cell: A1\n\
 - Header cell: ~1A\n\
- Footer cell: _1A\n\
- Left margin: [A1\n\
- Right margin: ]A1\n\
 - Cross-sheet refs use numeric IDs like #2!A1.\n\
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
    let fr = FOOTER_ROWS;
    let total = hr + mr + fr;
    let main_order = g.sorted_main_rows();
    let mut display_rows: Vec<usize> = Vec::with_capacity(total);
    display_rows.extend((0..hr).filter(|&r| g.logical_row_has_content(r) || cursor.row == r));
    display_rows.extend(main_order.iter().copied().map(|r| hr + r));
    display_rows.extend((0..fr).map(|r| hr + mr + r));

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

/// Column viewport with pinned totals and minimal-scroll movement.
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
    let dim = dim.max(1).min(total.max(1));
    let cur = cursor.col.min(total.saturating_sub(1));
    let cursor_in_left = cursor.col < lm;
    let cursor_in_right = cursor.col >= lm + mc;

    if total <= dim {
        return ((0..total).collect(), 0);
    }

    let (main_lo, main_hi) = main_col_window(state, cursor);
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

    let mut reserved: Vec<usize> = right_band;
    reserved.extend(left_band.iter().copied());
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
    let mut start = prev_start.min(max_start);
    if cur_pos < start {
        start = start.saturating_sub(1);
    } else if cur_pos >= start + available {
        start = (start + 1).min(max_start);
    }
    let end = (start + available).min(filtered.len());

    let mut out = filtered[start..end].to_vec();
    out.extend(reserved);
    out.sort_unstable();
    (out, start)
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
    grid.get(&CellAddr::Header {
        row: (HEADER_ROWS - 1) as u8,
        col: (MARGIN_COLS as u32) + main_col as u32,
    })
    .is_some_and(is_formula)
}

fn data_main_col_count(grid: &Grid) -> usize {
    let mc = grid.main_cols();
    for c in 0..mc {
        if right_col_agg_func(grid, MARGIN_COLS + c).is_some() {
            return c;
        }
    }
    mc
}

fn row_total_block_start(grid: &Grid, current_main_row: u32) -> u32 {
    let mut agg_row_start: u32 = 0;
    for candidate in (0..current_main_row).rev() {
        if left_margin_agg_func(grid, candidate).is_some() {
            break;
        }
        agg_row_start = candidate;
    }
    agg_row_start
}

fn left_margin_template_applies(grid: &Grid, main_row: usize) -> bool {
    grid.get(&CellAddr::Left {
        col: (MARGIN_COLS - 1) as u8,
        row: main_row as u32,
    })
    .is_some_and(is_formula)
}

// ── Display-time aggregate helpers ───────────────────────────────────────────

fn footer_row_agg_func(grid: &Grid, footer_row_idx: usize) -> Option<AggFunc> {
    let key_col = (MARGIN_COLS - 1) as u32;
    let val = grid.get(&CellAddr::Footer {
        row: footer_row_idx as u8,
        col: key_col,
    })?;
    match val.trim().to_uppercase().as_str() {
        "TOTAL" | "SUM" => Some(AggFunc::Sum),
        "MEAN" | "AVERAGE" | "AVG" => Some(AggFunc::Mean),
        "MEDIAN" => Some(AggFunc::Median),
        "MIN" | "MINIMUM" => Some(AggFunc::Min),
        "MAX" | "MAXIMUM" => Some(AggFunc::Max),
        "COUNT" => Some(AggFunc::Count),
        _ => None,
    }
}

fn right_col_agg_func(grid: &Grid, global_col: usize) -> Option<AggFunc> {
    let val = grid.get(&CellAddr::Header {
        row: (HEADER_ROWS - 1) as u8,
        col: global_col as u32,
    })?;
    match val.trim().to_uppercase().as_str() {
        "TOTAL" | "SUM" => Some(AggFunc::Sum),
        "MEAN" | "AVERAGE" | "AVG" => Some(AggFunc::Mean),
        "MEDIAN" => Some(AggFunc::Median),
        "MIN" | "MINIMUM" => Some(AggFunc::Min),
        "MAX" | "MAXIMUM" => Some(AggFunc::Max),
        "COUNT" => Some(AggFunc::Count),
        _ => None,
    }
}

fn parse_num(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    t.parse::<f64>().ok()
}

fn left_margin_agg_func(grid: &Grid, main_row: u32) -> Option<AggFunc> {
    let key_col = (MARGIN_COLS - 1) as u8;
    let val = grid.get(&CellAddr::Left {
        col: key_col,
        row: main_row,
    })?;
    match val.trim().to_uppercase().as_str() {
        "TOTAL" | "SUM" => Some(AggFunc::Sum),
        "MEAN" | "AVERAGE" | "AVG" => Some(AggFunc::Mean),
        "MEDIAN" => Some(AggFunc::Median),
        "MIN" | "MINIMUM" => Some(AggFunc::Min),
        "MAX" | "MAXIMUM" => Some(AggFunc::Max),
        "COUNT" => Some(AggFunc::Count),
        _ => None,
    }
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
    let right_func = right_col_agg_func(grid, global_col)?;
    let data_cols = data_main_col_count(grid).min(main_cols);
    let mut samples: Vec<f64> = Vec::new();
    for r in 0..main_rows {
        let row_val = compute_aggregate(
            grid,
            &AggregateDef {
                func: right_func,
                source: MainRange {
                    row_start: r as u32,
                    row_end: r as u32 + 1,
                    col_start: 0,
                    col_end: data_cols as u32,
                },
            },
        );
        if let Some(n) = parse_num(&row_val) {
            samples.push(n);
        }
    }
    Some(fold_numbers(footer_func, &samples))
}

// ── Cell-address shorthand ───────────────────────────────────────────────────

/// Parse `ADDR: VALUE` shorthand. Returns `(target_addr, value)` or `None`.
fn parse_cell_shorthand(buf: &str) -> Option<(CellAddr, String)> {
    let colon = buf.find(':')?;
    let addr_part = buf[..colon].trim();
    let value_part = buf[colon + 1..].trim_start().to_string();
    if addr_part.is_empty() {
        return None;
    }
    let (addr, n) = parse_cell_ref_at(addr_part)?;
    if n != addr_part.len() {
        return None;
    }
    Some((addr, value_part))
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

// ── App ───────────────────────────────────────────────────────────────────────

pub struct App {
    pub path: Option<PathBuf>,
    source_path: Option<PathBuf>,
    revision_limit: Option<usize>,
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
    help_scroll: usize,
    about_scroll: usize,
    pub op_history: Vec<Op>,
    selection_kind: SelectionKind,
    edit_special_palette: bool,
    edit_cursor: Option<usize>,
    input_cursor: Option<usize>,
    special_picker: Option<usize>,
    view_sheet_id: u32,
}

impl App {
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
            source_path,
            revision_limit,
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
            help_scroll: 0,
            about_scroll: 0,
            op_history: Vec::new(),
            selection_kind: SelectionKind::Cells,
            edit_special_palette: false,
            edit_cursor: None,
            input_cursor: None,
            special_picker: None,
            view_sheet_id: 1,
        }
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
                        for c in 0..self.state.grid.main_cols() {
                            self.state.grid.fit_column_to_content(MARGIN_COLS + c);
                        }
                        self.offset = data.len() as u64;
                        self.ops_applied =
                            data.lines().filter(|line| !line.trim().is_empty()).count();
                        self.path = Some(p.clone());
                        self.source_path = None;
                        self.revision_limit = None;
                        self.watcher = Some(LogWatcher::new(p.clone())?);
                        self.status = format!("Loaded workbook {}", p.display());
                        self.cursor.clamp(&self.state.grid);
                        return Ok(());
                    }
                    "tsv" => {
                        let data = std::fs::read_to_string(p).map_err(|e| IoError::Io(e))?;
                        crate::io::import_tsv(&data, &mut self.state);
                        self.path = Some(p.clone());
                        self.source_path = None;
                        self.revision_limit = None;
                        self.watcher = None;
                        for c in 0..self.state.grid.main_cols() {
                            self.state.grid.auto_fit_column(MARGIN_COLS + c);
                        }
                        self.status = format!("Imported TSV {}", p.display());
                    }
                    "csv" => {
                        let data = std::fs::read_to_string(p).map_err(|e| IoError::Io(e))?;
                        crate::io::import_csv(&data, &mut self.state);
                        self.path = Some(p.clone());
                        self.source_path = None;
                        self.revision_limit = None;
                        self.watcher = None;
                        for c in 0..self.state.grid.main_cols() {
                            self.state.grid.auto_fit_column(MARGIN_COLS + c);
                        }
                        self.status = format!("Imported CSV {}", p.display());
                    }
                    _ => {
                        let (off, n) = match linked_revision {
                            Some(limit) => load_revisions(p, limit, &mut self.state)?,
                            None => load_full(p, &mut self.state)?,
                        };
                        for c in 0..self.state.grid.main_cols() {
                            self.state.grid.auto_fit_column(MARGIN_COLS + c);
                        }
                        self.offset = off;
                        self.ops_applied = n;
                        if let Some(limit) = linked_revision {
                            self.source_path = Some(p.clone());
                            self.path = None;
                            self.watcher = None;
                            self.status = format!("Linked {} @ revision {}", p.display(), limit);
                        } else {
                            self.source_path = None;
                            self.path = Some(p.clone());
                            self.watcher = Some(LogWatcher::new(p.clone())?);
                            self.status = format!("Loaded {}", p.display());
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
    }

    fn current_selection_range(&self) -> Option<(Vec<usize>, Vec<usize>)> {
        let a = self.anchor?;
        let b = self.cursor;
        let r0 = a.row.min(b.row);
        let r1 = a.row.max(b.row);
        let c0 = a.col.min(b.col);
        let c1 = a.col.max(b.col);
        Some(((r0..=r1).collect(), (c0..=c1).collect()))
    }

    fn addr_at(&self, row: usize, col: usize) -> Option<CellAddr> {
        let preview_grid = if let Mode::Edit { buffer, .. } = &self.mode {
            let mut grid = self.state.grid.clone();
            let addr = self.cursor.to_addr(&self.state.grid);
            grid.set(&addr, buffer.clone());
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
                row: row as u8,
                col: col as u32,
            })
        } else if row < hr + mr {
            let mri = row - hr;
            if col < MARGIN_COLS {
                Some(CellAddr::Left {
                    col: col as u8,
                    row: mri as u32,
                })
            } else if col < MARGIN_COLS + mc {
                Some(CellAddr::Main {
                    row: mri as u32,
                    col: (col - MARGIN_COLS) as u32,
                })
            } else if col < MARGIN_COLS + mc + MARGIN_COLS {
                Some(CellAddr::Right {
                    col: (col - MARGIN_COLS - mc) as u8,
                    row: mri as u32,
                })
            } else {
                None
            }
        } else if row < hr + mr + FOOTER_ROWS {
            Some(CellAddr::Footer {
                row: (row - hr - mr) as u8,
                col: col as u32,
            })
        } else {
            None
        }
    }

    fn delete_selection(&mut self) -> bool {
        let Some((rows, cols)) = self.current_selection_range() else {
            return false;
        };
        let mut did_any = false;
        for r in rows {
            for c in cols.iter().copied() {
                let Some(addr) = self.addr_at(r, c) else {
                    continue;
                };
                if self.state.grid.get(&addr).is_some_and(|v| !v.is_empty()) {
                    let op = Op::SetCell {
                        addr: addr.clone(),
                        value: String::new(),
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
                    if let CellAddr::Main { col, .. } = addr {
                        self.state.grid.auto_fit_column(MARGIN_COLS + col as usize);
                    }
                    did_any = true;
                }
            }
        }
        if did_any {
            self.status = "Selection deleted".into();
            self.anchor = None;
        }
        did_any
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
        let fr = FOOTER_ROWS;
        let mut rows = Vec::with_capacity(hr + mr + fr);
        rows.extend(0..hr);
        rows.extend(g.sorted_main_rows().into_iter().map(|r| hr + r));
        rows.extend((0..fr).map(|r| hr + mr + r));
        rows
    }

    fn move_cursor_row_through_view(&mut self, down: bool) -> bool {
        if self.state.grid.view_sort_cols.is_empty() {
            return false;
        }

        let hr = HEADER_ROWS;
        let mr = self.state.grid.main_rows();
        let last_main = hr + mr.saturating_sub(1);
        let first_footer = hr + mr;
        let rows = self.view_row_order();
        let Some(pos) = rows.iter().position(|&r| r == self.cursor.row) else {
            return false;
        };
        let next_pos = if down {
            if self.cursor.row == last_main
                && trailing_blank_main_rows(&self.state) < NAV_BLANK_ROWS
            {
                self.state.grid.grow_main_row_at_bottom();
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
        true
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

    fn commit_edit_buffer(&mut self, buffer: &str) -> Result<(), RunError> {
        self.edit_special_palette = false;
        let (addr, value) = if let Some((a, v)) = parse_cell_shorthand(buffer) {
            (a, v)
        } else {
            (self.cursor.to_addr(&self.state.grid), buffer.to_string())
        };
        if self.state.grid.get(&addr).unwrap_or("") == value {
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
        Ok(())
    }

    fn autofit_column_from_current_cell(&mut self, addr: CellAddr) {
        match addr {
            CellAddr::Main { col, .. } => {
                self.state.grid.auto_fit_column(MARGIN_COLS + col as usize)
            }
            CellAddr::Left { col, .. } => self.state.grid.auto_fit_column(col as usize),
            CellAddr::Right { col, .. } => self
                .state
                .grid
                .auto_fit_column(MARGIN_COLS + self.state.grid.main_cols() + col as usize),
            CellAddr::Header { col, .. } | CellAddr::Footer { col, .. } => {
                self.state.grid.auto_fit_column(col as usize)
            }
        }
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

        self.anchor = Some(SheetCursor {
            row: HEADER_ROWS + from as usize,
            col: MARGIN_COLS,
        });
        self.cursor = SheetCursor {
            row: HEADER_ROWS + (from + count - 1) as usize,
            col: MARGIN_COLS + self.state.grid.main_cols().saturating_sub(1),
        };
        self.selection_kind = SelectionKind::Rows;
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
        self.anchor = Some(SheetCursor {
            row: HEADER_ROWS + row as usize,
            col: MARGIN_COLS,
        });
        self.cursor = SheetCursor {
            row: HEADER_ROWS + (row + count - 1) as usize,
            col: MARGIN_COLS + self.state.grid.main_cols().saturating_sub(1),
        };
        self.selection_kind = SelectionKind::Rows;
        self.status = if count == 1 {
            format!("Inserted 1 row above row {row}")
        } else {
            format!("Inserted {count} rows above row {row}")
        };
        Ok(true)
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
        let current = self.state.grid.get(&addr).unwrap_or("").trim();
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
        let current = self.state.grid.get(&addr).unwrap_or("").trim();
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
        match addr {
            CellAddr::Header { row, col } => format!(
                "~{}{}",
                HEADER_ROWS - *row as usize,
                formula_col_fragment(*col as usize, self.state.grid.main_cols())
            ),
            CellAddr::Footer { row, col } => format!(
                "_{}{}",
                *row as usize + 1,
                formula_col_fragment(*col as usize, self.state.grid.main_cols())
            ),
            CellAddr::Main { row, col } => {
                format!("{}{}", addr::excel_column_name(*col as usize), row + 1)
            }
            CellAddr::Left { col, row } => format!(
                "[{}{}",
                addr::mirror_margin_column_name(*col as usize, true),
                row + 1
            ),
            CellAddr::Right { col, row } => format!(
                "]{}{}",
                addr::mirror_margin_column_name(*col as usize, false),
                row + 1
            ),
        }
    }

    fn do_export(&mut self, csv: bool) -> String {
        crate::formula::refresh_spills(&mut self.state.grid);
        let mut buf = Vec::new();
        if csv {
            export::export_csv(&self.state.grid, &mut buf);
        } else {
            if self.anchor.is_some() {
                let rows = self
                    .current_selection_range()
                    .map(|(rows, _)| rows)
                    .unwrap_or_default();
                let cols = self
                    .current_selection_range()
                    .map(|(_, cols)| cols)
                    .unwrap_or_default();
                export::export_selection(&self.state.grid, &mut buf, &rows, &cols);
            } else {
                export::export_tsv(&self.state.grid, &mut buf);
            }
        }
        String::from_utf8_lossy(&buf).into_owned()
    }

    fn do_export_ascii(&mut self) -> String {
        crate::formula::refresh_spills(&mut self.state.grid);
        let mut buf = Vec::new();
        export::export_ascii_table(&self.state.grid, &mut buf);
        String::from_utf8_lossy(&buf).into_owned()
    }

    fn do_export_all(&mut self) -> String {
        crate::formula::refresh_spills(&mut self.state.grid);
        let mut buf = Vec::new();
        export::export_all(&self.state.grid, &mut buf);
        String::from_utf8_lossy(&buf).into_owned()
    }

    fn do_export_odt(&mut self) -> Vec<u8> {
        crate::formula::refresh_spills(&mut self.state.grid);
        export::export_odt_bytes(&self.state.grid).unwrap_or_default()
    }

    fn save_to_path(&mut self, path: &Path) -> Result<(), RunError> {
        self.commit_active_sheet_cache();
        let mut buf = String::new();
        for sheet in &self.workbook.sheets {
            buf.push_str(
                &crate::ops::WorkbookOp::NewSheet {
                    id: sheet.id,
                    title: sheet.title.clone(),
                }
                .to_log_line(),
            );
            buf.push('\n');
            for row in 0..sheet.state.grid.main_rows() {
                for col in 0..sheet.state.grid.main_cols() {
                    let addr = CellAddr::Main {
                        row: row as u32,
                        col: col as u32,
                    };
                    if let Some(value) = sheet.state.grid.get(&addr) {
                        if !value.is_empty() {
                            buf.push_str(
                                &crate::ops::WorkbookOp::SheetOp {
                                    sheet_id: sheet.id,
                                    op: Op::SetCell {
                                        addr: addr.clone(),
                                        value: value.to_string(),
                                    },
                                }
                                .to_log_line(),
                            );
                            buf.push('\n');
                        }
                    }
                }
            }
            buf.push_str(
                &crate::ops::WorkbookOp::SheetOp {
                    sheet_id: sheet.id,
                    op: Op::SetMainSize {
                        main_rows: sheet.state.grid.main_rows() as u32,
                        main_cols: sheet.state.grid.main_cols() as u32,
                    },
                }
                .to_log_line(),
            );
            buf.push('\n');
            if sheet.state.grid.max_col_width != 20 {
                buf.push_str(
                    &crate::ops::WorkbookOp::SheetOp {
                        sheet_id: sheet.id,
                        op: Op::SetMaxColWidth {
                            width: sheet.state.grid.max_col_width,
                        },
                    }
                    .to_log_line(),
                );
                buf.push('\n');
            }
            for (col, width) in &sheet.state.grid.col_width_overrides {
                buf.push_str(
                    &crate::ops::WorkbookOp::SheetOp {
                        sheet_id: sheet.id,
                        op: Op::SetColWidth {
                            col: *col,
                            width: Some(*width),
                        },
                    }
                    .to_log_line(),
                );
                buf.push('\n');
            }
            if !sheet.state.grid.view_sort_cols.is_empty() {
                buf.push_str(
                    &crate::ops::WorkbookOp::SheetOp {
                        sheet_id: sheet.id,
                        op: Op::SetViewSortCols {
                            cols: sheet.state.grid.view_sort_cols.clone(),
                        },
                    }
                    .to_log_line(),
                );
                buf.push('\n');
            }
        }
        std::fs::write(path, buf)?;
        self.path = Some(path.to_path_buf());
        self.source_path = None;
        self.revision_limit = None;
        self.status = format!("Saved {}", path.display());
        if self.watcher.is_none() {
            self.watcher = Some(LogWatcher::new(path.to_path_buf()).map_err(IoError::from)?);
        }
        Ok(())
    }

    #[allow(dead_code)]
    fn do_export_selection(&mut self) -> String {
        let mut buf = Vec::new();
        let rows: Vec<usize> = (0..self.state.grid.main_rows()).collect();
        let cols: Vec<usize> = (MARGIN_COLS..MARGIN_COLS + self.state.grid.main_cols()).collect();
        export::export_selection(&self.state.grid, &mut buf, &rows, &cols);
        String::from_utf8_lossy(&buf).into_owned()
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
        let data_cols = inner_w
            .saturating_sub(ROW_LABEL_CHARS)
            .checked_div(CELL_W)
            .unwrap_or(1)
            .max(1);

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

        let (row_ixs, next_row_scroll) =
            visible_row_indices(&self.state, self.cursor, data_rows, self.row_scroll);
        let (col_ixs, next_col_scroll) =
            visible_col_indices(&self.state, self.cursor, data_cols, self.col_scroll);
        self.row_scroll = next_row_scroll;
        self.col_scroll = next_col_scroll;

        // ── Menu bar ──────────────────────────────────────────────────────────
        let menubar = self.menu_bar_line();
        f.render_widget(
            Paragraph::new(menubar).style(Style::default().fg(Color::Black).bg(Color::Cyan)),
            menubar_area,
        );

        // ── Formula bar ───────────────────────────────────────────────────────
        let addr = self.cursor.to_addr(grid);
        let addr_str = addr_label(&addr, grid.main_cols());
        let prompt_style = Style::default().fg(Color::White).bg(Color::DarkGray);
        let prompt_style_bold = prompt_style.add_modifier(Modifier::BOLD);
        let caret_style = Style::default()
            .fg(Color::Black)
            .bg(Color::Yellow)
            .add_modifier(Modifier::BOLD);
        let formula_widget = match &self.mode {
            Mode::Edit { buffer, .. } => Paragraph::new(input_line(
                format!(" {addr_str}  "),
                buffer,
                self.edit_cursor.unwrap_or_else(|| buffer.chars().count()),
                prompt_style_bold,
                caret_style,
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
                " export ODT: ".to_string(),
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
            Mode::QuitPrompt => Paragraph::new(" Quit Corro? (Q)uit, (B)ack ")
                .style(Style::default().fg(Color::White).bg(Color::Red)),
            Mode::Help => Paragraph::new(" Help - Up/Down scroll, Esc closes ")
                .style(Style::default().fg(Color::White).bg(Color::Blue)),
            Mode::About => Paragraph::new(" About - Up/Down scroll, Esc closes ")
                .style(Style::default().fg(Color::White).bg(Color::Blue)),
            Mode::Menu { .. } => {
                let val = cell_effective_display(grid, &addr);
                let base = format!(" {addr_str}  {val}");
                let text = if self.status.is_empty() {
                    base
                } else {
                    format!("{base}   ·  {}", self.status)
                };
                Paragraph::new(text).style(Style::default().fg(Color::Cyan))
            }
            _ => {
                let val = cell_effective_display(grid, &addr);
                let base = format!(" {addr_str}  {val}");
                let text = if self.status.is_empty() {
                    base
                } else {
                    format!("{base}   ·  {}", self.status)
                };
                Paragraph::new(text).style(Style::default().fg(Color::Cyan))
            }
        };
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

        if matches!(&self.mode, Mode::Help | Mode::About) {
            let body = match &self.mode {
                Mode::Help => self.help_page_body(),
                Mode::About => self.about_page_body(),
                _ => String::new(),
            };
            let inner = Block::default().borders(Borders::ALL).inner(grid_area);
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
            let block = Block::default()
                .borders(Borders::ALL)
                .title(match self.mode {
                    Mode::Help => " Help ",
                    Mode::About => " About ",
                    _ => "",
                });
            let paragraph = Paragraph::new(visible)
                .block(block)
                .wrap(Wrap { trim: false });
            f.render_widget(Clear, grid_area);
            f.render_widget(paragraph, grid_area);
            return;
        }

        // ── Grid ──────────────────────────────────────────────────────────────
        let mut lines: Vec<Line> = Vec::new();

        {
            let lm = MARGIN_COLS;
            let mc = grid.main_cols();
            let show_right_divider = col_ixs.contains(&(lm + mc));
            let mut spans: Vec<Span> = vec![Span::styled(
                format!("{:>width$}", "", width = ROW_LABEL_CHARS),
                Style::default().add_modifier(Modifier::BOLD),
            )];
            for &c in &col_ixs {
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
                spans.push(Span::styled(format!("{:>w$}", name, w = w), style));
                if c == lm - 1 && lm > 0 && col_ixs.contains(&lm) {
                    spans.push(Span::styled("│", Style::default().fg(Color::DarkGray)));
                }
                if c == lm + mc - 1 && show_right_divider {
                    spans.push(Span::styled("│", Style::default().fg(Color::DarkGray)));
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
        for &r in row_ixs.iter().take(max_data_lines) {
            let active_row = r == self.cursor.row;
            let row_label_style = if active_row {
                Style::default()
                    .fg(Color::Black)
                    .bg(Color::Yellow)
                    .add_modifier(Modifier::BOLD)
            } else {
                Style::default().fg(Color::Yellow)
            };
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

            for &c in &col_ixs {
                let cur = SheetCursor { row: r, col: c };
                let cell_addr = cur.to_addr(grid);

                let text = if let Some(func) = footer_agg {
                    if right_col_agg_func(grid, c).is_some() {
                        footer_special_col_aggregate(grid, func, c, mr, data_main_col_count(grid))
                            .unwrap_or_else(|| cell_effective_display(grid, &cell_addr))
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
                        cell_effective_display(grid, &cell_addr)
                    }
                } else if left_margin_agg.is_some() {
                    cell_effective_display(grid, &cell_addr)
                } else if r >= hr && r < hr + mr {
                    if let Some(func) = right_col_agg_func(grid, c) {
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
                let disp = if text.chars().count() > cw {
                    format!("{}…", text.chars().take(cw).collect::<String>())
                } else if text.trim().parse::<f64>().is_ok() {
                    format!("{:>w$}", text, w = cw)
                } else {
                    format!("{:<w$}", text, w = cw)
                };
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
                let is_footer_border = r == hr + mr && row_ixs.contains(&(hr + mr));

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
                let is_last_data_row = r == hr + mr - 1;
                if is_last_data_row {
                    st = st.add_modifier(Modifier::UNDERLINED);
                }
                spans.push(Span::styled(disp, st));
                if c == lm - 1 && lm > 0 && col_ixs.contains(&lm) {
                    spans.push(Span::styled("│", Style::default().fg(Color::DarkGray)));
                }
                if c == lm + mc - 1 && show_right_divider {
                    spans.push(Span::styled("│", Style::default().fg(Color::DarkGray)));
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
                    "  shortcuts: e·edit v·select t·tsv c·csv o·open a·save as q·quit ?·help; printable letters edit cells unless reserved".into()
                }
            }
            Mode::Edit { .. } => {
                "  type to edit (or addr: val)   Enter·confirm   Esc·discard".into()
            }
            Mode::OpenPath { .. } => {
                "  type path or link <file> <revision>   Enter·open   Esc·cancel".into()
            }
            Mode::SavePath { .. } => "  type file path   Enter·save as   Esc·cancel".into(),
            Mode::ExportTsv { .. }
            | Mode::ExportCsv { .. }
            | Mode::ExportAscii { .. }
            | Mode::ExportAll { .. }
            | Mode::ExportOdt { .. }
            | Mode::SetMaxColWidth { .. }
            | Mode::SetColWidth { .. } => {
                "  type filename (blank=clipboard)   Enter·export   Esc·cancel".into()
            }
            Mode::SortView { .. } => {
                "  type sort columns like A,B,C   Enter·apply   Esc·cancel".into()
            }
            Mode::QuitPrompt => "  Q·quit   B·back   Esc·cancel".into(),
            Mode::Help => "  up/down·scroll   Esc·close   ?·help   A·about".into(),
            Mode::About => "  up/down·scroll   Esc·close   ?·help   A·about".into(),
            Mode::Menu { .. } => {
                "  right·open submenu   left·back   up/down·move   Enter/letter·open   Esc·close"
                    .into()
            }
        }
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
        let insert = if section == MenuSection::Insert {
            "[Insert]"
        } else {
            " Insert "
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
        format!(" {file}  {insert}  {help}{active}")
    }

    fn handle_key(&mut self, key: KeyEvent) -> Result<bool, RunError> {
        if key.kind == KeyEventKind::Release {
            return Ok(false);
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

        let mut mode = std::mem::replace(&mut self.mode, Mode::Normal);

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
                            mode = self.menu_target_mode(stack.as_slice(), menu_item.target);
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
                            mode = self.menu_target_mode(stack.as_slice(), menu_item.target);
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
                        self.open_menu_path(vec![MenuLevel {
                            section: MenuSection::Export,
                            item: 3,
                        }]);
                        return Ok(false);
                    }
                    'i' | 'I' => {
                        self.open_menu(MenuSection::Insert);
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
                    's' => {
                        self.open_menu_path(vec![MenuLevel {
                            section: MenuSection::File,
                            item: 3,
                        }]);
                        return Ok(false);
                    }
                    'S' => {
                        self.open_menu_path(vec![MenuLevel {
                            section: MenuSection::File,
                            item: 4,
                        }]);
                        return Ok(false);
                    }
                    'n' | 'N' => {
                        self.add_sheet(format!("Sheet{}", self.workbook.next_sheet_id));
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
                        self.mode = Mode::Edit {
                            buffer,
                            formula_cursor: None,
                        };
                        return Ok(false);
                    }
                    ':' => {
                        let buffer = chrono::Local::now().format("%H:%M:%S").to_string();
                        self.mode = Mode::Edit {
                            buffer,
                            formula_cursor: None,
                        };
                        return Ok(false);
                    }
                    ';' => {
                        let buffer = chrono::Local::now().format("%Y-%m-%d").to_string();
                        self.mode = Mode::Edit {
                            buffer,
                            formula_cursor: None,
                        };
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
            Mode::Menu { .. } => {}
            Mode::ExportTsv { buffer } => match key.code {
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
            },
            Mode::ExportCsv { buffer } => match key.code {
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
            },
            Mode::ExportAscii { buffer } => match key.code {
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
                                self.status = format!("ASCII table exported to {}", fname.trim())
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
            },
            Mode::ExportOdt { buffer } => match key.code {
                KeyCode::Enter => {
                    let fname = buffer.clone();
                    if fname.trim().is_empty() {
                        self.status = "ODT requires a filename".into();
                    } else {
                        match std::fs::write(fname.trim(), self.do_export_odt()) {
                            Ok(()) => self.status = format!("ODT saved to {}", fname.trim()),
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
            },
            Mode::ExportAll { buffer } => match key.code {
                KeyCode::Enter => {
                    let fname = buffer.clone();
                    if fname.trim().is_empty() {
                        let data = if self.anchor.is_some() {
                            self.do_export(false)
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
                            self.do_export(false)
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
            },
            Mode::SetMaxColWidth { buffer } => match key.code {
                KeyCode::Enter => {
                    if let Ok(width) = buffer.trim().parse::<usize>() {
                        if let Some(ref p) = self.path.clone() {
                            commit_line(
                                p,
                                &mut self.offset,
                                &mut self.state,
                                &format!("MAX_COL_WIDTH {width}"),
                            )?;
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
                            let line =
                                format!("COL_WIDTH {} {}", addr::excel_column_name(col), width);
                            if let Some(ref p) = self.path.clone() {
                                commit_line(p, &mut self.offset, &mut self.state, &line)?;
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
                        let line = format!("COL_WIDTH {}", addr::excel_column_name(col));
                        if let Some(ref p) = self.path.clone() {
                            commit_line(p, &mut self.offset, &mut self.state, &line)?;
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
                        let line = format!(
                            "SORT {}",
                            cols.iter()
                                .map(|spec| {
                                    let name = addr::excel_column_name(
                                        spec.col.saturating_sub(MARGIN_COLS),
                                    );
                                    if spec.desc {
                                        format!("!{name}")
                                    } else {
                                        name
                                    }
                                })
                                .collect::<Vec<_>>()
                                .join(" ")
                        );
                        if let Some(ref p) = self.path.clone() {
                            commit_line(p, &mut self.offset, &mut self.state, &line)?;
                            self.ops_applied = self.ops_applied.saturating_add(1);
                            self.start_log_watcher_if_needed()?;
                        } else {
                            self.state.grid.set_view_sort_cols(cols);
                        }
                    } else {
                        self.state.grid.set_view_sort_cols(cols);
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
            Mode::OpenPath { buffer } => match key.code {
                KeyCode::Enter => match parse_open_path_request(buffer) {
                    Err(OpenPathError::Empty) => {
                        self.status = "Path required".into();
                    }
                    Err(OpenPathError::InvalidRevisionSyntax) => {
                        self.status = "Link syntax: link <file> <revision>".into();
                    }
                    Ok(OpenPathRequest::Plain(path)) => {
                        self.path = Some(path.clone());
                        self.source_path = None;
                        self.offset = 0;
                        self.state = SheetState::new(1, 1);
                        self.ops_applied = 0;
                        self.revision_limit = None;
                        if path.exists() {
                            let ext = path
                                .extension()
                                .and_then(|e| e.to_str())
                                .unwrap_or("")
                                .to_lowercase();
                            match ext.as_str() {
                                "tsv" => {
                                    if let Ok(data) = std::fs::read_to_string(&path) {
                                        crate::io::import_tsv(&data, &mut self.state);
                                    }
                                }
                                "csv" => {
                                    if let Ok(data) = std::fs::read_to_string(&path) {
                                        crate::io::import_csv(&data, &mut self.state);
                                    }
                                }
                                _ => {
                                    let loaded = load_full(&path, &mut self.state);
                                    if let Ok((off, n)) = loaded {
                                        self.offset = off;
                                        self.ops_applied = n;
                                    }
                                }
                            }
                        }
                        self.watcher = if path.exists() {
                            Some(LogWatcher::new(path.clone()).map_err(IoError::from)?)
                        } else {
                            None
                        };
                        self.cursor = SheetCursor {
                            row: HEADER_ROWS,
                            col: MARGIN_COLS,
                        };
                        self.row_scroll = 0;
                        self.col_scroll = 0;
                        self.status = if path.exists() {
                            format!("Opened {}", path.display())
                        } else {
                            format!("New file {}", path.display())
                        };
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
                            if matches!(ext.as_str(), "csv" | "tsv") {
                                self.status = "Link only works for .corro logs".into();
                            } else if let Ok((off, n)) =
                                load_revisions(&path, revision, &mut self.state)
                            {
                                self.path = None;
                                self.source_path = Some(path.clone());
                                self.revision_limit = Some(revision);
                                self.offset = off;
                                self.ops_applied = n;
                                self.watcher = None;
                                self.cursor = SheetCursor {
                                    row: HEADER_ROWS,
                                    col: MARGIN_COLS,
                                };
                                self.row_scroll = 0;
                                self.col_scroll = 0;
                                self.status =
                                    format!("Linked {} @ revision {}", path.display(), revision);
                                mode = Mode::Normal;
                            } else {
                                self.status = format!("Link load failed: {}", path.display());
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
            Mode::Edit {
                buffer,
                formula_cursor,
            } => match key.code {
                KeyCode::Enter => {
                    self.commit_edit_buffer(buffer)?;
                    self.edit_cursor = None;
                    mode = Mode::Normal;
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
                    if self.cursor.row > 0 {
                        self.cursor.row = self.cursor.row.saturating_sub(1);
                    }
                    self.cursor.clamp(&self.state.grid);
                    self.state
                        .grid
                        .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
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
                    );
                }
                KeyCode::Down => {
                    self.edit_cursor = None;
                    let raw = buffer.clone();
                    self.commit_edit_buffer(&raw)?;
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
                    );
                }
                KeyCode::Esc => {
                    self.edit_cursor = None;
                    self.edit_special_palette = false;
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
                if key.modifiers.contains(KeyModifiers::CONTROL) && key.code == KeyCode::Char('q') {
                    self.mode = mode;
                    return Ok(true);
                }
                if key.modifiers.contains(KeyModifiers::CONTROL) && key.code == KeyCode::Char('z') {
                    if let Some(undo_op) = self.op_history.pop() {
                        if let Some(ref p) = self.path.clone() {
                            let mut active_sheet = self.view_sheet_id;
                            if let Err(e) = commit_workbook_op(
                                p,
                                &mut self.offset,
                                &mut self.workbook,
                                &mut active_sheet,
                                &crate::ops::WorkbookOp::SheetOp {
                                    sheet_id: self.view_sheet_id,
                                    op: undo_op.clone(),
                                },
                            ) {
                                self.status = format!("Undo failed: {}", e);
                            } else {
                                self.ops_applied = self.ops_applied.saturating_add(1);
                                self.sync_active_sheet_cache();
                                self.status = "Undo applied".to_string();
                            }
                        } else {
                            undo_op.apply(&mut self.state);
                            self.status = "Undo applied (memory only)".to_string();
                        }
                    } else {
                        self.status = "Nothing to undo".to_string();
                    }
                    self.mode = mode;
                    return Ok(false);
                }

                match key.code {
                    KeyCode::Esc => {
                        self.anchor = None;
                        if self.anchor.is_none() {
                            mode = Mode::QuitPrompt;
                        }
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
                            self.cursor.row = self.cursor.row.saturating_sub(1);
                            self.cursor.clamp(&self.state.grid);
                        }
                    }
                    KeyCode::Down if key.modifiers.contains(KeyModifiers::SHIFT) => {
                        if self.anchor.is_none() {
                            self.anchor = Some(self.cursor);
                        }
                        self.cursor.row = self.cursor.row.saturating_add(1);
                        self.cursor.clamp(&self.state.grid);
                        self.state
                            .grid
                            .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
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
                        mode = Mode::ExportTsv {
                            buffer: self.start_input_mode(String::new()),
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
                            mode = Mode::ExportCsv {
                                buffer: self.start_input_mode(String::new()),
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
                    KeyCode::Delete | KeyCode::Backspace => {
                        if !self.delete_selection() {
                            if let Some(addr) = self.addr_at(self.cursor.row, self.cursor.col) {
                                if self.state.grid.get(&addr).unwrap_or("").is_empty() {
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
                        if !self.move_cursor_row_through_view(false) {
                            self.cursor.row = self.cursor.row.saturating_sub(1);
                            self.cursor.clamp(&self.state.grid);
                            self.state
                                .grid
                                .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
                        }
                    }
                    KeyCode::Down | KeyCode::Char('j') => {
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
                    }
                    KeyCode::Char(c) if !c.is_control() => {
                        self.edit_special_palette = false;
                        let buffer = c.to_string();
                        mode = self.start_edit_mode(
                            buffer.clone(),
                            if buffer.trim() == "=" {
                                Some(self.cursor)
                            } else {
                                None
                            },
                            false,
                        );
                    }
                    _ => {}
                }
            }
        }

        self.mode = mode;
        Ok(false)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crossterm::event::{KeyCode, KeyEvent, KeyModifiers};

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
            app.state.grid.get(&CellAddr::Main { row: 0, col: 0 }),
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
            app.state.grid.get(&CellAddr::Main { row: 2, col: 0 }),
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
        assert_eq!(app.cursor.row, HEADER_ROWS + 1);
        assert_eq!(app.anchor.unwrap().row, HEADER_ROWS);
        assert_eq!(app.state.grid.get(&CellAddr::Main { row: 0, col: 0 }), None);
        assert_eq!(app.state.grid.get(&CellAddr::Main { row: 1, col: 0 }), None);
        assert_eq!(
            app.state.grid.get(&CellAddr::Main { row: 2, col: 0 }),
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
            app.state.grid.get(&CellAddr::Main { row: 1, col: 0 }),
            Some("top")
        );
        assert_eq!(
            app.state.grid.get(&CellAddr::Main { row: 2, col: 0 }),
            Some("bottom")
        );
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
                item: 1,
            }],
        };

        app.handle_key(KeyEvent::new(KeyCode::Enter, KeyModifiers::empty()))
            .unwrap();

        assert_eq!(app.state.grid.main_cols(), 3);
        assert_eq!(
            app.state.grid.get(&CellAddr::Main { row: 0, col: 0 }),
            Some("left")
        );
        assert_eq!(app.state.grid.get(&CellAddr::Main { row: 0, col: 1 }), None);
        assert_eq!(
            app.state.grid.get(&CellAddr::Main { row: 0, col: 2 }),
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
                item: 2,
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
                item: 2,
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
                item: 5,
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
                row: 25,
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
    fn startup_keeps_total_column_visible() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(None);
        app.state.grid.set_main_size(4, 3);
        app.state.grid.set(
            &CellAddr::Header {
                row: 25,
                col: MARGIN_COLS as u32 + 2,
            },
            "TOTAL".into(),
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
        use crate::io::load_revisions;
        use std::path::Path;

        let mut state = SheetState::new(1, 1);
        load_revisions(Path::new("test5.corro"), 141, &mut state).unwrap();

        assert_eq!(
            footer_special_col_aggregate(
                &state.grid,
                AggFunc::Sum,
                MARGIN_COLS + 2,
                state.grid.main_rows(),
                state.grid.main_cols(),
            ),
            Some("105".into())
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
                row: 25,
                col: MARGIN_COLS as u32 + 3,
            },
            "TOTAL".into(),
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
    fn escape_cancels_edit_without_committing() {
        let mut app = App::new(None);
        app.state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "orig".into());
        app.mode = Mode::Edit {
            buffer: "changed".into(),
            formula_cursor: None,
        };

        app.handle_key(KeyEvent::new(KeyCode::Esc, KeyModifiers::empty()))
            .unwrap();

        assert!(matches!(app.mode, Mode::Normal));
        assert_eq!(
            app.state.grid.get(&CellAddr::Main { row: 0, col: 0 }),
            Some("orig")
        );
    }

    #[test]
    fn open_path_parses_link_revision() {
        let parsed = parse_open_path_request("link test5.corro 2").unwrap();
        match parsed {
            OpenPathRequest::Revision { path, revision } => {
                assert_eq!(path, PathBuf::from("test5.corro"));
                assert_eq!(revision, 2);
            }
            other => panic!("unexpected parse: {other:?}"),
        }
    }

    #[test]
    fn linked_revision_uses_source_path_and_detaches_on_save() {
        let mut app = App::new_with_revision_limit(Some(PathBuf::from("test5.corro")), Some(2));
        assert!(app.path.is_none());
        assert_eq!(app.source_path, Some(PathBuf::from("test5.corro")));
        assert_eq!(app.revision_limit, Some(2));

        let tmp = tempfile::NamedTempFile::new().unwrap();
        app.save_to_path(tmp.path()).unwrap();

        assert_eq!(app.path, Some(tmp.path().to_path_buf()));
        assert_eq!(app.source_path, None);
        assert_eq!(app.revision_limit, None);
    }

    #[test]
    fn save_clears_revision_limit() {
        let mut app = App::new_with_revision_limit(Some(PathBuf::from("test5.corro")), Some(2));
        app.revision_limit = Some(2);
        let path = tempfile::NamedTempFile::new().unwrap();

        app.save_to_path(path.path()).unwrap();

        assert_eq!(app.revision_limit, None);
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
    fn workbook_edit_updates_visible_sheet_immediately() {
        let path = tempfile::NamedTempFile::new().unwrap();
        let mut app = App::new(Some(path.path().to_path_buf()));
        app.add_sheet("Sheet2".into());

        app.mode = Mode::Edit {
            buffer: "Sheet2 value".into(),
            formula_cursor: None,
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
                .and_then(|sheet| sheet.state.grid.get(&CellAddr::Main { row: 0, col: 0 })),
            Some("Sheet2 value")
        );
        assert_eq!(
            app.state.grid.get(&CellAddr::Main { row: 0, col: 0 }),
            Some("Sheet2 value")
        );
    }

    #[test]
    fn ctrl_page_switch_works_in_edit_mode() {
        let mut app = App::new(None);
        app.add_sheet("Sheet2".into());
        app.mode = Mode::Edit {
            buffer: "x".into(),
            formula_cursor: None,
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
            app.state.grid.get(&CellAddr::Main { row: 0, col: 0 }),
            Some("=$Sheet1:A1")
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
        };

        app.handle_key(KeyEvent::new(
            KeyCode::Char('+'),
            KeyModifiers::CONTROL | KeyModifiers::SHIFT,
        ))
        .unwrap();

        assert_eq!(app.state.grid.main_rows(), 3);
        assert_eq!(app.state.grid.get(&CellAddr::Main { row: 0, col: 0 }), None);
        assert_eq!(
            app.state.grid.get(&CellAddr::Main { row: 1, col: 0 }),
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
    fn left_wraps_from_help_to_file() {
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
                assert_eq!(stack[0].section, MenuSection::Insert);
            }
            other => panic!("unexpected mode: {other:?}"),
        }

        app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::empty()))
            .unwrap();

        match &app.mode {
            Mode::Menu { stack } => {
                assert_eq!(stack.len(), 1);
                assert_eq!(stack[0].section, MenuSection::File);
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
                assert_eq!(stack[0].section, MenuSection::Insert);
            }
            other => panic!("unexpected mode: {other:?}"),
        }

        app.handle_key(KeyEvent::new(KeyCode::Left, KeyModifiers::empty()))
            .unwrap();

        match &app.mode {
            Mode::Menu { stack } => {
                assert_eq!(stack.len(), 1);
                assert_eq!(stack[0].section, MenuSection::File);
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
                item: 2,
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
    fn aggregate_rows_draw_dividers_instead_of_underlines() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(Some(PathBuf::from("test5.corro")));
        app.load_initial().unwrap();

        let backend = TestBackend::new(140, 24);
        let mut terminal = Terminal::new(backend).unwrap();
        terminal.draw(|f| app.draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let mut saw_underlined_last_data_row = false;
        let mut saw_left_divider = false;
        let mut saw_right_divider = false;
        let mut last_data_row_y: Option<u16> = None;
        for y in 0..buffer.area.height {
            let line = (0..buffer.area.width)
                .map(|x| buffer[(x, y)].symbol())
                .collect::<String>();
            if line.contains("│  13") && line.contains("#NAME") {
                last_data_row_y = Some(y);
                break;
            }
        }
        for y in 0..buffer.area.height {
            for x in 0..buffer.area.width {
                let cell = &buffer[(x, y)];
                if last_data_row_y == Some(y) && cell.modifier.contains(Modifier::UNDERLINED) {
                    saw_underlined_last_data_row = true;
                }
                if cell.symbol() == "│" && x < 8 {
                    saw_left_divider = true;
                }
                if cell.symbol() == "│" && x > 25 {
                    saw_right_divider = true;
                }
            }
        }

        assert!(saw_underlined_last_data_row);
        assert!(saw_left_divider);
        assert!(saw_right_divider);
    }

    #[test]
    fn aggregate_divider_sits_after_row_labels() {
        use ratatui::backend::TestBackend;
        use ratatui::Terminal;

        let mut app = App::new(Some(PathBuf::from("test5.corro")));
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

        let mut app = App::new(Some(PathBuf::from("test5.corro")));
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
    match addr {
        CellAddr::Header { row, col } => format!(
            "~{}:{}",
            HEADER_ROWS - *row as usize,
            col_header_label(*col as usize, main_cols)
        ),
        CellAddr::Footer { row, col } => format!(
            "_{}:{}",
            *row as usize + 1,
            col_header_label(*col as usize, main_cols)
        ),
        CellAddr::Main { row, col } => {
            format!("{}{}", addr::excel_column_name(*col as usize), row + 1)
        }
        CellAddr::Left { col, row } => {
            format!(
                "[{}{}",
                addr::mirror_margin_column_name(*col as usize, true),
                row + 1
            )
        }
        CellAddr::Right { col, row } => {
            format!(
                "]{}{}",
                addr::mirror_margin_column_name(*col as usize, false),
                row + 1
            )
        }
    }
}

fn input_line(
    prefix: String,
    buffer: &str,
    cursor: usize,
    text_style: Style,
    caret_style: Style,
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

    Line::from(spans)
}

fn sheet_row_label(logical_row: usize, main_rows: usize) -> String {
    let hr = HEADER_ROWS;
    if logical_row < hr {
        format!("~{}", hr - logical_row)
    } else if logical_row < hr + main_rows {
        format!("{}", logical_row - hr + 1)
    } else {
        let fr = logical_row - hr - main_rows;
        format!("_{}", fr + 1)
    }
}

fn col_header_label(global_col: usize, main_cols: usize) -> String {
    let m = MARGIN_COLS;
    if global_col < m {
        format!("[{}", addr::mirror_margin_column_name(global_col, true))
    } else if global_col < m + main_cols {
        addr::excel_column_name(global_col - m)
    } else {
        format!(
            "]{}",
            addr::mirror_margin_column_name(global_col - m - main_cols, false)
        )
    }
}

fn formula_col_fragment(global_col: usize, main_cols: usize) -> String {
    let m = MARGIN_COLS;
    if global_col < m {
        format!("[{}", addr::mirror_margin_column_name(global_col, true))
    } else if global_col < m + main_cols {
        addr::excel_column_name(global_col - m)
    } else {
        format!(
            "]{}",
            addr::mirror_margin_column_name(global_col - m - main_cols, false)
        )
    }
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
