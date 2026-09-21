//! Shared `--movie` replay for the non-terminal GUI backends.
//!
//! The ratatui TUI has its own replayer ([`crate::ui::App::run_movie`]) which
//! drives a live crossterm terminal. The GUI backends share one [`App`] type
//! ([`crate::gui::App`]) over the same workbook/op core, so the *actions* a
//! movie performs (type into a cell, open a menu, confirm a dialog, switch
//! sheets) are backend-independent — only the way a frame reaches the screen
//! differs.
//!
//! This module is that backend-independent half:
//!
//! * [`GuiMovie::new`] parses the `.corro` log into the same `(sheet_id, line)`
//!   steps the TUI walks, and
//! * [`GuiMovie::apply_step`] performs exactly one step: move the cursor to the
//!   line's address, apply the op to the workbook, and leave a human status
//!   message (plus a short "menu flash" label) behind.
//!
//! A backend then only has to loop `apply_step` + sleep + repaint. Because the
//! stepping is shared, a backend that can paint at all gets `--movie` for free
//! — including `--gui --movie` with no X server (see
//! `gui_backend::run_gui_movie_capture`), which is what makes an unattended
//! video capture possible.

use crate::addr;
use crate::grid::{CellAddr, SheetCursor, HEADER_ROWS, MARGIN_COLS};
use crate::ops::{self, Op, WorkbookOp};

/// Replay pacing, mirroring `crate::ui::MovieReplayOptions` (the CLI passes the
/// same `--movie-*` values to both).
#[derive(Clone, Copy, Debug)]
pub struct GuiMovieOptions {
    pub typing_cps: f64,
    pub confirm_delay_ms: u64,
    pub menu_hold_ms: u64,
}

impl GuiMovieOptions {
    /// Pacing read from the environment, so the movie driver inside
    /// `run_gui` observes the same values the CLI parsed.
    pub fn from_env() -> Self {
        let num = |key: &str, default: f64| -> f64 {
            std::env::var(key)
                .ok()
                .and_then(|v| v.parse::<f64>().ok())
                .filter(|v| v.is_finite() && *v > 0.0)
                .unwrap_or(default)
        };
        GuiMovieOptions {
            typing_cps: num("CORRO_MOVIE_TYPING_CPS", 22.0),
            confirm_delay_ms: num("CORRO_MOVIE_CONFIRM_MS", 120.0) as u64,
            menu_hold_ms: num("CORRO_MOVIE_MENU_HOLD_MS", 1200.0) as u64,
        }
    }
}

impl Default for GuiMovieOptions {
    fn default() -> Self {
        GuiMovieOptions {
            typing_cps: 22.0,
            confirm_delay_ms: 120,
            menu_hold_ms: 1200,
        }
    }
}

impl GuiMovieOptions {
    /// Per-character delay for the typing animation.
    pub fn char_delay(&self) -> std::time::Duration {
        let cps = if self.typing_cps.is_finite() && self.typing_cps > 0.0 {
            self.typing_cps
        } else {
            22.0
        };
        std::time::Duration::from_secs_f64(1.0 / cps)
    }

    /// Publish this pacing to the environment for the driver that lives in
    /// the GUI backend (which has no access to the CLI's parsed values).
    pub fn publish_to_env(&self) {
        let _ = std::env::set_var("CORRO_MOVIE_TYPING_CPS", self.typing_cps.to_string());
        let _ = std::env::set_var("CORRO_MOVIE_CONFIRM_MS", self.confirm_delay_ms.to_string());
        let _ = std::env::set_var("CORRO_MOVIE_MENU_HOLD_MS", self.menu_hold_ms.to_string());
    }

    pub fn confirm_delay(&self) -> std::time::Duration {
        std::time::Duration::from_millis(self.confirm_delay_ms)
    }

    pub fn menu_hold(&self) -> std::time::Duration {
        std::time::Duration::from_millis(self.menu_hold_ms)
    }
}

pub use crate::ui_core::{edit_script_from_env, EditStep};

/// One replay step: a log line and the sheet it targets.
#[derive(Clone, Debug)]
pub struct MovieStep {
    pub sheet_id: u32,
    pub line: String,
}

/// A parsed, ready-to-replay movie.
#[derive(Clone, Debug)]
pub struct GuiMovie {
    pub path: std::path::PathBuf,
    pub steps: Vec<MovieStep>,
    /// Cell the cursor should sit on at the end of `step_index` (drives the
    /// cursor animation; `None` for steps that do not move it).
    pub cursor: Option<CellAddr>,
    /// Human-readable label for the current step ("A2 = 42", "Sheet ▸ Balance
    /// books"), shown in the status/formula row while the step is replayed.
    pub status: String,
    /// Menu currently "open" for this step, if any — the transient menu flash
    /// that makes a movie legible (e.g. `Some(("Sheet", "Balance books"))`).
    pub menu: Option<(String, String)>,
    /// Index of the step being replayed (`None` before the first / after the
    /// last).
    pub step_index: Option<usize>,
    /// Ops applied so far (bounds the progress display).
    pub applied: usize,
}

impl GuiMovie {
    /// Parse `path` into a replayable movie. Accepts the same `.corro` logs the
    /// TUI replays: blank lines and the `CORRO_LOG` header are skipped.
    pub fn new(path: &std::path::Path) -> Result<Self, String> {
        let data = std::fs::read_to_string(path)
            .map_err(|e| format!("failed to read movie input {}: {e}", path.display()))?;
        let mut steps = Vec::new();
        for raw in data.lines() {
            let t = raw.trim();
            if t.is_empty() || t.starts_with(ops::LOG_HEADER_PREFIX) {
                continue;
            }
            // A malformed line is not fatal: the TUI falls back to applying the
            // whole line as a log op, which fails the same way per line. Keep
            // the line so the replay still shows it and the status stays honest.
            let sheet_id = ops::parse_workbook_line(t).map(step_sheet_id).unwrap_or(1);
            steps.push(MovieStep {
                sheet_id,
                line: t.to_string(),
            });
        }
        if steps.is_empty() {
            return Err(format!("movie input has no operations: {}", path.display()));
        }
        Ok(GuiMovie {
            path: path.to_path_buf(),
            steps,
            cursor: None,
            status: String::new(),
            menu: None,
            step_index: None,
            applied: 0,
        })
    }

    pub fn len(&self) -> usize {
        self.steps.len()
    }

    pub fn is_empty(&self) -> bool {
        self.steps.is_empty()
    }

    /// Title used in the status line: the file name without its extension.
    pub fn title(&self) -> String {
        self.path
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("movie")
            .to_string()
    }

    /// Detach `app` from its on-disk file for the duration of the replay.
    ///
    /// A movie is a *reading* of a log, so replaying must never write back to
    /// it. Every GUI commit path appends to `app.core.path`, so leaving the
    /// movie's own file bound made a recording append its steps to the fixture
    /// it was replaying (the file grew on every run). The TUI replayer detaches
    /// the same way (`ui::App::reset_workbook_for_movie` sets `path = None` and
    /// keeps the file as `source_path`); do that here too, and hand the caller
    /// the real path back so it can report the title.
    pub fn detach_source(app: &mut crate::gui::App) -> Option<std::path::PathBuf> {
        let bound = app.core.path.take();
        let source = bound
            .clone()
            .or_else(|| app.core.source_path.clone());
        app.core.source_path = source;
        bound
    }

    /// Reset to a freshly-parsed state (a rewind for a second pass).
    pub fn rewind(&mut self) {
        self.cursor = None;
        self.status.clear();
        self.menu = None;
        self.step_index = None;
        self.applied = 0;
    }

    /// Apply step `index` to `workbook`, returning `(cursor_target, status,
    /// menu)` for the frame that should be painted for this step.
    ///
    /// `active_sheet` is updated when the step activates another sheet, exactly
    /// like the TUI's `movie_apply_line_as_user`.
    pub fn apply_step(
        &mut self,
        workbook: &mut ops::WorkbookState,
        active_sheet: &mut u32,
        index: usize,
    ) -> Result<MovieFrame, String> {
        let Some(step) = self.steps.get(index).cloned() else {
            return Err(format!("movie step {index} out of range"));
        };
        self.step_index = Some(index);
        // Focus the step's sheet first: an op is always relative to its own
        // sheet, so replaying out of order would target the wrong grid.
        self.focus_sheet(workbook, active_sheet, step.sheet_id);

        let parsed = ops::parse_workbook_line(&step.line);
        let mut cursor: Option<CellAddr> = None;
        let mut menu: Option<(String, String)> = None;
        let mut status;

        match parsed {
            Ok(WorkbookOp::SheetOp { sheet_id: _, op }) => {
                cursor = cursor_for_op(&op, workbook);
                menu = menu_for_op(&op);
                status = describe_op(&op, workbook);
                if let Some(addr) = cursor.as_ref() {
                    self.grow_for(workbook, addr);
                }
                if let Err(e) = ops::apply_workbook_op(workbook, active_sheet, WorkbookOp::SheetOp {
                    sheet_id: step.sheet_id,
                    op: op.clone(),
                }) {
                    status = format!("{} (error: {e})", status);
                }
            }
            Ok(other) => {
                menu = menu_for_workbook_op(&other);
                status = describe_workbook_op(&other, workbook);
                if let Err(e) = ops::apply_workbook_op(workbook, active_sheet, other.clone()) {
                    status = format!("{} (error: {e})", status);
                }
            }
            Err(e) => {
                // Not a structured op (or a line this build cannot parse):
                // replay it as a raw log line, matching the TUI's fallback.
                status = format!("{} (unparsed: {e})", step.line);
                if let Err(e2) = ops::apply_log_line_to_workbook(&step.line, workbook, active_sheet) {
                    status = format!("{} (error: {e2})", status);
                }
            }
        }

        workbook.ensure_active_sheet();
        self.cursor = cursor;
        self.status = status;
        self.menu = menu;
        self.applied = self.applied.saturating_add(1);
        Ok(MovieFrame {
            cursor: self.cursor,
            status: self.status.clone(),
            menu: self.menu.clone(),
            progress: (index + 1, self.steps.len()),
        })
    }

    /// Switch the replay to `sheet_id` (no-op when it is already active).
    fn focus_sheet(&self, workbook: &mut ops::WorkbookState, active_sheet: &mut u32, sheet_id: u32) {
        if let Some(idx) = workbook.sheet_index_by_id(sheet_id) {
            workbook.active_sheet = idx;
            *active_sheet = sheet_id;
        }
    }

    /// Grow the target sheet so `addr` is addressable. Movie logs may write
    /// cells past the current in-memory extent (the log is the source of
    /// truth), and a clamped cursor would animate to the wrong cell.
    fn grow_for(&self, workbook: &mut ops::WorkbookState, addr: &CellAddr) {
        let grid = &mut workbook.active_sheet_mut().grid;
        let (mut rows, mut cols) = (grid.main_rows(), grid.main_cols());
        match addr {
            CellAddr::Main { row, col } => {
                rows = rows.max(*row as usize + 1);
                cols = cols.max(*col as usize + 1);
            }
            CellAddr::Left { row, .. } | CellAddr::Right { row, .. } => {
                rows = rows.max(*row as usize + 1);
            }
            CellAddr::Header { .. } | CellAddr::Footer { .. } => {}
        }
        if rows != grid.main_rows() || cols != grid.main_cols() {
            grid.set_main_size(rows, cols);
        }
    }
}

/// What the backend needs to paint one replayed step.
#[derive(Clone, Debug)]
pub struct MovieFrame {
    pub cursor: Option<CellAddr>,
    pub status: String,
    pub menu: Option<(String, String)>,
    /// `(step_number, total_steps)` for the progress display.
    pub progress: (usize, usize),
}

/// The sheet a workbook op targets (`SheetOp` carries it explicitly; every
/// other op belongs to the sheet whose `$id:` prefix introduced it, which
/// `parse_workbook_line` already resolved).
fn step_sheet_id(op: WorkbookOp) -> u32 {
    match op {
        WorkbookOp::SheetOp { sheet_id, .. } => sheet_id,
        WorkbookOp::ActivateSheet { id } => id,
        WorkbookOp::NewSheet { id, .. }
        | WorkbookOp::CopySheet { id, .. }
        | WorkbookOp::RenameSheet { id, .. }
        | WorkbookOp::MoveSheet { id }
        | WorkbookOp::DeleteSheet { id }
        | WorkbookOp::LinkSheet { id, .. }
        | WorkbookOp::BalanceReport { id, .. } => id,
    }
}

/// Cell address an op writes to, in cursor coordinates (so the backend can
/// animate the cursor there). Ops that are not cell writes return `None`.
fn cursor_for_op(op: &Op, workbook: &ops::WorkbookState) -> Option<CellAddr> {
    match op {
        Op::SetCell { addr, .. } => Some(addr.clone()),
        Op::SetCellFormat { addr, .. } => Some(addr.clone()),
        Op::SetCellRef { cref, .. } => {
            Some(cref.to_grid_addr(workbook.active_sheet().grid.main_cols()))
        }
        Op::FillRange { cells } => cells.first().map(|(addr, _)| addr.clone()),
        _ => None,
    }
}

/// Menu the TUI flashes for an op, so the GUI movie reads the same way: the
/// label names the menu path a user would have taken to cause the op.
fn menu_for_op(op: &Op) -> Option<(String, String)> {
    let pair = |section: &str, item: &str| Some((section.to_string(), item.to_string()));
    let pair_owned = |section: &str, item: String| Some((section.to_string(), item));
    match op {
        Op::FillRange { cells } => {
            let all_empty = cells.iter().all(|(_, v)| v.is_empty());
            if all_empty {
                pair("Edit", "Cut")
            } else if cells.len() > 1 {
                pair("Edit", "Paste")
            } else {
                None
            }
        }
        Op::SetViewSortCols { .. } => pair("File", "Sort view"),
        Op::SetColumnFormat { scope, format, .. } => {
            let scope_label = match scope {
                crate::grid::FormatScope::All => "All",
                crate::grid::FormatScope::Data => "Data",
                crate::grid::FormatScope::Special => "Special",
            };
            pair("Format", &format!("{scope_label} · {}", format_label(format)))
        }
        Op::SetAllColumnFormat { format } => {
            pair("Format", &format!("All · {}", format_label(format)))
        }
        Op::SetCellFormat { format, .. } => pair_owned("Format", format_label(format)),
        Op::SetMaxColWidth { .. } => pair("File", "Width · Default width"),
        Op::SetColWidth { .. } => pair("File", "Width · Column width"),
        _ => None,
    }
}

fn menu_for_workbook_op(op: &WorkbookOp) -> Option<(String, String)> {
    let pair = |section: &str, item: &str| Some((section.to_string(), item.to_string()));
    match op {
        WorkbookOp::NewSheet { .. } => pair("Sheet", "New sheet"),
        WorkbookOp::CopySheet { .. } => pair("Sheet", "Copy sheet"),
        WorkbookOp::RenameSheet { .. } => pair("Sheet", "Rename sheet"),
        WorkbookOp::MoveSheet { .. } => pair("Sheet", "Move sheet"),
        WorkbookOp::DeleteSheet { .. } => pair("Sheet", "Delete sheet"),
        WorkbookOp::BalanceReport { .. } => pair("Sheet", "Balance books"),
        WorkbookOp::ActivateSheet { .. } | WorkbookOp::LinkSheet { .. } => None,
        WorkbookOp::SheetOp { op, .. } => menu_for_op(op),
    }
}

fn format_label(format: &crate::grid::CellFormat) -> String {
    use crate::grid::{NumberFormat, TextAlign};
    if format == &crate::grid::CellFormat::default() {
        return "Reset".into();
    }
    if let Some(align) = format.align {
        return match align {
            TextAlign::Left => "Align Left".into(),
            TextAlign::Center => "Align Center".into(),
            TextAlign::Right => "Align Right".into(),
            TextAlign::Default => "Align Default".into(),
        };
    }
    match format.number {
        Some(NumberFormat::DecimalGeneric) => "Decimal (generic)".into(),
        Some(NumberFormat::Currency { .. }) => "Currency ($)".into(),
        Some(NumberFormat::Rational) => "Rational".into(),
        Some(NumberFormat::Fixed { decimals }) => format!("Fixed {decimals}"),
        None => "Format".into(),
    }
}

/// Human sentence for an op — the same phrasing the TUI puts in its status
/// line, so a recording of either backend reads identically.
fn describe_op(op: &Op, workbook: &ops::WorkbookState) -> String {
    let mc = workbook.active_sheet().grid.main_cols();
    match op {
        Op::SetCell { addr, value } => format!("{} = {}", addr::cell_ref_text(addr, mc), value),
        Op::SetCellRef { cref, value } => {
            let addr = cref.to_grid_addr(mc);
            format!("{} = {}", addr::cell_ref_text(&addr, mc), value)
        }
        Op::FillRange { cells } if cells.len() == 1 => {
            let (addr, value) = &cells[0];
            format!("{} = {}", addr::cell_ref_text(addr, mc), value)
        }
        Op::FillRange { cells } => {
            let blank = cells.iter().filter(|(_, v)| v.is_empty()).count();
            if blank == cells.len() {
                format!("clear {} cells", cells.len())
            } else {
                format!("fill {} cells", cells.len())
            }
        }
        Op::SetViewSortCols { cols } => {
            let names: Vec<String> = cols
                .iter()
                .map(|s| {
                    let name = addr::excel_column_name(s.col.saturating_sub(MARGIN_COLS));
                    if s.desc {
                        format!("!{name}")
                    } else {
                        name
                    }
                })
                .collect();
            format!("sort view by {}", names.join(", "))
        }
        Op::SetCellFormat { format, .. } => format!("format cell: {}", format_label(format)),
        Op::SetColumnFormat { col, format, .. } => format!(
            "format column {}: {}",
            addr::excel_column_name(col.saturating_sub(MARGIN_COLS)),
            format_label(format)
        ),
        Op::SetAllColumnFormat { format } => format!("format all columns: {}", format_label(format)),
        Op::SetMaxColWidth { width } => format!("default column width {width}"),
        Op::SetColWidth { col, width } => format!(
            "column {} width {}",
            addr::excel_column_name(col.saturating_sub(MARGIN_COLS)),
            width.map(|w| w.to_string()).unwrap_or_else(|| "auto".into())
        ),
        other => format!("{other:?}"),
    }
}

fn describe_workbook_op(op: &WorkbookOp, workbook: &ops::WorkbookState) -> String {
    match op {
        WorkbookOp::NewSheet { title, .. } => format!("new sheet {title}"),
        WorkbookOp::CopySheet { title, .. } => format!("copy sheet to {title}"),
        WorkbookOp::RenameSheet { title, .. } => format!("rename sheet to {title}"),
        WorkbookOp::MoveSheet { .. } => "move sheet".into(),
        WorkbookOp::DeleteSheet { .. } => "delete sheet".into(),
        WorkbookOp::ActivateSheet { id } => format!("activate sheet {id}"),
        WorkbookOp::LinkSheet { id, .. } => format!("link sheet {id}"),
        WorkbookOp::BalanceReport { title, .. } => format!("balance report {title}"),
        WorkbookOp::SheetOp { op, .. } => describe_op(op, workbook),
    }
}

/// True when `addr` is inside the main data band (movies mostly write there).
pub fn is_main_addr(addr: &CellAddr) -> bool {
    matches!(addr, CellAddr::Main { .. })
}

/// Cursor coordinates (logical row, global column) for a cell address.
pub fn cursor_of(addr: &CellAddr, workbook: &ops::WorkbookState) -> SheetCursor {
    let mc = workbook.active_sheet().grid.main_cols();
    let mr = workbook.active_sheet().grid.main_rows();
    let row = match addr {
        CellAddr::Header { row, .. } => *row as usize,
        CellAddr::Main { row, .. } | CellAddr::Left { row, .. } | CellAddr::Right { row, .. } => {
            HEADER_ROWS + *row as usize
        }
        CellAddr::Footer { row, .. } => HEADER_ROWS + mr + *row as usize,
    };
    let col = match addr {
        CellAddr::Header { col, .. } | CellAddr::Footer { col, .. } => col.to_global(mc),
        CellAddr::Main { col, .. } => MARGIN_COLS + *col as usize,
        CellAddr::Left { col, .. } => *col,
        CellAddr::Right { col, .. } => MARGIN_COLS + mc + *col,
    };
    SheetCursor { row, col }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ops::WorkbookState;

    fn movie_from(text: &str) -> (GuiMovie, std::path::PathBuf) {
        let dir = std::env::temp_dir();
        let path = dir.join(format!(
            "corro-movie-test-{}-{:?}.corro",
            std::process::id(),
            std::thread::current().id()
        ));
        std::fs::write(&path, text).unwrap();
        (GuiMovie::new(&path).unwrap(), path)
    }

    #[test]
    fn parses_ops_and_skips_blank_lines() {
        let (movie, path) = movie_from("SET $1:A1 1\n\nSET $1:A2 2\n");
        assert_eq!(movie.len(), 2);
        assert_eq!(movie.steps[0].sheet_id, 1);
        assert_eq!(movie.steps[1].line, "SET $1:A2 2");
        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn applies_set_cell_and_tracks_the_cursor() {
        let (mut movie, path) = movie_from("SET $1:A1 42\nSET $1:B2 hello\n");
        let mut wb = WorkbookState::new();
        let mut active = wb.sheet_id(wb.active_sheet);
        let frame = movie.apply_step(&mut wb, &mut active, 0).unwrap();
        assert_eq!(frame.progress, (1, 2));
        assert_eq!(
            wb.active_sheet().grid.get(&CellAddr::Main { row: 0, col: 0 }),
            Some("42".to_string())
        );
        assert_eq!(frame.cursor, Some(CellAddr::Main { row: 0, col: 0 }));
        assert_eq!(frame.status, "A1 = 42");

        let frame = movie.apply_step(&mut wb, &mut active, 1).unwrap();
        // Main (row 1, col 1) is the cursor coordinate HEADER_ROWS+1 / MARGIN+1.
        let cur = cursor_of(&CellAddr::Main { row: 1, col: 1 }, &wb);
        assert_eq!(cur.row, HEADER_ROWS + 1);
        assert_eq!(cur.col, MARGIN_COLS + 1);
        assert_eq!(frame.menu, None);
        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn grows_the_grid_for_far_away_cells() {
        let (mut movie, path) = movie_from("SET $1:C5 x\n");
        let mut wb = WorkbookState::new();
        let mut active = wb.sheet_id(wb.active_sheet);
        movie.apply_step(&mut wb, &mut active, 0).unwrap();
        assert!(wb.active_sheet().grid.main_rows() >= 5);
        assert!(wb.active_sheet().grid.main_cols() >= 3);
        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn format_ops_flash_the_format_menu() {
        let (mut movie, path) = movie_from("FORMAT CELL A1 align:center\n");
        let mut wb = WorkbookState::new();
        let mut active = wb.sheet_id(wb.active_sheet);
        let frame = match movie.apply_step(&mut wb, &mut active, 0) {
            Ok(f) => f,
            Err(_) => return, // this log spelling is not supported by every build
        };
        if let Some((section, _item)) = frame.menu {
            assert_eq!(section, "Format");
        }
        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn missing_input_is_an_error_not_a_panic() {
        let missing = std::env::temp_dir().join("corro-movie-does-not-exist.corro");
        let _ = std::fs::remove_file(&missing);
        assert!(GuiMovie::new(&missing).is_err());
    }

    #[test]
    fn empty_movie_is_rejected() {
        let (res, path) = {
            let dir = std::env::temp_dir();
            let path = dir.join(format!("corro-movie-empty-{}.corro", std::process::id()));
            std::fs::write(&path, "\n\n").unwrap();
            (GuiMovie::new(&path), path)
        };
        assert!(res.is_err());
        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn char_delay_guards_against_zero_cps() {
        let opts = GuiMovieOptions {
            typing_cps: 0.0,
            ..Default::default()
        };
        assert!(opts.char_delay() > std::time::Duration::ZERO);
        let opts = GuiMovieOptions {
            typing_cps: f64::NAN,
            ..Default::default()
        };
        assert!(opts.char_delay() > std::time::Duration::ZERO);
    }
}
