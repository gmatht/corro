use crate::core::state::CoreApp;
mod app_alias;
use crate::grid::{CellAddr, SheetCursor, HEADER_ROWS, MARGIN_COLS};
use crate::io::load_workbook_revisions_partial;
use crate::io::PartialReplay;
use std::path::Path;
use crate::ops::WorkbookState;
use std::path::PathBuf;

pub mod actions;
pub mod clipboard;
pub mod viewport;
pub mod compute;
pub mod dialogs;
pub mod edit;
pub mod extrapolate;
pub mod keymap;
pub mod menu;
pub mod render;
pub mod sheet;
pub mod special_picker;

#[cfg(any(feature = "gui", all(feature = "wasm", target_arch = "wasm32")))]
mod gui_backend;
#[cfg(all(feature = "gui", target_os = "android"))]
mod android_backend;
#[cfg(feature = "pancurses")]
mod pnc_backend;

pub enum Backend {
    Gui,
    Pancurses,
}

pub struct App {
    pub core: CoreApp,
    rev_limit: Option<usize>,
    rev_browse: bool,
    backend: Option<Backend>,
    /// Interactive extrapolate modal state (shared by all GUI backends).
    pub extrapolate: Option<extrapolate::ExtrapolateModal>,
    /// Insert > Special Char picker selection (shared by all GUI backends).
    /// `Some(idx)` while the picker is open; backends render
    /// [`special_picker::items`] and drive it via that module's
    /// open/step/set/close/take functions.
    pub special_picker: Option<usize>,
}

impl App {
    pub fn new_with_paths(paths: Vec<PathBuf>) -> Self {
        // Fresh-document content mirrors the TUI (App::new_with_revision_limit):
        // a CORRO_TEMPLATE workbook when set and readable, else the built-in
        // seeded blank (margin TOTAL seeds). Loads replay file ops onto a
        // plain blank, exactly like the TUI, so only genuinely fresh docs
        // (no paths) branch here.
        let (workbook, template_note) = if paths.is_empty() {
            match crate::io::template_path_from_env() {
                Some(tpath) => match crate::io::load_workbook_template(tpath.as_path()) {
                    Ok(wb) => (wb, None),
                    Err(msg) => (
                        WorkbookState::new_seeded(),
                        Some(format!(
                            "Template {} failed ({msg}); opened blank instead",
                            tpath.display()
                        )),
                    ),
                },
                None => (WorkbookState::new_seeded(), None),
            }
        } else {
            (WorkbookState::new(), None)
        };
        App {
            core: CoreApp {
                path: paths.first().cloned(),
                import_source: None,
                source_path: None,
                revision_limit: None,
                revision_browse: false,
                revision_browse_limit: 0,
                offset: 0,
                state: Default::default(),
                workbook,
                cursor: SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS },
                anchor: None,
                watcher: None,
                ops_applied: 0,
                op_history: Vec::new(),
                redo_history: Vec::new(),
                view_sheet_id: 0,
                persisted_view_sort_cols: Default::default(),
                linked_source_mtimes: Default::default(),
                unsaved_file: None,
                unsaved_auto_create: false,
                status: template_note.unwrap_or_default(),
                exit_message: None,
                clipboard_snapshot: None,
                edit_target_addr: None,
                edit_range_addrs: None,
                pending_lost_edit: None,
                pending_fit_to_content_on_commit: false,
            },
            rev_limit: None,
            rev_browse: false,
            backend: None,
            extrapolate: None,
            special_picker: None,
        }
    }

    pub fn new_with_revision_browser(path: Option<PathBuf>) -> Self {
        let mut app = Self::new_with_paths(path.into_iter().collect());
        app.rev_browse = true;
        app.core.revision_browse = true;
        app
    }

    pub fn new_with_revision_limit(path: Option<PathBuf>, limit: Option<usize>) -> Self {
        let mut app = Self::new_with_paths(path.into_iter().collect());
        app.rev_limit = limit;
        app.core.revision_limit = limit;
        app
    }

    pub fn load_initial(&mut self) -> Result<(), Box<dyn std::error::Error>> {
        let path = self.core.path.clone();
        if let Some(p) = &path {
            if p.exists() {
                let ext = p.extension().and_then(|e| e.to_str()).unwrap_or("").to_ascii_lowercase();
                match ext.as_str() {
                    "corro" => {
                        let mut active_sheet = self.core.workbook.sheet_id(self.core.workbook.active_sheet);
                        let (offset, replay) = load_workbook_revisions_partial(
                            p, self.rev_limit.unwrap_or(usize::MAX),
                            &mut self.core.workbook, &mut active_sheet,
                        ).map_err(|e| format!("failed to load: {e}"))?;
                        self.core.offset = offset;
                        self.core.ops_applied = replay.op_count;
                        self.core.status = Self::replay_status("Loaded workbook", p, &replay);
                        if let Some(i) = self.core.workbook.sheets.iter().position(|s| s.id == active_sheet) {
                            self.core.workbook.active_sheet = i;
                        }
                    }
                    "ods" => {
                        let wb = crate::ods::import_ods_workbook(p).map_err(|e| format!("failed to import ODS: {e}"))?;
                        self.core.workbook = wb;
                    }
                    "tsv" => {
                        let data = std::fs::read_to_string(p).map_err(|e| format!("failed to read TSV: {e}"))?;
                        crate::io::import_tsv(&data, self.core.workbook.active_sheet_mut());
                    }
                    "csv" => {
                        let data = std::fs::read_to_string(p).map_err(|e| format!("failed to read CSV: {e}"))?;
                        crate::io::import_csv(&data, self.core.workbook.active_sheet_mut());
                    }
                    _ => return Err(format!("unsupported file type: {ext}").into()),
                }
            }
        }
        if self.core.workbook.sheets.is_empty() {
            self.core.workbook.sheets.push(crate::ops::SheetRecord {
                id: 1,
                title: "Sheet1".to_string(),
                state: crate::ops::SheetState::new(1, 1),
                linked_source: None,
            });
            self.core.workbook.active_sheet = 0;
        }
        Ok(())
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

    /// Fit all main columns to their rendered content (matching ratatui's
    /// fit_column_to_rendered_content called during load_initial).
    pub fn fit_main_columns_to_max_width(&mut self) {
        let sheet = self.core.workbook.active_sheet_mut();
        let grid = &mut sheet.grid;
        let mc = grid.main_cols();
        for c in 0..mc {
            let global_col = MARGIN_COLS + c;
            if let Some(rw) = crate::ui_core::rendered_width_for_column(grid, global_col) {
                let capped = rw.min(grid.max_col_width());
                grid.set_col_width(global_col, Some(capped));
            }
        }
    }

    pub fn set_backend(&mut self, backend: Backend) {
        self.backend = Some(backend);
    }

    /// Save the current workbook to its path (mirrors the "save" GUI action).
    pub fn save(&mut self) -> Result<(), Box<dyn std::error::Error>> {
        let p = self
            .core
            .path
            .clone()
            .ok_or_else(|| -> Box<dyn std::error::Error> { "no file path set".into() })?;
        crate::io::write_workbook_log(
            &p,
            &self.core.workbook,
            &self.core.persisted_view_sort_cols,
        )?;
        self.core.status = "Saved".into();
        Ok(())
    }

    /// Set a cell's text on the active sheet (headless/test helper).
    pub fn set_cell(&mut self, addr: CellAddr, text: String) {
        let sheet = self.core.workbook.active_sheet_mut();
        sheet.grid.set(&addr, text);
    }

    /// Sort the active sheet's main grid by `col` (0-based main column),
    /// ascending when `asc` is true. This is the routine the GTK
    /// "Sort ▸ Ascending / Descending" menu actions run once the sort dialog
    /// resolves a column and direction — i.e. the exact code path a faked
    /// keypress into that dialog's confirm button triggers. It applies a real
    /// view-sort to the underlying grid so the spreadsheet is actually sorted.
    pub fn sort_by_column(&mut self, col: usize, asc: bool) {
        use crate::grid::SortSpec;
        let global_col = MARGIN_COLS + col;
        let spec = SortSpec { col: global_col, desc: !asc };
        let sheet = self.core.workbook.active_sheet_mut();
        sheet.grid.set_view_sort_cols(vec![spec.clone()]);
        self.core
            .persisted_view_sort_cols
            .insert(self.core.workbook.active_sheet as u32, vec![spec]);
    }

    /// Borrow the active workbook. Used by the rustxWidgets terminal adapter
    /// (`crate::ui::rswidgets_term`) to build a `SpreadsheetModel`.
    pub fn workbook(&self) -> &WorkbookState {
        &self.core.workbook
    }

    pub fn run(&mut self) -> Result<(), Box<dyn std::error::Error>> {
        #[cfg(any(feature = "gui", all(feature = "wasm", target_arch = "wasm32")))]
        if self.backend.as_ref().map_or(true, |b| matches!(b, Backend::Gui)) {
            return gui_backend::run_gui(self);
        }
        #[cfg(feature = "pancurses")]
        if self.backend.as_ref().map_or(true, |b| matches!(b, Backend::Pancurses)) {
            return pnc_backend::run_pancurses(self);
        }
        Err("Unknown backend".into())
    }

    pub fn take_final_exit_hint(&mut self) -> Option<String> {
        self.core.exit_message.take()
    }
}
