use crate::grid::{CellAddr, MainRange, SheetCursor, SortSpec};
use crate::io::{tail_apply_workbook, IoError, LogWatcher};
use crate::ops::{apply_log_line_to_workbook, Op, SheetState, WorkbookState};
use std::collections::HashMap;
use std::path::PathBuf;
use std::time::SystemTime;

/// Bottom-strip hints for the normal state, shared by every UI: the
/// ratatui bottom row always shows these (never status — that lives in
/// the formula bar trailing), and the GUI bottom strip shows exactly the
/// same text. Pure over four booleans so terminal (`ui::App`) and GUI
/// (`CoreApp`) share one source of truth instead of two hint tables.
/// Callers pass live state (undo/redo availability, save target); the
/// text must stay byte-identical everywhere it renders.
pub fn normal_hints(anchored: bool, can_undo: bool, can_redo: bool, has_path: bool) -> String {
    if anchored {
        return "  r·move-rows   c·move-cols   v·deselect   Esc·cancel".into();
    }
    let mut hints = vec!["type/F2·edit", "Ctrl+C·copy", "Ctrl+X·cut", "Ctrl+V·paste"];
    if can_undo {
        hints.push("Ctrl+Z·undo");
    }
    if can_redo {
        hints.push("Ctrl+Y·redo");
    }
    hints.push("Ctrl+;·date");
    hints.push("Ctrl+:·time");
    hints.push(if has_path {
        "Ctrl+S·save"
    } else {
        "Ctrl+S·save as"
    });
    hints.push("F1·help");
    format!("  {}", hints.join("; "))
}

pub struct CoreApp {
    pub path: Option<PathBuf>,
    pub import_source: Option<PathBuf>,
    pub source_path: Option<PathBuf>,
    pub revision_limit: Option<usize>,
    pub revision_browse: bool,
    pub revision_browse_limit: usize,
    pub offset: u64,
    pub state: SheetState,
    pub workbook: WorkbookState,
    pub cursor: SheetCursor,
    pub anchor: Option<SheetCursor>,
    pub watcher: Option<LogWatcher>,
    pub status: String,
    pub ops_applied: usize,
    pub op_history: Vec<Op>,
    pub redo_history: Vec<Op>,
    pub view_sheet_id: u32,
    pub persisted_view_sort_cols: HashMap<u32, Vec<SortSpec>>,
    pub linked_source_mtimes: HashMap<PathBuf, SystemTime>,
    pub unsaved_file: Option<PathBuf>,
    pub unsaved_auto_create: bool,
    pub exit_message: Option<String>,
    pub clipboard_snapshot: Option<(MainRange, String)>,
    pub edit_target_addr: Option<CellAddr>,
    pub edit_range_addrs: Option<Vec<CellAddr>>,
    pub pending_lost_edit: Option<(CellAddr, String)>,
    pub pending_fit_to_content_on_commit: bool,
}

/// Log-tail polling for backends that drive [`CoreApp`] directly (the GUI).
///
/// Twin of the log-tail half of `ui::App::sync_external` (which the TUI
/// loop calls every iteration): when another window appends revisions to
/// the bound file, tail-apply them here so both windows show the same
/// cells. Commits append to the log immediately, so mirroring needs no
/// save step on either side. Keep the two in sync: same watcher-then-size
/// check order, same tail-apply-then-extent-restore sequence, same status
/// texts, same full-reload fallback when the file was rewritten (Save)
/// instead of appended to.
impl CoreApp {
    pub fn poll_log_tail(&mut self) -> Result<bool, IoError> {
        let Some(ref p) = self.path.clone() else {
            return Ok(false);
        };
        // notify cannot watch a path that does not exist yet; (re)try until
        // it does. Absent watcher => the size check below still fires.
        if self.watcher.is_none() {
            if let Ok(w) = LogWatcher::new(p.clone()) {
                self.watcher = Some(w);
            }
        }
        let mut should_tail = false;
        if let Some(w) = &self.watcher {
            if w.poll_dirty() {
                should_tail = true;
            }
        }
        if !should_tail {
            if let Ok(meta) = std::fs::metadata(p) {
                if meta.len() > self.offset {
                    should_tail = true;
                }
            }
        }
        if !should_tail {
            return Ok(false);
        }
        // Save the in-memory extent so we can restore it after the reload.
        let saved_rows = self.state.grid.main_rows();
        let saved_cols = self.state.grid.main_cols();
        match tail_apply_workbook(p, self.offset, &mut self.workbook, &mut self.view_sheet_id) {
            Ok(new_off) => {
                if new_off > self.offset {
                    self.offset = new_off;
                    self.sync_sheet_cache();
                    let cur_rows = self.state.grid.main_rows();
                    let cur_cols = self.state.grid.main_cols();
                    if saved_rows > cur_rows || saved_cols > cur_cols {
                        self.state
                            .grid
                            .set_main_size(saved_rows.max(cur_rows), saved_cols.max(cur_cols));
                    }
                    self.status = "External change applied".into();
                    Ok(true)
                } else {
                    self.offset = new_off;
                    Ok(false)
                }
            }
            Err(_) => {
                // Counterpart rewrote the file (Save) instead of appending:
                // fall back to a full reload.
                let data = std::fs::read_to_string(p)?;
                let mut workbook = WorkbookState::new();
                let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
                for line in data.lines() {
                    let t = line.trim();
                    if t.is_empty() {
                        continue;
                    }
                    apply_log_line_to_workbook(t, &mut workbook, &mut active_sheet)?;
                }
                self.workbook = workbook;
                self.view_sheet_id = active_sheet;
                self.sync_sheet_cache();
                let cur_rows = self.state.grid.main_rows();
                let cur_cols = self.state.grid.main_cols();
                if saved_rows > cur_rows || saved_cols > cur_cols {
                    self.state
                        .grid
                        .set_main_size(saved_rows.max(cur_rows), saved_cols.max(cur_cols));
                }
                self.offset = data.len() as u64;
                self.ops_applied = data.lines().filter(|line| !line.trim().is_empty()).count();
                self.status = "File reset; full reload".into();
                Ok(true)
            }
        }
    }

    fn sync_sheet_cache(&mut self) {
        self.workbook.ensure_active_sheet();
        if let Some(idx) = self.workbook.sheet_index_by_id(self.view_sheet_id) {
            self.workbook.active_sheet = idx;
            self.state = self.workbook.sheets[idx].state.clone();
        } else {
            self.view_sheet_id = self.workbook.sheet_id(self.workbook.active_sheet);
            self.state = self.workbook.active_sheet().clone();
        }
    }
}

#[cfg(test)]
mod tests {
    use super::normal_hints;
    use super::CoreApp;
    use crate::grid::CellAddr;
    use crate::io::{commit_workbook_op, load_workbook_revisions_partial};
    use crate::ops::{apply_workbook_op, Op, WorkbookOp, WorkbookState};

    fn test_core_app(path: Option<std::path::PathBuf>, workbook: WorkbookState, offset: u64) -> CoreApp {
        super::CoreApp {
            path,
            import_source: None,
            source_path: None,
            revision_limit: None,
            revision_browse: false,
            revision_browse_limit: 0,
            offset,
            state: Default::default(),
            workbook,
            cursor: crate::grid::SheetCursor { row: 0, col: 0 },
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
            status: String::new(),
            exit_message: None,
            clipboard_snapshot: None,
            edit_target_addr: None,
            edit_range_addrs: None,
            pending_lost_edit: None,
            pending_fit_to_content_on_commit: false,
        }
    }

    /// CoreApp twin of the two-window mirror: B.poll_log_tail() picks up
    /// A's committed cell. This is the exact path the GUI keystroke hook
    /// drives (gui_backend polls app.core on every key).
    #[test]
    fn core_second_window_tails_first_windows_commit() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("shared.corro");
        let cell = CellAddr::Main { row: 0, col: 0 };
        // Window A opens a fresh sheet bound to the file.
        let mut a = test_core_app(Some(path.clone()), WorkbookState::new_seeded(), 0);
        let mut a_active = a.workbook.sheets[0].id;
        let seed = WorkbookOp::SheetOp {
            sheet_id: a_active,
            op: Op::SetCell {
                addr: cell.clone(),
                value: "seed".into(),
            },
        };
        apply_workbook_op(&mut a.workbook, &mut a_active, seed.clone()).unwrap();
        commit_workbook_op(&path, &mut a.offset, &mut a.workbook, &mut a_active, &seed).unwrap();
        // Window B opens the file (full load, offset at end).
        let mut b_wb = WorkbookState::new();
        let mut b_active = 0u32;
        let (off, _) =
            load_workbook_revisions_partial(&path, usize::MAX, &mut b_wb, &mut b_active).unwrap();
        let mut b = test_core_app(Some(path.clone()), b_wb, off);
        b.view_sheet_id = b_active;
        // Window A commits a new cell (append to log + apply in memory).
        let op = WorkbookOp::SheetOp {
            sheet_id: a_active,
            op: Op::SetCell {
                addr: cell.clone(),
                value: "MIRROR".into(),
            },
        };
        apply_workbook_op(&mut a.workbook, &mut a_active, op.clone()).unwrap();
        commit_workbook_op(&path, &mut a.offset, &mut a.workbook, &mut a_active, &op).unwrap();
        // B is stale until it polls ...
        assert_ne!(
            b.workbook.sheets[0].state.grid.get(&cell).as_deref(),
            Some("MIRROR")
        );
        // ... then the tail applies A's commit with no save step.
        assert!(b.poll_log_tail().unwrap(), "B should observe A's commit");
        assert_eq!(
            b.workbook.sheets[0].state.grid.get(&cell).as_deref(),
            Some("MIRROR")
        );
        assert_eq!(b.offset, a.offset);
        assert_eq!(b.status, "External change applied");
        // A second poll is quiet.
        assert!(!b.poll_log_tail().unwrap());
    }

    #[test]
    fn normal_hints_covers_anchor_undo_redo_and_save_states() {
        let fresh = normal_hints(false, false, false, false);
        assert!(fresh.contains("type/F2·edit"), "{fresh:?}");
        assert!(fresh.contains("Ctrl+S·save as"), "{fresh:?}");
        assert!(!fresh.contains("undo"), "{fresh:?}");
        let undo = normal_hints(false, true, false, true);
        assert!(undo.contains("Ctrl+Z·undo"), "{undo:?}");
        assert!(undo.contains("Ctrl+S·save"), "{undo:?}");
        assert!(!undo.contains("save as"), "{undo:?}");
        let redo = normal_hints(false, false, true, false);
        assert!(redo.contains("Ctrl+Y·redo"), "{redo:?}");
        let anchored = normal_hints(true, true, true, true);
        assert!(anchored.contains("move-rows"), "{anchored:?}");
        assert!(!anchored.contains("Ctrl+C"), "{anchored:?}");
    }
}
