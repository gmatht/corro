use crate::grid::{CellAddr, MainRange, SheetCursor, SortSpec};
use crate::io::LogWatcher;
use crate::ops::{Op, SheetState, WorkbookState};
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

#[cfg(test)]
mod tests {
    use super::normal_hints;

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
