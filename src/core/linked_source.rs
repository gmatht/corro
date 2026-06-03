use crate::ops::WorkbookState;
use std::collections::HashMap;
use std::path::PathBuf;
use std::time::SystemTime;

/// Check if any linked source file has been modified since the given mtimes.
pub fn linked_sources_changed(
    workbook: &WorkbookState,
    mtimes: &HashMap<PathBuf, SystemTime>,
) -> bool {
    for sheet in &workbook.sheets {
        let Some(source) = &sheet.linked_source else {
            continue;
        };
        let Ok(meta) = std::fs::metadata(&source.path) else {
            continue;
        };
        let Ok(modified) = meta.modified() else {
            continue;
        };
        match mtimes.get(&source.path) {
            Some(prev) if *prev >= modified => {}
            _ => return true,
        }
    }
    false
}

/// Refresh cached mtimes from the linked sources in the workbook.
pub fn refresh_linked_source_mtimes(
    workbook: &WorkbookState,
    mtimes: &mut HashMap<PathBuf, SystemTime>,
) {
    mtimes.clear();
    for sheet in &workbook.sheets {
        if let Some(source) = &sheet.linked_source {
            if let Ok(meta) = std::fs::metadata(&source.path) {
                if let Ok(modified) = meta.modified() {
                    mtimes.insert(source.path.clone(), modified);
                }
            }
        }
    }
}
