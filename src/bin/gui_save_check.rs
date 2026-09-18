//! GUI save path check: the GUI always serializes a fresh workbook through
//! `io::write_workbook_log`, so a saved-then-reopened document must keep the
//! built-in margin TOTAL seeds. Mirrors src/gui/mod.rs's fresh-document
//! construction (WorkbookState::new_seeded) + the "save" menu action.
//!
//!   cargo run --features ratatui --bin gui-save-check [out.corro]

fn main() {
    let out = std::env::args()
        .nth(1)
        .map(std::path::PathBuf::from)
        .unwrap_or_else(|| std::env::temp_dir().join("corro-gui-save-check.corro"));
    let _ = std::fs::remove_file(&out);

    // Exactly what MenuInit hands the GUI for a no-paths launch.
    let mut workbook = corro::ops::WorkbookState::new_seeded();
    workbook
        .active_sheet_mut()
        .grid
        .set(&corro::grid::CellAddr::Main { row: 0, col: 0 }, "5".into());

    let margins = |wb: &corro::ops::WorkbookState| {
        let mut v: Vec<(corro::grid::CellAddr, String)> = wb
            .active_sheet()
            .grid
            .iter_nonempty()
            .filter(|(a, val)| {
                !matches!(a, corro::grid::CellAddr::Main { .. }) && !val.trim().is_empty()
            })
            .collect();
        v.sort_by_key(|(a, _)| format!("{a:?}"));
        v
    };
    let before = margins(&workbook);
    assert!(
        before.iter().any(|(_, v)| v.eq_ignore_ascii_case("TOTAL")),
        "fresh GUI workbook must carry TOTAL seeds"
    );

    // The GUI "save" action with a path: canonical CORRO_LOG writer.
    corro::io::write_workbook_log(&out, &workbook, &Default::default()).expect("gui save");

    let reopened = corro::io::load_workbook_file(&out).expect("reopen");
    let after = margins(&reopened);
    for (addr, val) in &before {
        assert_eq!(
            reopened.active_sheet().grid.get(addr).as_deref(),
            Some(val.as_str()),
            "GUI save/reopen lost margin cell {addr:?}"
        );
    }
    assert_eq!(before.len(), after.len(), "margin cell count changed");
    println!(
        "OK: GUI save path kept {} margin cells across reopen ({})",
        before.len(),
        out.display()
    );
}
