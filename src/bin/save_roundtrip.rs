//! Save→reopen round-trip check for the auto-added margin TOTALs.
//!
//! Builds a fresh (seeded) workbook, saves it exactly the way the TUI does,
//! reopens it, and diffs the margin seed cells. Regression guard for the
//! bug where the untitled log was seeded with a bare header, so Save (which
//! renames that log verbatim) lost every margin seed.
//!
//!   cargo run --release --features ratatui --bin save-roundtrip [out.corro]

fn margin_cells(app: &corro::ui::App) -> Vec<(corro::grid::CellAddr, String)> {
    let mut v: Vec<_> = app
        .state
        .grid
        .iter_nonempty()
        .filter(|(a, _)| {
            !matches!(
                a,
                corro::grid::CellAddr::Main { .. }
            )
        })
        .filter(|(_, val)| !val.trim().is_empty())
        .collect();
    v.sort_by_key(|(a, _)| format!("{a:?}"));
    v
}

fn main() {
    let out = std::env::args()
        .nth(1)
        .map(std::path::PathBuf::from)
        .unwrap_or_else(|| std::env::temp_dir().join("corro-save-roundtrip.corro"));
    let _ = std::fs::remove_file(&out);

    let mut app = corro::ui::App::new(None);
    app.load_initial().expect("load fresh workbook");
    let before = margin_cells(&app);
    println!("margins before save: {}", before.len());
    assert!(
        before.iter().any(|(_, v)| v.eq_ignore_ascii_case("TOTAL")),
        "fresh workbook should carry the built-in TOTAL seeds"
    );

    app.save_current_to_path(&out).expect("save");
    let text = std::fs::read_to_string(&out).expect("read saved log");
    assert!(
        text.contains("TOTAL"),
        "saved log must contain the TOTAL seeds:\n{text}"
    );

    let mut re = corro::ui::App::new(Some(out.clone()));
    re.load_initial().expect("reopen");
    let after = margin_cells(&re);
    println!("margins after reopen: {}", after.len());

    for (addr, val) in &before {
        let got = re.state.grid.get(addr);
        assert_eq!(
            got.as_deref(),
            Some(val.as_str()),
            "margin cell {addr:?} lost on reopen"
        );
    }
    assert_eq!(before.len(), after.len(), "margin cell count changed");
    println!("OK: {} margin cells survived save→reopen ({})", before.len(), out.display());
}
