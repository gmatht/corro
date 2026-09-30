/// Ensure the blessed test-data file contains no stale replayer-output lines
/// (elloH / estT).  Replayer test sessions used to modify the file in-place,
/// accumulating duplicate SET commands that broke the golden export tests.
///
/// # This test does not write anything
///
/// It used to copy `docs/tests/subtotal-tiny.corro` over a tracked
/// `test_rec5.corro` at the repo root, on the stated grounds that the
/// destination was "a working copy — it is git-ignored". It is not ignored; it
/// is tracked, and the test has been writing a tracked file on every run.
///
/// That was a no-op only by luck: the copy is byte-identical to its source, so
/// `git status` stayed clean and nothing looked wrong. Change the source and
/// the next run would silently stage the copy. It is the same shape as the bug
/// this repo has hit twice with `docs/ZORK_WALKTHROUGH.txt` — a test writing a
/// committed record and then passing because the write agreed with it.
///
/// The working copy other tests need is made in a temp directory instead, and
/// the assertion here is made against the source directly, because the source
/// is the only file in the pair that anyone maintains.
#[test]
fn test_data_files_are_consistent() {
    let manifest = std::path::Path::new(env!("CARGO_MANIFEST_DIR"));
    let src = manifest.join("docs/tests/subtotal-tiny.corro");

    // If a tracked `test_rec5.corro` is still around from before this test
    // stopped writing it, say so rather than quietly leaving it to rot: the
    // committed copy is byte-identical to the source today, so it is dead
    // weight that a future edit to one file would turn into a live diff.
    let stale = manifest.join("test_rec5.corro");
    if stale.exists() {
        let same = std::fs::read(&src).ok() == std::fs::read(&stale).ok();
        assert!(
            same,
            "{} differs from {} and is no longer written by any test, so \\
             nothing keeps it in step. Delete it, or point the tests that read \\
             it at the source.",
            stale.display(),
            src.display()
        );
    }

    let bytes = std::fs::read(&src)
        .unwrap_or_else(|e| panic!("read {}: {e}", src.display()));
    let text = String::from_utf8_lossy(&bytes);
    assert!(
        !text.contains("elloH"),
        "subtotal-tiny.corro must not contain 'elloH' (stale replayer output)"
    );
    assert!(
        !text.contains("estT"),
        "subtotal-tiny.corro must not contain 'estT' (stale replayer output)"
    );
}

#[cfg(feature = "ratatui")]
use std::path::PathBuf;

#[cfg(feature = "ratatui")]
#[test]
fn check_test_rec5_values() {
    let manifest = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let src = manifest.join("docs/tests/subtotal-tiny.corro");
    // Copy to a temporary file, so this test never touches a path in the
    // repo and cannot race with anything that reads the source.
    let tmp = tempfile::NamedTempFile::new().unwrap();
    std::fs::copy(&src, tmp.path()).unwrap_or_else(|e| {
        panic!("copy {} -> {}: {e}", src.display(), tmp.path().display())
    });
    let mut app = corro::ui::App::new(Some(tmp.path().to_path_buf()));
    app.load_initial().unwrap();
    let sheet = app.workbook.active_sheet().clone();
    let grid = &sheet.grid;
    
    let addr = corro::grid::CellAddr::Main { row: 0, col: 0 };
    eprintln!("A1 raw: {:?}", grid.get(&addr));
    eprintln!("A1 display: {:?}", corro::formula::cell_effective_display(grid, &addr));

    let addr2 = corro::grid::CellAddr::Main { row: 1, col: 0 };
    eprintln!("A2 raw: {:?}", grid.get(&addr2));
    eprintln!("A2 display: {:?}", corro::formula::cell_effective_display(grid, &addr2));

    eprintln!("Cursor: row={}, col={}", app.cursor.row, app.cursor.col);
    eprintln!("main_rows={} main_cols={}", grid.main_rows(), grid.main_cols());
    
    // Print col_width for key columns
    eprintln!("col_width(0)={}", grid.col_width(0));
    eprintln!("col_width(701)={}", grid.col_width(701));
    eprintln!("col_width(702)={}", grid.col_width(702));
    eprintln!("col_width(703)={}", grid.col_width(703));

    // Check rendered_width_for_column (ui_core version)
    let rw = corro::ui_core::rendered_width_for_column(grid, 702);
    eprintln!("rendered_width_for_column(702)={:?}", rw);
    
    let rw703 = corro::ui_core::rendered_width_for_column(grid, 703);
    eprintln!("rendered_width_for_column(703)={:?}", rw703);
    
    // Check content_width_for_column
    let cw = grid.content_width_for_column(702);
    eprintln!("content_width_for_column(702)={:?}", cw);
    
    // Check col_width_overrides
    eprintln!("has_override_702={}", grid.col_width_overrides().iter().any(|(c,_)| *c == 702));
    for (c, w) in grid.col_width_overrides() {
        if c >= 700 && c <= 710 {
            eprintln!("  override col={} width={}", c, w);
        }
    }

    // Simulate what fit_main_columns_to_max_width does:
    // It calls rendered_width_for_column and set_col_width
    for c in 0..grid.main_cols() {
        let global_col = corro::grid::MARGIN_COLS + c;
        if let Some(rw) = corro::ui_core::rendered_width_for_column(grid, global_col) {
            let capped = rw.min(grid.max_col_width());
            eprintln!("  fit col={} rw={} capped={}", global_col, rw, capped);
        }
    }
}
