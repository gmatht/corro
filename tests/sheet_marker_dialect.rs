//! Interop + migration tests for the `.corro` writer/reader.
//!
//! The canonical format is the `CORRO_LOG` op log, written by the *same*
//! function (`io::serialize_workbook_log`) for both the TUI and the GUI, so a
//! file saved by either backend must load in the other. The old `WORKBOOK`
//! snapshot dialect (which is where the stray `END_SHEET` came from) is kept
//! only as a read-compat path.

use corro::grid::{CellAddr, ColumnAddr, HEADER_ROWS, MARGIN_COLS};
use corro::ops::{WorkbookState, LOG_HEADER_PREFIX, LOG_VERSION};

fn main_cell(row: u32, col: u32) -> CellAddr {
    CellAddr::Main { row, col }
}

fn footer_seed() -> CellAddr {
    CellAddr::Footer {
        row: 0,
        col: ColumnAddr::Left(MARGIN_COLS - 1),
    }
}

fn header_seed() -> CellAddr {
    CellAddr::Header {
        row: (HEADER_ROWS - 1) as u32,
        col: ColumnAddr::Right(0),
    }
}

fn tmp_path(tag: &str) -> std::path::PathBuf {
    let p = std::env::temp_dir().join(format!(
        "corro_roundtrip_{}_{}.corro",
        std::process::id(),
        tag
    ));
    let _ = std::fs::remove_file(&p);
    p
}

fn save(path: &std::path::Path, wb: &WorkbookState) {
    corro::io::write_workbook_log(path, wb, &Default::default()).expect("save");
}

/// A GUI save must produce the TUI's format, and the TUI's loader must read
/// it back — that is the whole point of sharing one serializer.
#[test]
fn gui_save_is_readable_by_the_log_loader() {
    let path = tmp_path("gui_then_log");
    let mut wb = WorkbookState::new_seeded();
    wb.active_sheet_mut()
        .grid
        .set(&main_cell(0, 0), "42".into());
    save(&path, &wb);

    let text = std::fs::read_to_string(&path).unwrap();
    assert!(
        text.starts_with(&format!("{LOG_HEADER_PREFIX} {LOG_VERSION}\n")),
        "GUI saves must be a canonical CORRO_LOG, got:\n{text}"
    );
    assert!(
        !text.contains("WORKBOOK ") && !text.contains("END_SHEET"),
        "the snapshot dialect must be gone, got:\n{text}"
    );

    // The TUI's own replay path (log loader) must accept the file.
    let mut back = WorkbookState::new();
    let mut active = back.sheet_id(back.active_sheet);
    let (_off, replay) =
        corro::io::load_workbook_revisions_partial(&path, usize::MAX, &mut back, &mut active)
            .expect("load");
    assert_eq!(replay.failed_line, None, "log loader must not choke: {replay:?}");
    assert!(replay.op_count > 0, "ops should replay");
    assert_eq!(
        back.active_sheet().grid.get(&main_cell(0, 0)).as_deref(),
        Some("42")
    );
    let _ = std::fs::remove_file(&path);
}

/// Margin content (the built-in TOTAL seeds) must survive a save/load round
/// trip through the shared writer — this is the "missing TOTAL lines" bug.
#[test]
fn margin_total_seeds_round_trip() {
    let path = tmp_path("margins");
    save(&path, &WorkbookState::new_seeded());

    let text = std::fs::read_to_string(&path).unwrap();
    for label in ["[A_1", "]A~1"] {
        assert!(
            text.contains(label),
            "margin seed {label} missing from the log, got:\n{text}"
        );
    }

    let back = corro::io::load_workbook_file(&path).expect("reload");
    let grid = &back.active_sheet().grid;
    assert_eq!(
        grid.get(&footer_seed()).as_deref(),
        Some("TOTAL"),
        "footer TOTAL seed lost across save/load"
    );
    assert_eq!(
        grid.get(&header_seed()).as_deref(),
        Some("TOTAL"),
        "header TOTAL seed lost across save/load"
    );
    let _ = std::fs::remove_file(&path);
}

/// Multi-sheet round trip: sheet identity/titles and the active sheet survive.
#[test]
fn multi_sheet_round_trip() {
    let path = tmp_path("multisheet");
    let mut wb = WorkbookState::new();
    wb.sheets[0].title = "Alpha".into();
    wb.active_sheet_mut()
        .grid
        .set(&main_cell(0, 0), "one".into());
    wb.add_sheet("Beta".into(), corro::ops::SheetState::new(1, 1));
    // NOTE: add_sheet does not switch the active sheet; address Beta by its
    // record directly rather than through active_sheet.
    let beta = wb.sheets.iter().position(|s| s.title == "Beta").expect("Beta added");
    wb.sheets[beta].state.grid.set(&main_cell(1, 1), "two".into());
    save(&path, &wb);

    let back = corro::io::load_workbook_file(&path).expect("reload");
    assert_eq!(back.sheets.len(), 2, "both sheets survive");
    assert_eq!(back.sheets[0].title, "Alpha");
    assert_eq!(back.sheets[1].title, "Beta");
    assert_eq!(
        back.sheets[1].state.grid.get(&main_cell(1, 1)).as_deref(),
        Some("two")
    );
    let _ = std::fs::remove_file(&path);
}

/// Legacy read-compat: a file written by an older GUI build (snapshot dialect
/// with `END_SHEET`) still opens, including content appended after the marker.
#[test]
fn legacy_snapshot_files_still_open() {
    let path = tmp_path("legacy");
    std::fs::write(
        &path,
        "\
WORKBOOK 2 1
SHEET 1 Sheet1
VOLATILE_SEED 0
SET A1 1
SET A2 2
SET A3 3
END_SHEET
SET B1 f
SET B2 f
SET B3 a
SET B1 3
SET B2 4
SET B3 5
SET A4 1
SET A5 2
",
    )
    .unwrap();
    let wb = corro::io::load_workbook_file(&path).expect("legacy file must open");
    let grid = &wb.active_sheet().grid;
    for (addr, want) in [
        (main_cell(0, 0), "1"),
        (main_cell(0, 1), "3"),
        (main_cell(1, 1), "4"),
        (main_cell(2, 1), "5"),
        (main_cell(3, 0), "1"),
        (main_cell(4, 0), "2"),
    ] {
        assert_eq!(
            grid.get(&addr).as_deref(),
            Some(want),
            "legacy content dropped at {addr:?}"
        );
    }
    let _ = std::fs::remove_file(&path);
}
