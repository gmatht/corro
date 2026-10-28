//! Cross-check: the proved cursor <-> address round-trip against the real
//! `addr::sheet_cursor_to_addr` / `addr::addr_to_sheet_cursor`.
//!
//! `proofs/cursor_roundtrip.rs` proves the composition is the identity over the
//! five regions, and that `ColumnAddr::from_global`/`to_global` invert each
//! other. The proof models the arithmetic over `int`; the `as u32` casts in the
//! source are outside it, so this test is what covers them.
//!
//! The existing `cursor_addr_roundtrip_across_regions` in `src/addr.rs` checks
//! one address per region. This sweeps the *boundaries*, which is where an
//! off-by-one would live: `margin_cols - 1` / `margin_cols`, and
//! `margin_cols + main_cols - 1` / `margin_cols + main_cols`.
//!
//! This test writes nothing.

use corro::addr::{addr_to_sheet_cursor, sheet_cursor_to_addr, GlobalCol, LogicalRow, MainCols, MainRows};
use corro::grid::{CellAddr, ColumnAddr, HEADER_ROWS, MARGIN_COLS};

const MAIN_ROWS: usize = 3;
const MAIN_COLS: usize = 4;

fn roundtrip(addr: &CellAddr) -> CellAddr {
    let (r, c) = addr_to_sheet_cursor(addr, MainRows(MAIN_ROWS), MainCols(MAIN_COLS));
    sheet_cursor_to_addr(r, c, MainRows(MAIN_ROWS), MainCols(MAIN_COLS))
}

/// Every address kind round-trips through the cursor.
#[test]
fn all_five_regions_roundtrip() {
    let cases = [
        CellAddr::Header { row: 0, col: ColumnAddr::Left(0) },
        CellAddr::Header { row: 7, col: ColumnAddr::Main(2) },
        CellAddr::Left { col: 0, row: 0 },
        CellAddr::Left { col: MARGIN_COLS - 1, row: 1 },
        CellAddr::Main { row: 0, col: 0 },
        CellAddr::Main { row: 2, col: 3 },
        CellAddr::Right { col: 0, row: 2 },
        CellAddr::Right { col: 9, row: 0 },
        CellAddr::Footer { row: 0, col: ColumnAddr::Right(0) },
        CellAddr::Footer { row: 5, col: ColumnAddr::Left(3) },
    ];
    for addr in cases {
        assert_eq!(roundtrip(&addr), addr, "{addr:?} did not round-trip");
    }
}

/// The column-region boundaries, where an off-by-one would live: the last left
/// margin column, the first main column, the last main column, the first right
/// margin column.
#[test]
fn column_region_boundaries_roundtrip() {
    let boundary_globals = [
        0,
        MARGIN_COLS - 1,
        MARGIN_COLS,
        MARGIN_COLS + MAIN_COLS - 1,
        MARGIN_COLS + MAIN_COLS,
        MARGIN_COLS + MAIN_COLS + 1,
    ];
    for g in boundary_globals {
        let expect_left = g < MARGIN_COLS;
        let expect_main = g >= MARGIN_COLS && g < MARGIN_COLS + MAIN_COLS;
        let addr = sheet_cursor_to_addr(
            LogicalRow(HEADER_ROWS), // a main-band row, so the column decides
            GlobalCol(g),
            MainRows(MAIN_ROWS),
            MainCols(MAIN_COLS),
        );
        match addr {
            CellAddr::Left { .. } => assert!(expect_left, "global {g} should not be Left"),
            CellAddr::Main { .. } => assert!(expect_main, "global {g} should not be Main"),
            CellAddr::Right { .. } => {
                assert!(!expect_left && !expect_main, "global {g} should not be Right")
            }
            _ => panic!("main-band row {g} produced {addr:?}"),
        }
        assert_eq!(roundtrip(&addr), addr, "global column {g} did not round-trip");
    }
}

/// The row-region boundaries: last header row / first main row, last main row /
/// first footer row.
#[test]
fn row_region_boundaries_roundtrip() {
    let boundary_rows = [
        0,
        HEADER_ROWS - 1,
        HEADER_ROWS,
        HEADER_ROWS + MAIN_ROWS - 1,
        HEADER_ROWS + MAIN_ROWS,
        HEADER_ROWS + MAIN_ROWS + 1,
    ];
    for r in boundary_rows {
        let addr = sheet_cursor_to_addr(
            LogicalRow(r),
            GlobalCol(MARGIN_COLS), // a main column
            MainRows(MAIN_ROWS),
            MainCols(MAIN_COLS),
        );
        assert_eq!(roundtrip(&addr), addr, "logical row {r} did not round-trip");
    }
}

/// A cursor over the header band lands in a header address, and one in the
/// footer band in a footer address -- the `regions_are_distinct` claim.
#[test]
fn cursor_bands_map_to_the_right_kind() {
    let hdr = sheet_cursor_to_addr(LogicalRow(0), GlobalCol(0), MainRows(MAIN_ROWS), MainCols(MAIN_COLS));
    assert!(matches!(hdr, CellAddr::Header { .. }), "got {hdr:?}");

    let ftr = sheet_cursor_to_addr(
        LogicalRow(HEADER_ROWS + MAIN_ROWS),
        GlobalCol(0),
        MainRows(MAIN_ROWS),
        MainCols(MAIN_COLS),
    );
    assert!(matches!(ftr, CellAddr::Footer { .. }), "got {ftr:?}");
}
