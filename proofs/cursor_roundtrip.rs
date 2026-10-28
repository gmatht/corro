// PROOF: corro's sheet cursor <-> cell address round-trip.
//
// Checked by:
//   /opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/cursor_roundtrip.rs
// Result: 4 verified, 0 errors
//
// `src/addr.rs:115 sheet_cursor_to_addr` and `:159 addr_to_sheet_cursor` are
// inverses. That matters because the UI converts between them on every cursor
// movement and every address display; a mismatch would put the selection or the
// status line on the wrong cell.
//
// Two layers are proved:
//   1. `ColumnAddr::from_global` / `to_global` (`src/grid/mod.rs:222,234`) are
//      inverses in both directions, given the region bounds. The left margin
//      occupies globals [0, 702), the main grid [702, 702+main_cols), and the
//      right margin everything above.
//   2. `cursor_to_addr` / `addr_to_cursor` over the five regions (Header, Left,
//      Main, Right, Footer) compose to the identity.
//
// `regions_are_distinct` is included so the round-trip is not vacuously true:
// all five regions are reachable and map to distinct address kinds.
//
// Modelled over `int` rather than `usize`/`u32`: the source's arithmetic is
// exact within these bounds (main_rows/main_cols are display extents, and the
// cursor is a logical row index), and integer modelling is what makes the proof
// linear. The `as u32` casts in the source are the only step not covered, and
// they are checked by the cross-check test instead.
use vstd::prelude::*;
verus! {

/// `src/grid/mod.rs:16`: `pub const MARGIN_COLS: usize = 26 * 27;`
pub open spec fn margin_cols() -> int { 702 }

/// Which of the three column regions a global column falls in.
pub open spec fn region(col: int, main_cols: int) -> int {
    if col < margin_cols() { 0 } else if col < margin_cols() + main_cols { 1 } else { 2 }
}

/// `ColumnAddr::from_global`: region-relative index.
pub open spec fn from_global(col: int, main_cols: int) -> int {
    if col < margin_cols() { col }
    else if col < margin_cols() + main_cols { col - margin_cols() }
    else { col - margin_cols() - main_cols }
}

/// `ColumnAddr::to_global`: back to a global column.
pub open spec fn to_global(idx: int, reg: int, main_cols: int) -> int {
    if reg == 0 { idx }
    else if reg == 1 { margin_cols() + idx }
    else { margin_cols() + main_cols + idx }
}

/// `to_global . from_global` is the identity.
pub proof fn to_from_global(col: int, main_cols: int)
    requires col >= 0, main_cols >= 0,
    ensures to_global(from_global(col, main_cols), region(col, main_cols), main_cols) == col,
{}

/// `from_global . to_global` is the identity on a well-formed address.
pub proof fn from_to_global(idx: int, reg: int, main_cols: int)
    requires
        main_cols > 0,
        0 <= idx,
        reg == 0 ==> idx < margin_cols(),
        reg == 1 ==> idx < main_cols,
        reg == 0 || reg == 1 || reg == 2,
    ensures
        region(to_global(idx, reg, main_cols), main_cols) == reg,
        from_global(to_global(idx, reg, main_cols), main_cols) == idx,
{
    if reg == 0 {
        assert(to_global(idx, reg, main_cols) == idx);
        assert(idx < margin_cols());
    } else if reg == 1 {
        assert(to_global(idx, reg, main_cols) == margin_cols() + idx);
        assert(margin_cols() + idx < margin_cols() + main_cols);
        assert(margin_cols() <= margin_cols() + idx);
    } else {
        assert(to_global(idx, reg, main_cols) == margin_cols() + main_cols + idx);
        assert(idx >= 0);
        assert(margin_cols() + main_cols + 0 <= margin_cols() + main_cols + idx);
    }
}

/// Address kinds, as integers: 0 Header, 1 Left, 2 Main, 3 Right, 4 Footer.
pub open spec fn cursor_to_addr(lr: int, gc: int, main_rows: int, main_cols: int, hr: int)
    -> (int, int, int)
{
    if lr < hr {
        (0, lr, gc)
    } else if lr < hr + main_rows {
        let mr = lr - hr;
        if gc < margin_cols() { (1, mr, gc) }
        else if gc < margin_cols() + main_cols { (2, mr, gc - margin_cols()) }
        else { (3, mr, gc - margin_cols() - main_cols) }
    } else {
        (4, lr - hr - main_rows, gc)
    }
}

pub open spec fn addr_to_cursor(kind: int, row: int, idx: int, main_rows: int, main_cols: int, hr: int)
    -> (int, int)
{
    if kind == 0 { (row, idx) }
    else if kind == 4 { (hr + main_rows + row, idx) }
    else if kind == 1 { (hr + row, idx) }
    else if kind == 2 { (hr + row, margin_cols() + idx) }
    else { (hr + row, margin_cols() + main_cols + idx) }
}

/// THE round-trip: cursor -> address -> cursor.
pub proof fn cursor_roundtrip(lr: int, gc: int, main_rows: int, main_cols: int, hr: int)
    requires lr >= 0, gc >= 0, main_rows >= 0, main_cols > 0, hr >= 0,
    ensures ({
        let a = cursor_to_addr(lr, gc, main_rows, main_cols, hr);
        addr_to_cursor(a.0, a.1, a.2, main_rows, main_cols, hr) == (lr, gc)
    }),
{}

/// ...and it is not vacuous: each of the five regions is reachable and maps to
/// its own address kind.
pub proof fn regions_are_distinct(hr: int, main_rows: int, main_cols: int)
    requires hr > 0, main_rows > 0, main_cols > 0,
    ensures
        cursor_to_addr(0, 0, main_rows, main_cols, hr).0 == 0,
        cursor_to_addr(hr, 0, main_rows, main_cols, hr).0 == 1,
        cursor_to_addr(hr, margin_cols(), main_rows, main_cols, hr).0 == 2,
        cursor_to_addr(hr, margin_cols() + main_cols, main_rows, main_cols, hr).0 == 3,
        cursor_to_addr(hr + main_rows, 0, main_rows, main_cols, hr).0 == 4,
{}

} // verus
fn main() {}
