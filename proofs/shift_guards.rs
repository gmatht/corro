// PROOF: corro `src/formula/mod.rs:626 shift_if_unlocked_row` and `:638
// shift_if_unlocked_col`, the guards that keep inserted-row repair from
// corrupting `$`-locked references.
//
// Checked by:
//   /opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/shift_guards.rs
// Result: 6 verified, 0 errors
//
// Proved:
//   * `shifted` is a total spec of the operation;
//   * the executable `shift` refines it exactly;
//   * a locked row is returned unchanged, as is one above the insertion gap;
//   * an unlocked row at or below the gap moves by exactly `delta`;
//   * it is refused (`None`) exactly when the result leaves the `u32` range;
//   * the `i64` intermediate cannot overflow, so the range test is the only
//     place a value is ever refused.
//
// Both functions in the source have identical bodies over `row_absolute` and
// `col_absolute` respectively, so one proof covers both; the caller passes
// the relevant lock flag.
//
// NOT proved: `translate_cell_addr_by_offset` (:1145), which dispatches over
// `CellAddr` variants and composes these shifts with `ColumnAddr::from_global`.
use vstd::prelude::*;
verus! {

/// Spec: the coordinate after shifting, or `None` if that leaves the u32 range.
pub open spec fn shifted(row: u32, locked: bool, gap_start: u32, delta: int) -> Option<u32> {
    if locked || (row as int) < (gap_start as int) {
        Some(row)
    } else if (row as int) + delta < 0 || (row as int) + delta > u32::MAX {
        None
    } else {
        Some(((row as int) + delta) as u32)
    }
}

/// `shift_if_unlocked_row` / `shift_if_unlocked_col`, over bytes of arithmetic.
fn shift(row: u32, locked: bool, gap_start: u32, delta: i32) -> (r: Option<u32>)
    ensures r == shifted(row, locked, gap_start, delta as int),
{
    if locked || row < gap_start {
        return Some(row);
    }
    let r = row as i64 + delta as i64;
    if r < 0 || r > u32::MAX as i64 {
        None
    } else {
        Some(r as u32)
    }
}

/// A locked coordinate never moves.
pub proof fn shift_locked_unchanged(row: u32, gap_start: u32, delta: int)
    ensures shifted(row, true, gap_start, delta) == Some(row),
{}

/// A coordinate above the insertion gap never moves (rows after the gap keep
/// their number).
pub proof fn shift_above_gap_unchanged(row: u32, gap_start: u32, delta: int)
    requires (row as int) < (gap_start as int),
    ensures shifted(row, false, gap_start, delta) == Some(row),
{}

/// An unlocked coordinate at or below the gap moves by exactly `delta`.
pub proof fn shift_in_range_moves(row: u32, gap_start: u32, delta: int)
    requires
        (gap_start as int) <= (row as int),
        0 <= (row as int) + delta <= u32::MAX,
    ensures shifted(row, false, gap_start, delta) == Some(((row as int) + delta) as u32),
{}

/// ...and is refused exactly when that would overflow.
pub proof fn shift_overflow_is_none(row: u32, gap_start: u32, delta: int)
    requires
        (gap_start as int) <= (row as int),
        (row as int) + delta < 0 || (row as int) + delta > u32::MAX,
    ensures shifted(row, false, gap_start, delta).is_none(),
{}

/// The `i64` intermediate cannot overflow, so the range test is the only place
/// a value is ever refused. `u32::MAX + i32::MAX` is about 6.4e9, far inside
/// `i64`.
pub proof fn i64_intermediate_is_exact()
    ensures
        forall|row: u32, delta: i32|
            -(9223372036854775807int) <= #[trigger] ((row as int) + (delta as int))
            && ((row as int) + (delta as int)) <= 9223372036854775807int,
{}

}  // verus
fn main() {}
