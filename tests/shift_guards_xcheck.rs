//! Cross-check: the proved shift guards against corro's real row/column insert
//! repair, reached through `translate_formula_text_insert_main_rows`.
//!
//! `proofs/shift_guards.rs` proves `shift_if_unlocked_row`/`_col`: a locked
//! coordinate is returned unchanged, one above the gap is unchanged, one at or
//! below it moves by exactly `delta`, and `None` exactly on overflow.
//!
//! Note the coordinate system: `gap_start` is a **0-based** main-row index,
//! while formula text is 1-based, so `A5` is index 4 and shifts only when
//! `gap_start <= 4`.
//!
//! This test writes nothing.

use corro::formula::translate_formula_text_insert_main_rows as ins_rows;

fn ins(raw: &str, gap: u32, delta: i32) -> Option<String> {
    ins_rows(raw, gap, delta, None)
}

/// Refs above the gap are untouched.
#[test]
fn refs_above_the_gap_do_not_move() {
    assert_eq!(ins("=A1", 4, 2), Some("=A1".to_string()));
    assert_eq!(ins("=A4", 4, 2), Some("=A4".to_string()));
}

/// Refs at or below the gap move by exactly delta.
#[test]
fn refs_at_or_below_the_gap_shift() {
    assert_eq!(ins("=A5", 4, 2), Some("=A7".to_string()));
    assert_eq!(ins("=A6", 4, 2), Some("=A8".to_string()));
    // and in the other direction
    assert_eq!(ins("=A7", 4, -2), Some("=A5".to_string()));
}

/// `$` locks fix the relevant axis, which is the "locked unchanged" property.
#[test]
fn dollar_locks_are_preserved() {
    assert_eq!(ins("=$A5", 4, 2), Some("=$A7".to_string()), "column lock only");
    assert_eq!(ins("=A$5", 4, 2), Some("=A$5".to_string()), "row lock fixes the row");
    assert_eq!(ins("=$A$5", 4, 2), Some("=$A$5".to_string()), "both locked");
}

/// Every ref inside an expression obeys the same rule.
#[test]
fn every_ref_in_an_expression_shifts() {
    // A5 is index 4, which IS the gap, so it shifts too
    assert_eq!(ins("=A5+A6", 4, 2), Some("=(A7+A8)".to_string()));
    assert_eq!(ins("=SUM(A1:A6)", 4, 2), Some("=SUM(A1:A8)".to_string()));
}

/// Zero delta is the identity, matching the guard's `row < gap_start` branch.
#[test]
fn zero_delta_is_identity() {
    assert_eq!(ins("=A5+A6", 4, 0), Some("=A5+A6".to_string()));
}
