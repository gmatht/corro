//! Cross-check: the Verus-verified trailing-digit splitter against corro's
//! fill-inference, which is the only place it is reachable from outside.
//!
//! `proofs/split_digits.rs` proves, for its own copy of `split_trailing_digits`,
//! that the scan returns exactly `len - ndigits` and returns `None` iff there
//! is no trailing digit. That function is private, so the check has to go
//! through its public caller `infer_fill_value`.
//!
//! This test writes nothing.

use corro::extrapolate::{infer_fill_value, FillDirection};

fn fill(seed: &[&str]) -> Option<String> {
    let seed: Vec<String> = seed.iter().map(|s| s.to_string()).collect();
    infer_fill_value(&seed, 1, FillDirection::Right, 100)
}

/// The splitter is reachable only when a seed ends in digits; a seed with no
/// trailing digit falls through to the "repeat the last value" rule.
#[test]
fn no_trailing_digit_falls_through_to_repeat() {
    assert_eq!(fill(&["abc"]), Some("abc".to_string()));
    assert_eq!(fill(&["abc", "def"]), Some("def".to_string()));
}

/// One seed with digits: the splitter finds prefix "item" and digits "1", and
/// with no previous value the step is zero, so the value is reproduced with
/// the same width.
#[test]
fn single_seed_reproduces_itself() {
    assert_eq!(fill(&["item1"]), Some("item1".to_string()));
    assert_eq!(fill(&["A01"]), Some("A01".to_string()));
}

/// Two seeds: the step is `last - prev`, applied once.
#[test]
fn two_seeds_continue_the_sequence() {
    assert_eq!(fill(&["item1", "item2"]), Some("item3".to_string()));
    assert_eq!(fill(&["A01", "A02"]), Some("A03".to_string()));
    // width is preserved from the last seed
    assert_eq!(fill(&["A01", "A03"]), Some("A05".to_string()));
    // non-unit step
    assert_eq!(fill(&["x9", "x11"]), Some("x13".to_string()));
}

/// The split boundary is where the digits start: "A01" splits into "A" and
/// "01", so the zero padding is preserved rather than parsed away.
#[test]
fn split_boundary_preserves_zero_padding() {
    assert_eq!(fill(&["A01", "A02"]), Some("A03".to_string()));
    assert_eq!(fill(&["v0098", "v0099"]), Some("v0100".to_string()));
}

/// A seed whose *prefixes* differ is rejected: the splitter's prefix becomes
/// part of the contract when `infer_suffix_fill` compares them.
#[test]
fn differing_prefixes_do_not_continue() {
    // "a1" and "b2" have different prefixes, so suffix inference must decline.
    // What it falls back to is the repeat rule, i.e. the last seed.
    assert_eq!(fill(&["a1", "b2"]), Some("b2".to_string()));
}
