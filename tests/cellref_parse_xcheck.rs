//! Cross-check: the proved consumed-length contract against the real
//! `addr::parse_cell_ref_at`.
//!
//! `proofs/cellref_parse.rs` proves that for a plain `A1`-style ref the reported
//! length covers exactly the letters and digits, and the byte after it is not a
//! digit -- so callers that slice at that offset land on a token boundary.
//!
//! The proof covers the plain branch only; the `$` locks and the `~` / `_`
//! header/footer markers are checked here.
//!
//! This test writes nothing.

use corro::addr::parse_cell_ref_at;

fn consumed(s: &str) -> Option<usize> {
    parse_cell_ref_at(s, 100).map(|(_, _, n)| n)
}

/// The reported length stops exactly before the next non-digit.
#[test]
fn consumed_length_stops_at_a_token_boundary() {
    for (src, want) in [
        ("A1", 2),
        ("A5rest", 2),
        ("A5+A6", 2),
        ("A5:A6", 2),
        ("Z9", 2),
        ("AB12", 4),
        ("A5x", 2),
        ("A5 ", 2),
    ] {
        assert_eq!(consumed(src), Some(want), "wrong consumed length for {src:?}");
        let n = want;
        let rest = &src[n..];
        assert!(
            rest.is_empty() || !rest.starts_with(|c: char| c.is_ascii_digit()),
            "{src:?} left a digit behind at offset {n}: {rest:?}"
        );
    }
}

/// Dollar locks are part of the consumed length, so a caller does not re-parse
/// them.
#[test]
fn dollar_locks_are_consumed() {
    assert_eq!(consumed("$A5"), Some(3));
    assert_eq!(consumed("A$5"), Some(3));
    assert_eq!(consumed("$A$5"), Some(4));
    assert_eq!(consumed("$A$5+1"), Some(4));
}

/// Header/footer markers are consumed with their digits.
#[test]
fn header_footer_markers_are_consumed() {
    assert_eq!(consumed("A~1"), Some(3));
    assert_eq!(consumed("A_1"), Some(3));
    assert_eq!(consumed("A~1x"), Some(3));
    assert_eq!(consumed("AA~9999"), Some(7));
}

/// Bracketed margin refs are consumed including the bracket.
#[test]
fn margin_brackets_are_consumed() {
    assert_eq!(consumed("[AA5"), Some(4));
    assert_eq!(consumed("]AA5"), Some(4));
    assert_eq!(consumed("[A1"), Some(3));
}

/// Malformed input is rejected rather than consuming a wrong length.
#[test]
fn malformed_inputs_are_rejected() {
    for s in ["A", "$5", "A0", "A~0", "", "5", "[", "$"] {
        assert_eq!(parse_cell_ref_at(s, 100), None, "{s:?} should be rejected");
    }
}

/// A `$` only starts a lock when the right kind of byte follows it: a letter for
/// the column lock, a digit for the row lock.
#[test]
fn dollar_locks_require_the_right_following_byte() {
    // "$AX5" is a column lock ($ before 'A'), columns "AX", row "5": all 4 bytes.
    let (_, locks, n) = parse_cell_ref_at("$AX5", 100).expect("$AX5 should parse");
    assert!(locks.col_absolute, "the leading $ locks the column");
    assert!(!locks.row_absolute, "no row lock here");
    assert_eq!(n, 4);

    // "A$X5": the '$' is before 'X', not a digit, so it is NOT a row lock and
    // the row parse sees no digits at all -- rejected.
    assert_eq!(parse_cell_ref_at("A$X5", 100), None);

    // "A$5" is a row lock.
    let (_, locks, n) = parse_cell_ref_at("A$5", 100).expect("A$5 should parse");
    assert!(locks.row_absolute);
    assert!(!locks.col_absolute);
    assert_eq!(n, 3);
}
