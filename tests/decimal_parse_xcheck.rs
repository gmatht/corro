//! Cross-check: the proved accept/reject shape against the real decimal parser.
//!
//! `proofs/decimal_parse.rs` proves the body language: non-empty, at most one
//! `.`, every byte a digit or `.`. The real path additionally trims and tries a
//! complex literal first, and `effective_numeric` is the public way to reach it.
//!
//! This test writes nothing.

use corro::formula::effective_numeric;
use corro::grid::{CellAddr, Grid, GridBox};

/// The verified scanner, transcribed back to Rust.
fn accepts_body(s: &[u8]) -> bool {
    if s.is_empty() {
        return false;
    }
    let mut seen_dot = false;
    for &b in s {
        if b == b'.' {
            if seen_dot {
                return false;
            }
            seen_dot = true;
        } else if !b.is_ascii_digit() {
            return false;
        }
    }
    true
}

fn numeric(cell: &str) -> bool {
    let mut g = GridBox::new(Grid::new(3, 3));
    let addr = CellAddr::Main { row: 0, col: 0 };
    g.set(&addr, cell.to_string());
    let mut visiting = Vec::new();
    let mut budget = 1000usize;
    effective_numeric(&g, &addr, &mut visiting, &mut budget).is_some()
}

fn model(cell: &str) -> bool {
    let t = cell.trim();
    let body = t.strip_prefix('-').unwrap_or(t);
    accepts_body(body.as_bytes())
}

#[test]
fn plain_numbers_are_accepted() {
    for c in ["0", "1", "007", "-1", "1.5", "-1.5", ".5", "1.", "0.0", ".0", "-.5"] {
        assert!(numeric(c), "{c:?} should be numeric");
        assert_eq!(model(c), numeric(c), "model disagreed on {c:?}");
    }
}

#[test]
fn non_numbers_are_rejected() {
    for c in ["", "abc", "1e5", "+1", "--1", "1..2", "..", "-", "1.2.3", "1 _0"] {
        assert!(!numeric(c), "{c:?} should not be numeric");
        assert_eq!(model(c), numeric(c), "model disagreed on {c:?}");
    }
}

/// The quirk the proof's header calls out: a bare dot is accepted (as zero),
/// because neither side of the `.` being empty is rejected.
#[test]
fn a_lone_dot_parses_as_zero() {
    assert!(numeric("."), "'.' is accepted by the parser");
    assert!(numeric("-."), "'-.' is accepted by the parser");
    assert_eq!(model("."), true);
    assert_eq!(model("-."), true);
}

/// Trimming happens before the shape check, so surrounding space is irrelevant.
#[test]
fn surrounding_whitespace_is_trimmed() {
    for c in [" 12 ", "1 ", " 1", "\t1.5\n"] {
        assert!(numeric(c), "{c:?} should be numeric after trimming");
        assert_eq!(model(c), numeric(c), "model disagreed on {c:?}");
    }
}
