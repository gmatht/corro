//! Cross-check: the proved keyword decision table against the real
//! `ops::margin_key_agg_func`.
//!
//! `proofs/keyword_table.rs` proves the precedence rules over bytes. The real
//! function additionally trims and upper-cases first, and those are `str`
//! methods without vstd specs, so this test is what checks the composed
//! behaviour.
//!
//! This test writes nothing.

use corro::ops::{margin_key_agg_func, AggFunc};

fn kw(s: &str) -> Option<AggFunc> {
    margin_key_agg_func(s)
}

/// The bare keyword table.
#[test]
fn bare_keywords() {
    assert_eq!(kw("SUM"), Some(AggFunc::Sum));
    assert_eq!(kw("TOTAL"), Some(AggFunc::Sum));
    assert_eq!(kw("MEAN"), Some(AggFunc::Mean));
    assert_eq!(kw("AVERAGE"), Some(AggFunc::Mean));
    assert_eq!(kw("AVG"), Some(AggFunc::Mean));
    assert_eq!(kw("MEDIAN"), Some(AggFunc::Median));
    assert_eq!(kw("MIN"), Some(AggFunc::Min));
    assert_eq!(kw("MINIMUM"), Some(AggFunc::Min));
    assert_eq!(kw("MAX"), Some(AggFunc::Max));
    assert_eq!(kw("MAXIMUM"), Some(AggFunc::Max));
    assert_eq!(kw("COUNT"), Some(AggFunc::Count));
}

/// Case-insensitive, and trimmed.
#[test]
fn case_and_whitespace_insensitive() {
    for s in ["sum", "Sum", "sUM", "  SUM  ", "\tSUM\n"] {
        assert_eq!(kw(s), Some(AggFunc::Sum), "{s:?} should be SUM");
    }
}

/// Nothing outside the table is accepted.
#[test]
fn non_keywords_rejected() {
    for s in ["SUM1", "", " ", "AVERAGES", "TOTALS", "MINIMUMS", "xyz"] {
        assert_eq!(kw(s), None, "{s:?} should be rejected");
    }
}

/// A leading apostrophe makes the label literal text, whatever follows --
/// this is the `quoted_is_never_an_aggregate` property.
#[test]
fn apostrophe_forces_literal() {
    for s in ["'SUM", "'=SUM", "''SUM", "'  SUM", "'=TOTAL", "'==MIN"] {
        assert_eq!(kw(s), None, "{s:?} is quoted and must never be an aggregate");
    }
}

/// A single `=` accepts ONLY `TOTAL`. This is the property that keeps `=MIN`
/// and `=SUM` usable as ordinary formulas.
#[test]
fn single_equals_accepts_only_total() {
    assert_eq!(kw("=TOTAL"), Some(AggFunc::Sum));
    assert_eq!(kw("=total"), Some(AggFunc::Sum));
    for s in ["=SUM", "=MIN", "=MAX", "=AVG", "=MEAN", "=COUNT", "=MEDIAN", "=", "=="] {
        assert_eq!(kw(s), None, "{s:?} must stay a formula");
    }
}

/// `==WORD` and a bare `WORD` behave identically -- `double_eq_matches_bare`.
#[test]
fn double_equals_matches_bare() {
    for w in ["SUM", "TOTAL", "MEAN", "AVERAGE", "AVG", "MEDIAN", "MIN", "MINIMUM", "MAX", "MAXIMUM", "COUNT"] {
        assert_eq!(
            kw(&format!("=={w}")),
            kw(w),
            "=={w} should match bare {w}"
        );
    }
    // and non-keywords stay rejected in both forms
    assert_eq!(kw("==NOPE"), None);
    assert_eq!(kw("NOPE"), None);
}
