// PROOF: corro `src/ops/mod.rs:35 margin_key_agg_func` -- the margin-keyword
// decision table that decides whether a margin label is an aggregate.
//
// Checked by:
//   /opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/keyword_table.rs
// Result: 10 verified, 0 errors
//
// The function has four input forms, and the one that matters is the
// precedence between them. Proved here, over bytes:
//   * a leading apostrophe makes the label literal text -- never an aggregate,
//     whatever follows (`quoted_is_never_an_aggregate`);
//   * a single `=` prefix accepts ONLY `TOTAL` (`single_eq_accepts_only_total`),
//     so `=MIN` stays a formula;
//   * `==WORD` and a bare `WORD` share one table (`double_eq_matches_bare`);
//   * nothing outside the eleven literals is ever accepted
//     (`only_keywords_accepted`).
//
// The table itself: SUM/TOTAL -> Sum, MEAN/AVERAGE/AVG -> Mean, MEDIAN,
// MIN/MINIMUM, MAX/MAXIMUM, COUNT.
//
// NOT proved: the trimming and case folding (`str::trim`,
// `to_ascii_uppercase`) which the source applies before the table, and the
// `split_labeled_formula` pre-pass. Both are `str` methods without vstd specs;
// `tests/keyword_table_xcheck.rs` checks the composed behaviour on the real
// function instead.
use vstd::prelude::*;
verus! {

pub enum AggFunc { Sum, Mean, Median, Min, Max, Count }

/// Keyword table on an already upper-cased word.
pub open spec fn lookup(s: Seq<u8>) -> Option<AggFunc> {
    if s == seq![83u8, 85u8, 77u8] { Some(AggFunc::Sum) }                     // SUM
    else if s == seq![84u8, 79u8, 84u8, 65u8, 76u8] { Some(AggFunc::Sum) }    // TOTAL
    else if s == seq![77u8, 69u8, 65u8, 78u8] { Some(AggFunc::Mean) }         // MEAN
    else if s == seq![65u8, 86u8, 69u8, 82u8, 65u8, 71u8, 69u8] { Some(AggFunc::Mean) } // AVERAGE
    else if s == seq![65u8, 86u8, 71u8] { Some(AggFunc::Mean) }               // AVG
    else if s == seq![77u8, 69u8, 68u8, 73u8, 65u8, 78u8] { Some(AggFunc::Median) }
    else if s == seq![77u8, 73u8, 78u8] { Some(AggFunc::Min) }                // MIN
    else if s == seq![77u8, 73u8, 78u8, 73u8, 77u8, 85u8, 77u8] { Some(AggFunc::Min) } // MINIMUM
    else if s == seq![77u8, 65u8, 88u8] { Some(AggFunc::Max) }                // MAX
    else if s == seq![77u8, 65u8, 88u8, 73u8, 77u8, 85u8, 77u8] { Some(AggFunc::Max) } // MAXIMUM
    else if s == seq![67u8, 79u8, 85u8, 78u8, 84u8] { Some(AggFunc::Count) }  // COUNT
    else { None }
}

/// The four input forms, as an integer selector.
/// 0 = bare word, 1 = `==WORD`, 2 = `=WORD`, 3 = quoted.
pub open spec fn decides(mode: int, word: Seq<u8>) -> Option<AggFunc> {
    if mode == 0 || mode == 1 {
        lookup(word)
    } else if mode == 2 {
        // Legacy totals row: `=TOTAL` only.
        if word == seq![84u8, 79u8, 84u8, 65u8, 76u8] { Some(AggFunc::Sum) } else { None }
    } else {
        None
    }
}

pub open spec fn is_keyword(s: Seq<u8>) -> bool {
    s == seq![83u8, 85u8, 77u8]
    || s == seq![84u8, 79u8, 84u8, 65u8, 76u8]
    || s == seq![77u8, 69u8, 65u8, 78u8]
    || s == seq![65u8, 86u8, 69u8, 82u8, 65u8, 71u8, 69u8]
    || s == seq![65u8, 86u8, 71u8]
    || s == seq![77u8, 69u8, 68u8, 73u8, 65u8, 78u8]
    || s == seq![77u8, 73u8, 78u8]
    || s == seq![77u8, 73u8, 78u8, 73u8, 77u8, 85u8, 77u8]
    || s == seq![77u8, 65u8, 88u8]
    || s == seq![77u8, 65u8, 88u8, 73u8, 77u8, 85u8, 77u8]
    || s == seq![67u8, 79u8, 85u8, 78u8, 84u8]
}

/// A quoted label is never an aggregate, whatever follows the apostrophe.
pub proof fn quoted_is_never_an_aggregate(word: Seq<u8>)
    ensures decides(3, word).is_none(),
{}

/// A single `=` prefix accepts only `TOTAL`, so `=MIN` and `=SUM` stay formulas.
pub proof fn single_eq_accepts_only_total(word: Seq<u8>)
    ensures decides(2, word).is_some() <==> word == seq![84u8, 79u8, 84u8, 65u8, 76u8],
{}

/// `==WORD` and a bare `WORD` behave identically.
pub proof fn double_eq_matches_bare(word: Seq<u8>)
    ensures decides(1, word) == decides(0, word),
{}

/// Nothing outside the eleven literals is ever accepted.
pub proof fn only_keywords_accepted(s: Seq<u8>)
    ensures lookup(s).is_some() ==> is_keyword(s),
{
    if lookup(s).is_some() {
        assert(is_keyword(s));
    }
}

pub proof fn table_sum() ensures lookup(seq![83u8, 85u8, 77u8]) == Some(AggFunc::Sum), {}
pub proof fn table_mean() ensures lookup(seq![65u8, 86u8, 71u8]) == Some(AggFunc::Mean), {}
pub proof fn table_median() ensures lookup(seq![77u8, 69u8, 68u8, 73u8, 65u8, 78u8]) == Some(AggFunc::Median), {}
pub proof fn table_min() ensures lookup(seq![77u8, 73u8, 78u8]) == Some(AggFunc::Min), {}
pub proof fn table_max() ensures lookup(seq![77u8, 65u8, 88u8]) == Some(AggFunc::Max), {}
pub proof fn table_count() ensures lookup(seq![67u8, 79u8, 85u8, 78u8, 84u8]) == Some(AggFunc::Count), {}

} // verus
fn main() {}
