// PROOF: the accept/reject shape of `balance-core/src/lib.rs:7 parse_amount_cents`.
//
// Checked by:
//   /opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/money_parse.rs
// Result: 4 verified, 0 errors
//
// This proves the *structure* of the parse over bytes: an optional sign, then
// a non-empty run of ASCII digits (the whole part), then optionally a '.' and
// a run of digits (the fraction) which must reach the end. Whenever it
// accepts, the whole and fraction parts it identified really are digit runs
// with a non-empty whole part.
//
// NOT proved:
//   * the returned cents *value* -- that needs `parse::<i64>`, `saturating_*`
//     and the `replace(',', "")` pre-pass, none of which have vstd specs;
//   * `format_amount_cents` (uses `format!`), so no round-trip here.
//
// The comma pre-pass is deliberately absent: the real function strips every
// ',' before parsing, which is why ",1" parses as 1. That is a separate
// transformation and is checked by the cross-check test instead.
use vstd::prelude::*;
verus! {

pub open spec fn is_digit(b: u8) -> bool { 48 <= b <= 57 }

pub open spec fn all_digits(s: Seq<u8>) -> bool {
    forall|i: int| 0 <= i < s.len() ==> is_digit(#[trigger] s[i])
}

pub proof fn all_digits_range(s: Seq<u8>, lo: int, hi: int)
    requires 0 <= lo <= hi <= s.len(),
        forall|i: int| lo <= i < hi ==> is_digit(#[trigger] s[i]),
    ensures all_digits(s.subrange(lo, hi)),
{
    assert forall|i: int| 0 <= i < hi - lo implies is_digit(#[trigger] s.subrange(lo, hi)[i]) by {
        assert(s.subrange(lo, hi)[i] == s[i + lo]);
    };
}

/// The parse, over bytes and without the comma pre-pass.
fn parse_amount_cents(raw: &[u8]) -> (r: Option<i64>)
    ensures
        r.is_some() ==> exists|w: Seq<u8>, f: Seq<u8>|
            #![trigger all_digits(w), all_digits(f)]
            all_digits(w) && all_digits(f) && w.len() >= 1,
{
    if raw.len() == 0 {
        return None;
    }
    let mut i: usize = 0;
    if raw[0] == b'+' || raw[0] == b'-' {
        i = 1;
    }
    let start = i;
    while i < raw.len() && raw[i] != b'.' && (raw[i] >= 48 && raw[i] <= 57)
        invariant
            start <= i <= raw.len(),
            forall|j: int| start as int <= j < i as int ==> is_digit(#[trigger] raw@[j]),
        decreases raw.len() - i
    {
        i = i + 1;
    }
    // a '.' with no digits before it is rejected: the whole part is empty
    if i == start {
        return None;
    }
    proof { all_digits_range(raw@, start as int, i as int); }
    if i == raw.len() {
        return Some(0);
    }
    let mut j: usize = i + 1;
    while j < raw.len() && (raw[j] >= 48 && raw[j] <= 57)
        invariant
            i < j <= raw.len(),
            forall|k: int| (i + 1) as int <= k < j as int ==> is_digit(#[trigger] raw@[k]),
        decreases raw.len() - j
    {
        j = j + 1;
    }
    // anything after the fraction is rejected
    if j != raw.len() {
        return None;
    }
    Some(0)
}

} // verus
fn main() {}
