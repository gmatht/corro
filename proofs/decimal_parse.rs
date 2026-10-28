// PROOF: the accept/reject shape of corro's decimal-literal parser.
//
// Checked by:
//   /opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/decimal_parse.rs
// Result: 2 verified, 0 errors
//
// `src/formula/number.rs:462 parse_number_literal` decides whether a cell's
// text is a number, and `parse_decimal_rational` (:495) is what it delegates
// the numeric value to. Proved here is the part with no `BigInt` in it: **which
// bodies are accepted**, i.e. non-empty, at most one `.`, every byte a digit or
// `.`.
//
// A quirk worth naming, because a reader will assume otherwise: `"."` and
// `"-."` are ACCEPTED and parse as zero. Neither side of the `.` being empty is
// rejected, and `parse_decimal_rational` turns two empty parts into 0.
// `tests/decimal_parse_xcheck.rs` pins this against the real parser.
//
// NOT proved: the numeric value (that needs `BigInt`/`BigRational`), the `trim`
// and the complex-literal pre-pass (`strip_prefix`, `split_once`), and
// `rational_to_formula_literal`. So there is no round-trip here.
use vstd::prelude::*;
verus! {

pub open spec fn is_digit(b: u8) -> bool { 48 <= b <= 57 }
pub open spec fn is_dot(b: u8) -> bool { b == 46 }

/// At most one `.` in the sequence.
pub open spec fn at_most_one_dot(s: Seq<u8>) -> bool {
    forall|i: int, j: int| 0 <= i < j < s.len()
        ==> !(#[trigger] s[i] == 46u8 && #[trigger] s[j] == 46u8)
}

/// Every byte is a digit or a `.`.
pub open spec fn digits_or_dots(s: Seq<u8>) -> bool {
    forall|i: int| 0 <= i < s.len() ==> is_digit(#[trigger] s[i]) || is_dot(s[i])
}

/// The accepted language of the decimal body (after any leading `-`).
pub open spec fn accepted_body(s: Seq<u8>) -> bool {
    s.len() > 0 && at_most_one_dot(s) && digits_or_dots(s)
}

/// The executable check, mirroring the parser's scan.
fn accepts_body(s: &[u8]) -> (r: bool)
    ensures r == accepted_body(s@),
{
    if s.len() == 0 {
        return false;
    }
    let mut i: usize = 0;
    let mut seen_dot: bool = false;
    while i < s.len()
        invariant
            i <= s.len(),
            digits_or_dots(s@.subrange(0, i as int)),
            at_most_one_dot(s@.subrange(0, i as int)),
            seen_dot ==> (exists|j: int| 0 <= j < i as int && #[trigger] s@[j] == 46u8),
            !seen_dot ==> (forall|j: int| 0 <= j < i as int ==> #[trigger] s@[j] != 46u8),
        decreases s.len() - i
    {
        let b = s[i];
        if b == 46u8 {
            if seen_dot {
                return false;
            }
            seen_dot = true;
        } else if !(b >= 48 && b <= 57) {
            return false;
        }
        i = i + 1;
    }
    proof {
        assert(s@.subrange(0, i as int) == s@);
    }
    true
}

} // verus
fn main() {}
