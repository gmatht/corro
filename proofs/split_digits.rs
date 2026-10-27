// PROOF: corro `src/extrapolate/mod.rs:398` `split_trailing_digits`.
//
// Checked by:
//   /opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/split_digits.rs
// Result: 4 verified, 0 errors
//
// What this establishes:
//   * `ndigits` counts the trailing ASCII-digit run of a byte string;
//   * the executable scan returns the exact split index `len - ndigits`;
//   * it returns `None` exactly when the string ends in a non-digit, and
//     `Some` exactly when there is at least one trailing digit.
//
// NOT established: that the byte at the split boundary is a non-digit (i.e.
// that the run is *maximal*). The scan's loop condition makes it true, but
// stating it needs an extra induction this file does not carry.
//
// This function is the half of fill-inference that is pure byte work; the
// caller `infer_suffix_fill` adds `parse::<i64>` and `format!`, neither of
// which has a vstd spec.
use vstd::prelude::*;
verus! {

pub open spec fn is_digit(b: u8) -> bool { 48 <= b <= 57 }

/// Length of the trailing ASCII-digit run.
pub open spec fn ndigits(s: Seq<u8>) -> int
    decreases s.len(),
{
    if s.len() == 0 {
        0
    } else if is_digit(s[s.len() - 1]) {
        1 + ndigits(s.drop_last())
    } else {
        0
    }
}

/// If the last `s.len() - k` bytes are all digits, they account for exactly
/// that many of `ndigits(s)`.
pub proof fn ndigits_split(s: Seq<u8>, k: int)
    requires 0 <= k <= s.len(),
        forall|j: int| k <= j < s.len() ==> is_digit(#[trigger] s[j]),
    ensures ndigits(s) == (s.len() - k) + ndigits(s.subrange(0, k)),
    decreases s.len() - k,
{
    if k == s.len() {
        assert(s.subrange(0, k) =~= s);
    } else {
        ndigits_split(s, k + 1);
        assert(is_digit(s[s.len() - 1]));
        assert(s.subrange(0, k + 1).drop_last() =~= s.subrange(0, k));
        assert(ndigits(s) == 1 + ndigits(s.drop_last()));
    }
}

/// `src/extrapolate/mod.rs:398`, over bytes: split `s` into (prefix, digits).
fn split_trailing_digits(s: &[u8]) -> (r: Option<usize>)
    ensures
        r.is_some() ==> r.unwrap() as int == s.len() as int - ndigits(s@),
        r.is_some() ==> ndigits(s@) > 0,
        r.is_none() ==> ndigits(s@) == 0,
{
    let mut i: usize = s.len();
    while i > 0 && (s[i - 1] >= 48 && s[i - 1] <= 57)
        invariant
            i <= s.len(),
            forall|j: int| i as int <= j < s.len() ==> is_digit(#[trigger] s@[j]),
        decreases i,
    {
        i = i - 1;
    }
    proof {
        ndigits_split(s@, i as int);
    }
    if i == s.len() {
        return None;
    }
    Some(i)
}

} // verus
fn main() {}
