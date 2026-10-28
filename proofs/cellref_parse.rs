// PROOF: the consumed-length contract of `src/addr.rs:282 parse_cell_ref_at`.
//
// Checked by:
//   /opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/cellref_parse.rs
// Result: 7 verified, 0 errors
//
// `parse_cell_ref_at` returns `(addr, locks, n)`, and every caller slices the
// formula source at `n` to find the next token. If `n` were short or long by
// one, the remainder would be mis-tokenised -- `A5+A6` would become `A5` then
// `+A6` shifted, or worse. So the property worth proving is not the address
// value (already covered by `col_decode.rs`) but that **the reported length is
// exactly the letters and digits consumed, and the byte after it is not a
// digit**.
//
// Proved here, over bytes:
//   * `letter_run` / `digit_run` count maximal runs, and stop exactly where the
//     run ends (`run_stops`, `digit_run_stops`);
//   * a run can never pass the end of the input (`letter_run_bound`,
//     `digit_run_bound`);
//   * `a1_len` is within the input, and the byte at `a1_len` is not a digit
//     (`a1_consumed_is_exact`) -- so a caller slicing there leaves no digit
//     behind.
//
// NOT proved: the header/footer (`~` / `_`) branch, the `$`-lock prefix rules,
// or `cell_ref_text` (which is `format!`-bound). The cross-check test covers the
// composed behaviour including those branches.
use vstd::prelude::*;
verus! {

pub open spec fn is_upper(b: u8) -> bool { 65 <= b <= 90 }
pub open spec fn is_digit_b(b: u8) -> bool { 48 <= b <= 57 }

/// Length of the maximal run of `A`..`Z` starting at `i` (bounded by `n`).
pub open spec fn letter_run(s: Seq<u8>, i: int, n: int) -> int
    decreases n - i,
{
    if i >= n || !is_upper(s[i]) { 0 } else { 1 + letter_run(s, i + 1, n) }
}

/// Length of the maximal run of `0`..`9` starting at `i` (bounded by `n`).
pub open spec fn digit_run(s: Seq<u8>, i: int, n: int) -> int
    decreases n - i,
{
    if i >= n || !is_digit_b(s[i]) { 0 } else { 1 + digit_run(s, i + 1, n) }
}

/// A run never passes the end of the input.
pub proof fn letter_run_bound(s: Seq<u8>, i: int, n: int)
    requires 0 <= i <= n <= s.len(),
    ensures i + letter_run(s, i, n) <= n,
    decreases n - i,
{
    if i >= n {
    } else if !is_upper(s[i]) {
    } else {
        letter_run_bound(s, i + 1, n);
    }
}

pub proof fn digit_run_bound(s: Seq<u8>, i: int, n: int)
    requires 0 <= i <= n <= s.len(),
    ensures i + digit_run(s, i, n) <= n,
    decreases n - i,
{
    if i >= n {
    } else if !is_digit_b(s[i]) {
    } else {
        digit_run_bound(s, i + 1, n);
    }
}

/// The letter run is exactly the maximal prefix of letters.
pub proof fn run_stops(s: Seq<u8>, i: int, n: int)
    requires 0 <= i <= n <= s.len(),
    ensures
        forall|j: int| i <= j < i + letter_run(s, i, n) ==> is_upper(#[trigger] s[j]),
        i + letter_run(s, i, n) >= n || !is_upper(s[i + letter_run(s, i, n)]),
    decreases n - i,
{
    if i >= n {
    } else if !is_upper(s[i]) {
    } else {
        run_stops(s, i + 1, n);
    }
}

pub proof fn digit_run_stops(s: Seq<u8>, i: int, n: int)
    requires 0 <= i <= n <= s.len(),
    ensures
        forall|j: int| i <= j < i + digit_run(s, i, n) ==> is_digit_b(#[trigger] s[j]),
        i + digit_run(s, i, n) >= n || !is_digit_b(s[i + digit_run(s, i, n)]),
    decreases n - i,
{
    if i >= n {
    } else if !is_digit_b(s[i]) {
    } else {
        digit_run_stops(s, i + 1, n);
    }
}

/// Consumed length of a plain `A1`-style ref beginning at `i`.
pub open spec fn a1_len(s: Seq<u8>, i: int) -> int {
    let l = letter_run(s, i, s.len() as int);
    i + l + digit_run(s, i + l, s.len() as int)
}

/// THE contract: the reported length is in range, and the byte just past it is
/// not a digit -- so a caller slicing there is left on a token boundary.
pub proof fn a1_consumed_is_exact(s: Seq<u8>, i: int)
    requires
        0 <= i <= s.len(),
        letter_run(s, i, s.len() as int) > 0,
        digit_run(s, i + letter_run(s, i, s.len() as int), s.len() as int) > 0,
    ensures
        a1_len(s, i) <= s.len(),
        a1_len(s, i) >= s.len() || !is_digit_b(s[a1_len(s, i)]),
{
    let l = letter_run(s, i, s.len() as int);
    let d = digit_run(s, i + l, s.len() as int);
    run_stops(s, i, s.len() as int);
    letter_run_bound(s, i, s.len() as int);
    digit_run_stops(s, i + l, s.len() as int);
    digit_run_bound(s, i + l, s.len() as int);
    assert(a1_len(s, i) == i + l + d);
    assert(i + l + d <= s.len());
    // the digit run stopped either at the end or on a non-digit
    assert(i + l + d >= s.len() || !is_digit_b(s[i + l + d]));
}

} // verus
fn main() {}
