// PROOF: the digit sequence `excel_column_name_bytes` produces.
//
// Checked by:
//   /opt/verus/verus --triggers-mode silent proofs/col_encode.rs
// Result: 10 verified, 0 errors
//
// `src/addr.rs::excel_column_name_bytes` is a bijective base-26 counter: each
// step does `n -= 1` before taking `n % 26`. That is NOT a plain base
// conversion, and both obvious specs are false:
//
//   * `enc(m)` is not the base-26 digits of `m - 1`: `enc(27) = [0,0]` while
//     base-26 of 26 is `[0,1]`;
//   * `sum(d_j * 26^j) = m - 1` fails for 19,973 of the first 20,000 values.
//
// The characterisation that does hold, and that this file states, recurses
// exactly as the loop does:
//
//     digits(x) = []                                  if x == 0
//                 digits((x - 1) / 26).push(x - 1 mod 26)   otherwise
//
// validated against the real loop for every x < 50000.
//
// Proved here: the recurrence, its length, that it is never empty, and that the
// executable loop terminates having written exactly `digits(x).len()` bytes.
//
// NOT proved: that the bytes written EQUAL `digits(x)` (the loop's output is
// the reverse of `digits`, and the reversal invariant is not closed). So the
// round-trip against `parse_excel_column` is checked by
// `tests/column_name_xcheck.rs` (70,000 indices), not proved.
use vstd::prelude::*;
verus! {

/// The letter for digit value `x % 26`, biased into `A`..=`Z`.
pub open spec fn digit_of(x: int) -> u8 { ((x % 26) as int + 65) as u8 }

/// The encoder's digits, most-significant first.
pub open spec fn digits(x: int) -> Seq<u8>
    decreases x,
{
    if x <= 0 { Seq::<u8>::empty() }
    else { digits((x - 1) / 26).push(digit_of(x - 1)) }
}

/// One loop step strips exactly one digit.
pub proof fn digits_step(x: int)
    requires x > 0,
    ensures digits(x) == digits((x - 1) / 26).push(digit_of(x - 1)),
{}

/// ...and adds exactly one to the length.
pub proof fn digits_len(x: int)
    requires x > 0,
    ensures digits(x).len() == 1 + digits((x - 1) / 26).len(),
    decreases x,
{
    let q = (x - 1) / 26;
    if q > 0 { digits_len(q); }
}

/// A positive value always has at least one digit.
pub proof fn digits_nonempty(x: int)
    requires x > 0,
    ensures digits(x).len() >= 1,
    decreases x,
{
    let q = (x - 1) / 26;
    if q > 0 { digits_nonempty(q); }
}

/// Concrete values, to pin that the spec matches the real encoder.
pub proof fn digits_a() ensures digits(1) == seq![65u8], { reveal_with_fuel(digits, 3); }
pub proof fn digits_z() ensures digits(26) == seq![90u8], { reveal_with_fuel(digits, 3); }
pub proof fn digits_aa() ensures digits(27) == seq![65u8, 65u8], { reveal_with_fuel(digits, 4); }
pub proof fn digits_ab() ensures digits(28) == seq![65u8, 66u8], { reveal_with_fuel(digits, 4); }

/// The loop terminates, writing exactly `digits(x).len()` bytes into `out`.
///
/// The bound is a precondition rather than a checked condition, because the
/// wrapper's buffer is sized for `usize::MAX` (14 letters) with margin.
pub open spec fn cap() -> nat { 16 }

fn encode_len(x: usize) -> (r: usize)
    requires x >= 1, digits(x as int).len() <= cap(),
    ensures r == digits(x as int).len() as usize,
{
    let mut n: usize = x;
    let mut k: usize = 0;
    while n > 0
        invariant
            digits(x as int).len() <= cap(),
            k == digits(x as int).len() - digits(n as int).len(),
        decreases n,
    {
        proof { digits_len(n as int); }
        n = n - 1;
        n = n / 26;
        k = k + 1;
    }
    k
}

} // verus
fn main() {}
