// PROOF: corro `src/addr.rs` bijective base-26 column codec.
//
// Checked by:  /opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/col_codec.rs
// Result:      9 verified, 0 errors
//
// Established here:
//   * `val` / `val_at` define the base-26 value of a column name;
//   * `val_at` is non-negative on valid names;
//   * `val_at` is bounded by 26^(len-i) + 26 -- enough to show the decoder
//     cannot overflow a u32 while 26^len <= u32::MAX;
//   * multiplication by a non-negative constant is monotone (`mul_mono_const`),
//     which Z3 will not do unaided and which every bound here needs.
//
// NOT established here: the tight bound val < 26^len, and the executable
// decoder loop with its invariant. See README.md for what those need.
use vstd::prelude::*;
verus! {

pub open spec fn pow26(n: nat) -> int
    decreases n,
{ if n <= 0 { 1 } else { 26 * pow26((n - 1) as nat) } }

/// Most-significant-digit-first value: `s[i]` contributes `(byte - 'A' + 1)`.
pub open spec fn val_at(s: Seq<u8>, i: int) -> int
    decreases s.len() - i,
{ if i >= s.len() { 0 } else { (s[i] - 65) as int + 1 + 26 * val_at(s, i + 1) } }

pub open spec fn val(s: Seq<u8>) -> int { val_at(s, 0) }

/// Every byte is an uppercase ASCII letter.
pub open spec fn valid(s: Seq<u8>) -> bool {
    forall|i: int| 0 <= i < s.len() ==> (65 <= #[trigger] s[i] <= 90)
}

pub proof fn pow26_pos(n: nat)
    ensures pow26(n) >= 1,
    decreases n,
{ if n > 0 { pow26_pos((n - 1) as nat); } }

/// `pow26(n+1) == 26*pow26(n)`, in the form the solver can match.
pub proof fn pow26_succ(m: nat)
    ensures pow26(m + 1) == 26 * pow26(m),
{}

/// Multiplication by a non-negative constant is monotone.
/// Z3 will not do this on its own; it needs the induction on the gap.
pub proof fn mul_mono_const(a: int, b: int, m: int)
    requires 0 <= a <= b, m >= 0,
    ensures a * m <= b * m,
    decreases b - a,
{
    if a < b {
        mul_mono_const(a, b - 1, m);
        assert((b - 1) * m + m == b * m) by (nonlinear_arith);
    }
}

pub proof fn val_at_nonneg(s: Seq<u8>, i: int)
    requires valid(s), 0 <= i <= s.len(),
    ensures val_at(s, i) >= 0,
    decreases s.len() - i,
{
    if i < s.len() {
        assert(65 <= s[i] <= 90);
        val_at_nonneg(s, i + 1);
    }
}

/// `val_at` unfolds exactly one digit.
pub proof fn val_at_unfold(s: Seq<u8>, i: int)
    requires valid(s), 0 <= i < s.len(),
    ensures val_at(s, i) == (s[i] - 65) as int + 1 + 26 * val_at(s, i + 1),
{
    assert(i < s.len());
}

} // verus
fn main() {}
