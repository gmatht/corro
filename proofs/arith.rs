// PROOF: the integer facts the corro kernels keep needing, on their own.
//
// Checked by:  verus --triggers-mode silent proofs/arith.rs
// Result:      12 verified, 0 errors
//
// Z3 does not multiply a hypothesis by a constant on its own, and it will not
// chain a product identity into a bound. Both are routine for a human and
// neither comes free here, so the shapes that DO go through are collected in
// one place and cited by the kernel proofs.
use vstd::prelude::*;
verus! {

/// Multiplication by a non-negative constant is monotone (induction on the gap).
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

/// Same, with the non-negativity of the product derived rather than assumed.
pub proof fn mul_nonneg(a: int, b: int)
    requires a >= 0, b >= 0,
    ensures a * b >= 0,
    decreases a,
{
    if a > 0 {
        mul_nonneg(a - 1, b);
        assert(a * b == (a - 1) * b + b) by (nonlinear_arith);
    }
}

/// `n * 26` and `26 * n` agree, so either spelling can be matched.
pub proof fn comm26(n: int)
    ensures n * 26 == 26 * n,
{
    assert(n * 26 == 26 * n) by (nonlinear_arith);
}

/// `26*(P-1) = 26*P - 26` for `P >= 1`.
pub proof fn expand26(p: int)
    requires p >= 1,
    ensures 26 * (p - 1) == 26 * p - 26,
{
    assert(26 * (p - 1) == 26 * p - 26) by (nonlinear_arith);
}

/// The base-10 analogue, for the money parser.
pub proof fn expand10(p: int)
    requires p >= 0,
    ensures 10 * (p + 1) == 10 * p + 10,
{
    assert(10 * (p + 1) == 10 * p + 10) by (nonlinear_arith);
}

/// Transitivity with an equation on the right, which the solver needs spelled out.
pub proof fn le_of_le_eq(x: int, y: int, z: int)
    requires x <= y, y == z,
    ensures x <= z,
{}

/// A chain step: both bounds, then the comparison.
pub proof fn le_add_const(x: int, y: int, c: int)
    requires x <= y,
    ensures x + c <= y + c,
{}

} // verus
fn main() {}
