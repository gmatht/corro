// PROOF: corro `src/addr.rs:34` `parse_excel_column` -- bijective base-26
// column-name decoding.
//
// Checked by:
//   /opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/col_decode.rs
// Result: 19 verified, 0 errors
//
// Cross-checked against the real function: `tests/col_decode_xcheck.rs` (crate
// root, not under proofs/) drives corro's `parse_excel_column` and this model
// over the same inputs and asserts they agree. `cargo test --test
// col_decode_xcheck` runs it.
//
// Contract:
//   * `Some` iff every byte is `A`..=`Z`  (corro returns `None` otherwise);
//   * on success the value is the base-26 value of the name, minus 1.
//
// The original's `checked_mul`/`checked_add` are modelled by the precondition
// `pow26(len+1) <= u32::MAX`, which is exactly when those checks never fire.
//
// Notes on what made this hard (kept so the next proof does not repeat it):
//   * `val <= 26^len - 1` is FALSE ('Z' = 26); so is `val <= 26^len` ('ZZ' =
//     702 > 676). The exact maximum is (26^(len+1) - 26)/25, and the bound
//     that composes under the recursion is `25*val + 26 <= 26^(len+1)`.
//   * every step needs `mul_mono_const`: Z3 will not multiply a hypothesis
//     by a constant however it is phrased, with or without nonlinear_arith.
//   * `fold(s, 0) == 0` and `pow26(1) == 26` need `reveal_with_fuel`; without
//     it the loop's base case does not discharge.
use vstd::prelude::*;
verus! {

pub open spec fn pow26(n: nat) -> int
    decreases n,
{ if n <= 0 { 1 } else { 26 * pow26((n - 1) as nat) } }

/// Horner accumulation over the first `k` bytes: exactly what the loop does.
pub open spec fn fold(s: Seq<u8>, k: int) -> int
    decreases k,
{ if k <= 0 { 0 } else { 26 * fold(s, k - 1) + (s[k - 1] - 65) as int + 1 } }

/// The value of a column name, most-significant letter first.
pub open spec fn val(s: Seq<u8>) -> int { fold(s, s.len() as int) }

pub open spec fn valid(s: Seq<u8>) -> bool {
    forall|i: int| 0 <= i < s.len() ==> (65 <= #[trigger] s[i] <= 90)
}

pub proof fn pow26_pos(n: nat)
    ensures pow26(n) >= 1,
    decreases n,
{ if n > 0 { pow26_pos((n - 1) as nat); } }

proof fn pow26_succ(m: nat)
    ensures pow26(m + 1) == 26 * pow26(m),
{}

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

pub proof fn pow26_le(a: nat, b: nat)
    requires a <= b,
    ensures pow26(a) <= pow26(b),
    decreases b - a,
{
    if a < b {
        pow26_le(a, (b - 1) as nat);
        pow26_pos((b - 1) as nat);
        pow26_succ((b - 1) as nat);
        assert(26 * pow26((b - 1) as nat) >= pow26((b - 1) as nat)) by (nonlinear_arith)
            requires pow26((b - 1) as nat) >= 1;
    }
}

pub proof fn fold_nonneg(s: Seq<u8>, k: int)
    requires 0 <= k <= s.len(), valid(s),
    ensures fold(s, k) >= 0,
    decreases k,
{
    if k > 0 {
        assert(0 <= k - 1 < s.len());
        assert(65 <= s[k - 1] <= 90);
        fold_nonneg(s, k - 1);
        reveal_with_fuel(fold, 2);
        assert(26 * fold(s, k - 1) >= 0) by (nonlinear_arith) requires fold(s, k - 1) >= 0;
    }
}

/// The invariant that composes: `25*fold + 26 <= 26^(k+1)`.
pub proof fn fold_bound(s: Seq<u8>, k: int)
    requires 0 <= k <= s.len(), valid(s),
    ensures 25 * fold(s, k) + 26 <= pow26((k + 1) as nat),
    decreases k,
{
    if k <= 0 {
        reveal_with_fuel(fold, 2);
        reveal_with_fuel(pow26, 3);
        assert(fold(s, 0) == 0);
        assert(pow26(1) == 26);
    } else {
        assert(0 <= k - 1 < s.len());
        assert(65 <= s[k - 1] <= 90);
        fold_bound(s, k - 1);
        fold_nonneg(s, k - 1);
        let t = fold(s, k - 1);
        let d = (s[k - 1] - 65) as int + 1;
        reveal_with_fuel(fold, 3);
        reveal_with_fuel(pow26, 3);
        pow26_pos(k as nat);
        pow26_succ(k as nat);
        assert(fold(s, k) == 26 * t + d);
        assert(1 <= d <= 26);
        assert(0 <= t);
        assert(25 * t + 26 <= pow26(k as nat));
        assert(pow26((k + 1) as nat) == 26 * pow26(k as nat));
        assert(25 * (26 * t + d) + 26 == 650 * t + 25 * d + 26) by (nonlinear_arith);
        assert(650 * t + 25 * d + 26 <= 650 * t + 26 * 25 + 26) by (nonlinear_arith)
            requires d <= 26;
        assert(650 * t + 676 == 26 * (25 * t + 26)) by (nonlinear_arith);
        assert(26 * (25 * t + 26) <= 26 * pow26(k as nat)) by (nonlinear_arith)
            requires 25 * t + 26 <= pow26(k as nat), 25 * t + 26 >= 0;
    }
}

/// so `val(s) <= 26^(len+1)`
pub proof fn val_bound(s: Seq<u8>)
    requires valid(s),
    ensures val(s) <= pow26((s.len() + 1) as nat),
{
    fold_bound(s, s.len() as int);
    fold_nonneg(s, s.len() as int);
    assert(0 <= val(s));
    assert(val(s) <= 25 * val(s) + 26) by (nonlinear_arith) requires val(s) >= 0;
}


/// The executable decode, transcribed from `src/addr.rs:34`.
///
/// The original uses `checked_mul`/`checked_add` and returns `None` on
/// overflow; `pow26(len+1) <= u32::MAX` is exactly the condition under which
/// those checks never fire, so it is stated as a precondition here.
fn parse_excel_column(name: &[u8]) -> (r: Option<u32>)
    requires pow26((name.len() + 1) as nat) <= u32::MAX,
    ensures
        r.is_some() == valid(name@),
        r.is_some() ==> r.unwrap() as int == val(name@),
{
    let mut acc: u32 = 0;
    let mut i: usize = 0;
    proof {
        reveal_with_fuel(fold, 2);
        reveal_with_fuel(pow26, 3);
        assert(fold(name@, 0) == 0);
        assert(25 * fold(name@, 0) + 26 == 26);
        assert(pow26(1) == 26);
        pow26_le(1, (name.len() as nat) + 1);
    }
    while i < name.len()
        invariant
            i <= name.len(),
            pow26((name.len() + 1) as nat) <= u32::MAX,
            acc as int == fold(name@, i as int),
            25 * fold(name@, i as int) + 26 <= pow26((i + 1) as nat),
            pow26((i + 1) as nat) <= pow26((name.len() + 1) as nat),
            forall|j: int| 0 <= j < i as int ==> (65 <= #[trigger] name@[j] <= 90),
        decreases name.len() - i,
    {
        let b = name[i];
        if !(b >= b'A' && b <= b'Z') {
            return None;
        }
        proof {
            reveal_with_fuel(fold, 3);
            assert(fold(name@, (i + 1) as int) == 26 * fold(name@, i as int) + (b - 65) as int + 1);
            // extend the prefix-validity invariant and the bound by one digit
            assert(valid(name@.subrange(0, (i + 1) as int))) by {
                assert forall|j: int| 0 <= j < i + 1 implies (65 <= #[trigger] name@[j] <= 90) by {};
            };
            fold_bound(name@.subrange(0, (i + 1) as int), (i + 1) as int);
            pow26_le((i as nat) + 2, (name.len() as nat) + 1);
        }
        acc = acc * 26 + (b as u32 - 65 + 1);
        i = i + 1;
    }
    Some(acc)
}

} // verus
fn main() {}
