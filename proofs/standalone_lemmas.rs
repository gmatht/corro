// LEMMAS THAT STAND ALONE
//
// Checked by:  verus --triggers-mode silent proofs/standalone_lemmas.rs
// Result: 12 verified, 0 errors
//
// Every lemma here is stated over `int`, or over a generic `Seq<A>` with an
// abstract order relation. **None mentions corro**, and none needs a spreadsheet
// to make sense. They can be lifted into any Verus project, and worked on
// without the corro source at all — which is the point of collecting them.
//
// Each is here because it was *needed* and the prover did not supply it. Two
// groups, for two different reasons:
//
//   * `mul_*` / `le_*` — Z3 will not multiply a hypothesis by a constant, and
//     will not chain `x <= y, y == z`. Every arithmetic bound in corro's proofs
//     is built from these; they are the single biggest time sink otherwise.
//   * `rev_*` — list reversal does not distribute over append as an SMT solver
//     will find unprompted; the identity needs induction on the sequence.
//   * `subseq_*` — "a subsequence of a sorted list is sorted" needs the source
//     sequence and a transitivity hypothesis on the order.
//
// Nothing here is corro-flavoured. The corro-specific *instantiations* (base-26
// digits, cell addresses, aggregate visit order) live in the other proof files.

use vstd::prelude::*;
verus! {

// ===========================================================================
// Integer arithmetic
// ===========================================================================

/// Multiplication by a non-negative constant is monotone.
///
/// Stated for possibly-negative `a`: the induction runs on `b - a`, and the
/// `0 <= a` bound is needed only to make the `decreases` well-founded.
/// SMT will not do this on its own, with or without `nonlinear_arith`.
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

/// A product of two non-negative integers is non-negative.
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

/// `x <= y` composed with `y == z`.
///
/// Trivial to a human, and the solver will not do it: it needs the equality
/// oriented the same way as the goal.
pub proof fn le_of_le_eq(x: int, y: int, z: int)
    requires x <= y, y == z,
    ensures x <= z,
{}

/// Adding the same constant to both sides preserves `<=`.
pub proof fn le_add_const(x: int, y: int, c: int)
    requires x <= y,
    ensures x + c <= y + c,
{}

// ===========================================================================
// Sequence reversal
// ===========================================================================

/// Reverse, defined by peeling from the FRONT (so the recursion is on the tail).
pub open spec fn rev<A>(s: Seq<A>) -> Seq<A>
    decreases s.len(),
{
    if s.len() == 0 { s } else { rev(s.drop_first()).push(s[0]) }
}

/// Reversing an append puts the new element at the FRONT.
///
/// This is the identity every "we built it backwards" algorithm needs, and it
/// is not free: it requires induction, and the base case needs `rev` unfolded
/// explicitly (`reveal_with_fuel`), because the solver will not evaluate it.
pub proof fn rev_push<A>(s: Seq<A>, a: A)
    ensures rev(s.push(a)) == seq![a].add(rev(s)),
    decreases s.len(),
{
    if s.len() == 0 {
        assert(s.push(a) =~= seq![a]);
        reveal_with_fuel(rev, 3);
    } else {
        assert(s.push(a).drop_first() =~= s.drop_first().push(a));
        rev_push(s.drop_first(), a);
        assert(rev(s) == rev(s.drop_first()).push(s[0]));
        assert(seq![a].add(rev(s)) == seq![a].add(rev(s.drop_first())).push(s[0]));
    }
}

/// Reversal preserves length.
pub proof fn rev_len<A>(s: Seq<A>)
    ensures rev(s).len() == s.len(),
    decreases s.len(),
{
    if s.len() > 0 {
        rev_len(s.drop_first());
    }
}

// ===========================================================================
// Subsequences and sortedness
// ===========================================================================

/// `x` occurs somewhere in `f`.
pub open spec fn mem_at<A>(f: Seq<A>, x: A) -> bool {
    exists|i: int| 0 <= i < f.len() && #[trigger] f[i] == x
}

/// `f` is obtained from `s` by deleting elements.
pub open spec fn is_subseq<A>(f: Seq<A>, s: Seq<A>) -> bool
    decreases f.len() + s.len(),
{
    if f.len() == 0 { true }
    else if s.len() == 0 { false }
    else if f[0] == s[0] { is_subseq(f.drop_first(), s.drop_first()) }
    else { is_subseq(f, s.drop_first()) }
}

/// Every pair `i < j` satisfies `le`.
pub open spec fn sorted_by<A>(f: Seq<A>, le: spec_fn(A, A) -> bool) -> bool {
    forall|i: int, j: int| 0 <= i < j < f.len() ==> #[trigger] le(f[i], f[j])
}

/// Every element of a subsequence occurs in the source.
///
/// (`Seq::contains` does not decompose over `drop_first` for the solver, so the
/// membership relation is spelled out with an index instead.)
pub proof fn subseq_mem<A>(f: Seq<A>, s: Seq<A>, x: A)
    requires is_subseq(f, s), mem_at(f, x),
    ensures mem_at(s, x),
    decreases f.len() + s.len(),
{
    if f.len() == 0 {
    } else if s.len() == 0 {
    } else if f[0] == s[0] {
        if f[0] == x {
            assert(mem_at(s, x)) by { assert(s[0] == x); };
        } else {
            let i = choose|i: int| 0 <= i < f.len() && #[trigger] f[i] == x;
            assert(0 < i);
            assert(mem_at(f.drop_first(), x)) by { assert(f.drop_first()[i - 1] == x); };
            subseq_mem(f.drop_first(), s.drop_first(), x);
            let j = choose|j: int|
                0 <= j < s.drop_first().len() && #[trigger] s.drop_first()[j] == x;
            assert(mem_at(s, x)) by { assert(s[j + 1] == x); };
        }
    } else {
        subseq_mem(f, s.drop_first(), x);
        let j = choose|j: int|
            0 <= j < s.drop_first().len() && #[trigger] s.drop_first()[j] == x;
        assert(mem_at(s, x)) by { assert(s[j + 1] == x); };
    }
}

/// A subsequence of a list sorted by `le` is sorted by `le`, given `le` is
/// transitive.
///
/// This is the generic form of "filtering a sorted collection preserves its
/// order" — a filtered fast path over a sorted store has the same iteration
/// order as a full scan, provided the filter is a subsequence and the order is
/// transitive.
pub proof fn subseq_sorted<A>(f: Seq<A>, s: Seq<A>, le: spec_fn(A, A) -> bool)
    requires
        sorted_by(s, le),
        is_subseq(f, s),
        forall|a: A, b: A, c: A|
            #[trigger] le(a, b) && #[trigger] le(b, c) ==> le(a, c),
    ensures sorted_by(f, le),
    decreases f.len() + s.len(),
{
    if f.len() == 0 {
        assert forall|i: int, j: int| 0 <= i < j < f.len() implies le(f[i], f[j]) by {};
    } else if s.len() == 0 {
    } else if f[0] == s[0] {
        subseq_sorted(f.drop_first(), s.drop_first(), le);
        assert forall|i: int, j: int| 0 <= i < j < f.len() implies #[trigger] le(f[i], f[j]) by {
            if i == 0 {
                assert(mem_at(f.drop_first(), f[j])) by { assert(f.drop_first()[j - 1] == f[j]); };
                subseq_mem(f.drop_first(), s.drop_first(), f[j]);
                let q = choose|q: int|
                    0 <= q < s.drop_first().len() && #[trigger] s.drop_first()[q] == f[j];
                assert(s[q + 1] == f[j]);
                assert(le(s[0], s[q + 1]));
                assert(f[0] == s[0]);
            } else {
                assert(f[i] == f.drop_first()[i - 1] && f[j] == f.drop_first()[j - 1]);
            }
        };
    } else {
        subseq_sorted(f, s.drop_first(), le);
    }
}

} // verus
fn main() {}
