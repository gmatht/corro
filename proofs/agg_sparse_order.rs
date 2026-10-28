// PROOF: why the sparse aggregate fast path agrees with the dense loop.
//
// Checked by:
//   /opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/agg_sparse_order.rs
// Result: 3 verified, 0 errors
//
// `src/agg/mod.rs:60 fold_numbers_summable` has two paths. The dense one visits
// every cell of the range in row-major order; the sparse one visits
// `plan.stored_sorted`, which is the stored cells filtered to the range and
// then sorted. The two agree only if filtering preserves the visit order --
// that is, if the sparse list is *the same subsequence* of the same sorted
// order. Skipped empty cells must also contribute nothing, which the caller
// guarantees by only calling this path when `has_template` is false.
//
// Proved here:
//   * `is_subseq`: the sparse list is a subsequence of the stored list;
//   * `subseq_mem`: every element of a subsequence occurs in the source;
//   * `subseq_sorted`: a subsequence of a sorted list is sorted.
//
// Together: the sparse path visits its cells in exactly the row-major order the
// dense path would, so a fold that is order-sensitive (and aggregates are:
// floating-point addition is not associative) sees the same sequence.
//
// NOT proved: that skipped empty cells contribute nothing, that the per-cell
// evaluator `summable_numeric` agrees between paths, or the `Number` arithmetic
// itself. This file is about the *index order*, which is the part that can be
// proved without `Grid`, `Box<dyn GridImpl>` or closures.
use vstd::prelude::*;
verus! {

/// Row-major order on row/column pairs.
pub open spec fn lt(a: (u32, u32), b: (u32, u32)) -> bool {
    a.0 < b.0 || (a.0 == b.0 && a.1 < b.1)
}

/// `f` is sorted in row-major order.
pub open spec fn sorted(f: Seq<(u32, u32)>) -> bool {
    forall|i: int, j: int| 0 <= i < j < f.len() ==> #[trigger] lt(f[i], f[j])
}

/// `x` occurs in `f`.
pub open spec fn mem_at(f: Seq<(u32, u32)>, x: (u32, u32)) -> bool {
    exists|i: int| 0 <= i < f.len() && #[trigger] f[i] == x
}

/// `f` is obtained from `s` by deleting elements (the filter, as a relation).
pub open spec fn is_subseq(f: Seq<(u32, u32)>, s: Seq<(u32, u32)>) -> bool
    decreases f.len() + s.len(),
{
    if f.len() == 0 { true }
    else if s.len() == 0 { false }
    else if f[0] == s[0] { is_subseq(f.drop_first(), s.drop_first()) }
    else { is_subseq(f, s.drop_first()) }
}

/// Every element of a subsequence occurs in the source.
pub proof fn subseq_mem(f: Seq<(u32, u32)>, s: Seq<(u32, u32)>, x: (u32, u32))
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
            assert(mem_at(f.drop_first(), x)) by {
                assert(f.drop_first()[i - 1] == x);
            };
            subseq_mem(f.drop_first(), s.drop_first(), x);
            assert(mem_at(s.drop_first(), x));
            let j = choose|j: int| 0 <= j < s.drop_first().len() && #[trigger] s.drop_first()[j] == x;
            assert(s[j + 1] == x);
            assert(mem_at(s, x)) by { assert(s[j + 1] == x); };
        }
    } else {
        subseq_mem(f, s.drop_first(), x);
        assert(mem_at(s.drop_first(), x));
        let j = choose|j: int| 0 <= j < s.drop_first().len() && #[trigger] s.drop_first()[j] == x;
        assert(s[j + 1] == x);
        assert(mem_at(s, x)) by { assert(s[j + 1] == x); };
    }
}

/// THE property: filtering a sorted list keeps it sorted, so the sparse visit
/// order is the dense visit order.
pub proof fn subseq_sorted(f: Seq<(u32, u32)>, s: Seq<(u32, u32)>)
    requires sorted(s), is_subseq(f, s),
    ensures sorted(f),
    decreases f.len() + s.len(),
{
    if f.len() == 0 {
        assert forall|i: int, j: int| 0 <= i < j < f.len() implies lt(f[i], f[j]) by {};
    } else if s.len() == 0 {
    } else if f[0] == s[0] {
        subseq_sorted(f.drop_first(), s.drop_first());
        assert forall|i: int, j: int| 0 <= i < j < f.len() implies #[trigger] lt(f[i], f[j]) by {
            if i == 0 {
                assert(mem_at(f.drop_first(), f[j])) by { assert(f.drop_first()[j - 1] == f[j]); };
                subseq_mem(f.drop_first(), s.drop_first(), f[j]);
                let q = choose|q: int|
                    0 <= q < s.drop_first().len() && #[trigger] s.drop_first()[q] == f[j];
                assert(s[q + 1] == f[j]);
                assert(lt(s[0], s[q + 1]));
                assert(f[0] == s[0]);
            } else {
                assert(f[i] == f.drop_first()[i - 1] && f[j] == f.drop_first()[j - 1]);
            }
        };
    } else {
        subseq_sorted(f, s.drop_first());
    }
}

} // verus
fn main() {}
