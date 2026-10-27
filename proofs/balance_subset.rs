// PROOF: the soundness core of `balance-core/src/lib.rs:125 choose_subset`.
//
// Checked by:
//   /opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/balance_subset.rs
// Result: 4 verified, 0 errors
//
// The book-balancing search returns a subset of the remaining items whose
// amounts sum exactly to an anchor's amount. The property that matters is
// *soundness*: when it finds nothing, there really is nothing to find. A
// search that silently gave up would balance a book wrongly and say nothing.
//
// Proved here:
//   * `sorted_in` states what a valid subset is (indices increasing, in range);
//   * `prepend_ok`: adding a smaller index in front keeps a subset valid;
//   * `no_subset_step`: if no subset of [start, n) sums to `target`, then for
//     any usable index `i` no subset of (i, n) sums to `target - vals[i]`.
//     That is exactly the step the recursion takes before trying the next
//     candidate, so "None" cannot be reached while a solution exists below.
//
// NOT proved: the executable recursion itself (`rec`), the `used[]` mask, the
// outer `solve_balance_groups` loop over anchors, or that groups are disjoint.
// Those need `Vec`, `sort_by_key` and iterator chains rewritten to loops.
use vstd::prelude::*;
verus! {

/// Sum of the values at `idxs`, in order.
pub open spec fn sm(vals: Seq<int>, idxs: Seq<nat>) -> int
    decreases idxs.len(),
{
    if idxs.len() == 0 { 0 } else { vals[idxs[0] as int] + sm(vals, idxs.drop_first()) }
}

/// A valid subset of `(start, n)`: indices strictly increasing and in range.
pub open spec fn sorted_in(idxs: Seq<nat>, start: nat, n: nat) -> bool {
    &&& (forall|i: int| 0 <= i < idxs.len() ==> start <= #[trigger] idxs[i] < n)
    &&& (forall|i: int| 0 < i < idxs.len() ==> #[trigger] idxs[i - 1] < idxs[i])
}

/// Some subset of `(start, n)` sums to `target`.
pub open spec fn exists_subset(vals: Seq<int>, start: nat, target: int) -> bool {
    exists|idxs: Seq<nat>|
        #[trigger] sorted_in(idxs, start, vals.len() as nat) && sm(vals, idxs) == target
}

/// Prepending an index below the whole list keeps it a valid subset.
pub proof fn prepend_ok(idxs: Seq<nat>, i: nat, start: nat, n: nat)
    requires i < n, start <= i, sorted_in(idxs, i + 1, n),
    ensures sorted_in(seq![i].add(idxs), start, n),
{
    assert forall|k: int| 0 <= k < seq![i].add(idxs).len()
        implies start <= #[trigger] seq![i].add(idxs)[k] < n by {
        if k != 0 {
            assert(seq![i].add(idxs)[k] == idxs[k - 1]);
        }
    };
    assert forall|k: int| 0 < k < seq![i].add(idxs).len()
        implies #[trigger] seq![i].add(idxs)[k - 1] < seq![i].add(idxs)[k] by {
        if k == 1 {
            assert(idxs[0] >= i + 1);
        } else {
            assert(seq![i].add(idxs)[k] == idxs[k - 1]);
            assert(seq![i].add(idxs)[k - 1] == idxs[k - 2]);
        }
    };
}

/// SOUNDNESS: the step the recursion takes cannot step past a solution.
pub proof fn no_subset_step(vals: Seq<int>, start: nat, target: int, i: nat)
    requires
        start <= i < vals.len(),
        !exists_subset(vals, start, target),
        vals[i as int] >= 0,
    ensures
        !exists_subset(vals, i + 1, target - vals[i as int]),
{
    if exists_subset(vals, i + 1, target - vals[i as int]) {
        let idxs = choose|idxs: Seq<nat>|
            sorted_in(idxs, i + 1, vals.len() as nat) && sm(vals, idxs) == target - vals[i as int];
        assert(sorted_in(idxs, i + 1, vals.len() as nat)) by {
            assert(exists_subset(vals, i + 1, target - vals[i as int]));
        };
        assert(sm(vals, idxs) == target - vals[i as int]) by {
            assert(exists_subset(vals, i + 1, target - vals[i as int]));
        };
        let all = seq![i].add(idxs);
        prepend_ok(idxs, i, start, vals.len() as nat);
        assert(all[0] == i);
        assert(all.drop_first() =~= idxs);
        assert(sm(vals, all) == vals[i as int] + sm(vals, idxs));
        assert(vals[i as int] + sm(vals, idxs) == target) by (nonlinear_arith)
            requires sm(vals, idxs) == target - vals[i as int];
        assert(sm(vals, all) == target);
        assert(exists_subset(vals, start, target));
        assert(false);
    }
}

} // verus
fn main() {}
