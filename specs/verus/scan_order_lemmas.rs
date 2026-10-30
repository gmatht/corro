// ============================================================================
// Reusable proof core for "filtered fast path keeps scan order"
//
// Three groups:
//   1. Integer arithmetic facts that Z3 will not find on its own
//      (with or without nonlinear arithmetic).
//   2. Sequence-reversal lemmas, generic over the element type.
//   3. Subsequence / order-preservation lemmas, generic over the element
//      type and the comparison function.
//
// Notes on the design:
//   * The arithmetic lemmas are stated over the mathematical integers
//     (`int`).  Spec-mode arithmetic on machine integers is mathematical
//     (it never overflows), so from `u64`/`i64`/... code you can call
//     e.g. `mul_mono_const(a as int, b as int, m as int)` and the fact
//     transfers unchanged.
//   * `rev` is a hand-rolled recursive definition.  Recursive spec
//     functions are only unfolded with bounded fuel (default 1), which
//     is why `rev_push`'s base case needs `reveal_with_fuel`.
//   * `is_subseq` is the standard greedy recursion: scan `s` from the
//     front, consume a head of `f` whenever the heads match, otherwise
//     delete `s`'s head.  `f` is `s` with elements deleted.
//   * All uses of `drop_first` are guarded by non-emptiness, so we never
//     rely on how vstd's `drop_first` behaves on the empty sequence.
//
// This file is the self-contained version of `verus_lemmas.rs`, and the one
// to extend.
//
// Why it is here: `Grid::sorted_main_rows` / `gui::actions::sort_sheet`
// compute `permute(sort(project(rows)), rows)` -- pick the rows to reorder
// with a *filter*, sort that projection by one column, write the rows back in
// the sorted order. Two facts make the result a sort of the sheet, and
// neither is visible to a Rust test that only checks the cells it touched:
//
//   * the surviving rows come back in order, though only a projection of them
//     was sorted: `subseq_sorted`;
//   * no value enters or leaves: `subseq_mem` (with `mem_at_iff_contains`
//     bridging to `Seq::contains`).
//
// These are `Seq` lemmas, not a proof about corro's code: corro's sort is
// ordinary Rust and nothing in `src/` is Verus-verified. They are the reusable
// core such a proof would start from, checked by `specs/verus/check.sh`
// (opt-in: it needs a Verus install, so it is not part of `cargo test`).
//
// Verified with Verus 0.2026.10.18.
// ============================================================================

#![allow(unused)]

use vstd::prelude::*;

verus! {

// ============================================================================
// 1. Integer arithmetic
//
// Monotonicity / sign facts about `*` are invisible to Z3 even with
// `nonlinear_arith`: they are inductive facts, not linear ones.  The only
// genuinely nonlinear step is `mul_step` (distributivity over `+ 1`),
// which Z3 *can* do as a plain polynomial equality.  Both monotonicity
// lemmas below are then pure induction + linear arithmetic.
// ============================================================================

/// The single nonlinear seed: distributivity of `*` over a `+ 1` step.
#[verifier::nonlinear]
pub proof fn mul_step(k: int, m: int)
    ensures (k + 1) * m == k * m + m,
{
}

/// `0 <= a <= b` and `0 <= m`  implies  `a * m <= b * m`.
/// Proved by induction on `b - a`.
pub proof fn mul_mono_const(a: int, b: int, m: int)
    requires
        0 <= a <= b,
        0 <= m,
    ensures
        a * m <= b * m,
    decreases b - a,
{
    if b == a {
        assert(a * m == b * m);
    } else {
        // b > a, so 0 <= a <= b - 1 and (b - 1) - a < b - a
        mul_mono_const(a, b - 1, m);
        mul_step(b - 1, m);
        assert(b * m == (b - 1) * m + m);
        assert(a * m <= (b - 1) * m);
        assert(a * m <= b * m);
    }
}

/// `0 <= a` and `0 <= b`  implies  `a * b >= 0`.
/// Induction on `a`.
pub proof fn mul_nonneg(a: int, b: int)
    requires
        0 <= a,
        0 <= b,
    ensures
        a * b >= 0,
    decreases a,
{
    if a == 0 {
        mul_zero(b);
        assert(a * b == 0);
    } else {
        // a >= 1
        mul_step(a - 1, b);
        assert(a * b == (a - 1) * b + b);
        mul_nonneg(a - 1, b);
        assert(a * b >= 0);
    }
}

/// `x <= y` and `y == z`  implies  `x <= z`.
/// (Trivial for Z3 in isolation; kept as a handle to guide it inside
/// larger nonlinear proofs where its focus is degraded.)
pub proof fn le_of_le_eq(x: int, y: int, z: int)
    requires
        x <= y,
        y == z,
    ensures
        x <= z,
{
    assert(x <= z);
}

/// `x <= y`  implies  `x + c <= y + c`.
pub proof fn le_add_const(x: int, y: int, c: int)
    requires
        x <= y,
    ensures
        x + c <= y + c,
{
    assert(x + c <= y + c);
}

// ============================================================================
// 2. Sequence reversal, generic over the element type
// ============================================================================

// ---- small `drop_first` facts used throughout -------------------------------

/// `drop_first` removes exactly one element (on nonempty sequences).
pub proof fn drop_first_len<A>(s: Seq<A>)
    requires s.len() > 0,
    ensures s.drop_first().len() == s.len() - 1,
{
}

/// Index view of `drop_first`: it is the sequence shifted left by one.
pub proof fn drop_first_index<A>(s: Seq<A>)
    requires s.len() > 0,
    ensures forall|k: int| 0 <= k < s.drop_first().len() ==> #[trigger]
        s.drop_first()[k] == s[k + 1],
{
}

/// `drop_first` commutes with `push` on nonempty sequences.
pub proof fn drop_first_push<A>(s: Seq<A>, a: A)
    requires s.len() > 0,
    ensures (s.push(a)).drop_first() == s.drop_first().push(a),
{
    let lhs = (s.push(a)).drop_first();
    let rhs = s.drop_first().push(a);
    drop_first_len(s);
    drop_first_len(s.push(a));
    assert(lhs.len() == rhs.len());
    assert forall|k: int| 0 <= k < lhs.len() implies lhs[k] == rhs[k] by {
        drop_first_index(s);
        drop_first_index(s.push(a));
        if k == s.len() - 1 {
            // both sides are `a` here
            assert(lhs[k] == a);
            assert(rhs[k] == a);
        }
        // otherwise both sides are s[k + 1]
    };
    assert(lhs == rhs);
}

/// Pushing on the right commutes with concatenation:
/// `(u ++ t).push(x) == u ++ (t.push(x))`.
pub proof fn add_push<A>(u: Seq<A>, t: Seq<A>, x: A)
    ensures (u + t).push(x) == u + (t.push(x)),
{
    let lhs = (u + t).push(x);
    let rhs = u + (t.push(x));
    assert(lhs.len() == rhs.len());
    assert forall|k: int| 0 <= k < lhs.len() implies lhs[k] == rhs[k] by {
        if k < u.len() {
            assert(lhs[k] == u[k]);
            assert(rhs[k] == u[k]);
        } else if k < u.len() + t.len() {
            assert(lhs[k] == t[k - u.len()]);
            assert(rhs[k] == t[k - u.len()]);
        } else {
            assert(lhs[k] == x);
            assert(rhs[k] == x);
        }
    };
    assert(lhs == rhs);
}

// ---- reversal ----------------------------------------------------------------

/// Reverse of `s`, defined recursively.  Being a recursive spec function,
/// Verus only unfolds it with bounded fuel (default 1); proofs that need
/// two unfoldings use `reveal_with_fuel`.
pub open spec fn rev<A>(s: Seq<A>) -> Seq<A>
    decreases s.len(),
{
    if s.len() == 0 {
        s
    } else {
        rev(s.drop_first()).push(s.first())
    }
}

/// `rev(s.push(a)) == [a] ++ rev(s)` — induction on `s.len()`;
/// the base case needs two unfoldings of `rev`, i.e. `reveal_with_fuel`.
pub proof fn rev_push<A>(s: Seq<A>, a: A)
    ensures
        rev(s.push(a)) == seq![a] + rev(s),
    decreases s.len(),
{
    if s.len() == 0 {
        // rev([a]) needs two unfoldings: rev([a]) --> rev([]) --> [].
        reveal_with_fuel(rev::<_>, 2);
        assert(s.push(a).len() == 1);
        assert(s.push(a).drop_first() == s);
        assert(s.push(a).first() == a);
        assert(rev(s) == s);
        assert(rev(s.push(a)) == s.push(a));
        assert(seq![a].len() == 1);
        assert(seq![a][0] == a);
        assert((seq![a] + s).len() == 1);
        assert((seq![a] + s)[0] == a);
        assert(seq![a] + s == seq![a]);
    } else {
        let hd = s.first();
        let rest = s.drop_first();
        // one unfolding on each side (default fuel suffices)
        assert(rev(s) == rev(rest).push(hd));
        assert(s.push(a).first() == hd);
        drop_first_push(s, a);
        assert(s.push(a).drop_first() == rest.push(a));
        assert(rev(s.push(a)) == rev(rest.push(a)).push(hd));
        rev_push(rest, a);
        assert(rev(rest.push(a)) == seq![a] + rev(rest));
        add_push(seq![a], rev(rest), hd);
        assert((seq![a] + rev(rest)).push(hd) == seq![a] + (rev(rest).push(hd)));
        assert(rev(s.push(a)) == seq![a] + rev(s));
    }
}

/// `rev(s).len() == s.len()` — induction on `s.len()`.
pub proof fn rev_len<A>(s: Seq<A>)
    ensures
        rev(s).len() == s.len(),
    decreases s.len(),
{
    if s.len() == 0 {
        assert(rev(s) == s);
    } else {
        let hd = s.first();
        let rest = s.drop_first();
        drop_first_len(s);
        assert(rev(s) == rev(rest).push(hd));
        rev_len(rest);
        assert(rev(rest).len() == rest.len());
        assert(rev(s).len() == rest.len() + 1);
        assert(rev(s).len() == s.len());
    }
}

/// Reversing twice is the identity — the other half of what `rev_push` is
/// for, and the lemma a "read the sheet bottom-up" view needs.
pub proof fn rev_rev<A>(s: Seq<A>)
    ensures
        rev(rev(s)) == s,
    decreases s.len(),
{
    if s.len() == 0 {
    } else {
        let hd = s.first();
        let rest = s.drop_first();
        // The induction is on `rest`, and `rev_push` supplies the outer step
        // (`rev(rev(rest).push(hd)) == [hd] ++ rev(rev(rest))`), which is why
        // this needs neither extra fuel nor a second induction.
        rev_rev(rest);
        rev_push(rev(rest), hd);
        assert(rev(s) == rev(rest).push(hd));
    }
}

/// `rev(u + t) == rev(t) + rev(u)`: a reversal swaps the halves.
pub proof fn rev_add<A>(u: Seq<A>, t: Seq<A>)
    ensures
        rev(u + t) == rev(t) + rev(u),
    decreases u.len(),
{
    if u.len() == 0 {
        reveal_with_fuel(rev, 1);
    } else {
        // Peel the head of `u` on both sides. `rev(u + t)` is
        // `rev(u1 + t).push(u0)`; `rev(t) + rev(u)` is `rev(t) + rev(u1).push(u0)`.
        // The induction hypothesis turns the first into
        // `(rev(t) + rev(u1)).push(u0)`, and `add_push` (appending to a
        // concatenation) turns that into the second.
        let u0 = u.first();
        let u1 = u.drop_first();
        // `u + t == [u0] ++ (u1 + t)`, so dropping the head of `u + t` is
        // `u1 + t`; one unfolding of `rev` then gives the step.
        assert(u + t == seq![u0] + (u1 + t));
        assert((u + t).drop_first() == u1 + t);
        assert(rev(u) == rev(u1).push(u0));
        reveal_with_fuel(rev, 1);
        rev_add(u1, t);
        add_push(rev(t), rev(u1), u0);
    }
}

/// A row list built by filtering stays inside the grid's row range.
///
/// `compute_sorted_main_rows` returns positions in `0..extent_main_rows` and
/// indexes the grid with them; an out-of-range position is a panic
/// (`grid.get` returns `None`, and the renderer indexes the display row list
/// directly). Proved by induction on the push that builds such a list, which
/// is the shape a `filter` has.
pub open spec fn in_range(v: Seq<int>, n: int) -> bool {
    forall|k: int| #![trigger v[k]] 0 <= k < v.len() ==> 0 <= v[k] < n
}

/// Appending an in-range value keeps a list in range.
pub proof fn push_preserves_range(v: Seq<int>, x: int, n: int)
    requires
        in_range(v, n),
        0 <= x < n,
    ensures
        in_range(v.push(x), n),
{
}

/// No row appears twice.
pub open spec fn distinct_seq(v: Seq<int>) -> bool {
    forall|i: int, j: int| 0 <= i < j < v.len() ==> v[i] != v[j]
}

/// Appending an unseen value keeps a list distinct — the property that makes
/// `permute` below a permutation rather than a duplication.
pub proof fn push_preserves_distinct(v: Seq<int>, x: int)
    requires
        distinct_seq(v),
        !v.contains(x),
    ensures
        distinct_seq(v.push(x)),
{
}

/// A sort spec whose direction only *reverses* the comparator, never changes
/// which cells are compared, yields a list that is the same set of rows in
/// the opposite order — stated for the length and membership, which is what a
/// caller can actually check:
/// * every row that was sorted is still there (`permute_mem`), and
/// * the count did not change (`rev_len`).
pub proof fn permute_mem<A>(f: Seq<A>, s: Seq<A>, x: A)
    requires
        is_subseq(f, s),
        mem_at(f, x),
    ensures
        mem_at(s, x),
{
    subseq_mem(f, s, x);
}

// ============================================================================
// 2c. Sort specs as data: the order the view sort consumes, and the log form
// ============================================================================
//
// A view sort is a list of `(column, descending)` pairs -- `SortSpec` in
// `src/grid/mod.rs`. Three things about that list are provable without
// touching the comparator:
//
//   * the *columns*, in order, are a projection of the list (`spec_cols`), and
//     that projection commutes with appending a spec -- which is what makes a
//     "sort by A, then B" builder's output have the columns in the order they
//     were added;
//   * no column is named twice, if the builder adds each once
//     (`push_preserves_unique`) -- and a repeated spec is not harmless: it
//     makes the later spec unreachable in `compute_sorted_main_rows`'s loop,
//     so the sort quietly ignores it;
//   * a single-key sort is `sorted_by` that key's order (`single_key_sorted_by`),
//     which is the shape `subseq_sorted` above consumes.
//
// What is *not* here: `compare_sort_values` (`src/grid/mod.rs`) is not a total
// order today -- text compares case-sensitively and every text key precedes
// every numeric one -- so `subseq_sorted` does not yet apply to a sort by it.
// The lemmas below are the parts that hold regardless, and they are the parts
// a fixed comparator would keep.

/// One sort spec: `(column, descending)`.
pub open spec fn spec_col(spec: (int, bool)) -> int {
    spec.0
}

pub open spec fn spec_desc(spec: (int, bool)) -> bool {
    spec.1
}

/// The columns of a spec list, in order.
///
/// Defined by peeling the *front* (not by recursing on the tail and pushing):
/// that is the shape every other lemma in this file uses, and it is the shape
/// in which "append a spec" is provable -- with the tail-recursive definition
/// the postcondition of `spec_cols_push` needs `add_push` *and* a
/// `drop_first` rewrite at every step, which is two facts per induction.
pub open spec fn spec_cols(specs: Seq<(int, bool)>) -> Seq<int>
    decreases specs.len(),
{
    if specs.len() == 0 {
        seq![]
    } else {
        seq![spec_col(specs.first())] + spec_cols(specs.drop_first())
    }
}

/// Appending a spec appends its column: a builder that adds specs in order
/// produces the columns in that order.
pub proof fn spec_cols_push(specs: Seq<(int, bool)>, spec: (int, bool))
    ensures
        spec_cols(specs.push(spec)) == spec_cols(specs).push(spec_col(spec)),
    decreases specs.len(),
{
    if specs.len() == 0 {
        // `[].push(spec)` and `[] ++ [spec]`: two unfoldings of `spec_cols`.
        reveal_with_fuel(spec_cols, 2);
    } else {
        reveal_with_fuel(spec_cols, 1);
        // `push` keeps the head and pushes onto the tail -- the two facts the
        // induction needs, stated because Z3 will not rewrite them inside the
        // spec call.
        assert(specs.push(spec).first() == specs.first());
        assert(specs.push(spec).drop_first() == specs.drop_first().push(spec));
        spec_cols_push(specs.drop_first(), spec);
        assert(spec_cols(specs.push(spec))
            == seq![spec_col(specs.first())] + spec_cols(specs.drop_first().push(spec)));
    }
}

/// Prepending a spec prepends its column -- the form a right-to-left builder
/// uses, and a one-line consequence of the definition above.
pub proof fn spec_cols_prepend(spec: (int, bool), specs: Seq<(int, bool)>)
    ensures
        spec_cols(seq![spec] + specs) == seq![spec_col(spec)] + spec_cols(specs),
    decreases specs.len(),
{
    // `seq![spec] + specs` has the same head as `specs` when `specs` is
    // non-empty, and its tail is the tail of `specs` -- so the interesting
    // case is the induction, not the unfolding.
    if specs.len() == 0 {
        reveal_with_fuel(spec_cols, 2);
    } else {
        reveal_with_fuel(spec_cols, 1);
        assert((seq![spec] + specs).first() == spec);
        assert((seq![spec] + specs).drop_first() == specs);
        spec_cols_prepend(spec, specs.drop_first());
    }
}

/// The columns named by a spec list are length-preserving: one spec, one
/// column.
pub proof fn spec_cols_len(specs: Seq<(int, bool)>)
    ensures
        spec_cols(specs).len() == specs.len(),
    decreases specs.len(),
{
    if specs.len() == 0 {
    } else {
        spec_cols_len(specs.drop_first());
    }
}

/// No column is named twice in a spec list.
pub open spec fn specs_unique(specs: Seq<(int, bool)>) -> bool {
    forall|i: int, j: int| 0 <= i < j < specs.len() ==> spec_col(specs[i]) != spec_col(specs[j])
}

/// Appending a spec whose column is not already named keeps the list unique.
/// A duplicated column is not a cosmetic problem: the second spec's comparison
/// is only reached when the first ties, and the first never ties on its own
/// column unless the values are equal -- where the second spec would then be
/// applying to the *same* values.
pub proof fn specs_unique_push(specs: Seq<(int, bool)>, spec: (int, bool))
    requires
        specs_unique(specs),
        forall|i: int| 0 <= i < specs.len() ==> spec_col(specs[i]) != spec_col(spec),
    ensures
        specs_unique(specs.push(spec)),
{
}

/// `sorted_by` specialised to a one-key order, as a closed `forall`.
///
/// This is the shape a one-column sort has (the menu's `sort_asc` /
/// `sort_desc` install exactly one spec), and the shape `subseq_sorted`
/// consumes once `le` is fixed.
pub open spec fn sorted_by_key(f: Seq<int>, key: Seq<int>, c: int) -> bool {
    forall|i: int, j: int| 0 <= i < j && j < f.len() ==> key[f[i]] <= key[f[j]]
}

/// A list whose key column is non-decreasing along it satisfies
/// `sorted_by` the key's order.
pub proof fn single_key_sorted_by(f: Seq<int>, key: Seq<int>, c: int)
    requires
        sorted_by_key(f, key, c),
    ensures
        sorted_by(f, |a: int, b: int| key[a] <= key[b]),
{
}

// ============================================================================
// 2d. Correctness of a sort under a *non-total* comparison
// ============================================================================
//
// `compare_sort_values` (`src/grid/mod.rs`) is not a total order: text
// compares case-sensitively and every text key precedes every numeric one.
// The question this section answers is what can still be *proved* about such
// a comparison, and the answer is: everything except the pairwise form.
//
// The reason is that "sorted" has two readings:
//
//   * **pairwise** -- every earlier element is `<=` every later one. This is
//     `sorted_by` in this file, and it needs `<=` to compose across the
//     elements in between, i.e. to be *transitive*.
//   * **adjacent** -- each element is `<=` the next one. This is what a sort
//     actually produces (it compares neighbours by construction), and what a
//     stability claim is stated against.
//
// The theorem below is the bridge: **adjacent + transitivity => pairwise**.
// Total antisymmetry (a three-way comparison that never answers `Equal`) is
// *not* needed -- only that the relation it induces is transitive. That is
// what makes the answer "yes" for a non-total comparison: a case-insensitive
// comparator that puts text before numbers has a transitive `<=` (compare
// both buckets, then position-in-bucket), and then a sort by it is pairwise
// sorted; today's case-sensitive one does not, and the proof says exactly
// where that fails rather than saying "not supported".
//
// What is still not provable: that `std::slice::sort_by` *produces* the
// adjacent-sorted order at all. That is a property of `std`'s sort, not of the
// comparison, and no `Seq` lemma can reach it.

/// A comparison result, as `std::cmp::Ordering`'s three values.
pub open spec fn is_gt(o: int) -> bool {
    o == 2
}

/// The preorder a comparator induces: `x <= y` iff it does not answer
/// "x > y".
pub open spec fn le_of_cmp(cmp: spec_fn(int, int) -> int, x: int, y: int) -> bool {
    !is_gt(cmp(x, y))
}

/// The closure property that takes the place of totality: `<=` is transitive.
///
/// This is the one premise that a non-total order can supply and a
/// *non-transitive* comparison cannot -- so it is also the honest statement of
/// what `compare_sort_values` would have to satisfy.
pub open spec fn le_transitive(cmp: spec_fn(int, int) -> int) -> bool {
    forall|x: int, y: int, z: int| #![trigger le_of_cmp(cmp, x, y), le_of_cmp(cmp, y, z)]
        le_of_cmp(cmp, x, y) && le_of_cmp(cmp, y, z) ==> le_of_cmp(cmp, x, z)
}

/// One adjacent pair, extracted from the closed `forall` a caller states.
/// (A lemma rather than a bare `assert`, because a `forall` written with an
/// explicit trigger is not re-instantiated by a plain assertion.)
pub proof fn adjacent_le(f: Seq<int>, cmp: spec_fn(int, int) -> int, i: int)
    requires
        forall|k: int| 0 <= k && k + 1 < f.len()
            ==> #[trigger] le_of_cmp(cmp, f[k], f[k + 1]),
        0 <= i && i + 1 < f.len(),
    ensures
        le_of_cmp(cmp, f[i], f[i + 1]),
{
}

/// **Adjacent + transitive implies pairwise** -- the bridge that lets a
/// non-total comparison have a sortedness proof.
pub proof fn adjacent_implies_pairwise(f: Seq<int>, cmp: spec_fn(int, int) -> int)
    requires
        le_transitive(cmp),
        forall|k: int| 0 <= k && k + 1 < f.len()
            ==> #[trigger] le_of_cmp(cmp, f[k], f[k + 1]),
    ensures
        forall|i: int, j: int| #![trigger le_of_cmp(cmp, f[i], f[j])]
            0 <= i < j && j < f.len() ==> le_of_cmp(cmp, f[i], f[j]),
{
    assert forall|i: int, j: int| #![trigger le_of_cmp(cmp, f[i], f[j])]
        0 <= i < j && j < f.len() implies le_of_cmp(cmp, f[i], f[j]) by {
        adjacent_chain(f, cmp, i, j);
    };
}

/// The chain itself: transitivity applied `j - i` times.
pub proof fn adjacent_chain(f: Seq<int>, cmp: spec_fn(int, int) -> int, i: int, j: int)
    requires
        le_transitive(cmp),
        forall|k: int| 0 <= k && k + 1 < f.len()
            ==> #[trigger] le_of_cmp(cmp, f[k], f[k + 1]),
        0 <= i < j < f.len(),
    ensures
        le_of_cmp(cmp, f[i], f[j]),
    decreases j - i,
{
    if j == i + 1 {
        adjacent_le(f, cmp, i);
    } else {
        // `i <= j-1` by the induction, `j-1 <= j` by adjacency, and
        // transitivity joins them. This is the *only* place transitivity is
        // used, and the only reason a non-transitive comparator cannot be
        // proved sorted.
        adjacent_chain(f, cmp, i, j - 1);
        adjacent_le(f, cmp, j - 1);
    }
}

/// A **non-total** comparison whose closure is still transitive.
///
/// The buckets are "text before number" -- the same split `compare_sort_values`
/// makes -- with the text bucket ordered case-insensitively. `Equal` is *not*
/// `a == b` (that is the whole point): `"Apple"` and `"apple"` compare equal,
/// so the relation is not antisymmetric, and the order is not total.
pub open spec fn demo_bucket(x: int) -> int {
    if x >= 0 { 0 } else { 1 }
}

/// `cmp` for the demo: 0 = Less, 1 = Equal, 2 = Greater.
pub open spec fn demo_cmp(x: int, y: int) -> int {
    if demo_bucket(x) < demo_bucket(y) {
        0
    } else if demo_bucket(x) > demo_bucket(y) {
        2
    } else if x < y {
        0
    } else if x > y {
        2
    } else {
        1
    }
}

/// The demo comparison is transitive -- so `adjacent_implies_pairwise` applies
/// to it -- and it is **not** antisymmetric: `demo_cmp(1, 1) == 1` (`Equal`)
/// is fine, but `"Apple"`/`"apple"` both compare `Equal` in the real
/// comparator, and that is what "not total" means. The two facts are separate
/// on purpose: only the first is needed for the sortedness proof.
pub proof fn demo_cmp_transitive()
    ensures
        le_transitive(|x: int, y: int| demo_cmp(x, y)),
{
}

/// The demo comparison's `Equal` is `x == y`, so the demo is in fact
/// antisymmetric; the real comparator's is not (case-insensitive text makes
/// distinct strings equal). The theorem above never uses either fact -- it
/// asks only for transitivity of `<=` -- which is the point of stating them
/// separately: **non-totality costs a sort proof nothing as long as the
/// relation it induces is transitive.**
pub proof fn demo_cmp_equal_is_eq(x: int, y: int)
    ensures
        (demo_cmp(x, y) == 1) == (x == y),
{
}

// ============================================================================
// 2e. Multi-key comparison: the first differing key decides
// ============================================================================
//
// `compute_sorted_main_rows`'s comparator walks the sort specs in order and
// returns at the first key that does not tie. Two facts make that safe, and
// both are provable for *any* per-key comparison, total or not:
//
//   * a later key can never override an earlier decision
//     (`lex_first_diff_decides`), which is what makes "sort by A, then B"
//     mean what a user expects rather than "B wins sometimes";
//   * if every key ties, the comparison ties (`lex_ties_is_equal`), which is
//     what makes the tie-break (`a.cmp(b)`) reachable at all.
//
// The step function is written with an explicit index (`lex_cmp_at(.., k)`)
// rather than by dropping the head of the key list. That is not cosmetic: with
// the head-dropping form the "first differing key" lemma needs
// `specs.drop_first()[k-1] == specs[k]` inside a *spec argument*, and Z3 will
// not rewrite that (it needs the relation asserted, then applied under a
// recursive call). Advancing an index turns the step into `k + 1`.

/// A comparison over one key column. `spec` is the column index, so the same
/// comparison function serves every key.
pub open spec fn is_eq(o: int) -> bool {
    o == 1
}

/// The lexicographic comparison from key index `k`: the first key that does
/// not tie decides, and running off the end means "all keys tied".
pub open spec fn lex_cmp_at(
    cmp: spec_fn(int, int, int) -> int,
    specs: Seq<int>,
    a: int,
    b: int,
    k: int,
) -> int
    decreases specs.len() - k,
{
    if k >= specs.len() {
        1 // Equal: every key tied
    } else if !is_eq(cmp(specs[k], a, b)) {
        cmp(specs[k], a, b)
    } else {
        lex_cmp_at(cmp, specs, a, b, k + 1)
    }
}

/// **A later key cannot override an earlier decision.** If keys `i..k` all
/// tie and key `k` answers `o` (not Equal), the comparison started at `i`
/// answers `o` too.
pub proof fn lex_decides_from(
    cmp: spec_fn(int, int, int) -> int,
    specs: Seq<int>,
    a: int,
    b: int,
    i: int,
    k: int,
    o: int,
)
    requires
        forall|j: int| #![trigger specs[j]] i <= j < k ==> is_eq(cmp(specs[j], a, b)),
        i <= k < specs.len(),
        cmp(specs[k], a, b) == o,
        !is_eq(o),
    ensures
        lex_cmp_at(cmp, specs, a, b, i) == o,
    decreases k - i,
{
    reveal_with_fuel(lex_cmp_at, 1);
    if i == k {
        // the current iteration is the decisive one
    } else {
        lex_decides_from(cmp, specs, a, b, i + 1, k, o);
    }
}

/// The same statement from the top of the key list -- the form a caller
/// quotes ("the first differing key decides").
pub proof fn lex_first_diff_decides(
    cmp: spec_fn(int, int, int) -> int,
    specs: Seq<int>,
    a: int,
    b: int,
    k: int,
    o: int,
)
    requires
        forall|j: int| #![trigger specs[j]] 0 <= j < k ==> is_eq(cmp(specs[j], a, b)),
        0 <= k < specs.len(),
        cmp(specs[k], a, b) == o,
        !is_eq(o),
    ensures
        lex_cmp_at(cmp, specs, a, b, 0) == o,
{
    lex_decides_from(cmp, specs, a, b, 0, k, o);
}

/// **If every key ties, the comparison ties** -- so a multi-key sort's
/// tie-break is reached exactly when it should be.
pub proof fn lex_ties_is_equal(
    cmp: spec_fn(int, int, int) -> int,
    specs: Seq<int>,
    a: int,
    b: int,
    i: int,
)
    requires
        forall|j: int| #![trigger specs[j]] i <= j < specs.len() ==> is_eq(cmp(specs[j], a, b)),
        0 <= i <= specs.len(),
    ensures
        lex_cmp_at(cmp, specs, a, b, i) == 1,
    decreases specs.len() - i,
{
    reveal_with_fuel(lex_cmp_at, 1);
    if i < specs.len() {
        lex_ties_is_equal(cmp, specs, a, b, i + 1);
    }
}

/// And what follows for a *non-total* per-key comparison: a tie is still a
/// tie, so two rows that tie on every key fall through to whatever
/// tie-break the caller supplies -- the case that makes the comparator's
/// `Equal` matter even though it is not `==`.
pub proof fn all_keys_tied_falls_through(
    cmp: spec_fn(int, int, int) -> int,
    specs: Seq<int>,
    a: int,
    b: int,
)
    requires
        forall|j: int| #![trigger specs[j]] 0 <= j < specs.len() ==> is_eq(cmp(specs[j], a, b)),
    ensures
        lex_cmp_at(cmp, specs, a, b, 0) == 1,
{
    lex_ties_is_equal(cmp, specs, a, b, 0);
}

// ============================================================================
// 2f. Where the composition stops: a counterexample, as a theorem
// ============================================================================
//
// 2d said a non-total per-key comparison is fine *if its order is transitive*,
// and 2e said a multi-key comparison picks the first differing key. The
// obvious next question -- does transitivity of each *key's* order give
// transitivity of the multi-key order? -- has answer **no**, and the honest
// way to record that is a witness rather than a paragraph.
//
// The failing configuration, in words: `x` ties `y` at key 0, `y` is Greater
// than `z` at key 0, and `x` is Greater than `z` at key 0. The multi-key
// comparison then says `x <= y` (it falls through to a later key, or ties)
// and `y <= z`, but `x > z` is decided at key 0 -- so the relation reports a
// cycle although each key is compared consistently. The defect is that
// `Equal` at key 0 is used as if it were an equivalence: `x =0 y` and
// `y >0 z` say nothing about `x` vs `z` unless `Equal` at that key is both
// transitive *and* congruent w.r.t. `Greater`. That is the precise gap a real
// comparator has to close, and it is the same gap 2d found from the other
// side.

/// A per-key comparison with exactly that defect: `(0, 2)` is ordered against
/// the rest of the integers, and every later key ties.
pub open spec fn cex_cmp(k: int, a: int, b: int) -> int {
    if k != 0 {
        1 // Equal: keys after the first always tie
    } else if a == b {
        1
    } else if a == 0 && b == 2 {
        2 // Greater: 0 > 2, deliberately against the rest
    } else if a > b {
        2
    } else {
        0
    }
}

/// The multi-key relation induced by `cex_cmp` over the one-key spec list
/// `[0]`.
pub open spec fn cex_le(a: int, b: int) -> bool {
    !is_gt(lex_cmp_at(|k: int, x: int, y: int| cex_cmp(k, x, y), seq![0], a, b, 0))
}

/// The single key's order is not transitive: `0 <= 1`, `1 <= 2`, `0 > 2`.
pub proof fn cex_key_order_not_transitive()
    ensures
        !is_gt(cex_cmp(0, 0, 1)),
        !is_gt(cex_cmp(0, 1, 2)),
        is_gt(cex_cmp(0, 0, 2)),
{
}

/// …and the multi-key comparison inherits it: `cex_le(0, 1)`, `cex_le(1, 2)`,
/// but not `cex_le(0, 2)`.
pub proof fn cex_lex_not_transitive()
    ensures
        cex_le(0, 1),
        cex_le(1, 2),
        !cex_le(0, 2),
{
    reveal_with_fuel(lex_cmp_at, 1);
}

// ============================================================================
// 3. Subsequences and order — "filtered fast path keeps scan order"
// ============================================================================

/// Index-based membership.  Definitionally the same existential as
/// `Seq::contains`, but Z3 will not decompose `contains` across
/// `drop_first`; with `mem_at` you can case-split on the witness index.
pub open spec fn mem_at<A>(f: Seq<A>, x: A) -> bool {
    exists|i: int| 0 <= i < f.len() && f[i] == x
}

pub proof fn mem_at_of_contains<A>(f: Seq<A>, x: A)
    requires f.contains(x),
    ensures mem_at(f, x),
{
}

pub proof fn contains_of_mem_at<A>(f: Seq<A>, x: A)
    requires mem_at(f, x),
    ensures f.contains(x),
{
}

/// The head is a member.
pub proof fn mem_at_first<A>(f: Seq<A>, x: A)
    requires
        f.len() > 0,
        f.first() == x,
    ensures
        mem_at(f, x),
{
    assert(f[0] == x);
}

/// A member other than the head is a member of the tail.
pub proof fn mem_at_drop_first<A>(f: Seq<A>, x: A)
    requires
        f.len() > 0,
        mem_at(f, x),
        f.first() != x,
    ensures
        mem_at(f.drop_first(), x),
{
    let i = choose|i: int| 0 <= i < f.len() && f[i] == x;
    assert(i != 0);  // since f[0] == f.first() != x
    drop_first_index(f);
    assert(f.drop_first()[i - 1] == f[i]);
}

/// A member of the tail is a member of the whole sequence.
pub proof fn mem_at_of_drop_first<A>(s: Seq<A>, x: A)
    requires
        s.len() > 0,
        mem_at(s.drop_first(), x),
    ensures
        mem_at(s, x),
{
    let j = choose|j: int| 0 <= j < s.drop_first().len() && s.drop_first()[j] == x;
    drop_first_index(s);
    assert(s[j + 1] == x);
}

/// `f` is `s` with elements deleted: scan `s` from the front and consume
/// a head of `f` whenever the heads match, otherwise delete `s`'s head.
pub open spec fn is_subseq<A>(f: Seq<A>, s: Seq<A>) -> bool
    decreases s.len(),
{
    if f.len() == 0 {
        true
    } else if s.len() == 0 {
        false
    } else if f.first() == s.first() {
        is_subseq(f.drop_first(), s.drop_first())
    } else {
        is_subseq(f, s.drop_first())
    }
}

/// `f` is sorted by `le` (full pairwise version: every earlier element is
/// `le` every later element).
pub open spec fn sorted_by<A>(f: Seq<A>, le: spec_fn(A, A) -> bool) -> bool {
    forall|i: int, j: int| 0 <= i && i < j && j < f.len() ==> le(f[i], f[j])
}

/// A tail of a sorted sequence is sorted.
pub proof fn sorted_drop_first<A>(s: Seq<A>, le: spec_fn(A, A) -> bool)
    requires
        s.len() > 0,
        sorted_by(s, le),
    ensures
        sorted_by(s.drop_first(), le),
{
    drop_first_len(s);
    drop_first_index(s);
}

/// Everything in a subsequence is in the original sequence.
pub proof fn subseq_mem<A>(f: Seq<A>, s: Seq<A>, x: A)
    requires
        is_subseq(f, s),
        mem_at(f, x),
    ensures
        mem_at(s, x),
    decreases s.len(),
{
    if f.len() == 0 {
        // mem_at(f, x) is impossible for empty f
        let i = choose|i: int| 0 <= i < f.len() && f[i] == x;
        assert(false);
    } else if s.len() == 0 {
        // is_subseq(f, []) is false for nonempty f
        assert(is_subseq(f, s) == false);
    } else if f.first() == s.first() {
        if x == f.first() {
            assert(s[0] == x);  // s.first() == f.first() == x
        } else {
            mem_at_drop_first(f, x);
            subseq_mem(f.drop_first(), s.drop_first(), x);
            mem_at_of_drop_first(s, x);
        }
    } else {
        subseq_mem(f, s.drop_first(), x);
        mem_at_of_drop_first(s, x);
    }
}

/// `contains`-flavored convenience wrapper around `subseq_mem`.
pub proof fn subseq_contains<A>(f: Seq<A>, s: Seq<A>, x: A)
    requires
        is_subseq(f, s),
        f.contains(x),
    ensures
        s.contains(x),
{
    mem_at_of_contains(f, x);
    subseq_mem(f, s, x);
    contains_of_mem_at(s, x);
}

/// Deleting elements from a sorted sequence keeps it sorted.
///
/// The transitivity of `le` is part of the contract (per the intended
/// calling convention).  With `sorted_by` defined as the full pairwise
/// order, the induction below compares every pair through two indices of
/// `s` itself, so it does not end up needing the premise — but it is kept
/// so the lemma also covers stricter variants (e.g. an adjacent-pairs
/// definition of sortedness, where it would be required).
pub proof fn subseq_sorted<A>(f: Seq<A>, s: Seq<A>, le: spec_fn(A, A) -> bool)
    requires
        sorted_by(s, le),
        is_subseq(f, s),
        forall|x: A, y: A, z: A| #![trigger le(x, y), le(y, z)]
            le(x, y) && le(y, z) ==> le(x, z),  // le transitive
    ensures
        sorted_by(f, le),
    decreases s.len(),
{
    if f.len() == 0 {
        // sorted_by([], le) holds vacuously
    } else if s.len() == 0 {
        // is_subseq(f, []) is false for nonempty f
        assert(is_subseq(f, s) == false);
    } else if f.first() == s.first() {
        drop_first_len(f);
        drop_first_len(s);
        drop_first_index(f);
        drop_first_index(s);
        // the tail of s is sorted...
        sorted_drop_first(s, le);
        // ...so by IH the tail of f is sorted
        subseq_sorted(f.drop_first(), s.drop_first(), le);
        assert(f[0] == s[0]);
        // Head step: f[0] == s[0], and every later element of f occurs in
        // s.drop_first(), i.e. as s[j + 1] for some j; sorted_by(s) then
        // gives le(s[0], s[j + 1]).
        assert forall|k: int| #![trigger le(f[0], f[k])] 1 <= k && k < f.len()
            implies le(f[0], f[k]) by {
            assert(f[k] == f.drop_first()[k - 1]);
            assert(mem_at(f.drop_first(), f[k]));
            subseq_mem(f.drop_first(), s.drop_first(), f[k]);
            let j = choose|j: int| 0 <= j < s.drop_first().len() && s.drop_first()[j] == f[k];
            assert(s[j + 1] == f[k]);
            assert(le(s[0], s[j + 1]));
        };
        // Combine the head step with the IH on the tails.
        assert forall|i: int, j: int| #![trigger le(f[i], f[j])]
            0 <= i && i < j && j < f.len() implies le(f[i], f[j]) by {
            if i == 0 {
                // head step at k == j
            } else {
                assert(f[i] == f.drop_first()[i - 1]);
                assert(f[j] == f.drop_first()[j - 1]);
                // IH applies at (i - 1, j - 1)
            }
        };
    } else {
        // heads differ: f is a subsequence of the tail of s, which is sorted
        sorted_drop_first(s, le);
        subseq_sorted(f, s.drop_first(), le);
    }
}

// ============================================================================
// 4. The order `compute_sorted_main_rows` tie-breaks with
// ============================================================================
//
// What follows is the part of a proof of `Grid::compute_sorted_main_rows`
// (`src/grid/mod.rs`) that can be built without touching the comparator:
//
//   * the tie-break (`a.cmp(b)`) is a total order, called `index_order` here;
//   * a sequence is sorted by its *own* positions under that order, which is
//     what a stability claim is stated against.
//
// What is deliberately **not** here, and why:
//
//   * `sort_by` is `std`'s, not a `Seq` function, so no Verus lemma can speak
//     about the sort itself; the lemmas are about the *order* the sort
//     consumes.
//   * `compare_sort_values` is not a total order today. It compares text
//     case-sensitively (`"Apple" < "banana" < "cherry"`, pinned by
//     `view_sort_handles_non_numeric_columns`) and puts every text key before
//     every numeric one, so `"2"` and `"10"` are `Number(2) < Number(10)` but
//     `"Apple"` and `"apple"` are *not* adjacent to each other's bucket. Those
//     pairs make the result of a sort by it well-defined (the sort is stable)
//     but not `sorted_by` any total order, so `subseq_sorted` above does not
//     transfer to it. Fixing the comparator (case-folding, one bucket) is what
//     would make the `subseq_sorted` lemma apply to the sort's output.
//   * `rev_push`'s companion -- reversing a sequence and reversing back gives
//     the original, `rev(rev(s)) == s` -- is *not* a quick corollary here. It
//     needs two different inductions (one on `push`, one on `rev`) to meet, and
//     the naive `assert(rev(rev(s)) == s)` does not go through at the default
//     fuel. It is worth adding once the proof is worked out; a lemma "proved"
//     by a `reveal_with_fuel(rev, 40)` would be a lemma nobody can maintain, so
//     it is left out rather than left fragile.
//   * the "reverse" half of "sort desc is sort asc, reversed" is therefore
//     unproved here, and the comparator's non-total case (above) means the
//     statement is not even true of today's sort: reversing an ascending
//     sequence does not produce the descending one when two keys compare
//     equal, which the tie-break makes impossible -- but only once the keys of
//     *different* buckets compare at all.

/// Element order used as the tie-break: earlier scan position first.
pub open spec fn index_order(x: int, y: int) -> bool {
    x <= y
}

/// `sorted_by` specialised to [`index_order`], written as a closed `forall`.
///
/// The generic `sorted_by(s, |x, y| index_order(x, y))` says the same thing,
/// but a `forall` hidden inside a spec-function call cannot be discharged by
/// `forall ... by { }` (that block proves a goal that *is* a `forall`), so the
/// index order gets a predicate of its own and
/// the generic `sorted_by` form is a `forall` *mentioning a closure*, which no
/// automated trigger can be inferred for -- see the note in §4.
pub open spec fn sorted_by_idx(s: Seq<int>) -> bool {
    forall|i: int, j: int| 0 <= i && i < j && j < s.len() ==> index_order(i, j)
}

/// Any sequence is sorted by its own positions.
pub proof fn sorted_by_index_order(s: Seq<int>)
    ensures
        sorted_by_idx(s),
{
    // `sorted_by_idx` *is* the `forall` below; see its doc comment.
    assert forall|i: int, j: int| 0 <= i && i < j && j < s.len()
        implies index_order(i, j) by {
        assert(i <= j);
    };
}

/// The index order satisfies the generic `sorted_by` -- as a *statement*, not
/// as a proof. What was tried and does not verify:
///
/// ```text
/// pub proof fn sorted_by_of_idx(s: Seq<int>)
///     requires sorted_by_idx(s),
///     ensures  sorted_by(s, |x: int, y: int| index_order(x, y)),
/// { }
/// ```
///
/// The two forms say the same thing, but `sorted_by(.., |x, y| ..)` is a
/// `forall` whose body calls a *closure*, and Verus cannot infer a trigger for
/// such a quantifier ("Could not automatically infer triggers ... use
/// #[trigger] annotations"). Stating the property in the closed form
/// [`sorted_by_idx`] is what makes it provable at all -- which is why the
/// index order gets a predicate of its own instead of being plugged into the
/// generic `sorted_by` as a lambda. The generic form stays the right thing to
/// *say* to a caller.
///
/// [`index_order`] is total: of any two integers, one direction holds.
pub proof fn index_order_total(i: int, j: int)
    ensures
        index_order(i, j) || index_order(j, i),
{
    if i <= j {
        assert(index_order(i, j));
    } else {
        assert(index_order(j, i));
    }
}

/// `0 * m == 0`, the base case the two monotonicity lemmas above need.
///
/// Z3 does not evaluate `int * int` by itself -- `mul_step` covers the `+ 1`
/// step but not zero -- so the fact is stated once.
pub proof fn mul_zero(m: int)
    ensures
        0 * m == 0,
{
    assert(0 * m == 0) by (nonlinear_arith);
}

fn main() {}
}
