//! Standalone Verus lemmas that Z3 will not find on its own.
//!
//! Three groups:
//!   1. Integer arithmetic — multiplication monotonicity / nonnegativity
//!      (Z3 fails on these with or without `nonlinear_arith`, so we induct
//!      and discharge only single distributivity steps nonlinearly).
//!   2. Sequence reversal — `rev_push` / `rev_len`, generic over the element
//!      type.  `rev` is defined by peeling the *first* element, so pushing at
//!      the end needs induction; the base case evaluates `rev([a])`, which
//!      costs *two* unfoldings of `rev` while the default fuel is one —
//!      hence `reveal_with_fuel(rev, 2)`.
//!   3. Subsequences and order — the reusable core of "a filtered fast path
//!      keeps scan order": `mem_at` (a membership predicate that decomposes
//!      over `drop_first`, unlike `Seq::contains`), `is_subseq`, and
//!      `subseq_mem` / `subseq_sorted` (over `vstd::relations::sorted_by`).
//!
//! Verified with Verus 0.2026.10.18:
//! ```text
//! specs/verus/check.sh
//! ```
//!
//! **This file is the older of the two in this directory.** Its header used to
//! say "Verus 0.2026.06.07 (classic `vstd` `Seq` API)", and the imports below
//! (`vstd::relations::sorted_by`, `vstd::transitive`, `vstd::seq::*`,
//! `vstd::seq_lib::*`) are the shape that API had. Current vstd still exports
//! them, and this file still verifies against it unchanged (21 lemmas), but
//! `scan_order_lemmas.rs` makes the same points without leaning on any of it.
//! Keep this one as the record of what the API looked like, and prefer
//! `scan_order_lemmas.rs` for anything new.
//!
//! Note on `subseq_sorted`: with the pairwise `sorted_by` and the greedy
//! `is_subseq` below, transitivity of `le` turns out *not* to be needed —
//! every element of `f` relates directly to its position in `s`.  (You only
//! need `transitive` if your `sorted_by` speaks about *adjacent* elements.)

use vstd::prelude::*;
use vstd::relations::sorted_by;
use vstd::relations::transitive;
use vstd::seq::*;
use vstd::seq_lib::*;

verus! {

// =====================================================================
// 1. Integer arithmetic
// =====================================================================

/// `0 <= a <= b` and `0 <= m` implies `a * m <= b * m`.
///
/// Z3 will not derive monotonicity of multiplication by itself (with or
/// without nonlinear arithmetic), so we induct on the gap `b - a`.  The only
/// nonlinear step is one distributivity fact, discharged in a
/// `by (nonlinear_arith)` block; everything else is linear.
pub proof fn mul_mono_const(a: int, b: int, m: int)
    requires
        0 <= a && a <= b,
        0 <= m,
    ensures
        a * m <= b * m,
    decreases b - a,
{
    if a < b {
        mul_mono_const(a, b - 1, m);
        // a*m <= (b-1)*m   (inductive hypothesis)
        assert(b * m == (b - 1) * m + m) by (nonlinear_arith);
        // (b-1)*m <= (b-1)*m + m   since m >= 0 — linear from here
    }
    // a == b: a*m <= a*m, trivially
}

/// `0 <= a` and `0 <= b` implies `0 <= a * b`, by induction on `a`.
pub proof fn mul_nonneg(a: int, b: int)
    requires
        0 <= a,
        0 <= b,
    ensures
        0 <= a * b,
    decreases a,
{
    if a > 0 {
        mul_nonneg(a - 1, b);
        // 0 <= (a-1)*b   (inductive hypothesis)
        assert(a * b == (a - 1) * b + b) by (nonlinear_arith);
        // (a-1)*b + b >= 0 since both summands are — linear from here
    }
    // a == 0: 0 * b == 0
}

/// `x <= y` and `y == z` implies `x <= z`.
///
/// Trivial for Z3 once stated, but calling the lemma forces the instance in
/// contexts where the terms are buried under triggers.
pub proof fn le_of_le_eq(x: int, y: int, z: int)
    requires
        x <= y,
        y == z,
    ensures
        x <= z,
{
}

/// `x <= y` implies `x + c <= y + c`.
pub proof fn le_add_const(x: int, y: int, c: int)
    requires
        x <= y,
    ensures
        x + c <= y + c,
{
}

// =====================================================================
// 2. Sequence reversal
// =====================================================================

/// Reverse of `s`: peel the first element, reverse the rest, append.
pub open spec fn rev<A>(s: Seq<A>) -> Seq<A>
    decreases s.len(),
{
    if s.len() == 0 {
        seq![]
    } else {
        rev(s.drop_first()).push(s[0])
    }
}

/// Any nonempty sequence is its head consed onto its tail.
pub proof fn seq_cons<A>(s: Seq<A>)
    requires
        s.len() > 0,
    ensures
        s == seq![s[0]] + s.drop_first(),
{
    let stl = s.drop_first();
    assert(stl.len() == s.len() - 1);
    assert((seq![s[0]] + stl).len() == s.len());
    assert forall|i: int| 0 <= i < s.len() implies (seq![s[0]] + stl)[i] == s[i] by {
        if i == 0 {
            assert((seq![s[0]] + stl)[0] == s[0]);
        } else {
            assert((seq![s[0]] + stl)[i] == stl[i - 1]);
            assert(stl[i - 1] == s[i]);
        }
    };
    assert(s =~= seq![s[0]] + stl);
}

/// `rev(s.push(a)) == [a] ++ rev(s)` — the workhorse for reasoning about
/// pushing onto a sequence across a reversal.
///
/// Induction on `s.len()`.  The base case evaluates `rev([a])`, which needs
/// *two* unfoldings of `rev` (`[a]` itself, then the `rev(seq![])` inside
/// it) while the default fuel is one, hence `reveal_with_fuel(rev, 2)`.
pub proof fn rev_push<A>(s: Seq<A>, a: A)
    ensures
        rev(s.push(a)) == seq![a] + rev(s),
    decreases s.len(),
{
    if s.len() == 0 {
        reveal_with_fuel(rev, 2);
        assert(s =~= seq![]);
        assert(s == seq![]);
        assert(s.push(a) =~= seq![a]);
        assert(s.push(a) == seq![a]);
        // rev([a]) == rev([].drop_first()).push(a) == seq![a].push(a) == [a]  (fuel 2)
        assert(rev(seq![a]) == seq![a]);
        assert(rev(s) == seq![]);
        assert(seq![a] + seq![] =~= seq![a]);
        assert(seq![a] + seq![] == seq![a]);
    } else {
        reveal_with_fuel(rev, 1);
        let hd = s[0];
        let stl = s.drop_first();
        assert(stl.len() == s.len() - 1);
        seq_cons(s);

        // (1) s.push(a) == [hd] ++ (stl.push(a))
        assert(s.push(a).len() == s.len() + 1);
        assert((seq![hd] + stl.push(a)).len() == s.len() + 1);
        assert forall|i: int| #![auto] 0 <= i < s.push(a).len() implies s.push(a)[i]
            == (seq![hd] + stl.push(a))[i] by {
            if i == 0 {
                assert(s.push(a)[0] == s[0]);
                assert((seq![hd] + stl.push(a))[0] == hd);
            } else if i < s.len() {
                assert(s.push(a)[i] == s[i]);
                assert(s[i] == (seq![hd] + stl)[i]);
                assert((seq![hd] + stl)[i] == stl[i - 1]);
                assert(stl[i - 1] == stl.push(a)[i - 1]);
                assert(stl.push(a)[i - 1] == (seq![hd] + stl.push(a))[i]);
            } else {
                assert(i == s.len());
                assert(s.push(a)[i] == a);
                assert((seq![hd] + stl.push(a))[i] == stl.push(a)[i - 1]);
                assert(stl.push(a)[i - 1] == a);
            }
        };
        assert(s.push(a) =~= seq![hd] + stl.push(a));
        assert(s.push(a) == seq![hd] + stl.push(a));

        // (2) dropping the head of [hd] ++ y yields y
        assert((seq![hd] + stl.push(a)).drop_first().len() == stl.push(a).len());
        assert forall|i: int| #![auto] 0 <= i < stl.push(a).len() implies (seq![hd]
            + stl.push(a)).drop_first()[i] == stl.push(a)[i] by {
            assert((seq![hd] + stl.push(a)).drop_first()[i] == (seq![hd] + stl.push(a))[i + 1]);
            assert((seq![hd] + stl.push(a))[i + 1] == stl.push(a)[i]);
        };
        assert((seq![hd] + stl.push(a)).drop_first() =~= stl.push(a));
        assert(s.push(a).drop_first() == stl.push(a));
        assert(s.push(a)[0] == hd);

        // (3) inductive hypothesis on the tail
        rev_push(stl, a);

        // (4) unfold rev once on both sides
        assert(rev(s.push(a)) == (seq![a] + rev(stl)).push(hd));
        assert(rev(s) == rev(stl).push(hd));

        // (5) push distributes over ++ on the right:
        //     ([a] ++ R).push(hd) == [a] ++ (R.push(hd))
        assert((seq![a] + rev(stl)).push(hd).len() == (seq![a] + rev(stl).push(hd)).len());
        assert forall|i: int| 0 <= i < (seq![a] + rev(stl)).push(hd).len() implies (seq![a]
            + rev(stl)).push(hd)[i] == (seq![a] + rev(stl).push(hd))[i] by {
            if i == 0 {
                assert((seq![a] + rev(stl)).push(hd)[0] == (seq![a] + rev(stl))[0]);
                assert((seq![a] + rev(stl))[0] == (seq![a] + rev(stl).push(hd))[0]);
            } else if i <= rev(stl).len() {
                assert((seq![a] + rev(stl)).push(hd)[i] == (seq![a] + rev(stl))[i]);
                assert((seq![a] + rev(stl))[i] == rev(stl)[i - 1]);
                assert((seq![a] + rev(stl).push(hd))[i] == rev(stl).push(hd)[i - 1]);
                assert(rev(stl).push(hd)[i - 1] == rev(stl)[i - 1]);
            } else {
                assert(i == 1 + rev(stl).len());
                assert((seq![a] + rev(stl)).push(hd)[i] == hd);
                assert((seq![a] + rev(stl).push(hd))[i] == rev(stl).push(hd)[i - 1]);
                assert(rev(stl).push(hd)[i - 1] == hd);
            }
        };
        assert((seq![a] + rev(stl)).push(hd) =~= seq![a] + rev(stl).push(hd));
        // combine (3), (4), (5):  rev(s.push(a)) == [a] ++ rev(stl).push(hd) == [a] ++ rev(s)
    }
}

/// `(rev(s)).len() == s.len()`.
pub proof fn rev_len<A>(s: Seq<A>)
    ensures
        (rev(s)).len() == s.len(),
    decreases s.len(),
{
    if s.len() > 0 {
        let stl = s.drop_first();
        assert(stl.len() == s.len() - 1);
        rev_len(stl);
        // rev(s) == rev(stl).push(s[0]);
        // (rev(stl).push(s[0])).len() == rev(stl).len() + 1 == stl.len() + 1 == s.len()
    }
    // s.len() == 0: rev(s) unfolds once to seq![]
}

// =====================================================================
// 3. Subsequences and order
// =====================================================================

/// Membership witnessed by an index.  Unlike `Seq::contains`, `mem_at`
/// decomposes cleanly when a sequence is taken apart with `drop_first`
/// or `push` (see `mem_drop_first` / `mem_push`): with `contains`, the
/// existential's witness index gets stuck under the `drop_first` term and
/// Z3 will not instantiate the `subrange` axioms inside it.
pub open spec fn mem_at<A>(f: Seq<A>, x: A) -> bool {
    exists|i: int| 0 <= i < f.len() && f[i] == x
}

/// `f` is `s` with zero or more elements deleted (order preserved):
/// the standard greedy recursion on the head of `s`.
pub open spec fn is_subseq<A>(f: Seq<A>, s: Seq<A>) -> bool
    decreases s.len(),
{
    if s.len() == 0 {
        f.len() == 0
    } else if f.len() == 0 {
        true
    } else if s[0] == f[0] {
        is_subseq(f.drop_first(), s.drop_first())
    } else {
        is_subseq(f, s.drop_first())
    }
}

// `sorted_by(f, le)` is `vstd::relations::sorted_by`:
//     forall |i, j| 0 <= i < j < f.len() ==> le(f[i], f[j])
// `transitive(le)` is `vstd::relations::transitive` (not needed below).

/// Bridge between `mem_at` and `Seq::contains`.
pub proof fn mem_at_iff_contains<A>(f: Seq<A>, x: A)
    ensures
        mem_at(f, x) <==> f.contains(x),
{
}

/// `mem_at` over a cons.
pub proof fn mem_cons<A>(hd: A, tl: Seq<A>, x: A)
    ensures
        mem_at(seq![hd] + tl, x) <==> hd == x || mem_at(tl, x),
{
    let cons = seq![hd] + tl;
    assert(cons.len() == 1 + tl.len());
    if hd == x {
        assert(0 <= 0 < cons.len() && cons[0] == hd);
    }
    if mem_at(tl, x) {
        let i = choose|i: int| 0 <= i < tl.len() && tl[i] == x;
        assert(0 <= 1 + i < cons.len());
        assert(cons[1 + i] == tl[i]);
    }
    if mem_at(cons, x) && hd != x {
        let i = choose|i: int| 0 <= i < cons.len() && cons[i] == x;
        assert(i != 0) by {
            if i == 0 {
                assert(cons[0] == hd);
            }
        };
        assert(cons[i] == tl[i - 1]);
        assert(0 <= i - 1 < tl.len());
        assert(tl[i - 1] == x);
    }
}

/// `mem_at` decomposes over `drop_first` — this is the decomposition that
/// `Seq::contains` will not give you.
pub proof fn mem_drop_first<A>(s: Seq<A>, x: A)
    requires
        s.len() > 0,
    ensures
        mem_at(s, x) <==> s[0] == x || mem_at(s.drop_first(), x),
{
    seq_cons(s);
    mem_cons(s[0], s.drop_first(), x);
}

/// `mem_at` decomposes over `push` (companion to `mem_drop_first`).
pub proof fn mem_push<A>(s: Seq<A>, a: A, x: A)
    ensures
        mem_at(s.push(a), x) <==> a == x || mem_at(s, x),
{
    let sp = s.push(a);
    assert(sp.len() == s.len() + 1);
    if a == x {
        assert(0 <= s.len() < sp.len() && sp[s.len() as int] == a);
    }
    if mem_at(s, x) {
        let i = choose|i: int| 0 <= i < s.len() && s[i] == x;
        assert(sp[i] == s[i]);
    }
    if mem_at(sp, x) && a != x {
        let i = choose|i: int| 0 <= i < sp.len() && sp[i] == x;
        assert(i != s.len()) by {
            if i == s.len() {
                assert(sp[s.len() as int] == a);
            }
        };
        assert(0 <= i < s.len());
        assert(s[i] == sp[i]);
    }
}

/// Elements of a subsequence are elements of the supersequence.
pub proof fn subseq_mem<A>(f: Seq<A>, s: Seq<A>, x: A)
    requires
        is_subseq(f, s),
        mem_at(f, x),
    ensures
        mem_at(s, x),
    decreases s.len(),
{
    if s.len() == 0 {
        // is_subseq(f, seq![]) forces f.len() == 0, contradicting mem_at(f, x)
    } else if f.len() == 0 {
        // mem_at(f, x) is vacuous: there is no in-bounds index
    } else {
        let ftl = f.drop_first();
        let stl = s.drop_first();
        assert(stl.len() == s.len() - 1);
        if s[0] == f[0] {
            mem_drop_first(f, x);
            if f[0] == x {
                // s[0] == f[0] == x
                assert(0 <= 0 < s.len() && s[0] == x);
            } else {
                subseq_mem(ftl, stl, x);
                mem_drop_first(s, x);
            }
        } else {
            subseq_mem(f, stl, x);
            mem_drop_first(s, x);
        }
    }
}

/// A sorted sequence stays sorted when its head is dropped, and its head
/// precedes (by `le`) every element of the tail.
pub proof fn sorted_cons<A>(s: Seq<A>, le: spec_fn(A, A) -> bool)
    requires
        s.len() > 0,
        sorted_by(s, le),
    ensures
        sorted_by(s.drop_first(), le),
        forall|k: int|
            #![trigger s.drop_first()[k]]
            0 <= k < s.drop_first().len() ==> le(s[0], s.drop_first()[k]),
{
    let stl = s.drop_first();
    assert(stl.len() == s.len() - 1);
    assert forall|i: int, j: int| #![trigger le(stl[i], stl[j])] 0 <= i < j < stl.len() implies le(
        stl[i],
        stl[j],
    ) by {
        assert(stl[i] == s[i + 1]);
        assert(stl[j] == s[j + 1]);
        assert(le(s[i + 1], s[j + 1]));
    };
    assert forall|k: int| 0 <= k < stl.len() implies le(s[0], stl[k]) by {
        assert(stl[k] == s[k + 1]);
        assert(le(s[0], s[k + 1]));
    };
}

/// A subsequence of an `le`-sorted sequence is `le`-sorted — the core of
/// "a filtered fast path keeps scan order".
///
/// With the pairwise `sorted_by` and the greedy `is_subseq` above,
/// transitivity of `le` is *not* required: in the case `f[0] == s[0]`,
/// every later element of `f` is located in the tail of `s`, whose head
/// relates to it directly by sortedness of `s`.  (If your `sorted_by` is
/// only *adjacent*-sorted, add `requires transitive(le)` and chain.)
pub proof fn subseq_sorted<A>(f: Seq<A>, s: Seq<A>, le: spec_fn(A, A) -> bool)
    requires
        sorted_by(s, le),
        is_subseq(f, s),
    ensures
        sorted_by(f, le),
    decreases s.len(),
{
    if s.len() == 0 {
        // is_subseq(f, seq![]) forces f.len() == 0; sorted_by(f, le) is vacuous
    } else if f.len() == 0 {
        // sorted_by(f, le) is vacuous
    } else {
        let ftl = f.drop_first();
        let stl = s.drop_first();
        assert(stl.len() == s.len() - 1);
        sorted_cons(s, le);
        if s[0] == f[0] {
            subseq_sorted(ftl, stl, le);
            assert forall|i: int, j: int| #![trigger le(f[i], f[j])] 0 <= i < j < f.len() implies le(
                f[i],
                f[j],
            ) by {
                if i > 0 {
                    assert(f[i] == ftl[i - 1]);
                    assert(f[j] == ftl[j - 1]);
                    assert(le(ftl[i - 1], ftl[j - 1]));  // sorted_by(ftl, le)
                } else {
                    // i == 0 < j: locate f[j] in the tail of s
                    assert(f[0] == s[0]);
                    assert(f[j] == ftl[j - 1]);
                    assert(0 <= j - 1 < ftl.len());
                    assert(mem_at(ftl, f[j]));
                    subseq_mem(ftl, stl, f[j]);
                    let k = choose|k: int| 0 <= k < stl.len() && stl[k] == f[j];
                    assert(stl[k] == s[k + 1]);
                    assert(f[j] == s[k + 1]);
                    // head-minimality of s (sorted_cons): le(s[0], stl[k]) == le(f[0], f[j])
                }
            };
        } else {
            // s[0] was deleted: recurse directly
            subseq_sorted(f, stl, le);
        }
    }
}

// =====================================================================
// Smoke tests: how to *use* the lemmas
// =====================================================================

proof fn demo_arith(a: int, b: int, m: int)
    requires
        0 <= a && a <= b,
        0 <= m,
{
    // Z3 will not relate a*m and b*m on its own; the lemma does:
    mul_mono_const(a, b, m);
    assert(a * m <= b * m);
    mul_nonneg(a, m);
    assert(0 <= a * m);
    le_add_const(a, b, 1000);
    assert(a + 1000 <= b + 1000);
}

proof fn demo_rev() {
    let s1 = seq![1int];  // [1]
    let s2 = seq![1int].push(2);  // [1, 2]
    let s3 = seq![1int].push(2).push(3);  // [1, 2, 3]

    rev_push(s2, 3);  // rev(s3) == [3] ++ rev(s2)
    rev_push(s1, 2);  // rev(s2) == [2] ++ rev(s1)
    rev_push(seq![], 1int);  // rev([1]) == [1] ++ rev(seq![])

    assert(rev(Seq::<int>::empty()) == seq![]);  // one unfolding of rev
    assert(seq![1int] + seq![] == seq![1int]);
    assert(rev(s1) == seq![1int]);
    assert(rev(s2) == seq![2int].push(1));  // [2, 1]
    assert(rev(s3) == seq![3int].push(2).push(1));  // [3, 2, 1]

    rev_len(s3);
    assert((rev(s3)).len() == 3);
}

spec fn le_int(x: int, y: int) -> bool {
    x <= y
}

proof fn demo_subseq() {
    let s = seq![1int].push(2).push(3);  // [1, 2, 3]
    let f = seq![1int].push(3);  // [1, 3]  (s with 2 deleted)

    // is_subseq(f, s) needs four unfoldings of is_subseq:
    reveal_with_fuel(is_subseq, 4);
    assert(is_subseq(f, s));

    // s is sorted by <=:
    assert(sorted_by(s, |x: int, y: int| le_int(x, y)));

    // ... therefore the filtered subsequence is sorted, ...
    subseq_sorted(f, s, |x: int, y: int| le_int(x, y));
    assert(sorted_by(f, |x: int, y: int| le_int(x, y)));

    // ... and everything it contains is still in s:
    assert(0 <= 1 < f.len() && f[1] == 3);
    assert(mem_at(f, 3));
    subseq_mem(f, s, 3);
    assert(mem_at(s, 3));
    mem_at_iff_contains(s, 3);
    assert(s.contains(3));

    // (this `le` happens to be transitive, though subseq_sorted doesn't need it)
    assert(transitive(|x: int, y: int| le_int(x, y)));
}

} // verus!

fn main() {}
