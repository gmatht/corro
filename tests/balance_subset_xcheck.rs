//! Cross-check: the proved subset-sum soundness property against the real
//! book-balancing search in `balance_core::solve_balance_groups`.
//!
//! `proofs/balance_subset.rs` proves the step the search takes cannot step
//! past a solution: if nothing in `[start, n)` sums to `target`, then for any
//! usable index `i`, nothing in `(i, n)` sums to `target - vals[i]`.
//!
//! The observable consequence -- the one worth checking -- is that a group is
//! only reported when its members really do sum to the anchor it was matched
//! with, and that an item is never reported both inside a group and as a
//! leftover.
//!
//! This test writes nothing.

use balance_core::{solve_balance_groups, BalanceDirection, BalanceItem};

fn item(id: usize, amount_cents: i64) -> BalanceItem {
    BalanceItem { id, amount_cents }
}

/// Every reported group's members really sum to its reported total.
#[test]
fn groups_sum_to_their_totals() {
    let items = vec![
        item(0, 1000),
        item(1, -400),
        item(2, -600),
        item(3, -300),
        item(4, -250),
        item(5, 550),
    ];
    let sol = solve_balance_groups(&items, BalanceDirection::PosToNeg);
    for g in &sol.groups {
        let sum: i64 = g
            .item_ids
            .iter()
            .map(|id| items.iter().find(|it| it.id == *id).unwrap().amount_cents)
            .sum();
        assert_eq!(sum, g.total_cents, "group {g:?} does not sum to its total");
    }
}

/// A group is only formed when its non-anchor members sum to the anchor.
#[test]
fn each_group_balances_against_its_anchor() {
    let items = vec![
        item(0, 1000),
        item(1, -400),
        item(2, -600),
        item(3, -300),
        item(4, -250),
        item(5, 550),
    ];
    let sol = solve_balance_groups(&items, BalanceDirection::PosToNeg);
    for g in &sol.groups {
        // the anchor is the positive item, the rest are the negatives
        let anchor: i64 = g
            .item_ids
            .iter()
            .map(|id| items.iter().find(|it| it.id == *id).unwrap().amount_cents)
            .filter(|a| *a > 0)
            .sum();
        let rest: i64 = g
            .item_ids
            .iter()
            .map(|id| items.iter().find(|it| it.id == *id).unwrap().amount_cents)
            .filter(|a| *a <= 0)
            .sum();
        assert_eq!(anchor + rest, g.total_cents);
        assert_eq!(anchor + rest, 0, "group {g:?} does not balance to zero");
    }
}

/// Nothing is both inside a group and listed as a leftover.
#[test]
fn grouped_and_leftover_are_disjoint() {
    let items = vec![
        item(0, 1000),
        item(1, -400),
        item(2, -600),
        item(3, -300),
        item(4, -250),
        item(5, 550),
    ];
    let sol = solve_balance_groups(&items, BalanceDirection::PosToNeg);
    let grouped: Vec<usize> = sol.groups.iter().flat_map(|g| g.item_ids.clone()).collect();
    for l in &sol.leftovers {
        assert!(!grouped.contains(l), "item {l} is grouped and a leftover");
    }
    // and every input item appears exactly once, somewhere
    let mut seen = grouped.clone();
    seen.extend(sol.leftovers.iter().copied());
    seen.sort_unstable();
    let all: Vec<usize> = (0..items.len()).collect();
    assert_eq!(seen, all, "every item should be grouped or left over, once");
}

/// The search finds a solution when one exists: 1000 = 400 + 600.
#[test]
fn finds_an_exact_match() {
    let items = vec![item(0, 1000), item(1, -400), item(2, -600)];
    let sol = solve_balance_groups(&items, BalanceDirection::PosToNeg);
    assert_eq!(sol.groups.len(), 1);
    assert_eq!(sol.leftovers.len(), 0);
}

/// With no exact match, nothing is reported and everything is left over --
/// which is exactly the `None` branch the soundness lemma covers.
#[test]
fn reports_nothing_when_no_subset_matches() {
    let items = vec![item(0, 1000), item(1, -400), item(2, -500)];
    let sol = solve_balance_groups(&items, BalanceDirection::PosToNeg);
    assert_eq!(sol.groups.len(), 0, "400 + 500 != 1000, so no group");
    assert_eq!(sol.leftovers, vec![0, 1, 2]);
}
