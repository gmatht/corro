//! Cross-check: the proved order-preservation property against the real
//! `MainRangeEvalPlan` that the sparse aggregate path walks.
//!
//! `proofs/agg_sparse_order.rs` proves that a subsequence of a row-major
//! sorted list is still row-major sorted, so filtering the stored cells to a
//! range keeps the dense loop's visit order. This test checks the two
//! hypotheses of that lemma on a real grid: `stored_sorted` really is sorted,
//! and really does consist only of in-range cells.
//!
//! This test writes nothing.

use corro::grid::{CellAddr, Grid, MainRange};

fn grid_with(cells: &[(u32, u32, &str)]) -> Grid {
    let mut g = Grid::new(10, 5);
    for (r, c, v) in cells {
        g.set(&CellAddr::Main { row: *r, col: *c }, v.to_string());
    }
    g
}

fn row_major(a: (u32, u32), b: (u32, u32)) -> bool {
    a.0 < b.0 || (a.0 == b.0 && a.1 < b.1)
}

#[test]
fn plan_is_row_major_sorted() {
    // stored deliberately out of order
    let g = grid_with(&[
        (7, 3, "1"),
        (1, 1, "1"),
        (5, 0, "1"),
        (1, 4, "1"),
        (9, 2, "1"),
        (0, 0, "1"),
    ]);
    let range = MainRange { row_start: 0, row_end: 10, col_start: 0, col_end: 5 };
    let plan = g.main_range_eval_plan(&range);
    assert!(
        plan.stored_sorted.windows(2).all(|w| row_major(w[0], w[1])),
        "stored_sorted must be row-major: {:?}",
        plan.stored_sorted
    );
}

#[test]
fn plan_contains_only_in_range_cells() {
    let g = grid_with(&[(1, 1, "1"), (20, 20, "9"), (2, 2, "1")]);
    let range = MainRange { row_start: 0, row_end: 10, col_start: 0, col_end: 5 };
    let plan = g.main_range_eval_plan(&range);
    for &(r, c) in &plan.stored_sorted {
        assert!(range.contains(r, c), "({r},{c}) is outside the range");
    }
    assert!(
        !plan.stored_sorted.contains(&(20, 20)),
        "a cell outside the range must not be visited"
    );
}

#[test]
fn plan_is_the_sorted_filter_of_stored_cells() {
    let all = [(7u32, 3u32), (1, 1), (5, 0), (1, 4), (9, 2), (0, 0)];
    let g = grid_with(&all.map(|(r, c)| (r, c, "1")));
    let range = MainRange { row_start: 0, row_end: 10, col_start: 0, col_end: 5 };
    let plan = g.main_range_eval_plan(&range);

    // expected: the stored cells inside the range, sorted row-major
    let mut expected: Vec<(u32, u32)> = all
        .iter()
        .copied()
        .filter(|&(r, c)| range.contains(r, c))
        .collect();
    expected.sort_unstable();
    assert_eq!(plan.stored_sorted, expected);

    // and it is a subsequence of the (sorted) full stored set: every element
    // occurs in the source, which is what `subseq_mem` gives the proof.
    let mut full: Vec<(u32, u32)> = all.to_vec();
    full.sort_unstable();
    for p in &plan.stored_sorted {
        assert!(full.contains(p), "{p:?} is in the plan but not stored");
    }
}

/// A range with no stored cells yields an empty plan, so the sparse path folds
/// nothing -- matching a dense loop over empty cells.
#[test]
fn empty_range_gives_empty_plan() {
    let g = grid_with(&[(1, 1, "1")]);
    let range = MainRange { row_start: 5, row_end: 8, col_start: 3, col_end: 4 };
    let plan = g.main_range_eval_plan(&range);
    assert!(plan.stored_sorted.is_empty());
}
