//! Aggregate functions over main-region numeric samples.

use crate::formula::{self, Number};
use crate::grid::{CellAddr, GridBox as Grid, MainRange};
use crate::ops::{AggFunc, AggregateDef};

pub mod helpers;

/// Formatting for margin aggregates when only an [`f64`] is available (`Number::Approx` path).
///
/// Uses [`formula::format_decimal_generic`] (10 significant digits, scientific
/// notation for extreme magnitudes) rather than fixed decimal places: the old
/// `{:.10}` rendered `1e-99` as `0`, so a column like `1 + 10^-99 - 1` summed to
/// a display of `0` even though the value was not zero.
fn format_aggregate_approx(value: f64) -> String {
    formula::format_decimal_generic(value)
}

/// Preserve [`Number::Exact`] without a `float` round-trip; match cell-style rational display.
fn format_aggregate_number(n: &Number) -> String {
    n.format_eval_display(format_aggregate_approx)
}

fn cmp_number_aggregate(a: &Number, b: &Number) -> std::cmp::Ordering {
    a.partial_cmp(b)
        .unwrap_or(std::cmp::Ordering::Equal)
}

/// Median of sample values using the same ordering as formulas (IEEE-aware `Exact` vs `Approx`).
fn median_aggregate(mut xs: Vec<Number>) -> Option<Number> {
    if xs.is_empty() {
        return None;
    }
    xs.sort_by(|a, b| cmp_number_aggregate(a, b));
    let n = xs.len();
    if n % 2 == 1 {
        Some(xs[n / 2].clone())
    } else {
        let a = xs[n / 2 - 1].clone();
        let b = xs[n / 2].clone();
        Some(a.add(b).div(Number::from_i64(2)))
    }
}

fn collect_numbers_summable(grid: &Grid, range: &MainRange) -> Vec<Number> {
    let mut v = Vec::new();
    if range.is_empty() {
        return v;
    }
    // Median is the only caller that needs the whole vector; every other
    // aggregate streams through `fold_numbers_summable` instead
    // (ALGORITHMS.md §2.6).
    fold_numbers_summable(grid, range, &mut v, |v, n| v.push(n));
    v
}

/// Stream summable cells through `f` without materialising a vector.
/// Shares one `visiting` stack and eval budget across the range, exactly as
/// the old collect-then-scan did, so evaluation semantics are unchanged.
fn fold_numbers_summable<T>(
    grid: &Grid,
    range: &MainRange,
    acc: &mut T,
    mut f: impl FnMut(&mut T, Number),
) {
    if range.is_empty() {
        return;
    }
    let mut visiting = Vec::new();
    let mut budget = formula::EVAL_BUDGET_AGG;
    // Sparse fast path (ALGORITHMS.md §2.6): for ranges much larger than the
    // stored content with no applicable templates, evaluate only stored and
    // spilled cells in row-major order. Order matches the dense loop, and
    // skipped empty cells consume no budget and touch no visiting stack.
    if range.area() > 4 * grid.stored_main_count() as u64 {
        let plan = grid.main_range_eval_plan(range);
        if !plan.has_template {
            for (r, c) in plan.stored_sorted {
                let addr = CellAddr::Main { row: r, col: c };
                if let Some(n) = formula::summable_numeric(grid, &addr, &mut visiting, &mut budget)
                {
                    f(acc, n);
                }
            }
            return;
        }
    }
    for r in range.row_start..range.row_end {
        for c in range.col_start..range.col_end {
            let addr = CellAddr::Main { row: r, col: c };
            if let Some(n) = formula::summable_numeric(grid, &addr, &mut visiting, &mut budget) {
                f(acc, n);
            }
        }
    }
}

fn count_numeric_cells(grid: &Grid, range: &MainRange) -> usize {
    let mut n = 0usize;
    if range.is_empty() {
        return n;
    }
    let mut visiting = Vec::new();
    let mut budget = formula::EVAL_BUDGET_AGG;
    // Sparse fast path (ALGORITHMS.md §2.6): same argument as
    // `fold_numbers_summable` above.
    if range.area() > 4 * grid.stored_main_count() as u64 {
        let plan = grid.main_range_eval_plan(range);
        if !plan.has_template {
            for (r, c) in plan.stored_sorted {
                let addr = CellAddr::Main { row: r, col: c };
                if formula::effective_numeric(grid, &addr, &mut visiting, &mut budget).is_some() {
                    n += 1;
                }
            }
            return n;
        }
    }
    for r in range.row_start..range.row_end {
        for c in range.col_start..range.col_end {
            let addr = CellAddr::Main { row: r, col: c };
            if formula::effective_numeric(grid, &addr, &mut visiting, &mut budget).is_some() {
                n += 1;
            }
        }
    }
    n
}

/// Compute display string for an aggregate over `source` main cells.
pub fn compute_aggregate(grid: &Grid, def: &AggregateDef) -> String {
    match def.func {
        AggFunc::Count => {
            let n = count_numeric_cells(grid, &def.source);
            if n == 0 {
                String::new()
            } else {
                format!("{}", n)
            }
        }
        AggFunc::Sum => {
            let mut sum = Number::exact_zero();
            let mut count = 0usize;
            fold_numbers_summable(grid, &def.source, &mut (), |_, n| {
                sum = sum.clone().add(n);
                count += 1;
            });
            if count == 0 {
                String::new()
            } else {
                format_aggregate_number(&sum)
            }
        }
        AggFunc::Mean => {
            let mut sum = Number::exact_zero();
            let mut count = 0usize;
            fold_numbers_summable(grid, &def.source, &mut (), |_, n| {
                sum = sum.clone().add(n);
                count += 1;
            });
            if count == 0 {
                String::new()
            } else {
                let s = sum.div(Number::from_i64(count as i64));
                format_aggregate_number(&s)
            }
        }
        AggFunc::Median => {
            let xs = collect_numbers_summable(grid, &def.source);
            // Complex has no ordering: a non-real sample makes the result
            // undefined (#NUM!) rather than an order-dependent pick.
            if xs.iter().any(Number::is_nonreal) {
                "#NUM!".to_string()
            } else {
                median_aggregate(xs)
                    .map(|m| format_aggregate_number(&m))
                    .unwrap_or_default()
            }
        }
        AggFunc::Min => {
            let mut best: Option<Number> = None;
            // Complex has no ordering: a non-real sample makes the result
            // undefined (#NUM!) rather than an order-dependent pick.
            let mut undefined = false;
            fold_numbers_summable(grid, &def.source, &mut (), |_, n| {
                if n.is_nonreal() {
                    undefined = true;
                    return;
                }
                best = Some(match best.take() {
                    None => n,
                    // `Iterator::min_by` keeps the *last* equally-minimum
                    // element; match that tie rule exactly.
                    Some(b) => {
                        if cmp_number_aggregate(&n, &b) == std::cmp::Ordering::Greater {
                            b
                        } else {
                            n
                        }
                    }
                });
            });
            if undefined {
                "#NUM!".to_string()
            } else {
                best.map(|n| format_aggregate_number(&n)).unwrap_or_default()
            }
        }
        AggFunc::Max => {
            let mut best: Option<Number> = None;
            // Complex has no ordering: a non-real sample makes the result
            // undefined (#NUM!) rather than an order-dependent pick.
            let mut undefined = false;
            fold_numbers_summable(grid, &def.source, &mut (), |_, n| {
                if n.is_nonreal() {
                    undefined = true;
                    return;
                }
                best = Some(match best.take() {
                    None => n,
                    // `Iterator::max_by` keeps the *last* equally-maximum
                    // element; match that tie rule exactly.
                    Some(b) => {
                        if cmp_number_aggregate(&n, &b) == std::cmp::Ordering::Less {
                            b
                        } else {
                            n
                        }
                    }
                });
            });
            if undefined {
                "#NUM!".to_string()
            } else {
                best.map(|n| format_aggregate_number(&n)).unwrap_or_default()
            }
        }
    }
}

/// Raw cell value for display.
pub fn cell_display(grid: &Grid, addr: &CellAddr) -> String {
    // GridBox provides `text` which returns an owned String for the addr.
    grid.text(addr)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::grid::GridBox;
    use crate::grid::{Grid, HEADER_ROWS};

    #[test]
    fn sum_mean() {
        let mut g = Grid::new(2, 2);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "2".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "3".into());
        let def = AggregateDef {
            func: AggFunc::Sum,
            source: MainRange {
                row_start: 0,
                row_end: 2,
                col_start: 0,
                col_end: 2,
            },
        };
        let gb = GridBox::from(g);
        assert_eq!(compute_aggregate(&gb, &def), "5");
    }

    #[test]
    fn aggregate_includes_formula_numeric() {
        let mut g = Grid::new(1, 2);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=1+1".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "3".into());
        let def = AggregateDef {
            func: AggFunc::Sum,
            source: MainRange {
                row_start: 0,
                row_end: 1,
                col_start: 0,
                col_end: 2,
            },
        };
        let gb = GridBox::from(g);
        assert_eq!(compute_aggregate(&gb, &def), "5");
    }

    /// Regression: a SUM whose value is an extreme exact decimal (here
    /// `1 + 10^-99 - 1` = `10^-99`) must display in scientific notation, not
    /// as `0`. The old fixed-decimal aggregate formatter (`{:.10}`) rounded it
    /// to zero, so the column total looked empty/zero.
    #[test]
    fn sum_extreme_exact_decimal_displays_scientific_not_zero() {
        let mut g = Grid::new(3, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "1".into());
        g.set(&CellAddr::Main { row: 1, col: 0 }, "=10^-99".into());
        g.set(&CellAddr::Main { row: 2, col: 0 }, "-1".into());
        let gb = GridBox::from(g);
        let def = AggregateDef {
            func: AggFunc::Sum,
            source: MainRange {
                row_start: 0,
                row_end: 3,
                col_start: 0,
                col_end: 1,
            },
        };
        let shown = compute_aggregate(&gb, &def);
        assert_eq!(shown, "1e-99", "extreme exact sum must not display as 0");
    }

    /// The same value via the approximate (f64) path must also stay
    /// scientific: `Number::Approx(1e-99)` rather than a rounded `0`.
    #[test]
    fn approx_extreme_magnitude_formats_scientific() {
        assert_eq!(format_aggregate_approx(1e-99), "1e-99");
        assert_eq!(format_aggregate_approx(1.5e-99), "1.5e-99");
        assert_eq!(format_aggregate_approx(1e100), "1e100");
        // Human-scale values keep their familiar form.
        assert_eq!(format_aggregate_approx(5.0), "5");
        assert_eq!(format_aggregate_approx(1.5), "1.5");
    }

    #[test]
    fn sparse_range_aggregates_match_dense() {
        // 26x500 range with a handful of cells: area dwarfs stored count,
        // forcing the sparse plan. Results must equal dense evaluation.
        let mut g = Grid::new(500, 26);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "2".into());
        g.set(&CellAddr::Main { row: 499, col: 25 }, "4".into());
        g.set(&CellAddr::Main { row: 250, col: 13 }, "=2*3".into());
        g.set(&CellAddr::Main { row: 100, col: 1 }, "text".into());
        let gb = GridBox::from(g);
        let source = MainRange {
            row_start: 0,
            row_end: 500,
            col_start: 0,
            col_end: 26,
        };
        assert!(!gb.main_range_eval_plan(&source).has_template);
        assert!(gb.main_range_eval_plan(&source).stored_sorted.len() < 10);
        for (func, expect) in [
            (AggFunc::Sum, "12"),
            (AggFunc::Mean, "4"),
            (AggFunc::Min, "2"),
            (AggFunc::Max, "6"),
            (AggFunc::Count, "3"),
            (AggFunc::Median, "4"),
        ] {
            let def = AggregateDef {
                func,
                source: source.clone(),
            };
            assert_eq!(compute_aggregate(&gb, &def), expect, "{func:?}");
        }
    }

    #[test]
    fn ordering_aggregates_over_complex_cells_are_undefined() {
        // Row aggregates share the formula rule: no ordering for non-real
        // complex, so MIN/MAX/MEDIAN are #NUM!, not a pick over the reals.
        // SUM/MEAN over the same cells stay complex-valued.
        let mut g = Grid::new(2, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "3+4i".into());
        g.set(&CellAddr::Main { row: 1, col: 0 }, "5".into());
        let gb = GridBox::from(g);
        let source = MainRange {
            row_start: 0,
            row_end: 2,
            col_start: 0,
            col_end: 1,
        };
        for func in [AggFunc::Min, AggFunc::Max, AggFunc::Median] {
            let def = AggregateDef {
                func,
                source: source.clone(),
            };
            assert_eq!(compute_aggregate(&gb, &def), "#NUM!", "{func:?}");
        }
        let def = AggregateDef {
            func: AggFunc::Sum,
            source: source.clone(),
        };
        assert_eq!(compute_aggregate(&gb, &def), "8+4i");
    }

    #[test]
    fn sparse_range_includes_spill_followers() {
        let mut g = Grid::new(200, 4);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "3".into());
        // A spilled array value with no stored cell must still aggregate.
        g.set_spill_value(CellAddr::Main { row: 150, col: 0 }, "7".into());
        let gb = GridBox::from(g);
        let def = AggregateDef {
            func: AggFunc::Sum,
            source: MainRange {
                row_start: 0,
                row_end: 200,
                col_start: 0,
                col_end: 4,
            },
        };
        assert_eq!(compute_aggregate(&gb, &def), "10");
    }

    #[test]
    fn aggregate_ignores_template_zero_from_blank_references() {
        let mut g = GridBox::from(Grid::new(2, 2));
        g.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: crate::grid::ColumnAddr::Main(1),
            },
            "=A*0.1 -- TAX".into(),
        );

        let def = AggregateDef {
            func: AggFunc::Sum,
            source: MainRange {
                row_start: 0,
                row_end: 2,
                col_start: 1,
                col_end: 2,
            },
        };
        assert_eq!(compute_aggregate(&g, &def), "");

        let def = AggregateDef {
            func: AggFunc::Min,
            source: MainRange {
                row_start: 0,
                row_end: 2,
                col_start: 1,
                col_end: 2,
            },
        };
        assert_eq!(compute_aggregate(&g, &def), "");

        let def = AggregateDef {
            func: AggFunc::Mean,
            source: MainRange {
                row_start: 0,
                row_end: 2,
                col_start: 1,
                col_end: 2,
            },
        };
        assert_eq!(compute_aggregate(&g, &def), "");

        g.set(&CellAddr::Main { row: 1, col: 0 }, "0".into());
        assert_eq!(compute_aggregate(&g, &def), "0");
    }

    #[test]
    fn aggregate_sum_median_exact_decimal_display() {
        let mut g = Grid::new(1, 5);
        for (c, lit) in ["0.1", "0.2", "0.3", "0.4", "0.5"].iter().enumerate() {
            g.set(&CellAddr::Main { row: 0, col: c as u32 }, (*lit).into());
        }
        let range = MainRange {
            row_start: 0,
            row_end: 1,
            col_start: 0,
            col_end: 5,
        };

        let median_def = AggregateDef {
            func: AggFunc::Median,
            source: range,
        };
        let gb = GridBox::from(g);
        assert_eq!(compute_aggregate(&gb, &median_def), "0.3");

        let sum_def = AggregateDef {
            func: AggFunc::Sum,
            source: MainRange {
                row_start: 0,
                row_end: 1,
                col_start: 0,
                col_end: 5,
            },
        };
        assert_eq!(compute_aggregate(&gb, &sum_def), "1.5");
    }
}
