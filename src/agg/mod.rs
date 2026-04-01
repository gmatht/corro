//! Aggregate functions over main-region numeric samples.

use crate::formula;
use crate::grid::{CellAddr, Grid, MainRange};
use crate::ops::{AggFunc, AggregateDef};

fn collect_numbers(grid: &Grid, range: &MainRange) -> Vec<f64> {
    let mut v = Vec::new();
    if range.is_empty() {
        return v;
    }
    let mut visiting = Vec::new();
    let mut budget = formula::EVAL_BUDGET_AGG;
    for r in range.row_start..range.row_end {
        for c in range.col_start..range.col_end {
            let addr = CellAddr::Main { row: r, col: c };
            if let Some(n) =
                formula::effective_numeric(grid, &addr, &mut visiting, &mut budget)
            {
                v.push(n);
            }
        }
    }
    v
}

fn median(mut xs: Vec<f64>) -> f64 {
    if xs.is_empty() {
        return f64::NAN;
    }
    xs.sort_by(|a, b| a.partial_cmp(b).unwrap());
    let n = xs.len();
    if n % 2 == 1 {
        xs[n / 2]
    } else {
        (xs[n / 2 - 1] + xs[n / 2]) / 2.0
    }
}

/// Compute display string for an aggregate over `source` main cells.
pub fn compute_aggregate(grid: &Grid, def: &AggregateDef) -> String {
    let xs = collect_numbers(grid, &def.source);
    match def.func {
        AggFunc::Sum => {
            if xs.is_empty() {
                String::new()
            } else {
                let s: f64 = xs.iter().sum();
                format!("{s}")
            }
        }
        AggFunc::Mean => {
            if xs.is_empty() {
                String::new()
            } else {
                let s: f64 = xs.iter().sum::<f64>() / xs.len() as f64;
                format!("{s}")
            }
        }
        AggFunc::Median => {
            if xs.is_empty() {
                String::new()
            } else {
                let m = median(xs);
                format!("{m}")
            }
        }
        AggFunc::Min => xs
            .into_iter()
            .min_by(|a, b| a.partial_cmp(b).unwrap())
            .map(|m| format!("{m}"))
            .unwrap_or_default(),
        AggFunc::Max => xs
            .into_iter()
            .max_by(|a, b| a.partial_cmp(b).unwrap())
            .map(|m| format!("{m}"))
            .unwrap_or_default(),
        AggFunc::Count => {
            if xs.is_empty() {
                String::new()
            } else {
                format!("{}", xs.len())
            }
        }
    }
}

/// Raw cell value for display.
pub fn cell_display(grid: &Grid, addr: &CellAddr) -> String {
    grid.get(addr).unwrap_or("").to_string()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::grid::Grid;

    #[test]
    fn sum_mean() {
        let mut g = Grid::new(2, 2);
        g.set(
            &CellAddr::Main { row: 0, col: 0 },
            "2".into(),
        );
        g.set(
            &CellAddr::Main { row: 0, col: 1 },
            "3".into(),
        );
        let def = AggregateDef {
            func: AggFunc::Sum,
            source: MainRange {
                row_start: 0,
                row_end: 2,
                col_start: 0,
                col_end: 2,
            },
        };
        assert_eq!(compute_aggregate(&g, &def), "5");
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
        assert_eq!(compute_aggregate(&g, &def), "5");
    }
}
