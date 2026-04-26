//! Aggregate functions over main-region numeric samples.

use crate::formula;
use crate::grid::{CellAddr, GridBox as Grid, MainRange};
use crate::ops::{AggFunc, AggregateDef};

fn format_aggregate_value(value: f64) -> String {
    if !value.is_finite() {
        return value.to_string();
    }
    let s = format!("{value:.10}");
    if s.contains('.') {
        s.trim_end_matches('0').trim_end_matches('.').to_string()
    } else {
        s
    }
}

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
            if let Some(n) = formula::effective_numeric(grid, &addr, &mut visiting, &mut budget) {
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
                format_aggregate_value(s)
            }
        }
        AggFunc::Mean => {
            if xs.is_empty() {
                String::new()
            } else {
                let s: f64 = xs.iter().sum::<f64>() / xs.len() as f64;
                format_aggregate_value(s)
            }
        }
        AggFunc::Median => {
            if xs.is_empty() {
                String::new()
            } else {
                let m = median(xs);
                format_aggregate_value(m)
            }
        }
        AggFunc::Min => xs
            .into_iter()
            .min_by(|a, b| a.partial_cmp(b).unwrap())
            .map(format_aggregate_value)
            .unwrap_or_default(),
        AggFunc::Max => xs
            .into_iter()
            .max_by(|a, b| a.partial_cmp(b).unwrap())
            .map(format_aggregate_value)
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
    // GridBox provides `text` which returns an owned String for the addr.
    grid.text(addr)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::grid::GridBox;
    use crate::grid::{Grid, HEADER_ROWS, MARGIN_COLS};

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

    #[test]
    fn aggregate_ignores_template_zero_from_blank_references() {
        let mut g = GridBox::from(Grid::new(2, 2));
        g.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: MARGIN_COLS as u32 + 1,
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
}
