//! Aggregate functions over main-region numeric samples.

use crate::grid::{CellAddr, Grid, MainRange};
use crate::ops::{AggFunc, AggregateDef};

fn parse_number(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    t.parse::<f64>().ok()
}

fn collect_numbers(grid: &Grid, range: &MainRange) -> Vec<f64> {
    let mut v = Vec::new();
    if range.is_empty() {
        return v;
    }
    for r in range.row_start..range.row_end {
        for c in range.col_start..range.col_end {
            if let Some(s) = grid.get(&CellAddr::Main { row: r, col: c }) {
                if let Some(n) = parse_number(s) {
                    v.push(n);
                }
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
    let out = match def.func {
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
    };
    out
}

/// Value shown for `addr`: aggregate result if defined, else raw cell.
pub fn cell_display(
    grid: &Grid,
    aggregates: &std::collections::HashMap<CellAddr, AggregateDef>,
    addr: &CellAddr,
) -> String {
    if let Some(def) = aggregates.get(addr) {
        return compute_aggregate(grid, def);
    }
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
}
