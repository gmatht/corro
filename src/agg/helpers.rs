use crate::formula::cell_effective_display;
use crate::formula::Number;
use crate::grid::{CellAddr, GridBox as Grid, MainRange, MARGIN_COLS};
use crate::ops::{AggFunc, AggregateDef};
use num_complex::Complex64;

// Re-exported for the always-compiled default UI (ui/mod.rs, ui_core.rs) which
// cannot reach the gui-gated gui::compute version.
pub(crate) use crate::ods::footer_row_agg_func;

/// Compute a footer aggregate value across all main rows for a given column.
/// Mirrors gui::compute::footer_special_col_aggregate but lives in this always
/// compiled module so the default ratatui UI can use it.
pub(crate) fn footer_special_col_aggregate(
    grid: &Grid,
    footer_func: AggFunc,
    global_col: usize,
    main_rows: usize,
    main_cols: usize,
) -> Option<String> {
    let row_func = right_col_agg_func(grid, global_col);
    let data_cols = data_main_col_count(grid);
    let mut samples: Vec<f64> = Vec::new();
    let mut complex: Vec<Complex64> = Vec::new();
    for r in 0..main_rows {
        let row_val = if let Some(func) = row_func {
            crate::agg::compute_aggregate(
                grid,
                &AggregateDef {
                    func,
                    source: MainRange {
                        row_start: r as u32,
                        row_end: r as u32 + 1,
                        col_start: 0,
                        col_end: data_cols as u32,
                    },
                },
            )
        } else if global_col < MARGIN_COLS {
            String::new()
        } else if global_col < MARGIN_COLS + main_cols {
            cell_effective_display(
                grid,
                &CellAddr::Main {
                    row: r as u32,
                    col: (global_col - MARGIN_COLS) as u32,
                },
            )
        } else {
            cell_effective_display(
                grid,
                &CellAddr::Right {
                    col: (global_col - MARGIN_COLS - main_cols),
                    row: r as u32,
                },
            )
        };
        if let Some(n) = parse_num(&row_val) {
            samples.push(n);
        } else if let Some(c) = parse_complex_display(&row_val) {
            // A complex row total is a real computed value, not text: keep
            // it for the algebraic aggregates instead of dropping it.
            complex.push(c);
        }
    }
    Some(fold_numbers_with_complex(footer_func, &samples, &complex))
}

// Internal helpers kept private to this module
pub(crate) fn right_col_agg_func(grid: &Grid, global_col: usize) -> Option<AggFunc> {
    let main_cols = grid.main_cols();
    let mut labels: Vec<(u32, String)> = grid
        .iter_nonempty()
        .filter_map(|(addr, val)| match addr {
            CellAddr::Header { row, col } if col.to_global(main_cols) == global_col => Some((row, val)),
            _ => None,
        })
        .collect();
    labels.sort_unstable_by_key(|(row, _)| *row);
    for (_, val) in labels {
        if let Some(f) = crate::ops::margin_key_agg_func(&val) {
            return Some(f);
        }
    }
    None
}

pub(crate) fn left_margin_agg_func(grid: &Grid, main_row: u32) -> Option<AggFunc> {
    let key_col = MARGIN_COLS - 1;
    let val = grid.get(&CellAddr::Left { col: key_col, row: main_row })?;
    crate::ops::margin_key_agg_func(&val)
}

pub(crate) fn row_total_block_start(grid: &Grid, current_main_row: u32) -> u32 {
    for candidate in (0..current_main_row).rev() {
        if left_margin_agg_func(grid, candidate).is_some() {
            return candidate + 1;
        }
    }
    0
}

pub(crate) fn parse_num(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    t.parse::<f64>().ok()
}

/// Parse a complex row-total display back into a value.
///
/// Mirrors `format_complex` exactly: `{re}±{im}i` with an explicit sign and
/// `i` suffix (e.g. `"5+1i"`, `"2-3i"`, `"0+1i"`). Anything else is `None`,
/// so unparseable text keeps today's skip behavior — this only *adds* the
/// complex forms the evaluator itself emits, it never reinterprets text.
fn parse_complex_display(s: &str) -> Option<Complex64> {
    let body = s.trim().strip_suffix('i')?;
    // Split at the last interior sign so a negative real part survives.
    let sep = body[1..].find(['+', '-']).map(|i| i + 1)?;
    let (re_s, im_s) = body.split_at(sep);
    let re: f64 = re_s.parse().ok()?;
    let (neg, digits) = match im_s.strip_prefix('+') {
        Some(d) => (false, d),
        None => (true, im_s.strip_prefix('-')?),
    };
    let mut im: f64 = digits.parse().ok()?;
    if neg {
        im = -im;
    }
    if !re.is_finite() || !im.is_finite() {
        return None;
    }
    Some(Complex64::new(re, im))
}

/// Aggregate samples for the algebraic functions (SUM, MEAN) when some row
/// totals are complex.
///
/// Without complex samples this is exactly `fold_numbers` (bit-identical
/// output — the all-real path below delegates untouched). With complex
/// samples, reals join the complex sum and the result displays in the same
/// shape column totals already use for complex values. A zero imaginary part
/// collapses to the plain real rendering (`"5"`, not `"5+0i"`).
///
/// Ordering-based aggregates (MIN/MAX/MEDIAN) have no meaning for non-real
/// complex samples, so those surface `#NUM!` instead of silently ignoring
/// them. COUNT stays on real samples only, and zero-imaginary complex joins
/// the reals by its real part (the same rule formulas use).
fn fold_numbers_with_complex(func: AggFunc, reals: &[f64], complex: &[Complex64]) -> String {
    if complex.is_empty() {
        return fold_numbers(func, reals);
    }
    match func {
        AggFunc::Sum | AggFunc::Mean => {
            let mut acc = Complex64::new(reals.iter().sum(), 0.0);
            for c in complex {
                acc += *c;
            }
            let v = if matches!(func, AggFunc::Mean) {
                acc / ((reals.len() + complex.len()) as f64)
            } else {
                acc
            };
            if v.im == 0.0 {
                format!("{}", v.re)
            } else {
                super::format_aggregate_number(&Number::Complex(v))
            }
        }
        AggFunc::Min | AggFunc::Max | AggFunc::Median => {
            // No ordering for non-real complex: undefined (#NUM!), not a
            // pick over the reals. Zero-imaginary complex joins the reals
            // by its real part (the formula rule).
            let mut xs: Vec<f64> = reals.to_vec();
            for c in complex {
                if c.im != 0.0 {
                    return "#NUM!".to_string();
                }
                xs.push(c.re);
            }
            fold_numbers(func, &xs)
        }
        AggFunc::Count => fold_numbers(func, reals),
    }
}

pub(crate) fn fold_numbers(func: AggFunc, xs: &[f64]) -> String {
    if xs.is_empty() {
        return String::new();
    }
    match func {
        AggFunc::Sum => format!("{}", xs.iter().sum::<f64>()),
        AggFunc::Mean => format!("{}", xs.iter().sum::<f64>() / xs.len() as f64),
        AggFunc::Median => {
            let mut ys = xs.to_vec();
            ys.sort_by(|a, b| a.partial_cmp(b).unwrap());
            let n = ys.len();
            let m = if n % 2 == 1 { ys[n / 2] } else { (ys[n / 2 - 1] + ys[n / 2]) / 2.0 };
            format!("{m}")
        }
        AggFunc::Min => xs
            .iter()
            .copied()
            .min_by(|a, b| a.partial_cmp(b).unwrap())
            .map(|v| format!("{v}"))
            .unwrap_or_default(),
        AggFunc::Max => xs
            .iter()
            .copied()
            .max_by(|a, b| a.partial_cmp(b).unwrap())
            .map(|v| format!("{v}"))
            .unwrap_or_default(),
        AggFunc::Count => format!("{}", xs.len()),
    }
}

// Shared helper functions used by UI and ODS
pub(crate) fn data_main_col_count(grid: &Grid) -> usize {
    let mc = grid.main_cols();
    for c in 0..mc {
        if right_col_agg_func(grid, MARGIN_COLS + c).is_some() {
            return c + 1;
        }
    }
    mc
}

pub(crate) fn previous_raw_block(grid: &Grid, current_main_row: u32) -> Option<(u32, u32)> {
    let mut end = current_main_row;
    while end > 0 {
        let last_agg = (0..end)
            .rev()
            .find(|&r| left_margin_agg_func(grid, r).is_some())
            .unwrap_or(0);
        let prev_agg = if last_agg == 0 {
            None
        } else {
            (0..last_agg)
                .rev()
                .find(|&r| left_margin_agg_func(grid, r).is_some())
        };
        let start = prev_agg.map_or(0, |r| r + 1);
        if start < last_agg {
            return Some((start, last_agg));
        }
        if last_agg == 0 {
            return Some((0, end));
        }
        end = last_agg;
    }
    Some((0, current_main_row))
}

pub(crate) fn left_margin_main_col_aggregate(
    grid: &Grid,
    subtotal_func: AggFunc,
    main_row: u32,
    main_col: u32,
) -> String {
    let block_start = row_total_block_start(grid, main_row);
    let Some((start, end)) = (if block_start < main_row {
        Some((block_start, main_row))
    } else {
        previous_raw_block(grid, main_row)
    }) else {
        return String::new();
    };
    crate::agg::compute_aggregate(
        grid,
        &AggregateDef {
            func: subtotal_func,
            source: MainRange {
                row_start: start,
                row_end: end,
                col_start: main_col,
                col_end: main_col + 1,
            },
        },
    )
}

pub(crate) fn left_margin_special_col_aggregate(
    grid: &Grid,
    subtotal_func: AggFunc,
    global_col: usize,
    row_start: u32,
    row_end: u32,
    data_cols: usize,
) -> Option<String> {
    let row_func = right_col_agg_func(grid, global_col)?;
    let collect = |row_start: u32, row_end: u32| -> (Vec<f64>, Vec<Complex64>) {
        let mut samples: Vec<f64> = Vec::new();
        let mut complex: Vec<Complex64> = Vec::new();
        for r in row_start..row_end {
            let row_val = crate::agg::compute_aggregate(
                grid,
                &AggregateDef {
                    func: row_func,
                    source: MainRange {
                        row_start: r,
                        row_end: r + 1,
                        col_start: 0,
                        col_end: data_cols as u32,
                    },
                },
            );
            if let Some(n) = parse_num(&row_val) {
                samples.push(n);
            } else if let Some(c) = parse_complex_display(&row_val) {
                complex.push(c);
            }
        }
        (samples, complex)
    };

    let (mut samples, mut complex) = collect(row_start, row_end);
    let mut end = row_start;
    while samples.is_empty() && complex.is_empty() && end > 0 {
        let Some((fallback_start, fallback_end)) = previous_raw_block(grid, end) else {
            break;
        };
        (samples, complex) = collect(fallback_start, fallback_end);
        if fallback_start == 0 {
            break;
        }
        end = fallback_start;
    }
    Some(fold_numbers_with_complex(subtotal_func, &samples, &complex))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn complex_display_parses_back() {
        // Exactly the shapes `format_complex` emits; anything else stays None
        // (plain text keeps today's skip behavior).
        assert_eq!(
            parse_complex_display("5+1i"),
            Some(Complex64::new(5.0, 1.0))
        );
        assert_eq!(
            parse_complex_display("0+1i"),
            Some(Complex64::new(0.0, 1.0))
        );
        assert_eq!(
            parse_complex_display("2-3i"),
            Some(Complex64::new(2.0, -3.0))
        );
        assert_eq!(
            parse_complex_display("-2-3i"),
            Some(Complex64::new(-2.0, -3.0))
        );
        assert_eq!(
            parse_complex_display("  5+1i  "),
            Some(Complex64::new(5.0, 1.0))
        );
        assert_eq!(parse_complex_display("hello"), None);
        assert_eq!(parse_complex_display("3"), None);
        assert_eq!(parse_complex_display("1/3"), None);
        assert_eq!(parse_complex_display(""), None);
        assert_eq!(parse_complex_display("5+i"), None);
        assert_eq!(parse_complex_display("inf+1i"), None);
    }

    #[test]
    fn complex_fold_matches_real_fold_without_complex() {
        // No complex samples: bit-identical to the legacy f64 fold.
        for func in [
            AggFunc::Sum,
            AggFunc::Mean,
            AggFunc::Median,
            AggFunc::Min,
            AggFunc::Max,
            AggFunc::Count,
        ] {
            let reals = vec![1.0, 2.0, 3.0];
            assert_eq!(
                fold_numbers_with_complex(func, &reals, &[]),
                fold_numbers(func, &reals),
                "{func:?} must be unchanged without complex samples"
            );
        }
        assert_eq!(
            fold_numbers_with_complex(AggFunc::Sum, &[], &[]),
            "",
            "empty stays blank"
        );
    }

    #[test]
    fn complex_sum_and_mean() {
        assert_eq!(
            fold_numbers_with_complex(
                AggFunc::Sum,
                &[3.0],
                &[Complex64::new(5.0, 1.0)]
            ),
            "8+1i"
        );
        assert_eq!(
            fold_numbers_with_complex(AggFunc::Sum, &[], &[Complex64::new(0.0, 1.0)]),
            "0+1i"
        );
        assert_eq!(
            fold_numbers_with_complex(
                AggFunc::Mean,
                &[2.0],
                &[Complex64::new(4.0, 2.0)]
            ),
            "3+1i"
        );
        // Zero imaginary part collapses to the plain real rendering.
        assert_eq!(
            fold_numbers_with_complex(AggFunc::Sum, &[], &[Complex64::new(5.0, 0.0)]),
            "5"
        );
    }

    #[test]
    fn ordering_aggregates_over_complex_are_undefined() {
        // Complex has no ordering: MIN/MAX/MEDIAN over a non-real sample is
        // #NUM!, not a pick over the reals. COUNT still counts real samples
        // only, and zero-imaginary complex joins the reals by its real part.
        let complex = vec![Complex64::new(5.0, 1.0)];
        assert_eq!(
            fold_numbers_with_complex(AggFunc::Min, &[3.0], &complex),
            "#NUM!"
        );
        assert_eq!(
            fold_numbers_with_complex(AggFunc::Max, &[3.0], &complex),
            "#NUM!"
        );
        assert_eq!(
            fold_numbers_with_complex(AggFunc::Median, &[3.0], &complex),
            "#NUM!"
        );
        assert_eq!(
            fold_numbers_with_complex(AggFunc::Count, &[3.0], &complex),
            "1"
        );
        // Zero imaginary part participates by its real part.
        let real_complex = vec![Complex64::new(7.0, 0.0)];
        assert_eq!(
            fold_numbers_with_complex(AggFunc::Max, &[3.0], &real_complex),
            "7"
        );
        assert_eq!(
            fold_numbers_with_complex(AggFunc::Min, &[3.0], &real_complex),
            "3"
        );
    }
}
