//! Extrapolation utilities: detect simple sequences and generate preview/commit
//! operations. This module intentionally provides a small, well-documented API so
//! the UI can call into it for drag-preview and commit.

use crate::grid::{CellAddr, GridBox, MainRange};
use crate::formula::{translate_formula_text_by_offset, is_formula};

/// Direction for a 1-D extrapolation (used by the UI when inferring values).
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum FillDirection {
    Right,
    Down,
}

/// Infer a single fill value from a seed sequence. Matches the UI precedence:
/// 1) formula translation (translate_formula_text_by_offset)
/// 2) numeric linear extrapolation
/// 3) named-sequence (weekdays/months)
/// 4) suffix increment (preserve zero-padding width)
/// 5) fallback to the last seed value
pub fn infer_fill_value(
    seed: &[String],
    offset_from_last: i32,
    direction: FillDirection,
) -> Option<String> {
    let last = seed.last()?.clone();
    if is_formula(&last) {
        let (row_delta, col_delta) = match direction {
            FillDirection::Right => (0, offset_from_last),
            FillDirection::Down => (offset_from_last, 0),
        };
        if let Some(translated) = translate_formula_text_by_offset(&last, row_delta, col_delta) {
            return Some(translated);
        }
    }
    if let Some(v) = infer_numeric_fill(seed, offset_from_last) {
        return Some(v);
    }
    if let Some(v) = infer_named_sequence_fill(seed, offset_from_last) {
        return Some(v);
    }
    if let Some(v) = infer_suffix_fill(seed, offset_from_last) {
        return Some(v);
    }
    Some(last)
}

/// Basic preview cell: target address and the value to show in preview.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PreviewCell {
    pub addr: CellAddr,
    pub value: String,
}

/// Analyze a rectangular main-region selection and generate a preview fill for
/// the target range. This minimal implementation implements two simple rules:
/// - If the source contains a single formula cell, translate it by offsets (relative fill).
/// - Otherwise, repeat the last cell's text to the target region.
///
/// Parameters:
/// - grid: current sheet grid (GridBox) for value access.
/// - source: main-range being dragged from (row/col in main-space).
/// - target: main-range being filled to (row/col in main-space).
///
/// Returns a Vec of PreviewCell for the target cells in arbitrary order.
pub fn generate_preview(
    grid: &GridBox,
    source: &MainRange,
    target: &MainRange,
) -> Vec<PreviewCell> {
    // Collect source cells (row-major) into a vector of Option<String> so we
    // preserve missing vs present semantics from Grid::get.
    let mut src_values: Vec<Option<String>> = Vec::new();
    for r in source.row_start..source.row_end {
        for c in source.col_start..source.col_end {
            let addr = CellAddr::Main { row: r, col: c };
            let v = grid.get(&addr);
            src_values.push(v);
        }
    }

    // Count formula cells (only present values that start with '=') and
    // count non-empty present values (Some(s) where s is not empty).
    let formula_count = src_values
        .iter()
        .filter(|opt| opt.as_ref().map_or(false, |s| s.trim_start().starts_with('=')))
        .count();
    let nonempty_count = src_values.iter().filter(|opt| opt.as_ref().map_or(false, |s| !s.is_empty())).count();

    let mut out: Vec<PreviewCell> = Vec::new();
    if formula_count == 1 && nonempty_count == 1 {
        // Find the formula cell index and its source coords.
        if let Some((fi, formula_text)) = src_values
            .iter()
            .enumerate()
            .find_map(|(idx, opt)| {
                opt.as_ref()
                    .and_then(|s| s.trim_start().starts_with('=').then_some((idx, s.clone())))
            })
        {
            let src_cols = (source.col_end - source.col_start) as usize;
            let src_r = fi / src_cols;
            let src_c = fi % src_cols;

            for r in target.row_start..target.row_end {
                for c in target.col_start..target.col_end {
                    // compute row/col delta in main-space relative to source top-left
                    let row_delta = r as i32 - (source.row_start as i32 + src_r as i32);
                    let col_delta = c as i32 - (source.col_start as i32 + src_c as i32);
                    let translated = translate_formula_text_by_offset(&formula_text, row_delta, col_delta)
                        .unwrap_or_else(|| formula_text.clone());
                    out.push(PreviewCell {
                        addr: CellAddr::Main { row: r, col: c },
                        value: translated,
                    });
                }
            }
            return out;
        }
    }

    // Fallback: repeat last non-empty present source value across target.
    // If there is no non-empty present source value, return an empty preview
    // (do not fill with empty strings).
    let last_opt: Option<String> = src_values
        .into_iter()
        .rev()
        .find_map(|opt| opt.and_then(|s| (!s.is_empty()).then_some(s)));
    let last = match last_opt {
        Some(v) => v,
        None => return out, // empty
    };
    for r in target.row_start..target.row_end {
        for c in target.col_start..target.col_end {
            out.push(PreviewCell {
                addr: CellAddr::Main { row: r, col: c },
                value: last.clone(),
            });
        }
    }
    out
}

// The following helpers mirror the original UI inference routines. They are
// kept private to this module but exposed via `infer_fill_value` above so the
// UI can call a single centralized function.
fn infer_numeric_fill(seed: &[String], offset_from_last: i32) -> Option<String> {
    if !seed.iter().all(|v| v.trim().parse::<f64>().is_ok()) {
        return None;
    }
    let last = seed.last()?.trim().parse::<f64>().ok()?;
    let prev = if seed.len() >= 2 {
        seed[seed.len() - 2].trim().parse::<f64>().ok()?
    } else {
        last
    };
    let step = last - prev;
    Some(format!("{}", last + step * offset_from_last as f64))
}

fn infer_named_sequence_fill(seed: &[String], offset_from_last: i32) -> Option<String> {
    const WEEKDAYS: [&str; 7] = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"];
    const MONTHS: [&str; 12] = [
        "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC",
    ];
    let normalized: Vec<String> = seed.iter().map(|v| v.trim().to_ascii_uppercase()).collect();
    let last = normalized.last()?.as_str();
    if normalized.iter().all(|v| WEEKDAYS.contains(&v.as_str())) {
        let idx = WEEKDAYS.iter().position(|&v| v == last)?;
        return Some(
            WEEKDAYS
                [(idx as i32 + offset_from_last).rem_euclid(WEEKDAYS.len() as i32) as usize]
                .to_string(),
        );
    }
    if normalized.iter().all(|v| MONTHS.contains(&v.as_str())) {
        let idx = MONTHS.iter().position(|&v| v == last)?;
        return Some(
            MONTHS[(idx as i32 + offset_from_last).rem_euclid(MONTHS.len() as i32) as usize]
                .to_string(),
        );
    }
    None
}

fn infer_suffix_fill(seed: &[String], offset_from_last: i32) -> Option<String> {
    let last = seed.last()?.trim();
    let (prefix, digits) = split_trailing_digits(last)?;
    if seed
        .iter()
        .any(|v| split_trailing_digits(v.trim()).is_none_or(|(p, _)| p != prefix))
    {
        return None;
    }
    let width = digits.len();
    let last_num = digits.parse::<i64>().ok()?;
    let prev_num = if seed.len() >= 2 {
        let (_, prev_digits) = split_trailing_digits(seed[seed.len() - 2].trim())?;
        prev_digits.parse::<i64>().ok()?
    } else {
        last_num
    };
    let next = last_num + (last_num - prev_num) * offset_from_last as i64;
    Some(format!("{}{:0width$}", prefix, next, width = width))
}

fn split_trailing_digits(s: &str) -> Option<(&str, &str)> {
    let bytes = s.as_bytes();
    let mut i = bytes.len();
    while i > 0 && bytes[i - 1].is_ascii_digit() {
        i -= 1;
    }
    if i == bytes.len() {
        return None;
    }
    Some((&s[..i], &s[i..]))
}

/// Construct an Op::FillRange or Op::RelFillRange equivalent commit for the
/// given preview cells. The caller (UI) should wrap this into a WorkbookOp::SheetOp
/// and commit via the existing IO/ops flow.
pub fn commit_from_preview(cells: Vec<PreviewCell>) -> crate::ops::Op {
    let mapped: Vec<(CellAddr, String)> = cells.into_iter().map(|p| (p.addr, p.value)).collect();
    crate::ops::Op::FillRange { cells: mapped }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::grid::{Grid, GridBox, CellAddr, MainRange};

    #[test]
    fn generate_preview_translates_single_formula() {
        let mut gb = GridBox::from(Grid::new(4, 1));
        gb.set(&CellAddr::Main { row: 0, col: 0 }, "=A1".into());

        let source = MainRange { row_start: 0, row_end: 1, col_start: 0, col_end: 1 };
        let target = MainRange { row_start: 1, row_end: 3, col_start: 0, col_end: 1 };

        let out = generate_preview(&gb, &source, &target);
        assert_eq!(out.len(), 2);
        assert!(out.contains(&PreviewCell { addr: CellAddr::Main { row: 1, col: 0 }, value: "=A2".into() }));
        assert!(out.contains(&PreviewCell { addr: CellAddr::Main { row: 2, col: 0 }, value: "=A3".into() }));
    }

    #[test]
    fn generate_preview_repeats_last_nonempty() {
        let mut gb = GridBox::from(Grid::new(4, 2));
        gb.set(&CellAddr::Main { row: 0, col: 0 }, "one".into());

        let source = MainRange { row_start: 0, row_end: 1, col_start: 0, col_end: 2 };
        let target = MainRange { row_start: 1, row_end: 3, col_start: 0, col_end: 2 };

        let out = generate_preview(&gb, &source, &target);
        // 2 rows * 2 cols
        assert_eq!(out.len(), 4);
        for pc in out {
            assert_eq!(pc.value, "one");
        }
    }
}
