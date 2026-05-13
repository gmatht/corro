//! Extrapolation utilities: detect simple sequences and generate preview/commit
//! operations. This module intentionally provides a small, well-documented API so
//! the UI can call into it for drag-preview and commit.

use crate::grid::{CellAddr, GridBox, MainRange};
use crate::formula::translate_formula_text_by_offset;

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
    // Collect source cells (row-major) into a vector of strings.
    let mut src_values: Vec<String> = Vec::new();
    for r in source.row_start..source.row_end {
        for c in source.col_start..source.col_end {
            let addr = CellAddr::Main { row: r, col: c };
            let v = grid.get(&addr).unwrap_or_else(|| "".to_string());
            src_values.push(v);
        }
    }

    let single_formula = src_values
        .iter()
        .filter(|s| s.trim_start().starts_with('='))
        .count()
        == 1
        && src_values.iter().filter(|s| !s.is_empty()).count() == 1;

    let mut out: Vec<PreviewCell> = Vec::new();
    if single_formula {
        // Find the formula cell index and its source coords.
        let mut formula_idx = None;
        for (idx, v) in src_values.iter().enumerate() {
            if v.trim_start().starts_with('=') {
                formula_idx = Some(idx);
                break;
            }
        }
        if let Some(fi) = formula_idx {
            let src_cols = (source.col_end - source.col_start) as usize;
            let src_r = fi / src_cols;
            let src_c = fi % src_cols;
            let formula_text = src_values[fi].clone();

            for r in target.row_start..target.row_end {
                for c in target.col_start..target.col_end {
                    // compute row/col delta in main-space relative to source top-left
                    let row_delta = (r as i32 - (source.row_start as i32 + src_r as i32));
                    let col_delta = (c as i32 - (source.col_start as i32 + src_c as i32));
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

    // Fallback: repeat last non-empty source value across target
    let last = src_values
        .into_iter()
        .rev()
        .find(|s| !s.is_empty())
        .unwrap_or_else(|| "".to_string());
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

/// Construct an Op::FillRange or Op::RelFillRange equivalent commit for the
/// given preview cells. The caller (UI) should wrap this into a WorkbookOp::SheetOp
/// and commit via the existing IO/ops flow.
pub fn commit_from_preview(cells: Vec<PreviewCell>) -> crate::ops::Op {
    let mapped: Vec<(CellAddr, String)> = cells.into_iter().map(|p| (p.addr, p.value)).collect();
    crate::ops::Op::FillRange { cells: mapped }
}
