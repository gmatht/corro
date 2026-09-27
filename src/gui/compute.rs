use crate::agg::compute_aggregate;
use crate::agg::helpers::{
    data_main_col_count, left_margin_main_col_aggregate,
    left_margin_special_col_aggregate, previous_raw_block,
};
use crate::formula::cell_effective_display;
use crate::grid::{CellAddr, GridBox, MainRange};
use crate::ops::{AggFunc, AggregateDef};

/// Compute row aggregate info for each display row.
pub fn compute_row_agg_func(
    g: &GridBox,
    display_rows: &[usize],
    hr: usize,
    mr: usize,
) -> Vec<Option<AggFunc>> {
    let mut row_agg_func: Vec<Option<AggFunc>> = Vec::with_capacity(display_rows.len());
    for &lr in display_rows {
        let func = if lr < hr {
            None
        } else if lr < hr + mr {
            crate::agg::helpers::left_margin_agg_func(g, (lr - hr) as u32)
        } else {
            crate::agg::helpers::footer_row_agg_func(g, lr - hr - mr)
        };
        row_agg_func.push(func);
    }
    row_agg_func
}

/// Per-display-column aggregate directive, the column-side mirror of
/// [`compute_row_agg_func`].
///
/// Computed once per frame and indexed by the *column's position in
/// `col_ixs`* (like `row_agg_func` is indexed by display-row position), so
/// the paint loop can ask "is this column a totals column?" in O(1) instead
/// of re-scanning the grid's header cells for every cell it draws —
/// `right_col_agg_func` walks `grid.iter_nonempty()`, which is far too
/// expensive to call per painted cell.
pub fn compute_col_agg_func(g: &GridBox, col_ixs: &[usize]) -> Vec<Option<AggFunc>> {
    col_ixs
        .iter()
        .map(|&c| crate::agg::helpers::right_col_agg_func(g, c))
        .collect()
}

pub(crate) use crate::agg::helpers::footer_special_col_aggregate;

/// Right-margin aggregate key for a global column (alias of the shared
/// [`crate::agg::helpers::right_col_agg_func`]); public because the GUI
/// integration tests reach it through `compute::`.
pub fn right_col_agg(grid: &GridBox, global_col: usize) -> Option<AggFunc> {
    crate::agg::helpers::right_col_agg_func(grid, global_col)
}

/// Start of the aggregate block governing a main row (alias of the shared
/// [`crate::agg::helpers::row_total_block_start`]).
pub(crate) use crate::agg::helpers::row_total_block_start;

/// Grow the grid when the cursor sits on the last main row or column and
/// trailing blanks are below the navigation threshold.
///
/// Thin alias of the shared [`crate::ui_core::grow_grid_for_cursor`] so both
/// GUI backends keep calling it through `compute::` (their existing import
/// path) while the rule itself lives in one place.
pub fn grow_grid_for_cursor(grid: &mut GridBox, cursor_row: usize, cursor_col: usize) {
    crate::ui_core::grow_grid_for_cursor(grid, cursor_row, cursor_col)
}

/// Trailing blank main columns/rows (aliases of the shared
/// [`crate::ui_core`] counters, kept for the backends' `compute::` import
/// path).
pub fn trailing_blank_main_cols(grid: &GridBox) -> usize {
    crate::ui_core::trailing_blank_main_cols(grid)
}

/// See [`trailing_blank_main_cols`].
pub fn trailing_blank_main_rows(grid: &GridBox) -> usize {
    crate::ui_core::trailing_blank_main_rows(grid)
}

/// Result of computing a cell's effective display text, style, and metadata.
pub struct CellInfo {
    /// The formatted display text (via format_cell_display), NOT truncated for column width.
    /// The caller (fill_cells) handles truncation/ellipsis/alignment.
    pub formatted: String,
    pub style: CellDisplayStyle,
    pub raw_value: Option<String>,
    pub is_agg_cell: bool,
}

/// Determine the effective display text, style, and metadata for a single cell.
/// Returns the FORMATTED (but not truncated) text, the display style, and the raw value.
/// fill_cells uses this and handles truncation/alignment/spill.
pub fn compute_cell_info(
    g: &GridBox,
    addr: &CellAddr,
    is_cursor_cell: bool,
    row_agg: Option<AggFunc>,
    main_row: Option<u32>,
    footer_row_idx: Option<u32>,
    rca: Option<AggFunc>,
    global_col: usize,
    lm: usize,
    mc: usize,
    mr: usize,
) -> CellInfo {
    let effective = if is_cursor_cell {
        cell_effective_display(g, addr)
    } else if let Some(func) = row_agg {
        if let Some(_) = footer_row_idx {
            if rca.is_some() {
                footer_special_col_aggregate(g, func, global_col, mr, mc)
                    .unwrap_or_else(|| cell_effective_display(g, addr))
            } else if global_col >= lm && global_col < lm + mc {
                let main_col = (global_col - lm) as u32;
                compute_aggregate(
                    g,
                    &AggregateDef {
                        func,
                        source: MainRange {
                            row_start: 0,
                            row_end: mr as u32,
                            col_start: main_col,
                            col_end: main_col + 1,
                        },
                    },
                )
            } else {
                cell_effective_display(g, addr)
            }
        } else if let Some(mri) = main_row {
            if global_col >= lm && global_col < lm + mc {
                if rca.is_some() {
                    let data_cols = data_main_col_count(g);
                    let block_start = row_total_block_start(g, mri);
                    let result = if block_start < mri {
                        left_margin_special_col_aggregate(
                            g, func, global_col, block_start, mri, data_cols,
                        )
                    } else {
                        previous_raw_block(g, mri).and_then(
                            |(start, end)| {
                                left_margin_special_col_aggregate(
                                    g, func, global_col, start, end, data_cols,
                                )
                            },
                        )
                    };
                    result.unwrap_or_else(|| cell_effective_display(g, addr))
                } else {
                    let main_col = (global_col - lm) as u32;
                    left_margin_main_col_aggregate(g, func, mri, main_col)
                }
            } else if rca.is_some() {
                let data_cols = data_main_col_count(g);
                let block_start = row_total_block_start(g, mri);
                let result = if block_start < mri {
                    left_margin_special_col_aggregate(
                        g, func, global_col, block_start, mri, data_cols,
                    )
                } else {
                    previous_raw_block(g, mri).and_then(
                        |(start, end)| {
                            left_margin_special_col_aggregate(
                                g, func, global_col, start, end, data_cols,
                            )
                        },
                    )
                };
                result.unwrap_or_else(|| cell_effective_display(g, addr))
            } else {
                cell_effective_display(g, addr)
            }
        } else {
            cell_effective_display(g, addr)
        }
    } else if let (Some(mri), Some(agg_func)) = (main_row, rca) {
        let data_cols = data_main_col_count(g);
        let agg = compute_aggregate(
            g,
            &AggregateDef {
                func: agg_func,
                source: MainRange {
                    row_start: mri,
                    row_end: mri + 1,
                    col_start: 0,
                    col_end: data_cols as u32,
                },
            },
        );
        // Ratatui parity: the reference shows the aggregate directly, so a
        // numberless row renders blank — never a hardcoded "0" (commit
        // 4746249 added the zero thinking of Excel's SUM-of-empty, but an
        // empty TOTAL cell must stay empty). A margin cell with its own
        // content (manual override/label) still shows it.
        if agg.is_empty() {
            cell_effective_display(g, addr)
        } else {
            agg
        }
    } else {
        cell_effective_display(g, addr)
    };

    let is_agg_cell = if row_agg.is_some() {
        rca.is_some() || (global_col >= lm && global_col < lm + mc)
    } else if let Some(_) = main_row {
        rca.is_some()
    } else {
        false
    };

    let raw_value = g.get(addr);
    let formatted = crate::ui_core::format_cell_display(g, addr, effective);

    let style = if is_cursor_cell {
        CellDisplayStyle::Cursor
    } else if is_agg_cell {
        if footer_row_idx.is_some() && row_agg.is_some() {
            CellDisplayStyle::FooterAggregate
        } else {
            CellDisplayStyle::Aggregate
        }
    } else if crate::ui_core::hyperlink_target(&formatted).is_some() {
        // Hyperlinks render blue and underlined by default (same rule as
        // the ratatui reference). Cursor/aggregate highlights win.
        CellDisplayStyle::Hyperlink
    } else {
        CellDisplayStyle::Default
    };

    CellInfo { formatted, style, raw_value, is_agg_cell }
}

/// Display style for a cell, mapped to backend-specific rendering.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CellDisplayStyle {
    Default,
    Cursor,
    Aggregate,
    FooterAggregate,
    Selected,
    ActiveHeader,
    InactiveHeader,
    /// Hyperlink cell: blue + underlined on every backend.
    Hyperlink,
}

impl CellDisplayStyle {
    /// The inverse of [`GridSink::style_bits`] (and of
    /// [`Self::to_pancurses_style`], which shares its numbering): recover the
    /// style from a stored style bit.
    ///
    /// Needed once the model (rather than a per-backend map) holds the style:
    /// `SpreadsheetModel::cell_styles` stores the `u8`, while painting needs the
    /// enum. An unknown bit is treated as `Default` rather than panicking — a
    /// style written by a future version should degrade, not crash a repaint.
    ///
    /// Deliberately not `#[cfg(feature = "pancurses")]`, unlike its inverse:
    /// this direction is what lets a *canvas* backend read the model.
    pub fn from_style_bits(bits: u8) -> Self {
        match bits {
            1 => CellDisplayStyle::Cursor,
            2 => CellDisplayStyle::Aggregate,
            3 => CellDisplayStyle::FooterAggregate,
            4 => CellDisplayStyle::Selected,
            5 => CellDisplayStyle::ActiveHeader,
            6 => CellDisplayStyle::InactiveHeader,
            7 => CellDisplayStyle::Hyperlink,
            _ => CellDisplayStyle::Default,
        }
    }

    /// Map to pancurses CELL_STYLE_* constants.
    #[cfg(feature = "pancurses")]
    pub fn to_pancurses_style(self) -> u8 {
        match self {
            CellDisplayStyle::Default => 0,
            CellDisplayStyle::Cursor => 1,
            CellDisplayStyle::Aggregate => 2,
            CellDisplayStyle::FooterAggregate => 3,
            CellDisplayStyle::Selected => 4,
            CellDisplayStyle::ActiveHeader => 5,
            CellDisplayStyle::InactiveHeader => 6,
            CellDisplayStyle::Hyperlink => 7,
        }
    }
}

#[cfg(test)]
mod style_bit_roundtrip_tests {
    use super::CellDisplayStyle;
    // `render` (and its `GridSink`) is only compiled with the `gui` feature;
    // the pancurses-only and rswidgets-term sets have no such module.
    #[cfg(feature = "gui")]
    use crate::gui::render::GridSink;

    /// Every style must survive a round-trip through its stored bit.
    ///
    /// This is load-bearing now that the model (not a backend map) holds the
    /// style: a repaint converts bit → enum → bit, so a mismatch would silently
    /// repaint cells with the wrong style rather than failing loudly.
    ///
    /// The bit→enum→bit direction is written with the feature-independent
    /// [`GridSink::style_bits`] where it exists, falling back to
    /// [`CellDisplayStyle::to_pancurses_style`] (the pancurses map, also
    /// ungated) otherwise. Both share the numbering, and neither is gated on
    /// the other's feature, so the test compiles on every feature set.
    #[test]
    fn every_style_round_trips_through_its_bits() {
        for style in [
            CellDisplayStyle::Default,
            CellDisplayStyle::Cursor,
            CellDisplayStyle::Aggregate,
            CellDisplayStyle::FooterAggregate,
            CellDisplayStyle::Selected,
            CellDisplayStyle::ActiveHeader,
            CellDisplayStyle::InactiveHeader,
            CellDisplayStyle::Hyperlink,
        ] {
            #[cfg(feature = "gui")]
            let bits = GridSink::style_bits(style);
            #[cfg(all(not(feature = "gui"), feature = "pancurses"))]
            let bits = style.to_pancurses_style();
            #[cfg(all(not(feature = "gui"), not(feature = "pancurses")))]
            let bits = match style {
                CellDisplayStyle::Default => 0,
                CellDisplayStyle::Cursor => 1,
                CellDisplayStyle::Aggregate => 2,
                CellDisplayStyle::FooterAggregate => 3,
                CellDisplayStyle::Selected => 4,
                CellDisplayStyle::ActiveHeader => 5,
                CellDisplayStyle::InactiveHeader => 6,
                CellDisplayStyle::Hyperlink => 7,
            };
            assert_eq!(
                style,
                CellDisplayStyle::from_style_bits(bits),
                "style {style:?} must round-trip through bits {bits}"
            );
        }
    }

    /// An unknown bit must degrade to `Default`, not panic: a style written by
    /// a newer version would otherwise crash a repaint.
    #[test]
    fn unknown_bits_degrade_to_default() {
        for bits in [8u8, 42, 255] {
            assert_eq!(
                CellDisplayStyle::Default,
                CellDisplayStyle::from_style_bits(bits)
            );
        }
    }
}
