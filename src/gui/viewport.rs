//! Shared viewport controller for the spreadsheet grid.
//!
//! Every backend repaints the grid the same way: compute the visible
//! row/column indices, fit column widths, derive row labels / column layout /
//! aggregate functions, and refill a [`CellSink`] via [`crate::gui::render`].
//! This module holds that computation so pancurses, GTK and the canvas backends
//! stay in lock-step (render parity) and only differ in the widget calls they
//! make against the resulting [`Viewport`].
//!
//! Nothing here references a specific toolkit — it operates purely on
//! [`crate::gui::App`]/[`crate::core::state::CoreApp`] and the grid.

use crate::addr::{ui_column_fragment, ui_row_label};
use crate::grid::{GridBox, SheetCursor};
use crate::gui::compute;
use crate::gui::render::{self, CellSink};
use crate::gui::App;
use crate::ops::AggFunc;
use crate::ui_core;
use std::collections::HashMap;

/// A fully-computed, paint-ready view of the grid for one frame.
pub struct Viewport {
    pub display_rows: Vec<usize>,
    pub col_ixs: Vec<usize>,
    pub col_widths: HashMap<usize, usize>,
    pub row_labels: Vec<(u32, String)>,
    pub column_layout: Vec<(u32, u32, String)>,
    pub row_agg_func: Vec<Option<AggFunc>>,
    /// Aggregate directive per *displayed* column, positionally aligned with
    /// `col_ixs` (the column mirror of `row_agg_func`). `Some` marks a totals
    /// column, which the painter shades so computed values do not read as
    /// entered data.
    pub col_agg_func: Vec<Option<AggFunc>>,
    pub mr: usize,
    pub mc: usize,
    /// Character width the columns were trimmed to fit. `fill_cells` needs it
    /// to size text spill into the trailing blank space.
    pub data_width: usize,
}

impl Viewport {
    /// Border title string, e.g. `"corro  12r × 3c  ops 5"`. Must match the
    /// format used wherever the title is rendered.
    /// Character width this viewport was laid out for.
    pub fn data_width(&self) -> usize {
        self.data_width
    }

    pub fn border_title(&self, ops: usize) -> String {
        format!("corro  {}r × {}c  ops {}", self.mr, self.mc, ops)
    }

    /// Refill a `CellSink` from this viewport.
    pub fn refill(
        &self,
        sink: &mut dyn CellSink,
        g: &GridBox,
        hr: usize,
        lm: usize,
        data_width: usize,
        cursor_row: usize,
        cursor_col: usize,
    ) {
        render::fill_cells(
            sink,
            &self.display_rows,
            &self.col_ixs,
            &self.col_widths,
            g,
            hr,
            self.mr,
            self.mc,
            lm,
            data_width,
            cursor_row,
            cursor_col,
            &self.row_agg_func,
        );
    }

    /// Full viewport recompute (rows + columns) for `cursor`, mutating `app`
    /// (column-width trimming adjusts the grid). Matches ratatui's draw_visual.
    pub fn recompute(
        app: &mut App,
        cursor: SheetCursor,
        data_rows: usize,
        data_cols: usize,
        data_width: usize,
        hr: usize,
        lm: usize,
    ) -> Viewport {
        let rec = app.core.workbook.active_sheet().clone();
        let display_rows = ui_core::visible_row_indices(&rec, cursor, data_rows, 0).0;
        build(app, &display_rows, cursor, data_cols, data_width, hr, lm)
    }

    /// Recompute only the visible columns (rows unchanged), e.g. after the
    /// cursor moves horizontally or the grid grows. Mutates `app` for width
    /// trimming.
    pub fn recompute_columns(
        app: &mut App,
        display_rows: &[usize],
        cursor: SheetCursor,
        data_cols: usize,
        data_width: usize,
        hr: usize,
        lm: usize,
    ) -> Viewport {
        build(app, display_rows, cursor, data_cols, data_width, hr, lm)
    }

    /// Build a viewport from already-known `display_rows`/`col_ixs` and a
    /// ready column-width map, without recomputing visible indices (the
    /// in-viewport refresh path, where nothing scrolled) and without touching
    /// `app` beyond reading the active sheet's shape.
    ///
    /// `data_width` is the sum of the given widths (what the caller laid the
    /// frame out for); pass the map the caller actually paints with, so the
    /// gutter widths in `column_layout` and the spill budget in `fill_cells`
    /// agree. Both GUI backends build their frame through this one function
    /// (GTK supplies padlock-aware `display_col_width` values, pancurses its
    /// own), so the derived labels/aggregates cannot drift between them.
    pub fn from_parts(
        app: &App,
        display_rows: &[usize],
        col_ixs: &[usize],
        col_widths: &HashMap<usize, usize>,
        hr: usize,
    ) -> Viewport {
        let rec = app.core.workbook.active_sheet().clone();
        let g = &rec.grid;
        let mr = g.main_rows();
        let mc = g.main_cols();
        let row_labels: Vec<(u32, String)> = display_rows
            .iter()
            .enumerate()
            .map(|(idx, &r)| (idx as u32, ui_row_label(r, mr)))
            .collect();
        let column_layout: Vec<(u32, u32, String)> = col_ixs
            .iter()
            .map(|&c| {
                let w = *col_widths.get(&c).unwrap_or(&1);
                (c as u32, w as u32, ui_column_fragment(c, mc))
            })
            .collect();
        let row_agg_func = compute::compute_row_agg_func(g, display_rows, hr, mr);
        let col_agg_func = compute::compute_col_agg_func(g, &col_ixs);
        Viewport {
            display_rows: display_rows.to_vec(),
            col_ixs: col_ixs.to_vec(),
            col_widths: col_widths.clone(),
            row_labels,
            column_layout,
            row_agg_func,
            col_agg_func,
            mr,
            mc,
            data_width: col_widths.values().copied().sum(),
        }
    }

    /// Build a viewport from already-known `display_rows`/`col_ixs`, taking
    /// each column's width from the grid. `data_width` is left at 0 because
    /// the caller paints at its own layout width (see [`Self::refill`]).
    pub fn snapshot(
        app: &App,
        display_rows: &[usize],
        col_ixs: &[usize],
        hr: usize,
        _lm: usize,
    ) -> Viewport {
        let rec = app.core.workbook.active_sheet().clone();
        let g = &rec.grid;
        let col_widths: HashMap<usize, usize> =
            col_ixs.iter().map(|&c| (c, g.col_width(c).max(1))).collect();
        let mut vp = Self::from_parts(app, display_rows, col_ixs, &col_widths, hr);
        // `snapshot` callers size the paint separately; a summed width here
        // would silently re-trim their layout.
        vp.data_width = 0;
        vp
    }
}

/// Shared body of [`Viewport::recompute`] / [`Viewport::recompute_columns`]:
/// derive `col_ixs` (with width trimming), then row labels / column layout /
/// row aggregates from the current grid.
fn build(
    app: &mut App,
    display_rows: &[usize],
    cursor: SheetCursor,
    data_cols: usize,
    data_width: usize,
    hr: usize,
    _lm: usize,
) -> Viewport {
    let (mut col_ixs, _) =
        ui_core::visible_col_indices(&app.core.workbook.active_sheet(), cursor, data_cols, 0);
    // Trim columns to fit (matching ratatui: no proportional refit).
    {
        let sht = app.core.workbook.active_sheet_mut();
        ui_core::trim_visible_cols_to_width(&mut sht.grid, &mut col_ixs, cursor.col, data_width);
    }
    // Re-read the sheet after the width adjustments.
    let rec = app.core.workbook.active_sheet().clone();
    let g = &rec.grid;
    let mr = g.main_rows();
    let mc = g.main_cols();
    let col_widths: HashMap<usize, usize> =
        col_ixs.iter().map(|&c| (c, g.col_width(c).max(1))).collect();
    let row_labels: Vec<(u32, String)> = display_rows
        .iter()
        .enumerate()
        .map(|(idx, &r)| (idx as u32, ui_row_label(r, mr)))
        .collect();
    let column_layout: Vec<(u32, u32, String)> = col_ixs
        .iter()
        .map(|&c| {
            let w = g.col_width(c).max(1);
            (c as u32, w as u32, ui_column_fragment(c, mc))
        })
        .collect();
    let row_agg_func = compute::compute_row_agg_func(g, display_rows, hr, mr);
    let col_agg_func = compute::compute_col_agg_func(g, &col_ixs);
    Viewport {
        display_rows: display_rows.to_vec(),
        col_ixs,
        col_widths,
        row_labels,
        column_layout,
        row_agg_func,
        col_agg_func,
        mr,
        mc,
        data_width,
    }
}
