//! Interactive extrapolate mode for the GUI backends (pancurses/GTK/nwg),
//! mirroring the ratatui reference's modal flow:
//!   Enter -> compute fills and FILL-commit them
//!   Esc   -> cancel (restore the pre-preview grid)
//!   arrows/screen-key -> extend the selection (anchor anchored) and refresh the
//!                        in-grid preview
//!
//! The preview is written straight into the active sheet's grid so the grid
//! widget renders predicted values dimmed-without-persisting; `cancel` restores
//! the saved snapshot and `commit` persists exactly what is shown. The seed
//! range (the selection at the moment Extrapolate was invoked) is kept so only
//! those cells feed inference, matching ratatui's `Mode::Extrapolate`.

use crate::grid::{CellAddr, GridBox, SheetCursor};
use crate::ops::{Op, WorkbookOp};

/// Modal extrapolation state (stored on the GUI `App`, not CoreApp, so the
/// toolkit/app layering stays intact — Core lacks a notion of modal UI).
pub struct ExtrapolateModal {
    /// The selection at the moment Extrapolate was invoked: these logica
    /// rows/cols are the inference source.
    pub seed_anchor: SheetCursor,
    pub seed_cursor: SheetCursor,
    /// Snapshot of the active sheet grid taken on entry, used to restore the
    /// pre-preview state on cancel.
    pub saved_grid: GridBox,
}

/// Enter extrapolate mode on the current selection (or the cursor cell).
pub fn enter(app: &mut super::App) {
    let anchor = app.core.anchor.unwrap_or(app.core.cursor);
    let saved_grid = app.core.workbook.active_sheet().grid.clone();
    app.extrapolate = Some(ExtrapolateModal {
        seed_anchor: anchor,
        seed_cursor: app.core.cursor,
        saved_grid,
    });
    app.core.status = "Use arrows to extend selection, Enter to extrapolate, Esc to cancel".into();
}

/// True if extrapolate mode is currently active.
pub fn active(app: &super::App) -> bool {
    app.extrapolate.is_some()
}

/// The full materialized selection range (logical row indices, logical col
/// indices) spanned by the modal anchor..cursor, as ratatui computes it.
fn selection(app: &super::App) -> (Vec<usize>, Vec<usize>) {
    let a = app.core.anchor.unwrap_or(app.core.cursor);
    let b = app.core.cursor;
    let r0 = a.row.min(b.row);
    let r1 = a.row.max(b.row);
    let c0 = a.col.min(b.col);
    let c1 = a.col.max(b.col);
    ((r0..=r1).collect(), (c0..=c1).collect())
}

fn compute(app: &super::App) -> Vec<(CellAddr, String)> {
    let m = app.extrapolate.as_ref().expect("extrapolate active");
    let cells = {
        let (rows, cols) = selection(app);
        crate::extrapolate::extrapolate_cells(
            &app.core.workbook.active_sheet().grid,
            m.seed_anchor,
            m.seed_cursor,
            &rows,
            &cols,
        )
    };
    cells
}

/// The computed target fills for the current selection, WITHOUT mutating any
/// grid. Used by the pancurses backend to render the preview display-only.
pub fn preview_cells(app: &super::App) -> Vec<(CellAddr, String)> {
    if app.extrapolate.is_none() {
        return Vec::new();
    }
    compute(app)
}

/// Refresh the in-grid preview for the current selection: reset the grid to the
/// entry snapshot, then write the freshly-computed target fills so the widget
/// renders predicted values. The caller redraws/refills the widget afterwards.
pub fn refresh_preview(app: &mut super::App) {
    if app.extrapolate.is_none() {
        return;
    }
    let saved = app.extrapolate.as_ref().unwrap().saved_grid.clone();
    app.core.workbook.active_sheet_mut().grid = saved;
    let cells = compute(app);
    let sheet = app.core.workbook.active_sheet_mut();
    for (addr, value) in cells {
        sheet.grid.set(&addr, value);
    }
}

/// Cancel extrapolate mode, restoring the pre-preview grid (predicted values
/// are discarded, nothing is committed).
pub fn cancel(app: &mut super::App) {
    if let Some(m) = app.extrapolate.take() {
        app.core.workbook.active_sheet_mut().grid = m.saved_grid;
    }
    app.core.anchor = None;
    app.core.status.clear();
}

/// Commit the extrapolation, persisting a `FILL` op (matching the ratatui
/// reference's `Op::FillRange` serialization).
pub fn commit(app: &mut super::App) {
    let cells = compute(app);
    app.extrapolate = None;
    app.core.anchor = None;
    if cells.is_empty() {
        app.core.status = "Select cells with a pattern, then Extrapolate".into();
        return;
    }
    let sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
    // Ensure the preview values are present (they already are, but set them
    // again so a commit with a freshly-restored grid is also correct).
    let op = Op::FillRange { cells };
    if let Some(ref p) = app.core.path.clone() {
        let mut active_sheet = sheet_id;
        let _ = crate::io::commit_workbook_op(
            p,
            &mut app.core.offset,
            &mut app.core.workbook,
            &mut active_sheet,
            &WorkbookOp::SheetOp { sheet_id, op },
        );
        app.core.ops_applied = app.core.ops_applied.saturating_add(1);
    }
    app.core.status = "Extrapolated selection".into();
}
