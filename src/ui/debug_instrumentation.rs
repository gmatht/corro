//! Debug-only tracing helpers.
//!
//! The whole function is gated, not just its body: in a release build it is a
//! no-op, so keeping it would leave `addr`/`ui_main_cols`/`workbook_main_cols`
//! unused and warn on every release build. Gating the item keeps the callers
//! (themselves inside `debug_assertions` blocks) type-checked only where the
//! work actually happens.

#[cfg(debug_assertions)]
use crate::grid::CellAddr;

#[cfg(debug_assertions)]
pub fn trace_setcell_construction(addr: &CellAddr, ui_main_cols: usize, workbook_main_cols: usize) {
    crate::debug_log::log(&format!(
        "DEBUG SetCell constructed: addr={:?} ui_main_cols={} workbook_main_cols={}",
        addr, ui_main_cols, workbook_main_cols
    ));
}
