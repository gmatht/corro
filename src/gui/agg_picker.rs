//! Shared margin-aggregate picker state (GUI + pancurses backends).
//!
//! Mirrors [`super::special_picker`]: the selection machine lives here so
//! arrow/digit/Enter/Esc semantics can never drift between backends, and
//! backends are thin renderers that call [`open`] / [`step`] / [`set`] /
//! [`close`] / [`take`].
//!
//! The picker is offered when the user clicks (or invokes it on) a margin
//! aggregate *key* cell — the left/footer key column or the header/right key
//! band that [`crate::ops::margin_key_agg_func`] reads as a directive (e.g.
//! the seeded `TOTAL` in `]A~1` / `[A_1`). Committing writes the canonical
//! `==KEYWORD` directive, so the choice stays distinct from spreadsheet
//! formulas like `=MIN(A1)`.
//!
//! Pure logic over [`crate::gui::App`]; unit-tested below without a display.

use super::App;
use crate::grid::CellAddr;
use crate::ui_core::{
    agg_choice_directive, agg_choice_index_for_digit, agg_choice_index_for_func,
    agg_labelled_choices, agg_step_index, margin_agg_func_at,
};

/// Picker rows every backend renders (`"1: TOTAL"` … `"6: MEDIAN"`).
pub fn items() -> Vec<String> {
    agg_labelled_choices()
}

/// Current selection index (`None` = picker closed).
pub fn index(app: &App) -> Option<usize> {
    app.agg_picker
}

/// The cell whose key the picker is editing (`None` = closed).
pub fn target(app: &App) -> Option<CellAddr> {
    app.agg_picker_target.clone()
}

/// Whether `addr` is a cell the picker is *always* offered for, blank or
/// not: the aggregate **corner keys** `]?~1` (right-margin header band) and
/// `[A_1` (footer under the left-margin key column).
///
/// Per the design these always work — a blank `]B~1` is a valid "no
/// aggregate chosen yet" key, not a dead cell. The left-margin *row* column
/// (`[A1`, `[A2`, …) is excluded: it doubles as the item-name column, so
/// offering a picker on every blank cell there would fight typing item
/// names. A row key that already *holds* a directive still opens the picker
/// (see [`is_directive_cell`]).
pub fn is_always_agg_key_cell(addr: &CellAddr) -> bool {
    crate::ui_core::addr_is_always_agg_key(addr)
}

/// Whether `addr` currently holds an aggregate directive (TOTAL/MAX/…).
pub fn is_directive_cell(app: &App, addr: &CellAddr) -> bool {
    margin_agg_func_at(&app.core.workbook.active_sheet().grid, addr).is_some()
}

/// Whether the picker should open for `addr`: every always-key corner, and
/// any margin cell that already carries a directive (so `[A2 = =MAX` on a
/// row key still opens, while a plain item name does not).
pub fn is_agg_key_cell(app: &App, addr: &CellAddr) -> bool {
    is_always_agg_key_cell(addr) || is_directive_cell(app, addr)
}

/// Open the picker on `addr`, preselecting its current function (a `==MAX`
/// cell highlights MAX). Opens for the design's key cells even when blank —
/// a blank key defaults to TOTAL, matching the seeded `TOTAL` default.
pub fn open_for(app: &mut App, addr: &CellAddr) -> bool {
    let current = margin_agg_func_at(&app.core.workbook.active_sheet().grid, addr);
    if current.is_none() && !is_always_agg_key_cell(addr) {
        return false;
    }
    let idx = current.map(agg_choice_index_for_func).unwrap_or(0);
    app.agg_picker = Some(idx);
    app.agg_picker_target = Some(addr.clone());
    true
}

/// Open the picker from a cursor anywhere on the sheet, resolving the
/// aggregate key that governs it.
///
/// Resolution is [`crate::ui_core::resolve_margin_agg_key`] — the same rule
/// the ratatui and pancurses backends use — so all four cannot drift.
/// Returns false when no key governs the cursor, letting the caller explain
/// instead of showing an empty popup.
pub fn open_for_cursor(app: &mut App, cursor: &crate::grid::SheetCursor) -> bool {
    let grid = app.core.workbook.active_sheet().grid.clone();
    let Some((key_addr, func)) = crate::ui_core::resolve_margin_agg_key(&grid, cursor) else {
        return false;
    };
    app.agg_picker = Some(agg_choice_index_for_func(func));
    app.agg_picker_target = Some(key_addr);
    true
}

/// Close without committing.
pub fn close(app: &mut App) {
    app.agg_picker = None;
    app.agg_picker_target = None;
}

/// One arrow step (`delta` +1 for Down/Right, -1 for Up/Left), clamped.
/// No-op while closed.
pub fn step(app: &mut App, delta: i32) {
    if let Some(idx) = app.agg_picker {
        app.agg_picker = Some(agg_step_index(idx, delta));
    }
}

/// Absolute selection (widget → state sync), clamped. No-op while closed.
pub fn set(app: &mut App, idx: usize) {
    if app.agg_picker.is_some() {
        app.agg_picker = Some(idx.min(crate::ui_core::AGG_CHOICES.len() - 1));
    }
}

/// Commit the current selection: returns `(target, directive)` and closes.
/// `None` while closed.
pub fn take(app: &mut App) -> Option<(CellAddr, String)> {
    let idx = app.agg_picker.take()?;
    let target = app.agg_picker_target.take()?;
    let directive = agg_choice_directive(idx)?.to_string();
    Some((target, directive))
}

/// Digit hotkey → choice index (`1`..=`7` → 0..=6, bounded by the list).
pub fn index_for_digit(digit: char) -> Option<usize> {
    agg_choice_index_for_digit(digit)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::grid::{ColumnAddr, MARGIN_COLS, HEADER_ROWS};

    fn app_with_total_seed() -> (App, CellAddr) {
        let mut a = App::new_with_paths(vec![]);
        // Fresh documents are seeded, but unit tests elsewhere clear them;
        // set the canonical key explicitly so the target is deterministic.
        let addr = CellAddr::Footer {
            row: 0,
            col: ColumnAddr::Left(MARGIN_COLS - 1),
        };
        a.core
            .workbook
            .active_sheet_mut()
            .grid
            .set(&addr, "TOTAL".into());
        (a, addr)
    }

    #[test]
    fn rows_are_labelled_in_navigation_order() {
        let rows = items();
        let n = crate::ui_core::AGG_CHOICES.len();
        assert_eq!(rows.len(), n);
        assert_eq!(rows[0], "1: TOTAL");
        assert_eq!(rows[1], "2: MAX");
        assert_eq!(rows[5], "6: MEDIAN");
        // The trailing row is the canned column template.
        assert_eq!(rows[n - 1], format!("{n}: =A*B -- AB"));
    }

    #[test]
    fn open_preselects_current_function() {
        let (mut a, addr) = app_with_total_seed();
        a.core
            .workbook
            .active_sheet_mut()
            .grid
            .set(&addr, "==MAX".into());
        assert!(open_for(&mut a, &addr));
        assert_eq!(index(&a), Some(1), "==MAX highlights MAX");
        assert_eq!(target(&a), Some(addr));
    }

    #[test]
    fn open_refuses_non_agg_margin_cells() {
        let (mut a, addr) = app_with_total_seed();
        let plain = CellAddr::Left {
            col: MARGIN_COLS - 1,
            row: 3,
        };
        a.core
            .workbook
            .active_sheet_mut()
            .grid
            .set(&plain, "Hammers".into());
        assert!(
            !open_for(&mut a, &plain),
            "ordinary row-key text must not open the picker"
        );
        assert_eq!(index(&a), None);
        // Main cells are never keys either.
        assert!(!is_agg_key_cell(
            &a,
            &CellAddr::Main { row: 0, col: 0 }
        ));
        // ...but the seeded key is.
        assert!(is_agg_key_cell(&a, &addr));
    }

    #[test]
    fn open_step_take_writes_canonical_directive() {
        let (mut a, addr) = app_with_total_seed();
        assert!(open_for(&mut a, &addr));
        step(&mut a, 1); // TOTAL -> MAX
        let (picked, directive) = take(&mut a).expect("commit");
        assert_eq!(picked, addr);
        assert_eq!(directive, "==MAX", "preferred form stays distinct from =MAX(...)");
        assert_eq!(index(&a), None, "take closes the picker");
        assert_eq!(target(&a), None);
    }

    #[test]
    fn steps_clamp_and_digits_map() {
        let (mut a, addr) = app_with_total_seed();
        open_for(&mut a, &addr);
        step(&mut a, -1);
        assert_eq!(index(&a), Some(0));
        let last = crate::ui_core::AGG_CHOICES.len() - 1;
        step(&mut a, 25);
        assert_eq!(index(&a), Some(last));
        set(&mut a, 99);
        assert_eq!(index(&a), Some(last));
        assert_eq!(index_for_digit('1'), Some(0));
        assert_eq!(index_for_digit('6'), Some(5));
        assert_eq!(index_for_digit('7'), Some(6), "the template row");
        assert_eq!(index_for_digit('0'), None);
    }

    #[test]
    fn closed_picker_ignores_keys_and_take() {
        let (mut a, _) = app_with_total_seed();
        step(&mut a, 1);
        set(&mut a, 3);
        assert_eq!(take(&mut a), None);
        assert_eq!(index(&a), None);
    }

    /// Committing actually changes what the sheet computes: a MAX key over
    /// `1, 5` displays 5, where the seeded SUM displayed 6.
    #[test]
    fn committed_directive_changes_the_computed_display() {
        let (mut a, addr) = app_with_total_seed();
        a.core
            .workbook
            .active_sheet_mut()
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "1".into());
        a.core
            .workbook
            .active_sheet_mut()
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "5".into());

        // The key drives the column aggregate, so read the computed value
        // the sheet shows for that column.
        let shown = |a: &App| {
            crate::agg::helpers::footer_special_col_aggregate(
                &a.core.workbook.active_sheet().grid,
                crate::ops::AggFunc::Sum,
                crate::grid::MARGIN_COLS,
                2,
                1,
            )
            .unwrap_or_default()
        };
        assert_eq!(shown(&a), "6", "SUM of 1,5");

        assert!(open_for(&mut a, &addr));
        // Select MAX (row 2) and commit through the shared state machine.
        step(&mut a, 1);
        let (target, directive) = take(&mut a).expect("commit");
        super::super::actions::commit_cell(&mut a, target, directive);

        assert_eq!(
            a.core.workbook.active_sheet().grid.get(&addr).as_deref(),
            Some("==MAX")
        );
        // Reopening preselects MAX, proving the written directive parses
        // back through the same path the picker uses.
        assert!(open_for(&mut a, &addr));
        assert_eq!(index(&a), Some(1));
    }

    /// Every *aggregate* choice the picker can write must parse back to the
    /// same function, so reopening highlights what was just written. The
    /// canned template entry is not an aggregate (it has no function to
    /// highlight), so it is checked separately below.
    #[test]
    fn every_directive_round_trips_through_the_parser() {
        use crate::ops::margin_key_agg_func;
        for idx in 0..crate::ui_core::AGG_CHOICES.len() {
            let directive = agg_choice_directive(idx).unwrap();
            if crate::ui_core::is_template_choice_directive(directive) {
                assert!(
                    margin_key_agg_func(directive).is_none(),
                    "{directive:?} is a template, not an aggregate"
                );
                continue;
            }
            let func = margin_key_agg_func(directive)
                .unwrap_or_else(|| panic!("{directive:?} must parse as an aggregate"));
            assert_eq!(
                agg_choice_index_for_func(func),
                idx,
                "{directive:?} must reopen on its own row"
            );
        }
    }

    /// The canned template entry writes its text verbatim.
    #[test]
    fn the_template_choice_writes_the_literal_formula() {
        let idx = crate::ui_core::AGG_CHOICES
            .iter()
            .position(|(_, d)| crate::ui_core::is_template_choice_directive(d))
            .expect("one template entry");
        assert_eq!(agg_choice_directive(idx), Some("=A*B -- AB"));
        assert_eq!(
            crate::ui_core::agg_choice_label(idx),
            Some("=A*B -- AB")
        );
    }

    /// The key band predicate: header/right key cells and the left/footer
    /// key column qualify; main body cells never do.
    #[test]
    fn key_cell_predicate_covers_the_margin_key_bands() {
        assert!(crate::ui_core::addr_is_margin_agg_key(&CellAddr::Header {
            row: (HEADER_ROWS - 1) as u32,
            col: ColumnAddr::Right(0),
        }));
        // The right-margin key band is the header corner (`]A~1`), never the
        // right-margin body cells (`]A1`, which hold computed totals).
        assert!(!crate::ui_core::addr_is_margin_agg_key(&CellAddr::Right {
            col: 0,
            row: 0,
        }));
        assert!(crate::ui_core::addr_is_margin_agg_key(&CellAddr::Header {
            row: (HEADER_ROWS - 1) as u32,
            col: ColumnAddr::Right(1),
        }));
        // A main-column header is a label, not a key.
        assert!(!crate::ui_core::addr_is_margin_agg_key(&CellAddr::Header {
            row: (HEADER_ROWS - 1) as u32,
            col: ColumnAddr::Main(0),
        }));
        assert!(crate::ui_core::addr_is_margin_agg_key(&CellAddr::Left {
            col: MARGIN_COLS - 1,
            row: 0,
        }));
        assert!(crate::ui_core::addr_is_margin_agg_key(&CellAddr::Footer {
            row: 0,
            col: ColumnAddr::Left(MARGIN_COLS - 1),
        }));
        assert!(!crate::ui_core::addr_is_margin_agg_key(&CellAddr::Main {
            row: 0,
            col: 0,
        }));
        // A non-key left-margin column is not a key band.
        assert!(!crate::ui_core::addr_is_margin_agg_key(&CellAddr::Left {
            col: 0,
            row: 0,
        }));
    }
}

#[cfg(test)]
mod blank_key_repro {
    use super::*;
    use crate::grid::{ColumnAddr, HEADER_ROWS, MARGIN_COLS};

    /// REPRO: clicking `]B~1` when it is BLANK must still offer the picker.
    /// Per the design every `]?~1` / `[A??` key cell always works; requiring
    /// the cell to already contain TOTAL/MAX made blank ones dead.
    #[test]
    fn clicking_a_blank_right_margin_key_opens_the_picker() {
        let mut a = App::new_with_paths(vec![]);
        // Blank the seeded `]A~1` so the sheet has NO directives at all.
        a.core.workbook.active_sheet_mut().grid.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: ColumnAddr::Right(0),
            },
            String::new(),
        );
        let b_key = CellAddr::Header {
            row: (HEADER_ROWS - 1) as u32,
            col: ColumnAddr::Right(1),
        };
        assert_eq!(
            a.core.workbook.active_sheet().grid.get(&b_key),
            None,
            "precondition: ]B~1 is blank"
        );

        assert!(
            open_for(&mut a, &b_key),
            "clicking a blank ]B~1 must open the picker"
        );
        assert_eq!(index(&a), Some(0), "blank keys default to TOTAL");
    }

    /// The footer corner `[A_1` is an always-key cell: blank or not, it
    /// opens the picker (defaulting to TOTAL), matching `]?~1`.
    #[test]
    fn clicking_a_blank_footer_corner_opens_the_picker() {
        let mut a = App::new_with_paths(vec![]);
        let corner = CellAddr::Footer {
            row: 0,
            col: ColumnAddr::Left(MARGIN_COLS - 1),
        };
        a.core
            .workbook
            .active_sheet_mut()
            .grid
            .set(&corner, String::new());
        assert!(open_for(&mut a, &corner), "blank `[A_1` must open the picker");
        assert_eq!(index(&a), Some(0), "blank keys default to TOTAL");
    }

    /// A left-margin *row* key that already holds a directive still opens
    /// (the row column doubles as the item-name column, so only cells with
    /// a directive are offered there).
    #[test]
    fn a_row_key_holding_a_directive_still_opens() {
        let mut a = App::new_with_paths(vec![]);
        let key = CellAddr::Left {
            col: MARGIN_COLS - 1,
            row: 2,
        };
        a.core
            .workbook
            .active_sheet_mut()
            .grid
            .set(&key, "==MAX".into());
        assert!(open_for(&mut a, &key));
        assert_eq!(index(&a), Some(1), "==MAX highlights MAX");
    }

    /// A blank left-margin *row* cell is NOT offered: that column is also
    /// the item-name column (`Hammers`, …), and a picker on every blank
    /// cell there would fight typing names.
    #[test]
    fn a_blank_row_name_cell_does_not_open() {
        let mut a = App::new_with_paths(vec![]);
        let name_cell = CellAddr::Left {
            col: MARGIN_COLS - 1,
            row: 3,
        };
        a.core
            .workbook
            .active_sheet_mut()
            .grid
            .set(&name_cell, String::new());
        assert!(
            !open_for(&mut a, &name_cell),
            "blank item-name cells must not pop the picker"
        );
    }
}

#[cfg(test)]
mod footer_key_column_repro {
    use super::*;
    use crate::grid::{ColumnAddr, MARGIN_COLS};

    fn footer(app: &mut App, row: u32) -> CellAddr {
        let a = CellAddr::Footer {
            row,
            col: ColumnAddr::Left(MARGIN_COLS - 1),
        };
        app.core.workbook.active_sheet_mut().grid.set(&a, String::new());
        a
    }

    /// REPRO: the `[A_n` footer column is the aggregate **key column** — every
    /// row of it is a key (`_1` may be a label like "Grand", `_2`/`_3` hold
    /// TOTAL/MAX/MIN). Clicking `[A_2` / `[A_3` must pop the dropdown, blank
    /// or not, exactly like `[A_1` and `]?~1`.
    #[test]
    fn clicking_A_2_and_A_3_footer_keys_opens_the_dropdown() {
        for row in 1..=4u32 {
            let mut a = App::new_with_paths(vec![]);
            let key = footer(&mut a, row);
            assert!(
                is_agg_key_cell(&a, &key),
                "[A_{} is a key cell in the footer key column",
                row + 1
            );
            assert!(
                open_for(&mut a, &key),
                "clicking the blank [A_{} key must open the dropdown",
                row + 1
            );
            assert_eq!(index(&a), Some(0), "blank keys default to TOTAL");
        }
    }

    /// A footer key that already holds a directive preselects it.
    #[test]
    fn footer_key_with_a_directive_preselects_it() {
        let mut a = App::new_with_paths(vec![]);
        let key = CellAddr::Footer {
            row: 2,
            col: ColumnAddr::Left(MARGIN_COLS - 1),
        };
        a.core
            .workbook
            .active_sheet_mut()
            .grid
            .set(&key, "==MAX".into());
        assert!(open_for(&mut a, &key));
        assert_eq!(index(&a), Some(1), "==MAX highlights MAX");
    }

    /// Footer cells under a MAIN column are values, not keys.
    #[test]
    fn footer_cells_outside_the_key_column_are_not_keys() {
        let mut a = App::new_with_paths(vec![]);
        for col in [ColumnAddr::Main(0), ColumnAddr::Left(0), ColumnAddr::Right(0)] {
            let addr = CellAddr::Footer { row: 1, col };
            a.core
                .workbook
                .active_sheet_mut()
                .grid
                .set(&addr, String::new());
            assert!(
                !is_agg_key_cell(&a, &addr),
                "{addr:?} is not in the footer key column"
            );
        }
    }

    /// The F3 / menu resolver agrees for the whole footer key column.
    #[test]
    fn resolver_finds_any_footer_key_row() {
        let mut a = App::new_with_paths(vec![]);
        let key = footer(&mut a, 2);
        let grid = a.core.workbook.active_sheet().grid.clone();
        let (lr, gc) = crate::addr::addr_to_sheet_cursor(
            &key,
            crate::addr::MainRows(grid.main_rows()),
            crate::addr::MainCols(grid.main_cols()),
        );
        let cursor = crate::grid::SheetCursor { row: lr.0, col: gc.0 };
        let (found, func) =
            crate::ui_core::resolve_margin_agg_key(&grid, &cursor).expect("footer key resolves");
        assert_eq!(found, key);
        assert_eq!(func, crate::ops::AggFunc::Sum, "blank defaults to TOTAL");
    }
}
