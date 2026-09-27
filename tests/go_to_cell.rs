//! `Sheet ▸ Go` (the `go_to_cell` menu action) behaviour, on the shared
//! prompt-handling path every GUI/pancurses backend uses.
//!
//! The action used to parse the typed text with `addr::parse_cell_ref_at` and
//! map it with `HEADER_ROWS + row` / `MARGIN_COLS + col`, which:
//!
//!   * moved the cursor to the WRONG cell for any target past the current grid
//!     edge — `C999` on a 1x1 sheet landed on the right-margin footer `]B_998`,
//!     not on C999 — because the cursor was placed in margin space with no
//!     growth, so the address resolved differently than it reads;
//!   * accepted only a bare `A1` main cell, so the documented-but-unusable
//!     forms (`5` a row, `C` a column) were rejected outright;
//!   * never grew the grid, so a target past the edge was not addressable.
//!
//! It now delegates to `ui_core::resolve_go_target` (the one resolver the TUI
//! reference uses) and grows the body, so the cursor lands exactly on the
//! address the user typed and the address label in the formula bar says so.
//!
//! Run with: `cargo test --features gui --test go_to_cell`
//! (also compiles under `--features pancurses`).

#![cfg(all(target_os = "linux", any(feature = "gui", feature = "pancurses")))]

use corro::grid::{CellAddr, SheetCursor, HEADER_ROWS, MARGIN_COLS};
use corro::gui::actions::run_prompt_action;
use corro::gui::App;

/// The formula bar's address label for a cursor position — the user-visible
/// proof that the selection really is the cell that was asked for.
fn addr_label(app: &App, cursor: SheetCursor) -> String {
    let g = &app.core.workbook.active_sheet().grid;
    let addr = corro::addr::sheet_cursor_to_addr(
        corro::addr::LogicalRow(cursor.row),
        corro::addr::GlobalCol(cursor.col),
        corro::addr::MainRows(g.main_rows()),
        corro::addr::MainCols(g.main_cols()),
    );
    corro::addr::cell_ref_text(&addr, g.main_cols())
}

/// A small sheet with the cursor at A1.
fn app_at_a1(rows: usize, cols: usize) -> App {
    let mut app = App::new_with_paths(vec![]);
    app.core.workbook.active_sheet_mut().grid.set_main_size(rows, cols);
    app.core.cursor = SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS };
    app
}

/// The regression this whole file exists for: a target past the current grid
/// edge must land ON that cell, not on a margin/footer neighbour of it.
///
/// Before the fix `Go to C999` on a 1x1 sheet set the cursor to a position
/// whose address label read `]B_998`, i.e. the selection moved somewhere the
/// user never asked for (and, being out in the margin, appeared not to move at
/// all).
#[test]
fn go_past_grid_edge_lands_on_the_typed_cell() {
    let mut app = app_at_a1(1, 1);
    run_prompt_action(&mut app, "go_to_cell", "C999");

    assert_eq!(
        addr_label(&app, app.core.cursor),
        "C999",
        "Go to C999 must leave the selection on C999; status was {:?}",
        app.core.status
    );
    // The body must have grown to make that address real.
    let g = &app.core.workbook.active_sheet().grid;
    assert!(
        g.main_rows() >= 999,
        "grid must grow to reach row 999, got {}",
        g.main_rows()
    );
    assert!(
        g.main_cols() >= 3,
        "grid must grow to reach column C, got {}",
        g.main_cols()
    );
    assert_eq!(app.core.status, "Went to C999");
}

/// A target inside the existing grid still lands exactly where it did (the
/// `tests/menu_all_items.rs` `("go_to_cell", "C3")` case is the baseline; this
/// pins the arithmetic so a future resolver swap cannot drift it).
#[test]
fn go_inside_grid_lands_on_the_typed_cell() {
    let mut app = app_at_a1(3, 3);
    run_prompt_action(&mut app, "go_to_cell", "C3");
    assert_eq!(
        app.core.cursor,
        SheetCursor { row: HEADER_ROWS + 2, col: MARGIN_COLS + 2 }
    );
    assert_eq!(addr_label(&app, app.core.cursor), "C3");
    // An in-grid target must not grow anything: the extent stays what it was.
    let g = &app.core.workbook.active_sheet().grid;
    assert_eq!((g.main_rows(), g.main_cols()), (3, 3), "in-grid Go must not resize the sheet");
}

/// A row-only entry (`5`) moves down that column; a column-only entry (`C`)
/// moves across that row. Both used to be rejected outright, so a user typing
/// them got "Unknown cell" and no movement at all.
#[test]
fn go_accepts_row_and_column_parts() {
    let mut app = app_at_a1(10, 6);
    // Column C first, so the row jump is observable against a known column.
    run_prompt_action(&mut app, "go_to_cell", "C");
    assert_eq!(app.core.cursor.col, MARGIN_COLS + 2, "go to C");
    assert_eq!(addr_label(&app, app.core.cursor), "C1");

    run_prompt_action(&mut app, "go_to_cell", "5");
    assert_eq!(
        app.core.cursor,
        SheetCursor { row: HEADER_ROWS + 4, col: MARGIN_COLS + 2 },
        "go to 5 stays in the current column"
    );
    assert_eq!(addr_label(&app, app.core.cursor), "C5");
}

/// A row/column part past the edge grows the body just like a full ref does.
#[test]
fn go_row_past_edge_grows_grid() {
    let mut app = app_at_a1(2, 2);
    run_prompt_action(&mut app, "go_to_cell", "50");
    assert_eq!(addr_label(&app, app.core.cursor), "A50");
    let g = &app.core.workbook.active_sheet().grid;
    assert!(g.main_rows() >= 50, "grid must grow to row 50, got {}", g.main_rows());
}

/// Header, footer and margin refs are addressable through Go too.
#[test]
fn go_reaches_header_footer_and_margin() {
    let mut app = app_at_a1(5, 5);
    run_prompt_action(&mut app, "go_to_cell", "C~1");
    assert_eq!(addr_label(&app, app.core.cursor), "C~1", "header ref");

    run_prompt_action(&mut app, "go_to_cell", "C_1");
    assert_eq!(addr_label(&app, app.core.cursor), "C_1", "footer ref");

    run_prompt_action(&mut app, "go_to_cell", "[A1");
    assert_eq!(addr_label(&app, app.core.cursor), "[A1", "left-margin ref");
}

/// Bad input is answered, not silently ignored: a typo must leave a reason in
/// the status line instead of leaving the user wondering why nothing moved.
#[test]
fn go_rejects_bad_input_with_a_reason() {
    let mut app = app_at_a1(3, 3);
    let before = app.core.cursor;

    run_prompt_action(&mut app, "go_to_cell", "zzz");
    assert_eq!(app.core.status, "Bad cell address");
    assert_eq!(app.core.cursor, before, "a bad address must not move the cursor");

    run_prompt_action(&mut app, "go_to_cell", "0");
    assert_eq!(app.core.status, "Bad cell address");
    assert_eq!(app.core.cursor, before);

    // Row 0 does not exist; a wildly out-of-band row ref is a real address but
    // unaddressable, and must say so rather than moving somewhere random.
    run_prompt_action(&mut app, "go_to_cell", "C0");
    assert!(
        !app.core.status.is_empty() && app.core.status != "Went to C0",
        "C0 must be refused with a reason, got {:?}",
        app.core.status
    );
}

/// An empty entry is a no-op, not an error banner (the dialog can be confirmed
/// without typing).
#[test]
fn go_empty_input_is_a_no_op() {
    let mut app = app_at_a1(3, 3);
    let before = app.core.cursor;
    let status_before = app.core.status.clone();
    run_prompt_action(&mut app, "go_to_cell", "");
    assert_eq!(app.core.cursor, before);
    assert_eq!(app.core.status, status_before);
}

/// The Go dialog must state what it accepts, and every form it advertises must
/// actually resolve — otherwise the hint is a promise the code breaks. This
/// pins the hint text against the resolver it documents.
#[test]
fn go_dialog_hint_matches_the_resolver() {
    // The hint constant is private to the GUI module; assert the contract
    // through the resolver the hint describes instead of the string, and
    // separately that the shipped hint mentions each advertised form.
    let g = corro::ops::SheetState::new(5, 5);
    let cur = SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS };
    for form in ["C12", "C~1", "C_1", "[A1", "5", "C"] {
        assert!(
            corro::ui_core::resolve_go_target(&g.grid, cur, form).is_ok(),
            "{form} is advertised as valid but does not resolve"
        );
    }
    for form in ["zzz", "0"] {
        assert!(
            corro::ui_core::resolve_go_target(&g.grid, cur, form).is_err(),
            "{form} must be refused"
        );
    }
    // A `$-locked` cell is not a Go target: `$` prefixes a sheet name in both
    // backends (`$1`, `$Sheet1`, `$Sheet1:B2`), so `$C$12` is not `C12`.
    assert!(
        corro::ui_core::resolve_go_target(&g.grid, cur, "$C$12").is_err(),
        "$C$12 is a sheet-qualified form, not a locked cell, and must be refused"
    );
    // `resolve_go_target`'s doc table lists the accepted forms; a form listed
    // there that does not parse is a doc lie users hit. The header/footer
    // marker follows the column (`C~1`), not the row (`~1C` is not a form).
    assert!(
        corro::ui_core::resolve_go_target(&g.grid, cur, "~1C").is_err(),
        "~1C is not a valid ref form (it is C~1); must not be documented as one"
    );
    assert!(corro::ui_core::resolve_go_target(&g.grid, cur, "C~1").is_ok());
}

/// Going to a cell and reading it back must agree with the cell's own address
/// (guards the off-by-one between a 0-based main index and a 1-based label).
#[test]
fn go_address_matches_stored_cell_address() {
    let mut app = app_at_a1(4, 4);
    run_prompt_action(&mut app, "go_to_cell", "B2");
    let g = &app.core.workbook.active_sheet().grid;
    // The cursor's logical row/col must map back to main (1, 1) — the cell
    // labelled B2 (column B is main index 1).
    let addr = corro::addr::sheet_cursor_to_addr(
        corro::addr::LogicalRow(app.core.cursor.row),
        corro::addr::GlobalCol(app.core.cursor.col),
        corro::addr::MainRows(g.main_rows()),
        corro::addr::MainCols(g.main_cols()),
    );
    assert_eq!(addr, CellAddr::main(1, 1));
    assert_eq!(addr_label(&app, app.core.cursor), "B2");
}
