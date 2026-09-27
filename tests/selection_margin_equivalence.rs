//! Equivalence: extending a selection (Shift+arrows) never leaves the main
//! area.
//!
//! The default (ratatui) backend is the oracle. Its Shift+arrow arms refuse
//! to step out of the body at the top/left edge and *grow the main area* at
//! the bottom/right edge, so a selection begun inside the body always ends
//! inside the body. Plain navigation is deliberately different: it only grows
//! while trailing blanks are below the navigation threshold, so plain Down
//! *can* land in the footer (and plain Right in the right margin).
//!
//! The GUI/pancurses backends share one implementation of the same rule,
//! [`corro::ui_core::grow_grid_for_selection_edge`]. This test drives both
//! with the same fake key sequence and asserts they agree on the cursor, the
//! anchor and the resulting main extent at every step — that is what stops the
//! non-ratatui backends drifting back into the margins on Shift+Down/Right.
//!
//! Runs headlessly on the default features. The file is gated on `ratatui`:
//! it drives `corro::ui::App` (the reference backend, the oracle here) with
//! crossterm key events, neither of which is compiled into a gui/pancurses-only
//! build.

#![cfg(feature = "ratatui")]

use corro::grid::{Grid, GridBox, SheetCursor, HEADER_ROWS, MARGIN_COLS};
use corro::ui::App;
use crossterm::event::{KeyCode, KeyEvent, KeyModifiers};

fn shift(app: &mut App, code: KeyCode) {
    app.bench_handle_key(KeyEvent::new(code, KeyModifiers::SHIFT))
        .ok();
}

/// The cursor is the last main cell of an empty `rows`x`cols` body — the
/// bottom/right corner, where a selection step would otherwise walk into the
/// footer / right margin. An empty body has >= NAV_BLANK_ROWS blank rows, so
/// the plain-navigation growth rule does not fire; only the selection rule
/// keeps the cursor inside.
fn corner_cursor(rows: usize, cols: usize) -> SheetCursor {
    SheetCursor {
        row: HEADER_ROWS + rows - 1,
        col: MARGIN_COLS + cols - 1,
    }
}

/// Whether `cursor` is inside the main body of a `rows`x`cols` grid.
fn in_main(cursor: SheetCursor, rows: usize, cols: usize) -> bool {
    cursor.row >= HEADER_ROWS
        && cursor.row < HEADER_ROWS + rows
        && cursor.col >= MARGIN_COLS
        && cursor.col < MARGIN_COLS + cols
}

#[test]
fn shift_down_from_bottom_stays_in_main_like_ratatui() {
    let mut app = App::new(None);
    app.state.grid.set_main_size(2, 2);
    app.cursor = corner_cursor(2, 2);
    app.anchor = None;

    let mut core = GridBox::from(Grid::new(2, 2));
    let (mut row, mut col) = (app.cursor.row, app.cursor.col);

    for step in 0..4 {
        shift(&mut app, KeyCode::Down);
        let (nr, nc) = corro::ui_core::grow_grid_for_selection_edge(&mut core, row, col, 1, 0);
        row = nr;
        col = nc;

        assert!(
            in_main(
                app.cursor,
                app.state.grid.main_rows(),
                app.state.grid.main_cols()
            ),
            "step {step}: ratatui arrowed out of the main area into {:?}",
            app.cursor
        );
        assert!(
            in_main(SheetCursor { row, col }, core.main_rows(), core.main_cols()),
            "step {step}: shared rule arrowed out of the main area into ({row}, {col})"
        );
        assert_eq!(
            app.cursor.row, row,
            "step {step}: ratatui cursor row diverged from the shared rule"
        );
        assert_eq!(
            app.cursor.col, col,
            "step {step}: ratatui cursor col diverged from the shared rule"
        );
        assert_eq!(
            app.state.grid.main_rows(),
            core.main_rows(),
            "step {step}: main_rows diverged"
        );
        assert_eq!(
            app.state.grid.main_cols(),
            core.main_cols(),
            "step {step}: main_cols diverged"
        );
        // The anchor is set by the first step and preserved by the rest.
        assert_eq!(
            app.anchor,
            Some(corner_cursor(2, 2)),
            "step {step}: anchor must stay at the selection start"
        );
    }
}

#[test]
fn shift_right_from_right_edge_stays_in_main_like_ratatui() {
    let mut app = App::new(None);
    app.state.grid.set_main_size(2, 2);
    app.cursor = corner_cursor(2, 2);
    app.anchor = None;

    let mut core = GridBox::from(Grid::new(2, 2));
    let (mut row, mut col) = (app.cursor.row, app.cursor.col);

    for step in 0..4 {
        shift(&mut app, KeyCode::Right);
        let (nr, nc) = corro::ui_core::grow_grid_for_selection_edge(&mut core, row, col, 0, 1);
        row = nr;
        col = nc;

        assert!(
            in_main(
                app.cursor,
                app.state.grid.main_rows(),
                app.state.grid.main_cols()
            ),
            "step {step}: ratatui arrowed out of the main area into {:?}",
            app.cursor
        );
        assert!(
            in_main(SheetCursor { row, col }, core.main_rows(), core.main_cols()),
            "step {step}: shared rule arrowed out of the main area into ({row}, {col})"
        );
        assert_eq!(app.cursor.row, row, "step {step}: cursor row diverged");
        assert_eq!(app.cursor.col, col, "step {step}: cursor col diverged");
        assert_eq!(
            app.state.grid.main_cols(),
            core.main_cols(),
            "step {step}: main_cols diverged"
        );
        assert_eq!(
            app.anchor,
            Some(corner_cursor(2, 2)),
            "step {step}: anchor must stay at the selection start"
        );
    }
}

/// Walking down/right alternately from the bottom-right corner: every step is
/// a selection step, so nothing may ever leave the body — while the main area
/// grows under it, exactly as ratatui does.
#[test]
fn shift_zigzag_from_the_corner_never_leaves_the_body() {
    let mut app = App::new(None);
    app.state.grid.set_main_size(2, 2);
    app.cursor = corner_cursor(2, 2);
    app.anchor = None;

    let mut core = GridBox::from(Grid::new(2, 2));
    let (mut row, mut col) = (app.cursor.row, app.cursor.col);

    let seq = [
        (KeyCode::Down, 1isize, 0isize),
        (KeyCode::Right, 0, 1),
        (KeyCode::Down, 1, 0),
        (KeyCode::Right, 0, 1),
        (KeyCode::Up, -1, 0),
        (KeyCode::Left, 0, -1),
    ];
    for (i, (code, dr, dc)) in seq.into_iter().enumerate() {
        shift(&mut app, code);
        let (nr, nc) = corro::ui_core::grow_grid_for_selection_edge(&mut core, row, col, dr, dc);
        row = nr;
        col = nc;
        assert!(
            in_main(
                app.cursor,
                app.state.grid.main_rows(),
                app.state.grid.main_cols()
            ),
            "step {i}: ratatui left the body at {:?}",
            app.cursor
        );
        assert!(
            in_main(SheetCursor { row, col }, core.main_rows(), core.main_cols()),
            "step {i}: shared rule left the body at ({row}, {col})"
        );
        assert_eq!((app.cursor.row, app.cursor.col), (row, col), "step {i} diverged");
        assert_eq!(
            (app.state.grid.main_rows(), app.state.grid.main_cols()),
            (core.main_rows(), core.main_cols()),
            "step {i}: extent diverged"
        );
    }
}

/// The contrast that makes the rule meaningful: *plain* Down from the same
/// corner does leave the body (into the footer), because plain navigation uses
/// the trailing-blank growth policy instead. If this ever stops holding, the
/// selection rule above is no longer testing anything.
#[test]
fn plain_down_from_the_corner_does_enter_the_footer_like_ratatui() {
    let mut app = App::new(None);
    app.state.grid.set_main_size(2, 2);
    app.cursor = corner_cursor(2, 2);
    app.anchor = None;

    app.bench_handle_key(KeyEvent::new(KeyCode::Down, KeyModifiers::NONE))
        .ok();

    assert!(
        !in_main(
            app.cursor,
            app.state.grid.main_rows(),
            app.state.grid.main_cols()
        ),
        "precondition: plain Down on an all-blank body must reach the footer, got {:?}",
        app.cursor
    );
}
