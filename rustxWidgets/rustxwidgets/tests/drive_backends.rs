// Copyright (c) 2026, Corro Project.
// Licensed under the Apache License, Version 2.0.
// See the LICENSE file in the project root for license information.

//! Driver tests for the pancurses backend: drive its *real* paint and input
//! code paths (no live terminal required — `render_model_to_grid` paints into
//! an in-memory `CellGrid`, and `InputEvent::from_pancurses` translates raw
//! pancurses keys). This is the pancurses counterpart of the ratatui/GTK3
//! driver tests: we feed the backend real data + real key input and assert the
//! resulting grid/event, instead of testing the model in isolation.

#![cfg(feature = "pancurses")]

use rustxwidgets::backends::pancurses_draw::render_model_to_grid;
use rustxwidgets::core::InputEvent;
use rustxwidgets::spreadsheet::SpreadsheetModel;

/// Build a 3-row single-column model (rows 1..=3) holding the given values.
/// Values go in model column 1 (column 0 is the row-label margin), and we use
/// letters so they never collide with the numeric row labels.
fn column_model(values: [&str; 3]) -> SpreadsheetModel {
    let mut m = SpreadsheetModel::new(6, 4);
    m.set_border_title("PanSort");
    m.set_cell(1, 1, values[0]);
    m.set_cell(2, 1, values[1]);
    m.set_cell(3, 1, values[2]);
    m
}

/// Find the first display row (y) that contains `val` in the painted grid.
fn first_row_with(grid: &rustxwidgets::backends::pancurses_draw::CellGrid, val: &str) -> Option<usize> {
    grid.cells
        .iter()
        .position(|row| row.iter().any(|c| c.ch.to_string() == val))
}

#[test]
fn drive_pancurses_backend_paints_unsorted_then_sorted() {
    // 1) Drive the real pancurses paint path with an UNSORTED model.
    let unsorted = column_model(["C", "A", "B"]);
    let grid = render_model_to_grid(&unsorted, 60u16, 16u16);
    let joined = grid.row_strings().join("");
    assert!(joined.contains("C"));
    assert!(joined.contains("A"));
    assert!(joined.contains("B"));

    // 2) Drive a "sort": reorder the cells (what a sort feature does) and
    //    re-render through the *same* pancurses paint path.
    let sorted = column_model(["A", "B", "C"]);
    let grid_sorted = render_model_to_grid(&sorted, 60u16, 16u16);

    // The painted grid must show the values in ascending visual order:
    // value 'A' sits above 'B' sits above 'C'.
    let y1 = first_row_with(&grid_sorted, "A").expect("sorted grid shows 'A'");
    let y2 = first_row_with(&grid_sorted, "B").expect("sorted grid shows 'B'");
    let y3 = first_row_with(&grid_sorted, "C").expect("sorted grid shows 'C'");
    assert!(
        y1 < y2 && y2 < y3,
        "sorted values must appear top-to-bottom as A,B,C (got rows {y1},{y2},{y3})"
    );

    // And the unsorted render must NOT already be in that order.
    let uy3 = first_row_with(&grid, "C").expect("unsorted grid shows 'C'");
    let uy1 = first_row_with(&grid, "A").expect("unsorted grid shows 'A'");
    assert_ne!(
        (uy1, uy3),
        (y1, y3),
        "precondition: unsorted render must differ from sorted render"
    );
}

#[test]
fn drive_pancurses_input_translation() {
    // Drive the real pancurses key -> InputEvent translation (the input path
    // the spreadsheet consumes).
    assert_eq!(
        InputEvent::from_pancurses(pancurses::Input::KeyUp),
        InputEvent::ArrowUp
    );
    assert_eq!(
        InputEvent::from_pancurses(pancurses::Input::KeyDown),
        InputEvent::ArrowDown
    );
    assert_eq!(
        InputEvent::from_pancurses(pancurses::Input::KeyLeft),
        InputEvent::ArrowLeft
    );
    assert_eq!(
        InputEvent::from_pancurses(pancurses::Input::KeyRight),
        InputEvent::ArrowRight
    );
    assert_eq!(
        InputEvent::from_pancurses(pancurses::Input::KeyEnter),
        InputEvent::Enter
    );
    assert_eq!(
        InputEvent::from_pancurses(pancurses::Input::KeyDC),
        InputEvent::Delete
    );
    assert_eq!(
        InputEvent::from_pancurses(pancurses::Input::Character('a')),
        InputEvent::Char('a')
    );
    assert_eq!(
        InputEvent::from_pancurses(pancurses::Input::KeyF1),
        InputEvent::F(1)
    );
}
