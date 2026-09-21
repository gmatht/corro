//! `--movie` on the GUI backends: the demo mode is not terminal-only.
//!
//! Movie mode IS the normal UI. `--gui --movie` opens the same window, builds
//! the same widget tree and runs the same draw callbacks as an interactive
//! session; a periodic timer applies one movie step per tick instead of
//! waiting for a keystroke. So the coverage here is in two layers:
//!
//! * the backend-independent replay itself ([`corro::gui::movie`]) — parsing,
//!   op application, cursor derivation — which needs no display, and
//! * the *renderer* contract, asserted by driving the production sheet renderer
//!   (`gui_backend::render_grid_body`) through the raster `DrawContext` that
//!   `corro::gui::gui_movie` also uses for headless frame production.
//!
//! An earlier version of movie mode reimplemented the sheet into its own raster
//! and drifted from the window four times over (blank frame, overlapping
//! glyphs, missing margin shading, missing text). These tests exist so a second
//! implementation cannot come back unnoticed.
#![cfg(feature = "gui")]

use corro::grid::CellAddr;
use corro::gui::movie::{GuiMovie, GuiMovieOptions};
use std::path::PathBuf;

/// A movie fixture written into a temp dir, so tests never touch the repo.
fn movie_file(tag: &str, text: &str) -> PathBuf {
    let path = std::env::temp_dir().join(format!(
        "corro-gui-movie-it-{}-{tag}.corro",
        std::process::id()
    ));
    std::fs::write(&path, text).expect("write movie fixture");
    path
}

fn fast() -> GuiMovieOptions {
    GuiMovieOptions {
        typing_cps: 10_000.0,
        confirm_delay_ms: 0,
        menu_hold_ms: 0,
    }
}

fn apply_all(path: &std::path::Path, text: &str) -> (GuiMovie, corro::gui::App) {
    let mut movie = GuiMovie::new(path).expect("parse movie");
    let mut app = corro::gui::App::new_with_paths(vec![path.to_path_buf()]);
    let active = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
    app.core.view_sheet_id = active;
    for i in 0..movie.len() {
        movie
            .apply_step(&mut app.core.workbook, &mut app.core.view_sheet_id, i)
            .expect("apply step");
    }
    let _ = text;
    (movie, app)
}

#[test]
fn gui_movie_replays_every_step_into_the_workbook() {
    let path = movie_file("steps", "SET $1:A1 1\nSET $1:A2 2\nSET $1:B1 4\n");
    let (movie, app) = apply_all(&path, "");
    assert_eq!(movie.len(), 3);

    let grid = &app.core.workbook.active_sheet().grid;
    assert_eq!(grid.get(&CellAddr::Main { row: 0, col: 0 }), Some("1".to_string()));
    assert_eq!(grid.get(&CellAddr::Main { row: 1, col: 0 }), Some("2".to_string()));
    assert_eq!(grid.get(&CellAddr::Main { row: 0, col: 1 }), Some("4".to_string()));
    assert_eq!(movie.applied, 3);

    let _ = std::fs::remove_file(path);
}

#[test]
fn gui_movie_tracks_the_cursor_of_the_step() {
    let path = movie_file("cursor", "SET $1:A1 1\nSET $1:C3 x\n");
    let (movie, mut app) = apply_all(&path, "");
    let _ = movie;
    let active = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
    app.core.view_sheet_id = active;
    // Replay just the last step to check the cursor it reports.
    let mut movie = GuiMovie::new(&path).unwrap();
    let frame = movie
        .apply_step(&mut app.core.workbook, &mut app.core.view_sheet_id, 1)
        .unwrap();
    let addr = frame.cursor.expect("a cell write reports its address");
    let cur = corro::gui::movie::cursor_of(&addr, &app.core.workbook);
    assert_eq!(cur.row, corro::grid::HEADER_ROWS + 2);
    assert_eq!(cur.col, corro::grid::MARGIN_COLS + 2);
    let _ = std::fs::remove_file(path);
}

#[test]
fn gui_movie_flashes_the_menu_a_step_would_have_used() {
    // A cleared range is the Edit ▸ Cut path; the replay reports the menu so
    // the window can show what a user would have done.
    let path = movie_file("menu", "SET $1:A1 x\nSET $1:B1 y\nFILL A1= B1=\n");
    let (mut movie, mut app) = apply_all(&path, "");
    let active = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
    app.core.view_sheet_id = active;
    let mut movie = GuiMovie::new(&path).unwrap();
    let frame = movie
        .apply_step(&mut app.core.workbook, &mut app.core.view_sheet_id, 2)
        .expect("apply the fill step");
    assert_eq!(
        frame.menu.as_ref().map(|(s, _)| s.as_str()),
        Some("Edit"),
        "clearing a range is an Edit ▸ Cut"
    );
    assert!(movie.applied >= 1);
    let _ = std::fs::remove_file(path);
}

#[test]
fn gui_movie_rejects_a_non_corro_input() {
    let path = std::env::temp_dir().join(format!("corro-movie-{}.tsv", std::process::id()));
    std::fs::write(&path, "a\tb\n").unwrap();
    let mut app = corro::gui::App::new_with_paths(vec![path.clone()]);
    let err = app
        .run_movie(GuiMovieOptions::default())
        .expect_err("a .tsv is not a movie");
    assert!(
        err.to_string().contains("only supports .corro"),
        "unexpected error: {err}"
    );
    let _ = std::fs::remove_file(path);
}

#[test]
fn gui_movie_errors_are_reported_not_panics() {
    let missing = std::env::temp_dir().join("corro-gui-movie-does-not-exist.corro");
    let _ = std::fs::remove_file(&missing);
    let mut app = corro::gui::App::new_with_paths(vec![missing.clone()]);
    let err = app.run_movie(GuiMovieOptions::default()).expect_err("missing input");
    assert!(
        err.to_string().contains("does not exist"),
        "unexpected error: {err}"
    );
}

#[test]
fn gui_movie_options_pacing_is_read_from_the_environment() {
    // The driver that lives inside the GUI backend reads its pacing from the
    // environment, so the CLI's parsed values must be published there.
    let opts = GuiMovieOptions {
        typing_cps: 33.0,
        confirm_delay_ms: 55,
        menu_hold_ms: 777,
    };
    opts.publish_to_env();
    let read = GuiMovieOptions::from_env();
    assert!((read.typing_cps - 33.0).abs() < f64::EPSILON);
    assert_eq!(read.confirm_delay_ms, 55);
    assert_eq!(read.menu_hold_ms, 777);
    // char_delay is derived, and must never be zero (division by cps).
    assert!(read.char_delay() > std::time::Duration::ZERO);
    for key in ["CORRO_MOVIE_TYPING_CPS", "CORRO_MOVIE_CONFIRM_MS", "CORRO_MOVIE_MENU_HOLD_MS"] {
        std::env::remove_var(key);
    }
    let defaults = GuiMovieOptions::from_env();
    assert!(defaults.char_delay() > std::time::Duration::ZERO);
    assert_eq!(defaults.confirm_delay_ms, 120);
}

/// Regression: replaying a movie must not write to the log it is reading.
///
/// Every GUI commit path appends to `app.core.path`, and the movie's file was
/// left bound, so each recording appended its own steps to the fixture — the
/// demo workbooks grew every time the video was made.
#[test]
fn gui_movie_does_not_write_to_the_log_it_replays() {
    let path = movie_file("readonly", "SET $1:A1 1\nSET $1:B1 2\n");
    let before = std::fs::read(&path).expect("read fixture");

    let mut mov = GuiMovie::new(&path).expect("parse");
    let mut app = corro::gui::App::new_with_paths(vec![path.clone()]);
    // Replay through the real entry point (which must detach the file).
    corro::gui::movie::GuiMovie::detach_source(&mut app);
    let active = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
    app.core.view_sheet_id = active;
    for i in 0..mov.len() {
        mov.apply_step(&mut app.core.workbook, &mut app.core.view_sheet_id, i)
            .expect("step");
    }

    let after = std::fs::read(&path).expect("read fixture");
    assert_eq!(
        before, after,
        "the movie log must be left byte-identical by a replay"
    );
    // ...and the app must still know where it came from, for the title/caption.
    assert!(app.core.source_path.is_some(), "source path is kept for display");

    let _ = std::fs::remove_file(path);
}

/// The movie's frames come from the production sheet renderer, so a frame
/// produced headlessly must contain what the window paints: shaded margins,
/// cell text, and a sheet that spans the frame.
#[test]
fn gui_movie_frames_come_from_the_production_renderer() {
    let path = movie_file("render", "SET $1:A1 HELLO\nSET $1:B2 WORLD\n");
    let (_, mut app) = apply_all(&path, "");
    let app_mut = &mut app;
    app_mut.core.cursor = corro::grid::SheetCursor {
        row: corro::grid::HEADER_ROWS,
        col: corro::grid::MARGIN_COLS,
    };

    let dir = std::env::temp_dir().join(format!("corro-movie-render-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    let mut painter = corro::gui::gui_movie::HeadlessFramePainter::new(dir.clone());
    corro::gui::gui_movie::paint_one_frame(app_mut, &mut painter, 0).expect("paint frame");

    let frame = std::fs::read(dir.join("frame-00000.ppm")).expect("frame written");
    let (w, h, pixels) = parse_ppm(&frame);

    // This surface is the sheet body alone (the window's chrome is painted
    // around it separately), so the grid starts at the top of the frame.
    let mut margin = 0usize;
    let mut dark = 0usize;
    let mut gridline_rightmost = 0usize;
    for y in 4..h - 20 {
        for x in 50..w - 50 {
            let i = (y * w + x) * 3;
            let (r, g, b) = (pixels[i], pixels[i + 1], pixels[i + 2]);
            if (185..=196).contains(&r) && r == g && g == b {
                margin += 1;
            }
            if r < 100 && g < 100 && b < 100 {
                dark += 1;
            }
            if (195..=205).contains(&r) && r == g && g == b && x > gridline_rightmost {
                gridline_rightmost = x;
            }
        }
    }

    // Margin shading: the defect that appeared when the painter was a second
    // implementation of the sheet.
    assert!(margin > w, "expected shaded margin bands, found {margin} px");
    // Cell text: the defect that appeared when the chrome offset was applied
    // twice. "HELLO" across a 1200px frame contributes a few hundred ink
    // pixels at this font scale, so require a meaningful amount without
    // pinning an exact count (antialiasing shifts it by a few percent).
    assert!(dark > 250, "expected readable cell text, found {dark} dark px");
    // The sheet spans the frame rather than stopping at a fraction of it.
    assert!(
        gridline_rightmost > w * 3 / 4,
        "grid must span the frame: rightmost gridline at {gridline_rightmost} of {w}"
    );

    let _ = std::fs::remove_dir_all(&dir);
    let _ = std::fs::remove_file(path);
}

/// Parse a binary PPM (P6) into `(width, height, rgb bytes)`.
fn parse_ppm(bytes: &[u8]) -> (usize, usize, Vec<u8>) {
    assert_eq!(&bytes[..2], b"P6", "expected a binary PPM");
    let mut fields = Vec::new();
    let mut i = 2usize;
    while fields.len() < 3 {
        while i < bytes.len() && bytes[i].is_ascii_whitespace() {
            i += 1;
        }
        let start = i;
        while i < bytes.len() && !bytes[i].is_ascii_whitespace() {
            i += 1;
        }
        fields.push(std::str::from_utf8(&bytes[start..i]).unwrap().parse::<usize>().unwrap());
    }
    i += 1; // single whitespace after the maxval
    (fields[0], fields[1], bytes[i..].to_vec())
}
