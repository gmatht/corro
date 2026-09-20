//! `--movie` on the GUI backends: the demo mode is no longer terminal-only.
//!
//! The ratatui TUI has its own replayer that drives a live terminal. The GUI
//! backends share one `App` over the same workbook/op core, so the replay
//! itself (`corro::gui::movie`) and the frame painting
//! (`corro::gui::gui_movie`) are backend-independent and testable without a
//! display server — which is also what makes an unattended video capture
//! possible (`corro --gui --movie --movie-frames DIR FILE.corro`).
#![cfg(feature = "gui")]

use corro::gui::gui_movie::{run_replay, FramePainter, MovieFrameView, NullPainter};
use corro::gui::movie::{GuiMovie, GuiMovieOptions};
use corro::grid::CellAddr;
use std::path::{Path, PathBuf};

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

/// A painter that records the captions it was asked to paint, so the test can
/// assert the replay produced the frames a viewer would see.
#[derive(Default)]
struct CaptionLog {
    captions: Vec<String>,
    menus: Vec<String>,
}

impl FramePainter for CaptionLog {
    fn size(&self) -> (i32, i32) {
        (320, 200)
    }
    fn is_live(&self) -> bool {
        false
    }
    fn paint(
        &mut self,
        _app: &corro::gui::App,
        frame: &MovieFrameView,
        n: usize,
    ) -> Result<(), String> {
        self.captions.push(frame.status.clone());
        if let Some((section, item)) = frame.menu.as_ref() {
            self.menus.push(format!("{section} ▸ {item}"));
        }
        Ok(())
    }
}

#[test]
fn gui_movie_replays_every_step_into_the_workbook() {
    let path = movie_file("steps", "SET $1:A1 1\nSET $1:A2 2\nSET $1:B1 4\n");
    let mut movie = GuiMovie::new(&path).expect("parse movie");
    assert_eq!(movie.len(), 3);

    let mut app = corro::gui::App::new_with_paths(vec![path.clone()]);
    let mut painter = CaptionLog::default();
    let frames = run_replay(&mut app, &mut movie, fast(), &mut painter).expect("replay");

    // One typing frame per character, one settle frame per step, and the
    // closing "movie complete" frame — the settle frame carries the same
    // caption as the last typed frame, so count captions rather than frames.
    assert!(frames >= 4, "captions: {:?}", painter.captions);
    assert_eq!(
        painter
            .captions
            .iter()
            .filter(|c| c.as_str() == "A1 = 1")
            .count(),
        2,
        "one typing frame plus one settle frame for A1: {:?}",
        painter.captions
    );
    assert!(painter.captions.last().unwrap().starts_with("Movie complete"));

    let grid = &app.core.workbook.active_sheet().grid;
    assert_eq!(grid.get(&CellAddr::Main { row: 0, col: 0 }), Some("1".to_string()));
    assert_eq!(grid.get(&CellAddr::Main { row: 1, col: 0 }), Some("2".to_string()));
    assert_eq!(grid.get(&CellAddr::Main { row: 0, col: 1 }), Some("4".to_string()));

    // The cursor follows the replay instead of sitting at A1.
    let expected = corro::gui::movie::cursor_of(&CellAddr::Main { row: 0, col: 1 }, &app.core.workbook);
    assert_eq!(app.core.cursor.row, expected.row);
    assert_eq!(app.core.cursor.col, expected.col);

    let _ = std::fs::remove_file(path);
}

#[test]
fn gui_movie_animates_typed_values_character_by_character() {
    let path = movie_file("typing", "SET $1:A1 hello\n");
    let mut movie = GuiMovie::new(&path).expect("parse movie");
    let mut app = corro::gui::App::new_with_paths(vec![path.clone()]);
    let mut painter = CaptionLog::default();
    run_replay(&mut app, &mut movie, fast(), &mut painter).expect("replay");

    // "hello" is five characters: the captions must walk the partial text so a
    // recording looks like someone typing, then land on the committed value.
    for partial in ["h", "he", "hel", "hell", "hello"] {
        assert!(
            painter.captions.iter().any(|c| c.ends_with(&format!("A1 = {partial}"))),
            "missing typing frame for {partial:?}: {:?}",
            painter.captions
        );
    }
    let _ = std::fs::remove_file(path);
}

#[test]
fn gui_movie_flashes_the_menu_a_step_would_have_used() {
    // A range of cleared cells is the Edit ▸ Cut path; the replay must show
    // that menu before the change lands, or the step is unreadable.
    let path = movie_file("menu", "SET $1:A1 x\nSET $1:B1 y\nFILL A1= B1=\n");
    let mut movie = GuiMovie::new(&path).expect("parse movie");
    let mut app = corro::gui::App::new_with_paths(vec![path.clone()]);
    let mut painter = CaptionLog::default();
    run_replay(&mut app, &mut movie, fast(), &mut painter).expect("replay");

    assert!(
        painter.menus.iter().any(|m| m == "Edit ▸ Cut"),
        "expected an Edit ▸ Cut flash, got {:?}",
        painter.menus
    );
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
fn gui_movie_frames_are_written_when_a_directory_is_configured() {
    // The capture path is what makes the video possible: with
    // CORRO_MOVIE_FRAMES set, the GUI movie run must leave one image per
    // painted frame behind.
    let path = movie_file("capture", "SET $1:A1 7\n");
    let dir = std::env::temp_dir().join(format!("corro-movie-frames-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    std::env::set_var("CORRO_MOVIE_FRAMES", &dir);

    let mut app = corro::gui::App::new_with_paths(vec![path.clone()]);
    app.run_movie(fast()).expect("capture run");

    let frames: Vec<PathBuf> = std::fs::read_dir(&dir)
        .unwrap()
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().is_some_and(|e| e == "ppm"))
        .collect();
    assert!(frames.len() >= 2, "expected per-frame images, got {frames:?}");
    let bytes = std::fs::read(&frames[0]).unwrap();
    assert_eq!(&bytes[..2], b"P6", "frames are binary PPMs");

    std::env::remove_var("CORRO_MOVIE_FRAMES");
    let _ = std::fs::remove_dir_all(&dir);
    let _ = std::fs::remove_file(path);
}

#[test]
fn gui_movie_null_painter_still_counts_frames() {
    // `--gui --movie` without a frames directory must still replay (and
    // report progress) so the mode is usable without a display.
    let path = movie_file("null", "SET $1:A1 1\n");
    let mut movie = GuiMovie::new(&path).expect("parse movie");
    let mut app = corro::gui::App::new_with_paths(vec![path.clone()]);
    let mut painter = NullPainter::default();
    let frames = run_replay(&mut app, &mut movie, fast(), &mut painter).expect("replay");
    assert_eq!(painter.painted, frames);
    assert!(frames >= 2);
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
    let _: &Path = Path::new(".");
}
