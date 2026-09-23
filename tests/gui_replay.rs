//! File ▸ Replay must actually reload the bound `.corro` log, and must never
//! dead-end silently when there is nothing to replay.
//!
//! The shared dispatcher (`gui::actions::dispatch_menu_action`) is what every
//! non-ratatui backend runs, so these tests pin the contract for the GUI and
//! pancurses at once:
//!
//!   1. a bound `.corro` file is reloaded from disk (an externally appended
//!      revision shows up), with the cursor reset and the revision-browse
//!      bookkeeping set the way the ratatui reference does;
//!   2. a detached `source_path` is accepted too (`path.or(source_path)`),
//!      so a workbook shown from a movie/revision source can be replayed;
//!   3. with no source at all the status is actionable (names File ▸ Open)
//!      rather than a silent no-op;
//!   4. a missing or non-`.corro` source reports what is wrong.
//!
//! Run with: `cargo test --features gui --test gui_replay`
//! (also compiles under `--features pancurses`).

#![cfg(all(target_os = "linux", any(feature = "gui", feature = "pancurses")))]

use corro::gui::actions::{dispatch_menu_action, MenuDispatch};
use corro::gui::App;

static COUNTER: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

fn temp_corro(lines: &str) -> std::path::PathBuf {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let p = std::env::temp_dir().join(format!("corro-replay-{}-{id}.corro", std::process::id()));
    std::fs::write(&p, format!("CORRO_LOG 1\n{lines}")).unwrap();
    p
}

fn replay(app: &mut App) -> MenuDispatch {
    let mut scope = 0u8;
    let mut clipboard = String::new();
    dispatch_menu_action(app, "replay", &mut scope, &mut clipboard)
}

fn status(d: MenuDispatch) -> String {
    match d {
        MenuDispatch::Status(s) => s,
        other => panic!(
            "replay must dispatch Status, got {}",
            match other {
                MenuDispatch::Prompt(..) => "Prompt",
                MenuDispatch::SpecialPicker => "SpecialPicker",
                MenuDispatch::AggregatePicker => "AggregatePicker",
                MenuDispatch::Edit { .. } => "Edit",
                MenuDispatch::About { .. } => "About",
                MenuDispatch::HelpFull { .. } => "HelpFull",
                MenuDispatch::HelpKeybinds { .. } => "HelpKeybinds",
                MenuDispatch::BalanceBooks => "BalanceBooks",
                MenuDispatch::Status(_) => "Status",
            }
        ),
    }
}

/// The bound file is re-read from disk: a revision appended after the app
/// loaded it appears, the cursor resets to the home cell, and the replay is
/// recorded as a browsable revision set (as the ratatui reference does).
#[test]
fn replay_reloads_the_bound_file_from_disk() {
    let path = temp_corro("SET $1:A1 first\n");
    let mut app = App::new_with_paths(vec![path.clone()]);
    app.load_initial().unwrap();
    assert_eq!(
        app.core.workbook.active_sheet().grid.get(&corro::grid::CellAddr::main(0, 0)),
        Some("first".to_string())
    );

    // A revision lands on disk that the app has not seen.
    std::fs::write(&path, "CORRO_LOG 1\nSET $1:A1 first\nSET $1:A2 added\n").unwrap();
    // Move the cursor away so the reset is observable.
    app.core.cursor = corro::grid::SheetCursor { row: 5, col: 5 };

    let s = status(replay(&mut app));
    assert!(s.starts_with("Replayed "), "status must confirm the replay, got {s:?}");
    assert!(s.contains("@ revision 2"), "status must name the revision, got {s:?}");

    let g = &app.core.workbook.active_sheet().grid;
    assert_eq!(
        g.get(&corro::grid::CellAddr::main(1, 0)),
        Some("added".to_string()),
        "the externally-appended revision must be applied by the replay"
    );
    assert_eq!(
        app.core.cursor,
        corro::grid::SheetCursor {
            row: corro::grid::HEADER_ROWS,
            col: corro::grid::MARGIN_COLS
        },
        "replay resets the cursor to the home cell"
    );
    assert!(app.core.revision_browse, "replay marks the log as a browsable revision set");
    assert_eq!(app.core.revision_browse_limit, 2);
    let _ = std::fs::remove_file(&path);
}

/// A detached `source_path` (no `path`) is replayable, matching the
/// reference's `path.or(source_path)`.
#[test]
fn replay_accepts_a_detached_source_path() {
    let path = temp_corro("SET $1:A1 detached\n");
    let mut app = App::new_with_paths(vec![]);
    app.load_initial().unwrap();
    app.core.source_path = Some(path.clone());
    assert!(app.core.path.is_none(), "setup: no bound path");

    let s = status(replay(&mut app));
    assert!(s.starts_with("Replayed "), "detached source must replay, got {s:?}");
    assert_eq!(
        app.core.workbook.active_sheet().grid.get(&corro::grid::CellAddr::main(0, 0)),
        Some("detached".to_string())
    );
    let _ = std::fs::remove_file(&path);
}

/// With nothing to replay the status must point at File ▸ Open, not read as a
/// silent no-op.
#[test]
fn replay_without_a_file_says_how_to_get_one() {
    let mut app = App::new_with_paths(vec![]);
    app.load_initial().unwrap();
    assert!(app.core.path.is_none() && app.core.source_path.is_none());

    let s = status(replay(&mut app));
    assert!(
        s.contains("Open") && s.contains(".corro"),
        "no-file replay must tell the user to open a .corro file, got {s:?}"
    );
}

/// A path that no longer exists is reported, not silently ignored.
#[test]
fn replay_reports_a_missing_file() {
    let mut app = App::new_with_paths(vec![]);
    app.load_initial().unwrap();
    app.core.path = Some(std::env::temp_dir().join("corro-replay-does-not-exist.corro"));

    let s = status(replay(&mut app));
    assert!(s.contains("not found"), "missing file must be reported, got {s:?}");
}

/// A non-`.corro` file (e.g. a `.tsv` opened for import) is not a replayable
/// log, and saying so beats a confusing parse error.
#[test]
fn replay_rejects_a_non_corro_source() {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let p = std::env::temp_dir().join(format!("corro-replay-{}-{id}.tsv", std::process::id()));
    std::fs::write(&p, "a\tb\n1\t2\n").unwrap();

    let mut app = App::new_with_paths(vec![]);
    app.load_initial().unwrap();
    app.core.path = Some(p.clone());

    let s = status(replay(&mut app));
    assert!(s.contains("not a .corro log"), "non-corro source must be named, got {s:?}");
    let _ = std::fs::remove_file(&p);
}
