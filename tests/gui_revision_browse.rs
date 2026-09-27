//! File ▸ Replay's revision-browse mode: Left/Right step through the current
//! `.corro` log's revisions, matching the ratatui reference.
//!
//! The reference implements this as `Mode::RevisionBrowse` in `src/ui/mod.rs`
//! (the arrow arm is at `src/ui/mod.rs:10714`, the reload at `:3486`, and its
//! own parity test at `:20089`). The GUI reaches the same behaviour through
//! `dispatch_menu_action` + `CoreApp::step_revision_*`, so these tests pin the
//! *contract* for both GUI backends at once — and pin the same numbers the
//! reference's test pins, so a drift in either implementation is caught.
//!
//! The sequence is the interesting part, because every boundary is a place
//! where an off-by-one is invisible at a glance:
//!
//!   * Left at revision 1 is a no-op (guard `limit > 1`), never a wrap;
//!   * Right past the newest revision is a no-op (the re-replay clamps);
//!   * the cells change with the revision, not just the counter.
//!
//! Run with: `cargo test --features gui --test gui_revision_browse`
//! (also compiles under `--features pancurses`).

#![cfg(all(target_os = "linux", any(feature = "gui", feature = "pancurses")))]

use corro::core::state::RevisionStep;
use corro::gui::actions::{dispatch_menu_action, MenuDispatch};
use corro::gui::App;

static COUNTER: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

fn temp_corro(lines: &str) -> std::path::PathBuf {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let p = std::env::temp_dir().join(format!("corro-revbrowse-{}-{id}.corro", std::process::id()));
    // Remove any leftover from an earlier run that happened to reuse this
    // pid+id slot: `replay_inside_browse_mode_rereads_the_log` appends a
    // revision to the file on disk, so starting from a stale copy would make
    // that test's revision count depend on run order.
    let _ = std::fs::remove_file(&p);
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
        _ => panic!("replay must dispatch Status"),
    }
}

fn cell(app: &App, row: u32, col: u32) -> Option<String> {
    app.core
        .workbook
        .active_sheet()
        .grid
        .get(&corro::grid::CellAddr::main(row, col))
}

/// A three-revision log: A1, then B1, then A1 overwritten.
fn three_revision_app() -> (App, std::path::PathBuf) {
    let path = temp_corro("SET $1:A1 one\nSET $1:B1 two\nSET $1:A1 three\n");
    let app = App::new_with_paths(vec![path.clone()]);
    (app, path)
}

/// File▸Replay leaves the app browsing, and the initial replay says
/// "Replayed" (the reference's `replay_status("Replayed", ...)`); only
/// stepping switches the wording to "Browsing".
#[test]
fn replay_enters_revision_browse_at_the_newest_revision() {
    let (mut app, path) = three_revision_app();
    let s = status(replay(&mut app));

    assert!(s.starts_with("Replayed "), "initial replay says Replayed, got {s:?}");
    assert!(s.contains("@ revision 3"), "got {s:?}");
    assert!(app.core.revision_browse, "replay marks the log browsable");
    assert_eq!(app.core.revision_browse_limit, 3);
    assert!(app.core.path.is_none(), "a browsed revision set is detached");
    // The newest revision's cells are what a full replay produces.
    assert_eq!(cell(&app, 0, 0).as_deref(), Some("three"));
    assert_eq!(cell(&app, 0, 1).as_deref(), Some("two"));
    let _ = std::fs::remove_file(&path);
}

/// The core requirement: Left shows the previous revision's cells, Right
/// moves back up, and both ends clamp. This is the same sequence the
/// reference's own test walks for its two-revision log.
#[test]
fn left_and_right_step_through_the_logs_cells() {
    let (mut app, path) = three_revision_app();
    let _ = status(replay(&mut app));

    // Left -> revision 2: A1 is "one" again, B1 still set.
    assert_eq!(app.core.step_revision_back().unwrap(), RevisionStep::Moved);
    assert_eq!(app.core.revision_browse_limit, 2);
    assert_eq!(cell(&app, 0, 0).as_deref(), Some("one"));
    assert_eq!(cell(&app, 0, 1).as_deref(), Some("two"));
    assert!(app.core.status.contains("@ revision 2"), "{}", app.core.status);

    // Left -> revision 1: B1 has not happened yet.
    assert_eq!(app.core.step_revision_back().unwrap(), RevisionStep::Moved);
    assert_eq!(app.core.revision_browse_limit, 1);
    assert_eq!(cell(&app, 0, 0).as_deref(), Some("one"));
    assert_eq!(cell(&app, 0, 1), None);

    // Boundary: Left at the oldest revision does nothing.
    assert_eq!(app.core.step_revision_back().unwrap(), RevisionStep::AtOldest);
    assert_eq!(app.core.revision_browse_limit, 1);
    assert_eq!(cell(&app, 0, 0).as_deref(), Some("one"));

    // Right walks back up.
    assert_eq!(app.core.step_revision_forward().unwrap(), RevisionStep::Moved);
    assert_eq!(app.core.revision_browse_limit, 2);
    assert_eq!(cell(&app, 0, 0).as_deref(), Some("one"));
    assert_eq!(cell(&app, 0, 1).as_deref(), Some("two"));
    assert_eq!(app.core.step_revision_forward().unwrap(), RevisionStep::Moved);
    assert_eq!(app.core.revision_browse_limit, 3);
    assert_eq!(cell(&app, 0, 0).as_deref(), Some("three"));

    // Boundary: Right past the newest revision is a no-op, not a phantom
    // revision and not an empty sheet.
    assert_eq!(app.core.step_revision_forward().unwrap(), RevisionStep::AtNewest);
    assert_eq!(app.core.revision_browse_limit, 3);
    assert_eq!(cell(&app, 0, 0).as_deref(), Some("three"));

    let _ = std::fs::remove_file(&path);
}

/// Stepping backwards then forwards returns to the same cells: the state is
/// a pure function of the revision number, so nothing accumulates.
#[test]
fn stepping_is_reversible() {
    let (mut app, path) = three_revision_app();
    let _ = status(replay(&mut app));
    let before = (cell(&app, 0, 0), cell(&app, 0, 1));

    app.core.step_revision_back().unwrap();
    app.core.step_revision_back().unwrap();
    app.core.step_revision_forward().unwrap();
    app.core.step_revision_forward().unwrap();

    assert_eq!(app.core.revision_browse_limit, 3);
    assert_eq!((cell(&app, 0, 0), cell(&app, 0, 1)), before);
    let _ = std::fs::remove_file(&path);
}

/// Browsing is read-only: the log on disk is untouched no matter how far the
/// user steps, and no commit path is armed.
#[test]
fn browsing_never_writes_to_the_log() {
    let (mut app, path) = three_revision_app();
    let _ = status(replay(&mut app));
    let before = std::fs::read_to_string(&path).unwrap();

    app.core.step_revision_back().unwrap();
    app.core.step_revision_back().unwrap();
    app.core.step_revision_forward().unwrap();

    assert_eq!(std::fs::read_to_string(&path).unwrap(), before);
    assert!(app.core.path.is_none(), "still detached from the commit path");
    let _ = std::fs::remove_file(&path);
}

/// The hint the GUIs show while browsing is the reference's own string, so
/// the available keys are discoverable and the text cannot drift between
/// backends.
#[test]
fn browse_hint_matches_the_reference() {
    assert_eq!(
        corro::core::state::REVISION_BROWSE_HINTS,
        "  left/right·step revisions   Enter·close   Esc·close"
    );
    let (mut app, path) = three_revision_app();
    let _ = status(replay(&mut app));
    assert!(app.core.revision_browse, "the hint only shows while browsing");
    let _ = std::fs::remove_file(&path);
}

/// Replaying again from inside browse mode re-reads the log: a revision
/// appended on disk is picked up, which is the "replay before saving" case
/// from the reference's `path.or(source_path)` handling.
#[test]
fn replay_inside_browse_mode_rereads_the_log() {
    let (mut app, path) = three_revision_app();
    let _ = status(replay(&mut app));
    app.core.step_revision_back().unwrap();
    assert_eq!(app.core.revision_browse_limit, 2);

    // A fourth revision lands while we are browsing an older one.
    std::fs::write(
        &path,
        "CORRO_LOG 1\nSET $1:A1 one\nSET $1:B1 two\nSET $1:A1 three\nSET $1:A2 four\n",
    )
    .unwrap();
    let s = status(replay(&mut app));
    assert!(s.contains("@ revision 4"), "replay must see the new revision, got {s:?}");
    assert_eq!(app.core.revision_browse_limit, 4);
    assert_eq!(cell(&app, 1, 0).as_deref(), Some("four"));
    let _ = std::fs::remove_file(&path);
}

/// A fresh app with no source reports rather than panics: the GUI can reach
/// this if the file is renamed away mid-session.
#[test]
fn stepping_without_a_source_is_reported() {
    let mut app = App::new_with_paths(vec![]);
    assert!(app.core.source_path.is_none(), "setup: no source");
    assert_eq!(app.core.step_revision_back().unwrap(), RevisionStep::NoSource);
    assert_eq!(app.core.step_revision_forward().unwrap(), RevisionStep::NoSource);
}
