//! Live pancurses parity for the overwrite confirm dialog.
//!
//! Typed TUI paths have no native file dialog, so Save As / exports onto an
//! existing file must ask first (ratatui reference: `Mode::ConfirmOverwrite`;
//! GUI backends: native chooser prompt). This drives the real `--pancurses`
//! binary in tmux: File → Save as, type an existing path, Enter — the
//! `Overwrite …? (y/n)` prompt must appear and the file must stay intact;
//! `n` cancels without writing, `y` replaces wholesale. A missing path
//! saves with no prompt.
//!
//! Appearance/disappearance are events (pane-text waits); keys settle before
//! the verdict. No fixed sleeps decide anything.
//!
//! Requires: Linux, tmux, `pancurses` + `ratatui` features. Run with:
//!   cargo test --features pancurses --test pnc_overwrite_confirm
#![cfg(all(target_os = "linux", feature = "pancurses", feature = "ratatui"))]

use std::path::PathBuf;
use std::time::Duration;

static COUNTER: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

mod tmux {
    use std::process::Command;
    pub fn new_session(session: &str, command: &str) {
        Command::new("tmux")
            .args(["new-session", "-d", "-s", session, "-x", "120", "-y", "40", command])
            .status()
            .ok();
    }
    pub fn send_keys(session: &str, key: &str) {
        Command::new("tmux").args(["send-keys", "-t", session, key]).status().ok();
    }
    pub fn send_text(session: &str, text: &str) {
        Command::new("tmux")
            .args(["send-keys", "-t", session, "-l", text])
            .status()
            .ok();
    }
    pub fn capture_pane(session: &str) -> String {
        let o = Command::new("tmux")
            .args(["capture-pane", "-t", session, "-p", "-S", "-200"])
            .output()
            .unwrap();
        String::from_utf8_lossy(&o.stdout).to_string()
    }
    pub fn kill_session(session: &str) {
        Command::new("tmux").args(["kill-session", "-t", session]).output().ok();
    }
}

/// Poll the pane until it contains `needle` (the app actually rendered it).
fn wait_for_text(session: &str, needle: &str) {
    for _ in 0..50 {
        if tmux::capture_pane(session).contains(needle) {
            return;
        }
        std::thread::sleep(Duration::from_millis(200));
    }
    panic!("timed out waiting for pane to contain {needle:?}");
}

/// Poll until `needle` is gone from the pane (prompt closed).
fn wait_for_absence(session: &str, needle: &str) {
    for _ in 0..50 {
        if !tmux::capture_pane(session).contains(needle) {
            return;
        }
        std::thread::sleep(Duration::from_millis(200));
    }
    panic!("timed out waiting for pane to lose {needle:?}");
}

/// Send one key, then wait for the frame it produces to settle (two
/// consecutive identical captures).
fn send_settled(session: &str, key: &str) {
    tmux::send_keys(session, key);
    let mut prev = tmux::capture_pane(session);
    for _ in 0..40 {
        std::thread::sleep(Duration::from_millis(30));
        let cur = tmux::capture_pane(session);
        if cur == prev {
            return;
        }
        prev = cur;
    }
}

fn spawn_session() -> (String, PathBuf, PathBuf) {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let tag = format!("pnc-ow-{}-{id}", std::process::id());
    let work = PathBuf::from(format!("{tag}.corro"));
    let target = std::env::temp_dir().join(format!("{tag}-target.corro"));
    std::fs::write(&work, "CORRO_LOG 1\nSET $1:A1 SESSIONMARK\n").unwrap();
    std::fs::write(&target, "OLD-CONTENT\n").unwrap();
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    tmux::new_session(&tag, &format!("{bin} --pancurses {}; sleep 2", work.display()));
    wait_for_text(&tag, "[File]");
    (tag, work, target)
}

/// Open File → Save as (third item; New sits at 0) and wait for the path prompt.
fn open_save_as(session: &str) {
    send_settled(session, "M-f");
    wait_for_text(session, "Save as");
    send_settled(session, "Down");
    send_settled(session, "Down");
    send_settled(session, "Enter");
    wait_for_text(session, "Save as: ");
}

/// Existing target: prompt appears, file intact; `n` cancels cleanly.
#[test]
fn pnc_saveas_existing_prompts_and_n_cancels() {
    let (session, work, target) = spawn_session();
    open_save_as(&session);
    tmux::send_text(&session, &target.to_string_lossy());
    send_settled(&session, "Enter");
    // Structural: the confirm prompt names the action and the file.
    wait_for_text(&session, "Overwrite");
    wait_for_text(&session, "(y/n)");
    let pane = tmux::capture_pane(&session);
    assert!(
        pane.contains("target.corro") || pane.contains(target.file_name().unwrap().to_string_lossy().as_ref()),
        "confirm prompt must name the file, got:\n{pane}"
    );
    assert_eq!(
        std::fs::read_to_string(&target).unwrap(),
        "OLD-CONTENT\n",
        "showing the prompt must not touch the file"
    );
    tmux::send_text(&session, "n");
    send_settled(&session, "Enter");
    wait_for_absence(&session, "(y/n)");
    assert_eq!(
        std::fs::read_to_string(&target).unwrap(),
        "OLD-CONTENT\n",
        "cancelling must not touch the file"
    );
    tmux::kill_session(&session);
    std::fs::remove_file(&work).ok();
    std::fs::remove_file(&target).ok();
}

/// Existing target: `y` replaces wholesale with the session content.
#[test]
fn pnc_saveas_existing_y_replaces() {
    let (session, work, target) = spawn_session();
    open_save_as(&session);
    tmux::send_text(&session, &target.to_string_lossy());
    send_settled(&session, "Enter");
    wait_for_text(&session, "(y/n)");
    tmux::send_text(&session, "y");
    send_settled(&session, "Enter");
    wait_for_absence(&session, "(y/n)");
    let after = std::fs::read_to_string(&target).unwrap();
    assert!(
        after.contains("SESSIONMARK") && !after.contains("OLD-CONTENT"),
        "confirming must replace wholesale, got:\n{after}"
    );
    tmux::kill_session(&session);
    std::fs::remove_file(&work).ok();
    std::fs::remove_file(&target).ok();
}

/// Missing target: saves with no prompt.
#[test]
fn pnc_saveas_missing_saves_without_prompt() {
    let (session, work, target) = spawn_session();
    std::fs::remove_file(&target).unwrap();
    open_save_as(&session);
    tmux::send_text(&session, &target.to_string_lossy());
    send_settled(&session, "Enter");
    std::thread::sleep(Duration::from_millis(1500));
    let pane = tmux::capture_pane(&session);
    assert!(
        !pane.contains("(y/n)"),
        "missing target must not prompt, got:\n{pane}"
    );
    assert!(
        target.is_file(),
        "missing target must be written directly"
    );
    tmux::kill_session(&session);
    std::fs::remove_file(&work).ok();
    std::fs::remove_file(&target).ok();
}
