//! Live pancurses parity for the Insert → Special Char picker.
//!
//! The ratatui reference is the oracle: opening the picker, pressing Down*n
//! and Enter splices `CHOICES[min(n,9)]` into the edit, and a second Enter
//! commits it. This drives the real `--pancurses` binary in tmux through the
//! same gesture: Alt+I, arrows to Special Char, Enter (picker popup with a
//! `▸`-highlighted row), Down*n, Enter (splice), Enter (commit) — and
//! asserts on the committed `.corro` file plus the popup structure (border,
//! highlight marker, title).
//!
//! Appearance/disappearance are events (pane-text waits); keys settle before
//! the verdict. No fixed sleeps decide anything.
//!
//! Requires: Linux, tmux, `pancurses` + `ratatui` features. Run with:
//!   cargo test --features pancurses --test pnc_special_picker
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

/// Poll until `needle` is gone from the pane (popup closed).
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

/// Poll the `.corro` log for an exact committed line.
fn wait_file_line(path: &PathBuf, want: &str) {
    for _ in 0..100 {
        if let Ok(content) = std::fs::read_to_string(path) {
            if content.lines().any(|l| l == want) {
                return;
            }
        }
        std::thread::sleep(Duration::from_millis(100));
    }
    panic!(
        "timed out waiting for {want:?} in {}\ncontent: {:?}",
        path.display(),
        std::fs::read_to_string(path).unwrap_or_default()
    );
}

fn spawn_session() -> (String, PathBuf) {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("pnc-pick-{}-{id}", std::process::id());
    let tmp = std::env::temp_dir().join(format!("{session}.corro"));
    std::fs::write(&tmp, "CORRO_LOG 1\n").unwrap();
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    tmux::new_session(&session, &format!("{bin} --pancurses {}; sleep 2", tmp.display()));
    wait_for_text(&session, "[File]");
    (session, tmp)
}

/// Open the Insert menu and activate Special Char (index 4) to open the
/// picker popup. Returns after the picker rows are visible.
fn open_picker(session: &str) {
    send_settled(session, "M-i");
    wait_for_text(session, "┌Insert");
    // Settle once more so the popup is fully interactive before navigating.
    send_settled(session, "Escape");
    send_settled(session, "M-i");
    wait_for_text(session, "┌Insert");
    for _ in 0..4 {
        send_settled(session, "Down");
    }
    send_settled(session, "Enter");
    wait_for_text(session, "1: ∞");
}

/// Down*2 + Enter must splice Ω (structural: popup border + `▸` highlight
/// on `3: Ω`), then Enter must commit `SET A1 Ω` (proves the Downs
/// navigated rather than committing the first item).
#[test]
fn pnc_special_picker_down2_enter_commits_omega() {
    let (session, tmp) = spawn_session();
    open_picker(&session);
    let pane = tmux::capture_pane(&session);
    // Structural popup assertions: box border, title, first-row highlight.
    for border in ["┌", "┐", "└", "┘", "─", "│"] {
        assert!(pane.contains(border), "picker popup must draw a box border, got:\n{pane}");
    }
    assert!(pane.contains("Special Char"), "picker popup must be titled, got:\n{pane}");
    assert!(
        pane.lines().any(|l| l.contains('▸') && l.contains("1: ∞")),
        "picker must highlight the first row, got:\n{pane}"
    );
    send_settled(&session, "Down");
    send_settled(&session, "Down");
    let pane = tmux::capture_pane(&session);
    assert!(
        pane.lines().any(|l| l.contains('▸') && l.contains("3: Ω")),
        "Down*2 must highlight `3: Ω`, got:\n{pane}"
    );
    send_settled(&session, "Enter");
    wait_for_absence(&session, "1: ∞");
    // Second Enter commits the staged edit.
    send_settled(&session, "Enter");
    wait_file_line(&tmp, "SET A1 Ω");
    tmux::kill_session(&session);
    // No first-item commit may sneak in alongside.
    let content = std::fs::read_to_string(&tmp).unwrap_or_default();
    assert!(
        !content.lines().any(|l| l == "SET A1 ∞"),
        "Down*2 must not commit the first item, got:\n{content}"
    );
}

/// Esc in the picker must close it and commit nothing.
#[test]
fn pnc_special_picker_esc_cancels_without_commit() {
    let (session, tmp) = spawn_session();
    open_picker(&session);
    send_settled(&session, "Down");
    send_settled(&session, "Escape");
    wait_for_absence(&session, "1: ∞");
    std::thread::sleep(Duration::from_millis(500));
    tmux::kill_session(&session);
    let content = std::fs::read_to_string(&tmp).unwrap_or_default();
    assert!(
        !content.lines().any(|l| l.starts_with("SET ")),
        "Esc must commit nothing, got:\n{content}"
    );
}
