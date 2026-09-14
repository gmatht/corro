//! Live GTK parity for the Insert → Special Char picker dialog.
//!
//! The ratatui reference is the oracle: opening the picker, pressing Down*n
//! and Enter splices `CHOICES[min(n,9)]` into the edit, and a second Enter
//! commits it (`SET A1 Ω` for n=2). These tests drive the real GTK GUI
//! through the same gesture and assert on the committed `.corro` file:
//! dialog appearance is an event (a new "Insert special char" window under
//! the app PID), verdicts are file polls with deadlines — never fixed
//! sleeps for the verdict, no pixel matching.
//!
//! Phase A (Down Down + Insert + Enter → `SET A1 Ω`) proves navigation
//! actually selects: if Downs were swallowed, A1 would commit `∞` (the
//! first item) instead. Phase B (no Downs → `SET A2 ∞`) proves Enter
//! confirms the initial selection. The Esc test proves cancel commits
//! nothing.
//!
//! Requires: Linux, an X server, xdotool. Run with:
//!   xvfb-run -a cargo test --features gtk --test gui_special_picker
#![cfg(all(target_os = "linux", feature = "gui"))]

use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::time::{Duration, Instant};

static COUNTER: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

/// Live GUI tests share one X server and global input focus: serialize them.
/// Verdicts stay event-based (window appearance + file polls with deadlines).
static GUI_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

fn xdotool(args: &[&str]) -> String {
    String::from_utf8_lossy(
        &Command::new("xdotool").args(args).output().expect("xdotool failed").stdout,
    )
    .to_string()
}

fn find_corro_window(child_pid: u32, deadline: Instant) -> String {
    let pid = child_pid.to_string();
    loop {
        let out = xdotool(&["search", "--pid", &pid]);
        for id in out.split_whitespace() {
            let geo = xdotool(&["getwindowgeometry", "--shell", id]);
            let wide = geo.lines().any(|l| {
                l.strip_prefix("WIDTH=")
                    .and_then(|v| v.trim().parse::<i32>().ok())
                    .map(|w| w > 100)
                    .unwrap_or(false)
            });
            if wide {
                return id.to_string();
            }
        }
        if Instant::now() > deadline {
            panic!("timed out waiting for corro window");
        }
        std::thread::sleep(Duration::from_millis(300));
    }
}

/// Child process handle that kills (and reaps) the app on drop, including
/// on test panic. Without this, a panicking test leaks its window, which
/// steals X focus and blanks/keys later tests (cascading flakes).
struct KillOnDrop(Child);
impl std::ops::Deref for KillOnDrop {
    type Target = Child;
    fn deref(&self) -> &Child { &self.0 }
}
impl std::ops::DerefMut for KillOnDrop {
    fn deref_mut(&mut self) -> &mut Child { &mut self.0 }
}
impl Drop for KillOnDrop {
    fn drop(&mut self) {
        let _ = self.0.kill();
        let _ = self.0.wait();
    }
}

fn spawn_gui(path: &PathBuf) -> KillOnDrop {
    let bin = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("target/debug/corro");
    KillOnDrop(
        Command::new(&bin)
            .arg("--gui")
            .arg(path)
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .spawn()
            .expect("spawn corro --gui"),
    )
}

fn window_name(id: &str) -> String {
    xdotool(&["getwindowname", id]).trim().to_string()
}

/// Poll with a deadline for the "Insert special char" dialog window to
/// appear (or disappear) under the app PID. Appearance/disappearance is the
/// event — never a fixed sleep.
fn wait_dialog(pid: u32, deadline: Instant, want_present: bool) -> String {
    loop {
        let mut found = String::new();
        for id in xdotool(&["search", "--pid", &pid.to_string()]).split_whitespace() {
            if window_name(id) == "Insert special char" {
                found = id.to_string();
                break;
            }
        }
        if found.is_empty() != want_present {
            return found;
        }
        if Instant::now() > deadline {
            panic!(
                "timed out waiting for picker dialog to {}",
                if want_present { "appear" } else { "disappear" }
            );
        }
        std::thread::sleep(Duration::from_millis(150));
    }
}

/// Poll with a deadline for an exact committed line. Returns all lines.
fn wait_file_line(path: &PathBuf, want: &str, deadline: Instant) -> Vec<String> {
    loop {
        if let Ok(content) = std::fs::read_to_string(path) {
            let lines: Vec<String> = content.lines().map(|l| l.to_string()).collect();
            if lines.iter().any(|l| l == want) {
                return lines;
            }
        }
        if Instant::now() > deadline {
            panic!(
                "timed out waiting for {want:?} in {}\ncontent: {:?}",
                path.display(),
                std::fs::read_to_string(path).unwrap_or_default()
            );
        }
        std::thread::sleep(Duration::from_millis(100));
    }
}

fn file_lines(path: &PathBuf) -> Vec<String> {
    std::fs::read_to_string(path)
        .unwrap_or_default()
        .lines()
        .map(|l| l.to_string())
        .collect()
}

/// Open the picker dialog via Alt+I, s and return the dialog window id.
fn open_picker(child_pid: u32, wid: &str) -> String {
    xdotool(&["key", "--window", wid, "alt+i"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", wid, "s"]);
    let dlg = wait_dialog(child_pid, Instant::now() + Duration::from_secs(8), true);
    xdotool(&["windowactivate", "--sync", &dlg]);
    std::thread::sleep(Duration::from_millis(300));
    dlg
}

/// Down*n, Insert (Return), wait for the dialog to close, then commit the
/// staged edit with Return on the main window.
fn pick_with_downs(wid: &str, dlg: &str, child_pid: u32, downs: usize) {
    for _ in 0..downs {
        xdotool(&["key", "--window", dlg, "Down"]);
        std::thread::sleep(Duration::from_millis(150));
    }
    xdotool(&["key", "--window", dlg, "Return"]);
    wait_dialog(child_pid, Instant::now() + Duration::from_secs(8), false);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "--window", wid, "Return"]);
}

/// Down*2 + Insert + Enter must stage and commit Ω to A1 (proves the Downs
/// navigated: a swallowed navigation would commit ∞), then a no-Down pick
/// must commit ∞ to A2 (proves Enter confirms the initial selection).
#[test]
fn gui_special_picker_down2_enter_commits_omega() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-guipick-{}-{id}.corro", std::process::id()));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, "CORRO_LOG 1\n").expect("write fixture");

    let mut child = spawn_gui(&path);
    let pid = child.id();
    let wid = find_corro_window(pid, Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));

    // Phase A: Down*2 → Ω into A1.
    let dlg = open_picker(pid, &wid);
    pick_with_downs(&wid, &dlg, pid, 2);
    let lines = wait_file_line(&path, "SET A1 Ω", Instant::now() + Duration::from_secs(10));
    assert!(
        !lines.iter().any(|l| l == "SET A1 ∞"),
        "Down*2 must navigate past the first item (∞), got: {lines:?}"
    );

    // Phase B: no Downs → ∞ into A2 (cursor moved down after the commit).
    let dlg = open_picker(pid, &wid);
    pick_with_downs(&wid, &dlg, pid, 0);
    let lines = wait_file_line(&path, "SET A2 ∞", Instant::now() + Duration::from_secs(10));
    let _ = child.kill();
    let _ = child.wait();

    assert!(
        lines.iter().any(|l| l == "SET A1 Ω"),
        "phase A must have committed Ω to A1, got: {lines:?}"
    );
    assert!(
        lines.iter().any(|l| l == "SET A2 ∞"),
        "phase B must commit ∞ to A2, got: {lines:?}"
    );
}

/// Esc in the picker dialog must close it and commit nothing.
#[test]
fn gui_special_picker_esc_cancels_without_commit() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-guipickesc-{}-{id}.corro", std::process::id()));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, "CORRO_LOG 1\n").expect("write fixture");

    let mut child = spawn_gui(&path);
    let pid = child.id();
    let wid = find_corro_window(pid, Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));

    let dlg = open_picker(pid, &wid);
    xdotool(&["key", "--window", &dlg, "Down"]);
    std::thread::sleep(Duration::from_millis(200));
    xdotool(&["key", "--window", &dlg, "Escape"]);
    wait_dialog(pid, Instant::now() + Duration::from_secs(8), false);
    std::thread::sleep(Duration::from_millis(500));
    let _ = child.kill();
    let _ = child.wait();

    let lines = file_lines(&path);
    assert!(
        !lines.iter().any(|l| l.starts_with("SET ")),
        "Esc must commit nothing, got: {lines:?}"
    );
}
