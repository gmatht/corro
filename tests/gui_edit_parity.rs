//! Live GUI parity for type-first editing: Right,Right,A,Enter.
//!
//! The ratatui reference commits exactly "A" to C1 and leaves the cursor on
//! C2. This test drives the real GTK GUI and asserts on the committed
//! `.corro` file (no pixel matching, no fixed sleeps for the verdict — the
//! file is polled with a deadline).
//!
//! Requires: Linux, an X server, xdotool. Run with:
//!   xvfb-run -a cargo test --features gtk --test gui_edit_parity
#![cfg(all(target_os = "linux", feature = "gui"))]

use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::time::{Duration, Instant};

static COUNTER: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

/// Live GUI tests share one X server and global input focus: two corro
/// windows driven in parallel steal focus/keys from each other (flaky empty
/// commits). Serialize the tests in this binary; the verdict itself is still
/// event-based (file polled with a deadline), never a fixed sleep.
static GUI_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

fn xdotool(args: &[&str]) -> String {
    String::from_utf8_lossy(
        &Command::new("xdotool").args(args).output().expect("xdotool failed").stdout,
    )
    .to_string()
}

fn find_corro_window(child_pid: u32, deadline: Instant) -> String {
    // Match by PID, not by name: tests run in parallel and several corro
    // windows can exist at once; name matching would drive the wrong app.
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

fn wait_file_lines(path: &PathBuf, deadline: Instant) -> Vec<String> {
    loop {
        if let Ok(content) = std::fs::read_to_string(path) {
            let lines: Vec<String> = content.lines().map(|l| l.to_string()).collect();
            if lines.iter().any(|l| l.starts_with("SET ")) {
                return lines;
            }
        }
        if Instant::now() > deadline {
            panic!(
                "timed out waiting for commit in {}\ncontent: {:?}",
                path.display(),
                std::fs::read_to_string(path).unwrap_or_default()
            );
        }
        std::thread::sleep(Duration::from_millis(100));
    }
}

fn spawn_gui(path: &PathBuf) -> Child {
    let bin = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("target/debug/corro");
    Command::new(&bin)
        .arg("--gui")
        .arg(path)
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .expect("spawn corro --gui")
}

/// Right,Right,A,Enter must commit exactly "A" to C1 (cursor was on A1).
/// Regression: the GUI backend dropped both arrows (cursor never grew past
/// A1 — move just clamped) and doubled the typed char, committing "AA" to A1.
#[test]
fn gui_right_right_a_enter_commits_c1_single_a() {
    // Tolerate poisoning: a failed sibling must not mask this test's verdict.
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-gui-parity-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);

    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "--window", &wid, "Right"]);
    std::thread::sleep(Duration::from_millis(300));
    xdotool(&["key", "--window", &wid, "Right"]);
    std::thread::sleep(Duration::from_millis(300));
    xdotool(&["type", "--window", &wid, "A"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "Return"]);
    let lines = wait_file_lines(&path, Instant::now() + Duration::from_secs(10));
    let _ = child.kill();
    let _ = child.wait();

    assert!(
        lines.iter().any(|l| l == "SET C1 A"),
        "expected exactly `SET C1 A`, got lines: {lines:?}"
    );
    assert!(
        !lines.iter().any(|l| l.starts_with("SET A1")),
        "must not commit to A1 (arrows must move first), got: {lines:?}"
    );
    assert!(
        !lines.iter().any(|l| l.contains("AA")),
        "typed char must not double, got: {lines:?}"
    );
}

/// Type HELLO, Backspace, Enter in A1: repeated chars must all land (no
/// press/release-dedup drops) and Backspace must pop exactly once.
/// Regression: the GUI backend lost the second L and double-popped,
/// committing "HLO" instead of "HELL".
#[test]
fn gui_hello_backspace_enter_commits_hell() {
    // Tolerate poisoning: a failed sibling must not mask this test's verdict.
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-gui-hello-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);

    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["type", "--window", &wid, "HELLO"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "BackSpace"]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "--window", &wid, "Return"]);
    let lines = wait_file_lines(&path, Instant::now() + Duration::from_secs(10));
    let _ = child.kill();
    let _ = child.wait();

    assert!(
        lines.iter().any(|l| l == "SET A1 HELL"),
        "expected exactly `SET A1 HELL`, got lines: {lines:?}\n--- keylog tail ---\n{}",
        keylog_tail()
    );
}

/// Last lines of the GUI keylog (/tmp/corro_keylog.txt) for failure diagnosis.
fn keylog_tail() -> String {
    let content = std::fs::read_to_string("/tmp/corro_keylog.txt").unwrap_or_default();
    content.lines().rev().take(40).collect::<Vec<_>>().into_iter().rev()
        .collect::<Vec<_>>().join("\n")
}
