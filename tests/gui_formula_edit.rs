//! Clicking into the formula bar must EDIT the formula, not replace it.
//!
//! Regression: clicking the formula entry (which displays the cell's
//! formula) and pressing a key wiped the formula to just that keystroke —
//! the keypress took the canvas-typing path (`start_edit_with`, fresh
//! replace) instead of inserting at the entry cursor. The fix adopts the
//! displayed text into the edit buffer on entry focus and lets the native
//! widget insert at its cursor while focused.
//!
//! Method (no fixed sleeps for the verdict): fixture A1 holds `AB`; click
//! the entry zone; type `X`; press Return; poll the `.corro` log with a
//! deadline for the committed value. Fixed code commits `ABX` (insert at
//! end after a past-the-text click); the bug committed bare `X`.
//!
//! Requires: Linux, an X server, xdotool. Run with:
//!   xvfb-run -a cargo test --features gtk --test gui_formula_edit
#![cfg(all(target_os = "linux", feature = "gui"))]

use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::time::{Duration, Instant};

static COUNTER: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

/// See gui_edit_parity.rs: tests in one binary share an X server and global
/// focus, so serialize them. The verdict itself stays event-based.
static GUI_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

fn xdotool(args: &[&str]) -> String {
    String::from_utf8_lossy(
        &Command::new("xdotool").args(args).output().expect("xdotool failed").stdout,
    )
    .to_string()
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

fn find_corro_window(child_pid: u32, deadline: Instant) -> String {
    // Same as gui_edit_parity: skip tiny InputOnly windows, take the real one.
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

fn win_xy(id: &str) -> (i32, i32) {
    let g = xdotool(&["getwindowgeometry", "--shell", id]);
    let mut x = -1;
    let mut y = -1;
    for l in g.lines() {
        if let Some(v) = l.strip_prefix("X=") {
            x = v.trim().parse().unwrap_or(-1);
        }
        if let Some(v) = l.strip_prefix("Y=") {
            y = v.trim().parse().unwrap_or(-1);
        }
    }
    (x, y)
}

fn screenshot(wid: &str, tag: &str) -> PathBuf {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let xwd = std::env::temp_dir().join(format!("corro-formula-{tag}-{id}.xwd"));
    let png = std::env::temp_dir().join(format!("corro-formula-{tag}-{id}.png"));
    Command::new("xwd")
        .args(["-id", wid, "-out"])
        .arg(&xwd)
        .status()
        .expect("xwd");
    Command::new("convert")
        .arg(&xwd)
        .arg(&png)
        .status()
        .expect("convert xwd->png");
    let _ = std::fs::remove_file(&xwd);
    png
}

/// OCR over the formula row's left (entry) zone, upscaled 2x — same geometry
/// the status-bar tests use (500x35+0+20).
fn ocr_entry_zone(shot: &PathBuf) -> String {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let crop = std::env::temp_dir().join(format!("corro-formula-entry-{id}.png"));
    Command::new("convert")
        .arg(shot)
        .args(["-crop", "500x35+0+20", "+repage", "-colorspace", "Gray", "-resize", "200%"])
        .arg(&crop)
        .status()
        .expect("convert entry crop");
    let out = Command::new("tesseract")
        .arg(&crop)
        .arg("stdout")
        .args(["--psm", "7"])
        .output()
        .expect("tesseract");
    let _ = std::fs::remove_file(&crop);
    String::from_utf8_lossy(&out.stdout).to_string()
}

/// Wait for the first painted frame (deadline-polled appearance event).
/// Fresh Xvfb servers sometimes stall present()'s pump, leaving a blank
/// window for many seconds; a Down/Up poke partway kicks the frame clock
/// (net-zero cursor move). Same pattern as the status-bar suite.
fn await_first_paint(wid: &str) {
    let verdict = Instant::now() + Duration::from_secs(45);
    let mut poked = false;
    loop {
        let shot = screenshot(wid, "firstpaint");
        let text = ocr_entry_zone(&shot);
        let _ = std::fs::remove_file(&shot);
        if text.contains('A') || text.contains("fx") || Instant::now() > verdict {
            assert!(
                text.contains('A') || text.contains("fx"),
                "app never painted (present-pump stall), OCR: {text:?}"
            );
            break;
        }
        if !poked && Instant::now() > verdict - Duration::from_secs(30) {
            poked = true;
            xdotool(&["key", "--window", wid, "Down"]);
            xdotool(&["key", "--window", wid, "Up"]);
        }
        std::thread::sleep(Duration::from_millis(400));
    }
}

fn spawn_gui(path: &PathBuf) -> KillOnDrop {
    // CORRO_TEST_BIN overrides the spawned binary (hermetic against
    // concurrent sessions rebuilding target/debug/corro with different
    // features mid-run); otherwise the usual cargo-built binary.
    let bin = std::env::var("CORRO_TEST_BIN")
        .map(PathBuf::from)
        .unwrap_or_else(|_| PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("target/debug/corro"));
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

/// Click into the formula entry, type two keys, commit with Return, and
/// prove the committed log value kept the original formula text in order
/// (insert) instead of collapsing to the keystrokes (replace — the bug).
/// Two keys also guard append order: the second must land after the first.
#[test]
fn gui_formula_bar_click_type_inserts() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let stem = format!("corro-formulaedit-{}-{}.corro", std::process::id(), id);
    let path = std::env::temp_dir().join(&stem);
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, "CORRO_LOG 1\nSET A1 AB\n").expect("write fixture");

    let bin = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("target/debug/corro");
    let mut child = KillOnDrop(
        Command::new(&bin)
            .arg("--gui")
            .arg(&path)
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .spawn()
            .expect("spawn corro --gui"),
    );
    let pid = child.id();
    let wid = find_corro_window(pid, Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    await_first_paint(&wid);

    // Formula entry zone: full-width row near the top (the status-bar
    // tests crop it as 500x35+0+20). Under Xvfb there is no window manager,
    // so client coords equal root coords: click mid-entry, well past the
    // short `AB` text so the caret lands at the end.
    let (wx, wy) = win_xy(&wid);
    assert!(wx >= 0 && wy >= 0, "no window position for {wid}");
    xdotool(&[
        "mousemove",
        &format!("{}", wx + 300),
        &format!("{}", wy + 37),
        "click",
        "1",
    ]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["type", "XY"]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "Return"]);

    // Ground truth is the log, not pixels: the commit must append
    // `SET A1 ABXY` in order. The bug committed bare keystrokes (`XY`),
    // and a caret-reset regression would show `YABX` or similar.
    let deadline = Instant::now() + Duration::from_secs(8);
    loop {
        let log = std::fs::read_to_string(&path).unwrap_or_default();
        if log.lines().any(|l| l.trim() == "SET A1 ABXY") {
            break;
        }
        if Instant::now() > deadline {
            let _ = child.kill();
            panic!(
                "formula-bar typing did not insert in order (log lacks `SET A1 ABXY`):\n{log}"
            );
        }
        std::thread::sleep(Duration::from_millis(150));
    }
    let _ = std::fs::remove_file(&path);
}

/// Native text selection works in the formula bar: click in, Ctrl+A to
/// select all, type to replace the selection, commit. Proves the entry
/// handles its own select/replace natively (no app interception) — the
/// same machinery that makes mid-text clicks insert mid-text.
#[test]
fn gui_formula_bar_select_all_replace() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path =
        std::env::temp_dir().join(format!("corro-formulasel-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, "CORRO_LOG 1\nSET A1 AB\n").expect("write fixture");

    let mut child = spawn_gui(&path);
    let pid = child.id();
    let wid = find_corro_window(pid, Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    await_first_paint(&wid);
    let (wx, wy) = win_xy(&wid);
    xdotool(&[
        "mousemove",
        &format!("{}", wx + 300),
        &format!("{}", wy + 37),
        "click",
        "1",
    ]);
    std::thread::sleep(Duration::from_millis(400));
    // Ctrl+A must select (native), not type or bubble to the window.
    xdotool(&["key", "ctrl+a"]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["type", "Z"]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "Return"]);
    // The selection replace commits exactly `Z`: a leaked `a` (window
    // pushed the Ctrl combo as text) or a kept `AB` (select failed) fails.
    let deadline = Instant::now() + Duration::from_secs(8);
    loop {
        let log = std::fs::read_to_string(&path).unwrap_or_default();
        if log.lines().any(|l| l.trim() == "SET A1 Z") {
            break;
        }
        if Instant::now() > deadline {
            let _ = child.kill();
            panic!("select-all replace failed, log:\n{log}");
        }
        std::thread::sleep(Duration::from_millis(150));
    }
    let _ = std::fs::remove_file(&path);
}

/// Clicking a filled grid cell must show its value in the formula bar
/// (select semantics, ratatui parity) — not leave the bar blank.
///
/// Regression: grid clicks ended in `start_edit`, which clears the entry,
/// so the bar went empty on every click (and stayed empty while navigating
/// with the keyboard, since the update guard skipped rewrites while the
/// editing flag was on with an empty buffer).
#[test]
fn gui_grid_click_shows_cell_value() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path =
        std::env::temp_dir().join(format!("corro-formulaclick-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);
    // Long value on purpose: 2-char strings are OCR-invisible at entry
    // scale (tesseract returns ""), while longer words read reliably.
    std::fs::write(&path, "CORRO_LOG 1\nSET A1 HELLOFORMULA\n").expect("write fixture");

    let mut child = spawn_gui(&path);
    let pid = child.id();
    let wid = find_corro_window(pid, Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    await_first_paint(&wid);
    // A1's grid position: gutter (50) + half default column (~29) across,
    // header band (24) + half row (10) below the canvas top (~54 under Xvfb
    // with no window manager, so client coords equal root coords).
    let (wx, wy) = win_xy(&wid);
    // Drain present()'s event pump first: on virtual displays its ~1500
    // blocking iterations stall the frame clock for tens of seconds, and
    // paint verdicts before it drains read stale frames. Down/Up is a
    // net-zero cursor poke that keeps events flowing (same pattern as the
    // status-bar suite's first-paint wait).
    xdotool(&["key", "--window", &wid, "Down"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "Up"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&[
        "mousemove",
        &format!("{}", wx + 110),
        &format!("{}", wy + 88),
        "click",
        "1",
    ]);
    // Appearance event with a deadline: the entry must show the value.
    let deadline = Instant::now() + Duration::from_secs(45);
    let mut last = String::new();
    loop {
        let shot = screenshot(&wid, "clickshow");
        last = ocr_entry_zone(&shot);
        let _ = std::fs::remove_file(&shot);
        if last.contains("HELLO") {
            break;
        }
        if Instant::now() > deadline {
            let _ = child.kill();
            panic!("grid click did not show the cell value, entry OCR: {last:?}");
        }
        std::thread::sleep(Duration::from_millis(300));
    }
    let _ = std::fs::remove_file(&path);
}

/// Grid click, then formula-bar click, then typing must keep the formula
/// (the original wipe report through the grid-first flow): clicking the
/// grid selects (editing on, empty buffer), clicking the bar arms the
/// adopt, and the keystroke restores the displayed text around itself.
#[test]
fn gui_grid_click_then_entry_type_keeps_formula() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path =
        std::env::temp_dir().join(format!("corro-formulagrid-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, "CORRO_LOG 1\nSET A1 AB\n").expect("write fixture");

    let mut child = spawn_gui(&path);
    let pid = child.id();
    let wid = find_corro_window(pid, Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    await_first_paint(&wid);
    let (wx, wy) = win_xy(&wid);
    // Grid cell first (selects), then the formula bar (arms the adopt).
    // A1's middle sits near x+110 (the gutter + margin columns occupy the
    // first ~95px; verified against the click-address keylog).
    xdotool(&[
        "mousemove",
        &format!("{}", wx + 110),
        &format!("{}", wy + 88),
        "click",
        "1",
    ]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&[
        "mousemove",
        &format!("{}", wx + 300),
        &format!("{}", wy + 37),
        "click",
        "1",
    ]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["type", "X"]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "Return"]);
    // Without the adopt the log would show bare `SET A1 X`.
    let deadline = Instant::now() + Duration::from_secs(8);
    loop {
        let log = std::fs::read_to_string(&path).unwrap_or_default();
        if log.lines().any(|l| l.trim() == "SET A1 ABX") {
            break;
        }
        if Instant::now() > deadline {
            let _ = child.kill();
            panic!("grid-then-bar typing lost the formula, log:\n{log}");
        }
        std::thread::sleep(Duration::from_millis(150));
    }
    let _ = std::fs::remove_file(&path);
}
