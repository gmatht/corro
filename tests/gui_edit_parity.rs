//! Live GUI parity for type-first editing.
//!
//! The ratatui reference is the oracle: on a fresh 1x1 sheet Right,Right
//! leaves the main area (trailing-blank growth stops at 2x2), so A,Enter
//! commits to the RIGHT-MARGIN cell (`SET ]A1 A`), and Down x3 / Down x40
//! land in the footer (`SET A_2 Z` / `SET A_39 Q`) for the same reason.
//! These tests drive the real GTK GUI and assert on the committed `.corro`
//! file (no pixel matching, no fixed sleeps for the verdict — the file is
//! polled with a deadline).
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

/// Right,Right,A,Enter must commit exactly "A" to the right-margin cell
/// ]A1 (the two Rights leave the 1x1-start main area: trailing-blank growth
/// stops, so the cursor sits in the margin — exactly what the ratatui
/// reference commits: `SET ]A1 A`). First margin row/col addresses (and
/// saves) as margin, same as ratatui.
/// Regression: the GUI backend built CellAddr::Main unconditionally, so the
/// margin edit misrouted into an out-of-range main cell (accidentally
/// growing the grid and serializing as `SET C1 A`).
#[test]
fn gui_right_right_a_enter_commits_margin_bracket_a1() {
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
        lines.iter().any(|l| l == "SET ]A1 A"),
        "expected exactly `SET ]A1 A` (margin commit, matching ratatui), got lines: {lines:?}"
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

/// Up,Left,A,Enter must commit exactly "A" to the margin-header corner
/// cell `[A~1` (Up reaches the header band, Left steps into the left margin
/// — exactly what the ratatui reference commits: `SET [A~1 A`).
/// Regression: the GUI cursor clamped into the main area (Up/Left from A1
/// were no-ops) and commits built CellAddr::Main unconditionally, so the
/// sequence committed `SET A1 A` instead.
#[test]
fn gui_up_left_a_enter_commits_margin_header_corner() {
    // Tolerate poisoning: a failed sibling must not mask this test's verdict.
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-gui-seq1-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);

    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "--window", &wid, "Up"]);
    std::thread::sleep(Duration::from_millis(300));
    xdotool(&["key", "--window", &wid, "Left"]);
    std::thread::sleep(Duration::from_millis(300));
    xdotool(&["type", "--window", &wid, "A"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "Return"]);
    let lines = wait_file_lines(&path, Instant::now() + Duration::from_secs(10));
    let _ = child.kill();
    let _ = child.wait();

    assert!(
        lines.iter().any(|l| l == "SET [A~1 A"),
        "expected exactly `SET [A~1 A` (margin-header commit, matching ratatui), got lines: {lines:?}"
    );
    assert!(
        !lines.iter().any(|l| l == "SET A1 A"),
        "must not commit to main A1 (Up/Left must leave the main area), got: {lines:?}"
    );
}

/// Delete on a margin-header cell must clear THAT cell (empty SET op for
/// `[A~1`), not the clamped main cell. Regression: handle_delete built
/// CellAddr::Main unconditionally, so deleting a header/margin cell wiped
/// an unrelated main cell instead.
#[test]
fn gui_delete_clears_margin_header_cell() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-gui-delhdr-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);

    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    // Commit "A" to the margin-header corner, then delete it.
    for key in ["Up", "Left"] {
        xdotool(&["key", "--window", &wid, key]);
        std::thread::sleep(Duration::from_millis(300));
    }
    xdotool(&["type", "--window", &wid, "A"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "Return"]);
    // Wait for the commit, then delete and wait for the clear op.
    let deadline = Instant::now() + Duration::from_secs(10);
    loop {
        let content = std::fs::read_to_string(&path).unwrap_or_default();
        if content.lines().any(|l| l == "SET [A~1 A") { break; }
        if Instant::now() > deadline { panic!("no header commit in {}", path.display()); }
        std::thread::sleep(Duration::from_millis(100));
    }
    // Move back up onto the header cell (Enter moved down) and delete.
    xdotool(&["key", "--window", &wid, "Up"]);
    std::thread::sleep(Duration::from_millis(300));
    xdotool(&["key", "--window", &wid, "Delete"]);
    let lines = {
        let deadline = Instant::now() + Duration::from_secs(10);
        loop {
            let content = std::fs::read_to_string(&path).unwrap_or_default();
            let have: Vec<String> = content.lines().map(|l| l.to_string()).collect();
            // The commit plus the clear op: two ops for `[A~1`.
            if have.iter().filter(|l| l.starts_with("SET [A~1")).count() >= 2 { break have; }
            if Instant::now() > deadline {
                panic!("no clear op after Delete, got: {have:?}");
            }
            std::thread::sleep(Duration::from_millis(100));
        }
    };
    let _ = child.kill();
    let _ = child.wait();
    let last = lines.last().cloned().unwrap_or_default();
    assert!(
        last.starts_with("SET [A~1") && !last.ends_with(" A"),
        "Delete must clear the header cell (empty op for `[A~1`), got lines: {lines:?}"
    );
    assert!(
        !lines.iter().any(|l| l.starts_with("SET A1")),
        "Delete must not touch main A1, got: {lines:?}"
    );
}

/// A,Right,Right,C,Enter must commit "A" to A1 then "C" to C1: the first
/// Right grows the grid on the just-committed content (trailing < 2) and
/// the second grows again, exactly like the ratatui reference
/// (`SET A1 A`, `SET C1 C`).
/// Regression: commit-then-Right moved with a direct +1 that skipped the
/// growth check, so the second Right escaped into the margin and committed
/// `SET ]B1 C` instead of growing to C1.
#[test]
fn gui_a_right_right_c_enter_grows_to_c1() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-gui-growr-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, "CORRO_LOG 1\n").expect("write fixture");

    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    for key in ["A", "Right", "Right", "C", "Return"] {
        xdotool(&["key", "--window", &wid, key]);
        std::thread::sleep(Duration::from_millis(400));
    }
    let lines = wait_file_lines(&path, Instant::now() + Duration::from_secs(10));
    let _ = child.kill();
    let _ = child.wait();

    assert!(
        lines.iter().any(|l| l == "SET A1 A"),
        "expected `SET A1 A`, got lines: {lines:?}"
    );
    assert!(
        lines.iter().any(|l| l == "SET C1 C"),
        "expected `SET C1 C` (grid must grow on commit-then-Right, matching ratatui), got lines: {lines:?}"
    );
    assert!(
        !lines.iter().any(|l| l.contains("]")),
        "must not escape into the margin, got: {lines:?}"
    );
}

/// A,Down,Down,C,Enter must commit "A" to A1 then "C" to A3: commit-then-
/// Down grows like plain Down (ratatui reference: `SET A1 A`, `SET A3 C`).
/// Regression: same skipped-growth bug as the Right variant (committed to
/// the footer instead of growing to A3).
#[test]
fn gui_a_down_down_c_enter_grows_to_a3() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-gui-growd-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, "CORRO_LOG 1\n").expect("write fixture");

    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    for key in ["A", "Down", "Down", "C", "Return"] {
        xdotool(&["key", "--window", &wid, key]);
        std::thread::sleep(Duration::from_millis(400));
    }
    let lines = wait_file_lines(&path, Instant::now() + Duration::from_secs(10));
    let _ = child.kill();
    let _ = child.wait();

    assert!(
        lines.iter().any(|l| l == "SET A1 A"),
        "expected `SET A1 A`, got lines: {lines:?}"
    );
    assert!(
        lines.iter().any(|l| l == "SET A3 C"),
        "expected `SET A3 C` (grid must grow on commit-then-Down, matching ratatui), got lines: {lines:?}"
    );
    assert!(
        !lines.iter().any(|l| l.starts_with("SET A_")),
        "must not escape into the footer, got: {lines:?}"
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

/// Screenshot the window to `tag`.png (via xwd + convert). No focus needed.
fn screenshot(wid: &str, tag: &str) -> PathBuf {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let xwd = std::env::temp_dir().join(format!("corro-shot-{id}-{tag}.xwd"));
    let png = std::env::temp_dir().join(format!("corro-shot-{id}-{tag}.png"));
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

/// Selection-tint census of a screenshot, via an inline python3+PIL script.
/// Prints `marker_y sel_total sel_max_x sel_below_cutoff cursor_blue`.
/// Colors (cairo rounds halves up): selection fill (0.9,0.95,1.0) renders as
/// (230,243,255); cursor fill (0.8,0.9,1.0) as (204,230,255); the canvas
/// test marker is (254,237,190); the cursor border core is (0,102,204).
const SEL_ANALYZER: &str = r#"
import sys
from PIL import Image
png, y_cut = sys.argv[1], int(sys.argv[2])
img = Image.open(png).convert('RGB')
W, H = img.size
px = img.load()
MARK = (254, 237, 190)
SEL = (230, 243, 255)
BLUE = (0, 102, 204)
my = [y for y in range(H) for x in range(0, W, 4) if px[x, y] == MARK]
marker_y = min(my) if my else -1
sel_total = sel_max_x = sel_below = blue = 0
for y in range(H):
    for x in range(0, W, 2):
        c = px[x, y]
        if c == SEL:
            sel_total += 1
            if x > sel_max_x: sel_max_x = x
            if y > y_cut: sel_below += 1
        elif c == BLUE:
            blue += 1
print(f"{marker_y} {sel_total} {sel_max_x} {sel_below} {blue}")
"#;

struct SelCensus {
    marker_y: i32,
    sel_total: i32,
    sel_max_x: i32,
    sel_below: i32,
    blue: i32,
}

fn analyze_selection(png: &PathBuf, y_cut: i32) -> SelCensus {
    let script = std::env::temp_dir().join(format!(
        "corro-sel-an-{}.py",
        COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst)
    ));
    std::fs::write(&script, SEL_ANALYZER).expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(png)
        .arg(y_cut.to_string())
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let text = String::from_utf8_lossy(&out.stdout).trim().to_string();
    let p: Vec<i32> = text.split_whitespace().map(|s| s.parse().expect("ints")).collect();
    assert_eq!(p.len(), 5, "bad analyzer output: {text:?}");
    SelCensus { marker_y: p[0], sel_total: p[1], sel_max_x: p[2], sel_below: p[3], blue: p[4] }
}

/// Plain arrows must move WITHOUT painting a selection band (the anchor
/// collapses). Regression: the anchor stuck at startup, so Down x3 painted
/// a full-width band over rows 1-4 (~4700 tint px); now expect ~0.
/// Movement itself is proven by typing Z + Enter afterwards: on the fresh
/// 1x1 sheet Down x3 lands in the footer (growth stops at 2 rows), so the
/// file holds `SET A_2 Z` — exactly the ratatui reference.
#[test]
fn gui_plain_arrows_paint_no_selection_band() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(std::env::var("DISPLAY").is_ok(), "requires X server (run under xvfb-run -a)");
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-gui-nosel-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);

    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    for _ in 0..3 {
        xdotool(&["key", "--window", &wid, "Down"]);
        std::thread::sleep(Duration::from_millis(300));
    }
    std::thread::sleep(Duration::from_millis(500));
    // Settle: two captures must agree the selection count is stable, so a
    // mid-render tear cannot produce the verdict.
    let png = screenshot(&wid, "nosel");
    let c1 = analyze_selection(&png, 0);
    std::thread::sleep(Duration::from_millis(400));
    let png2 = screenshot(&wid, "nosel2");
    let c2 = analyze_selection(&png2, 0);
    assert!(
        (c1.sel_total - c2.sel_total).abs() < 200,
        "selection count unsettled ({} vs {}), retry",
        c1.sel_total, c2.sel_total
    );
    assert!(
        c2.sel_total < 500,
        "plain Down x3 must paint no selection band (tint px={}), marker_y={}",
        c2.sel_total, c2.marker_y
    );
    // Prove the cursor really reached A4 (guards a vacuous pass where no
    // keys landed at all).
    xdotool(&["type", "--window", &wid, "Z"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "Return"]);
    let lines = wait_file_lines(&path, Instant::now() + Duration::from_secs(10));
    let _ = child.kill();
    let _ = child.wait();
    assert!(
        lines.iter().any(|l| l == "SET A_2 Z"),
        "expected footer commit `SET A_2 Z` after Down x3 (matching ratatui), got: {lines:?}"
    );
}

/// Shift+Right extends a true rectangle (anchor..cursor, cols included) —
/// not a full-width band. Starts with a plain Right (leaves startup edit
/// mode, cursor B1, no selection); then Shift+Right x2 selects B1..D1.
/// Regression: shift was ignored (full-width band) or dropped.
#[test]
fn gui_shift_right_extends_rect_not_band() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(std::env::var("DISPLAY").is_ok(), "requires X server (run under xvfb-run -a)");
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-gui-selrect-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);

    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "--window", &wid, "Right"]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "--window", &wid, "Shift+Right"]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "--window", &wid, "Shift+Right"]);
    std::thread::sleep(Duration::from_millis(600));
    let png = screenshot(&wid, "selrect");
    let probe = analyze_selection(&png, 0);
    assert!(probe.marker_y > 0, "canvas marker must render");
    // Row 1 (first main row) ends at marker+24 (header) +20 (row); cut below it.
    let c = analyze_selection(&png, probe.marker_y + 24 + 20 + 8);
    let _ = child.kill();
    let _ = child.wait();
    assert!(
        c.sel_total > 50,
        "Shift+Right x2 must highlight a selection (tint px={})", c.sel_total
    );
    assert!(
        c.sel_max_x < 400,
        "selection must be a rect (cols B..D), not a full-width band (max_x={})", c.sel_max_x
    );
    assert!(
        c.sel_below == 0,
        "selection must stay in row 1 (below-cutoff tint px={})", c.sel_below
    );
}

/// Cursor must stay visible: after Down x40 (past the ~30 visible rows)
/// the selected cell's blue border must render on screen, and typing + Enter
/// must commit at the arrived address (`SET A_39 Q`: on the fresh 1x1 sheet
/// the Downs land in the footer — exactly the ratatui reference), proving
/// the viewport followed.
#[test]
fn gui_deep_move_keeps_cursor_visible() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(std::env::var("DISPLAY").is_ok(), "requires X server (run under xvfb-run -a)");
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-gui-follow-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);

    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    for _ in 0..40 {
        xdotool(&["key", "--window", &wid, "Down"]);
        std::thread::sleep(Duration::from_millis(120));
    }
    std::thread::sleep(Duration::from_millis(600));
    // Cursor border core is (0,102,204); the fill is (204,230,255). Both
    // render exact (verified: the formula entry's focus ring never produces
    // the exact core color). Either one proves the cursor cell painted.
    let png = screenshot(&wid, "follow");
    let script = std::env::temp_dir().join(format!(
        "corro-follow-an-{}.py",
        COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst)
    ));
    std::fs::write(&script, "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\nW,H = img.size\npx = img.load()\nCF=(204,230,255)\nBLUE=(0,102,204)\ncf=sum(1 for y in range(H) for x in range(0,W,2) if px[x,y]==CF)\nbl=sum(1 for y in range(H) for x in range(0,W,2) if px[x,y]==BLUE)\nprint(f'{cf} {bl}')\n").expect("write analyzer");
    let out = Command::new("python3").arg(&script).arg(&png).output().expect("python3");
    let _ = std::fs::remove_file(&script);
    let nums: Vec<i32> = String::from_utf8_lossy(&out.stdout).trim().split_whitespace().map(|s| s.parse().unwrap_or(-99)).collect();
    let (fill, blue) = (nums[0], nums[1]);
    // Prove arrival deep down via commit (guards a vacuous pass where no
    // keys landed: unmoved cursor would commit at A1 instead). The footer
    // address matches the ratatui reference for Down x40 on a fresh sheet.
    xdotool(&["type", "--window", &wid, "Q"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "Return"]);
    let lines = wait_file_lines(&path, Instant::now() + Duration::from_secs(10));
    let _ = child.kill();
    let _ = child.wait();
    assert!(
        fill > 100 || blue > 10,
        "cursor cell must render after Down x40 (fill px={fill}, border px={blue}) — viewport did not follow"
    );
    assert!(
        lines.iter().any(|l| l == "SET A_39 Q"),
        "expected footer commit `SET A_39 Q` after Down x40 (matching ratatui), got: {lines:?}"
    );
}

/// Scrollbar presence: a vertical scrollbar trough must run along the right
/// edge of the grid band (uniform non-white column where grid content would
/// otherwise reach the edge). Regression guard: canvases without scrollbars
/// (or with dead ones) show grid/border pixels there instead.
#[test]
fn gui_vertical_scrollbar_present() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(std::env::var("DISPLAY").is_ok(), "requires X server (run under xvfb-run -a)");
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-gui-sbpres-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);

    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(800));
    let png = screenshot(&wid, "sbpres");
    let script = std::env::temp_dir().join(format!(
        "corro-sbpres-an-{}.py",
        COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst)
    ));
    // Right-edge column over the grid band: a scrollbar trough is one uniform
    // non-white gray; grid content reaching the edge is mostly white with
    // border/text pixels mixed in.
    std::fs::write(&script, "import sys\nfrom PIL import Image\nfrom collections import Counter\nimg = Image.open(sys.argv[1]).convert('RGB')\nW,H = img.size\npx = img.load()\ncol = [px[W-10, y] for y in range(150, H-150, 2)]\nmc = Counter(col).most_common(1)[0]\nprint(f'{mc[0][0]} {mc[0][1]} {mc[0][2]} {mc[1]} {len(col)}')\n").expect("write analyzer");
    let out = Command::new("python3").arg(&script).arg(&png).output().expect("python3");
    let _ = std::fs::remove_file(&script);
    let _ = child.kill();
    let _ = child.wait();
    let p: Vec<i32> = String::from_utf8_lossy(&out.stdout).trim().split_whitespace().map(|s| s.parse().unwrap_or(-99)).collect();
    assert_eq!(p.len(), 5, "bad analyzer output");
    let (r, g, b, count, total) = (p[0], p[1], p[2], p[3], p[4]);
    let uniform_frac = count as f64 / total as f64;
    let is_white = r > 250 && g > 250 && b > 250;
    assert!(
        !is_white && uniform_frac > 0.85,
        "expected scrollbar trough (uniform non-white) at right edge, got rgb=({r},{g},{b}) frac={uniform_frac:.2}"
    );
}

/// Scrollbar function: clicking the trough below the thumb must move the
/// selection down (formula row advances), proving the bars drive the sheet.
#[test]
fn gui_scrollbar_trough_click_moves_down() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(std::env::var("DISPLAY").is_ok(), "requires X server (run under xvfb-run -a)");
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-gui-sbclick-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);

    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    // Window screen geometry (no WM under Xvfb: client origin == window origin).
    let geo = xdotool(&["getwindowgeometry", "--shell", &wid]);
    let mut wx = 0i32;
    let mut wy = 0i32;
    let mut ww = 0i32;
    let mut wh = 0i32;
    for line in geo.lines() {
        if let Some(v) = line.strip_prefix("X=") { wx = v.trim().parse().unwrap_or(0); }
        if let Some(v) = line.strip_prefix("Y=") { wy = v.trim().parse().unwrap_or(0); }
        if let Some(v) = line.strip_prefix("WIDTH=") { ww = v.trim().parse().unwrap_or(0); }
        if let Some(v) = line.strip_prefix("HEIGHT=") { wh = v.trim().parse().unwrap_or(0); }
    }
    assert!(ww > 100 && wh > 100, "bad geometry {ww}x{wh}");
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    // Click the vertical trough well below the top (thumb sits at top while
    // the cursor is on row 1): trough click pages the selection down.
    xdotool(&["mousemove", "--sync", &(wx + ww - 10).to_string(), &(wy + wh * 3 / 4).to_string()]);
    xdotool(&["click", "1"]);
    std::thread::sleep(Duration::from_millis(800));
    // Prove the selection moved down via commit (guards clicks that land on
    // nothing: an unmoved cursor would commit at A1).
    xdotool(&["type", "--window", &wid, "W"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "Return"]);
    let lines = wait_file_lines(&path, Instant::now() + Duration::from_secs(10));
    let _ = child.kill();
    let _ = child.wait();
    // The trough pages down past the 2-row main area into the footer, so
    // the commit serializes with the footer prefix (`SET A_11 W`); accept
    // both main (`A<num>`) and footer (`A_<num>`) forms and assert the row.
    let committed_row: Option<u32> = lines.iter().find_map(|l| {
        let rest = l.strip_prefix("SET A").or_else(|| l.strip_prefix("SET A_"))?;
        let rest = rest.strip_prefix('_').unwrap_or(rest);
        let num: String = rest.chars().take_while(|c| c.is_ascii_digit()).collect();
        num.parse().ok()
    });
    assert!(
        committed_row.map_or(false, |r| r > 5),
        "trough click must move selection down several rows, got lines: {lines:?}"
    );
}

/// Click at a fraction down the currently-open menu popup (found as the
/// non-main window of this pid). Synthetic key events cannot drive GTK's
/// grab-based menu navigation, but synthetic button presses reach the popup
/// directly, so clicking a menu row is the reliable activation path.
fn click_popup_fraction(pid: u32, wid: &str, frac: f64) {
    let pid = pid.to_string();
    let mut popup = String::new();
    let deadline = Instant::now() + Duration::from_secs(8);
    loop {
        let out = xdotool(&["search", "--pid", &pid]);
        for id in out.split_whitespace() {
            if id == wid {
                continue;
            }
            let geo = xdotool(&["getwindowgeometry", "--shell", id]);
            let wide = geo.lines().any(|l| {
                l.strip_prefix("WIDTH=")
                    .and_then(|v| v.trim().parse::<i32>().ok())
                    .map(|w| w > 40)
                    .unwrap_or(false)
            });
            if wide {
                popup = id.to_string();
                break;
            }
        }
        if !popup.is_empty() {
            break;
        }
        if Instant::now() > deadline {
            panic!("timed out waiting for menu popup to click");
        }
        std::thread::sleep(Duration::from_millis(150));
    }
    let geo = xdotool(&["getwindowgeometry", "--shell", &popup]);
    let (mut px, mut py, mut ph) = (0, 0, 0);
    for l in geo.lines() {
        if let Some(v) = l.strip_prefix("X=") {
            px = v.trim().parse().unwrap_or(0);
        }
        if let Some(v) = l.strip_prefix("Y=") {
            py = v.trim().parse().unwrap_or(0);
        }
        if let Some(v) = l.strip_prefix("HEIGHT=") {
            ph = v.trim().parse().unwrap_or(0);
        }
    }
    assert!(ph > 60, "popup height implausible ({ph})");
    // Settle so the popup is mapped and input-ready (a click in the same
    // instant as appearance can land before GTK finishes mapping it).
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&[
        "mousemove",
        &format!("{}", px + 50),
        &format!("{}", py + (ph as f64 * frac) as i32),
        "click",
        "1",
    ]);
}

/// Poll a `.corro` file with a deadline until `pred` holds of its lines.
fn wait_file_pred(
    path: &PathBuf,
    deadline: Instant,
    what: &str,
    pred: impl Fn(&[String]) -> bool,
) -> Vec<String> {
    loop {
        if let Ok(content) = std::fs::read_to_string(path) {
            let lines: Vec<String> = content.lines().map(|l| l.to_string()).collect();
            if pred(&lines) {
                return lines;
            }
        }
        if Instant::now() > deadline {
            panic!(
                "timed out waiting for {what} in {}\ncontent: {:?}",
                path.display(),
                std::fs::read_to_string(path).unwrap_or_default()
            );
        }
        std::thread::sleep(Duration::from_millis(100));
    }
}

fn date_today() -> String {
    String::from_utf8_lossy(
        &Command::new("date")
            .arg("+%F")
            .output()
            .expect("date +%F")
            .stdout,
    )
    .trim()
    .to_string()
}

/// Insert > Date must preset today's date and commit it on Enter (GTK live).
/// Regression: the GTK backend answered "Insert date not yet implemented"
/// and committed nothing; the shared dispatch this now delegates to presets
/// the edit buffer exactly like pancurses/ratatui.
#[test]
fn gui_insert_date_commits_today() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-gui-insdate-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, "CORRO_LOG 1\n").expect("write fixture");

    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    // Insert menu (Alt+I), Date is the 'D' mnemonic item.
    xdotool(&["key", "--window", &wid, "alt+i"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "d"]);
    std::thread::sleep(Duration::from_millis(500));
    // Commit the preset date.
    xdotool(&["key", "--window", &wid, "Return"]);
    let today = date_today();
    let expected = format!("SET A1 {today}");
    let lines = wait_file_pred(&path, Instant::now() + Duration::from_secs(10), "date commit", |ls| {
        ls.iter().any(|l| l == &expected)
    });
    let _ = child.kill();
    let _ = child.wait();
    assert!(
        lines.iter().any(|l| l == &expected),
        "Insert Date should commit `SET A1 {today}`, got: {lines:?}"
    );
    let _ = std::fs::remove_file(&path);
}

/// Insert > Time must preset the clock time and commit it on Enter (GTK live).
/// Same Edit-preset mechanism as Date; the committed HH:MM:SS must be within
/// two minutes of now (guards second-boundary flakes without fixed sleeps).
#[test]
fn gui_insert_time_commits_time() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-gui-instime-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, "CORRO_LOG 1\n").expect("write fixture");

    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    // Insert menu (Alt+I), Time is the 'T' mnemonic item.
    xdotool(&["key", "--window", &wid, "alt+i"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "t"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "Return"]);
    let lines = wait_file_pred(&path, Instant::now() + Duration::from_secs(10), "time commit", |ls| {
        ls.iter().any(|l| {
            l.strip_prefix("SET A1 ")
                .map(|v| {
                    v.len() == 8
                        && v.as_bytes()[2] == b':'
                        && v.as_bytes()[5] == b':'
                        && v[..2].parse::<u32>().is_ok()
                })
                .unwrap_or(false)
        })
    });
    let _ = child.kill();
    let _ = child.wait();
    let committed: String = lines
        .iter()
        .find_map(|l| l.strip_prefix("SET A1 "))
        .unwrap_or_default()
        .to_string();
    // Within two minutes of now (robust to second ticks, no fixed sleeps).
    // Compare in local wall-clock: the app formats Local time, and `date`
    // without -u agrees with it (epoch seconds would be UTC-shifted).
    let now_out = Command::new("date").arg("+%H %M %S").output().expect("date");
    let now_hms = String::from_utf8_lossy(&now_out.stdout);
    let now_parts: Vec<i64> = now_hms
        .split_whitespace()
        .map(|p| p.parse().unwrap_or(-1))
        .collect();
    assert_eq!(now_parts.len(), 3, "could not read local time");
    let now_secs = now_parts[0] * 3600 + now_parts[1] * 60 + now_parts[2];
    let parts: Vec<i64> = committed.split(':').map(|p| p.parse().unwrap_or(-1)).collect();
    assert_eq!(parts.len(), 3, "committed time should be HH:MM:SS, got {committed:?}");
    let got_secs = parts[0] * 3600 + parts[1] * 60 + parts[2];
    assert!(
        (now_secs - got_secs).abs() <= 120,
        "committed time {committed:?} is not within two minutes of now"
    );
    let _ = std::fs::remove_file(&path);
}

/// Edit > Cut then Edit > Paste round-trips the cell value through the menu
/// (GTK live). Regression: the GTK backend answered "Cut/Paste not yet
/// implemented". This also exercises the GUI clipboard plumbing behind the
/// shared dispatch: paste must restore exactly what cut took.
#[test]
fn gui_cut_paste_roundtrip_via_menu() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-gui-cutpaste-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, "CORRO_LOG 1\n").expect("write fixture");

    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    // Seed A1 with "A" (Enter commits and moves down to A2).
    xdotool(&["type", "--window", &wid, "A"]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "--window", &wid, "Return"]);
    wait_file_pred(&path, Instant::now() + Duration::from_secs(10), "seed commit", |ls| {
        ls.iter().any(|l| l == "SET A1 A")
    });
    // Move back up to A1 (Enter advanced to A2); Cut/Paste target A1.
    xdotool(&["key", "--window", &wid, "Up"]);
    std::thread::sleep(Duration::from_millis(300));
    // Edit menu (Alt+E), Cut is row 1 of 7. Cut has no keyboard mnemonic
    // (its documented shortcut X appears nowhere in "Cut"), so activate by
    // click like the menu-key tests: clears A1.
    xdotool(&["key", "--window", &wid, "alt+e"]);
    std::thread::sleep(Duration::from_millis(500));
    click_popup_fraction(child.id(), &wid, 0.5 / 7.0);
    wait_file_pred(&path, Instant::now() + Duration::from_secs(10), "cut clear", |ls| {
        ls.iter().any(|l| l.starts_with("SET A1") && l != "SET A1 A")
    });
    // Edit menu, Paste is row 3 of 7: restores "A" from the menu clipboard.
    xdotool(&["key", "--window", &wid, "alt+e"]);
    std::thread::sleep(Duration::from_millis(500));
    click_popup_fraction(child.id(), &wid, 2.5 / 7.0);
    let lines = wait_file_pred(&path, Instant::now() + Duration::from_secs(10), "paste restore", |ls| {
        ls.iter().filter(|l| *l == "SET A1 A").count() >= 2
    });
    let _ = child.kill();
    let _ = child.wait();
    assert!(
        lines.iter().filter(|l| *l == "SET A1 A").count() >= 2,
        "paste should restore a second `SET A1 A` after cut cleared it, got: {lines:?}"
    );
    let _ = std::fs::remove_file(&path);
}
