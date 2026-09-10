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
/// Movement itself is proven by typing Z + Enter afterwards (file = A4).
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
        lines.iter().any(|l| l == "SET A4 Z"),
        "expected commit at A4 after Down x3, got: {lines:?}"
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
/// must commit at the arrived address (A41), proving the viewport followed.
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
    // Prove arrival at A41 via commit (guards a vacuous pass where no keys
    // landed: unmoved cursor would commit at A1 instead).
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
        lines.iter().any(|l| l == "SET A41 Q"),
        "expected commit at A41 after Down x40, got: {lines:?}"
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
    let committed_row: Option<u32> = lines.iter().find_map(|l| {
        let rest = l.strip_prefix("SET A")?;
        let num: String = rest.chars().take_while(|c| c.is_ascii_digit()).collect();
        num.parse().ok()
    });
    assert!(
        committed_row.map_or(false, |r| r > 5),
        "trough click must move selection down several rows, got lines: {lines:?}"
    );
}
