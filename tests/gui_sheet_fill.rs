//! GUI sheet-fill parity: after menus/formula bar/status chrome, the sheet
//! grid must extend to the window edges — no huge blank areas.
//!
//! Regression: the GUI backend sized the viewport from hardcoded counts
//! (12 columns), so the grid stopped ~400px into a 1200px window and the
//! rest stayed blank window background. The count is now derived from the
//! live canvas size every frame.
//!
//! Method (no pixel-perfect coupling, no shared files): screenshot the real
//! window, locate the canvas via the 8x8 0xFEEDBE test marker it draws at
//! its origin, then measure how far grid-white pixels reach toward the
//! right/bottom edges. Asserts at two window sizes — the second proves the
//! sheet EXPANDS (a fixed-size canvas passes small and fails large).
//!
//! Requires: Linux, an X server, xdotool, xwd, convert (ImageMagick),
//! python3+PIL. Run with:
//!   xvfb-run -a cargo test --features gtk --test gui_sheet_fill
#![cfg(all(target_os = "linux", feature = "gui"))]

use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::time::{Duration, Instant};

static COUNTER: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

/// Serialize the live-GUI tests in this binary: parallel windows steal focus
/// and blank each other's screenshots (same pattern as gui_edit_parity.rs).
static GUI_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

fn xdotool(args: &[&str]) -> String {
    String::from_utf8_lossy(
        &Command::new("xdotool").args(args).output().expect("xdotool failed").stdout,
    )
    .to_string()
}

/// Python analyzer: prints `marker_x marker_y right_gap bottom_gap`.
/// All measurements are structural (marker bbox, white-pixel extents), so
/// font/theme shifts move them by px, not by hundreds (blank-area scale).
const ANALYZER: &str = r#"
import sys
from PIL import Image
path = sys.argv[1]
img = Image.open(path).convert('RGB')
W, H = img.size
px = img.load()
MARK = (254, 237, 190)
WHITE = (255, 255, 255)
MARGIN = (191, 191, 191)
def is_grid(px):
    return px == WHITE or px == MARGIN
# Canvas origin: bbox of the marker the canvas draws at (0,0).
mx0, my0, mx1, my1 = W, H, -1, -1
for y in range(0, H, 2):
    for x in range(0, W, 2):
        if px[x, y] == MARK:
            if x < mx0: mx0 = x
            if x + 1 > mx1: mx1 = x + 1
            if y < my0: my0 = y
            if y + 1 > my1: my1 = y + 1
if mx1 < 0:
    print("NO_MARKER")
    sys.exit(0)
# Refine to exact pixels around the coarse bbox.
ex0 = max(0, mx0 - 2); ex1 = min(W - 1, mx1 + 2)
ey0 = max(0, my0 - 2); ey1 = min(H - 1, my1 + 2)
bx0, by0, bx1, by1 = W, H, -1, -1
for y in range(ey0, ey1 + 1):
    for x in range(ex0, ex1 + 1):
        if px[x, y] == MARK:
            bx0 = min(bx0, x); by0 = min(by0, y)
            bx1 = max(bx1, x); by1 = max(by1, y)
# Right extent: rightmost grid-cell interior in the grid band (white body
# or dimmed-margin gray; cell interiors are ~28px runs, text antialiasing
# never spans that far, and chrome grays differ from both).
right = -1
for y in range(by1 + 10, H - 30, 4):
    # scan from the right; stop at first grid cell
    for x in range(W - 1, bx0, -1):
        if is_grid(px[x, y]):
            if x > right: right = x
            break
# Bottom extent: bottommost grid-cell pixel anywhere below the marker.
# (Window chrome below the grid — status label — is neither white nor
# margin gray, so grid colors always mean grid.)
bottom = -1
for y in range(H - 1, by1, -1):
    found = False
    for x in range(bx0, W, 3):
        if is_grid(px[x, y]):
            bottom = y
            found = True
            break
    if found:
        break
print(f"{bx0} {by0} {W - 1 - right} {H - 1 - bottom}")
"#;

struct Fill {
    marker_x: i32,
    marker_y: i32,
    right_gap: i32,
    bottom_gap: i32,
}

fn analyze(png: &PathBuf) -> Fill {
    let script = std::env::temp_dir().join(format!(
        "corro-fill-an-{}.py",
        COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst)
    ));
    std::fs::write(&script, ANALYZER).expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(png)
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let text = String::from_utf8_lossy(&out.stdout).trim().to_string();
    assert_ne!(text, "NO_MARKER", "canvas test marker not found in {png:?} (grid not drawn?)");
    let p: Vec<i32> = text.split_whitespace().map(|s| s.parse().expect("ints")).collect();
    assert_eq!(p.len(), 4, "bad analyzer output: {text:?}");
    Fill { marker_x: p[0], marker_y: p[1], right_gap: p[2], bottom_gap: p[3] }
}

// Non-panicking analyzer for settle loops: None while the window hasn't
// rendered yet (blank first frames under Xvfb are normal).
fn try_analyze(png: &PathBuf) -> Option<Fill> {
    let script = std::env::temp_dir().join(format!(
        "corro-fill-try-{}.py",
        COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst)
    ));
    std::fs::write(&script, ANALYZER).expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(png)
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let text = String::from_utf8_lossy(&out.stdout).trim().to_string();
    if text == "NO_MARKER" {
        return None;
    }
    let p: Vec<i32> = text.split_whitespace().map(|s| s.parse().expect("ints")).collect();
    if p.len() != 4 {
        return None;
    }
    Some(Fill { marker_x: p[0], marker_y: p[1], right_gap: p[2], bottom_gap: p[3] })
}

/// Screenshot until the canvas marker renders (deadline), then settle until
/// two captures 400ms apart agree — so neither a blank first frame nor a
/// mid-move tear can produce the verdict.
fn capture_settled(wid: &str, tag: &str) -> Fill {
    let deadline = Instant::now() + Duration::from_secs(15);
    let mut prev: Option<Fill> = None;
    loop {
        if Instant::now() > deadline {
            panic!("timed out waiting for rendered canvas in settled capture");
        }
        let shot = screenshot(wid, tag);
        let cur = try_analyze(&shot);
        let _ = std::fs::remove_file(&shot);
        match (prev.take(), cur) {
            (Some(p), Some(c))
                if p.marker_x == c.marker_x && p.marker_y == c.marker_y =>
            {
                return c
            }
            (_, c) => {
                prev = c;
                std::thread::sleep(Duration::from_millis(400));
            }
        }
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

fn find_window(child_pid: u32) -> String {
    let pid = child_pid.to_string();
    let deadline = Instant::now() + Duration::from_secs(25);
    loop {
        let out = xdotool(&["search", "--pid", &pid]);
        for cand in out.split_whitespace() {
            let geo = xdotool(&["getwindowgeometry", "--shell", cand]);
            let wide = geo
                .lines()
                .any(|l| {
                    l.strip_prefix("WIDTH=")
                        .and_then(|v| v.trim().parse::<i32>().ok())
                        .map(|w| w > 100)
                        .unwrap_or(false)
                });
            if wide {
                return cand.to_string();
            }
        }
        if Instant::now() > deadline {
            panic!("timed out waiting for corro window");
        }
        std::thread::sleep(Duration::from_millis(300));
    }
}

fn screenshot(wid: &str, tag: &str) -> PathBuf {
    let xwd = std::env::temp_dir().join(format!("corro-fill-{tag}.xwd"));
    let png = std::env::temp_dir().join(format!("corro-fill-{tag}.png"));
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

/// Assert the sheet fills the window at (w, h): marker present near the
/// top-left chrome, grid white reaching the right edge (a fixed 12-column
/// viewport left ~800px blank here) and near the bottom chrome (fixed 30
/// rows leave ~285px blank at 1000px tall).
fn assert_fills(wid: &str, tag: &str, w: i32, h: i32) {
    xdotool(&["windowsize", wid, &w.to_string(), &h.to_string()]);
    // Settle: two captures 400ms apart must agree on the marker bbox, so a
    // mid-resize tear cannot produce the verdict.
    let mut prev = analyze(&screenshot(wid, tag));
    let mut tries = 0;
    let f = loop {
        std::thread::sleep(Duration::from_millis(400));
        let next = analyze(&screenshot(wid, tag));
        tries += 1;
        if next.marker_x == prev.marker_x && next.marker_y == prev.marker_y {
            break next;
        }
        prev = next;
        // Bounded: even a perpetually-shifting layout must produce a verdict
        // (the structural asserts below still apply to the last capture).
        if tries > 10 {
            break prev;
        }
    };
    assert!(
        f.marker_x <= 20,
        "canvas should start at the left edge, marker_x={} at {w}x{h}",
        f.marker_x
    );
    assert!(
        f.marker_y <= 250,
        "canvas should start below menu+formula chrome, marker_y={} at {w}x{h}",
        f.marker_y
    );
    assert!(
        f.right_gap <= 60,
        "grid must reach the right edge (right_gap={}) at {w}x{h}",
        f.right_gap
    );
    assert!(
        f.bottom_gap <= 160,
        "grid must reach the bottom chrome (bottom_gap={}) at {w}x{h}",
        f.bottom_gap
    );
}

#[test]
fn gui_sheet_fills_window_after_chrome() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-fill-{}-{}.corro", std::process::id(), id));
    // Tall sheet (100 rows): the grid CAN fill a 1000px-tall window, so a
    // bottom blank means the viewport never expanded (fixed 30 rows leave
    // ~285px blank). With a tiny sheet the rows legitimately end early on
    // every backend, which would make the assertion meaningless.
    let mut log = String::from("CORRO_LOG 1\n");
    for r in 1..=100 {
        log.push_str(&format!("SET A{r} x\n"));
    }
    std::fs::write(&path, log).expect("write fill fixture");

    let mut child = spawn_gui(&path);
    let wid = find_window(child.id());
    // Small window: catches a viewport stuck at a fixed count (12 cols left
    // ~800px blank at 1200 wide).
    assert_fills(&wid, "small", 1200, 800);
    // Large window (fits a 1280x1024 Xvfb screen): catches a viewport that
    // never expands — proves live expansion, not a fixed size.
    assert_fills(&wid, "large", 1280, 1000);
    let _ = child.kill();
    let _ = child.wait();
    let _ = std::fs::remove_file(&path);
}


/// Poll a `.corro` file with a deadline until one line satisfies `pred`.
fn wait_file_pred(
    path: &std::path::PathBuf,
    what: &str,
    pred: impl Fn(&str) -> bool,
) -> Vec<String> {
    let deadline = Instant::now() + Duration::from_secs(10);
    loop {
        if let Ok(content) = std::fs::read_to_string(path) {
            let lines: Vec<String> = content.lines().map(|l| l.to_string()).collect();
            if lines.iter().any(|l| pred(l)) {
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

/// Type Q + Enter and assert the commit lands in a margin cell (`SET [`):
/// proves the Left keys actually reached the margin (guards a vacuous pass
/// where lost keys leave the cursor on A1 and the fill trivially holds).
fn assert_margin_commit(wid: &str, path: &std::path::PathBuf) {
    xdotool(&["key", "--window", wid, "q"]);
    std::thread::sleep(Duration::from_millis(300));
    xdotool(&["key", "--window", wid, "Return"]);
    let lines = wait_file_pred(path, "margin-zone commit", |l| l.starts_with("SET ["));
    assert!(
        lines.iter().any(|l| l.starts_with("SET [")),
        "typed Q should commit into the margin zone after Left, got: {lines:?}"
    );
}

/// Walk Left one step at a time from A1 (depths 1..=10 into the margin),
/// asserting after EVERY step that the sheet still reaches the right window
/// edge. A single end-state assertion would miss depths where the viewport
/// strands narrow; the loop pins the whole margin band. Finishes with a
/// margin-zone commit proving the keys really landed in the margin
/// (non-vacuous: lost keys would leave the cursor on A1).
fn left_arrow_sweep(wid: &str, path: &std::path::PathBuf, tag: &str) {
    for depth in 1..=10 {
        xdotool(&["key", "--window", wid, "Left"]);
        std::thread::sleep(Duration::from_millis(300));
        let f = capture_settled(wid, tag);
        assert!(
            f.right_gap <= 60,
            "grid must still reach the right edge after {depth} Left(s) into the margin (right_gap={})",
            f.right_gap
        );
    }
    // Rows are unchanged by horizontal moves (covered by
    // gui_sheet_fills_window_after_chrome); the regression is horizontal.
    assert_margin_commit(wid, path);
}

/// After Left-arrowing into the left margin, the sheet must still reach the
/// right window edge — no huge blank area.
///
/// Regression: with the cursor in the left margin, `visible_col_indices`
/// returns dim-1 columns, so `cols_to_fill_px`'s `cols.len() < dim` exit
/// fired after two iterations (data_cols=9) and the grid stopped ~460px
/// into a 1280px window. Populated sheet (overflow.corro exercises margin
/// columns plus footer rows).
#[test]
fn gui_left_arrow_fills_window() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-fill-left-{}-{}.corro", std::process::id(), id));
    std::fs::copy(
        format!("{}/docs/tests/overflow.corro", env!("CARGO_MANIFEST_DIR")),
        &path,
    )
    .expect("copy overflow fixture");

    let mut child = spawn_gui(&path);
    let wid = find_window(child.id());
    xdotool(&["windowsize", &wid, "1200", "800"]);
    std::thread::sleep(Duration::from_millis(400));
    // Cursor starts on A1; sweep 1..=10 columns into the left margin.
    left_arrow_sweep(&wid, &path, "left");
    let _ = child.kill();
    let _ = child.wait();
    let _ = std::fs::remove_file(&path);
}

/// Same as above on a tiny (1x1, empty) sheet: margin columns always exist,
/// so even the smallest grid must fill the window after Left.
#[test]
fn gui_left_arrow_fills_window_empty() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-fill-left-empty-{}-{}.corro", std::process::id(), id));
    std::fs::write(&path, "CORRO_LOG 1\n").expect("write empty fixture");

    let mut child = spawn_gui(&path);
    let wid = find_window(child.id());
    xdotool(&["windowsize", &wid, "1200", "800"]);
    std::thread::sleep(Duration::from_millis(400));
    left_arrow_sweep(&wid, &path, "leftempty");
    let _ = child.kill();
    let _ = child.wait();
    let _ = std::fs::remove_file(&path);
}

/// Horizontal-thumb census of a screenshot: (count, min_x, max_x) of
/// dark-slate thumb pixels in the h-scrollbar band (y 770..782, 1200x800
/// window). Matches the normal (126,129,130) and pressed (86,91,92)
/// thumb shades; trough (206) and text (<100) do not match.
fn hthumb(shot: &PathBuf) -> (i32, i32, i32) {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fill-thumb-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\nW,H = img.size\npx = img.load()\nxs=[x for y in range(770,782) for x in range(0,W) if 75<=px[x,y][0]<=155 and abs(px[x,y][0]-px[x,y][1])<14 and abs(px[x,y][1]-px[x,y][2])<14]\nprint(f'{len(xs)} {(min(xs) if xs else -1)} {(max(xs) if xs else -1)}')\n",
    )
    .expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(shot)
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let _ = std::fs::remove_file(shot);
    let text = String::from_utf8_lossy(&out.stdout).trim().to_string();
    let p: Vec<i32> = text.split_whitespace().map(|s| s.parse().unwrap_or(-99)).collect();
    assert_eq!(p.len(), 3, "bad analyzer output: {text:?}");
    (p[0], p[1], p[2])
}

/// Vertical-thumb census of a screenshot: (count, min_y, max_y) of
/// dark-slate thumb pixels in the v-scrollbar column (x = W-8, y 40..H-40
/// of the 1200x800 window). Matches normal and pressed thumb shades;
/// trough (206) and text do not match.
fn vthumb(shot: &PathBuf) -> (i32, i32, i32) {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fill-vthumb-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\nW,H = img.size\npx = img.load()\nys=[y for y in range(40,H-40) for x in [W-8] if 75<=px[x,y][0]<=155 and abs(px[x,y][0]-px[x,y][1])<14 and abs(px[x,y][1]-px[x,y][2])<14]\nprint(f'{len(ys)} {(min(ys) if ys else -1)} {(max(ys) if ys else -1)}')\n",
    )
    .expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(shot)
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let _ = std::fs::remove_file(shot);
    let text = String::from_utf8_lossy(&out.stdout).trim().to_string();
    let p: Vec<i32> = text.split_whitespace().map(|s| s.parse().unwrap_or(-99)).collect();
    assert_eq!(p.len(), 3, "bad analyzer output: {text:?}");
    (p[0], p[1], p[2])
}

/// Cursor centroid of the blue cursor fill (204,230,255), or None when
/// invisible. Proves frames are live (redraws present cursor moves).
fn vcursor_xy(shot: &PathBuf) -> Option<(i32, i32)> {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fill-vcurxy-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\nW,H = img.size\npx = img.load()\nC=(204,230,255)\nxs=[x for y in range(60,500) for x in range(0,W) if px[x,y]==C]\nys=[y for y in range(60,500) for x in range(0,W) if px[x,y]==C]\nprint(f'{len(xs)} {(sum(xs)//len(xs)) if xs else -1} {(sum(ys)//len(ys)) if ys else -1}')\n",
    )
    .expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(shot)
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let _ = std::fs::remove_file(shot);
    let text = String::from_utf8_lossy(&out.stdout).trim().to_string();
    let p: Vec<i32> = text.split_whitespace().map(|s| s.parse().unwrap_or(-99)).collect();
    assert_eq!(p.len(), 3, "bad analyzer output: {text:?}");
    if p[0] < 50 {
        return None;
    }
    Some((p[1], p[2]))
}

/// In-viewport arrows must not move the vertical thumb: it tracks the
/// viewport (like Excel), not the cursor. Regression: cursor-tracking
/// slid the thumb on every step although the view never scrolled.
/// (The cursor-centroid gate guards a vacuous pass on frozen frames:
/// the highlight must demonstrably reach row 3 while the thumb sits.)
#[test]
fn gui_viewport_thumb_static_on_inviewport_arrows() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-fill-vpstatic-{}-{}.corro", std::process::id(), id));
    std::fs::write(&path, "CORRO_LOG 1\nSET A1 a\nSET A2 b\nSET A3 c\nSET A4 d\nSET A5 e\nSET A6 f\n")
        .expect("write fixture");
    let mut child = spawn_gui(&path);
    let wid = find_window(child.id());
    xdotool(&["windowsize", &wid, "1200", "800"]);
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(800));
    let shot = screenshot(&wid, "vpstatic0");
    let (n0, min0, max0) = vthumb(&shot);
    assert!(n0 > 100, "v-thumb must render (px: {n0})");
    xdotool(&["key", "--window", &wid, "Down"]);
    std::thread::sleep(Duration::from_millis(300));
    xdotool(&["key", "--window", &wid, "Down"]);
    // Gate: highlight reaches row 3 (centroid y 108..136), proving live
    // frames and delivered keys — a frozen frame would trivially agree.
    let deadline = Instant::now() + Duration::from_secs(12);
    loop {
        let shot = screenshot(&wid, "vpstaticcur");
        if let Some((_, cy)) = vcursor_xy(&shot) {
            if (108..=136).contains(&cy) {
                break;
            }
        }
        if Instant::now() > deadline {
            panic!("cursor highlight never reached row 3 after Down x2");
        }
        std::thread::sleep(Duration::from_millis(300));
    }
    // Settle: two reads 500ms apart must agree (no mid-repaint tear).
    let shot = screenshot(&wid, "vpstatic1");
    let t1 = vthumb(&shot);
    std::thread::sleep(Duration::from_millis(500));
    let shot = screenshot(&wid, "vpstatic2");
    let t2 = vthumb(&shot);
    assert_eq!(t1, t2, "thumb reads must settle (live frames, no tear): {t1:?} vs {t2:?}");
    assert_eq!(
        (t1.1, t1.2), (min0, max0),
        "thumb must not move on in-viewport arrows (viewport static): {min0}..{max0} -> {:?}",
        (t1.1, t1.2)
    );
    let _ = child.kill();
    let _ = child.wait();
    let _ = std::fs::remove_file(&path);
}

/// Scrolling the viewport must move the vertical thumb with it (it tracks
/// the viewport origin). Companion to the static test: the thumb is live,
/// not nailed on.
#[test]
fn gui_viewport_thumb_follows_scrolled_viewport() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-fill-vpfollow-{}-{}.corro", std::process::id(), id));
    std::fs::write(&path, "CORRO_LOG 1\nSET A1 a\nSET A2 b\nSET A3 c\nSET A4 d\nSET A5 e\nSET A6 f\n")
        .expect("write fixture");
    let mut child = spawn_gui(&path);
    let wid = find_window(child.id());
    xdotool(&["windowsize", &wid, "1200", "800"]);
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(800));
    let shot = screenshot(&wid, "vpfollow0");
    let (n0, min0, _) = vthumb(&shot);
    assert!(n0 > 100, "v-thumb must render (px: {n0})");
    for _ in 0..42 {
        xdotool(&["key", "--window", &wid, "Down"]);
        std::thread::sleep(Duration::from_millis(120));
    }
    // Poll for the thumb to ride down with the viewport (frames lag keys
    // under load; deadline, not sleep).
    let deadline = Instant::now() + Duration::from_secs(20);
    let (n1, min1, max1) = loop {
        let shot = screenshot(&wid, "vpfollow");
        let t = vthumb(&shot);
        if t.1 - min0 > 30 {
            break t;
        }
        if Instant::now() > deadline {
            panic!("thumb never followed the scrolled viewport (top {min0} -> {})", t.1);
        }
        std::thread::sleep(Duration::from_millis(500));
    };
    assert!(
        (n1 - n0).abs() < n0 / 3,
        "thumb size must stay plausible while following (px {n0} -> {n1})"
    );
    assert!(
        max1 < 740,
        "thumb must track the viewport origin, not slam to the bottom (bottom {max1})"
    );
    let _ = child.kill();
    let _ = child.wait();
    let _ = std::fs::remove_file(&path);
}
fn cursor_x0(shot: &PathBuf) -> Option<i32> {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fill-curx-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\nW,H = img.size\npx = img.load()\nxs=[x for y in range(60,200) for x in range(0,W) if px[x,y]==(204,230,255)]\nprint(f'{len(xs)} {(min(xs) if xs else -1)}')\n",
    )
    .expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(shot)
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let _ = std::fs::remove_file(shot);
    let text = String::from_utf8_lossy(&out.stdout).trim().to_string();
    let p: Vec<i32> = text.split_whitespace().map(|s| s.parse().unwrap_or(-99)).collect();
    assert_eq!(p.len(), 2, "bad analyzer output: {text:?}");
    if p[0] < 50 {
        return None;
    }
    Some(p[1])
}

/// Dragging the horizontal thumb while the cursor sits in the left margin
/// must neither snap the thumb back nor move the selection: the thumb has
/// nothing new to show (sync pushes only on genuine change), and the
/// margin address is outside the thumb's expressible domain.
/// Regression: every draw re-pushed hv=0, so post-mouseup draws snapped a
/// dragged thumb back to the left edge.
#[test]
fn gui_margin_drag_thumb_does_not_snap_back() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-fill-drag-{}-{}.corro", std::process::id(), id));
    std::fs::write(&path, "CORRO_LOG 1\n").expect("write empty fixture");
    let mut child = spawn_gui(&path);
    let wid = find_window(child.id());
    xdotool(&["windowsize", &wid, "1200", "800"]);
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(500));
    // Into the margin (delivery gate: poll for the highlight to leave A1).
    xdotool(&["key", "--window", &wid, "Left"]);
    let deadline = Instant::now() + Duration::from_secs(12);
    loop {
        let shot = screenshot(&wid, "dragcur");
        if let Some(x0) = cursor_x0(&shot) {
            if x0 < 80 {
                break;
            }
        }
        if Instant::now() > deadline {
            panic!("cursor highlight never reached the margin after Left");
        }
        std::thread::sleep(Duration::from_millis(300));
    }
    // Thumb before: must render (thousands of px on an empty doc).
    let shot = screenshot(&wid, "dragbefore");
    let (n0, min0, max0) = hthumb(&shot);
    assert!(n0 > 1000, "h-thumb must render (px: {n0})");
    let cx0 = (min0 + max0) / 2;
    // Drag +150px and release away (un-hover restores normal colors).
    xdotool(&["mousemove", "--window", &wid, &cx0.to_string(), "776", "mousedown", "1"]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["mousemove", "--window", &wid, &(cx0 + 150).to_string(), "776"]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["mouseup", "1", "mousemove", "--window", &wid, "300", "300"]);
    std::thread::sleep(Duration::from_millis(1000));
    let shot = screenshot(&wid, "dragafter");
    let (n1, min1, max1) = hthumb(&shot);
    let cx1 = (min1 + max1) / 2;
    assert!(
        (n1 - n0).abs() < n0 / 7,
        "drag must not resize the thumb (px {n0} -> {n1})"
    );
    assert!(
        cx1 - cx0 > 60,
        "thumb must follow the drag, not snap back (center {cx0} -> {cx1})"
    );
    assert!(
        min1 > 50,
        "thumb must not snap back to the left edge (min x {min1})"
    );
    // Selection still in the margin (drag must not move it: {x} vs A1 87+).
    let shot = screenshot(&wid, "dragcur2");
    let x0 = cursor_x0(&shot).expect("cursor highlight must still render");
    assert!(
        x0 < 80,
        "margin selection must survive the thumb drag (cursor x0 {x0})"
    );
    let _ = child.kill();
    let _ = child.wait();
    let _ = std::fs::remove_file(&path);
}
