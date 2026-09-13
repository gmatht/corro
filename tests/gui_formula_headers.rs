//! Formula-bar address label + header presentation (GTK3).
//!
//! 1. The address left of `fx` must show the user-facing cell name (`A1`,
//!    `[A1`, `A~1`, ... — `addr::cell_ref_text`, like the ratatui reference),
//!    not `CellAddr`'s internal rendering (`(0, 0)`, `<701>(0)`, ...).
//!    Regression: `update_formula_bar` formatted the address with
//!    `to_string()` (internal repr) instead of `cell_ref_text`.
//! 2. Body row/column gutter headers render bold (ratatui renders all column
//!    headers and the active/footer row labels bold).
//! 3. Margin-zone data cells render at 75% background brightness.
//! 4. Short (<=2 char, non-empty) gutter headers carry a padlock affordance;
//!    clicking toggles the row/column pin (stays visible while scrolling).
//!
//! Method: screenshot + tesseract OCR for the address label (exact match
//! against `cell_ref_text` computed in-test); pixel analysis for style and
//! padlocks. Verdicts poll with deadlines, never fixed sleeps.
//!
//! Requires: Linux, an X server, xdotool, xwd, convert (ImageMagick),
//! python3+PIL, tesseract. Run with:
//!   xvfb-run -a cargo test --features gtk --test gui_formula_headers
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

fn screenshot(wid: &str, tag: &str) -> PathBuf {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let xwd = std::env::temp_dir().join(format!("corro-fhdr-{tag}-{id}.xwd"));
    let png = std::env::temp_dir().join(format!("corro-fhdr-{tag}-{id}.png"));
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
    let _ = std::fs::remove_file(&xwd);    png
}

/// OCR the formula-bar address label (top-left, left of `fx`). Crops just
/// the label strip (48px — calibrated: wider crops catch the `fx` caption)
/// upscales 3x, and reads one line with an address whitelist (uppercase +
/// digits + symbols; address labels never contain lowercase).
fn ocr_addr_label(png: &PathBuf) -> String {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let crop = std::env::temp_dir().join(format!("corro-fhdr-crop-{id}.png"));
    Command::new("convert")
        .arg(png)
        .args(["-crop", "48x34+0+20", "-resize", "300%"])
        .arg(&crop)
        .status()
        .expect("convert crop");
    let out = Command::new("tesseract")
        .arg(&crop)
        .arg("stdout")
        .args(["--psm", "7", "-c", "tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ[]()~,0123456789_ "])
        .output()
        .expect("tesseract");
    let _ = std::fs::remove_file(&crop);
    String::from_utf8_lossy(&out.stdout)
        .chars()
        .filter(|c| !c.is_whitespace())
        .collect()
}

/// Poll OCR of the address label until it reads `expected` (deadline, not
/// sleep): proves the label actually updated, not just rendered once.
fn wait_addr_label(wid: &str, tag: &str, expected: &str) {
    let deadline = Instant::now() + Duration::from_secs(12);
    loop {
        let shot = screenshot(wid, tag);
        let text = ocr_addr_label(&shot);
        let _ = std::fs::remove_file(&shot);
        if text == expected {
            return;
        }
        if Instant::now() > deadline {
            panic!(
                "formula-bar address never read {expected:?} (last OCR: {text:?}); \
                 internal-repr rendering like `(0, 0)` fails this test"
            );
        }
        std::thread::sleep(Duration::from_millis(300));
    }
}

/// Centroid of the TOPMOST unlocked-padlock slate cluster in a screen rect:
/// the click target for that padlock. Scans top-down and averages only the
/// first icon's band (padlocks are ~12px tall with gaps between rows), so
/// the result is one padlock, not the centroid of all of them. Panics
/// (loudly, not silently) when no padlock pixels exist there.
fn find_slate_centroid(wid: &str, x0: i32, x1: i32, y0: i32, y1: i32) -> (i32, i32) {
    // Settle poll: a fresh app may not have rendered padlocks on the first
    // screenshots; verdicts stay pixel-based, only the wait is timed.
    for _ in 0..16 {
    let shot = screenshot(wid, "slatefind");
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fhdr-slate-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\npx = img.load()\nSLATE=(115,115,115)\nx0,x1,y0,y1 = map(int, sys.argv[2:6])\nysorted=sorted(y for y in range(y0,y1) for x in range(x0,x1) if px[x,y]==SLATE)\ny0top=ysorted[0] if ysorted else y1\nxs=[x for y in range(y0top,min(y0top+16,y1)) for x in range(x0,x1) if px[x,y]==SLATE]\nys=[y for y in range(y0top,min(y0top+16,y1)) for x in range(x0,x1) if px[x,y]==SLATE]\nprint(f'{len(xs)} {sum(xs)//max(1,len(xs))} {sum(ys)//max(1,len(ys))}')\n",
    )
    .expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(&shot)
        .args([&x0.to_string(), &x1.to_string(), &y0.to_string(), &y1.to_string()])
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let _ = std::fs::remove_file(&shot);
    let text = String::from_utf8_lossy(&out.stdout).trim().to_string();
    let p: Vec<i32> = text.split_whitespace().map(|s| s.parse().unwrap_or(-99)).collect();
    if p.len() == 3 && p[0] > 5 {
        return (p[1], p[2]);
    }
    std::thread::sleep(Duration::from_millis(500));
    }
    panic!(
        "no padlock pixels in rect ({x0},{y0})-({x1},{y1}) after settle poll; padlock missing where test expects one"
    );
}

/// Centroid of the single LOCKED-padlock dark cluster in a screen rect:
/// the click target for unpinning. After scrolling with a pin engaged, the
/// pinned row's padlock is the only locked one on screen, so it is
/// unambiguous; unlocked-slate search would find nothing (pinned icon is
/// dark, scrolled rows have long labels without padlocks). Settle-polls
/// like the slate finders, panics loudly after the deadline.
fn find_locked_centroid(wid: &str, x0: i32, x1: i32, y0: i32, y1: i32) -> (i32, i32) {
    for _ in 0..16 {
    let shot = screenshot(wid, "lockedfind");
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fhdr-lockedfind-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\npx = img.load()\nDARK=(51,51,51)\nx0,x1,y0,y1 = map(int, sys.argv[2:6])\nxs=[x for y in range(y0,y1) for x in range(x0,x1) if px[x,y]==DARK]\nys=[y for y in range(y0,y1) for x in range(x0,x1) if px[x,y]==DARK]\nprint(f'{len(xs)} {sum(xs)//max(1,len(xs))} {sum(ys)//max(1,len(ys))}')\n",
    )
    .expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(&shot)
        .args([&x0.to_string(), &x1.to_string(), &y0.to_string(), &y1.to_string()])
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let _ = std::fs::remove_file(&shot);
    let text = String::from_utf8_lossy(&out.stdout).trim().to_string();
    let p: Vec<i32> = text.split_whitespace().map(|s| s.parse().unwrap_or(-99)).collect();
    if p.len() == 3 && p[0] > 5 {
        return (p[1], p[2]);
    }
    std::thread::sleep(Duration::from_millis(500));
    }
    panic!(
        "no locked padlock in rect ({x0},{y0})-({x1},{y1}) after settle poll; pin did not engage?"
    );
}

/// True when dark-slate (locked padlock, 51,51,51) pixels appear near a
/// point: proves a click actually toggled a padlock to locked (as opposed
/// to clicking empty chrome, which changes nothing).
fn has_locked_pixel_near(wid: &str, x: i32, y: i32) -> bool {
    let shot = screenshot(wid, "lockednear");
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fhdr-locked-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\npx = img.load()\nDARK=(51,51,51)\nx,y = map(int, sys.argv[2:4])\nprint(sum(1 for yy in range(max(0,y-10), y+11) for xx in range(max(0,x-10), x+11) if px[xx,yy]==DARK))\n",
    )
    .expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(&shot)
        .args([&x.to_string(), &y.to_string()])
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let _ = std::fs::remove_file(&shot);
    String::from_utf8_lossy(&out.stdout).trim().parse().unwrap_or(0) > 3
}

/// Centroid of the LEFTMOST unlocked-padlock slate cluster in a screen
/// rect: the click target for that padlock in a horizontal strip (column
/// headers all share one y-band, so topmost-averaging would merge every
/// padlock; leftmost isolates the first column's). Panics loudly when no
/// padlock pixels exist there.
fn find_slate_leftmost(wid: &str, x0: i32, x1: i32, y0: i32, y1: i32) -> (i32, i32) {
    // Settle poll, like find_slate_centroid above.
    for _ in 0..16 {
    let shot = screenshot(wid, "slateleft");
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fhdr-slateL-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\npx = img.load()\nSLATE=(115,115,115)\nx0,x1,y0,y1 = map(int, sys.argv[2:6])\nxsl=sorted(x for y in range(y0,y1) for x in range(x0,x1) if px[x,y]==SLATE)\nx0l=xsl[0] if xsl else x1\nxs=[x for y in range(y0,y1) for x in range(x0l,min(x0l+14,x1)) if px[x,y]==SLATE]\nys=[y for y in range(y0,y1) for x in range(x0l,min(x0l+14,x1)) if px[x,y]==SLATE]\nprint(f'{len(xs)} {sum(xs)//max(1,len(xs))} {sum(ys)//max(1,len(ys))}')\n",
    )
    .expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(&shot)
        .args([&x0.to_string(), &x1.to_string(), &y0.to_string(), &y1.to_string()])
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let _ = std::fs::remove_file(&shot);
    let text = String::from_utf8_lossy(&out.stdout).trim().to_string();
    let p: Vec<i32> = text.split_whitespace().map(|s| s.parse().unwrap_or(-99)).collect();
    if p.len() == 3 && p[0] > 5 {
        return (p[1], p[2]);
    }
    std::thread::sleep(Duration::from_millis(500));
    }
    panic!(
        "no padlock pixels in rect ({x0},{y0})-({x1},{y1}) after settle poll; padlock missing where test expects one"
    );
}

/// OCR all visible grid text (grid band cropped, upscaled). Binarize with
/// gamma-preserving `-colorspace Gray` (NOT `-grayscale Rec709Luminance`,
/// which linearizes margin gray 191 down to 132, flipping dimmed margins
/// black and fusing PINME into a black bar). Crop starts below the header
/// strip so header text/padlocks never join the text block; `+repage`
/// clears the virtual canvas so later ops register correctly.
fn ocr_grid_text(wid: &str) -> String {
    let shot = screenshot(wid, "gridtext");
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let crop = std::env::temp_dir().join(format!("corro-fhdr-grid-{id}.png"));
    Command::new("convert")
        .arg(&shot)
        .args(["-crop", "500x348+50+56", "+repage", "-resize", "200%", "-colorspace", "Gray", "-threshold", "60%"])
        .arg(&crop)
        .status()
        .expect("convert crop");
    let _ = std::fs::remove_file(&shot);
    let out = Command::new("tesseract")
        .arg(&crop)
        .arg("stdout")
        .args(["--psm", "6"])
        .output()
        .expect("tesseract");
    let _ = std::fs::remove_file(&crop);
    String::from_utf8_lossy(&out.stdout).to_string()
}

/// Grid OCR with a settle poll: screenshots can catch a partial redraw
/// (empty text) right after a key burst; poll for non-empty text with a
/// deadline, then let the content assertions (not timing) deliver the
/// verdict.
fn ocr_grid_text_settled(wid: &str) -> String {
    let mut last = String::new();
    for _ in 0..10 {
        last = ocr_grid_text(wid);
        if !last.trim().is_empty() {
            return last;
        }
        std::thread::sleep(Duration::from_millis(500));
    }
    last
}

/// OCR the column-header strip (top band, upscaled).
fn ocr_header_strip(wid: &str) -> String {
    let shot = screenshot(wid, "headstrip");
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let crop = std::env::temp_dir().join(format!("corro-fhdr-strip-{id}.png"));
    Command::new("convert")
        .arg(&shot)
        .args(["-crop", "760x24+0+48", "-resize", "300%"])
        .arg(&crop)
        .status()
        .expect("convert crop");
    let _ = std::fs::remove_file(&shot);
    let out = Command::new("tesseract")
        .arg(&crop)
        .arg("stdout")
        .args(["--psm", "6"])
        .output()
        .expect("tesseract");
    let _ = std::fs::remove_file(&crop);
    String::from_utf8_lossy(&out.stdout).to_string()
}

fn fresh_fixture(tag: &str) -> PathBuf {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path =
        std::env::temp_dir().join(format!("corro-fhdr-{tag}-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, "CORRO_LOG 1\n").expect("write fixture");
    path
}

/// Address label must show the user-facing name `A2` after Down (not the
/// internal `(1, 0)`). Down (not startup) forces a real label update, so a
/// stale placeholder cannot pass.
#[test]
fn gui_formula_bar_shows_cell_name() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture("name");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "Down"]);
    wait_addr_label(&wid, "name", "A2");
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}

/// Left from A1 must highlight margin column [A and repaint the formula
/// bar for the margin address. Pixel ground truth (cursor bbox + label
/// repaint), because OCR of short labels confuses 1/L and drops brackets.
/// Cursor bounding box of the blue cursor fill (204,230,255) in the grid
/// band: (x0, y0, x1, y1), or None when no fill is visible. Pixel ground
/// truth for where the highlight sits (OCR of short labels confuses
/// 1/L and drops brackets, so text reads cannot distinguish A1 from [A1).
fn cursor_bbox(shot: &std::path::PathBuf) -> Option<(i32, i32, i32, i32)> {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fhdr-curbox-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\nW,H = img.size\npx = img.load()\nC=(204,230,255)\nxs=[x for y in range(60,500) for x in range(0,W) if px[x,y]==C]\nys=[y for y in range(60,500) for x in range(0,W) if px[x,y]==C]\nprint(f'{len(xs)} {(min(xs) if xs else -1)} {(min(ys) if ys else -1)} {(max(xs) if xs else -1)} {(max(ys) if ys else -1)}')\n",
    )
    .expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(shot)
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let text = String::from_utf8_lossy(&out.stdout).trim().to_string();
    let p: Vec<i32> = text.split_whitespace().map(|s| s.parse().unwrap_or(-99)).collect();
    assert_eq!(p.len(), 5, "bad analyzer output: {text:?}");
    if p[0] < 50 {
        return None;
    }
    Some((p[1], p[2], p[3], p[4]))
}

/// Dark-pixel count difference of the formula address zone (x0..52,y14..42)
/// between two screenshots: proves the label repainted (a margin address
/// adds a bracket and shifts glyphs, tens of pixels).
fn label_zone_diff(a: &std::path::PathBuf, b: &std::path::PathBuf) -> i32 {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fhdr-labeld-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nA = Image.open(sys.argv[1]).convert('RGB')\nB = Image.open(sys.argv[2]).convert('RGB')\npa, pb = A.load(), B.load()\ndef dark(p):\n    r,g,b = p\n    return r < 110 and g < 110 and b < 110\nn = sum(1 for y in range(14, 42) for x in range(0, 52) if dark(pa[x,y]) != dark(pb[x,y]))\nprint(n)\n",
    )
    .expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(a)
        .arg(b)
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    String::from_utf8_lossy(&out.stdout).trim().parse().unwrap_or(-99)
}

/// Dark-pixel (ink) count of the formula address zone (x0..64,y14..42).
/// `"A1"` renders two glyphs, `"[A1"` three: the ink count must grow when
/// the bar switches to the margin address (a repaint alone could redraw
/// identical text — growth proves the displayed address changed).
fn label_zone_ink(shot: &std::path::PathBuf) -> i32 {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fhdr-labelink-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\npx = img.load()\nn = sum(1 for y in range(14, 42) for x in range(0, 64) if px[x,y][0] < 110 and px[x,y][1] < 110 and px[x,y][2] < 110)\nprint(n)\n",
    )
    .expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(shot)
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    String::from_utf8_lossy(&out.stdout).trim().parse().unwrap_or(-99)
}

/// Spawn the app on a new (empty) document and activate its window.
/// Returns the child guard, window id, and fixture path.
/// Callers must hold GUI_LOCK for the whole test (this helper does not).
fn start_new_doc_app(tag: &str) -> (KillOnDrop, String, PathBuf) {
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture(tag);
    let child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(500));
    (child, wid, path)
}

#[test]
fn gui_new_doc_left_updates_formula_bar() {
    // New document, press Left: the formula bar must display the margin
    // cell. Pixel proof (repaint + ink growth for the added bracket),
    // because short-label OCR confuses 1/L and drops brackets, reading
    // both A1 and [A1 as "AL".
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    let (mut child, wid, path) = start_new_doc_app("formulamargin");
    let before_shot = screenshot(&wid, "fbar0");
    let ink_before = label_zone_ink(&before_shot);
    assert!(ink_before > 10, "address label must render ink (got {ink_before})");
    xdotool(&["key", "--window", &wid, "Left"]);
    // Gate on the highlight moving (the bar follows the cursor): poll for
    // any bbox change, then prove the bar's displayed text changed too.
    let deadline = Instant::now() + Duration::from_secs(12);
    let after_shot = loop {
        let shot = screenshot(&wid, "fbarmove");
        let bb = cursor_bbox(&shot);
        let moved = bb.map(|b| (b.0, b.1)) != cursor_bbox(&before_shot).map(|b| (b.0, b.1));
        if moved && bb.is_some() {
            break shot;
        }
        let _ = std::fs::remove_file(&shot);
        if Instant::now() > deadline {
            panic!("cursor highlight never moved after Left");
        }
        std::thread::sleep(Duration::from_millis(300));
    };
    let repainted = label_zone_diff(&before_shot, &after_shot);
    assert!(
        repainted > 30,
        "formula bar must repaint for the margin address (diff px: {repainted})"
    );
    // The new label must carry more ink ("[A1" gains a bracket over "A1").
    // Poll for a stable grown reading: a screenshot can catch a torn frame
    // mid-repaint under load, but the settled label persists.
    let deadline = Instant::now() + Duration::from_secs(8);
    let ink_after = loop {
        let shot = screenshot(&wid, "fbarink");
        let ink = label_zone_ink(&shot);
        let _ = std::fs::remove_file(&shot);
        if ink as f64 > ink_before as f64 * 1.1 {
            break ink;
        }
        if Instant::now() > deadline {
            break ink;
        }
        std::thread::sleep(Duration::from_millis(400));
    };
    assert!(
        ink_after as f64 > ink_before as f64 * 1.1,
        "formula bar must show a longer address ([A1 gains a bracket): ink {ink_before} -> {ink_after}"
    );
    let _ = std::fs::remove_file(&before_shot);
    let _ = std::fs::remove_file(&after_shot);
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}

#[test]
fn gui_new_doc_left_moves_blue_selection() {
    // New document, press Left: the blue highlight itself must move from
    // A1 into margin column [A (left edge at the gutter's right edge,
    // x=50=ROW_LABEL_W, fill inset ~2px hence 50..=58; same row). A
    // double-move would land a full column further right (empty-fixture
    // margins are ~28px wide); no move would still overlap `before`.
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    let (mut child, wid, path) = start_new_doc_app("bluemargin");
    // Cursor bbox before the key (A1): the gutter ends at x=50, so the
    // margin column [A starts there; col A starts a column-width right.
    let before_shot = screenshot(&wid, "blue0");
    let before = cursor_bbox(&before_shot)
        .expect("cursor highlight must be visible before Left");
    let _ = std::fs::remove_file(&before_shot);
    xdotool(&["key", "--window", &wid, "Left"]);
    // Poll for the highlight to move (deadline, not sleep): event delivery
    // through the present-pump is timing-dependent, so pixel position (not
    // text) is the verdict.
    let deadline = Instant::now() + Duration::from_secs(12);
    let after = loop {
        let shot = screenshot(&wid, "bluemove");
        if let Some(b) = cursor_bbox(&shot) {
            if (b.0 - before.0).abs() > 5 || (b.1 - before.1).abs() > 5 {
                let _ = std::fs::remove_file(&shot);
                break b;
            }
        }
        let _ = std::fs::remove_file(&shot);
        if Instant::now() > deadline {
            panic!(
                "blue selection never moved after Left (stayed {before:?})"
            );
        }
        std::thread::sleep(Duration::from_millis(300));
    };
    assert!(
        (50..=58).contains(&after.0),
        "Left from A1 must land the blue selection in margin column [A (x0 in 50..=58, got {after:?}, was {before:?})"
    );
    assert!(
        (after.1 - before.1).abs() <= 6,
        "Left must not change rows (y0 {} vs {})",
        after.1,
        before.1
    );
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}

/// Address label must show the header name after Up (cursor in `~1` row).
#[test]
fn gui_formula_bar_shows_header_name() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture("header");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "Up"]);
    let expected = corro::addr::cell_ref_text(
        &corro::addr::sheet_cursor_to_addr(
            corro::addr::LogicalRow(corro::grid::HEADER_ROWS - 1),
            corro::addr::GlobalCol(corro::grid::MARGIN_COLS),
            corro::addr::MainRows(1),
            corro::addr::MainCols(1),
        ),
        1,
    );
    assert!(
        expected.contains('~'),
        "test bug: expected a header label, got {expected:?}"
    );
    wait_addr_label(&wid, "header", &expected);
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}

/// Margin-zone data cells must render at 75% background brightness (exact
/// 191 gray) while the body stays white. Measures whole-grid-band fractions
/// on the populated overflow fixture: margins dominate the area, so gray
/// must be substantial but not total (which would mean dimming leaked into
/// main cells).
#[test]
fn gui_margin_cells_dimmer_than_body() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture("dim");
    std::fs::write(&path, "CORRO_LOG 1\nSET A1 Hello World!\nSET A2 x\n").expect("seed fixture");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(800));
    let shot = screenshot(&wid, "dim");
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fhdr-dim-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\nW,H = img.size\npx = img.load()\nGRAY=(191,191,191)\nWHITE=(255,255,255)\n# whole grid band below the 24px header strip\ntot=gray=white=0\nfor y in range(30, 750, 3):\n    row=px\n    for x in range(50, 1150, 3):\n        tot+=1\n        if px[x,y]==GRAY: gray+=1\n        elif px[x,y]==WHITE: white+=1\nprint(f'{gray} {white} {tot}')\n",
    )
    .expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(&shot)
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let _ = std::fs::remove_file(&shot);
    let text = String::from_utf8_lossy(&out.stdout).trim().to_string();
    let p: Vec<i32> = text.split_whitespace().map(|s| s.parse().unwrap_or(-99)).collect();
    assert_eq!(p.len(), 3, "bad analyzer output: {text:?}");
    let (gray, tot) = (p[0] as f64, p[2].max(1) as f64);
    assert!(
        gray / tot > 0.3,
        "margin zones should render mostly 75%-gray (gray fraction {}/{tot}); margins render full-white",
        p[0]
    );
    assert!(
        gray / tot < 0.95,
        "main cells must stay white (gray fraction {}/{tot}); dimming leaked into body content",
        p[0]
    );
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}

/// Body gutter headers render their labels in place (column strip on top,
/// row gutter at left). Boldness itself is pinned deterministically by the
/// `gutter_tests` unit tests (RecordingDrawContext asserts weight == 1 on
/// every gutter label op); pixel-ink comparison cannot prove boldness live
/// because grid glyphs render larger than gutter glyphs at this scale.
/// This test proves the gutter paint path executes end-to-end in the live
/// app: if header painting regresses (wrong strip, missing labels), it fails.
#[test]
fn gui_body_headers_render_labels() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture("headers");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(800));
    let text = ocr_header_strip(&wid);
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
    assert!(
        text.contains('A'),
        "column-header strip should show the A label, OCR got: {text:?}"
    );
    assert!(
        text.contains("[A") || text.contains("]A"),
        "column-header strip should show a margin label, OCR got: {text:?}"
    );
}

/// Short gutter headers carry padlock affordances: exact slate pixels must
/// appear in BOTH the row-label gutter and the column-header strip.
/// Unlocked padlocks paint slate-gray (115,115,115); nothing else in the UI
/// uses that color. Pre-fix: zero such pixels anywhere.
#[test]
fn gui_short_headers_show_padlocks() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture("locks");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(800));
    let shot = screenshot(&wid, "locks");
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fhdr-lock-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\nW,H = img.size\npx = img.load()\nSLATE=(115,115,115)\n# row-label gutter (x<50), grid band\nrow = sum(1 for y in range(70, 780) for x in range(0, 50) if px[x,y]==SLATE)\n# column-header strip (below menu+formula chrome, y 48..72), full width\ncol = sum(1 for y in range(48, 72, 1) for x in range(0, W, 2) if px[x,y]==SLATE)\nprint(f'{row} {col}')\n",
    )
    .expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(&shot)
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let _ = std::fs::remove_file(&shot);
    let text = String::from_utf8_lossy(&out.stdout).trim().to_string();
    let p: Vec<i32> = text.split_whitespace().map(|s| s.parse().unwrap_or(-99)).collect();
    assert_eq!(p.len(), 2, "bad analyzer output: {text:?}");
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
    assert!(
        p[0] > 100,
        "row gutter should show unlocked-padlock pixels (found {}); short headers lack padlocks",
        p[0]
    );
    assert!(
        p[1] > 100,
        "column header strip should show unlocked-padlock pixels (found {}); short headers lack padlocks",
        p[1]
    );
}

/// Clicking a row padlock pins the row: after scrolling far down, the pinned
/// row's content stays rendered while unpinned rows scroll away. Clicking
/// again unpins (row scrolls away normally). GONEBYE proves the viewport
/// really moved (guards a vacuous pass where scrolling never happened).
#[test]
fn gui_padlock_click_pins_row_visible() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path =
        std::env::temp_dir().join(format!("corro-fhdr-pin-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, "CORRO_LOG 1\nSET A1 PINME\nSET A5 GONEBYE\n").expect("write fixture");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(800));
    // Topmost padlock in the row gutter = display row 0 (fresh app starts
    // at A1). Clicking pins that row.
    let pad = find_slate_centroid(&wid, 0, 50, 70, 780);
    xdotool(&["mousemove", "--window", &wid, &pad.0.to_string(), &pad.1.to_string(), "click", "1"]);
    std::thread::sleep(Duration::from_millis(600));
    for _ in 0..50 {
        xdotool(&["key", "--window", &wid, "Down"]);
        std::thread::sleep(Duration::from_millis(120));
    }
    std::thread::sleep(Duration::from_millis(800));
    let grid = ocr_grid_text_settled(&wid);
    assert!(
        grid.contains("PINME"),
        "pinned row 1 must stay visible after scrolling (grid text: {grid:?})"
    );
    assert!(
        !grid.contains("GONEBYE"),
        "unpinned row 5 must scroll away (viewport never moved? grid text: {grid:?})"
    );
    // Click the LOCKED padlock to unlock (the only dark icon on screen;
    // unlocked-slate search would find nothing now). The pinned row renders
    // first, so it sits at the gutter top in either chrome variant.
    let pad2 = find_locked_centroid(&wid, 0, 50, 40, 130);
    xdotool(&["mousemove", "--window", &wid, &pad2.0.to_string(), &pad2.1.to_string(), "click", "1"]);
    std::thread::sleep(Duration::from_millis(600));
    std::thread::sleep(Duration::from_millis(800));
    let grid2 = ocr_grid_text_settled(&wid);
    assert!(
        !grid2.contains("PINME"),
        "unpinned row 1 must scroll away like any other row (grid text: {grid2:?})"
    );
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}

/// Clicking a column-header padlock pins the column: after scrolling far
/// left, the pinned column's locked padlock stays frozen at the same screen
/// spot while the formula bar proves the cursor moved on. Clicking again
/// unpins (the lock disappears). Pixel-exact padlock checks, not OCR words:
/// padlock glyphs contaminate header OCR with f/g misreads.
/// First margin column reached by keys commits as margin: build trailing
/// blanks with clicks (pointer growth chains), step onto the first margin
/// column with the keyboard (no NAV growth fires while 2+ blanks stand
/// past content), then commit. Must file exactly `SET ]A1 Q`, same as
/// ratatui — never data. The single Right carries the usual autorepeat
/// caveat.
#[test]
fn gui_ring_key_step_commits_margin() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture("ringkeystep");
    std::fs::write(&path, "CORRO_LOG 1\nSET A1 x\n").expect("seed fixture");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(500));
    // Chain clicks rightward to open blanks (each click on the last column
    // grows one more): B1, then C1.
    for x in [130, 165] {
        xdotool(&["mousemove", "--window", &wid, &x.to_string(), "82", "click", "1"]);
        std::thread::sleep(Duration::from_millis(600));
        xdotool(&["key", "--window", &wid, "Escape"]);
        std::thread::sleep(Duration::from_millis(300));
    }
    // Step right twice onto the ring (D1 main, then E1 ring — no NAV growth
    // with 2+ blanks standing).
    xdotool(&["key", "--window", &wid, "Right"]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "--window", &wid, "Right"]);
    // Gate: highlight reaches the E1 column zone (x0 > 200).
    let deadline = Instant::now() + Duration::from_secs(12);
    loop {
        let shot = screenshot(&wid, "ringkeycur");
        let far_enough = cursor_bbox(&shot).map(|b| b.0 > 200).unwrap_or(false);
        let _ = std::fs::remove_file(&shot);
        if far_enough {
            break;
        }
        if Instant::now() > deadline {
            panic!("cursor highlight never reached the ring column");
        }
        std::thread::sleep(Duration::from_millis(300));
    }
    xdotool(&["key", "--window", &wid, "Q"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "Return"]);
    let lines = wait_file_lines(&path, Instant::now() + Duration::from_secs(10));
    let _ = child.kill();
    let _ = child.wait();
    assert!(
        lines.iter().any(|l| l == "SET ]A1 Q"),
        "first margin column reached by keys must file as margin ]A1, same as ratatui (lines: {lines:?})"
    );
    assert!(
        !lines.iter().any(|l| l == "SET E1 Q"),
        "first margin column must never file as data E1 (lines: {lines:?})"
    );
    let _ = std::fs::remove_file(&path);
}

/// Poll the fixture file until a SET line lands (commit proof), or panic.
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
        std::thread::sleep(Duration::from_millis(200));
    }
}

/// True margins stay margins: clicking the right-margin column files a
/// `SET ]..` commit (never data), and the formula bar shows the longer
/// margin label. Guards over-correction of the ring-as-data change.
#[test]
fn gui_true_margin_stays_margin() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    let (mut child, wid, path) = start_new_doc_app("truemargin");
    let before_shot = screenshot(&wid, "truemargin0");
    let ink_before = label_zone_ink(&before_shot);
    let _ = std::fs::remove_file(&before_shot);
    // Click deep into the right-margin area (past the ring column).
    xdotool(&["mousemove", "--window", &wid, "300", "82", "click", "1"]);
    std::thread::sleep(Duration::from_millis(700));
    xdotool(&["key", "--window", &wid, "Escape"]);
    std::thread::sleep(Duration::from_millis(300));
    // Formula bar shows a longer (3-glyph) margin label.
    let deadline = Instant::now() + Duration::from_secs(8);
    let ink_after = loop {
        let shot = screenshot(&wid, "truemarginink");
        let ink = label_zone_ink(&shot);
        let _ = std::fs::remove_file(&shot);
        if ink as f64 > ink_before as f64 * 1.1 {
            break ink;
        }
        if Instant::now() > deadline {
            break ink;
        }
        std::thread::sleep(Duration::from_millis(400));
    };
    assert!(
        ink_after as f64 > ink_before as f64 * 1.1,
        "true margin must show a longer label than A1 (ink {ink_before} -> {ink_after})"
    );
    // ...and commits as margin (any ]X1), never as data.
    xdotool(&["key", "--window", &wid, "W"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "Return"]);
    let lines = wait_file_lines(&path, Instant::now() + Duration::from_secs(10));
    let _ = child.kill();
    let _ = child.wait();
    assert!(
        lines.iter().any(|l| l.starts_with("SET ]")),
        "true-margin click must file a margin commit (lines: {lines:?})"
    );
    assert!(
        !lines.iter().any(|l| l.starts_with("SET B") || l.starts_with("SET C")),
        "true-margin click must not misroute into data (lines: {lines:?})"
    );
    let _ = std::fs::remove_file(&path);
}
#[test]
fn gui_padlock_click_pins_column_visible() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path =
        std::env::temp_dir().join(format!("corro-fhdr-pincol-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, "CORRO_LOG 1\nSET A1 PINME\n").expect("write fixture");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(800));
    // Leftmost padlock in the header strip = first visible column's padlock.
    let pad = find_slate_leftmost(&wid, 0, 760, 40, 76);
    let formula0 = ocr_addr_label(&screenshot(&wid, "pincol0"));
    xdotool(&["mousemove", "--window", &wid, &pad.0.to_string(), &pad.1.to_string(), "click", "1"]);
    std::thread::sleep(Duration::from_millis(600));
    // The click must have toggled that padlock to locked (dark slate at the
    // click site); otherwise we clicked empty chrome and nothing is pinned.
    assert!(
        has_locked_pixel_near(&wid, pad.0, pad.1),
        "clicked padlock should render locked after click"
    );
    for _ in 0..40 {
        xdotool(&["key", "--window", &wid, "Left"]);
        std::thread::sleep(Duration::from_millis(120));
    }
    std::thread::sleep(Duration::from_millis(800));
    let formula1 = ocr_addr_label(&screenshot(&wid, "pincol1"));
    assert!(
        formula0 != formula1,
        "cursor must actually move for the scroll proof (formula {formula0:?} vs {formula1:?})"
    );
    // Pixel-exact frozen check (no OCR words: padlock glyphs contaminate
    // header OCR with f/g misreads): the locked padlock must still sit at
    // the same screen spot after scrolling — the pinned column never left.
    assert!(
        has_locked_pixel_near(&wid, pad.0, pad.1),
        "pinned column's locked padlock must stay frozen at ({},{}) after scrolling", pad.0, pad.1
    );
    // Click again to unlock (same spot: the frozen column renders first);
    // the locked padlock must disappear from the strip entirely.
    xdotool(&["mousemove", "--window", &wid, &pad.0.to_string(), &pad.1.to_string(), "click", "1"]);
    std::thread::sleep(Duration::from_millis(800));
    assert!(
        !has_locked_pixel_near(&wid, pad.0, pad.1),
        "unpinned column must release its lock (dark padlock still at ({},{}))", pad.0, pad.1
    );
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}

/// True footers past the ring keep footer addressing: Down x3 on an empty
/// sheet lands the first true footer row (ring row hr+1 is data A2 now) and
/// commits exactly `SET A_2 Q` — never data A3, never A1. Pins the footer
/// scheme intact beyond the ring (the ring consumed the old _1 slot).
#[test]
fn gui_footer_commit_past_ring() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    let (mut child, wid, path) = start_new_doc_app("footerpast");
    for _ in 0..3 {
        xdotool(&["key", "--window", &wid, "Down"]);
        std::thread::sleep(Duration::from_millis(300));
    }
    xdotool(&["key", "--window", &wid, "Q"]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "Return"]);
    let lines = wait_file_lines(&path, Instant::now() + Duration::from_secs(10));
    let _ = child.kill();
    let _ = child.wait();
    assert!(
        lines.iter().any(|l| l == "SET A_2 Q"),
        "first true footer must file as A_2 (lines: {lines:?})"
    );
    assert!(
        !lines.iter().any(|l| l == "SET A3 Q"),
        "true footer must not misroute into data A3 (lines: {lines:?})"
    );
    assert!(
        !lines.iter().any(|l| l == "SET A1 Q"),
        "keys must have moved (lines: {lines:?})"
    );
    let _ = std::fs::remove_file(&path);
}
