//! Trailing blank data row/column in the GUI (GTK3).
//!
//! The grid always shows one blank BODY row below (and one blank BODY column
//! right of) the last non-blank one, counting the cursor's row/column as
//! non-blank. Without this, the cells beyond the content are margins, so the
//! mouse can never select a fresh data cell. Keyboard navigation maintains
//! its own two content-counted blanks (NAV_BLANK_ROWS/COLS); this is the
//! GUI floor of one cursor-counted blank, visible even at load.
//!
//! - Fresh 2x2 content renders blank body row "3" and body column "C"
//!   (white, labeled, clickable) instead of footer/margin chrome there.
//! - Clicking the blank row selects a real data address (formula bar reads
//!   e.g. A3, no commit from the click alone) and opens one more blank
//!   beyond it (cursor counts as non-blank).
//!
//! Method (no fixed sleeps for verdicts, no shared files): pixel-structure
//! analysis (white body bands vs gray margins) plus formula-bar OCR, all
//! settled/polled with deadlines.
//!
//! Requires: Linux, an X server, xdotool, xwd, convert (ImageMagick),
//! python3+PIL, tesseract. Run with:
//!   xvfb-run -a cargo test --features gtk --test gui_trailing_blank
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
    let xwd = std::env::temp_dir().join(format!("corro-blank-{tag}-{id}.xwd"));
    let png = std::env::temp_dir().join(format!("corro-blank-{tag}-{id}.png"));
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

/// OCR the formula-bar address label (top-left, left of `fx`).
fn ocr_addr_label(png: &PathBuf) -> String {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let crop = std::env::temp_dir().join(format!("corro-blank-crop-{id}.png"));
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
/// sleep, delivers the verdict).
fn wait_addr_label(wid: &str, expected: &str, deadline: Instant) -> String {
    loop {
        let shot = screenshot(wid, "blankaddr");
        let text = ocr_addr_label(&shot);
        let _ = std::fs::remove_file(&shot);
        if text == expected || Instant::now() > deadline {
            return text;
        }
        std::thread::sleep(Duration::from_millis(300));
    }
}

/// Body census of a screenshot: white (255,255,255) runs of 10+ px in the
/// row band `y0..y1` (one per body column), and the number of body rows
/// down column `x` (white region height / 20px rows). Body cells paint
/// white (plus dark text); footer/right margins paint solid 191 gray, so
/// the white region's bottom edge (8 consecutive all-gray rows) marks the
/// last body row. Prints `colruns bodyrows`.
fn band_census(shot: &PathBuf, y0: i32, y1: i32, x: i32) -> (i32, i32, i32) {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-blank-census-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\nW,H = img.size\npx = img.load()\ny0,y1,x = map(int, sys.argv[2:5])\nWHITE=(255,255,255)\n# white runs in the row band (each body column contributes one)\nruns=0\nrun=0\nrunstart=0\nfirstcx=-1\nfor xx in range(50, min(700, W)):\n    colwhite = sum(1 for yy in range(y0, y1, 3) if px[xx,yy]==WHITE)\n    if colwhite * 3 >= (y1 - y0) - 4:\n        if run == 0:\n            runstart = xx\n        run += 1\n    else:\n        if run >= 10:\n            runs += 1\n            if firstcx < 0:\n                firstcx = (runstart + xx) // 2\n        run = 0\nif run >= 10:\n    runs += 1\n    if firstcx < 0:\n        firstcx = (runstart + min(700, W)) // 2\n# body rows: white-region height down column x (grid top onward)\ndef isgray(c):\n    r,g,b = c\n    return 150 <= r <= 235 and 150 <= g <= 235 and 150 <= b <= 235\n# grid top: first gutter ink (row numbers live only in grid rows; the\n# header-strip corner above them is blank). Glyphs start a couple px below\n# the row edge; rounding absorbs that. Robust across chrome variants.\ndef hasink(c):\n    return max(c) < 130\ntop = next((yy for yy in range(48, 220) if any(hasink(px[xx,yy]) for xx in range(30, 48))), 72)\nbottom = top\ngrayrun = 0\nfor yy in range(top, min(600, H)):\n    if all(isgray(px[xx,yy]) for xx in range(x-4, x+5)):\n        grayrun += 1\n        if grayrun >= 8:\n            bottom = yy - 7\n            break\n    else:\n        grayrun = 0\n        bottom = yy\nprint(f'{runs} {round((bottom - top) / 20)} {firstcx}')\n",
    )
    .expect("write analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(shot)
        .args([&y0.to_string(), &y1.to_string(), &x.to_string()])
        .output()
        .expect("python3 analyzer");
    let _ = std::fs::remove_file(&script);
    let _ = std::fs::remove_file(shot);
    let text = String::from_utf8_lossy(&out.stdout).trim().to_string();
    let p: Vec<i32> = text.split_whitespace().map(|s| s.parse().unwrap_or(-99)).collect();
    assert_eq!(p.len(), 3, "bad analyzer output: {text:?}");
    (p[0], p[1], p[2])
}

fn fresh_fixture(tag: &str, body: &str) -> PathBuf {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path =
        std::env::temp_dir().join(format!("corro-blank-{tag}-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, body).expect("write fixture");
    path
}

/// Fresh 2x2 content renders a blank BODY row 3 and BODY column C (white,
/// labeled, clickable) — not footer/margin chrome there. PINME widens
/// column A so the click target below stays inside it.
#[test]
fn gui_blank_row_and_col_render_as_body() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture("body", "CORRO_LOG 1\nSET A1 PINME\nSET B2 q\n");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(800));
    // Row 3 band (fresh unscrolled chrome: grid top 72, 20px rows).
    // Column A need not start at any fixed x (a leading margin of
    // content-adaptive width precedes it), so locate it dynamically: the
    // first white run's center, then census body rows down that column.
    let shot = screenshot(&wid, "bodybands");
    let (colruns, _, firstcx) = band_census(&shot, 114, 130, 80);
    assert_eq!(
        colruns, 3,
        "row 3 must show three white body columns A,B,C (runs: {colruns})"
    );
    assert!(
        firstcx > 50,
        "first body column must be right of the gutter (x: {firstcx})"
    );
    let shot = screenshot(&wid, "bodyrows");
    let (_, bodybands, _) = band_census(&shot, 114, 130, firstcx);
    assert_eq!(
        bodybands, 3,
        "column A must show three body rows 1,2,3 (rows: {bodybands})"
    );
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}

/// Clicking the blank body row selects a real data address (formula bar
/// reads A3, and the click alone commits nothing), and the cursor parking
/// there opens one more blank beyond it (row 4 renders as body).
#[test]
fn gui_click_blank_row_selects_data_cell() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture("clickblank", "CORRO_LOG 1\nSET A1 PINME\nSET B2 q\n");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(800));
    // Aim by the located first body column (row 3 band center).
    let shot = screenshot(&wid, "clickcol");
    let (_, _, firstcx) = band_census(&shot, 114, 130, 80);
    assert!(
        firstcx > 50,
        "first body column must be right of the gutter (x: {firstcx})"
    );
    xdotool(&["mousemove", "--window", &wid, &firstcx.to_string(), "122", "click", "1"]);
    let addr = wait_addr_label(&wid, "A3", Instant::now() + Duration::from_secs(10));
    assert_eq!(
        addr, "A3",
        "clicking blank row 3 must select data cell A3 (formula: {addr:?})"
    );
    // Cursor on row 3 counts as non-blank: row 4 opens beyond it.
    let shot = screenshot(&wid, "bodybands2");
    let (_, bodybands, _) = band_census(&shot, 114, 130, firstcx);
    assert_eq!(
        bodybands, 4,
        "cursor on row 3 must open body row 4 (rows: {bodybands})"
    );
    // The click alone commits nothing.
    let log = std::fs::read_to_string(&path).unwrap_or_default();
    assert!(
        !log.contains("SET A3"),
        "clicking must not commit a cell (log: {log:?})"
    );
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}
