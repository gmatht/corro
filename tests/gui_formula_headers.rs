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
    let _ = std::fs::remove_file(&xwd);
    png
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

/// Address label must show the margin name after Left (cursor in `[A`
/// column). Expected text is computed with the same `cell_ref_text` the
/// ratatui reference uses, on a 1x1 empty grid.
#[test]
fn gui_formula_bar_shows_margin_name() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture("margin");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(500));
    xdotool(&["key", "--window", &wid, "Left"]);
    let expected = corro::addr::cell_ref_text(
        &corro::addr::sheet_cursor_to_addr(
            corro::addr::LogicalRow(corro::grid::HEADER_ROWS),
            corro::addr::GlobalCol(corro::grid::MARGIN_COLS - 1),
            corro::addr::MainRows(1),
            corro::addr::MainCols(1),
        ),
        1,
    );
    assert!(
        expected.starts_with('['),
        "test bug: expected a margin label, got {expected:?}"
    );
    wait_addr_label(&wid, "margin", &expected);
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
