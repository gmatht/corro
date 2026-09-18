//! Live GTK: headers covered by a selection use the body selection fill.
//!
//! The ratatui reference highlights covered row/column headers; the GTK
//! canvas must do the same (same fill the selected body cells use). This
//! drives a real window: pin the cursor to A1 (click + type + Enter both
//! proves the geometry and returns focus to the canvas), extend with
//! Shift+Right, then count selection-blue pixels in header-only regions.
//! No body cells may intrude into those regions, so the counts are about
//! chrome, not content.
//!
//! Requires: Linux, X server, xdotool, xwd + ImageMagick, python3 + PIL.
//! Run with: xvfb-run -a cargo test --features gtk --test gui_selected_headers
#![cfg(all(target_os = "linux", feature = "gui"))]

use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::time::{Duration, Instant};

static COUNTER: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

/// Live GUI tests share one X server and global input focus: serialize them.
static GUI_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

/// Canvas origin in window pixels (menubar + formula bar sit above the
/// canvas, gutter chrome at its left), calibrated by click→type→read-back:
/// (112, 82) addresses A1.
const CANVAS_X: i32 = 50;
const CANVAS_Y: i32 = 48;
/// Selection fill (0.9, 0.95, 1.0) as rendered by cairo (rounds halves up).
const SEL_RGB: (u8, u8, u8) = (230, 243, 255);

fn xdotool(args: &[&str]) -> String {
    String::from_utf8_lossy(
        &Command::new("xdotool").args(args).output().expect("xdotool failed").stdout,
    )
    .to_string()
}

struct KillOnDrop(Child);
impl std::ops::Deref for KillOnDrop {
    type Target = Child;
    fn deref(&self) -> &Child {
        &self.0
    }
}
impl std::ops::DerefMut for KillOnDrop {
    fn deref_mut(&mut self) -> &mut Child {
        &mut self.0
    }
}
impl Drop for KillOnDrop {
    fn drop(&mut self) {
        let _ = self.0.kill();
        let _ = self.0.wait();
    }
}

fn find_corro_window(pid: u32, deadline: Instant) -> String {
    let pid = pid.to_string();
    loop {
        for id in xdotool(&["search", "--pid", &pid]).split_whitespace() {
            let geo = xdotool(&["getwindowgeometry", "--shell", id]);
            let wide = geo
                .lines()
                .any(|l| l.strip_prefix("WIDTH=").and_then(|v| v.trim().parse::<i32>().ok()).map(|w| w > 100).unwrap_or(false));
            if wide {
                return id.to_string();
            }
        }
        if Instant::now() > deadline {
            panic!("timed out waiting for the corro window");
        }
        std::thread::sleep(Duration::from_millis(300));
    }
}

fn screenshot(wid: &str, tag: &str) -> PathBuf {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let xwd = std::env::temp_dir().join(format!("corro-selhead-{id}-{tag}.xwd"));
    let png = std::env::temp_dir().join(format!("corro-selhead-{id}-{tag}.png"));
    Command::new("xwd")
        .args(["-id", wid, "-out"])
        .arg(&xwd)
        .status()
        .expect("xwd");
    Command::new("convert")
        .arg(&xwd)
        .arg(&png)
        .status()
        .expect("convert");
    png
}

/// Count exact-match pixels of SEL_RGB in a window-relative rect.
fn count_sel(png: &PathBuf, x0: i32, y0: i32, x1: i32, y1: i32) -> u64 {
    let script = r#"
import sys
from PIL import Image
im = Image.open(sys.argv[1]).convert('RGB')
px = im.load()
w, h = im.size
want = (230, 243, 255)
n = 0
for y in range(max(0, int(sys.argv[3])), min(h, int(sys.argv[5]))):
    for x in range(max(0, int(sys.argv[2])), min(w, int(sys.argv[4]))):
        if px[x, y] == want:
            n += 1
print(n)
"#;
    let out = Command::new("python3")
        .arg("-c")
        .arg(script)
        .arg(png)
        .arg(x0.to_string())
        .arg(y0.to_string())
        .arg(x1.to_string())
        .arg(y1.to_string())
        .output()
        .expect("python3 failed");
    String::from_utf8_lossy(&out.stdout).trim().parse().expect("pixel count")
}

fn wait_file_has(path: &PathBuf, want: &str, deadline: Instant) {
    loop {
        if let Ok(content) = std::fs::read_to_string(path) {
            if content.lines().any(|l| l == want) {
                return;
            }
        }
        if Instant::now() > deadline {
            panic!(
                "timed out waiting for {want:?} in {}\ncontent:\n{}",
                path.display(),
                std::fs::read_to_string(path).unwrap_or_default()
            );
        }
        std::thread::sleep(Duration::from_millis(100));
    }
}

/// Click (single xdotool invocation: separate move/click calls deliver no
/// button event), then type text and commit with Return.
fn click(wid: &str, x: i32, y: i32) {
    xdotool(&[
        "mousemove", "--window", wid, &x.to_string(), &y.to_string(), "click", "1",
    ]);
}

#[test]
fn covered_headers_use_selection_fill() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(std::env::var("DISPLAY").is_ok(), "requires X (run under xvfb-run -a)");

    let dir = std::env::temp_dir().join(format!("corro-selhead-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let file = dir.join("wb.corro");
    std::fs::write(&file, "CORRO_LOG 1\n").unwrap();

    let bin = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("target/debug/corro");
    let child = KillOnDrop(
        Command::new(&bin)
            .arg("--gui")
            .arg(&file)
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .spawn()
            .expect("spawn corro --gui"),
    );
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(30));
    xdotool(&["windowactivate", "--sync", &wid]);
    xdotool(&["windowfocus", &wid]);
    std::thread::sleep(Duration::from_millis(800));

    // Pin the cursor to A1 AND return focus to the canvas: click (starts an
    // edit), type, commit (commit returns canvas focus). The file assertion
    // proves the geometry; without it a later header count would be
    // evidence about the wrong cells. (Commit moves the cursor down to A2,
    // mirroring the ratatui commit-and-move-down, so the selection below
    // covers row 2 — verified via the formula-bar address in a probe run.)
    click(&wid, CANVAS_X + 62, CANVAS_Y + 34);
    std::thread::sleep(Duration::from_millis(600));
    xdotool(&["key", "--window", &wid, "q"]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "--window", &wid, "Return"]);
    wait_file_has(&file, "SET A1 q", Instant::now() + Duration::from_secs(10));

    // Extend to A2:B2. Covered: row 2 gutter, col A/B headers.
    xdotool(&["key", "--window", &wid, "Shift+Right"]);
    std::thread::sleep(Duration::from_millis(600));
    let png = screenshot(&wid, "sel");

    // Gutter band of the covered row: selection fill present in force.
    let gutter_covered = count_sel(&png, CANVAS_X, CANVAS_Y + 44, CANVAS_X + 50, CANVAS_Y + 64);
    assert!(
        gutter_covered > 100,
        "covered row-2 gutter must use selection fill, got {gutter_covered} px"
    );
    // Header strip over the covered columns: selection fill present.
    let header_covered = count_sel(&png, CANVAS_X + 50, CANVAS_Y, CANVAS_X + 150, CANVAS_Y + 24);
    assert!(
        header_covered > 100,
        "covered col headers must use selection fill, got {header_covered} px"
    );
    // Same-size regions outside the selection must stay clean: the row-1
    // gutter above (plain navigation chrome) and a far-right header span
    // (margin headers).
    let gutter_elsewhere = count_sel(&png, CANVAS_X, CANVAS_Y + 24, CANVAS_X + 50, CANVAS_Y + 44);
    assert_eq!(
        gutter_elsewhere, 0,
        "uncovered gutter must not use selection fill, got {gutter_elsewhere} px"
    );
    let header_elsewhere = count_sel(&png, CANVAS_X + 350, CANVAS_Y, CANVAS_X + 450, CANVAS_Y + 24);
    assert_eq!(
        header_elsewhere, 0,
        "uncovered headers must not use selection fill, got {header_elsewhere} px"
    );

    std::fs::remove_dir_all(&dir).ok();
}
