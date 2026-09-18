//! Live GUI regression: editing in one window must appear in another
//! *without* the user touching the second window.
//!
//! corro is an append-only collaborative spreadsheet, so a window that has a
//! file open tails the log (`CoreApp::poll_log_tail`) and repaints when
//! another window commits. The GUI used to poll only inside `handle_key`, so
//! the mirror arrived only after you pressed a key in the window you were
//! watching (the TUI polls every loop iteration and was never affected).
//!
//! This test reproduces "another window edited the file" the same way the app
//! does it — by appending a log line — and then requires the GUI to repaint
//! with **no input whatsoever sent to it**. Before the fix the screen stayed
//! frozen at the baseline; the control check at the end proves the change is
//! attributable to the append (the same window is otherwise pixel-stable).
//!
//! Requires: Linux, an X server, xdotool, xwd + ImageMagick, python3 + PIL.
//! Run with: xvfb-run -a cargo test --features gtk --test gui_window_mirror
#![cfg(all(target_os = "linux", feature = "gui"))]

use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::time::{Duration, Instant};

/// Live GUI tests share one X server and global focus: serialize them.
static GUI_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

/// Rendering a committed value changes hundreds of pixels; the window also
/// has low-level churn of its own (cursor/anti-aliasing) worth a few dozen.
/// Requiring a real repaint keeps the verdict about content, not noise.
/// Measured: noise ~17 px, a repainted cell ≫ this.
const MIN_REPAINT_PIXELS: u64 = 200;

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

/// Screenshot the window (xwd -> PNG).
fn shot(wid: &str, tag: &str) -> PathBuf {
    let png = std::env::temp_dir().join(format!(
        "corro-mirror-{}-{tag}.png",
        std::process::id()
    ));
    let xwd = png.with_extension("xwd");
    Command::new("xwd")
        .args(["-id", wid, "-out"])
        .arg(&xwd)
        .status()
        .expect("xwd failed");
    Command::new("convert")
        .arg(&xwd)
        .arg(&png)
        .status()
        .expect("convert failed");
    png
}

/// Number of pixels differing by more than 16 levels between two screenshots.
fn pixel_diff(a: &PathBuf, b: &PathBuf) -> u64 {
    let script = r#"
import sys
from PIL import Image, ImageChops
a = Image.open(sys.argv[1]).convert('L')
b = Image.open(sys.argv[2]).convert('L')
if a.size != b.size:
    print(10**9); raise SystemExit
d = ImageChops.difference(a, b)
print(sum(1 for v in d.getdata() if v > 16))
"#;
    let out = Command::new("python3")
        .arg("-c")
        .arg(script)
        .arg(a)
        .arg(b)
        .output()
        .expect("python3 failed");
    String::from_utf8_lossy(&out.stdout).trim().parse().expect("pixel count")
}

/// Wait until two consecutive captures are identical (the window has settled),
/// returning the stable screenshot. Deadline-based, never a fixed sleep.
fn settled_shot(wid: &str, tag: &str, deadline: Instant) -> PathBuf {
    let mut prev = shot(wid, &format!("{tag}-a"));
    loop {
        let cur = shot(wid, &format!("{tag}-b"));
        if pixel_diff(&prev, &cur) == 0 {
            return cur;
        }
        if Instant::now() > deadline {
            panic!("window never settled for {tag}");
        }
        prev = cur;
        std::thread::sleep(Duration::from_millis(250));
    }
}

#[test]
fn external_commit_repaints_without_input() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());

    let dir = std::env::temp_dir().join(format!("corro-mirror-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let file = dir.join("wb.corro");
    // A single-sheet log: appends from the "other window" use plain `SET A1 v`.
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
    let pid = child.id();
    let wid = find_corro_window(pid, Instant::now() + Duration::from_secs(30));
    xdotool(&["windowactivate", "--sync", &wid]);
    xdotool(&["windowfocus", &wid]);

    // Baseline: the window must be pixel-stable *before* anything is written,
    // otherwise a "changed" verdict below would be meaningless.
    let deadline = Instant::now() + Duration::from_secs(30);
    let base = settled_shot(&wid, "base", deadline);

    // Another window commits: append the op the app itself would append.
    // NO keys, NO mouse: nothing is sent to this window from here on.
    {
        use std::io::Write;
        let mut f = std::fs::OpenOptions::new().append(true).open(&file).unwrap();
        f.write_all(b"SET A1 777\n").unwrap();
        f.flush().unwrap();
    }

    // The mirror must arrive on its own (timer-driven poll, ~250 ms).
    let deadline = Instant::now() + Duration::from_secs(15);
    let mut changed: u64;
    loop {
        let cur = shot(&wid, "after");
        changed = pixel_diff(&base, &cur);
        if changed > MIN_REPAINT_PIXELS {
            break;
        }
        if Instant::now() > deadline {
            panic!(
                "window did not repaint after another window committed (no input was sent): \
                 the log tail is not being polled while idle"
            );
        }
        std::thread::sleep(Duration::from_millis(200));
    }
    assert!(
        changed > MIN_REPAINT_PIXELS,
        "expected a visible repaint from the external commit, got {changed} changed pixels"
    );

    // Sanity: the app really did apply the op (not just a cosmetic repaint).
    let text = std::fs::read_to_string(&file).unwrap();
    assert!(text.contains("SET A1 777"), "append survived: {text}");

    std::fs::remove_dir_all(&dir).ok();
}
