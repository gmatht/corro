//! Live GUI: clicking away from an in-progress edit must **commit** it to the
//! cell it was entered in, then move to the clicked cell.
//!
//! Reference behaviour (standard spreadsheet UX, and what coro's own edit
//! model implies): type `5` into A1, click A3 — A1 keeps `5` and the cursor
//! moves to A3. Clicking a different cell used to call `start_edit`, which
//! clears `edit_buf`, so the typed value was silently discarded: it was never
//! committed and never restored (unlike Esc, which remembers it via
//! `pending_lost_edit`).
//!
//! Deterministic geometry: the test only relies on ROW_LABEL_W / HEADER_H /
//! ROW_H (published constants), staying inside column A so no column-width
//! guess is needed.
//!
//! Requires: Linux, X server, xdotool, xwd + ImageMagick. Run with:
//!   xvfb-run -a cargo test --features gtk --test gui_click_commit
#![cfg(all(target_os = "linux", feature = "gui"))]

use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::time::{Duration, Instant};

static GUI_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

/// Layout constants mirrored from `gui_backend` (px).
const ROW_H: i32 = 20;
const HEADER_H: i32 = 24;
const ROW_LABEL_W: i32 = 50;
/// Window coordinates of the canvas origin: the menubar and the formula bar
/// sit *above* the canvas, so a click must add this offset before applying
/// the canvas-relative geometry. Calibrated here (click, type, read the
/// committed address): `(112, 82 + row*ROW_H)` addresses A1, A2, A3, ….
const CANVAS_ORIGIN_X: i32 = 50;
const CANVAS_ORIGIN_Y: i32 = 48;

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

/// Centre of main-row `row0` in column A (row0 = 0 is the first data row).
fn cell_a_centre(row0: i32) -> (i32, i32) {
    let x = CANVAS_ORIGIN_X + ROW_LABEL_W + 12;
    let y = CANVAS_ORIGIN_Y + HEADER_H + row0 * ROW_H + ROW_H / 2;
    (x, y)
}

fn click_cell(wid: &str, row0: i32) {
    let (x, y) = cell_a_centre(row0);
    // Move AND click in one xdotool invocation: separate calls leave the
    // pointer move unapplied for the click, and the app sees no button event
    // (the other live click tests do the same for this reason).
    xdotool(&["mousemove", "--window", wid, &x.to_string(), &y.to_string(), "click", "1"]);
}

/// Poll the log until `want` appears as a line (or fail loudly).
fn wait_file_line(path: &PathBuf, want: &str, deadline: Instant) {
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

fn file_has(path: &PathBuf, want: &str) -> bool {
    std::fs::read_to_string(path)
        .unwrap_or_default()
        .lines()
        .any(|l| l == want)
}

#[test]
fn clicking_another_cell_commits_the_edit_to_its_original_cell() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());

    let dir = std::env::temp_dir().join(format!("corro-clickcommit-{}", std::process::id()));
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

    // Edit A1: click it, then type 5.
    click_cell(&wid, 0);
    std::thread::sleep(Duration::from_millis(600));
    xdotool(&["key", "--window", &wid, "5"]);
    std::thread::sleep(Duration::from_millis(600));

    // Click a DIFFERENT cell (A2) while the edit is still in flight.
    // NOTE: only rows that exist stay in the main region — a fresh sheet has
    // two main rows here, and the *next* displayed row is the footer `_1`,
    // so clicking "row 2" would commit into the footer.
    click_cell(&wid, 1);
    std::thread::sleep(Duration::from_millis(1000));

    // The in-flight value must be committed to the cell it was entered in.
    wait_file_line(&file, "SET A1 5", Instant::now() + Duration::from_secs(10));

    // ...and it must not follow the click to the new cell.
    assert!(
        !file_has(&file, "SET A2 5"),
        "the edit must stay on A1, not move with the click:\n{}",
        std::fs::read_to_string(&file).unwrap_or_default()
    );

    // The cursor really did move: typing here must land on A2.
    xdotool(&["key", "--window", &wid, "Z"]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "--window", &wid, "Return"]);
    wait_file_line(&file, "SET A2 Z", Instant::now() + Duration::from_secs(10));

    std::fs::remove_dir_all(&dir).ok();
}
