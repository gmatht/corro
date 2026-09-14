//! GUI status chrome, ratatui parity (GTK3).
//!
//! Ratatui contract (mirrored here):
//! - bottom row always shows the hints line, never status;
//! - the formula bar shows `· status` trailing only when status is
//!   non-empty and no edit/input owns the row (Edit arms show the buffer
//!   with no status span).
//!
//! GUI mapping under test:
//! - T1: with a failing CORRO_TEMPLATE the startup status names the
//!   template in the TOP formula row (its own suffix label).
//! - T2: the BOTTOM strip shows the shared hints line ("F2" etc.).
//! - T3: the status text is NOT inside the editable entry (left part of
//!   the formula row lacks it) — suffix is a separate widget, so status
//!   can never leak into the edit buffer.
//! (Hide-path layout — empty status leaves the entry full width — is
//! guarded by the fx-bar width assertions in the nwg cram/GTK fill tests.)
//!
//! Method: bad-template startup makes the status deterministic from the
//! first settled frame (no keystrokes needed); OCR polled with deadlines,
//! no fixed sleeps for verdicts, no shared files.
//!
//! Requires: Linux, an X server, xdotool, xwd, convert (ImageMagick),
//! python3+PIL, tesseract. Run with:
//!   xvfb-run -a cargo test --features gtk --test gui_status_bar
#![cfg(all(target_os = "linux", feature = "gui"))]

use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::time::{Duration, Instant};

static COUNTER: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
static GUI_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

fn xdotool(args: &[&str]) -> String {
    String::from_utf8_lossy(
        &Command::new("xdotool").args(args).output().expect("xdotool failed").stdout,
    )
    .to_string()
}

struct KillOnDrop(Child);
impl std::ops::Deref for KillOnDrop {
    type Target = Child;
    fn deref(&self) -> &Child { &self.0 }
}
impl Drop for KillOnDrop {
    fn drop(&mut self) {
        let _ = self.0.kill();
        let _ = self.0.wait();
    }
}

fn spawn_gui_status(template: &str) -> KillOnDrop {
    let bin = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("target/debug/corro");
    KillOnDrop(
        Command::new(&bin)
            // No file arg: an explicit path wins over the template (same as
            // the terminal), so a fresh doc is the only template trigger.
            .arg("--gui")
            .env("CORRO_TEMPLATE", template)
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .spawn()
            .expect("spawn corro --gui"),
    )
}

fn find_corro_window(child_pid: u32, deadline: Instant) -> String {
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

fn window_size(wid: &str) -> (i32, i32) {
    let geo = xdotool(&["getwindowgeometry", "--shell", wid]);
    let mut w = 0;
    let mut h = 0;
    for l in geo.lines() {
        if let Some(v) = l.strip_prefix("WIDTH=") {
            w = v.trim().parse().unwrap_or(0);
        }
        if let Some(v) = l.strip_prefix("HEIGHT=") {
            h = v.trim().parse().unwrap_or(0);
        }
    }
    (w, h)
}

fn screenshot(wid: &str, tag: &str) -> PathBuf {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let xwd = std::env::temp_dir().join(format!("corro-status-{tag}-{id}.xwd"));
    let png = std::env::temp_dir().join(format!("corro-status-{tag}-{id}.png"));
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

/// OCR over the formula row's left (entry) zone, upscaled 2x: entry
/// content renders at x~80 while the status suffix hugs the right end
/// (x>1000 in the 1200px window), so this crop structurally excludes
/// suffix text — the separation verdict below rests on that geometry.
/// Sensor calibration: at entry-zone scale tesseract renders a small
/// "QQ" as "RQ" (Q→R confusion, verified on kept frames where the entry
/// visibly contains QQ — same class as the tabs test matching Sheet1's
/// "1" as "|"). Match either; the absence verdicts below stay strict
/// (long words read reliably, so a missing "Template" is meaningful).
fn sees_typed_qq(text: &str) -> bool {
    text.contains("QQ") || text.contains("RQ")
}

/// OCR over the formula row's right (suffix) end, upscaled 2x as a
/// single line: the short "Set cell A1" status vanishes in full-row
/// reads (only fragments survive) but the isolated region reads
/// reliably — calibrated as "Set cell Al" (1/l confusion, same class as
/// the tabs test matching Sheet1's "1" as "|").
fn ocr_suffix_zone(shot: &PathBuf, w: i32) -> String {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let crop = std::env::temp_dir().join(format!("corro-status-suffix-{id}.png"));
    Command::new("convert")
        .arg(shot)
        .args(["-crop", &format!("400x50+{}+0", (w - 400).max(0)), "+repage", "-colorspace", "Gray", "-resize", "200%"])
        .arg(&crop)
        .status()
        .expect("convert suffix crop");
    let out = Command::new("tesseract")
        .arg(&crop)
        .arg("stdout")
        .args(["--psm", "7"])
        .output()
        .expect("tesseract");
    let _ = std::fs::remove_file(&crop);
    String::from_utf8_lossy(&out.stdout).to_string()
}

fn ocr_entry_zone(shot: &PathBuf) -> String {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let crop = std::env::temp_dir().join(format!("corro-status-entry-{id}.png"));
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

fn ocr_region(shot: &PathBuf, x: i32, y: i32, w: i32, h: i32, psm: &str) -> String {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let crop = std::env::temp_dir().join(format!("corro-status-ocr-{id}.png"));
    Command::new("convert")
        .arg(shot)
        .args(["-crop", &format!("{w}x{h}+{x}+{y}"), "+repage", "-colorspace", "Gray"])
        .arg(&crop)
        .status()
        .expect("convert crop");
    // Single-threaded: tesseract defaults to all cores, and a sustained
    // poll storm would starve the app's frame clock (queued label
    // repaints never fire, screenshots stay stale). Verdicts stay polled.
    let out = Command::new("tesseract")
        .env("OMP_NUM_THREADS", "1")
        .arg(&crop)
        .arg("stdout")
        .args(["--psm", psm])
        .output()
        .expect("tesseract");
    let _ = std::fs::remove_file(&crop);
    String::from_utf8_lossy(&out.stdout).to_string()
}


/// Await first paint (deadline): present()'s pump stalls for highly
/// variable wall time on virtual displays — screenshots before it show
/// menu chrome with an unpainted body. Any formula-row content (address,
/// fx, or a status) proves paint flowed. Once past ~15s without paint,
/// send one Down+Up round trip as a best-effort pump kick (moves A1→A2→A1:
/// no commit, no status change, verdict-neutral). A timeout here means a
/// sick environment, not a chrome bug.
fn await_first_paint(wid: &str, w: i32) {
    let verdict = Instant::now() + Duration::from_secs(45);
    let mut poked = false;
    loop {
        let shot = screenshot(wid, "firstpaint");
        let text = ocr_region(&shot, 0, 0, w, 50, "6");
        let _ = std::fs::remove_file(&shot);
        if text.contains('A')
            || text.contains("fx")
            || text.contains("Template")
            || Instant::now() > verdict
        {
            assert!(
                text.contains('A') || text.contains("fx") || text.contains("Template"),
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
/// Passive chrome (zero input, immune to key-delivery flakiness by
/// construction): a bad CORRO_TEMPLATE names the template in the TOP
/// formula row (suffix label), while the BOTTOM strip shows the shared
/// hints line.
#[test]
fn gui_startup_status_lives_in_formula_row_hints_at_bottom() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    let _child = spawn_gui_status("/nonexistent-dir/corro-template-note.corro");
    let pid = _child.id();
    let deadline = Instant::now() + Duration::from_secs(30);
    let wid = find_corro_window(pid, deadline);
    let (w, h) = window_size(&wid);
    assert!(w > 100 && h > 100, "tiny window {w}x{h}");

    await_first_paint(&wid, w);
    // T1 (polled): the template-failure status appears in the top formula row.
    // Require CONSECUTIVE sightings: a single frame can catch the suffix
    // mid-layout (present()'s pump paints transient states), and typing
    // into a transient frame goes nowhere. Stable twice == settled.
    let verdict = Instant::now() + Duration::from_secs(30);
    let mut stable = 0;
    loop {
        let shot = screenshot(&wid, "formularow");
        let text = ocr_region(&shot, 0, 0, w, 50, "6");
        let _ = std::fs::remove_file(&shot);
        stable = if text.contains("Template") { stable + 1 } else { 0 };
        if stable >= 2 || Instant::now() > verdict {
            assert!(
                text.contains("Template"),
                "formula row lacks startup status, OCR: {text:?}"
            );
            break;
        }
        std::thread::sleep(Duration::from_millis(400));
    }
    // T2 (polled): bottom strip shows the shared hints line (ratatui parity text).
    let verdict = Instant::now() + Duration::from_secs(20);
    let bottom = loop {
        let shot = screenshot(&wid, "bottomstrip");
        let text = ocr_region(&shot, 0, h - 30, w, 30, "6");
        let _ = std::fs::remove_file(&shot);
        if text.contains("F2") || Instant::now() > verdict {
            break text;
        }
        std::thread::sleep(Duration::from_millis(400));
    };
    assert!(
        bottom.contains("F2"),
        "bottom strip lacks hints line, OCR: {bottom:?}"
    );
}

/// Interactive chrome (input early, mirroring the reliably-passing edit
/// tests): typing then Return commits, the suffix reports "Set cell",
/// and navigating back shows the raw cell value in the entry with no
/// status text (separate suffix widget — status can never leak into the
/// edit buffer). All keystrokes precede all verdict polls.
#[test]
fn gui_commit_status_and_entry_separation() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    let _child = spawn_gui_status("/nonexistent-dir/corro-template-note.corro");
    let pid = _child.id();
    let deadline = Instant::now() + Duration::from_secs(30);
    let wid = find_corro_window(pid, deadline);
    let (w, _h) = window_size(&wid);
    assert!(w > 100, "tiny window {w}");
    // Activate late and quiesce: screenshots need no focus, but synthetic
    // keys do — and the polling storm above starves the pump, so activating
    // early lets focus rot before typing. The sleep is setup quiet, not a
    // verdict (every verdict below stays deadline-polled); it mirrors the
    // idle beat shell-driven runs have between activate and type.
    xdotool(&["windowactivate", "--sync", &wid]);
    await_first_paint(&wid, w);
    std::thread::sleep(Duration::from_secs(2));
    // Settle on appearance between keystrokes: back-to-back synthetic
    // type+Return gets swallowed by the double-fire guards, so the poll
    // turns below are the (event-based) delay, not sleeps-as-verdicts.
    xdotool(&["type", "--window", &wid, "QQ"]);
    // Consecutive sightings: a lone "RQ" can be OCR garbage on a Template
    // frame, but a persistent signature means rendered QQ. Keep the
    // matching frame for post-mortem.
    let verdict = Instant::now() + Duration::from_secs(20);
    let mut stable = 0;
    loop {
        let shot = screenshot(&wid, "typedqq");
        let text = ocr_entry_zone(&shot);
        stable = if sees_typed_qq(&text) { stable + 1 } else { 0 };
        if stable >= 2 {
            let dst = std::env::temp_dir().join("corro-status-typedkeep.png");
            let _ = std::fs::copy(&shot, &dst);
        }
        let _ = std::fs::remove_file(&shot);
        if stable >= 2 || Instant::now() > verdict {
            assert!(stable >= 2, "typed QQ never settled in the entry, got: {text:?}");
            break;
        }
        std::thread::sleep(Duration::from_millis(400));
    }
    // Quiesce before Return: firing it in the same breath as the poll
    // storm starves the pump mid-dispatch (setup quiet — every verdict
    // stays deadline-polled).
    std::thread::sleep(Duration::from_secs(2));
    xdotool(&["key", "--window", &wid, "Return"]);
    // Idle after Return before polling: the queued label repaint needs
    // frame-clock turns, which a back-to-back poll storm would starve
    // (setup quiet — the verdict below stays deadline-polled).
    std::thread::sleep(Duration::from_secs(5));
    // T3: the commit status reached the formula-row suffix (suffix live
    // post-commit, not just at startup).
    let verdict = Instant::now() + Duration::from_secs(20);
    loop {
        let shot = screenshot(&wid, "commitrow");
        let text = ocr_suffix_zone(&shot, w);
        let _ = std::fs::remove_file(&shot);
        if text.contains("Set cell") || Instant::now() > verdict {
            assert!(
                text.contains("Set cell"),
                "commit must report status in formula row, OCR: {text:?}"
            );
            break;
        }
        std::thread::sleep(Duration::from_millis(400));
    }
    // T3b: separation. Navigate back to A1: the entry shows the raw cell
    // value. A buffer leak would render status text inside the entry zone;
    // the zone crop structurally excludes the suffix, so status words
    // here fail. Reader control: the committed value itself must read —
    // a blind or empty read fails there, so it can never fake "clean"
    // (single-char landmarks like "fx" mangle unpredictably at this
    // scale, but the two-char value signature reads reliably).
    xdotool(&["key", "--window", &wid, "Up"]);
    let verdict = Instant::now() + Duration::from_secs(20);
    let entry = loop {
        let shot = screenshot(&wid, "entryzone");
        let text = ocr_entry_zone(&shot);
        let _ = std::fs::remove_file(&shot);
        if sees_typed_qq(&text) || Instant::now() > verdict {
            break text;
        }
        std::thread::sleep(Duration::from_millis(400));
    };
    assert!(sees_typed_qq(&entry), "committed value missing from entry, got: {entry:?}");
    assert!(
        !entry.contains("Set cell") && !entry.contains("Template"),
        "status leaked into editable entry, OCR: {entry:?}"
    );
}
