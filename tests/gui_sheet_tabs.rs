//! Sheet tab bar below the grid (GTK3).
//!
//! - One sheet: no tab bar (the strip stays grid/status chrome).
//! - Sheet > New sheet: the tab bar appears with one tab per sheet, the new
//!   sheet active (bold on yellow, mirroring the terminal's black-on-yellow
//!   active tab; inactive tabs gutter-gray).
//! - Clicking a tab switches to that sheet (status names it, highlight moves).
//!
//! Method (no fixed sleeps for verdicts, no shared files): pixel-structure
//! analysis (exact tab fill/divider colors) plus OCR for titles/status, all
//! polled with deadlines. Sheet creation goes through the real Sheet menu
//! (Alt+S, then clicking "New sheet" — synthetic keys cannot drive GTK's
//! grab-based menus, per gui_menu_key.rs).
//!
//! Requires: Linux, an X server, xdotool, xwd, convert (ImageMagick),
//! python3+PIL, tesseract. Run with:
//!   xvfb-run -a cargo test --features gtk --test gui_sheet_tabs
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

fn screenshot(wid: &str, tag: &str) -> PathBuf {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let xwd = std::env::temp_dir().join(format!("corro-tabs-{tag}-{id}.xwd"));
    let png = std::env::temp_dir().join(format!("corro-tabs-{tag}-{id}.png"));
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

/// Windows of `pid` excluding the main window and tiny ones: menu popups.
fn popup_candidates(pid: u32, main_wid: &str) -> Vec<String> {
    let out = xdotool(&["search", "--pid", &pid.to_string()]);
    out.split_whitespace()
        .filter(|id| *id != main_wid)
        .filter(|id| {
            xdotool(&["getwindowgeometry", "--shell", id])
                .lines()
                .any(|l| {
                    l.strip_prefix("WIDTH=")
                        .and_then(|v| v.trim().parse::<i32>().ok())
                        .map(|w| w > 40)
                        .unwrap_or(false)
                })
        })
        .map(|s| s.to_string())
        .collect()
}

/// Poll with a deadline for a menu popup to appear. Returns the popup id.
fn wait_popup(pid: u32, main_wid: &str, deadline: Instant) -> Vec<String> {
    loop {
        let found = popup_candidates(pid, main_wid);
        if !found.is_empty() {
            return found;
        }
        if Instant::now() > deadline {
            panic!("timed out waiting for menu popup to appear");
        }
        std::thread::sleep(Duration::from_millis(150));
    }
}

/// Tab-strip metrics for a screenshot: yellow (active tab) pixel count and
/// bbox, divider x positions, and the strip row. Exact tab colors: active
/// fill (255,255,153), dividers (140,140,140) — dividers scan independently
/// of yellow so absence (no tab bar at all) is measurable too. Prints
/// `yellow_n yellow_x0 yellow_x1 divs... | stripy`.
fn strip_metrics(shot: &PathBuf) -> (i32, i32, i32, Vec<i32>, i32) {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-tabs-metrics-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg = Image.open(sys.argv[1]).convert('RGB')\nW,H = img.size\npx = img.load()\nYL=(255,255,153)\nDIV=(140,140,140)\nys=[y for y in range(H-100,H) for x in range(0,W,3) if px[x,y]==YL]\nxs=[x for y in range(H-100,H) for x in range(0,W) if px[x,y]==YL]\nstripy=(sum(ys)//len(ys)) if ys else -1\ndivs=sorted(set(x for x in range(0,min(600,W)) for y0 in range(max(0,H-100),H-12) if all(px[x,y]==DIV for y in range(y0,y0+12))))\nprint(f'{len(xs)} {(min(xs) if xs else -1)} {(max(xs) if xs else -1)} {\" \".join(map(str,divs))} | {stripy}')\n",
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
    let parts: Vec<&str> = text.split('|').collect();
    assert_eq!(parts.len(), 2, "bad analyzer output: {text:?}");
    let nums: Vec<i32> = parts[0].split_whitespace().map(|s| s.parse().unwrap_or(-99)).collect();
    assert!(nums.len() >= 3, "bad analyzer output: {text:?}");
    let stripy: i32 = parts[1].trim().parse().unwrap_or(-99);
    (nums[0], nums[1], nums[2], nums[3..].to_vec(), stripy)
}

/// OCR a window-relative region (upscaled 300%).
fn ocr_region(shot: &PathBuf, x: i32, y: i32, w: i32, h: i32, psm: &str) -> String {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let crop = std::env::temp_dir().join(format!("corro-tabs-ocr-{id}.png"));
    Command::new("convert")
        .arg(shot)
        .args(["-crop", &format!("{w}x{h}+{x}+{y}"), "+repage", "-resize", "300%"])
        .arg(&crop)
        .status()
        .expect("convert crop");
    let out = Command::new("tesseract")
        .arg(&crop)
        .arg("stdout")
        .args(["--psm", psm])
        .output()
        .expect("tesseract");
    let _ = std::fs::remove_file(&crop);
    String::from_utf8_lossy(&out.stdout).to_string()
}

fn fresh_fixture(tag: &str, body: &str) -> PathBuf {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path =
        std::env::temp_dir().join(format!("corro-tabs-{tag}-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, body).expect("write fixture");
    path
}

/// One sheet: no tab bar. The bottom chrome must show neither the active-tab
/// yellow nor any long gutter-gray tab box run (gridlines are 1px; a tab box
/// is a 20px+ run of (230,230,230)).
#[test]
fn gui_single_sheet_hides_tab_bar() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture("single", "CORRO_LOG 1\n");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(800));
    let shot = screenshot(&wid, "single");
    let (yellow, _, _, divs, _) = strip_metrics(&shot);
    assert_eq!(
        yellow, 0,
        "single sheet must show no active-tab yellow (found {yellow}px)"
    );
    assert!(
        divs.is_empty(),
        "single sheet must show no tab boxes (divider pixels at {divs:?})"
    );
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}

/// Sheet > New sheet: the tab bar appears with one tab per sheet, the new
/// sheet active. The 8-item Sheet menu puts "New sheet" 3rd (~0.31 down).
#[test]
fn gui_new_sheet_shows_two_tabs() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture("newtab", "CORRO_LOG 1\n");
    let mut child = spawn_gui(&path);
    let pid = child.id();
    let wid = find_corro_window(pid, Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "--window", &wid, "alt+s"]);
    let found = wait_popup(pid, &wid, Instant::now() + Duration::from_secs(8));
    let popup = found[0].clone();
    let geo = xdotool(&["getwindowgeometry", "--shell", &popup]);
    let (mut px, mut py, mut pw, mut ph): (i32, i32, i32, i32) = (0, 0, 0, 0);
    for l in geo.lines() {
        if let Some(v) = l.strip_prefix("X=") {
            px = v.trim().parse().unwrap_or(0);
        }
        if let Some(v) = l.strip_prefix("Y=") {
            py = v.trim().parse().unwrap_or(0);
        }
        if let Some(v) = l.strip_prefix("WIDTH=") {
            pw = v.trim().parse().unwrap_or(0);
        }
        if let Some(v) = l.strip_prefix("HEIGHT=") {
            ph = v.trim().parse().unwrap_or(0);
        }
    }
    assert!(ph > 60, "Sheet popup height implausible ({ph})");
    // Settle: the popup must report the same geometry twice before the
    // click, or the item rows may still be laying out (click would miss).
    let deadline_settle = Instant::now() + Duration::from_secs(5);
    loop {
        std::thread::sleep(Duration::from_millis(300));
        let geo2 = xdotool(&["getwindowgeometry", "--shell", &popup]);
        let mut ph2 = 0;
        for l in geo2.lines() {
            if let Some(v) = l.strip_prefix("HEIGHT=") {
                ph2 = v.trim().parse().unwrap_or(0);
            }
        }
        if (ph2 - ph).abs() <= 4 || Instant::now() > deadline_settle {
            ph = ph2.max(ph);
            break;
        }
        ph = ph2;
    }
    // Screen coords for the override-redirect popup (as in gui_menu_key).
    xdotool(&[
        "mousemove",
        &format!("{}", px + pw / 2),
        &format!("{}", py + (ph as f64 * 0.31) as i32),
        "click",
        "1",
    ]);
    // Proof of activation independent of rendering: the New-sheet op must
    // land in the live log file (deadline, not sleep).
    let deadline_log = Instant::now() + Duration::from_secs(10);
    loop {
        let log = std::fs::read_to_string(&path).unwrap_or_default();
        if log.contains("$2:NEW_SHEET") || Instant::now() > deadline_log {
            assert!(
                log.contains("$2:NEW_SHEET"),
                "menu click must create Sheet2 (log: {log:?})"
            );
            break;
        }
        std::thread::sleep(Duration::from_millis(300));
    }
    // Dismiss the menu (a grab-open menu can starve main-window redraws):
    // click the status label, which has no click handler and is harmless.
    xdotool(&["mousemove", "--window", &wid, "600", "785", "click", "1"]);
    // Appearance event: the active-tab yellow must show up (deadline).
    let deadline = Instant::now() + Duration::from_secs(12);
    let (yellow, y0, y1, divs, stripy) = loop {
        let shot = screenshot(&wid, "newtab");
        let m = strip_metrics(&shot);
        if m.0 > 200 || Instant::now() > deadline {
            break m;
        }
        std::thread::sleep(Duration::from_millis(300));
    };
    assert!(
        yellow > 200,
        "new sheet must raise the tab bar (active-tab yellow {yellow}px)"
    );
    assert!(
        y1 - y0 > 30 && y1 - y0 < 400,
        "active tab must be tab-sized, not specks or a full-strip flood (yellow {y0}..{y1})"
    );
    assert!(
        divs.len() >= 2,
        "two tabs need two trailing dividers, got {divs:?}"
    );
    // Titles, not just boxes: "Sheet" reads twice (tesseract renders
    // Sheet1's 1 as `|`), Sheet2 reads exact, and the yellow starts past the
    // first divider — the new tab is titled, second, and active.
    let shot = screenshot(&wid, "newtabtitles");
    let text = ocr_region(&shot, 0, stripy - 20, 700, 44, "6");
    let _ = std::fs::remove_file(&shot);
    assert!(
        text.matches("Sheet").count() >= 2,
        "tab bar must title both sheets, OCR got: {text:?}"
    );
    assert!(
        text.contains("Sheet2"),
        "new tab must read Sheet2, OCR got: {text:?}"
    );
    assert!(
        y0 > divs[0],
        "new sheet must be the active (second) tab (yellow starts {y0}, dividers {divs:?})"
    );
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}

/// Clicking a tab switches to that sheet: the status names it and the
/// active highlight moves to its tab.
#[test]
fn gui_tab_click_switches_active_sheet() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture("clicktab", "CORRO_LOG 1\n$2:NEW_SHEET Sheet2\n");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(800));
    // Locate the strip: dividers on the active-tab row; click the second
    // tab's middle (fully dynamic — no hardcoded tab widths).
    let shot = screenshot(&wid, "clicktab0");
    let (_, y0, _, divs, stripy) = strip_metrics(&shot);
    assert!(
        divs.len() >= 2,
        "two tabs need two dividers to aim between, got {divs:?}"
    );
    assert!(
        y1_in_first_tab(y0, &divs),
        "fresh two-sheet app must start on Sheet1 (yellow {y0}, dividers {divs:?})"
    );
    let target = (divs[0] + divs[1]) / 2;
    xdotool(&[
        "mousemove",
        "--window",
        &wid,
        &target.to_string(),
        &stripy.to_string(),
        "click",
        "1",
    ]);
    // Verdict 1 (polled): the status names the new active sheet.
    let deadline = Instant::now() + Duration::from_secs(10);
    let status = loop {
        std::thread::sleep(Duration::from_millis(300));
        let shot = screenshot(&wid, "clicktabstatus");
        // Status label lives at the very bottom of the 1200x800 window.
        let text = ocr_region(&shot, 0, 770, 1200, 30, "6");
        let _ = std::fs::remove_file(&shot);
        if text.contains("Sheet 2 of 2") || Instant::now() > deadline {
            break text;
        }
    };
    assert!(
        status.contains("Sheet 2 of 2"),
        "clicking Sheet2's tab must activate it, status OCR got: {status:?}"
    );
    // Verdict 2: the active highlight moved past the first divider.
    let shot = screenshot(&wid, "clicktab1");
    let (_, ny0, _, _, _) = strip_metrics(&shot);
    assert!(
        ny0 > divs[0],
        "active highlight must move to the second tab (yellow starts {ny0}, first divider {})",
        divs[0]
    );
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}

/// Fresh two-sheet app starts on the first sheet: the active-tab yellow
/// begins left of the first divider.
fn y1_in_first_tab(yellow_x0: i32, divs: &[i32]) -> bool {
    yellow_x0 >= 0 && yellow_x0 < divs[0]
}

/// Window id of the dialog titled `title` (deadline): None on timeout.
/// Dialog titles come from prompt_chrome — a mislabeled dialog (Rename
/// showing "Find") fails the title assertion instead of reaching users.
fn wait_dialog_title(pid: u32, title: &str, deadline: Instant) -> Option<String> {
    loop {
        for id in xdotool(&["search", "--pid", &pid.to_string()]).split_whitespace() {
            if xdotool(&["getwindowname", id]).trim() == title {
                return Some(id.to_string());
            }
        }
        if Instant::now() > deadline {
            return None;
        }
        std::thread::sleep(Duration::from_millis(200));
    }
}


/// Sheet > Rename sheet renames the active sheet end to end. Regression:
/// Rename opened a dialog titled "Find" (the prompt funnel recycled
/// find_dialog for every prompt action). The title assertion pins the fix;
/// the tab re-title, status, and log op prove the rename itself.
#[test]
fn gui_rename_sheet_retitles_active_tab() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture("renametab", "CORRO_LOG 1\n$2:NEW_SHEET Sheet2\n");
    let mut child = spawn_gui(&path);
    let pid = child.id();
    let wid = find_corro_window(pid, Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "--window", &wid, "alt+s"]);
    let found = wait_popup(pid, &wid, Instant::now() + Duration::from_secs(8));
    let popup = found[0].clone();
    let geo = xdotool(&["getwindowgeometry", "--shell", &popup]);
    let (mut px, mut py, mut pw, mut ph): (i32, i32, i32, i32) = (0, 0, 0, 0);
    for l in geo.lines() {
        if let Some(v) = l.strip_prefix("X=") {
            px = v.trim().parse().unwrap_or(0);
        }
        if let Some(v) = l.strip_prefix("Y=") {
            py = v.trim().parse().unwrap_or(0);
        }
        if let Some(v) = l.strip_prefix("WIDTH=") {
            pw = v.trim().parse().unwrap_or(0);
        }
        if let Some(v) = l.strip_prefix("HEIGHT=") {
            ph = v.trim().parse().unwrap_or(0);
        }
    }
    assert!(ph > 60, "Sheet popup height implausible ({ph})");
    // Rename is 4th of 8 items (~0.44 down).
    xdotool(&[
        "mousemove",
        &format!("{}", px + pw / 2),
        &format!("{}", py + (ph as f64 * 0.44) as i32),
        "click",
        "1",
    ]);
    // The dialog must be titled "Rename sheet" — not "Find".
    let dlg = wait_dialog_title(pid, "Rename sheet", Instant::now() + Duration::from_secs(10))
        .expect("Rename sheet must open a dialog titled 'Rename sheet'");
    // Entry arrives focused with the current title selected: typing replaces.
    // Settle for focus (the dialog just mapped; verdicts below stay polled).
    std::thread::sleep(Duration::from_millis(800));
    // Confirm with Enter (ratatui parity: type + Enter renames, Esc cancels).
    // Dialog screen geometry reports (0,0), so coordinate clicks cannot aim
    // at it — keyboard confirmation is also the honest end-to-end proof.
    xdotool(&["type", "--window", &dlg, "Budget"]);
    std::thread::sleep(Duration::from_millis(400));
    xdotool(&["key", "--window", &dlg, "Return"]);
    // The dialog must close (deadline) and the rename must land: tab title,
    // status, and log op.
    let deadline = Instant::now() + Duration::from_secs(10);
    loop {
        if wait_dialog_title(pid, "Rename sheet", Instant::now()) .is_none() {
            break;
        }
        if Instant::now() > deadline {
            panic!("Rename dialog never closed after confirming");
        }
        std::thread::sleep(Duration::from_millis(300));
    }
    let log = std::fs::read_to_string(&path).unwrap_or_default();
    assert!(
        log.contains("RENAME_SHEET") && log.contains("Budget"),
        "rename must log RENAME_SHEET Budget (log: {log:?})"
    );
    let shot = screenshot(&wid, "renamedtab");
    let (_, ny0, _, ndivs, nstripy) = strip_metrics(&shot);
    assert!(
        ndivs.len() >= 2 && ny0 < ndivs[0],
        "renamed tab stays first and active (yellow {ny0}, dividers {ndivs:?})"
    );
    // Poll for the re-titled text (the strip repaints asynchronously after
    // the dialog closes); the content assertion below is the verdict.
    let deadline = Instant::now() + Duration::from_secs(8);
    let text = loop {
        let shot = screenshot(&wid, "renamedtabtext");
        let text = ocr_region(&shot, 0, nstripy - 20, 700, 44, "6");
        let _ = std::fs::remove_file(&shot);
        if text.contains("Budget") || Instant::now() > deadline {
            break text;
        }
        std::thread::sleep(Duration::from_millis(300));
    };
    assert!(
        text.contains("Budget"),
        "first tab must read Budget after rename, OCR got: {text:?}"
    );
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}
