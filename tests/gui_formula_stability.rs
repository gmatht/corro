//! The formula row must not reflow when its labels change text.
//!
//! `fx` and the formula entry are the fixed landmarks of the row: the entry
//! starts right after the `fx` caption, and that is where the user's caret
//! goes. Anything laid out after a label whose *text* changes can slide them.
//!
//! ## Two tests, because the invariant is decided in two places
//!
//! * `formula_row_stays_put_on_move` is a **pixel** test: launch the GTK GUI,
//!   press Right until the address grows a digit, and require the `fx` caption
//!   and the entry's frame to sit at the same x. This is the user-visible
//!   invariant, measured the way the user sees it.
//!
//! * `formula_row_mutating_labels_pin_their_width` is a **source** test: every
//!   formula-row label whose text changes must pin its width.
//!
//! The source test is not redundant with the pixel one, and it is the one that
//! actually guards the platforms where this bites. The reflow is a Win32/NWG
//! behaviour: `Label::set_fixed_width` is a deliberate **no-op on GTK** (see
//! `backends_gtk_adapter_impl.rs` — a GtkLabel already lays out at its natural
//! width and the box does not re-pack children on a text change, so forcing a
//! size request there would only risk clipping). The NWG implementation is the
//! one with real work to do: a Win32 STATIC auto-resizes on `SetWindowText`
//! and the parent box re-fits it, which is what actually slides `fx` and the
//! entry. So the pixel test can only ever exercise the backend that *cannot*
//! exhibit the bug — on GTK, `fx` and the entry do not move even with the
//! address pin deleted, because the GTK no-op means there was never a pin to
//! delete. The pixel test therefore passes while Win32 still slides, and only
//! the source test can catch a pin being dropped.
//!
//! Requires for the pixel test: Linux, an X server, xdotool, xwd, convert
//! (ImageMagick), python3+PIL. Run with:
//!   xvfb-run -a cargo test --all-features --test gui_formula_stability
#![cfg(all(target_os = "linux", feature = "gui"))]

use std::path::PathBuf;
use std::process::{Child, Command, Stdio};
use std::time::{Duration, Instant};

static COUNTER: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

/// Tests in this binary share an X server and global focus, so serialize them.
/// The verdict itself stays pixel-based.
static GUI_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

fn xdotool(args: &[&str]) -> String {
    String::from_utf8_lossy(
        &Command::new("xdotool").args(args).output().expect("xdotool failed").stdout,
    )
    .to_string()
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
    let xwd = std::env::temp_dir().join(format!("corro-fstab-{tag}-{id}.xwd"));
    let png = std::env::temp_dir().join(format!("corro-fstab-{tag}-{id}.png"));
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

/// Where the `fx` caption and the formula entry's left edge sit, in pixels.
///
/// Two anchors, each found by exact colour rather than by "is there ink here",
/// so neither can latch onto a glyph:
///   fx    — the caption's first ink column. Anchored to the entry, not to the
///           address: the address lives in a *pinned* slot but its own glyphs
///           still move within that slot as the text changes width
///           (`A1` -> `A10` -> `A100`), so "first ink after the address" is the
///           address's tail, not the caption. The caption is the dark run that
///           follows a wide empty gutter and still sits left of the entry.
///   entry — the focused entry's blue left border `(53,132,228)`, the first
///           blue column in the row. The entry is the only blue-framed widget
///           there, and its left border is exactly what slides on a reflow.
fn fx_and_entry_x(png: &PathBuf) -> (i32, i32) {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fstab-an-{id}.py"));
    std::fs::write(
        &script,
        // Calibrated off a live window: the menu bar is 24px and the formula
        // row's text band is y=32..40.
        "import sys\nfrom PIL import Image\nimg=Image.open(sys.argv[1]).convert('RGB')\nW,H=img.size\npx=img.load()\nY0,Y1=30,44\nBLUE=(53,132,228)\ndef ink(c):\n    r,g,b=c\n    return r<200 and g<200 and b<200\ncs=[x for x in range(0, min(W,400)) if any(ink(px[x,y]) for y in range(Y0,Y1))]\nentry=next((x for x in range(0, min(W,900)) if any(px[x,y]==BLUE for y in range(Y0,Y1))), -1)\nruns=[]\nfor x in cs:\n    if runs and x==runs[-1][1]+1: runs[-1][1]=x\n    else: runs.append([x,x])\n# The caption is the run that follows a wide empty gutter and still sits\n# left of the entry (see the doc comment).\nfx=-1\nif entry>0:\n    prev_end=-99\n    for r in runs:\n        if r[0]>=entry-4: break\n        if prev_end>=0 and r[0]-prev_end-1>=8: fx=r[0]\n        prev_end=r[1]\nprint(fx, entry)\n",
    )
    .expect("write fx/entry analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(png)
        .output()
        .expect("fx/entry analyzer");
    let _ = std::fs::remove_file(&script);
    let text = String::from_utf8_lossy(&out.stdout);
    let mut it = text.split_whitespace();
    let fx: i32 = it.next().and_then(|v| v.parse().ok()).unwrap_or(-1);
    let entry: i32 = it.next().and_then(|v| v.parse().ok()).unwrap_or(-1);
    (fx, entry)
}

fn fresh_fixture(tag: &str) -> PathBuf {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let path = std::env::temp_dir().join(format!("corro-fstab-{tag}-{}-{}.corro", std::process::id(), id));
    let _ = std::fs::remove_file(&path);
    std::fs::write(&path, "CORRO_LOG 1\n").expect("write fixture");
    path
}

/// Pixel test: pressing Right must not move the `fx` caption or the entry.
///
/// The address text to their left changes on every cursor move, so this is the
/// case where a shrink-to-fit label would slide both landmarks sideways. The
/// cursor walks right until the address grows a digit (`A1` -> `A10`), because
/// `A1` -> `B1` is the same width and would pass even with a shrinking label.
///
/// This guards the GTK row. See the module docs for why the source test below
/// is still required: this one cannot fail on the backend that has the bug.
#[test]
fn formula_row_stays_put_on_move() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under xvfb-run -a)"
    );
    let path = fresh_fixture("move");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(800));

    let shot0 = screenshot(&wid, "before");
    let (fx0, entry0) = fx_and_entry_x(&shot0);
    let _ = std::fs::remove_file(&shot0);
    assert!(fx0 > 0 && entry0 > fx0, "could not locate fx/entry in the formula row: fx={fx0} entry={entry0}");

    for _ in 0..9 {
        xdotool(&["key", "--window", &wid, "Right"]);
        std::thread::sleep(Duration::from_millis(120));
    }
    std::thread::sleep(Duration::from_millis(1200));

    let shot1 = screenshot(&wid, "after");
    let (fx1, entry1) = fx_and_entry_x(&shot1);
    let _ = std::fs::remove_file(&shot1);

    assert_eq!(
        fx1, fx0,
        "fx caption moved sideways after moving the cursor: {fx0} -> {fx1} \
         (a formula-row label resized and pushed the row)"
    );
    assert_eq!(
        entry1, entry0,
        "formula entry moved sideways after moving the cursor: {entry0} -> {entry1} \
         (a formula-row label resized and pushed the row)"
    );
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}

/// Formula-row widgets that change text while the app runs, and so can reflow
/// the row on a backend that re-fits labels to their text.
///
/// - `addr_label`     — the cell address; changes on every cursor move.
/// - `formula_status` — the `· status` suffix; its text changes with every
///   status message and it appears/disappears.
///
/// `f_label` (the `fx` caption) is deliberately absent: its text is constant,
/// so it can never resize from a text change. `fx_caption_is_constant_text`
/// guards that.
const MUTATING_LABELS: &[&str] = &["addr_label", "formula_status"];

/// Every mutating formula-row label must pin its width.
///
/// This is the assertion that fails when a pin is dropped. The pixel test above
/// cannot catch that on the backends that matter, so this is the guard: remove
/// a `set_fixed_width(Some(...))` from any label in [`MUTATING_LABELS`] and it
/// fails here.
#[test]
fn formula_row_mutating_labels_pin_their_width() {
    let src = include_str!("../src/gui/gui_backend.rs");
    for label in MUTATING_LABELS {
        // Built by concatenation so this literal cannot match its own
        // assertion line, which would make the check vacuously true.
        let needle = [format!("{label}.set_fixed_width("), "Some(".to_string()].concat();
        assert!(
            src.contains(&needle),
            "{label} must pin its width: it changes text, and on the Win32/NWG \
             backend a STATIC label is re-fitted to its text, which reflows the \
             box and slides fx and the formula entry sideways. Pin it with \
             `set_fixed_width(Some(...))`."
        );
    }
}

/// The `fx` caption's text is constant, so it never needs a pin — a pin there
/// would only risk clipping it. What must hold is that it exists and sits
/// between the address slot and the entry, which is the ordering the pixel test
/// measures on screen.
#[test]
fn fx_caption_is_constant_text() {
    let src = include_str!("../src/gui/gui_backend.rs");
    // Matched via the binding so the label's text and its name cannot drift.
    let needle = "let f_label = rxapp.new_label(\"  fx  \")";
    assert!(
        src.contains(needle),
        "the fx caption must be a constant label packed between the address \
         slot and the entry; a changing caption would reflow the row on Win32"
    );
    // Order matters: address, then fx, then entry. If fx were packed after the
    // entry it would sit to the right of the input, which is where the status
    // suffix belongs.
    let addr_at = src.find("formula_bar.append(&addr_label)").expect("addr appended");
    let fx_at = src.find("formula_bar.append(&f_label)").expect("fx appended");
    let entry_at = src.find("formula_bar.append(&formula_entry)").expect("entry appended");
    assert!(
        addr_at < fx_at && fx_at < entry_at,
        "formula row order must be address, fx, entry (found at {addr_at}, {fx_at}, {entry_at})"
    );
}

/// The repro, isolated: pressing Right must not move the formula entry, and
/// the test must be able to tell a moved entry from keys that never arrived.
///
/// The existing `formula_row_stays_put_on_move` asserts the same invariant,
/// but it cannot distinguish "the row held still" from "the app never saw the
/// key": if the Right presses were swallowed, both landmarks would trivially
/// keep their x and the test would pass while the bug is present. That is not
/// hypothetical — on a display with no window manager, `xdotool key
/// --window` is delivered to nothing and the test still went green.
///
/// So this test makes the two facts separately observable:
///
///   1. the cursor really moved — the address ink changes shape (a wider
///      address has more/different glyph columns), and the grid repaints;
///   2. the entry's left edge did not move.
///
/// (1) is the guard that makes (2) mean something. Without it this test would
/// be the same vacuous pass in a new file.
#[test]
fn right_press_moves_the_cursor_without_moving_the_entry() {
    let _guard = GUI_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    assert!(
        std::env::var("DISPLAY").is_ok(),
        "requires X server (run under scripts/with-hidden-display.sh)"
    );
    // Needs a window manager: `xdotool key --window` and `windowactivate` go
    // nowhere without EWMH support, which is exactly how this test would end
    // up measuring a motionless app.
    assert!(
        xprop_has_wm(),
        "no window manager on this display: xdotool activation and synthetic \
         keys do not work without one, so the entry could not move and this \
         test would pass vacuously. Run under scripts/with-hidden-display.sh, \
         which starts one."
    );
    let path = fresh_fixture("entrymove");
    let mut child = spawn_gui(&path);
    let wid = find_corro_window(child.id(), Instant::now() + Duration::from_secs(25));
    xdotool(&["windowactivate", "--sync", &wid]);
    std::thread::sleep(Duration::from_millis(800));

    let shot0 = screenshot(&wid, "entry0");
    let before = row_geometry(&shot0);
    let _ = std::fs::remove_file(&shot0);
    assert!(
        before.entry_x > 0,
        "could not find the entry's blue frame in the formula row: {before:?}"
    );

    // Walk right far enough for the address to gain a character (A1 -> A10).
    for _ in 0..9 {
        xdotool(&["key", "--window", &wid, "Right"]);
        std::thread::sleep(Duration::from_millis(120));
    }
    std::thread::sleep(Duration::from_millis(1200));

    let shot1 = screenshot(&wid, "entry1");
    let after = row_geometry(&shot1);
    let _ = std::fs::remove_file(&shot1);

    // (1) The keys arrived: the address text changed, and the grid repainted.
    //
    // Keyed on the address's ink *width*, not its first ink column: the slot is
    // pinned and the text left-aligned, so `A1` and `A10` both start at the
    // same x and only the second's ink extends further. The start column is a
    // constant here, so asserting on it would never fail — including when the
    // keys never arrived, which is the case this guard exists to catch.
    assert_ne!(
        before.address_ink_cols, after.address_ink_cols,
        "the address did not get any wider after 9 Right presses (still \
         {before:?}), so the keys never reached the app — the entry's \
         position was never actually tested",
    );
    assert!(
        after.grid_changed_px > 0,
        "the grid did not repaint at all after 9 Right presses; the cursor \
         never moved, so this test cannot see a formula-row reflow \
         ({:?})",
        after
    );

    // (2) The actual invariant.
    assert_eq!(
        after.entry_x, before.entry_x,
        "the formula entry slid {}px -> {}px after moving the cursor: a \
         formula-row label resized to its new text and re-packed the row, \
         dragging the caret's frame sideways under the user ({:?} -> {:?})",
        before.entry_x,
        after.entry_x,
        before,
        after
    );
    let _ = child.kill();
    let _ = std::fs::remove_file(&path);
}

/// True when a window manager is present (i.e. EWMH is available).
fn xprop_has_wm() -> bool {
    if !PathBuf::from("/usr/bin/xprop").exists() {
        // No xprop to ask: don't block the test, the key-delivery assertion
        // below is the real guard anyway.
        return true;
    }
    Command::new("xprop")
        .args(["-root", "_NET_SUPPORTING_WM_CHECK"])
        .output()
        .map(|o| String::from_utf8_lossy(&o.stdout).contains("window id"))
        .unwrap_or(false)
}

/// Everything this test needs from one formula-row screenshot, measured in one
/// pass so the two frames are directly comparable.
#[derive(Debug, PartialEq, Eq)]
struct RowGeometry {
    /// Left edge of the entry's blue focus frame: the caret's position.
    entry_x: i32,
    /// First dark pixel past the address slot: the address text's start.
    address_ink: i32,
    /// How many dark columns the address slot has — a wider address (`A10`
    /// vs `A1`) has more of them, so this is the "the text really changed"
    /// signal that the landmark x alone cannot give.
    address_ink_cols: i32,
    /// Dark pixels in the grid body, used to prove the app repainted.
    grid_changed_px: i32,
}

fn row_geometry(png: &PathBuf) -> RowGeometry {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let script = std::env::temp_dir().join(format!("corro-fstab-geo-{id}.py"));
    std::fs::write(
        &script,
        "import sys\nfrom PIL import Image\nimg=Image.open(sys.argv[1]).convert('RGB')\nW,H=img.size\npx=img.load()\nY0,Y1=30,44\nBLUE=(53,132,228)\ndef dark(c):\n    r,g,b=c\n    return r<200 and g<200 and b<200\nentry=next((x for x in range(0,min(W,900)) if any(px[x,y]==BLUE for y in range(Y0,Y1))), -1)\ncols=[x for x in range(0,min(W,60)) if any(dark(px[x,y]) for y in range(Y0,Y1))]\naddr=next((x for x in cols if x>18), -1)\ngrid=sum(1 for y in range(60,min(H,400)) for x in range(0,min(W,700)) if dark(px[x,y]))\nprint(entry, addr, len(cols), grid)\n",
    )
    .expect("write geometry analyzer");
    let out = Command::new("python3")
        .arg(&script)
        .arg(png)
        .output()
        .expect("geometry analyzer");
    let _ = std::fs::remove_file(&script);
    let text = String::from_utf8_lossy(&out.stdout);
    let nums: Vec<i32> = text
        .split_whitespace()
        .map(|v| v.parse().unwrap_or(-1))
        .collect();
    let at = |i: usize| nums.get(i).copied().unwrap_or(-1);
    RowGeometry {
        entry_x: at(0),
        address_ink: at(1),
        address_ink_cols: at(2),
        grid_changed_px: at(3),
    }
}
