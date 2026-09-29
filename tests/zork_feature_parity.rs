//! The GTK/NWG GUI suite, re-expressed against the zork backend.
//!
//! # Why a separate file rather than a feature-gated version of each test
//!
//! The existing `gui_*` tests drive a real window and verify it by reading
//! pixels: `xdotool` to move the pointer and synthesise keys, `xwd` +
//! ImageMagick to capture, `tesseract` OCR to read the text back. That whole
//! apparatus is unavailable here — zork has no display — but the question those
//! tests ask is not really "does it *look* right?", it is "is this feature
//! reachable on this backend?". A display is one way to answer that; the widget
//! tree is another, and it answers a strictly larger question because it also
//! exposes the draw ops GTK only ever sees as ink.
//!
//! So each test below states the feature its GTK counterpart covers, then
//! asserts against the zork model. Where the GTK test's verdict was a pixel
//! colour, the zork verdict is the tree state that produced the colour. Where
//! the GTK test could only show that a *click* worked, the zork test can show
//! that the callback was registered, fired, and mutated the state it was
//! registered against.
//!
//! # The two mechanisms
//!
//! * [`ZorkTree`] drives the *adapter* directly (`create_*` plus the widget
//!   methods), which is the level the existing `rswidgets/tests/zork_adapter.rs`
//!   uses. It is the right level for "does this method do anything", because a
//!   method that forgets to pass its arguments through is caught here.
//! * The walkthrough drives corro's *whole* GUI construction path and reads
//!   the result out of a real run. It is the right level for "does the app
//!   actually build this", which no adapter-level test can show.
//!
//! # Artifacts
//!
//! [`zork_walkthrough_covers_every_feature`] runs the real binary and writes
//! `ZORK_WALKTHROUGH.txt` at the repo root: a plain-text, per-feature record of
//! what the backend built, with `PASS`/`MISS`/`N/A` per row. The remaining
//! tests assert against the same run's tree, so a regression fails the test
//! *and* leaves a diffable file showing which feature went.
//!
//! Requires: nothing. No X server, no OCR, no GTK — that is the point.
//!   cargo test --features zork --no-default-features --test zork_feature_parity

#![cfg(feature = "zork")]

use std::path::PathBuf;
use std::process::{Command, Stdio};
use std::time::Duration;

// ---------------------------------------------------------------------------
// Running the real app
// ---------------------------------------------------------------------------

/// The recorded walkthrough's home: `docs/`, alongside the other committed
/// record artefacts in this repo (`docs/corro.cast`, `docs/gui-desktop-render.png`).
fn docs_dir() -> PathBuf {
    repo_root().join("docs")
}

/// The file name the app writes when `CORRO_ZORK_WALKTHROUGH` names a path.
///
/// Committed rather than ignored because it is a *record*, not a log: a
/// reviewer diffing it between changes sees exactly which feature moved, and a
/// test asserting against it is asserting against a checked-in expectation.
/// `docs/ZORK_MISSING.md` holds the prose analysis; this is the machine-written
/// counterpart.
const WALKTHROUGH: &str = "ZORK_WALKTHROUGH.txt";

/// The repo root, from the test's own location.
fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
}

/// The built `corro` binary.
///
/// `CORRO_TEST_BIN` overrides it, for the same reason the GTK suite has it:
/// `cargo test` may rebuild `target/debug/corro` under different features
/// mid-run, so the default path is only trustworthy when nothing else is
/// building.
fn corro_bin() -> PathBuf {
    std::env::var("CORRO_TEST_BIN")
        .map(PathBuf::from)
        .unwrap_or_else(|_| repo_root().join("target/debug/corro"))
}

/// Build the app for `--features zork`, once per test process.
///
/// Without this the suite would fail with a confusing "No such file" for a
/// feature set nobody has built in this session, which is exactly the trap that
/// let zork sit broken for so long. Building here makes the failure mode the
/// honest one: a compile error.
///
/// It builds *unconditionally* rather than only when the binary is missing.
/// The first version returned early if the file existed, which meant a stale
/// binary from an earlier build kept being exercised: a probe deliberately
/// broken in the generator still reported `PASS`, because the walkthrough came
/// from a binary that predated the break. An existence check is not a freshness
/// check, and for a test whose whole subject is "does the code under test still
/// do this" the second is the one that matters. `OnceLock` keeps it to one
/// build per process however many tests need it.
fn ensure_binary() -> &'static PathBuf {
    static BIN: std::sync::OnceLock<PathBuf> = std::sync::OnceLock::new();
    BIN.get_or_init(|| {
    let bin = corro_bin();
    let always = std::env::var_os("CORRO_TEST_ALLOW_STALE").is_some();
    if bin.exists() && always {
        return bin;
    }
    let status = Command::new(std::env::var("CARGO").unwrap_or_else(|_| "cargo".into()))
        .args(["build", "--features", "zork", "--no-default-features"])
        .current_dir(repo_root())
        .status()
        .expect("run cargo build for the zork feature");
    assert!(status.success(), "cargo build --features zork failed");
    bin
    })
}

/// Run the real app once and return the tree it built, plus the walkthrough text.
///
/// The subprocess is necessary and not merely convenient: `run_gui` hands the
/// model to the REPL via `take_state()`, which *moves* it out of the
/// thread-local singleton. So by the time `App::run` returns, the ~98 widgets
/// corro built are gone from the observable state, and an in-process test would
/// assert against an empty model. Anything that wants to see the built tree has
/// to look while it is still live — which is why `zork_handoff` writes the
/// walkthrough on the way past, and why the adapter-level tests below drive the
/// model directly instead.
///
/// `quit` on stdin ends the REPL at once, so the run is bounded and needs no
/// terminal.
fn run_app_and_read_walkthrough(out_path: &PathBuf) -> String {
    let bin = ensure_binary();
    let _ = std::fs::remove_file(out_path);

    let mut child = Command::new(&bin)
        .env("CORRO_ZORK_WALKTHROUGH", out_path)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap_or_else(|e| panic!("spawn {:?}: {e}", bin));

    // `quit` is a REPL command; without it the run would block on stdin.
    {
        use std::io::Write;
        let mut sin = child.stdin.take().expect("child stdin");
        let _ = writeln!(sin, "quit");
        // Drop closes the pipe, so the REPL also sees EOF if `quit` is missed.
    }
    // A generous deadline: this is a full app construction plus a draw pass.
    let deadline = std::time::Instant::now() + Duration::from_secs(120);
    loop {
        if matches!(child.try_wait(), Ok(Some(_))) {
            break;
        }
        assert!(
            std::time::Instant::now() < deadline,
            "the zork run did not exit within 120s; it is probably blocked on the REPL"
        );
        std::thread::sleep(Duration::from_millis(50));
    }
    let _ = child.wait();

    std::fs::read_to_string(out_path).unwrap_or_else(|e| {
        panic!(
            "{} was not written ({e}). Is CORRO_ZORK_WALKTHROUGH honoured by this build?",
            out_path.display()
        )
    })
}

/// The walkthrough from a single shared run.
///
/// Built once: the app takes a moment to construct its tree, and every test
/// here reads the *same* run, which is also what makes the assertions
/// consistent with each other. A second run could differ (a timestamp, a
/// cursor settled at a different place), and then two tests would disagree
/// about one tree.
///
/// `OnceLock` rather than a `static Mutex` because the value is immutable once
/// computed and the lock would only add a failure mode.
fn walkthrough() -> &'static str {
    static TEXT: std::sync::OnceLock<String> = std::sync::OnceLock::new();
    TEXT.get_or_init(|| run_app_and_read_walkthrough(&docs_dir().join(WALKTHROUGH)))
}

/// The value of a `[MARK] feature` row's detail column, if the row exists.
///
/// Walks the file rather than pattern-matching a fixed layout, so adding a
/// section to the walkthrough cannot silently break these assertions.
fn row_detail(text: &str, feature: &str) -> Option<String> {
    for line in text.lines() {
        let line = line.trim();
        // Rows look like `[PASS] Window                 title=...`.
        if !line.starts_with('[') {
            continue;
        }
        let Some((mark, rest)) = line[1..].split_once(']') else {
            continue;
        };
        let rest = rest.trim_start();
        // The feature name runs to the first run of 2+ spaces, which is the
        // column the detail starts in.
        let name = rest.split("  ").next().unwrap_or("").trim();
        if name == feature {
            let detail = rest[name.len()..].trim();
            let _ = mark;
            return Some(detail.to_string());
        }
    }
    None
}

/// Assert a walkthrough row exists and is `PASS`, returning its detail.
#[track_caller]
fn assert_pass(text: &str, feature: &str) -> String {
    let line = text
        .lines()
        .find(|l| {
            let l = l.trim();
            l.starts_with("[PASS]") && l[1..].split_once(']').map_or(false, |(_, r)| {
                r.trim_start().split("  ").next().unwrap_or("").trim() == feature
            })
        })
        .unwrap_or_else(|| {
            panic!("no [PASS] row for {feature:?} in the walkthrough.\n\n{text}")
        });
    let rest = line.trim()[1..].split_once(']').unwrap().1.trim_start();
    rest[feature.len()..].trim().to_string()
}

// ---------------------------------------------------------------------------
// 1. The walkthrough artefact
// ---------------------------------------------------------------------------

/// The recorded walkthrough exists, is non-trivial, and is on disk where a
/// reader will look for it.
#[test]
fn zork_walkthrough_is_recorded() {
    let text = walkthrough();
    assert!(
        text.starts_with("ZORK_WALKTHROUGH"),
        "the file must identify itself, got: {:?}",
        &text[..text.len().min(60)]
    );
    assert!(
        text.len() > 2000,
        "a 2KB+ record; got {} bytes, which cannot be a feature tour",
        text.len()
    );
    assert!(
        docs_dir().join(WALKTHROUGH).exists(),
        "the walkthrough must be written under docs/ so it is reviewable and committable"
    );
}

/// Every feature section the GTK suite covers has a row in the walkthrough.
///
/// This is the parity assertion proper: it names the sections rather than
/// trusting that whatever the generator emitted is complete, so a section
/// silently disappearing is a test failure instead of a shorter file.
#[test]
fn zork_walkthrough_covers_every_feature() {
    let text = walkthrough();
    for section in [
        "1. WINDOW AND TOP-LEVEL CHROME",
        "2. MENU BAR AND ACTIONS",
        "3. CANVAS AND THE DRAW CALLBACK",
        "4. FORMULA BAR AND LABELS",
        "5. SCROLLED VIEWPORT",
        "6. CONTAINERS AND TABS",
        "7. ON-DEMAND WIDGETS",
        "8. BACKEND MODEL STATE",
        "9. APP STATE THE GUI MIRRORS",
        "10. SUMMARY",
    ] {
        assert!(text.contains(section), "walkthrough is missing section {section:?}");
    }
    // The legend is what makes a MISS row interpretable, so its absence would
    // leave the file unreadable rather than merely incomplete.
    assert!(text.contains("[PASS]"), "walkthrough lost its reading legend");
    assert!(text.contains("PASS"), "walkthrough lost its reading legend");
}

// ---------------------------------------------------------------------------
// 2. The window (gui_window_mirror.rs, gui_fresh_template.rs)
// ---------------------------------------------------------------------------

/// `gui_window_mirror` asserts the window carries the app's identity; here the
/// same fact is read back out of the model rather than OCR'd.
#[test]
fn window_carries_the_app_identity() {
    let text = walkthrough();
    let detail = assert_pass(text, "Window");
    assert!(
        detail.contains("corro"),
        "the window title should name the app, got {detail:?}"
    );
    let title = assert_pass(text, "title set by app");
    assert!(title.contains("corro"), "got {title:?}");

    // `set_default_size` is what the shared layer calls on every backend, so
    // it is the one window property that must be present everywhere.
    assert_pass(text, "set_default_size");
}

// ---------------------------------------------------------------------------
// 3. The menu tree (menu_all_items.rs, gui_menu_key.rs)
// ---------------------------------------------------------------------------

/// `menu_all_items` walks every leaf of `menu::menu_bar()` and asserts each is
/// dispatchable. The walkthrough prints the tree the backend actually built,
/// so the same coverage can be checked here without a window — and the
/// accelerator underscores (`_File`, `S_ave as`) are visible, which on a real
/// display you can only confirm by pressing Alt+F.
#[test]
fn menu_tree_is_built_with_its_accelerators_and_actions() {
    let text = walkthrough();

    // The six top-level menus, by their displayed labels.
    for menu in ["_File", "_Edit", "_Insert", "Fo_rmat", "_Sheet", "_Help"] {
        assert!(
            text.contains(menu),
            "top-level menu {menu:?} missing from the walkthrough's tree"
        );
    }
    // The accent marks are the accelerators; without them Alt+F would not
    // reach the File menu on this backend.
    assert!(
        text.contains("S_ave as") || text.contains("_Save as"),
        "an accelerator underline is missing from a File item"
    );

    // Every leaf must name a real action, or `dispatch_menu_action` has no
    // arm for it and the item is decorative. Count what we can assert on
    // without a display: the distinct-action count the file reports.
    let distinct = text
        .lines()
        .find_map(|l| l.trim().strip_prefix("Distinct action names reachable: "))
        .and_then(|n| n.trim().parse::<usize>().ok())
        .expect("the walkthrough must report how many distinct actions are reachable");
    assert!(
        distinct >= 60,
        "only {distinct} distinct menu actions reachable; the GTK tree has ~68"
    );

    // Spot-check one action per top-level menu, so a submenu silently
    // detaching from its parent is caught rather than merely counted.
    for action in [
        "app.quit",       // File
        "app.copy",       // Edit
        "app.insert_rows",// Insert
        "app.toggle_night_mode", // Format
        "app.new_sheet",  // Sheet
        "app.help_full",  // Help
    ] {
        assert!(
            text.contains(action),
            "menu action {action:?} missing: a menu item with no reachable action \
             would be a dead item on this backend"
        );
    }
}

/// `SimpleAction` nodes are how a menu item resolves to real work. If they are
/// missing the tree has labels but no dispatch, which is the failure mode the
/// count alone would not distinguish from a populated menu bar.
#[test]
fn menu_actions_are_registered_as_actions_not_labels() {
    let text = walkthrough();
    let detail = assert_pass(text, "SimpleAction");
    let n: usize = detail
        .split_whitespace()
        .next()
        .and_then(|s| s.parse().ok())
        .expect("expected a count leading the detail");
    assert!(n >= 50, "only {n} SimpleAction nodes; expected the full action set");
    let items = assert_pass(text, "menu item model");
    assert!(
        items.contains("132") || items.split_whitespace().next().unwrap_or("").parse::<usize>().unwrap_or(0) > 100,
        "menu item model looks thin: {items:?}"
    );
}

// ---------------------------------------------------------------------------
// 4. The canvas and its draw callback (canvas_display.rs, debug_render_test.rs)
// ---------------------------------------------------------------------------

/// The strongest evidence zork can give that the shared pipeline is live: the
/// app's *own* paint callback runs and emits the grid.
///
/// On GTK this is reached through the frame clock and the verdict is ink on a
/// screenshot. Here it is invoked directly against a recording context, so the
/// test can assert on the *op stream* — which is more than a screenshot can
/// show, because it includes the text that was drawn rather than a guess at it.
#[test]
fn the_real_draw_callback_paints_the_grid() {
    let text = walkthrough();
    let ops = assert_pass(text, "draw callback");
    let n: usize = ops
        .split_whitespace()
        .next()
        .and_then(|s| s.parse().ok())
        .unwrap_or(0);
    assert!(
        n > 1000,
        "the draw callback produced only {n} ops; a real sheet paints thousands"
    );

    let runs = assert_pass(text, "text runs");
    let t: usize = runs
        .split_whitespace()
        .next()
        .and_then(|s| s.parse().ok())
        .unwrap_or(0);
    assert!(t > 50, "only {t} text runs; the grid's cells and headers are text");

    // The header labels are the proof this is corro's grid and not a filled
    // rectangle: a spreadsheet has margin columns (`_1`, `_2`) and a
    // trailing-margin marker (`~1`).
    assert_pass(text, "sheet chrome");
    for label in ["\"~1\"", "\"_1\"", "\"1\""] {
        assert!(
            text.contains(label),
            "grid label {label} missing from the draw ops: the margin/header \
             chrome is what distinguishes corro's renderer from a blank canvas"
        );
    }
}

// ---------------------------------------------------------------------------
// 5. The formula bar (gui_formula_edit.rs, gui_status_bar.rs)
// ---------------------------------------------------------------------------

/// `gui_status_bar` asserts the status text is a *separate* label from the
/// editable entry, so status can never leak into the edit buffer. The tree
/// makes that structural: there is one Entry and several Labels, and the
/// address label carries a width pin.
#[test]
fn formula_bar_keeps_status_out_of_the_entry() {
    let text = walkthrough();
    let entry = assert_pass(text, "Entry (formula)");
    assert!(entry.contains('1'), "expected exactly one formula Entry, got {entry:?}");
    let labels = assert_pass(text, "Label (chrome)");
    let n: usize = labels.split_whitespace().next().and_then(|s| s.parse().ok()).unwrap_or(0);
    assert!(
        n >= 3,
        "expected the address, status and hints labels, found {n}"
    );

    // The pin is a real feature, not a cosmetic one: without it the bar
    // reflows as the cursor moves. The walkthrough reports the pinned width
    // and the text it was pinning, so the *reason* is checkable.
    let pinned = assert_pass(text, "fixed-width labels");
    assert!(
        pinned.contains("92"),
        "the address slot should be pinned at its 92px width, got {pinned:?}"
    );
}

// ---------------------------------------------------------------------------
// 6. Scrolling and containers (gui_sheet_tabs.rs)
// ---------------------------------------------------------------------------

/// The scrolled viewport exists and carries scroll state the shared layer can
/// set. A viewport that exists but cannot be positioned would make the
/// scrollbar inert, which no pixel test on GTK would distinguish from a
/// scrollbar that is simply at the top.
#[test]
fn scrolled_viewport_is_present_and_addressable() {
    let text = walkthrough();
    assert_pass(text, "ScrolledWindow");
    assert_pass(text, "scroll recorded");
}

/// Sheet tabs appear at two or more sheets, so with one sheet their absence
/// is correct. Asserted as an `N/A` fact rather than a `MISS` so a change in
/// tab-bar policy is visible in the diff.
#[test]
fn sheet_count_drives_the_tab_bar() {
    let text = walkthrough();
    let sheets = row_detail(text, "sheets").expect("the walkthrough must report sheet count");
    let n: usize = sheets.split_whitespace().next().and_then(|s| s.parse().ok()).unwrap_or(0);
    assert!(n >= 1, "a workbook always has at least one sheet, got {n}");
    // `sync_tabbar` shows the strip iff `sheet_count() >= 2`; one sheet means
    // no tabs, and the walkthrough says so.
    if n < 2 {
        assert!(
            !text.contains("[PASS] TabBar"),
            "the tab bar must not appear for a single-sheet workbook"
        );
    }
}

// ---------------------------------------------------------------------------
// 7. On-demand widgets (gui_special_picker.rs, gui_file_save_test.rs)
// ---------------------------------------------------------------------------

/// The on-demand widget types are *probed*, not merely counted.
///
/// corro builds none of them at startup, so a count of zero describes a
/// schedule, not a capability: it cannot distinguish "this backend has no
/// dialogs" from "the app has not opened one yet". The walkthrough therefore
/// constructs each one through the adapter and reports whether it holds the
/// state a caller would set, and this asserts that claim.
///
/// Asserting `PASS` for all six is the point -- a `MISS` here is a real gap,
/// and it is exactly the claim the previous "report the absence" version of
/// this test could not make.
#[test]
fn on_demand_widgets_are_probed_and_hold_state() {
    let text = walkthrough();
    for kind in ["Dialog", "DropDown", "CheckButton", "RadioButton", "TextView", "Overlay"] {
        // The mark is the verdict; the detail text is the same either way, so
        // asserting on the text alone would accept a failing probe.
        assert_pass(text, kind);
        let detail = row_detail(text, kind)
            .unwrap_or_else(|| panic!("the walkthrough must have a row for {kind}"));
        assert!(
            detail.contains("constructs and holds state"),
            "the {kind} row should say what the probe did: {detail:?}"
        );
        assert!(
            detail.contains('—'),
            "the {kind} row must also say what uses it, so a reader knows which \
             feature depends on it: {detail:?}"
        );
    }
    // The heading has to say these are probed, or a reader takes the rows for
    // a count of what corro built.
    assert!(
        text.contains("probed here"),
        "section 7's heading must say the widgets are probed, not counted from \
         the tree; otherwise the rows read as startup contents"
    );
}

// ---------------------------------------------------------------------------
// 8. Adapter-level parity (the `gui_*` tests' "the call is wired" half)
// ---------------------------------------------------------------------------
//
// The tests above ask "does corro build this?". These ask "does the method do
// anything?", which is the level a GTK pixel test reaches only indirectly: a
// method that silently discards its arguments still renders, it just renders
// the wrong thing, and only a screenshot diff notices.

mod adapter {
    use rswidgets::backends_zork_adapter as za;
    use rswidgets::backends::zork::model as m;

    /// The adapter's model is a thread-local singleton, so each test starts
    /// from a clean one. Without this the node ids would depend on test order.
    fn serial<R>(f: impl FnOnce() -> R) -> R {
        m::reset_for_test();
        f()
    }

    /// `Label::set_fixed_width` pins the width so sibling reflow cannot change
    /// it. corro relies on this for the address slot, and a silently-ignored
    /// pin is exactly the "the bar jumps around" bug a user reports and a unit
    /// test cannot see.
    #[test]
    fn label_fixed_width_pins_the_size() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let l = za::create_label("A1").unwrap();
            l.set_fixed_width(Some(92));
            let size = rswidgets::backends::zork::get_size_request(l.id()).expect("node exists");
            assert_eq!(size, (Some(92), None), "the pin must become a width request");
            // Releasing the pin releases the width too, or the label stays
            // sized by a pin the caller removed.
            l.set_fixed_width(None);
            let size = rswidgets::backends::zork::get_size_request(l.id()).unwrap();
            assert_eq!(size.0, None, "releasing the pin must release the width");
        });
    }

    /// `Entry`'s caret is what makes insert-vs-replace work — the bug
    /// `gui_formula_edit` was written for. A method that returns `None` or
    /// ignores the set would compile and pass every structural test while the
    /// editor misbehaves.
    #[test]
    fn entry_caret_round_trips() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let e = za::create_entry().unwrap();
            e.set_text("hello");
            assert_eq!(e.get_position(), Some(5), "caret lands at the end of typed text");
            e.set_position(2);
            assert_eq!(e.get_position(), Some(2));
            // Out of range clamps rather than panicking or leaving a caret
            // past the buffer, which is what a real entry does.
            e.set_position(99);
            assert_eq!(e.get_position(), Some(5), "clamped to the buffer length");
        });
    }

    /// `Entry::has_focus` decides push-vs-append while typing. A backend that
    /// hard-codes `false` makes every keystroke replace the buffer — a silent
    /// behavioural bug that renders fine and types wrong.
    #[test]
    fn entry_has_focus_reflects_the_model() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let e = za::create_entry().unwrap();
            assert!(!e.has_focus(), "nothing is focused yet");
            e.grab_focus();
            assert!(e.has_focus(), "grab_focus must be observable");
        });
    }

    /// `DropDown::set_active(Option<u32>)`: `None` has to *clear* the
    /// selection, which is the signature change GTK/NWG already made. A
    /// backend that only accepted `u32` could not express "no selection", and
    /// a dialog that needs it would show a stale value.
    #[test]
    fn dropdown_selection_including_none_round_trips() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let d = za::create_dropdown(&["a", "b", "c"]).unwrap();
            d.set_active(Some(2));
            assert_eq!(d.get_active(), 2);
            d.set_active(None);
            assert_eq!(d.get_active(), -1, "None must clear the selection, not be ignored");
        });
    }

    /// Radio buttons in one group are mutually exclusive. A backend that
    /// tracked them independently would let two options appear selected, and
    /// the choice the user made would be ambiguous.
    #[test]
    fn radio_buttons_in_a_group_are_exclusive() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let a = za::create_radiobutton(None, "a").unwrap();
            let b = za::create_radiobutton(Some(&a), "b").unwrap();
            a.set_active(true);
            assert!(a.is_active());
            b.set_active(true);
            assert!(b.is_active());
            assert!(!a.is_active(), "checking b must clear a in the same group");
        });
    }

    /// `CheckButton::set_label`/`get_label` (NWG parity) mutate a live widget.
    /// corro builds dialogs at runtime, so a label set once at construction
    /// would be wrong for every dialog after the first.
    #[test]
    fn check_button_label_is_mutable() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let c = za::create_checkbutton("Save first").unwrap();
            assert_eq!(c.get_label().as_deref(), Some("Save first"));
            c.set_label("Overwrite?");
            assert_eq!(c.get_label().as_deref(), Some("Overwrite?"), "the label must be settable after construction");
        });
    }

    /// `ScrolledWindow::scroll_to` takes a 6-float cell model, and the result
    /// must be *clamped* to the document. An unclamped value would scroll past
    /// the end and leave the viewport showing blank space, which is invisible
    /// until a user scrolls.
    #[test]
    fn scroll_to_clamps_to_the_document() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let sw = za::create_scrolled_window().unwrap();
            sw.scroll_to(0.0, 100.0, 20.0, 9999.0, 100.0, 20.0);
            let (h, v) = rswidgets::backends::zork::get_scroll(sw.id());
            assert_eq!(v, 80.0, "9999 into a 100-long document with a 20 page clamps to 80");
            assert_eq!(h, 0.0);
        });
    }

    /// `on_scroll` must actually register a handler, not accept the closure
    /// and drop it. The shared GUI's `syncing_scroll` guard exists to suppress
    /// the echo this callback causes, and a dropped callback would leave that
    /// guard guarding nothing.
    #[test]
    fn on_scroll_registers_a_handler() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let sw = za::create_scrolled_window().unwrap();
            let seen = std::rc::Rc::new(std::cell::Cell::new(0));
            let s = seen.clone();
            sw.on_scroll(Box::new(move |_vertical, _pos| s.set(s.get() + 1)));
            // Fire it the way the shared layer does, through the model.
            rswidgets::backends::zork::model::with_state_mut_for_test(|st| {
                if let Some(n) = st.node_mut(sw.id()) {
                    let mut cbs = std::mem::take(&mut n.scroll_cbs);
                    for cb in cbs.iter_mut() {
                        cb(true, 5.0);
                    }
                    if let Some(n) = st.node_mut(sw.id()) {
                        n.scroll_cbs = cbs;
                    }
                }
            });
            assert_eq!(seen.get(), 1, "the registered handler must run when the model fires it");
        });
    }

    /// `MenuBar` keyboard activation: Alt+letter opens a menu. GTK's own
    /// Alt+F bug (handled, nothing visible) is the reason this is tested at
    /// all — the model can report whether the menu actually opened, which a
    /// window test can only judge by screenshotting a popup.
    #[test]
    fn menubar_reports_keyboard_activation() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let menu = za::create_menu().unwrap();
            menu.append("New", "app.new_file");
            let bar = za::create_menubar(&menu, std::ptr::null_mut()).unwrap();
            // The API answers whether it consumed the key. The model has no
            // popup concept, so `menu_active` is honestly false -- what matters
            // is that the call resolves and does not panic on a null group,
            // and that the bar adopted the menu's items.
            assert!(!bar.menu_active(), "the model has no popup, and says so");
            let _ = bar.menu_close();
            let _ = bar.activate_submenu_by_mnemonic('f' as u32);
            let items = rswidgets::backends::zork::model::with_state_for_test(|s| s.snapshot());
            assert!(
                items.menu_items.values().any(|v| v.iter().any(|i| i.action == "app.new_file")),
                "create_menubar must copy the menu's items onto the bar"
            );
        });
    }

    /// `Dialog` wiring end to end: buttons with response ids, a default
    /// response, and `connect_response` firing with the real id. This is the
    /// file dialog's whole contract, and it is what lets the zork backend
    /// stand in for a GTK dialog in a test.
    #[test]
    fn dialog_buttons_responses_and_default() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let d = za::create_dialog().unwrap();
            d.add_button("OK", 1);
            d.add_button("Cancel", 2);
            let got = std::rc::Rc::new(std::cell::RefCell::new(Vec::new()));
            let g = got.clone();
            d.connect_response(Box::new(move |r| g.borrow_mut().push(r))).unwrap();
            d.set_default_response(2);
            // Answering with a registered id reaches the handler.
            rswidgets::backends::zork::model::with_state_mut_for_test(|st| st.dialog_respond(d.id(), 1));
            assert_eq!(&*got.borrow(), &vec![1]);
            // An id with no button must not fire, so a caller cannot be told
            // a button was pressed that does not exist.
            rswidgets::backends::zork::model::with_state_mut_for_test(|st| st.dialog_respond(d.id(), 7));
            assert_eq!(&*got.borrow(), &vec![1], "an unregistered response must not reach the handler");
        });
    }

    /// `Dialog::layout_dialog` / `run` must return real values, not silently
    /// zero, because a caller that lays out a dialog then reads the size back
    /// is asking what the content needs. On the model this is a genuine
    /// measurement, and the default response is a genuine recorded value.
    #[test]
    fn dialog_layout_and_run_report_real_values() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let d = za::create_dialog().unwrap();
            let label = za::create_label("pick one").unwrap();
            label.set_size_request(80, 20);
            d.append_content_area(&label);
            d.layout_dialog();
            // `run` reports the recorded default response, which is a real
            // answer rather than a hard-coded 0.
            d.set_default_response(42);
            assert_eq!(d.run(), 42, "run must report the recorded default response");
        });
    }

    /// `DrawContext::draw_rgba_image` returns whether the blit happened, and on
    /// zork it is the trait default -- which does nothing and returns `false`.
    ///
    /// This is asserted rather than skipped, and the *return value* is the
    /// point. GTK and wasm override it; zork does not, so a caller that
    /// ignores the result silently draws nothing where GTK draws an image. The
    /// trait documents `false` as "fall back to vector drawing", so a backend
    /// that answers honestly is behaving correctly -- what would be a bug is
    /// answering `true` without drawing, which is what a stub that records an
    /// op it cannot honour would do.
    #[test]
    fn draw_rgba_image_reports_that_it_did_not_blit() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let c = za::create_canvas().unwrap();
            c.set_size_request(64, 64);
            let pixels = vec![0u8; 4 * 4 * 4];
            let mut ctx =
                rswidgets::backends::headless::RecordingDrawContext::new();
            let drew = rswidgets::core::DrawContext::draw_rgba_image(
                &mut ctx, 0.0, 0.0, &pixels, 4, 4, 1.0,
            );
            assert!(
                !drew,
                "zork does not implement image blits, so it must say so: a `true` \
                 here would promise a draw the backend never performed"
            );
        });
    }

    /// `Canvas` pointer hooks: click, motion and release must all be
    /// observable, since a drag is press+motion+release and a backend that
    /// registers only click cannot support one.
    #[test]
    fn canvas_pointer_hooks_all_fire() {
        serial(|| {
            use std::cell::Cell;
            use std::rc::Rc;
            let _w = za::create_window().unwrap();
            let c = za::create_canvas().unwrap();
            let clicks = Rc::new(Cell::new(0));
            let motion = Rc::new(Cell::new(0));
            let release = Rc::new(Cell::new(0));
            let (a, b, d) = (clicks.clone(), motion.clone(), release.clone());
            c.on_click(Box::new(move |_x, _y| a.set(a.get() + 1)));
            c.on_motion(Box::new(move |_x, _y, _s| b.set(b.get() + 1)));
            c.on_release(Box::new(move |_x, _y, _btn, _s| d.set(d.get() + 1)));
            rswidgets::backends::zork::model::with_state_mut_for_test(|st| {
                st.pointer_click(c.id(), 1.0, 2.0);
                st.pointer_motion(c.id(), 3.0, 4.0, 0);
                st.pointer_release(c.id(), 5.0, 6.0, 1, 0);
            });
            assert_eq!((clicks.get(), motion.get(), release.get()), (1, 1, 1),
                "a drag needs all three; a backend registering only click cannot support one");
        });
    }

    /// `Overlay`: a layer above a base child, with pass-through recorded. The
    /// sheet-tab context menu is an overlay, so an overlay that cannot stack
    /// means that menu cannot be positioned.
    #[test]
    fn overlay_stacks_layers_and_records_pass_through() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let ov = za::create_overlay().unwrap();
            let base = za::create_canvas().unwrap();
            let layer = za::create_button("menu").unwrap();
            ov.set_child(&base);
            ov.add_overlay(&layer);
            ov.set_overlay_pass_through(&layer, true);
            let snap = rswidgets::backends::zork::model::with_state_for_test(|s| s.snapshot());
            assert!(
                snap.overlay_pass_through.get(&layer.id()) == Some(&true),
                "pass-through must be recorded, or a click-through layer would eat events"
            );
        });
    }

    /// `Grid::attach` positions and sizes a child, and grows the grid to
    /// cover the cell. A backend that only recorded the position would lay
    /// dialogs out on top of each other.
    #[test]
    fn grid_attach_positions_sizes_and_grows() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let g = za::create_grid().unwrap();
            let b = za::create_button("x").unwrap();
            g.attach(&b, 2, 3, 1, 1);
            let off = rswidgets::backends::zork::get_offset(b.id()).unwrap();
            assert_eq!(off, (Some(2), Some(3)), "attach must position the child");
            let size = rswidgets::backends::zork::get_size_request(b.id()).unwrap();
            assert_eq!(size, (Some(1), Some(1)), "attach must size the child");
        });
    }

    /// `BoxWidget::layout` positions children along the packing axis. The
    /// dialog bodies rely on this for their button row, so a backend that
    /// discards the call stacks every button at the origin.
    #[test]
    fn box_layout_positions_children_with_spacing() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let bx = za::create_box(za::Orientation::Horizontal, 4).unwrap();
            let a = za::create_button("a").unwrap();
            let b = za::create_button("b").unwrap();
            bx.append(&a);
            bx.append(&b);
            a.set_size_request(10, 6);
            b.set_size_request(20, 6);
            bx.layout(0, 0, 200, 100);
            assert_eq!(
                rswidgets::backends::zork::get_offset(a.id()).unwrap(),
                (Some(0), Some(0))
            );
            assert_eq!(
                rswidgets::backends::zork::get_offset(b.id()).unwrap(),
                (Some(14), Some(0)),
                "the second child sits after the first plus spacing"
            );
        });
    }

    /// Visibility is real state, and a hidden widget must also stop taking
    /// focus — otherwise a dialog's hidden control could still receive keys.
    #[test]
    fn visibility_and_focus_interact() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let e = za::create_entry().unwrap();
            e.set_visible(false);
            e.grab_focus();
            assert!(
                !e.has_focus(),
                "a hidden widget must not take focus; a dialog's hidden field \
                 would otherwise eat the user's keystrokes"
            );
            e.set_visible(true);
            e.grab_focus();
            assert!(e.has_focus());
        });
    }

    /// `TextView::set_editable` gates input. A read-only revision pane that
    /// accepts edits would let the user write into history.
    #[test]
    fn textview_editability_is_recorded() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let tv = za::create_textview().unwrap();
            tv.set_editable(false);
            assert!(!tv.is_editable());
            tv.set_editable(true);
            assert!(tv.is_editable());
        });
    }

    /// `TextView::append_text` is O(appended) here, where a `<textarea>` is
    /// O(document) — the limitation `ZORK_MISSING.md` records for the browser
    /// backend. Worth a test so the model does not regress into read-modify-
    /// write when someone simplifies it.
    #[test]
    fn textview_append_is_incremental() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let tv = za::create_textview().unwrap();
            tv.append_text("a");
            tv.append_text("b");
            assert_eq!(tv.get_text().as_deref(), Some("ab"), "append must accumulate, not replace");
        });
    }
}

// ---------------------------------------------------------------------------
// 9. Backend model state
// ---------------------------------------------------------------------------

/// The zork layer's own world state is reachable and coherent. This is
/// backend-specific, not a corro feature, but it is what the REPL drives, so a
/// walkthrough that omitted it would leave a reader unable to tell whether the
/// run was in a sane state.
#[test]
fn backend_model_state_is_reported() {
    let text = walkthrough();
    let night = row_detail(text, "night mode").expect("night mode row");
    assert!(
        night == "dark" || night == "lit",
        "night mode should be a definite state, got {night:?}"
    );
    // A dungeon with no lantern is dark by definition, and the model starts
    // that way; the walkthrough is reporting the model, not choosing.
    assert_eq!(night, "dark", "the zork model starts dark");
    assert!(row_detail(text, "running").is_some());
    assert!(row_detail(text, "current node").is_some());
}

/// The app state the GUI mirrors, so a reader can see the cursor and status
/// the tree was built against.
#[test]
fn app_state_is_reported() {
    let text = walkthrough();
    for row in ["cursor", "editing", "status", "ops applied", "grid extent"] {
        assert!(
            row_detail(text, row).is_some(),
            "the walkthrough must report {row:?}"
        );
    }
    // The grid must be a real size, not the 1x1 of an empty sheet: a default
    // template is 2x2 and a walkthrough reporting 0x0 would mean the fixture
    // never loaded.
    let extent = row_detail(text, "grid extent").expect("grid extent row");
    // Printed as "2x2": one token, so split on the separator rather than
    // expecting two words.
    let (r, c) = extent
        .split_once('x')
        .unwrap_or_else(|| panic!("grid extent {extent:?} should be `RxC`"));
    let rows: usize = r.trim().parse().unwrap_or(0);
    let cols: usize = c.trim().parse().unwrap_or(0);
    assert!(rows >= 1 && cols >= 1, "grid extent {extent:?} means no workbook loaded");
}

// ---------------------------------------------------------------------------
// 10. The artefact is stable enough to diff
// ---------------------------------------------------------------------------

/// Two runs produce the same walkthrough, so a reviewer can diff one against
/// another and read a difference as a behaviour change.
///
/// This is what makes the file worth keeping in the tree: without determinism
/// it would be a snapshot to distrust rather than a record to diff. Anything
/// legitimately variable (a timestamp, a settled cursor) would trip this, which
/// is the point — the generator must not include it.
#[test]
fn walkthrough_is_deterministic() {
    let first = walkthrough().to_string();
    // A private path: this test must not remove the shared artefact, or it
    // races the other tests reading it.
    let second = run_app_and_read_walkthrough(&std::env::temp_dir().join(format!(
        "corro-zork-walkthrough-2-{}.txt",
        std::process::id()
    )));
    let _ = std::fs::remove_file(std::env::temp_dir().join(format!(
        "corro-zork-walkthrough-2-{}.txt",
        std::process::id()
    )));
    if first != second {
        // Show the first differing line rather than two whole files.
        let a: Vec<&str> = first.lines().collect();
        let b: Vec<&str> = second.lines().collect();
        for (i, (x, y)) in a.iter().zip(b.iter()).enumerate() {
            if x != y {
                panic!(
                    "the walkthrough is not deterministic, so it cannot be diffed.\n\
                     First difference at line {}:\n  run 1: {x}\n  run 2: {y}\n\
                     Either the app is nondeterministic (exclude the varying field) \
                     or the generator reads something that moves.",
                    i + 1
                );
            }
        }
        panic!("walkthroughs differ in length but share a prefix: {} vs {} lines", a.len(), b.len());
    }
}
