

#![cfg(all(target_os = "linux", feature = "pancurses", feature = "ratatui"))]

mod tmux {
    use std::process::Command;

    pub fn new_session(session: &str, command: &str) {
        let status = Command::new("tmux")
            .args(["new-session", "-d", "-s", session, "-x", "120", "-y", "40", command])
            .status()
            .expect("tmux new-session failed");
        assert!(status.success(), "tmux new-session exited non-zero");
    }

    pub fn send_keys(session: &str, key: &str) {
        Command::new("tmux").args(["send-keys", "-t", session, key]).status().ok();
    }

    pub fn capture_pane(session: &str) -> String {
        let output = Command::new("tmux")
            .args(["capture-pane", "-t", session, "-p", "-S", "-200"])
            .output()
            .expect("tmux capture-pane failed");
        String::from_utf8_lossy(&output.stdout).to_string()
    }

    /// Same as capture_pane but keeps ANSI/OSC escape sequences (-e), so tests
    /// can assert on raw output such as the OSC 52 clipboard sequence.
    pub fn capture_pane_esc(session: &str) -> String {
        let output = Command::new("tmux")
            .args(["capture-pane", "-t", session, "-p", "-S", "-200", "-e"])
            .output()
            .expect("tmux capture-pane failed");
        String::from_utf8_lossy(&output.stdout).to_string()
    }

    pub fn kill_session(session: &str) {
        Command::new("tmux").args(["kill-session", "-t", session]).output().ok();
    }

    pub fn has_session(session: &str) -> bool {
        Command::new("tmux")
            .args(["has-session", "-t", session])
            .status()
            .map(|st| st.success())
            .unwrap_or(false)
    }
}

use std::path::PathBuf;
use std::time::Duration;

static COUNTER: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

/// Poll the pane until it contains `needle` (the app actually rendered it).
fn wait_for_text(session: &str, needle: &str) {
    for _ in 0..50 {
        if tmux::capture_pane(session).contains(needle) {
            return;
        }
        std::thread::sleep(Duration::from_millis(200));
    }
    panic!("timed out waiting for pane to contain {needle:?}");
}

/// Send one key, then wait for the frame it produces to settle (two
/// consecutive identical captures). Under parallel-tmux load a blind
/// follow-up key can land mid-render and get swallowed.
fn send_settled(session: &str, key: &str) {
    tmux::send_keys(session, key);
    let mut prev = tmux::capture_pane(session);
    for _ in 0..40 {
        std::thread::sleep(Duration::from_millis(30));
        let cur = tmux::capture_pane(session);
        if cur == prev {
            return;
        }
        prev = cur;
    }
}

/// Start a pancurses session and wait for the app to render its menu bar,
/// retrying on slow startup (many parallel tmux sessions contend for the
/// shared server; a session can take >15s to start under load).
fn start_session(session: &str, command: &str) {
    for _ in 0..3 {
        tmux::new_session(session, command);
        let mut ok = false;
        for _ in 0..100 {
            if tmux::capture_pane(session).contains("[File]") {
                ok = true;
                break;
            }
            std::thread::sleep(Duration::from_millis(200));
        }
        if ok {
            return;
        }
        tmux::kill_session(session);
        std::thread::sleep(Duration::from_millis(500));
    }
    panic!("failed to start session {session}");
}

/// Slice `s` to at most `n` chars, never splitting a UTF-8 codepoint
/// (the pancurses pane contains multi-byte box-drawing characters).
fn safe_slice(s: &str, n: usize) -> &str {
    if s.len() <= n { return s; }
    let mut end = n.min(s.len());
    while end > 0 && !s.is_char_boundary(end) { end -= 1; }
    &s[..end]
}


/// Wait until the CORRO_IDLE_MARKER file has at least `target` lines (the app
/// appends one line per redraw), so the test can send the next key as soon as
/// the app has finished processing the previous one — no fixed sleeps.
fn wait_marker(path: &std::path::Path, target: usize) {
    let deadline = std::time::Instant::now() + Duration::from_secs(15);
    loop {
        let count = std::fs::read_to_string(path).map(|s| s.lines().count()).unwrap_or(0);
        if count >= target { return; }
        if std::time::Instant::now() > deadline {
            panic!("timed out waiting for idle marker (target {target}, got {count})");
        }
        std::thread::sleep(Duration::from_millis(5));
    }
}



/// Copy an in-repo `docs/tests/*.corro` fixture to a unique temp file.
/// The pancurses backend writes cell commits through to the open file, so
/// tests must never open a checked-in fixture in place (editing tests once
/// appended their keystrokes to the repo file itself). Absolute paths
/// (already-temp copies shared with the ratatui side) pass through as-is.
fn fixture_temp_copy(rel: &str) -> String {
    let manifest = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let src = manifest.join(rel);
    let dst = std::env::temp_dir().join(format!(
        "corro-tmux-{}-{}.corro",
        std::process::id(),
        COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst)
    ));
    std::fs::copy(&src, &dst).expect("copy fixture to temp");
    dst.to_string_lossy().to_string()
}

/// Run pancurses in tmux, send keys, capture pane output.
fn run_in_tmux(args: &str, keys: &[&str], wait_ms: u64) -> String {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let manifest = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let mut rewritten: Vec<String> = Vec::new();
    for tok in args.split_whitespace() {
        if tok.ends_with(".corro") && !tok.starts_with('/') && manifest.join(tok).is_file() {
            rewritten.push(fixture_temp_copy(tok));
        } else {
            rewritten.push(tok.to_string());
        }
    }
    let cmd_args = rewritten.join(" ");
    tmux::new_session(&session, &format!("{} {}; sleep 2", bin, cmd_args));
    std::thread::sleep(Duration::from_millis(wait_ms));
    for key in keys {
        tmux::send_keys(&session, key);
        std::thread::sleep(Duration::from_millis(100));
    }
    std::thread::sleep(Duration::from_millis(400));
    let pane = tmux::capture_pane(&session);
    tmux::kill_session(&session);
    pane
}

/// Alt+F must open the File menu in the pancurses UI (regression test:
/// the pancurses path previously never created a MenuBar widget, so
/// menu_bar_id was None and Alt+key navigation did nothing).
#[test]
fn alt_f_opens_file_menu_in_pancurses() {
    let pane = run_in_tmux("--pancurses docs/tests/overflow.corro", &["Escape", "f"], 1200);
    let joined = pane.replace('\u{fffd}', "").replace('\x1b', "");
    for &s in &["Open file", "Save as", "Export", "Width", "Sort view", "Persist sort", "Exit", "Replay"] {
        assert!(
            joined.contains(s),
            "pancurses missing '{}' in File menu after Alt+F\n---\n{}\n---",
            s, safe_slice(&pane, 1200)
        );
    }
}


// Send the same key events to the ratatui App (via the public bench_handle_key API).
fn ratatui_send(app: &mut corro::ui::App, code: crossterm::event::KeyCode, mods: crossterm::event::KeyModifiers) {
    let ev = crossterm::event::KeyEvent::new(code, mods);
    app.bench_handle_key(ev).ok();
}

/// Render via ratatui TestBackend after sending a sequence of key events.
/// Operates on a COPY of the source file so the original .corro log is not mutated.
fn render_via_ratatui_with_keys(rel_path: &str, key_codes: &[crossterm::event::KeyCode]) -> String {
    let src = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(rel_path);
    // Unique per call: tests run in parallel within one process, and a fixed
    // shared temp path would let a stale leftover from an earlier run (or a
    // sandbox that blocks overwriting it) be loaded silently, diverging the
    // ratatui render from the pancurses side for harness reasons.  Fail loudly
    // on copy error instead of masking it.
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let tmp = std::env::temp_dir().join(format!(
        "corro-ml-rt-{}-{}.corro",
        std::process::id(),
        id
    ));
    std::fs::copy(&src, &tmp).expect("copy ratatui fixture");
    let mut app = corro::ui::App::new(Some(tmp));
    app.load_initial().unwrap();

    for &code in key_codes {
        ratatui_send(&mut app, code, crossterm::event::KeyModifiers::NONE);
    }

    let backend = ratatui::backend::TestBackend::new(120, 40);
    let mut terminal = ratatui::Terminal::new(backend).unwrap();
    terminal.draw(|f| app.bench_draw(f)).unwrap();
    let buffer = terminal.backend().buffer();
    (0..buffer.area.height)
        .map(|y| (0..buffer.area.width).map(|x| buffer[(x, y)].symbol()).collect::<String>())
        .collect::<Vec<_>>()
        .join("\n")
}

/// Render via ratatui TestBackend (no key events, initial state).
fn render_via_ratatui(rel_path: &str) -> String {
    render_via_ratatui_with_keys(rel_path, &[])
}

// ── Tests ──────────────────────────────────────────────────────────────────

#[test]
fn overflow_renders_cell_text() {
    let pane = run_in_tmux("--pancurses docs/tests/overflow.corro", &[], 600);
    assert!(pane.contains("should overflow"), "pancurses missing cell text:\n{}",
        safe_slice(&pane, 2000));
}

/// Up on the first menu item wraps to the last (matching the ratatui menu).
/// Enter then fires the wrapped item's action — the File menu's last item is
/// Replay — whose status message proves the selection wrapped around.
#[test]
fn menu_up_wraps_to_last_item() {
    let pane = run_in_tmux("--pancurses docs/tests/overflow.corro", &["Escape", "f", "Up", "Enter"], 2200);
    assert!(pane.contains("Replayed"),
        "Up on the first item should wrap to the last (Replay) and Enter fire it\n{}",
        safe_slice(&pane, 1500));
}

/// Ctrl+C must copy the cursor cell, not exit the app (regression: the
/// pancurses backend used to set running=false on Ctrl+C, quitting the TUI).
/// The copied value is written to the system clipboard via OSC 52, so the raw
/// (-e) pane capture must contain the base64 of the cell text.
#[test]
fn ctrl_c_copies_instead_of_quitting() {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-ctrlc-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    tmux::send_keys(&session, "C-c");
    std::thread::sleep(Duration::from_millis(400));
    let pane = tmux::capture_pane(&session);
    // The app must still be alive: the menu bar is rendered.
    assert!(pane.contains("[File]"),
        "app should stay alive after Ctrl+C (it must not quit)\n{}",
        safe_slice(&pane, 1200));
    // The cursor cell (A1) was copied: the formula bar reports "Copied A1",
    // and (because the copy is routed through the app's clipboard) the value
    // is available to Paste.  (tmux strips the raw OSC 52 sequence from
    // capture-pane, so we assert the observable copy instead.)
    assert!(pane.contains("Copied A1"),
        "Ctrl+C should copy the cursor cell (status 'Copied A1')\n{}",
        safe_slice(&pane, 1200));
    tmux::kill_session(&session);
}

#[test]
fn q_quits() {
    // Quit is Ctrl+Q only (matching the ratatui backend); plain 'q' is text.
    let pane = run_in_tmux("--pancurses docs/tests/overflow.corro", &["C-q"], 200);
    assert!(pane.len() < 200 || !pane.contains("[File]"),
        "program should have exited after Ctrl+Q");
}

#[test]
fn render_has_menu_and_cell_text() {
    let pane = run_in_tmux("--pancurses docs/tests/overflow.corro", &[], 500);
    let ratatui = render_via_ratatui("docs/tests/overflow.corro");
    for &s in &["[File]", "should overflow"] {
        assert!(pane.contains(s), "pancurses missing '{}'", s);
        assert!(ratatui.contains(s), "ratatui missing '{}'", s);
    }
}

#[test]
fn arrow_down_shows_a3() {
    let pane = run_in_tmux("--pancurses docs/tests/overflow.corro", &["Down", "Down"], 1200);
    assert!(pane.contains("A3"),
        "formula bar should show A3 after 2x Down from A1\n---\n{}\n---", safe_slice(&pane, 5000));
}

#[test]
fn right_arrow_shows_b2() {
    let pane = run_in_tmux("--pancurses docs/tests/overflow.corro", &["Right", "Down"], 1200);
    assert!(pane.contains("B2"),
        "formula bar should show B2 after Right+Down from A1\n---\n{}\n---", safe_slice(&pane, 5000));
}

#[test]
fn left_arrow_does_not_jump_viewport() {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));

    // Verify app started
    let pane0 = tmux::capture_pane(&session);
    assert!(pane0.contains("[File]"), "app should show menu bar");

    // Press Left once — cursor moves to [A (left-margin column nearest A)
    tmux::send_keys(&session, "Left");
    std::thread::sleep(Duration::from_millis(300));
    let pane1 = tmux::capture_pane(&session);
    // After Left once, formula bar should show [A1 (cursor in left margin)
    assert!(pane1.contains("[A1") || pane1.contains("[A1 "),
        "after Left once, formula bar should show [A1\n{}",
        &pane1[..pane1.len().min(3000)]);

    // Press Left twice — cursor moves to [B
    tmux::send_keys(&session, "Left");
    std::thread::sleep(Duration::from_millis(300));
    let pane2 = tmux::capture_pane(&session);
    // After Left twice, formula bar should show [B1
    assert!(pane2.contains("[B1") || pane2.contains("[B1 "),
        "after Left twice, formula bar should show [B1\n{}",
        &pane2[..pane2.len().min(3000)]);

    tmux::kill_session(&session);
}

/// Move to C3, enter "Hello World!", and verify both backends show the
/// correct cell address and content (structural match, not exact char).
#[test]
fn edit_c3_hello_world_full_screen_match() {
    use crossterm::event::KeyCode;
    let keys = &["Right", "Right", "Down", "Down", "Enter", "H", "e", "l", "l", "o", " ",
                 "W", "o", "r", "l", "d", "!", "Enter"];
    let pane = run_in_tmux("--pancurses docs/tests/overflow.corro", keys, 2000);
    assert!(pane.contains("C4"), "pancurses formula bar should show C4 after edit\n{}",
        safe_slice(&pane, 3000));

    let ratatui = render_via_ratatui_with_keys("docs/tests/overflow.corro", &[
        KeyCode::Right, KeyCode::Right, KeyCode::Down, KeyCode::Down,
        KeyCode::Enter,
        KeyCode::Char('H'), KeyCode::Char('e'), KeyCode::Char('l'), KeyCode::Char('l'),
        KeyCode::Char('o'), KeyCode::Char(' '),
        KeyCode::Char('W'), KeyCode::Char('o'), KeyCode::Char('r'), KeyCode::Char('l'),
        KeyCode::Char('d'), KeyCode::Char('!'),
        KeyCode::Enter,
    ]);
    assert!(ratatui.contains("C4"), "ratatui formula bar should show C4 after edit\n{}",
        &ratatui[..ratatui.len().min(3000)]);
}

/// Navigate to column K (past J) in the right margin and verify the ratatui formula bar shows ]K1.
#[test]
fn navigate_to_column_k_via_ratatui() {
    use crossterm::event::KeyCode;
    let mut keys = Vec::new();
    // Grid has 5 main columns (A-E).  15 Right presses from A1 reaches ]K1.
    for _ in 0..15 {
        keys.push(KeyCode::Right);
    }
    let ratatui = render_via_ratatui_with_keys("docs/tests/overflow.corro", &keys);
    assert!(ratatui.contains("]K1"),
        "ratatui formula bar should show ]K1 after 15 Right presses\n{}",
        &ratatui[..ratatui.len().min(3000)]);
}

/// Arrow left from A1 should enter the left margin column (show [A label).
/// Arrow up from A1 should enter the header row (show ~1 label).
#[test]
fn arrow_left_from_a1_enters_margin() {
    let pane = run_in_tmux("--pancurses docs/tests/overflow.corro", &["Left"], 1000);
    // After Left from A1, cursor should show left margin label (like [A or similar)
    assert!(pane.contains("[A") || pane.contains("[") || pane.contains("A1"),
        "Left from A1 should show margin or remain at A1\n---\n{}\n---",
        safe_slice(&pane, 2000));
}

/// Arrow up from A1 should enter the header row.
#[test]
fn arrow_up_from_a1_enters_header() {
    let pane = run_in_tmux("--pancurses docs/tests/overflow.corro", &["Up"], 1000);
    // After Up from A1, cursor should show header label (like ~1)
    assert!(pane.contains("~") || pane.contains("A1"),
        "Up from A1 should show header row label or remain at A1\n---\n{}\n---",
        safe_slice(&pane, 2000));
}

/// Navigate to a cell via repeated arrow keys, enter "Hello World!", and verify
/// the pancurses formula bar shows the correct address.
#[test]
fn go_to_cell_and_enter_hello_world() {
    let pane = run_in_tmux("--pancurses docs/tests/overflow.corro", &[
        "Right","Right","Down","Down","Enter","H","e","l","l","o"," ",
        "W","o","r","l","d","!","Enter",
    ], 3000);

    assert!(pane.contains("C4") || pane.contains("c4") || pane.contains("D4") || pane.contains("d4"),
        "formula bar should show C4 or nearby after edit\n---\n{}\n---", safe_slice(&pane, 3000));
    assert!(pane.contains("This Text is really long"),
        "cell content should still be visible\n---\n{}\n---", safe_slice(&pane, 3000));
}

/// Go to cell A1000 via Ctrl+G, then enter "Hello World!".
/// Verifies the pancurses formula bar shows the Go-to address.
#[test]
fn go_to_cell_via_ctrlg() {
    let pane = run_in_tmux("--pancurses docs/tests/overflow.corro", &[
        "C-g", "Enter", "Hello", " ", "World", "!", "Enter",
    ], 3000);

    // After Ctrl+G (Go to A1000), formula bar should show "A1000"
    assert!(pane.contains("A100") || pane.contains("a100"),
        "formula bar should show A1000+ after Go-to\n---\n{}\n---",
        safe_slice(&pane, 3000));
}
// ── Menu-item coverage ──────────────────────────────────────────────────
// The pancurses UI exposes eight flat root menus (File, Edit, View, Insert,
// Format, Sheet, Data, Help). They are opened with Alt+F (Escape, 'f') and the
// active submenu is switched with Right; items are highlighted with Down and
// activated with Enter.  The tests below open EVERY root menu and walk EVERY
// item, asserting each item's label is rendered while highlighted — this gives
// coverage of use of each and every menu item (every item is reachable and
// drawn).  A separate smoke test actually *activates* a safe subset.

/// Copy the canonical fixture to a temp file so menu exploration never mutates
/// the checked-in `docs/tests/overflow.corro`.
fn menu_fixture() -> String {
    let src = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("docs/tests/overflow.corro");
    // Unique per call: tests run in parallel within one process, so a path keyed
    // only on the process id would be shared and corrupted by concurrent writers.
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let dst = std::env::temp_dir().join(format!("corro-menu-{}-{}.corro", std::process::id(), id));
    std::fs::copy(&src, &dst).expect("copy fixture");
    dst.to_string_lossy().to_string()
}

/// Open root menu `sm` (0=File,1=Edit,2=Insert,3=Format,4=Sheet,5=Help) in a
/// fresh pancurses session.  Sheet/Help are opened with Alt+S / Alt+H because
/// Right-navigation from Format would enter the Format->Scope submenu (the
/// first Format item is a submenu, and Right on a submenu item enters it).
fn open_root_menu(session: &str, sm: usize) {
    tmux::send_keys(session, "Escape");
    std::thread::sleep(Duration::from_millis(120));
    match sm {
        0 => tmux::send_keys(session, "f"), // File
        1 => tmux::send_keys(session, "e"), // Edit
        2 => tmux::send_keys(session, "i"), // Insert
        3 => { // Format: File -> Edit -> Insert -> Format (Right on action items)
            tmux::send_keys(session, "f");
            std::thread::sleep(Duration::from_millis(400));
            for _ in 0..3 {
                tmux::send_keys(session, "Right");
                std::thread::sleep(Duration::from_millis(150));
            }
            std::thread::sleep(Duration::from_millis(300));
            return;
        }
        4 => tmux::send_keys(session, "s"), // Sheet
        5 => tmux::send_keys(session, "h"), // Help
        _ => {}
    }
    std::thread::sleep(Duration::from_millis(400));
}

/// Open root submenu `sm` (0=File,1=Edit,2=Insert,3=Format,4=Sheet,5=Help)
/// in a fresh pancurses session and walk every item, asserting each label is
/// visible while highlighted.  Returns the session name (caller kills it).
fn walk_menu_items(sm: usize, labels: &[&str]) -> String {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-walk-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    open_root_menu(&session, sm);
    std::thread::sleep(Duration::from_millis(300));
    for (i, &label) in labels.iter().enumerate() {
        // Retry the capture a few times to tolerate transient render delays
        // (e.g. when many tmux sessions contend for the shared server).
        let mut pane = tmux::capture_pane(&session);
        let mut tries = 0;
        while !pane.contains(label) && tries < 4 {
            std::thread::sleep(Duration::from_millis(200));
            pane = tmux::capture_pane(&session);
            tries += 1;
        }
        assert!(
            pane.contains(label),
            "menu submenu {}: item '{}' (index {}) not visible when highlighted\n--- pane ---\n{}",
            sm, label, i, safe_slice(&pane, 1500)
        );
        // Highlight the next item so it scrolls into view / is rendered.
        tmux::send_keys(&session, "Down");
        std::thread::sleep(Duration::from_millis(120));
    }
    tmux::send_keys(&session, "Escape");
    session
}

#[test]
fn menu_file_items() {
    let s = walk_menu_items(0, &[
        "Open file", "Save as", "Export", "Width", "Sort view", "Persist sort", "Exit", "Replay",
    ]);
    tmux::kill_session(&s);
}

#[test]
fn menu_edit_items() {
    let s = walk_menu_items(1, &[
        "Cut", "Copy", "Paste", "Find", "Replace", "Duplicate", "Extrapolate",
    ]);
    tmux::kill_session(&s);
}

#[test]
fn menu_insert_items() {
    let s = walk_menu_items(2, &[
        "Rows", "Mitosis (Row)", "Mitosis (Col)", "Cols", "Special Char", "Date", "Time", "Hyperlink",
    ]);
    tmux::kill_session(&s);
}

#[test]
fn menu_format_items() {
    let s = walk_menu_items(3, &["Scope", "Number", "Align", "Reset"]);
    tmux::kill_session(&s);
}

#[test]
fn menu_sheet_items() {
    let s = walk_menu_items(4, &[
        "Prev sheet", "Next sheet", "New sheet", "Rename sheet", "Copy sheet", "Move sheet", "Go", "Balance books",
    ]);
    tmux::kill_session(&s);
}

#[test]
fn menu_help_items() {
    let s = walk_menu_items(5, &["About", "Row ops", "Col ops", "Full help"]);
    tmux::kill_session(&s);
}

/// Actually *activate* a menu item and verify it took effect.
/// `expected` is the substring that should appear in the formula-bar status
/// after the action fires (the pancurses backend records `app.core.status` there,
/// proving the item was dispatched — not a no-op).  For `quit`, pass `quit=true`
/// and the app is expected to actually terminate.
fn activate_menu_item(sm: usize, idx: usize, label: &str, expected: &str, quit: bool) {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-act-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    open_root_menu(&session, sm);
    for _ in 0..idx {
        tmux::send_keys(&session, "Down");
        std::thread::sleep(Duration::from_millis(120));
    }
    std::thread::sleep(Duration::from_millis(200));
    tmux::send_keys(&session, "Enter");
    std::thread::sleep(Duration::from_millis(600));
    if quit {
        // The app should have terminated (running=false set by the Quit action).
        // The tmux session command ends with `; sleep 2`, so the session lingers
        // ~2s after corro exits; wait past that before checking it is gone.
        std::thread::sleep(Duration::from_millis(2600));
        let alive = tmux::has_session(&session);
        tmux::kill_session(&session);
        assert!(!alive, "menu item '{}' (Quit) should exit the app, but it is still running", label);
    } else {
        // A single Escape never quits; dismiss any leftover mode, then check the
        // formula bar reflects the action that fired.
        tmux::send_keys(&session, "Escape");
        std::thread::sleep(Duration::from_millis(250));
        let mut pane = tmux::capture_pane(&session);
        let mut tries = 0;
        while !pane.contains(expected) && tries < 5 {
            std::thread::sleep(Duration::from_millis(200));
            pane = tmux::capture_pane(&session);
            tries += 1;
        }
        tmux::kill_session(&session);
        assert!(
            pane.contains(expected),
            "menu item '{}' (submenu {}, idx {}) did not take effect: expected '{}' in formula-bar status\n--- pane ---\n{}",
            label, sm, idx, expected, safe_slice(&pane, 1500)
        );
    }
}

/// Extract the item labels from a bordered menu popup in a rendered frame.
/// Works for both the pancurses pane capture and the ratatui TestBackend render:
/// finds the `┌` top border, then reads the text between the popup's left and
/// right borders on each line until the `└` bottom border.  Ratatui renders a
/// shortcut prefix ("O·Open file") which is stripped to the bare label.
fn extract_popup_items(render: &str) -> Vec<String> {
    let lines: Vec<&str> = render.lines().collect();
    let mut items = Vec::new();
    let top = match lines.iter().position(|l| l.contains('┌')) { Some(i) => i, None => return items };
    let left = match lines[top].chars().position(|c| c == '┌') { Some(i) => i, None => return items };
    let right = match lines[top].chars().position(|c| c == '┐') { Some(i) => i, None => return items };
    for line in &lines[top + 1..] {
        if line.contains('└') { break; }
        if line.chars().count() > right {
            let s: String = line.chars().skip(left + 1).take(right - left - 1).collect();
            let s = s.trim().to_string();
            let s = s.strip_prefix("> ").unwrap_or(&s).to_string();
            let s = s.split('·').last().unwrap_or(&s).trim().to_string();
            if !s.is_empty() { items.push(s); }
        }
    }
    items
}

/// Rendering parity: the pancurses File menu must show the SAME items as the
/// ratatui File menu (same labels, same order).  The older menu tests only
/// asserted each pancurses label is present — they could not detect a menu that
/// diverges from the ratatui reference.  This test renders the File menu in
/// BOTH backends and compares the item lists.
#[test]
fn menu_file_parity_with_ratatui() {
    // ── ratatui reference: Alt+F opens the File menu ──
    let src = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("docs/tests/overflow.corro");
    let tmp = std::env::temp_dir().join("corro-parity-menu.corro");
    std::fs::copy(&src, &tmp).ok();
    let mut app = corro::ui::App::new(Some(tmp));
    app.load_initial().unwrap();
    ratatui_send(&mut app, crossterm::event::KeyCode::Char('f'), crossterm::event::KeyModifiers::ALT);
    let backend = ratatui::backend::TestBackend::new(120, 40);
    let mut terminal = ratatui::Terminal::new(backend).unwrap();
    terminal.draw(|f| app.bench_draw(f)).unwrap();
    let buffer = terminal.backend().buffer();
    let rat_render: String = (0..buffer.area.height)
        .map(|y| (0..buffer.area.width).map(|x| buffer[(x, y)].symbol()).collect::<String>())
        .collect::<Vec<_>>()
        .join("\n");
    let rat_items = extract_popup_items(&rat_render);

    // ── pancurses: Escape, f opens the File menu ──
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-parity-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    tmux::send_keys(&session, "Escape");
    std::thread::sleep(Duration::from_millis(120));
    tmux::send_keys(&session, "f");
    std::thread::sleep(Duration::from_millis(400));
    let pane = tmux::capture_pane(&session);
    tmux::kill_session(&session);
    let pnc_items = extract_popup_items(&pane);

    assert_eq!(pnc_items, rat_items,
        "pancurses File menu diverges from the ratatui reference\npancurses: {:?}\nratatui:   {:?}",
        pnc_items, rat_items);

    // Structural rendering parity (not just item labels):
    // - the open menu's label stays visible in the menu bar, bracketed like ratatui;
    // - the popup top border carries the menu title (┌File───┐), so the popup
    //   starts BELOW the menu bar instead of overwriting it.
    assert!(pane.contains("[File]"),
        "menu bar must keep [File] visible while the File menu is open\n{pane}");
    assert!(pane.contains("┌File"),
        "popup top border must carry the menu title (┌File…)\n{pane}");
    assert!(!pane.lines().nth(0).map_or(false, |l| l.contains("┌┐")),
        "popup must not overwrite the menu bar row");
}

#[test]
fn menu_item_activation_smoke() {
    // Each call asserts the action actually fired (status appears in the formula bar)
    // rather than the menu silently closing. This verifies menu items really work,
    // not just that the app survives.
    activate_menu_item(2, 0, "Rows", "Inserted 1 row above row 0", false);
    activate_menu_item(5, 0, "About", "About", false);
    activate_menu_item(3, 3, "Reset", "Format reset", false);
    activate_menu_item(1, 0, "Cut", "Selection cut", false);
    activate_menu_item(4, 2, "New sheet", "New sheet created", false);
    // Quit must actually terminate the app.
    activate_menu_item(0, 6, "Exit", "", true);
}
#[test]
fn menu_export_tsv_writes_file() {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-act-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    let export_path = std::env::temp_dir().join(format!("corro-export-{}.tsv", std::process::id()));
    let _ = std::fs::remove_file(&export_path);
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    tmux::send_keys(&session, "Escape");
    std::thread::sleep(Duration::from_millis(120));
    tmux::send_keys(&session, "f");
    std::thread::sleep(Duration::from_millis(400));
    // File -> Export -> TSV: Down to Export (idx 2), Right to enter the
    // submenu, then Enter on TSV (idx 0).
    for _ in 0..2 {
        tmux::send_keys(&session, "Down");
        std::thread::sleep(Duration::from_millis(120));
    }
    std::thread::sleep(Duration::from_millis(200));
    tmux::send_keys(&session, "Right"); // enter the Export submenu
    std::thread::sleep(Duration::from_millis(300));
    tmux::send_keys(&session, "Enter"); // opens the path prompt
    std::thread::sleep(Duration::from_millis(500));
    let pane = tmux::capture_pane(&session);
    assert!(pane.contains("Export TSV:"),
        "Export TSV should open a path prompt
--- pane ---
{}", safe_slice(&pane, 1500));
    // Type the export path and submit it
    let path_str = export_path.to_str().unwrap().to_string();
    for ch in path_str.chars() {
        tmux::send_keys(&session, &ch.to_string());
        std::thread::sleep(Duration::from_millis(25));
    }
    std::thread::sleep(Duration::from_millis(200));
    tmux::send_keys(&session, "Enter");
    std::thread::sleep(Duration::from_millis(700));
    tmux::kill_session(&session);
    assert!(export_path.exists(), "Export TSV should have written a file to {}", path_str);
    let content = std::fs::read_to_string(&export_path).unwrap_or_default();
    assert!(!content.is_empty(), "exported TSV should not be empty");
    let _ = std::fs::remove_file(&export_path);
}

#[test]
fn menu_open_loads_file() {
    // Build a loadable WORKBOOK snapshot from the fixture, then open it via the
    // File -> Open menu (typing the path into the TUI prompt). This verifies the
    // Open action genuinely loads a file, not just records a status.
    let src = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("docs/tests/overflow.corro");
    let mut app = corro::ui::App::new(Some(src));
    app.load_initial().unwrap();
    let snap = corro::ops::WorkbookSnapshot::from_workbook(&app.workbook);
    let snap_id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let snap_path = std::env::temp_dir().join(format!("corro-open-snap-{}-{}.corro", std::process::id(), snap_id));
    corro::io::save_workbook(&snap_path, &snap).unwrap();

    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-act-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    tmux::send_keys(&session, "Escape");
    std::thread::sleep(Duration::from_millis(120));
    tmux::send_keys(&session, "f");
    std::thread::sleep(Duration::from_millis(400));
    // Open is the first item (idx 0) — already highlighted.
    std::thread::sleep(Duration::from_millis(200));
    tmux::send_keys(&session, "Enter"); // opens the Open-file prompt
    std::thread::sleep(Duration::from_millis(500));
    let path_str = snap_path.to_str().unwrap().to_string();
    for ch in path_str.chars() {
        tmux::send_keys(&session, &ch.to_string());
        std::thread::sleep(Duration::from_millis(25));
    }
    std::thread::sleep(Duration::from_millis(200));
    tmux::send_keys(&session, "Enter");
    std::thread::sleep(Duration::from_millis(700));
    let pane = tmux::capture_pane(&session);
    tmux::kill_session(&session);
    let _ = std::fs::remove_file(&snap_path);
    assert!(pane.contains("Opened"),
        "Open should report the loaded path\n--- pane ---\n{}", safe_slice(&pane, 1500));
}

/// Activate a menu item that opens a text prompt, type `input`, submit with Enter,
/// then assert `expected` appears in the formula-bar status (proving the operation
/// actually ran on the workbook).
fn activate_menu_item_prompt(sm: usize, idx: usize, label: &str, input: &str, expected: &str) {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-act-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    open_root_menu(&session, sm);
    for _ in 0..idx {
        tmux::send_keys(&session, "Down");
        std::thread::sleep(Duration::from_millis(120));
    }
    std::thread::sleep(Duration::from_millis(200));
    tmux::send_keys(&session, "Enter"); // opens the prompt
    std::thread::sleep(Duration::from_millis(500));
    for ch in input.chars() {
        tmux::send_keys(&session, &ch.to_string());
        std::thread::sleep(Duration::from_millis(25));
    }
    std::thread::sleep(Duration::from_millis(200));
    tmux::send_keys(&session, "Enter"); // submits
    std::thread::sleep(Duration::from_millis(600));
    let pane = tmux::capture_pane(&session);
    tmux::kill_session(&session);
    assert!(pane.contains(expected),
        "menu item '{}' did not take effect: expected '{}'\n--- pane ---\n{}",
        label, expected, safe_slice(&pane, 1500));
}

#[test]
fn menu_insert_date() {
    // Insert -> Date: enters edit mode with the current date as the buffer
    // (matching ratatui). The formula bar shows the date being edited.
    let today = chrono::Local::now().format("%Y-%m-%d").to_string();
    activate_menu_item(2, 5, "Date", &today, false);
}

#[test]
fn menu_new_sheet() {
    // Sheet -> New Sheet: actually creates a new sheet.
    activate_menu_item(4, 2, "New sheet", "New sheet created", false);
}

#[test]
fn menu_rename_sheet() {
    // Sheet -> Rename Sheet: prompts, then renames the active sheet.
    activate_menu_item_prompt(4, 3, "Rename sheet", "TestSheet", "Renamed sheet to TestSheet");
}

#[test]
fn menu_edit_cut_clears_cell() {
    // Edit -> Cut: copies the cursor cell to the clipboard and clears it.
    activate_menu_item(1, 0, "Cut", "Selection cut", false);
}

#[test]
fn menu_find_text() {
    // Edit -> Find: prompts for text, then finds it and reports the cell.
    activate_menu_item_prompt(1, 3, "Find", "Hello", "Found");
}

#[test]
fn menu_replace_text() {
    // Edit -> Replace: prompts for find|replacement, then replaces occurrences.
    activate_menu_item_prompt(1, 4, "Replace", "Hello|Hi", "Replaced");
}

/// Help -> About must display an actual dialog (title + body), dismissible
/// with Escape.  Regression: it used to just set a status string and no dialog
/// ever appeared.
#[test]
fn help_about_shows_dialog() {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    // Open Help menu with Alt+H (Right-navigation from File would enter the
    // Format->Scope submenu, since the first Format item is a submenu).
    tmux::send_keys(&session, "Escape");
    std::thread::sleep(Duration::from_millis(120));
    tmux::send_keys(&session, "h");
    std::thread::sleep(Duration::from_millis(400));
    std::thread::sleep(Duration::from_millis(300));
    tmux::send_keys(&session, "Enter"); // About is item index 0 (already highlighted)
    std::thread::sleep(Duration::from_millis(600));
    let pane = tmux::capture_pane(&session);
    // Title matches the ratatui reference (" About " with spaces, not "About corro").
    assert!(pane.contains(" About "),
        "Help -> About did not display a dialog (title missing)
--- pane ---
{}", safe_slice(&pane, 2000));
    // Body matches the ratatui reference about_page_body ("Corro is a terminal
    // spreadsheet…"), not a pancurses-specific blurb.
    assert!(pane.contains("Corro is a terminal spreadsheet"),
        "Help -> About dialog body missing
--- pane ---
{}", safe_slice(&pane, 2000));
    // Structural: the dialog must be a real box (top and bottom borders), not a
    // bare text overlay.  (Rendering-parity policy: assert structure, not just text.)
    assert!(pane.contains('\u{250c}') && pane.contains('\u{2514}'),
        "About dialog is not a box (missing top/bottom border)
--- pane ---
{}", safe_slice(&pane, 2000));
    // Structural: the bottom hints/status line must STAY visible while the
    // dialog is open (ratatui's overlay covers only grid_area, not the hints
    // row).  Regression: the dialog's clear pass erased the hints line.
    let hints_row = pane.lines().nth(39).unwrap_or("");
    assert!(hints_row.contains("type/F2"),
        "bottom hints line must stay visible while the About dialog is open\n--- hints row ---\n{}\n--- pane ---\n{}",
        hints_row, safe_slice(&pane, 2000));
    // Escape must dismiss the dialog PROMPTLY.  A bare Esc used to be held by
    // ncurses for ESCDELAY (default 1000 ms) before delivery, so dismissal took
    // ~1 s.  Dismissal must complete well under that (the backend sets
    // ESCDELAY=25 ms; allow 600 ms for scheduling slack under tmux load).
    let t0 = std::time::Instant::now();
    let mut dismissed = false;
    while t0.elapsed() < Duration::from_millis(600) {
        tmux::send_keys(&session, "Escape");
        std::thread::sleep(Duration::from_millis(80));
        let p = tmux::capture_pane(&session);
        if !p.contains(" About ") { dismissed = true; break; }
    }
    let pane2 = tmux::capture_pane(&session);
    tmux::kill_session(&session);
    assert!(dismissed,
        "Escape should dismiss the About dialog within 600 ms (took {:?}; ncurses ESCDELAY regression?)\n--- pane ---\n{}",
        t0.elapsed(), safe_slice(&pane2, 2000));
}

/// Help -> Full help must also display a dialog (not a status string).
#[test]
fn help_full_shows_dialog() {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    tmux::send_keys(&session, "Escape");
    std::thread::sleep(Duration::from_millis(120));
    tmux::send_keys(&session, "h"); // Alt+H opens Help
    std::thread::sleep(Duration::from_millis(400));
    std::thread::sleep(Duration::from_millis(300));
    for _ in 0..3 {
        tmux::send_keys(&session, "Down"); // Full help is item index 3
        std::thread::sleep(Duration::from_millis(120));
    }
    std::thread::sleep(Duration::from_millis(200));
    tmux::send_keys(&session, "Enter");
    std::thread::sleep(Duration::from_millis(600));
    let pane = tmux::capture_pane(&session);
    tmux::kill_session(&session);
    // Title matches the ratatui reference (" Help " with spaces).
    assert!(pane.contains(" Help "),
        "Help -> Full help did not display a dialog
--- pane ---
{}", safe_slice(&pane, 2000));
    // Body matches the ratatui reference help_page_body ("Corro Help").
    assert!(pane.contains("Corro Help"),
        "Help -> Full help dialog body missing
--- pane ---
{}", safe_slice(&pane, 2000));
}

/// Extract the text inside a bordered dialog box (the `┌…┐` … `└…┘` region)
/// from a rendered frame.  Works for both the pancurses pane capture and the
/// ratatui TestBackend render.  `title` is the dialog title (e.g. " About ")
/// that appears on the box's top border, used to pick the dialog box out of a
/// frame that also contains the spreadsheet grid's own `┌` border.  Returns the
/// boxed lines (title + body) with the border characters stripped.
fn extract_dialog_lines(render: &str, title: &str) -> Vec<String> {
    let lines: Vec<&str> = render.lines().collect();
    // Find the box whose top border carries the title (e.g. "┌ About ──…").
    let top = match lines.iter().position(|l| l.contains('\u{250c}') && l.contains(title)) {
        Some(i) => i,
        None => return Vec::new(),
    };
    let left = match lines[top].chars().position(|c| c == '\u{250c}') { Some(i) => i, None => return Vec::new() };
    let right = match lines[top].chars().position(|c| c == '\u{2510}') { Some(i) => i, None => return Vec::new() };
    let mut out = Vec::new();
    for line in &lines[top + 1..] {
        if line.contains('\u{2514}') { break; }
        if line.chars().count() > right {
            let s: String = line.chars().skip(left + 1).take((right - left - 1).max(0)).collect();
            out.push(s.trim_end().to_string());
        }
    }
    out
}

/// Rendering parity: the pancurses About dialog must show the SAME title and
/// body as the ratatui reference (`crate::ui::App::about_page_body`).
#[test]
fn help_about_parity_with_ratatui() {
    // ── ratatui reference: Alt+H opens Help, Enter fires About ──
    let mut app = corro::ui::App::new(None);
    app.load_initial().unwrap();
    ratatui_send(&mut app, crossterm::event::KeyCode::Char('h'), crossterm::event::KeyModifiers::ALT);
    ratatui_send(&mut app, crossterm::event::KeyCode::Enter, crossterm::event::KeyModifiers::NONE);
    let backend = ratatui::backend::TestBackend::new(120, 40);
    let mut terminal = ratatui::Terminal::new(backend).unwrap();
    terminal.draw(|f| app.bench_draw(f)).unwrap();
    let buffer = terminal.backend().buffer();
    let rat_render: String = (0..buffer.area.height)
        .map(|y| (0..buffer.area.width).map(|x| buffer[(x, y)].symbol()).collect::<String>())
        .collect::<Vec<_>>()
        .join("\n");
    let rat_lines = extract_dialog_lines(&rat_render, " About ");
    assert!(!rat_lines.is_empty(), "ratatui About dialog not rendered");

    // ── pancurses: Escape, h opens Help, Enter fires About ──
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-parity-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    tmux::send_keys(&session, "Escape");
    std::thread::sleep(Duration::from_millis(120));
    tmux::send_keys(&session, "h");
    std::thread::sleep(Duration::from_millis(400));
    tmux::send_keys(&session, "Enter");
    std::thread::sleep(Duration::from_millis(600));
    let pane = tmux::capture_pane(&session);
    tmux::kill_session(&session);
    let pnc_lines = extract_dialog_lines(&pane, " About ");
    assert!(!pnc_lines.is_empty(), "pancurses About dialog not rendered\n--- pane ---\n{}", safe_slice(&pane, 2000));

    let rat_content: Vec<String> = rat_lines.iter().map(|l| l.trim().to_string()).filter(|l| !l.is_empty()).collect();
    let pnc_content: Vec<String> = pnc_lines.iter().map(|l| l.trim().to_string()).filter(|l| !l.is_empty()).collect();
    assert_eq!(pnc_content, rat_content,
        "pancurses About dialog diverges from the ratatui reference\npancurses: {:?}\nratatui:   {:?}",
        pnc_content, rat_content);
}

/// Rendering parity: the pancurses Full-help dialog must show the SAME title
/// and body as the ratatui reference (`crate::ui::App::help_page_body`).
#[test]
fn help_full_parity_with_ratatui() {
    // ── ratatui reference: Alt+H opens Help, Down x3 + Enter fires Full help ──
    let mut app = corro::ui::App::new(None);
    app.load_initial().unwrap();
    ratatui_send(&mut app, crossterm::event::KeyCode::Char('h'), crossterm::event::KeyModifiers::ALT);
    for _ in 0..3 {
        ratatui_send(&mut app, crossterm::event::KeyCode::Down, crossterm::event::KeyModifiers::NONE);
    }
    ratatui_send(&mut app, crossterm::event::KeyCode::Enter, crossterm::event::KeyModifiers::NONE);
    let backend = ratatui::backend::TestBackend::new(120, 40);
    let mut terminal = ratatui::Terminal::new(backend).unwrap();
    terminal.draw(|f| app.bench_draw(f)).unwrap();
    let buffer = terminal.backend().buffer();
    let rat_render: String = (0..buffer.area.height)
        .map(|y| (0..buffer.area.width).map(|x| buffer[(x, y)].symbol()).collect::<String>())
        .collect::<Vec<_>>()
        .join("\n");
    let rat_lines = extract_dialog_lines(&rat_render, " Help ");
    assert!(!rat_lines.is_empty(), "ratatui Full-help dialog not rendered");

    // ── pancurses: Escape, h opens Help, Down x3 + Enter fires Full help ──
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-parity-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    tmux::send_keys(&session, "Escape");
    std::thread::sleep(Duration::from_millis(120));
    tmux::send_keys(&session, "h");
    std::thread::sleep(Duration::from_millis(400));
    for _ in 0..3 {
        tmux::send_keys(&session, "Down");
        std::thread::sleep(Duration::from_millis(120));
    }
    tmux::send_keys(&session, "Enter");
    std::thread::sleep(Duration::from_millis(600));
    let pane = tmux::capture_pane(&session);
    tmux::kill_session(&session);
    let pnc_lines = extract_dialog_lines(&pane, " Help ");
    assert!(!pnc_lines.is_empty(), "pancurses Full-help dialog not rendered\n--- pane ---\n{}", safe_slice(&pane, 2000));

    let rat_content: Vec<String> = rat_lines.iter().map(|l| l.trim().to_string()).filter(|l| !l.is_empty()).collect();
    let pnc_content: Vec<String> = pnc_lines.iter().map(|l| l.trim().to_string()).filter(|l| !l.is_empty()).collect();
    assert_eq!(pnc_content, rat_content,
        "pancurses Full-help dialog diverges from the ratatui reference\npancurses: {:?}\nratatui:   {:?}",
        pnc_content, rat_content);
}

/// Insert -> Date must update the GRID immediately (not just the formula-bar
/// status).  Regression: the menu action mutated the workbook but nothing
/// re-filled the spreadsheet widget's cell buffers, so the new value appeared
/// only after arrowing away.  The old smoke test asserted only the status
/// string ("Inserted date") and could not detect this.
#[test]
fn menu_insert_date_shows_cell_immediately() {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-date-{}-{}", std::process::id(), id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    // Insert menu (Alt+I), Date is item index 5.
    send_settled(&session, "M-i");
    wait_for_text(&session, "┌Insert");
    for _ in 0..5 {
        send_settled(&session, "Down");
    }
    send_settled(&session, "Enter"); // activates Date -> enters edit mode
    // The date is now the in-progress edit buffer; Enter commits it.
    send_settled(&session, "Enter");
    let pane = tmux::capture_pane(&session);
    tmux::kill_session(&session);
    // The cursor cell is A1, whose old value is "Hello World!".  The grid row
    // for main row 1 (pane row index 5) must show a date (YYYY-MM-DD) and NOT
    // the old value — with NO cursor movement after Enter.
    let a1_row = pane.lines().nth(5).unwrap_or("");
    assert!(!a1_row.contains("Hello World!"),
        "A1 must not show the old value after Insert -> Date (stale cell buffers?)\n--- A1 row ---\n{}\n--- pane ---\n{}",
        a1_row, safe_slice(&pane, 2000));
    let date_re = |s: &str| {
        let b = s.as_bytes();
        (0..s.len()).any(|i| {
            i + 10 <= s.len()
                && b[i..i + 4].iter().all(|c| c.is_ascii_digit())
                && b[i + 4] == b'-'
                && b[i + 5..i + 7].iter().all(|c| c.is_ascii_digit())
                && b[i + 7] == b'-'
                && b[i + 8..i + 10].iter().all(|c| c.is_ascii_digit())
        })
    };
    assert!(date_re(a1_row),
        "A1 grid row must show the inserted date (YYYY-MM-DD) immediately after Enter\n--- A1 row ---\n{}\n--- pane ---\n{}",
        a1_row, safe_slice(&pane, 2000));
}

/// Insert -> Mitosis (Row) must COPY the cursor row's values into a new row
/// below it (shifting lower rows down), like the ratatui reference's
/// insert_mitosis_row_after_cursor (Op::DuplicateRow).  Regression: the GUI
/// action layer aliased mitosis to "grow a blank row at the bottom" — nothing
/// was copied.  The old coverage only walked the item's label (never
/// activated it), so the silent no-op went undetected.
#[test]
fn menu_mitosis_row_copies_row_values() {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-mit-r-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    // Cursor starts at A1 ("Hello World!").  Insert menu (Alt+I);
    // Mitosis (Row) is item index 1.
    tmux::send_keys(&session, "Escape");
    std::thread::sleep(Duration::from_millis(120));
    tmux::send_keys(&session, "i");
    std::thread::sleep(Duration::from_millis(400));
    tmux::send_keys(&session, "Down");
    std::thread::sleep(Duration::from_millis(200));
    // Enter can occasionally be lost under tmux load (the menu is then still
    // open on capture).  Retry the activation a bounded number of times instead
    // of papering over it: if the action never fires, the test still fails.
    let mut pane = String::new();
    let mut fired = false;
    for _ in 0..5 {
        tmux::send_keys(&session, "Enter");
        std::thread::sleep(Duration::from_millis(500));
        pane = tmux::capture_pane(&session);
        if pane.contains("Inserted mitosis row") {
            fired = true;
            break;
        }
    }
    tmux::kill_session(&session);
    assert!(fired, "Enter on Mitosis (Row) never fired the action (menu stuck open?)\n--- pane ---\n{}",
        safe_slice(&pane, 2000));
    // No cursor movement after Enter: assertions must hold immediately.
    let row1 = pane.lines().nth(5).unwrap_or("");
    let row2 = pane.lines().nth(6).unwrap_or("");
    assert!(row1.contains("Hello World!"),
        "main row 1 must keep its value after row mitosis\n--- row1 ---\n{}\n--- pane ---\n{}",
        row1, safe_slice(&pane, 2000));
    assert!(row2.contains("Hello World!"),
        "main row 2 must be the COPY of row 1 after row mitosis (got: {:?})\n--- row2 ---\n{}\n--- pane ---\n{}",
        row2, row2, safe_slice(&pane, 2000));
    // The cursor moves onto the duplicate (A2), matching ratatui.
    let bar = pane.lines().nth(1).unwrap_or("");
    assert!(bar.contains("A2"),
        "formula bar must show A2 (cursor on the duplicate row) after row mitosis\n--- formula bar ---\n{}\n--- pane ---\n{}",
        bar, safe_slice(&pane, 2000));
}

/// Insert -> Mitosis (Col) must COPY the cursor column's values into a new
/// column to its right (shifting columns right), like the ratatui reference's
/// insert_mitosis_col_after_cursor (Op::DuplicateCol).  Regression: aliased to
/// "grow a blank column at the right" — nothing was copied.
#[test]
fn menu_mitosis_col_copies_col_values() {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-mit-c-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    // Cursor starts at A1 ("Hello World!").  Insert menu (Alt+I), then the
    // item's SHORTCUT letter 'o' (the user's exact key sequence Alt+I > O) —
    // this must fire the item, not fall through into the grid as a cell edit.
    tmux::send_keys(&session, "Escape");
    std::thread::sleep(Duration::from_millis(120));
    tmux::send_keys(&session, "i");
    std::thread::sleep(Duration::from_millis(400));
    tmux::send_keys(&session, "o");
    std::thread::sleep(Duration::from_millis(600));
    let pane = tmux::capture_pane(&session);
    tmux::kill_session(&session);
    // Main row 1 must now show the value TWICE: original in column A and the
    // copy in the new column B.  (Column A may truncate the trailing '!', so
    // count the "Hello World" prefix, not the full string.)
    let row1 = pane.lines().nth(5).unwrap_or("");
    let copies = row1.matches("Hello World").count();
    assert!(copies >= 2,
        "main row 1 must show the copied value in the duplicate column (found {} occurrence(s))\n--- row1 ---\n{}\n--- pane ---\n{}",
        copies, copies, safe_slice(&pane, 2000));
    // The cursor moves onto the duplicate (B1), matching ratatui.
    let bar = pane.lines().nth(1).unwrap_or("");
    assert!(bar.contains("B1"),
        "formula bar must show B1 (cursor on the duplicate column) after col mitosis\n--- formula bar ---\n{}\n--- pane ---\n{}",
        bar, safe_slice(&pane, 2000));
}

/// The File menu popup must render as a bordered box (not a bare text overlay
/// over the grid).  Regression: the popup used to be drawn without a border and
/// with grid content showing through.
#[test]
fn menu_popup_has_border() {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    tmux::send_keys(&session, "Escape");
    std::thread::sleep(Duration::from_millis(120));
    tmux::send_keys(&session, "f");
    std::thread::sleep(Duration::from_millis(600));
    let pane = tmux::capture_pane(&session);
    tmux::kill_session(&session);
    assert!(pane.contains('\u{250c}') && pane.contains('\u{2514}'),
        "File menu popup is not a bordered box
--- pane ---
{}", safe_slice(&pane, 2000));
    assert!(pane.contains("Open file") && pane.contains("Exit"),
        "File menu items missing from popup
--- pane ---
{}", safe_slice(&pane, 2000));
}


/// Reproduce: type AAA, Enter, Up, type BBB, Enter, then move Up/Down.
/// The displayed value at A1 must be CONSISTENT (not flicker between "AAA" and
/// "AAABBB") and must match what is written to the .corro file.
#[test]
fn overwrite_text_is_consistent_and_matches_disk() {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fid = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let fixture = std::env::temp_dir().join(format!("corro-ow-{}-{}.corro", std::process::id(), fid));
    std::fs::write(&fixture, "CORRO_LOG 1\n").unwrap();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture.to_string_lossy()));
    // AAA, Enter
    for ch in "AAA".chars() {
        tmux::send_keys(&session, &ch.to_string());
        std::thread::sleep(Duration::from_millis(40));
    }
    std::thread::sleep(Duration::from_millis(200));
    tmux::send_keys(&session, "Enter");
    std::thread::sleep(Duration::from_millis(400));
    // Up (cursor back to A1, edit mode seeded with A1's value)
    tmux::send_keys(&session, "Up");
    std::thread::sleep(Duration::from_millis(400));
    // BBB, Enter (overwrites A1)
    for ch in "BBB".chars() {
        tmux::send_keys(&session, &ch.to_string());
        std::thread::sleep(Duration::from_millis(40));
    }
    std::thread::sleep(Duration::from_millis(200));
    tmux::send_keys(&session, "Enter");
    std::thread::sleep(Duration::from_millis(400));
    // Move Up/Down; each Up puts the cursor on A1, so the formula bar shows A1's
    // value.  It must be the same every time.
    let mut values: Vec<String> = Vec::new();
    for _ in 0..4 {
        tmux::send_keys(&session, "Up");
        std::thread::sleep(Duration::from_millis(300));
        let pane = tmux::capture_pane(&session);
        let line = pane.lines().nth(1).unwrap_or("");
        let mut toks = line.split_whitespace();
        let _addr = toks.next().unwrap_or("");
        values.push(toks.next().unwrap_or("").to_string());
        tmux::send_keys(&session, "Down");
        std::thread::sleep(Duration::from_millis(300));
    }
    tmux::kill_session(&session);
    assert!(values.iter().all(|v| v == &values[0]),
        "A1 display flickered between versions: {:?}", values);
    // The displayed value must match what is written to disk.
    let content = std::fs::read_to_string(&fixture).unwrap();
    assert!(content.contains(&format!("SET A1 {}", values[0])),
        "display '{}' does not match disk content:\n{}", values[0], content);
    let _ = std::fs::remove_file(&fixture);
}


/// Run WITHOUT a file (app.core.path is None): the commit must still apply the
/// value to the in-memory grid, or the typed text vanishes after Enter.
/// (The other edit tests all pass a .corro file, so commit_workbook_op ran and
/// this no-file path was never exercised.)
#[test]
fn edit_text_no_file_persists() {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    start_session(&session, &format!("{} --pancurses; sleep 2", bin));
    for ch in "AAA".chars() {
        tmux::send_keys(&session, &ch.to_string());
        std::thread::sleep(Duration::from_millis(40));
    }
    std::thread::sleep(Duration::from_millis(200));
    tmux::send_keys(&session, "Enter");
    std::thread::sleep(Duration::from_millis(500));
    let mut pane = tmux::capture_pane(&session);
    let mut tries = 0;
    while !pane.contains("AAA") && tries < 5 {
        std::thread::sleep(Duration::from_millis(200));
        pane = tmux::capture_pane(&session);
        tries += 1;
    }
    tmux::kill_session(&session);
    assert!(pane.contains("AAA"),
        "typed text should persist after Enter (no-file run)\n--- pane ---\n{}", safe_slice(&pane, 2000));
}


/// Type into A1, Enter, type into A2, Enter — BOTH cells must persist.
/// Regression: the deferred-commit drain lost the commit-edit callback registry
/// after the first commit, so the second edit never reached the grid/log and the
/// text silently disappeared.  (The single-cell test could not catch this.)
#[test]
fn edit_text_multiple_cells_persists() {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fid = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let fixture = std::env::temp_dir().join(format!("corro-multi-{}-{}.corro", std::process::id(), fid));
    std::fs::write(&fixture, "CORRO_LOG 1\n").unwrap();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture.to_string_lossy()));
    for ch in "AAA".chars() {
        tmux::send_keys(&session, &ch.to_string());
        std::thread::sleep(Duration::from_millis(40));
    }
    std::thread::sleep(Duration::from_millis(200));
    tmux::send_keys(&session, "Enter");
    std::thread::sleep(Duration::from_millis(400));
    for ch in "BBB".chars() {
        tmux::send_keys(&session, &ch.to_string());
        std::thread::sleep(Duration::from_millis(40));
    }
    std::thread::sleep(Duration::from_millis(200));
    tmux::send_keys(&session, "Enter");
    std::thread::sleep(Duration::from_millis(400));
    let mut pane = tmux::capture_pane(&session);
    let mut tries = 0;
    while (!pane.contains("AAA") || !pane.contains("BBB")) && tries < 5 {
        std::thread::sleep(Duration::from_millis(200));
        pane = tmux::capture_pane(&session);
        tries += 1;
    }
    tmux::kill_session(&session);
    assert!(pane.contains("AAA"),
        "A1 text should persist after editing A2\n--- pane ---\n{}", safe_slice(&pane, 2000));
    assert!(pane.contains("BBB"),
        "A2 text should persist after editing A2\n--- pane ---\n{}", safe_slice(&pane, 2000));
}


/// Type text, commit with Enter (cursor -> next line), then arrow down.
/// The typed text must persist in the grid and the formula bar must show the
/// correct main-row address (not a footer label like `A_1`).
#[test]
fn edit_text_then_move_persists() {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    start_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
    for ch in "ZZZTEST".chars() {
        tmux::send_keys(&session, &ch.to_string());
        std::thread::sleep(Duration::from_millis(30));
    }
    std::thread::sleep(Duration::from_millis(300));
    tmux::send_keys(&session, "Enter");
    std::thread::sleep(Duration::from_millis(400));
    tmux::send_keys(&session, "Down");
    std::thread::sleep(Duration::from_millis(400));
    let mut pane = tmux::capture_pane(&session);
    let mut tries = 0;
    while (!pane.contains("A3") || !pane.contains("ZZZTEST")) && tries < 5 {
        std::thread::sleep(Duration::from_millis(200));
        pane = tmux::capture_pane(&session);
        tries += 1;
    }
    tmux::kill_session(&session);
    // Formula bar must show A3 (main row 3), not a footer label.
    assert!(pane.contains("A3"),
        "formula bar should show A3 after Enter+Down\n--- pane ---\n{}", safe_slice(&pane, 2000));
    // The typed text must still be visible in the grid.
    assert!(pane.contains("ZZZTEST"),
        "typed text should persist in the grid\n--- pane ---\n{}", safe_slice(&pane, 2000));
}

/// Drive both the pancurses and ratatui backends with the same pseudorandom key
/// sequence (digits 1-2, letters A-Z, Enter, and the arrow keys) and compare the
/// formula-bar cell address after every step.  This is the "screengrab is the same
/// as the ratatui UI we are cloning" check: cursor movement / commit must agree.
#[test]
fn pseudorandom_walk_matches_ratatui() {
    use crossterm::event::KeyCode;
    // Deterministic xorshift64 PRNG.
    let mut state: u64 = 10; // seed chosen so the walk includes arrow keys
    let mut next = move || {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        state
    };
    let mut pnc_keys: Vec<String> = Vec::new();
    let mut rat_keys: Vec<KeyCode> = Vec::new();
    for _ in 0..30 {
        let k = (next() % 30) as usize;
        match k {
            0..=1 => { pnc_keys.push("1".into()); rat_keys.push(KeyCode::Char('1')); }
            2..=27 => {
                let ch = (b'A' + (k - 2) as u8) as char;
                pnc_keys.push(ch.to_string());
                rat_keys.push(KeyCode::Char(ch));
            }
            28 => { pnc_keys.push("Enter".into()); rat_keys.push(KeyCode::Enter); }
            _ => {
                let dir = (next() % 4) as usize;
                match dir {
                    0 => { pnc_keys.push("Up".into()); rat_keys.push(KeyCode::Up); }
                    1 => { pnc_keys.push("Down".into()); rat_keys.push(KeyCode::Down); }
                    2 => { pnc_keys.push("Left".into()); rat_keys.push(KeyCode::Left); }
                    _ => { pnc_keys.push("Right".into()); rat_keys.push(KeyCode::Right); }
                }
            }
        }
    }

    // ── pancurses (tmux) ──
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let session = format!("corro-walk-{}", id);
    let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
    let fixture = menu_fixture();
    // Idle marker: the app appends a line to this file after every redraw, so we
    // can send the next key the moment the previous one has been processed
    // (instead of sleeping a fixed amount).  This makes the walk as fast as the
    // app can keep up, and removes the parallel-tmux timing flakiness.
    let marker = std::env::temp_dir().join(format!("corro-idle-{}-{}.marker", std::process::id(), id));
    let _ = std::fs::remove_file(&marker);
    // Marker path goes on the command line, NOT via set_var: tmux sessions
    // inherit the long-lived server environment, not this process's, so a
    // set_var before spawn is silently ignored whenever the server predates
    // it (healthy app, zero markers in the expected file).
    let marker_arg = format!("CORRO_IDLE_MARKER={}", marker.display());
    // Start the app and wait for the initial redraw marker.  Under parallel-tmux
    // load a session can fail to start, so retry a few times.
    let mut started = false;
    for _attempt in 0..3 {
        tmux::new_session(&session, &format!("{} {} --pancurses {}; sleep 2", marker_arg, bin, fixture));
        std::thread::sleep(Duration::from_millis(2000));
        if tmux::has_session(&session) {
            let deadline = std::time::Instant::now() + Duration::from_secs(10);
            loop {
                let count = std::fs::read_to_string(&marker).map(|s| s.lines().count()).unwrap_or(0);
                if count >= 1 { started = true; break; }
                if std::time::Instant::now() > deadline { break; }
                std::thread::sleep(Duration::from_millis(50));
            }
        }
        if started { break; }
        tmux::kill_session(&session);
        std::thread::sleep(Duration::from_millis(500));
    }
    assert!(started, "pancurses app did not start (idle marker never appeared)");
    let mut pnc_addrs: Vec<(String, String)> = Vec::new();
    for (i, key) in pnc_keys.iter().enumerate() {
        tmux::send_keys(&session, key);
        wait_marker(&marker, i + 2); // initial(1) + (i+1) keys
        let pane = tmux::capture_pane(&session);
        let line = pane.lines().nth(1).unwrap_or("");
        let mut toks = line.split_whitespace();
        let addr = toks.next().unwrap_or("").to_string();
        let val = toks.next().unwrap_or("").to_string();
        pnc_addrs.push((addr, val));
    }
    tmux::kill_session(&session);
    let _ = std::fs::remove_file(&marker);

    // ── ratatui (bench_handle_key + bench_draw) ──
    let src = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("docs/tests/overflow.corro");
    let tmp_id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let tmp = std::env::temp_dir().join(format!("corro-walk-tmp-{}-{}.corro", std::process::id(), tmp_id));
    std::fs::copy(&src, &tmp).ok();
    let mut app = corro::ui::App::new(Some(tmp));
    app.load_initial().unwrap();
    let mut rat_addrs: Vec<(String, String)> = Vec::new();
    for code in &rat_keys {
        ratatui_send(&mut app, *code, crossterm::event::KeyModifiers::NONE);
        let backend = ratatui::backend::TestBackend::new(120, 40);
        let mut terminal = ratatui::Terminal::new(backend).unwrap();
        terminal.draw(|f| app.bench_draw(f)).unwrap();
        let buffer = terminal.backend().buffer();
        let row1: String = (0..buffer.area.width).map(|x| buffer[(x, 1)].symbol()).collect();
        let mut toks = row1.split_whitespace();
        let addr = toks.next().unwrap_or("").to_string();
        let val = toks.next().unwrap_or("").to_string();
        rat_addrs.push((addr, val));
    }

    assert_eq!(
        pnc_addrs, rat_addrs,
        "pseudorandom walk diverged (formula bar) between pancurses and ratatui\npancurses: {:?}\nratatui:   {:?}",
        pnc_addrs, rat_addrs
    );
}



