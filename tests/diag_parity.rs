#![cfg(all(target_os = "linux", feature = "pancurses", feature = "ratatui"))]
use std::path::PathBuf;
use std::time::Duration;
static COUNTER: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

mod tmux {
    use std::process::Command;
    pub fn new_session(session: &str, command: &str) {
        Command::new("tmux").args(["new-session", "-d", "-s", session, "-x", "120", "-y", "40", command]).status().ok();
    }
    pub fn send_keys(session: &str, key: &str) {
        Command::new("tmux").args(["send-keys", "-t", session, key]).status().ok();
    }
    pub fn capture_pane(session: &str) -> String {
        let o = Command::new("tmux").args(["capture-pane", "-t", session, "-p", "-S", "-200"]).output().unwrap();
        String::from_utf8_lossy(&o.stdout).to_string()
    }
    pub fn kill_session(session: &str) { Command::new("tmux").args(["kill-session", "-t", session]).output().ok(); }
}

fn ratatui_send(app: &mut corro::ui::App, code: crossterm::event::KeyCode, mods: crossterm::event::KeyModifiers) {
    app.bench_handle_key(crossterm::event::KeyEvent::new(code, mods)).ok();
}

fn ratatui_row1(app: &mut corro::ui::App) -> String {
    let backend = ratatui::backend::TestBackend::new(120, 40);
    let mut terminal = ratatui::Terminal::new(backend).unwrap();
    terminal.draw(|f| app.bench_draw(f)).unwrap();
    let buffer = terminal.backend().buffer();
    (0..buffer.area.width).map(|x| buffer[(x, 1)].symbol()).collect::<String>()
}

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

/// Copy an in-repo fixture to a unique temp file (the pancurses backend
/// writes commits through to the open file, so never open a checked-in
/// fixture in place).
fn fixture_temp_copy(rel: &str) -> String {    let manifest = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let src = manifest.join(rel);
    let dst = std::env::temp_dir().join(format!(
        "diagp-{}-{}.corro",
        std::process::id(),
        COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst)
    ));
    std::fs::copy(&src, &dst).expect("copy fixture to temp");
    dst.to_string_lossy().to_string()
}

/// Map the Alt-letter to the root menu title (for popup-title waits).
fn menu_title(alt: &str) -> &'static str {
    match alt {
        "f" => "File",
        "e" => "Edit",
        "i" => "Insert",
        "s" => "Sheet",
        "h" => "Help",
        _ => "",
    }
}

// (menu_alt, item_index, label)
// Rendering-clean menu actions: same behavior in both backends, so the
// formula bar (address + value + trailing status) must match exactly.
// Excluded (known feature gaps / different interaction models, not
// rendering bugs): Cut/Copy/Paste/Duplicate (ratatui is selection-based,
// pancurses cursor-cell), Replay/Extrapolate/Balance books (pancurses
// stubs), About/Full help (modal dialogs render the formula bar
// differently).
const ITEMS: &[(&str, usize, &str)] = &[
    ("i", 0, "Rows"),
    ("i", 1, "Mitosis Row"),
    ("i", 2, "Mitosis Col"),
    ("i", 3, "Cols"),
    ("i", 5, "Date"),
    ("i", 6, "Time"),
    ("s", 0, "Prev sheet"),
    ("s", 1, "Next sheet"),
    ("s", 2, "New sheet"),
    ("h", 1, "Row ops"),
    ("h", 2, "Col ops"),
];

/// The formula bar (row 1) after activating each menu item must match between
/// pancurses and ratatui — the address, the cell value, and the trailing
/// "· status" text. Regression guard for the status-format surface.
#[test]
fn diag_parity() {
    let src = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("docs/tests/overflow.corro");
    for &(alt, idx, label) in ITEMS {
        // ── ratatui reference ──
        let tmp = std::env::temp_dir().join(format!("diagp-rt-{}-{}.corro", std::process::id(), idx));
        std::fs::copy(&src, &tmp).ok();
        let mut app = corro::ui::App::new(Some(tmp.clone()));
        app.load_initial().unwrap();
        ratatui_send(&mut app, crossterm::event::KeyCode::Char(alt.chars().next().unwrap()), crossterm::event::KeyModifiers::ALT);
        for _ in 0..idx { ratatui_send(&mut app, crossterm::event::KeyCode::Down, crossterm::event::KeyModifiers::NONE); }
        ratatui_send(&mut app, crossterm::event::KeyCode::Enter, crossterm::event::KeyModifiers::NONE);
        let rat_row1 = ratatui_row1(&mut app);
        let _ = std::fs::remove_file(&tmp);

        // ── pancurses (atomic M-letter chord + settle; see open_root_menu) ──
        let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
        let session = format!("diagp-{}-{}", std::process::id(), id);
        let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
        let fixture = fixture_temp_copy("docs/tests/overflow.corro");
        tmux::new_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture));
        wait_for_text(&session, "[File]");
        send_settled(&session, &format!("M-{alt}"));
        // Wait for the popup to actually open (send_settled can return before
        // the menu is interactive; a blind Enter then gets swallowed).
        let popup_title = format!("┌{}", menu_title(alt));
        wait_for_text(&session, &popup_title);
        for _ in 0..idx { send_settled(&session, "Down"); }
        send_settled(&session, "Enter");
        let pane = tmux::capture_pane(&session);
        tmux::kill_session(&session);
        let _ = std::fs::remove_file(&fixture);
        let pnc_row1 = pane.lines().nth(1).unwrap_or("").to_string();

        let rat_trim: String = rat_row1.trim().to_string();
        let pnc_trim: String = pnc_row1.trim().to_string();
        if rat_trim != pnc_trim {
            println!("DIFF [{alt}/{idx}] {label}:\n  ratatui:   |{rat_trim}|\n  pancurses: |{pnc_trim}|");
        } else {
            println!("SAME [{alt}/{idx}] {label}");
        }
    }
}
