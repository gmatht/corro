#![cfg(all(target_os = "linux", feature = "pancurses", feature = "ratatui"))]
use std::path::PathBuf;
use std::time::Duration;
static COUNTER: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

mod tmux {
    use std::process::Command;
    pub fn new_session(session: &str, command: &str) {
        Command::new("tmux").args(["new-session","-d","-s",session,"-x","120","-y","40",command]).status().ok();
    }
    pub fn send_keys(session: &str, key: &str) {
        Command::new("tmux").args(["send-keys","-t",session,key]).status().ok();
    }
    pub fn capture_pane(session: &str) -> String {
        let o = Command::new("tmux").args(["capture-pane","-t",session,"-p","-S","-200"]).output().unwrap();
        String::from_utf8_lossy(&o.stdout).to_string()
    }
    pub fn kill_session(session: &str) { Command::new("tmux").args(["kill-session","-t",session]).output().ok(); }
}

fn ratatui_send(app: &mut corro::ui::App, code: crossterm::event::KeyCode, mods: crossterm::event::KeyModifiers) {
    app.bench_handle_key(crossterm::event::KeyEvent::new(code, mods)).ok();
}

fn ratatui_row1(app: &mut corro::ui::App) -> String {
    let backend = ratatui::backend::TestBackend::new(120, 40);
    let mut terminal = ratatui::Terminal::new(backend).unwrap();
    terminal.draw(|f| app.bench_draw(f)).unwrap();
    let buffer = terminal.backend().buffer();
    (0..buffer.area.width).map(|x| buffer[(x,1)].symbol()).collect::<String>()
}

// (menu_alt, item_index, label)
const ITEMS: &[(&str, usize, &str)] = &[
    ("f", 4, "Sort view"),
    ("f", 5, "Persist sort"),
    ("f", 7, "Replay"),
    ("e", 0, "Cut"),
    ("e", 1, "Copy"),
    ("e", 2, "Paste"),
    ("e", 5, "Duplicate"),
    ("e", 6, "Extrapolate"),
    ("i", 0, "Rows"),
    ("i", 1, "Mitosis Row"),
    ("i", 2, "Mitosis Col"),
    ("i", 3, "Cols"),
    ("i", 5, "Date"),
    ("i", 6, "Time"),
    ("s", 0, "Prev sheet"),
    ("s", 1, "Next sheet"),
    ("s", 2, "New sheet"),
    ("s", 5, "Move sheet"),
    ("s", 7, "Balance books"),
    ("h", 0, "About"),
    ("h", 1, "Row ops"),
    ("h", 2, "Col ops"),
    ("h", 3, "Full help"),
];

#[test]
fn diag_parity() {
    let src = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("docs/tests/overflow.corro");
    for &(alt, idx, label) in ITEMS {
        // ratatui
        let tmp = std::env::temp_dir().join(format!("diagp-rt-{}-{}.corro", std::process::id(), idx));
        std::fs::copy(&src, &tmp).ok();
        let mut app = corro::ui::App::new(Some(tmp.clone()));
        app.load_initial().unwrap();
        ratatui_send(&mut app, crossterm::event::KeyCode::Char(alt.chars().next().unwrap()), crossterm::event::KeyModifiers::ALT);
        for _ in 0..idx { ratatui_send(&mut app, crossterm::event::KeyCode::Down, crossterm::event::KeyModifiers::NONE); }
        ratatui_send(&mut app, crossterm::event::KeyCode::Enter, crossterm::event::KeyModifiers::NONE);
        let rat_row1 = ratatui_row1(&mut app);
        let _ = std::fs::remove_file(&tmp);

        // pancurses
        let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
        let session = format!("diagp-{}", id);
        let bin = format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"));
        let fixture = std::env::temp_dir().join(format!("diagp-fix-{}.corro", id));
        std::fs::copy(&src, &fixture).ok();
        tmux::new_session(&session, &format!("{} --pancurses {}; sleep 2", bin, fixture.to_string_lossy()));
        std::thread::sleep(Duration::from_millis(1500));
        tmux::send_keys(&session, "Escape");
        std::thread::sleep(Duration::from_millis(120));
        tmux::send_keys(&session, alt);
        std::thread::sleep(Duration::from_millis(400));
        for _ in 0..idx { tmux::send_keys(&session, "Down"); std::thread::sleep(Duration::from_millis(120)); }
        tmux::send_keys(&session, "Enter");
        std::thread::sleep(Duration::from_millis(600));
        let pane = tmux::capture_pane(&session);
        tmux::kill_session(&session);
        let _ = std::fs::remove_file(&fixture);
        let pnc_row1 = pane.lines().nth(1).unwrap_or("").to_string();

        let rat_trim: String = rat_row1.trim().to_string();
        let pnc_trim: String = pnc_row1.trim().to_string();
        let same = rat_trim == pnc_trim;
        println!("[{}] {alt}/{idx} {label}: {}", if same {"SAME"} else {"DIFF"}, if same {rat_trim} else {format!("rat={:?} pnc={:?}", rat_trim, pnc_trim)});
    }
}
