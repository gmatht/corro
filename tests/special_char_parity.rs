//! Insert → Special Char: Down*n+Enter parity between the ratatui picker
//! and the shared picker state machine (GUI/pancurses backends).
//!
//! Contract: every backend offers the same 10 labelled choices in the same
//! order (`"1: ∞"` … `"0: θ"`); Down*n lands on the nth item (clamped to the
//! last); digits `1`–`9`,`0` hotkey choices 0–8,9; Enter commits the selected
//! choice; Esc cancels with no commit. The ratatui `App` implements this
//! inline; all other backends drive `gui::special_picker` state (thin
//! renderers over one machine), so arrow/digit/Enter/Esc semantics cannot
//! drift.
//!
//! These tests pin the contract headlessly (no X, no tmux, no sleeps): the
//! ratatui side is driven via `bench_handle_key` against the committed
//! `.corro` log (full gesture through second Enter), the shared side via
//! `dispatch_menu_action` + the picker machine (open/step/take — exactly
//! what backend dialogs and the pancurses key hook call). Outcomes are
//! compared as selected/committed choice strings.
//!
//! History: the shared side used to be a free-text prompt (Down a no-op,
//! Enter confirming ""), diverging for every n. The old simulation caught
//! it; this file now guards the fixed contract.
//!
//! Run with: `cargo test --features pancurses --test special_char_parity`
//! (needs `ratatui` for `ui::App` plus `gui`/`pancurses` for the shared layer).

#![cfg(all(
    target_os = "linux",
    feature = "ratatui",
    any(feature = "gui", feature = "pancurses")
))]

use corro::gui::actions::dispatch_menu_action;
use corro::gui::special_picker;
use corro::gui::App as GuiApp;
use corro::ui::App as TuiApp;
use crossterm::event::{KeyCode, KeyEvent, KeyModifiers};
use std::path::PathBuf;

static COUNTER: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);

/// Ratatui picker's choices in navigation order (mirrors the reference
/// table; `special_picker::items` must render these same rows).
const CHOICES: [&str; 10] = ["∞", "Σ", "Ω", "π", "μ", "Δ", "√", "φ", "λ", "θ"];

fn key(c: KeyCode) -> KeyEvent {
    KeyEvent::new(c, KeyModifiers::NONE)
}
fn alt(c: char) -> KeyEvent {
    KeyEvent::new(KeyCode::Char(c), KeyModifiers::ALT)
}

fn fresh_log(tag: &str) -> PathBuf {
    let id = COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
    let p = std::env::temp_dir().join(format!(
        "corro-special-parity-{tag}-{}-{id}.corro",
        std::process::id()
    ));
    std::fs::write(&p, "CORRO_LOG 1\n").unwrap();
    p
}

fn render(app: &mut TuiApp) -> String {
    use ratatui::backend::TestBackend;
    let mut terminal = ratatui::Terminal::new(TestBackend::new(120, 40)).unwrap();
    terminal.draw(|f| app.bench_draw(f)).unwrap();
    let buf = terminal.backend().buffer();
    (0..buf.area.height)
        .map(|y| {
            (0..buf.area.width)
                .map(|x| buf[(x, y)].symbol())
                .collect::<String>()
        })
        .collect::<Vec<_>>()
        .join("\n")
}

/// Committed `SET A1 ...` values in a `.corro` log (empty when nothing was
/// committed — the observable outcome of a no-op gesture).
fn committed_a1(log: &PathBuf) -> Vec<String> {
    std::fs::read_to_string(log)
        .unwrap_or_default()
        .lines()
        .filter_map(|l| l.strip_prefix("SET A1 "))
        .map(|v| v.to_string())
        .collect()
}

/// Drive the ratatui picker through Down*n, then Enter (splice into the edit
/// buffer) + Enter (commit). Returns the highlighted picker row after the
/// Downs and the values committed to A1.
fn tui_down_n_enter(downs: usize) -> (String, Vec<String>) {
    let log = fresh_log(&format!("tui{downs}"));
    let mut app = TuiApp::new(Some(log.clone()));
    app.load_initial().unwrap();
    // Insert → Special Char via the real menu hotkeys.
    app.bench_handle_key(alt('i')).ok();
    app.bench_handle_key(key(KeyCode::Char('s'))).ok();
    let opened = render(&mut app);
    assert!(
        opened.contains("Suggestions"),
        "Down*{downs}: picker must open (appearance event), got no Suggestions surface"
    );
    for _ in 0..downs {
        app.bench_handle_key(key(KeyCode::Down)).ok();
    }
    let rt = render(&mut app);
    let highlighted = rt
        .lines()
        .find(|l| l.contains('▸'))
        .unwrap_or("")
        .to_string();
    // Enter splices the choice into the edit buffer; Enter again commits it.
    app.bench_handle_key(key(KeyCode::Enter)).ok();
    let formula = render(&mut app).lines().nth(1).unwrap_or("").to_string();
    assert!(
        formula.contains(CHOICES[downs.min(9)]),
        "Down*{downs}: formula bar must show the spliced choice {:?} after picker Enter, got {formula:?}",
        CHOICES[downs.min(9)],
    );
    app.bench_handle_key(key(KeyCode::Enter)).ok();
    (highlighted, committed_a1(&log))
}

/// Drive the shared picker machine through the same gesture every backend
/// performs: dispatch opens it, Down*n steps, Enter takes the choice.
/// Returns the committed choice (what the backend splices into its edit).
fn shared_down_n_enter(downs: usize) -> Option<String> {
    let mut app = GuiApp::new_with_paths(vec![]);
    let mut scope = 0u8;
    let mut clipboard = String::new();
    match dispatch_menu_action(&mut app, "insert_special_chars", &mut scope, &mut clipboard) {
        corro::gui::actions::MenuDispatch::SpecialPicker => {}
        d => panic!(
            "Down*{downs}: dispatch must open the picker, got {}",
            match d {
                corro::gui::actions::MenuDispatch::Status(_) => "Status",
                corro::gui::actions::MenuDispatch::Prompt(..) => "Prompt",
                corro::gui::actions::MenuDispatch::SpecialPicker => "SpecialPicker",
                corro::gui::actions::MenuDispatch::Edit { .. } => "Edit",
                corro::gui::actions::MenuDispatch::About { .. } => "About",
                corro::gui::actions::MenuDispatch::HelpFull { .. } => "HelpFull",
                corro::gui::actions::MenuDispatch::HelpKeybinds { .. } => "HelpKeybinds",
                corro::gui::actions::MenuDispatch::AggregatePicker => "AggregatePicker",
            }
        ),
    }
    for _ in 0..downs {
        special_picker::step(&mut app, 1);
    }
    special_picker::take(&mut app)
}

/// Oracle: Down*n+Enter highlights the nth choice (clamped) and commits it.
/// Covers 0, interior, the last index, and beyond-the-end (clamp).
#[test]
fn ratatui_picker_down_n_enter_commits_nth_choice() {
    let ns = [0usize, 1, 2, 5, 9, 10, 25];
    let mut highlights = Vec::new();
    let mut committed = Vec::new();
    for n in ns {
        let (hl, values) = tui_down_n_enter(n);
        highlights.push((n, hl));
        committed.push((n, values));
    }
    for (n, hl) in &highlights {
        let k = (*n).min(9);
        let want = format!("{}: {}", label(k), CHOICES[k]);
        assert!(
            hl.contains(&want),
            "Down*{n}: highlighted picker row must be {want:?}, got {hl:?}"
        );
    }
    for (n, values) in &committed {
        let k = (*n).min(9);
        assert_eq!(
            values,
            &vec![CHOICES[k].to_string()],
            "Down*{n}+Enter+Enter: A1 must commit {:?}",
            CHOICES[k],
        );
    }
}

/// Picker labels mirror the reference: 1..9 then 0.
fn label(idx: usize) -> char {
    if idx < 9 {
        char::from_digit(idx as u32 + 1, 10).unwrap()
    } else {
        '0'
    }
}

/// Equivalence: the shared machine selects the same choice the ratatui
/// picker commits, for every n (including clamp). The backend dialogs and
/// the pancurses key hook all funnel through these same calls.
#[test]
fn shared_picker_down_n_enter_matches_ratatui() {
    let ns = [0usize, 1, 2, 5, 9, 10, 25];
    let mut mismatched = Vec::new();
    for n in ns {
        let (_, tui_values) = tui_down_n_enter(n);
        let shared = shared_down_n_enter(n);
        let tui_first = tui_values.first().cloned();
        if shared != tui_first {
            mismatched.push(format!("n={n}: ratatui commits {tui_first:?}, shared takes {shared:?}"));
        }
    }
    assert!(
        mismatched.is_empty(),
        "Insert→Special Char Down*n+Enter diverged:\n{}",
        mismatched.join("\n")
    );
}

/// Equivalence for digit hotkeys: picker `3` selects Ω on both sides.
#[test]
fn shared_picker_digit_hotkey_matches_ratatui() {
    // Ratatui: picker open, `3` hotkeys Ω, Enter commits it to A1.
    let tui_log = fresh_log("tuidigit");
    let mut tui = TuiApp::new(Some(tui_log.clone()));
    tui.load_initial().unwrap();
    tui.bench_handle_key(alt('i')).ok();
    tui.bench_handle_key(key(KeyCode::Char('s'))).ok();
    tui.bench_handle_key(key(KeyCode::Char('3'))).ok();
    tui.bench_handle_key(key(KeyCode::Enter)).ok();
    let tui_values = committed_a1(&tui_log);
    // Shared: digit maps to the same index; take yields the same choice.
    assert_eq!(special_picker::index_for_digit('3'), Some(2));
    let mut shared = GuiApp::new_with_paths(vec![]);
    special_picker::open(&mut shared);
    let idx = special_picker::index_for_digit('3').unwrap();
    special_picker::set(&mut shared, idx);
    let shared_choice = special_picker::take(&mut shared);
    assert_eq!(
        shared_choice,
        tui_values.first().cloned(),
        "digit hotkey `3` must select Ω on both sides"
    );
}

/// Equivalence that holds by construction: Esc after navigating cancels
/// everywhere (no commit on either side).
#[test]
fn esc_after_navigation_cancels_on_both_sides() {
    // Ratatui: picker open, Down×3, Esc → picker gone, nothing committed.
    let tui_log = fresh_log("tuiesc");
    let mut tui = TuiApp::new(Some(tui_log.clone()));
    tui.load_initial().unwrap();
    tui.bench_handle_key(alt('i')).ok();
    tui.bench_handle_key(key(KeyCode::Char('s'))).ok();
    for _ in 0..3 {
        tui.bench_handle_key(key(KeyCode::Down)).ok();
    }
    tui.bench_handle_key(key(KeyCode::Esc)).ok();
    let rt = render(&mut tui);
    assert!(
        !rt.contains("Suggestions"),
        "Esc must dismiss the picker, surface still shows it"
    );
    let tui_values = committed_a1(&tui_log);
    // Shared: close takes nothing.
    let mut shared = GuiApp::new_with_paths(vec![]);
    special_picker::open(&mut shared);
    for _ in 0..3 {
        special_picker::step(&mut shared, 1);
    }
    special_picker::close(&mut shared);
    assert_eq!(special_picker::take(&mut shared), None);
    assert!(
        tui_values.is_empty(),
        "Esc must commit nothing, got {tui_values:?}"
    );
}
