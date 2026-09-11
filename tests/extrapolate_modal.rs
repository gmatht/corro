//! Modal extrapolation parity test for the reference (ratatui) backend.
//!
//! Sequence: `1` `Down` `2` `Shift+Up` `Alt+E` `E` `Down` `Down`
//!   - seeds A1=1, A2=2 (committed as SET ops), Shift+Up selects A1:A2;
//!   - Extrapolate enters the interactive modal (status hint shown);
//!   - Down Down extends the selection onto A3, which previews 3;
//!   - Enter commits `FILL A3=3`.
//!
//! This is a regression guard for two things: (1) the interactive modal's seed
//! must be the *saved* selection A1:A2 (not just the single cursor cell), and
//! (2) the modal Enter handler must actually be able to commit — it previously
//! never could, because `handle_key` had already mem::replace'd the mode out
//! before `extrapolate_selection()` matched on it, so the seed lookup fell
//! through and Enter produced "Select cells with a pattern, then Extrapolate".
//!
//! Runs headlessly on the default (ratatui) features — no tmux needed.

use crossterm::event::{KeyCode, KeyEvent, KeyModifiers};

fn key(c: KeyCode) -> KeyEvent {
    KeyEvent::new(c, KeyModifiers::NONE)
}
fn shift(c: KeyCode) -> KeyEvent {
    KeyEvent::new(c, KeyModifiers::SHIFT)
}
fn alt(c: char) -> KeyEvent {
    KeyEvent::new(KeyCode::Char(c), KeyModifiers::ALT)
}

fn render(app: &mut corro::ui::App) -> String {
    use ratatui::backend::TestBackend;
    let mut terminal = ratatui::Terminal::new(TestBackend::new(120, 40)).unwrap();
    terminal.draw(|f| app.bench_draw(f)).unwrap();
    let buf = terminal.backend().buffer();
    (0..buf.area.height)
        .map(|y| (0..buf.area.width).map(|x| buf[(x, y)].symbol()).collect::<String>())
        .collect::<Vec<_>>()
        .join("\n")
}

#[test]
fn extrapolate_modal_1_down_2_shift_up_down_down_enter() {
    let tmp = std::env::temp_dir().join(format!(
        "corro-ext-modal-{}.corro",
        std::process::id()
    ));
    std::fs::write(&tmp, "CORRO_LOG 1\n").unwrap();
    let mut app = corro::ui::App::new(Some(tmp.clone()));
    app.load_initial().unwrap();

    // Seed the column and select A1:A2 via Shift+Up.
    for k in [
        key(KeyCode::Char('1')),
        key(KeyCode::Down),
        key(KeyCode::Char('2')),
        shift(KeyCode::Up),
        alt('e'),
        key(KeyCode::Char('e')),
    ] {
        app.bench_handle_key(k).ok();
    }

    // In extrapolate mode: status prompt is shown.
    assert!(
        app.status.contains("Use arrows to extend selection"),
        "expected extrapolate prompt, got {:?}",
        app.status
    );

    // Down Down extends the selection onto A3 (formula bar address + preview).
    app.bench_handle_key(key(KeyCode::Down)).ok();
    app.bench_handle_key(key(KeyCode::Down)).ok();
    let rt = render(&mut app);
    let formula = rt.lines().nth(1).unwrap_or("");
    assert!(
        formula.contains("A3") && formula.contains("3"),
        "reference formula bar should show A3 with preview 3\n{formula:?}"
    );

    // Enter must commit the extrapolation (this was dead before the fix).
    app.bench_handle_key(key(KeyCode::Enter)).ok();
    let content = std::fs::read_to_string(&tmp).unwrap_or_default();
    let fills: Vec<String> = content
        .lines()
        .filter(|l| l.starts_with("FILL "))
        .map(|l| l.to_string())
        .collect();
    assert_eq!(
        fills,
        vec!["FILL A3=3".to_string()],
        "expected exactly FILL A3=3 (seed A1:A2, not just A1) — full file:\n{content}"
    );
    assert!(
        !content.lines().any(|l| l == "FILL A3=1"),
        "seed must be A1:A2 (incrementing), not a single-cell repeat:\n{content}"
    );
    assert!(
        app.status.contains("Extrapolated selection"),
        "expected commit status, got {:?}",
        app.status
    );
}
