#![cfg(all(feature = "gui", feature = "ratatui"))]
//! Fresh-document parity: the GUI backend must start new docs from the same
//! content the TUI uses — a `CORRO_TEMPLATE` workbook when set, else the
//! built-in margin TOTAL seeds. Before the wiring, fresh GUI docs started
//! truly blank (no TOTAL row) while the TUI started seeded.

use corro::grid::{CellAddr, ColumnAddr, HEADER_ROWS, MARGIN_COLS};

// CORRO_TEMPLATE is process-global: these tests mutate it, so serialize
// them against each other (default parallel threads would race set/remove
// across tests and flake).
static ENV_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

fn footer_seed() -> CellAddr {
    CellAddr::Footer {
        row: 0,
        col: ColumnAddr::Left(MARGIN_COLS - 1),
    }
}

fn header_seed() -> CellAddr {
    CellAddr::Header {
        row: (HEADER_ROWS - 1) as u32,
        col: ColumnAddr::Right(0),
    }
}

fn lock_env() -> std::sync::MutexGuard<'static, ()> {
    ENV_LOCK.lock().unwrap_or_else(|e| e.into_inner())
}

fn without_template() -> Option<String> {
    let prev = std::env::var("CORRO_TEMPLATE").ok();
    std::env::remove_var("CORRO_TEMPLATE");
    prev
}

#[test]
fn fresh_gui_doc_carries_total_seeds() {
    let _guard = lock_env();
    let prev = without_template();
    let app = corro::gui::App::new_with_paths(vec![]);
    let sheet = app.core.workbook.active_sheet();
    assert_eq!(
        sheet.grid.get(&footer_seed()).as_deref(),
        Some("TOTAL"),
        "fresh GUI doc must seed the margin footer TOTAL"
    );
    assert_eq!(
        sheet.grid.get(&header_seed()).as_deref(),
        Some("TOTAL"),
        "fresh GUI doc must seed the margin header TOTAL"
    );
    assert!(
        app.core.status.is_empty(),
        "no template note without CORRO_TEMPLATE, got {:?}",
        app.core.status
    );
    if let Some(v) = prev {
        std::env::set_var("CORRO_TEMPLATE", v);
    }
}

#[test]
fn fresh_gui_matches_tui_seeds() {
    let _guard = lock_env();
    let prev = without_template();
    let gui = corro::gui::App::new_with_paths(vec![]);
    let tui = corro::ui::App::new(None);
    for addr in [footer_seed(), header_seed()] {
        assert_eq!(
            gui.core.workbook.active_sheet().grid.get(&addr),
            tui.workbook.active_sheet().grid.get(&addr),
            "fresh GUI and TUI docs must start with identical TOTAL seeds at {addr:?}"
        );
    }
    if let Some(v) = prev {
        std::env::set_var("CORRO_TEMPLATE", v);
    }
}

#[test]
fn fresh_gui_honors_template_workbook() {
    let _guard = lock_env();
    let dir = std::env::temp_dir().join(format!("corro_tpl_{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let tpl = dir.join("t.corro");
    std::fs::write(&tpl, "SET $1:A1 7\n").unwrap();
    let prev = std::env::var("CORRO_TEMPLATE").ok();
    std::env::set_var("CORRO_TEMPLATE", &tpl);
    let app = corro::gui::App::new_with_paths(vec![]);
    let sheet = app.core.workbook.active_sheet();
    assert_eq!(
        sheet.grid.get(&CellAddr::Main { row: 0, col: 0 }).as_deref(),
        Some("7"),
        "CORRO_TEMPLATE workbook must become the fresh GUI document"
    );
    match prev {
        Some(v) => std::env::set_var("CORRO_TEMPLATE", v),
        None => std::env::remove_var("CORRO_TEMPLATE"),
    }
    std::fs::remove_dir_all(&dir).ok();
}

#[test]
fn fresh_gui_bad_template_falls_back_to_seeds_with_note() {
    let _guard = lock_env();
    let prev = std::env::var("CORRO_TEMPLATE").ok();
    std::env::set_var(
        "CORRO_TEMPLATE",
        "/nonexistent-dir-xyz/nope.corro",
    );
    let app = corro::gui::App::new_with_paths(vec![]);
    let sheet = app.core.workbook.active_sheet();
    assert_eq!(
        sheet.grid.get(&footer_seed()).as_deref(),
        Some("TOTAL"),
        "unreadable template must fall back to seeded blank"
    );
    assert!(
        app.core.status.contains("opened blank instead"),
        "template failure must leave a status note, got {:?}",
        app.core.status
    );
    match prev {
        Some(v) => std::env::set_var("CORRO_TEMPLATE", v),
        None => std::env::remove_var("CORRO_TEMPLATE"),
    }
}
