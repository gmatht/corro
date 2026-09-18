//! Render-level proof for the lost-TOTALs regression.
//!
//! Drives the exact path the bug lived in: materialize a fresh document's
//! untitled on-disk log (which Save then renames verbatim onto the target),
//! save, reopen, and require the right-margin TOTAL row to still render.
//! The row-number gutter is stripped because a saved+reopened document
//! legitimately renumbers rows, which would otherwise drown the real diff.
//!
//!   cargo run --release --features ratatui --bin save-roundtrip-render

use ratatui::backend::TestBackend;
use ratatui::Terminal;

fn screen(app: &mut corro::ui::App) -> String {
    let backend = TestBackend::new(100, 24);
    let mut term = Terminal::new(backend).unwrap();
    term.draw(|f| app.bench_draw(f)).unwrap();
    let buf = term.backend().buffer();
    (0..buf.area.height)
        .map(|y| {
            (0..buf.area.width)
                .map(|x| buf[(x, y)].symbol())
                .collect::<String>()
        })
        .collect::<Vec<_>>()
        .join("\n")
}

/// Right-margin column only (everything right of the main block), with the
/// row-number gutter removed and the variable-width main block trimmed.
fn right_margin_rows(s: &str) -> Vec<String> {
    let mut out: Vec<String> = Vec::new();
    for line in s.lines().skip(3) {
        if line.trim_start().starts_with("type/F2") {
            break;
        }
        // The right margin starts after the last `│`-delimited main column.
        if let Some(pos) = line.rfind("│ ") {
            let tail = line[pos..].trim_end().to_string();
            if !tail.is_empty() {
                out.push(tail);
            }
        }
    }
    out
}

fn main() {
    let dir = std::env::temp_dir().join("corro-save-roundtrip-render");
    let _ = std::fs::create_dir_all(&dir);
    unsafe { std::env::set_var("CORRO_UNSAVED_TEST_DIR", &dir) };
    let out = dir.join("out.corro");
    let _ = std::fs::remove_file(&out);

    let mut fresh = corro::ui::App::new(None);
    fresh.load_initial().expect("fresh");
    let before = right_margin_rows(&screen(&mut fresh));
    assert!(
        before.iter().any(|r| r.contains("TOTAL")),
        "fresh right margin should render TOTAL: {before:?}"
    );

    // Materialize the untitled log: the file Save renames onto `out`.
    let unsaved = fresh.debug_ensure_unsaved_file().expect("unsaved log");
    let untitled = std::fs::read_to_string(&unsaved).expect("read untitled");
    let log_has_total = untitled.contains("TOTAL");
    println!("untitled log carries TOTAL seeds: {log_has_total}");
    fresh.save_current_to_path(&out).expect("save");

    let mut reopened = corro::ui::App::new(Some(out.clone()));
    reopened.load_initial().expect("reopen");
    let after = right_margin_rows(&screen(&mut reopened));

    if before != after {
        println!("right margin before: {before:?}");
        println!("right margin after : {after:?}");
        panic!("reopened document lost its right-margin TOTAL row");
    }
    println!("OK: right-margin TOTAL row survives save→reopen");
}
