//! Parity tests: prove the shared, backend-agnostic spreadsheet model paints
//! identically (in content) on the headless recorder and the ratatui terminal
//! backend. These run with `--features ratatui,headless` (no display server).
#![cfg(all(feature = "ratatui", feature = "headless"))]

use rustxwidgets::backends::headless::RecordingDrawContext;
use rustxwidgets::backends::ratatui::{demo_model, render_demo_to_test_backend};
use rustxwidgets::spreadsheet::{paint, SpreadsheetModel};

/// Join a ratatui `Buffer` into per-row strings so we can search for text.
fn row_strings(buf: &ratatui::buffer::Buffer) -> Vec<String> {
    let w = buf.area.width as usize;
    (0..buf.area.height)
        .map(|y| {
            (0..w)
                .map(|x| buf.cell((x as u16, y as u16)).map(|c| c.symbol()).unwrap_or("").to_string())
                .collect()
        })
        .collect()
}

fn model_with(title: &str, cells: &[((u32, u32), &str)]) -> SpreadsheetModel {
    let mut m = SpreadsheetModel::new(6, 4);
    m.set_border_title(title);
    // `cells` are already real data coordinates (row >= header, col >= margin).
    for &((r, c), t) in cells {
        m.set_cell(r, c, t);
    }
    m
}

#[test]
fn headless_records_expected_text() {
    let m = model_with("PARITY-TEST", &[((1, 1), "Hello"), ((2, 2), "World")]);
    let mut dc = RecordingDrawContext::new();
    paint(&m, &mut dc, 80, 24);

    assert!(!dc.ops.is_empty(), "paint should record draw operations");
    assert!(dc.has_text("PARITY-TEST"), "border title should be painted");
    assert!(dc.has_text("Hello"), "cell A1 text should be painted");
    assert!(dc.has_text("World"), "cell B2 text should be painted");
}

#[test]
fn headless_records_clear_and_rects() {
    let m = model_with("RECTS", &[]);
    let mut dc = RecordingDrawContext::new();
    paint(&m, &mut dc, 80, 24);

    assert!(dc.ops.iter().any(|o| matches!(o, rustxwidgets::backends::headless::DrawOp::Clear(..))));
    assert!(!dc.fill_rects().is_empty(), "background regions should be filled");
}

#[test]
fn ratatui_backend_renders_same_model() {
    let buf = render_demo_to_test_backend(80, 24);
    let rows = row_strings(&buf);
    for (i, r) in rows.iter().enumerate() {
        if !r.trim().is_empty() {
            eprintln!("row {i:>2}: {r}");
        }
    }
    let ratatui_joined = rows.join("");

    assert!(ratatui_joined.contains("rustxWidgets"), "border title must appear on the ratatui canvas");
    assert!(ratatui_joined.contains("Hello"), "cell text must appear on the ratatui canvas");
    assert!(ratatui_joined.contains("World"), "cell text must appear on the ratatui canvas");
    assert!(ratatui_joined.contains("=SUM(A1:A3)"), "formula cell must appear on the ratatui canvas");
}

#[test]
fn ratatui_and_headless_agree_on_content() {
    // Both backends consume the same model; the set of drawn text must match,
    // proving the pixel-space paint is backend-agnostic in *what* it draws.
    let m = demo_model();
    let mut dc = RecordingDrawContext::new();
    paint(&m, &mut dc, 80, 24);
    let headless_texts: std::collections::HashSet<String> = dc.texts().iter().map(|s| s.to_string()).collect();

    let buf = render_demo_to_test_backend(80, 24);
    let rows = row_strings(&buf);
    let ratatui_joined = rows.join("");

    for t in &headless_texts {
        if t.len() >= 4 {
            assert!(ratatui_joined.contains(t), "headless text {t:?} missing from ratatui render");
        }
    }
}
