use crate::grid::{CellAddr, ColumnAddr, GridBox, SheetCursor, CellFormat, NumberFormat, TextAlign, HEADER_ROWS, MARGIN_COLS};
use crate::ops::{Op, WorkbookOp, SheetState};
use crate::ui_core;
use std::collections::HashMap;
use rustxwidgets::backends_pancurses_adapter::*;

use unicode_width::UnicodeWidthStr;

use super::compute;
use super::render::{self, CellSink};

/// Pancurses adapter: wraps a Spreadsheet ref as a CellSink for the generic fill_cells.
struct SpreadsheetSink<'a> {
    ss: &'a Spreadsheet,
}

impl<'a> SpreadsheetSink<'a> {
    fn new(ss: &'a Spreadsheet) -> Self {
        SpreadsheetSink { ss }
    }
}

impl CellSink for SpreadsheetSink<'_> {
    fn set_cell(&mut self, row: u32, col: u32, text: &str) {
        self.ss.set_cell(row, col, text);
    }
    fn set_cell_style(&mut self, row: u32, col: u32, style: compute::CellDisplayStyle) {
        self.ss.set_cell_style(row, col, style.to_pancurses_style());
    }
    fn set_raw_cell(&mut self, row: u32, col: u32, text: &str) {
        self.ss.set_raw_cell(row, col, text);
    }
    fn set_cursor(&mut self, row: u32, col: u32) {
        self.ss.set_cursor(row, col);
    }
}

/// Populate the spreadsheet widget with cell data for the given viewport.
#[allow(clippy::too_many_arguments)]
fn fill_cells(
    spreadsheet: &Spreadsheet,
    display_rows: &[usize],
    col_ixs: &[usize],
    col_widths: &HashMap<usize, usize>,
    g: &GridBox,
    hr: usize, mr: usize, mc: usize, lm: usize,
    data_width: usize,
    display_cursor_row: usize, display_cursor_col: usize,
    row_agg_func: &[Option<crate::ops::AggFunc>],
) {
    let mut sink = SpreadsheetSink::new(spreadsheet);
    render::fill_cells(
        &mut sink, display_rows, col_ixs, col_widths, g,
        hr, mr, mc, lm, data_width,
        display_cursor_row, display_cursor_col, row_agg_func,
    );
}


/// Set a main cell value in the grid and log it to the live .corro file (if any).
fn commit_cell(app: &mut super::App, addr: CellAddr, value: String) {
    let sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
    app.core.workbook.active_sheet_mut().grid.set(&addr, value.clone());
    let op = Op::SetCell { addr, value };
    let wbo = WorkbookOp::SheetOp { sheet_id, op };
    if let Some(ref p) = app.core.path.clone() {
        let mut active_sheet = sheet_id;
        let _ = crate::io::commit_workbook_op(
            p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo,
        );
        app.core.ops_applied = app.core.ops_applied.saturating_add(1);
    }
}

/// Simple A1-style label for a main cell (column letters + 1-indexed row).
fn main_addr_label(row: u32, col: u32) -> String {
    let mut name = String::new();
    let mut c = col;
    loop {
        name.insert(0, (b'A' + (c % 26) as u8) as char);
        if c < 26 { break; }
        c = c / 26 - 1;
    }
    format!("{}{}", name, row + 1)
}

/// Apply `fmt` to a format target scope (0=Cell,1=All,2=Full column,3=Data,4=Special,5=Selection).
fn apply_format(app: &mut super::App, scope: u8, main_row: u32, main_col: u32, fmt: CellFormat) {
    let (mr, mc) = {
        let sheet = app.core.workbook.active_sheet();
        (sheet.grid.main_rows(), sheet.grid.main_cols())
    };
    let (r0, r1, c0, c1) = match scope {
        1 | 3 | 4 => (0, mr, 0, mc),
        2 => (0, mr, main_col as usize, main_col as usize + 1),
        5 => {
            let (ar, ac) = app.core.anchor
                .map(|a| (a.row.saturating_sub(HEADER_ROWS), a.col.saturating_sub(MARGIN_COLS)))
                .unwrap_or((main_row as usize, main_col as usize));
            (
                ar.min(main_row as usize),
                ar.max(main_row as usize) + 1,
                ac.min(main_col as usize),
                ac.max(main_col as usize) + 1,
            )
        }
        _ => (main_row as usize, main_row as usize + 1, main_col as usize, main_col as usize + 1),
    };
    for r in r0.min(mr)..r1.min(mr) {
        for c in c0.min(mc)..c1.min(mc) {
            app.core.workbook.active_sheet_mut().grid.set_cell_format(
                CellAddr::Main { row: r as u32, col: c as u32 },
                fmt,
            );
        }
    }
}

/// Reorder the used main rows of the active sheet by `col` (ascending or descending).
/// Returns the number of rows sorted.
fn sort_sheet(app: &mut super::App, col: usize, asc: bool) -> usize {
    let g = app.core.workbook.active_sheet().grid.clone();
    let mr = g.main_rows();
    let mc = g.main_cols();
    if col >= mc || mr < 2 { return 0; }
    let used: Vec<usize> = (0..mr)
        .filter(|&r| (0..mc).any(|c| !g.get(&CellAddr::Main { row: r as u32, col: c as u32 }).unwrap_or_default().trim().is_empty()))
        .collect();
    if used.len() < 2 { return 0; }
    let mut pairs: Vec<(usize, String)> = used.iter()
        .map(|&r| (r, g.get(&CellAddr::Main { row: r as u32, col: col as u32 }).unwrap_or_default()))
        .collect();
    pairs.sort_by(|a, b| {
        let na: Option<f64> = a.1.trim().parse().ok();
        let nb: Option<f64> = b.1.trim().parse().ok();
        let ord = match (na, nb) {
            (Some(x), Some(y)) => x.partial_cmp(&y).unwrap_or(std::cmp::Ordering::Equal),
            _ => a.1.cmp(&b.1),
        };
        if asc { ord } else { ord.reverse() }
    });
    let mut cells: Vec<((usize, usize), String)> = Vec::new();
    for &r in &used {
        for c in 0..mc {
            let v = g.get(&CellAddr::Main { row: r as u32, col: c as u32 }).unwrap_or_default();
            if !v.trim().is_empty() {
                cells.push(((r, c), v));
            }
        }
    }
    {
        let g2 = app.core.workbook.active_sheet_mut();
        for &r in &used {
            for c in 0..mc {
                g2.grid.set(&CellAddr::Main { row: r as u32, col: c as u32 }, String::new());
            }
        }
    }
    for (i, &(old, _)) in pairs.iter().enumerate() {
        let new_row = used[i];
        for ((r, c), v) in cells.iter() {
            if *r == old {
                app.core.workbook.active_sheet_mut().grid.set(
                    &CellAddr::Main { row: new_row as u32, col: *c as u32 }, v.clone(),
                );
            }
        }
    }
    used.len()
}


/// Show a modal info dialog (About / Help) in the pancurses UI.  The backend
/// draws it via SGR on top of the spreadsheet output (ncurses widgets get
/// overwritten by the spreadsheet's direct SGR writes).  Without this, Help ->
/// About just set a status string and no dialog ever appeared.
fn show_info_dialog(title: &str, text: &str) {
    rustxwidgets::backends::pancurses::show_dialog(title, text);
}

pub fn run_pancurses(app: &mut super::App) -> Result<(), Box<dyn std::error::Error>> {
    let _backend = rustxwidgets::backends::pancurses::init()
        .map_err(|e| format!("pancurses init failed: {e}"))?;

    // Test-harness idle marker: the toolkit exposes a generic after-redraw
    // callback; corro wires it to append to the CORRO_IDLE_MARKER file so a
    // test can detect when a frame is fully flushed (instead of sleeping).
    if let Ok(path) = std::env::var("CORRO_IDLE_MARKER") {
        rustxwidgets::backends::pancurses::set_after_redraw_callback(Box::new(move || {
            let _ = std::fs::OpenOptions::new().create(true).append(true).open(&path)
                .and_then(|mut f| { use std::io::Write; writeln!(f, "idle") });
        }));
    }

    // Alt+letter shortcuts matching the ratatui reference: Alt+O/T/W/A/X open
    // specific File items/submenus.  The toolkit's generic alt-key callback
    // lets the app decide; the backend itself knows nothing about corro's menus.
    rustxwidgets::backends::pancurses::set_alt_key_callback(Box::new(|ch: char| {
        match ch.to_ascii_lowercase() {
            'o' => { rustxwidgets::backends::pancurses::open_menu(0, vec![], 0); true } // File -> Open file
            't' => { rustxwidgets::backends::pancurses::open_menu(0, vec![2], 0); true } // File -> Export
            'w' => { rustxwidgets::backends::pancurses::open_menu(0, vec![3], 0); true } // File -> Width
            'a' => { rustxwidgets::backends::pancurses::open_menu(0, vec![2], 2); true } // File -> Export -> ASCII table
            'x' => { rustxwidgets::backends::pancurses::open_menu(0, vec![3], 1); true } // File -> Width -> Column width
            _ => false, // let the backend fall back to the root-menu prefix match
        }
    }));

    let win = create_window()?;
    win.set_title("corro");

    // Tighten main columns to match ratatui's fit_column_to_rendered_content
    // (called during ui::App::load_initial). Without this, columns set wider
    // than max_col_width by auto_fit_column would remain uncapped.
    app.fit_main_columns_to_max_width();

    // ── Available data width / rows (matching ratatui's draw_visual) ──
    // Use environment variables to allow test backends to control size.
    let (term_cols, term_rows) = {
        let env_cols: Option<usize> = std::env::var("CORRO_TERM_COLS").ok().and_then(|s| s.parse().ok());
        let env_rows: Option<usize> = std::env::var("CORRO_TERM_ROWS").ok().and_then(|s| s.parse().ok());
        if let (Some(c), Some(r)) = (env_cols, env_rows) {
            (c, r)
        } else {
            #[cfg(unix)]
            {
                let mut ws: libc::winsize = unsafe { std::mem::zeroed() };
                if unsafe { libc::ioctl(libc::STDOUT_FILENO, libc::TIOCGWINSZ, &mut ws) } == 0 && ws.ws_col > 0
                {
                    (ws.ws_col as usize, ws.ws_row as usize)
                } else {
                    let cols: usize = std::env::var("COLUMNS")
                        .ok()
                        .and_then(|s| s.parse().ok())
                        .unwrap_or(80);
                    (cols, 50usize)
                }
            }
            #[cfg(not(unix))]
            {
                let cols: usize = std::env::var("COLUMNS")
                    .ok()
                    .and_then(|s| s.parse().ok())
                    .unwrap_or(80);
                (cols, 50usize)
            }
        }
    };
    let data_width = term_cols
        .saturating_sub(2)
        .saturating_sub(ui_core::ROW_LABEL_CHARS)
        .max(1);
    let data_cols = data_width.checked_div(2).unwrap_or(1).max(1);

    // ── Viewport rows (matching ratatui's draw_visual) ──────────────────
    // The widget layout is: menu(1) + formula(1) + border(1) + header(1) +
    // separator(1) + data_rows + border_bottom(1) + status(1) = term_rows
    // Ratatui uses the same layout: menu(1) + formula(1) + grid(h) + hints(1)
    // where grid area height = term_rows - 3, inner_h = h - 2 (borders),
    // and data_rows = inner_h - 1 = term_rows - 6.
    let data_rows = term_rows
        .saturating_sub(6)
        .max(1);

    // ── Visible columns (matching ratatui's visible_col_indices) ──────
    let hr = HEADER_ROWS;

    let display_cursor_row = HEADER_ROWS;
    let display_cursor_col = MARGIN_COLS;
    app.core.cursor.row = display_cursor_row;
    app.core.cursor.col = display_cursor_col;
    app.core.anchor = Some(SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS });

    let cursor = SheetCursor {
        row: display_cursor_row,
        col: display_cursor_col,
    };

    // ── Visible rows (matching ratatui's visible_row_indices) ──────────
    let sheet_rec = app.core.workbook.active_sheet().clone();
    let (display_rows, _row_scroll) =
        ui_core::visible_row_indices(&sheet_rec, cursor, data_rows, 0);

    // Use the live (pre-clone) grid for column width fitting, then clone
    // so the resulting overrides are present in the snapshot.
    let (mut col_ixs, _col_scroll) =
        ui_core::visible_col_indices(&sheet_rec, cursor, data_cols, 0);
    {
        let sht = app.core.workbook.active_sheet_mut();
        let grd = &mut sht.grid;
        // Match ratatui: trim columns that don't fit (no proportional refit).
        crate::ui_core::trim_visible_cols_to_width(grd, &mut col_ixs, cursor.col, data_width);
    }

    // Re-read the sheet after width adjustments.
    let sheet_rec = app.core.workbook.active_sheet().clone();
    let g = &sheet_rec.grid;
    let mr = g.main_rows();
    let mc = g.main_cols();
    let lm = MARGIN_COLS;

    // ── Column layout with widths matching ratatui's grid.col_width() ──
    // In ratatui, header and data rows use 1-char gaps everywhere (including
    // at left-margin→main and main→right-margin boundaries).  Only the
    // separator row draws a `│` at the boundary, handled by the widget.
    let mut layout: Vec<(u32, u32, String)> = Vec::new();
    let mut col_widths: HashMap<usize, usize> = HashMap::new();
    for &c in col_ixs.iter() {
        let w = g.col_width(c).max(1);
        col_widths.insert(c, w);
        let label = crate::addr::ui_column_fragment(c, mc);
        layout.push((c as u32, w as u32, label));
    }

    // ── Precompute aggregate info for each visible row ────────────────
    let row_agg_func = compute::compute_row_agg_func(g, &display_rows, hr, mr);

    // ── Spreadsheet ────────────────────────────────────────────────────
    let total_rows = display_rows.len() as u32;
    let total_cols = layout.len() as u32;
    let spreadsheet = create_spreadsheet(total_rows, total_cols)?;

    // Row labels
    let mut row_labels: Vec<(u32, String)> = Vec::new();
    for (idx, &r) in display_rows.iter().enumerate() {
        let label = crate::addr::ui_row_label(r, mr);
        row_labels.push((idx as u32, label));
    }
    spreadsheet.set_row_labels(row_labels);

    // ── Cell data for ALL visible rows and columns ────────────────────
    fill_cells(&spreadsheet, &display_rows, &col_ixs, &col_widths,
        g, hr, mr, mc, lm, data_width,
        display_cursor_row, display_cursor_col, &row_agg_func);

    spreadsheet.set_column_layout(layout);
    spreadsheet.set_grid_config(lm as u32, mc as u32);
    spreadsheet.set_row_counts(hr as u32, mr as u32);

    // Store cursor cell raw value at the cursor's position for formula bar lookup
    {
        let cursor_main_row = display_cursor_row.saturating_sub(hr);
        let cursor_col_addr = ColumnAddr::from_global(display_cursor_col, mc);
        let cursor_addr = if display_cursor_row < hr {
            CellAddr::Header { row: display_cursor_row as u32, col: cursor_col_addr }
        } else if display_cursor_row < hr + mr {
            if display_cursor_col < lm {
                CellAddr::Left { row: cursor_main_row as u32, col: display_cursor_col }
            } else if display_cursor_col < lm + mc {
                CellAddr::Main { row: cursor_main_row as u32, col: (display_cursor_col - lm) as u32 }
            } else {
                CellAddr::Right { row: cursor_main_row as u32, col: display_cursor_col - lm - mc }
            }
        } else {
            CellAddr::Footer { row: (display_cursor_row - hr - mr) as u32, col: cursor_col_addr }
        };
        let cursor_display_ri = display_rows.iter().position(|&r| r == display_cursor_row).unwrap_or(0);
        let cursor_raw_val = g.get(&cursor_addr).unwrap_or_default();
        spreadsheet.set_raw_cell(cursor_display_ri as u32, display_cursor_col as u32, &cursor_raw_val);
        // Also re-store the formatted display text for the cursor cell, so the
        // cells map always has correctly-aligned text regardless of any prior
        // fill_cells state.  Compute the effective display and align it to the
        // column width, matching fill_cells logic.
        {
            let cw_cursor = g.col_width(display_cursor_col).max(1);
            let effective_cursor = crate::formula::cell_effective_display(g, &cursor_addr);
            let formatted_cursor = crate::ui_core::format_cell_display(g, &cursor_addr, effective_cursor);
            let fw = formatted_cursor.width();
            let align_cursor = crate::ui_core::effective_cell_align(g, &cursor_addr, &formatted_cursor);
            let cursor_display_text = if fw > cw_cursor
                && (align_cursor.is_none() || align_cursor == Some(crate::grid::TextAlign::Left))
            {
                // Text would spill into adjacent columns — keep the full text
                // (matching fill_cells spill logic) so the widget renders
                // the overflow correctly instead of truncating it.
                formatted_cursor
            } else {
                crate::ui_core::align_cell_display(formatted_cursor, cw_cursor, align_cursor)
            };
            if !cursor_display_text.trim().is_empty() {
                spreadsheet.set_cell(cursor_display_ri as u32, display_cursor_col as u32, &cursor_display_text);
                spreadsheet.set_cell_style(cursor_display_ri as u32, display_cursor_col as u32, compute::CellDisplayStyle::Cursor.to_pancurses_style());
            }
        }
        spreadsheet.set_cursor(cursor_display_ri as u32, display_cursor_col as u32);

    }

    // Tab bar (styled matching ratatui: inactive=white fg+gray bg, active=bold+black fg+yellow bg)
    if app.core.workbook.sheet_count() > 1 {
        let titles: Vec<String> = app.core.workbook.sheets.iter()
            .map(|s| s.title.clone())
            .collect();
        let active = app.core.workbook.active_sheet;
        spreadsheet.set_tab_data(&titles, active);
    }

    // Border title
    let total_ops = app.core.ops_applied;
    let border_title =
        format!("corro  {}r × {}c  ops {}", mr, mc, total_ops);
    spreadsheet.set_border_title(&border_title);

    // Menu bar text (rendered) + the MenuBar *widget* so Alt+key navigation
    // (Alt+F opens File, etc.) works — matching the GTK path which creates a
    // MenuBar via build_menu(). Without the widget, menu_bar_id is None and
    // Alt+key does nothing in the pancurses backend.
    spreadsheet.set_menu_text(" [File]   Edit    Insert    Format    Sheet    Help");
    // Build the menu bar from the shared, backend-agnostic menu model.  The
    // same Menu type is used by every rustxwidgets backend, so the menu
    // definitions in crate::gui::menu are not tied to the pancurses backend.
    let menubar_model = rustxwidgets::backends_pancurses_adapter::create_menu()?;
    for root in crate::gui::menu::menu_bar() {
        let sub = rustxwidgets::backends_pancurses_adapter::create_menu()?;
        crate::gui::menu::build_menu_model(&sub, root.submenu.as_deref().unwrap_or(&[]));
        menubar_model.append_submenu(root.label, &sub);
    }
    let _menubar = rustxwidgets::backends_pancurses_adapter::create_menubar(&menubar_model, std::ptr::null_mut())?;

    // Formula bar trailing: show app status text (matches ratatui's
    // mode_prompt_widget which appends "   ·  {status}" after the cell value).
    if !app.core.status.is_empty() {
        spreadsheet.set_formula_bar_trailing(&format!("   ·  {}", app.core.status));
    } else {
        spreadsheet.set_formula_bar_trailing("");
    }

    win.set_child(&spreadsheet);
    rustxwidgets::backends::pancurses::set_focus(spreadsheet.id());
    win.present();

    // ── Cursor move callback: grow grid extent + update viewport ─────────
    // The ratatui backend recomputes the viewport each frame so pressing arrow
    // keys scrolls the visible columns/rows and the grid extent grows as needed.
    let display_rows_for_cb = std::rc::Rc::new(std::cell::RefCell::new(display_rows.clone()));
    let display_rows_for_ce = display_rows_for_cb.clone();
    // Clones captured by the go-to (Ctrl+G) callback, taken here BEFORE the
    // cursor-move / commit-edit closures move the originals below.
    let dr_for_goto_cb = display_rows_for_cb.clone();
    let dr_for_goto_ce = display_rows_for_ce.clone();
    let mut col_ixs_cb = col_ixs.clone();
    let sheet_cb = spreadsheet.clone();
    let sid = spreadsheet.id();
    let app_ptr: *mut super::App = app;

    // Wire menu item activation: when a menu item is chosen in the pancurses UI,
    // dispatch its action here.  "quit" is handled by the backend itself
    // (sets running=false); everything else performs the corresponding op on the
    // App or records status so the user sees the action fired.  Previously the
    // pancurses Enter handler just closed the menu and did nothing.
    let menu_ss = spreadsheet.clone();
    // Persist the pending format target across the Scope/Number/Align picks,
    // mirroring the ratatui menu (scope is chosen first, then the format).
    let pending_scope: std::rc::Rc<std::cell::RefCell<u8>> = std::rc::Rc::new(std::cell::RefCell::new(0));
    // Simple session clipboard for Copy/Cut/Paste.
    let clipboard: std::rc::Rc<std::cell::RefCell<String>> = std::rc::Rc::new(std::cell::RefCell::new(String::new()));
    // Clone for the Ctrl+C handler (the menu callback below moves `clipboard`).
    let ctrlc_clip = clipboard.clone();
    rustxwidgets::backends::pancurses::set_menu_action_callback(Box::new(move |name: String| {
        let app = unsafe { &mut *app_ptr };
        let hr = HEADER_ROWS;
        let lm = MARGIN_COLS;
        let main_row = app.core.cursor.row.saturating_sub(hr) as u32;
        let main_col = app.core.cursor.col.saturating_sub(lm) as u32;
        let addr = CellAddr::Main { row: main_row, col: main_col };
        let mut status: String = String::new();
        // ── Actions that need a text prompt (path / name / search text) ──
        match name.as_str() {
            "open" | "save_as" | "export_tsv" | "export_csv" | "export_ods" | "export_ascii"
            | "export_all" | "set_col_width" | "set_max_col_width" | "go_to_cell"
            | "find" | "replace" | "rename_sheet" | "copy_sheet" | "delete_sheet"
            | "insert_special_chars" | "insert_hyperlink" => {
                let label = match name.as_str() {
                    "open" => "Open file",
                    "save_as" => "Save as",
                    "export_tsv" => "Export TSV",
                    "export_csv" => "Export CSV",
                    "export_ods" => "Export ODS",
                    "export_ascii" => "Export ASCII",
                    "export_all" => "Export all to",
                    "set_col_width" => "Column width",
                    "set_max_col_width" => "Default width",
                    "go_to_cell" => "Go to cell",
                    "find" => "Find",
                    "replace" => "Replace (find|replacement)",
                    "rename_sheet" => "Rename sheet to",
                    "copy_sheet" => "Copy sheet as",
                    "delete_sheet" => "Delete sheet named",
                    "insert_special_chars" => "Insert special char",
                    "insert_hyperlink" => "Insert hyperlink",
                    _ => "File",
                };
                rustxwidgets::backends::pancurses::set_prompt(label, &name);
                return;
            }
            "save" => {
                if let Some(ref p) = app.core.path.clone() {
                    let snap = crate::ops::WorkbookSnapshot::from_workbook(&app.core.workbook);
                    match crate::io::save_workbook(p, &snap) {
                        Ok(()) => { app.core.status = "Saved".into(); menu_ss.set_formula_bar_trailing("   ·  Saved"); }
                        Err(e) => { app.core.status = format!("Save error: {e}"); menu_ss.set_formula_bar_trailing(&format!("   ·  Save error: {e}")); }
                    }
                } else {
                    rustxwidgets::backends::pancurses::set_prompt("Save as", "save_as");
                }
                return;
            }
            _ => {}
        }
        // ── Real operations ──
        match name.as_str() {
            "insert_rows" | "insert_mitosis_row" => {
                app.core.workbook.active_sheet_mut().grid.grow_main_row_at_bottom();
                status = "Inserted row".into();
            }
            "insert_cols" | "insert_mitosis_col" => {
                app.core.workbook.active_sheet_mut().grid.grow_main_col_at_right();
                status = format!("Inserted column ({name})");
            }
            "insert_date" => {
                let d = chrono::Local::now().format("%Y-%m-%d").to_string();
                commit_cell(app, addr.clone(), d.clone());
                status = format!("Inserted date {d} at {}", main_addr_label(main_row, main_col));
            }
            "insert_time" => {
                let t = chrono::Local::now().format("%H:%M:%S").to_string();
                commit_cell(app, addr.clone(), t.clone());
                status = format!("Inserted time {t} at {}", main_addr_label(main_row, main_col));
            }
            "delete_cell" | "delete" => {
                commit_cell(app, addr.clone(), String::new());
                status = format!("Cleared {}", main_addr_label(main_row, main_col));
            }
            "select_all" => {
                let (mr, mc) = {
                    let sheet = app.core.workbook.active_sheet();
                    (sheet.grid.main_rows(), sheet.grid.main_cols())
                };
                if mr > 0 && mc > 0 {
                    app.core.anchor = Some(SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS });
                    app.core.cursor = SheetCursor {
                        row: HEADER_ROWS + mr.saturating_sub(1),
                        col: MARGIN_COLS + mc.saturating_sub(1),
                    };
                }
                status = "Selected all".into();
            }
            "new_sheet" => {
                let id = app.core.workbook.next_sheet_id;
                let title = format!("Sheet {id}");
                let idx = app.core.workbook.add_sheet(title.clone(), SheetState::new(1, 1));
                app.core.workbook.active_sheet = idx;
                app.core.view_sheet_id = id;
                app.core.cursor = SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS };
                let wbo = WorkbookOp::NewSheet { id, title: title.clone() };
                if let Some(ref p) = app.core.path.clone() {
                    let mut active_sheet = id;
                    let _ = crate::io::commit_workbook_op(p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo);
                    app.core.ops_applied = app.core.ops_applied.saturating_add(1);
                }
                status = format!("New sheet created ({title})");
            }
            "copy" => {
                let val = app.core.workbook.active_sheet().grid.get(&addr).unwrap_or_default();
                *clipboard.borrow_mut() = val;
                status = format!("Copied {}", main_addr_label(main_row, main_col));
            }
            "cut" => {
                let val = app.core.workbook.active_sheet().grid.get(&addr).unwrap_or_default();
                *clipboard.borrow_mut() = val;
                commit_cell(app, addr.clone(), String::new());
                status = format!("Cut {}", main_addr_label(main_row, main_col));
            }
            "paste" => {
                let val = clipboard.borrow().clone();
                if val.is_empty() {
                    status = "Clipboard empty (use Copy/Cut first)".into();
                } else {
                    commit_cell(app, addr.clone(), val);
                    status = format!("Pasted at {}", main_addr_label(main_row, main_col));
                }
            }
            "sort_asc" | "sort_desc" | "sort_view" => {
                let asc = name != "sort_desc";
                let n = sort_sheet(app, main_col as usize, asc);
                status = if n > 0 {
                    format!("Sorted {n} rows by column {}", main_addr_label(0, main_col))
                } else {
                    "Nothing to sort".into()
                };
            }
            "persist_sort" => status = "Persist sort: not available in the pancurses build".into(),
            "replay" => status = "Replay: not available in the pancurses build".into(),
            "extrapolate" => status = "Extrapolate: not available in the pancurses build".into(),
            "duplicate" => {
                let val = app.core.workbook.active_sheet().grid.get(&addr).unwrap_or_default();
                if val.is_empty() {
                    status = "Nothing to duplicate".into();
                } else {
                    let below = CellAddr::Main { row: main_row + 1, col: main_col };
                    commit_cell(app, below.clone(), val);
                    status = format!("Duplicated {} to {}", main_addr_label(main_row, main_col), main_addr_label(main_row + 1, main_col));
                }
            }
            "sheet_prev" | "sheet_next" => {
                let n = app.core.workbook.sheet_count();
                if n > 1 {
                    let delta = if name == "sheet_prev" { -1i32 } else { 1i32 };
                    let cur = app.core.workbook.active_sheet as i32;
                    let next = (cur + delta).rem_euclid(n as i32) as usize;
                    app.core.workbook.active_sheet = next;
                    app.core.view_sheet_id = app.core.workbook.sheet_id(next);
                    app.core.cursor = SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS };
                    status = format!("Sheet: {}", app.core.workbook.sheet_title(next));
                } else {
                    status = "Only one sheet".into();
                }
            }
            "move_sheet" => {
                let n = app.core.workbook.sheet_count();
                if n > 1 {
                    let cur = app.core.workbook.active_sheet;
                    let rec = app.core.workbook.sheets.remove(cur);
                    app.core.workbook.sheets.push(rec);
                    app.core.workbook.active_sheet = n - 1;
                    app.core.view_sheet_id = app.core.workbook.sheet_id(n - 1);
                    status = format!("Moved sheet to end ({})", app.core.workbook.sheet_title(n - 1));
                } else {
                    status = "Only one sheet".into();
                }
            }
            "help_rows" => {
                show_info_dialog(
                    "Row ops",
                    "Insert rows: Insert menu → Rows\nMitosis (Row): Insert menu → Mitosis (Row)\nCtrl+arrows move the cursor",
                );
                status = "Help: Row ops".into();
            }
            "help_cols" => {
                show_info_dialog(
                    "Col ops",
                    "Insert cols: Insert menu → Cols\nMitosis (Col): Insert menu → Mitosis (Col)\nCtrl+arrows move the cursor",
                );
                status = "Help: Col ops".into();
            }
            "help_full" => {
                show_info_dialog(
                    "Full help",
                    "F2 edit\narrows move\nEnter commit\nCtrl+G go-to\nCtrl+Q quit\nAlt+letter opens a menu",
                );
                status = "Help: Full help".into();
            }
            "format_apply_all" => { *pending_scope.borrow_mut() = 1; status = "Format scope: All".into(); }
            "format_apply_full_column" => { *pending_scope.borrow_mut() = 2; status = "Format scope: Full column".into(); }
            "format_apply_data" => { *pending_scope.borrow_mut() = 3; status = "Format scope: Data".into(); }
            "format_apply_special" => { *pending_scope.borrow_mut() = 4; status = "Format scope: Special".into(); }
            "format_apply_cell" => { *pending_scope.borrow_mut() = 0; status = "Format scope: Cell".into(); }
            "format_apply_selection" => { *pending_scope.borrow_mut() = 5; status = "Format scope: Selection".into(); }
            "format_decimal_generic" => { apply_format(app, *pending_scope.borrow(), main_row, main_col, CellFormat { number: Some(NumberFormat::DecimalGeneric), align: None }); status = "Format: Decimal (generic)".into(); }
            "format_currency" => { apply_format(app, *pending_scope.borrow(), main_row, main_col, CellFormat { number: Some(NumberFormat::Currency { decimals: 2 }), align: None }); status = "Format: Currency ($)".into(); }
            "format_rational" => { apply_format(app, *pending_scope.borrow(), main_row, main_col, CellFormat { number: Some(NumberFormat::Rational), align: None }); status = "Format: Rational".into(); }
            "format_fixed_0" => { apply_format(app, *pending_scope.borrow(), main_row, main_col, CellFormat { number: Some(NumberFormat::Fixed { decimals: 0 }), align: None }); status = "Format: Fixed 0".into(); }
            "format_fixed_1" => { apply_format(app, *pending_scope.borrow(), main_row, main_col, CellFormat { number: Some(NumberFormat::Fixed { decimals: 1 }), align: None }); status = "Format: Fixed 1".into(); }
            "format_fixed_2" => { apply_format(app, *pending_scope.borrow(), main_row, main_col, CellFormat { number: Some(NumberFormat::Fixed { decimals: 2 }), align: None }); status = "Format: Fixed 2".into(); }
            "format_fixed_custom" => { apply_format(app, *pending_scope.borrow(), main_row, main_col, CellFormat { number: Some(NumberFormat::Fixed { decimals: 2 }), align: None }); status = "Format: Fixed n".into(); }
            "format_align_left" => { apply_format(app, *pending_scope.borrow(), main_row, main_col, CellFormat { number: None, align: Some(TextAlign::Left) }); status = "Format: Align Left".into(); }
            "format_align_center" => { apply_format(app, *pending_scope.borrow(), main_row, main_col, CellFormat { number: None, align: Some(TextAlign::Center) }); status = "Format: Align Center".into(); }
            "format_align_right" => { apply_format(app, *pending_scope.borrow(), main_row, main_col, CellFormat { number: None, align: Some(TextAlign::Right) }); status = "Format: Align Right".into(); }
            "format_align_default" => { apply_format(app, *pending_scope.borrow(), main_row, main_col, CellFormat { number: None, align: Some(TextAlign::Default) }); status = "Format: Align Default".into(); }
            "format_reset" => { apply_format(app, *pending_scope.borrow(), main_row, main_col, CellFormat { number: None, align: None }); status = "Format reset".into(); }
            "about" => {
                show_info_dialog(
                    "About corro",
                    "corro v0.6.0 — spreadsheet TUI (pancurses backend)",
                );
                status = "About".into();
            }
            "help_keybinds" => {
                show_info_dialog(
                    "Keybindings",
                    "F2 edit\narrows move\nEnter commit\nCtrl+G go-to\nCtrl+Q quit",
                );
                status = "Help".into();
            }
            "toggle_headers" | "toggle_margins" => status = format!("Toggle: {name} (fixed chrome in this build)"),
            "balance_books" => status = "Balance books: create a report from a data sheet — not available in the pancurses build".into(),
            "undo" | "redo" => status = format!("{name}: no undo/redo history in the pancurses build yet"),
            _ => status = format!("Menu action: {name}"),
        }
        if !status.is_empty() {
            app.core.status = status.clone();
            menu_ss.set_formula_bar_trailing(&format!("   ·  {}", status));
        }
    }));

    // Ctrl+C copies the cursor cell (standard terminal copy; it does NOT quit).
    // The value goes to the in-app clipboard (for Paste) and, via OSC 52, to the
    // terminal's system clipboard so it can be pasted elsewhere.
    let ctrlc_ss = spreadsheet.clone();
    rustxwidgets::backends::pancurses::add_key_callback('\x03', Box::new(move || {
        let app = unsafe { &mut *app_ptr };
        let hr = HEADER_ROWS;
        let lm = MARGIN_COLS;
        let main_row = app.core.cursor.row.saturating_sub(hr) as u32;
        let main_col = app.core.cursor.col.saturating_sub(lm) as u32;
        let addr = CellAddr::Main { row: main_row, col: main_col };
        let val = app.core.workbook.active_sheet().grid.get(&addr).unwrap_or_default();
        if !val.is_empty() {
            *ctrlc_clip.borrow_mut() = val.clone();
            rustxwidgets::backends::pancurses::set_clipboard_text(&val);
            app.core.status = format!("Copied {} ({})", main_addr_label(main_row, main_col), val);
            ctrlc_ss.set_formula_bar_trailing(&format!("   ·  Copied {}", main_addr_label(main_row, main_col)));
        } else {
            app.core.status = format!("Nothing to copy at {}", main_addr_label(main_row, main_col));
        }
    }));

    // Prompt callback: perform the real file operation for path/name actions
    // (Open/Save As/Export) submitted via the TUI text prompt.
    let prompt_ss = spreadsheet.clone();
    rustxwidgets::backends::pancurses::set_prompt_callback(Box::new(move |action: String, text: String| {
        let app = unsafe { &mut *app_ptr };
        let path = text.trim().to_string();
        match action.as_str() {
            "open" => {
                if !path.is_empty() {
                    match crate::io::load_workbook_snapshot(std::path::Path::new(&path)) {
                        Ok(snap) => {
                            app.core.workbook = crate::ops::WorkbookState::from_snapshot(&snap);
                            app.core.offset = 0;
                            app.core.ops_applied = 0;
                            app.core.path = Some(std::path::PathBuf::from(path.clone()));
                            app.core.status = format!("Opened {}", path);
                        }
                        Err(e) => app.core.status = format!("Open error: {e}"),
                    }
                }
            }
            "save_as" => {
                if !path.is_empty() {
                    let p = std::path::Path::new(&path);
                    let snap = crate::ops::WorkbookSnapshot::from_workbook(&app.core.workbook);
                    match crate::io::save_workbook(p, &snap) {
                        Ok(()) => {
                            app.core.path = Some(std::path::PathBuf::from(path.clone()));
                            app.core.status = format!("Saved to {}", path);
                        }
                        Err(e) => app.core.status = format!("Save error: {e}"),
                    }
                }
            }
            "export_tsv" | "export_csv" | "export_ods" | "export_ascii" | "export_all" => {
                if !path.is_empty() {
                    let g = app.core.workbook.active_sheet().grid.clone();
                    match std::fs::File::create(&path) {
                        Ok(mut f) => {
                            let r: Result<(), String> = match action.as_str() {
                                "export_tsv" => { crate::export::export_tsv(&g, &mut f); Ok(()) }
                                "export_csv" => { crate::export::export_csv(&g, &mut f); Ok(()) }
                                "export_ascii" => { crate::export::export_ascii_table(&g, &mut f, true); Ok(()) }
                                "export_ods" => match crate::ods::export_ods_bytes(&g) {
                                    Ok(bytes) => std::io::Write::write_all(&mut f, &bytes).map_err(|e| e.to_string()),
                                    Err(e) => Err(e.to_string()),
                                },
                                "export_all" => { crate::export::export_tsv(&g, &mut f); Ok(()) }
                                _ => Ok(()),
                            };
                            match r {
                                Ok(()) => app.core.status = format!("Exported {} to {}", action, path),
                                Err(e) => app.core.status = format!("Export error: {e}"),
                            }
                        }
                        Err(e) => app.core.status = format!("Export error: {e}"),
                    }
                }
            }
            "set_col_width" | "set_max_col_width" => {
                if let Ok(w) = path.trim().parse::<u32>() {
                    app.core.status = format!("Column width set to {w} (applied on next render)");
                } else {
                    app.core.status = "Column width: enter a number".into();
                }
            }
            "go_to_cell" => {
                if !path.is_empty() {
                    if let Some((addr, _, _)) = crate::addr::parse_cell_ref_at(&path, 0) {
                        match addr {
                            crate::grid::CellAddr::Main { row, col } => {
                                app.core.cursor.row = HEADER_ROWS + row as usize;
                                app.core.cursor.col = MARGIN_COLS + col as usize;
                                app.core.status = format!("Go to {}", path);
                            }
                            _ => app.core.status = format!("Unknown cell '{}'", path),
                        }
                    } else {
                        app.core.status = format!("Unknown cell '{}'", path);
                    }
                }
            }
            "find" => {
                if !path.is_empty() {
                    let g = app.core.workbook.active_sheet().grid.clone();
                    let mut found = None;
                    'find_loop: for r in 0..g.main_rows() {
                        for c in 0..g.main_cols() {
                            let v = g.get(&CellAddr::Main { row: r as u32, col: c as u32 }).unwrap_or_default();
                            if !v.trim().is_empty() && v.contains(&path) {
                                found = Some((r, c));
                                break 'find_loop;
                            }
                        }
                    }
                    if let Some((r, c)) = found {
                        app.core.cursor.row = HEADER_ROWS + r;
                        app.core.cursor.col = MARGIN_COLS + c;
                        app.core.status = format!("Found '{}' at {}", path, main_addr_label(r as u32, c as u32));
                    } else {
                        app.core.status = format!("'{}' not found", path);
                    }
                }
            }
            "replace" => {
                let (find, repl) = match path.split_once('|') {
                    Some((f, r)) => (f.to_string(), r.to_string()),
                    None => (path.clone(), String::new()),
                };
                if find.is_empty() {
                    app.core.status = "Replace: enter find|replacement".into();
                } else {
                    let g = app.core.workbook.active_sheet().grid.clone();
                    let mr = g.main_rows();
                    let mc = g.main_cols();
                    let mut count = 0usize;
                    for r in 0..mr {
                        for c in 0..mc {
                            let v = g.get(&CellAddr::Main { row: r as u32, col: c as u32 }).unwrap_or_default();
                            if v.contains(&find) {
                                let nv = v.replace(&find, &repl);
                                commit_cell(app, CellAddr::Main { row: r as u32, col: c as u32 }, nv);
                                count += 1;
                            }
                        }
                    }
                    app.core.status = format!("Replaced {count} occurrence(s)");
                }
            }
            "rename_sheet" => {
                if !path.is_empty() {
                    let id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
                    if let Some(sheet) = app.core.workbook.sheets.iter_mut().find(|s| s.id == id) {
                        sheet.title = path.clone();
                    }
                    let wbo = WorkbookOp::RenameSheet { id, title: path.clone() };
                    if let Some(ref p) = app.core.path.clone() {
                        let mut active_sheet = id;
                        let _ = crate::io::commit_workbook_op(p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo);
                        app.core.ops_applied = app.core.ops_applied.saturating_add(1);
                    }
                    app.core.status = format!("Renamed sheet to {}", path);
                }
            }
            "copy_sheet" => {
                if !path.is_empty() {
                    let source_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
                    let id = app.core.workbook.next_sheet_id;
                    if let Some(source) = app.core.workbook.sheets.iter().find(|s| s.id == source_id) {
                        let nidx = app.core.workbook.add_sheet_record(crate::ops::SheetRecord {
                            id,
                            title: path.clone(),
                            state: source.state.clone(),
                            linked_source: source.linked_source.clone(),
                        });
                        app.core.workbook.active_sheet = nidx;
                    }
                    let wbo = WorkbookOp::CopySheet { source_id, id, title: path.clone() };
                    if let Some(ref p) = app.core.path.clone() {
                        let mut active_sheet = source_id;
                        let _ = crate::io::commit_workbook_op(p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo);
                        app.core.ops_applied = app.core.ops_applied.saturating_add(1);
                    }
                    app.core.status = format!("Copied sheet as {}", path);
                }
            }
            "delete_sheet" => {
                if app.core.workbook.sheets.len() > 1 {
                    let idx = if path.is_empty() {
                        app.core.workbook.active_sheet
                    } else {
                        app.core.workbook.sheets.iter().position(|s| s.title == path).unwrap_or(app.core.workbook.active_sheet)
                    };
                    let id = app.core.workbook.sheets[idx].id;
                    let wbo = WorkbookOp::DeleteSheet { id };
                    if let Some(ref p) = app.core.path.clone() {
                        let mut active_sheet = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
                        let _ = crate::io::commit_workbook_op(p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo);
                        app.core.ops_applied = app.core.ops_applied.saturating_add(1);
                    }
                    app.core.workbook.sheets.remove(idx);
                    if app.core.workbook.active_sheet >= idx {
                        app.core.workbook.active_sheet = app.core.workbook.active_sheet.saturating_sub(1);
                    }
                    app.core.status = if path.is_empty() { "Deleted active sheet".into() } else { format!("Deleted sheet {}", path) };
                } else {
                    app.core.status = "Cannot delete the last sheet".into();
                }
            }
            "insert_special_chars" | "insert_hyperlink" => {
                if !path.is_empty() {
                    let hr = HEADER_ROWS;
                    let lm = MARGIN_COLS;
                    let addr = CellAddr::Main {
                        row: app.core.cursor.row.saturating_sub(hr) as u32,
                        col: app.core.cursor.col.saturating_sub(lm) as u32,
                    };
                    commit_cell(app, addr.clone(), path.clone());
                    app.core.status = if action == "insert_special_chars" {
                        format!("Inserted '{}'", path)
                    } else {
                        format!("Inserted hyperlink {}", path)
                    };
                }
            }
            _ => { app.core.status = format!("Menu: {}", action); }
        }
        prompt_ss.set_formula_bar_trailing(&format!("   ·  {}", app.core.status));
    }));
    let hr_cb = hr;
    let hr_ce = hr;
    let mr_cb = mr;
    let lm_ce = lm;
    let data_rows_cb = data_rows;
    let data_cols_cb = data_cols;
    let data_width_cb = data_width;
    let mut prev_cursor_col = cursor.col;
    let mut prev_cursor_row = cursor.row;
    add_cursor_move_callback(move |_display_row, _display_col| {
        // SAFETY: app is &mut App alive for the entire event loop
        let app = unsafe { &mut *app_ptr };

        // Sync formula bar trailing with app status (ratatui shows status in formula bar)
        if !app.core.status.is_empty() {
            sheet_cb.set_formula_bar_trailing(&format!("   ·  {}", app.core.status));
        } else {
            sheet_cb.set_formula_bar_trailing("");
        }

        let display_idx = _display_row as usize;
        let mut need_viewport_recompute = false;

        // Handle sentinel values for scrolling past viewport boundaries
        if _display_row == u32::MAX {
            // Scroll up sentinel: user pressed Up at the first visible row
            if app.core.cursor.row > 0 {
                app.core.cursor.row -= 1;
            }
            need_viewport_recompute = true;
        } else if _display_row == u32::MAX - 1 {
            // Scroll down sentinel: user pressed Down at the last visible row.
            // Grow the grid if the cursor is at the last main row with few
            // trailing blanks (matching ratatui's move_cursor_one_row_vertical).
            let cursor_row = app.core.cursor.row;
            {
                let sheet = app.core.workbook.active_sheet_mut();
                let mr = sheet.grid.main_rows();
                if cursor_row == hr_cb + mr.saturating_sub(1)
                    && compute::trailing_blank_main_rows(&sheet.grid) < crate::ui_core::NAV_BLANK_ROWS
                {
                    sheet.grid.grow_main_row_at_bottom();
                }
            }
            app.core.cursor.row += 1;
            need_viewport_recompute = true;
        } else if let Some(&logical_row) = display_rows_for_cb.borrow().get(display_idx) {
            app.core.cursor.row = logical_row;

            // Grow the grid if cursor moves from the last main row with few
            // trailing blanks (matching ratatui's move_cursor_one_row_vertical).
            if logical_row >= hr_cb + mr_cb {
                let sheet = app.core.workbook.active_sheet_mut();
                let cur_mr = sheet.grid.main_rows();
                if prev_cursor_row == hr_cb + cur_mr.saturating_sub(1)
                    && compute::trailing_blank_main_rows(&sheet.grid) < crate::ui_core::NAV_BLANK_ROWS
                {
                    sheet.grid.grow_main_row_at_bottom();
                }
            }

            // Check if cursor moved into header or footer region; if so,
            // recompute the viewport so those rows become visible.
            if logical_row < hr_cb || logical_row >= hr_cb + mr_cb {
                need_viewport_recompute = true;
            }
        }
        app.core.cursor.col = _display_col as usize;

        if need_viewport_recompute {
            let cursor = app.core.cursor;
            // Determine viewport and fit columns BEFORE cloning so the
            // resulting width overrides are reflected in the snapshot.
            let (new_display_rows, new_mr, new_mc, new_ixs) = {
                let rec = app.core.workbook.active_sheet().clone();
                let (new_display_rows, _) =
                    crate::ui_core::visible_row_indices(&rec, cursor, data_rows_cb, 0);
                let new_mr = rec.grid.main_rows();
                let new_mc = rec.grid.main_cols();
                let (mut new_ixs, _) =
                    crate::ui_core::visible_col_indices(&rec, cursor, data_cols_cb, 0);
                // Trim columns to fit (matching ratatui: no proportional refit).
                {
                    let sht = app.core.workbook.active_sheet_mut();
                    crate::ui_core::trim_visible_cols_to_width(
                        &mut sht.grid, &mut new_ixs, cursor.col, data_width_cb,
                    );
                }
                (new_display_rows, new_mr, new_mc, new_ixs)
            };
            // Re-read the sheet after width adjustments.
            let rec = app.core.workbook.active_sheet().clone();
            // Update border title when grid grew (matching the non-recompute path)
            let boundary_title = format!(
                "corro  {}r × {}c  ops {}",
                new_mr, new_mc, app.core.ops_applied
            );
            spreadsheet_set_border_title(sid, &boundary_title);
            let new_labels: Vec<(u32, String)> = new_display_rows.iter()
                .enumerate()
                .map(|(idx, &r)| {
                    let label = crate::addr::ui_row_label(r, new_mr);
                    (idx as u32, label)
                })
                .collect();
            spreadsheet_set_row_labels(sid, new_labels);
            // Update column layout with fitted widths
            {
                let g = &rec.grid;
                let new_layout: Vec<(u32, u32, String)> = new_ixs
                    .iter()
                    .map(|&c| {
                        let w = g.col_width(c).max(1);
                        let label = crate::addr::ui_column_fragment(c, new_mc);
                        (c as u32, w as u32, label)
                    })
                    .collect();
                spreadsheet_set_column_layout(sid, new_layout);
                col_ixs_cb = new_ixs.clone();
                spreadsheet_set_grid_config(sid, MARGIN_COLS as u32, new_mc as u32);
            }
            // Repopulate all visible cells for the new viewport
            let new_col_widths: HashMap<usize, usize> = col_ixs_cb.iter()
                .map(|&c| (c, rec.grid.col_width(c).max(1)))
                .collect();
            let new_row_agg = compute::compute_row_agg_func(&rec.grid, &new_display_rows, hr_cb, new_mr);
            fill_cells(
                &sheet_cb, &new_display_rows, &col_ixs_cb, &new_col_widths,
                &rec.grid, hr_cb, new_mr, new_mc, MARGIN_COLS, data_width_cb,
                cursor.row, cursor.col, &new_row_agg,
            );
            if let Some(new_display_ri) = new_display_rows.iter().position(|&r| r == cursor.row) {
                let cursor_addr = crate::addr::sheet_cursor_to_addr(
                    crate::addr::LogicalRow(cursor.row),
                    crate::addr::GlobalCol(cursor.col),
                    crate::addr::MainRows(new_mr),
                    crate::addr::MainCols(new_mc),
                );
                if let Some(raw_val) = rec.grid.get(&cursor_addr) {
                    sheet_cb.set_raw_cell(new_display_ri as u32, cursor.col as u32, &raw_val);
                } else {
                    sheet_cb.set_raw_cell(new_display_ri as u32, cursor.col as u32, "");
                }
                sheet_cb.set_cursor(new_display_ri as u32, cursor.col as u32);
            }
            *display_rows_for_cb.borrow_mut() = new_display_rows;
            prev_cursor_row = app.core.cursor.row;
            prev_cursor_col = app.core.cursor.col;
            return;
        }

        if let Some(&logical_row) = display_rows_for_cb.borrow().get(display_idx) {
            let sheet = app.core.workbook.active_sheet_mut();
            let prev_mr = sheet.grid.main_rows();
            let prev_mc = sheet.grid.main_cols();
            let lm = MARGIN_COLS;

            // Expand main columns when moving right past the last main column
            // (matching ratatui's move_cursor_one_col_horizontal).
            let new_col = _display_col as usize;
            if new_col > prev_cursor_col {
                if prev_cursor_col == lm + prev_mc.saturating_sub(1)
                    && compute::trailing_blank_main_cols(&sheet.grid) < crate::ui_core::NAV_BLANK_COLS
                {
                    sheet.grid.grow_main_col_at_right();
                }
            }

            // Expand main rows when moving down from the last main row
            // (matching ratatui's move_cursor_one_row_vertical).
            if logical_row > prev_cursor_row {
                if prev_cursor_row == hr_cb + prev_mr.saturating_sub(1)
                    && compute::trailing_blank_main_rows(&sheet.grid) < crate::ui_core::NAV_BLANK_ROWS
                {
                    sheet.grid.grow_main_row_at_bottom();
                }
            }

            prev_cursor_col = new_col;
            prev_cursor_row = logical_row;

            // When cursor moves to the row just beyond the current extent,
            // grow the grid (matching ratatui's move_cursor_one_row_vertical)
            // for cases where the cursor jumps multiple rows at once.
            // Use the current main row count (after any growth above) to
            // avoid redundant growth: prev_mr may be stale if the earlier
            // row-growth condition at line ~753 already fired.
            let cur_mr = sheet.grid.main_rows();
            if logical_row >= hr_cb + cur_mr {
                // Grow the grid so the cursor row becomes a main row (matching
                // ratatui's move_cursor_one_row_vertical, which extends the grid
                // when the cursor moves down past the last main row).  Growing by
                // one is not enough when the cursor jumps several rows at once.
                let need = logical_row.saturating_sub(hr_cb) + 1;
                if sheet.grid.main_rows() < need {
                    sheet.grid.set_main_size(need, sheet.grid.main_cols());
                }
            }
            sheet.grid.ensure_extent_for_cursor(logical_row, _display_col as usize);
            if sheet.grid.main_rows() != prev_mr || sheet.grid.main_cols() != prev_mc {
                // Grid grew — update border title and row labels
                let mr = sheet.grid.main_rows();
                let mc = sheet.grid.main_cols();
                let boundary_title = format!(
                    "corro  {}r × {}c  ops {}",
                    mr, mc, app.core.ops_applied
                );
                spreadsheet_set_border_title(sid, &boundary_title);
                let new_labels: Vec<(u32, String)> = display_rows_for_cb.borrow().iter()
                    .enumerate()
                    .map(|(idx, &r)| {
                        let label = crate::addr::ui_row_label(r, mr);
                        (idx as u32, label)
                    })
                    .collect();
                spreadsheet_set_row_labels(sid, new_labels);
                // Also update column layout when main columns grew
                let rec = app.core.workbook.active_sheet().clone();
                let cursor = app.core.cursor;
                let (mut new_ixs, _) =
                    crate::ui_core::visible_col_indices(&rec, cursor, data_cols_cb, 0);
                // Trim columns to fit (matching ratatui: no proportional refit).
                {
                    let sht = app.core.workbook.active_sheet_mut();
                    crate::ui_core::trim_visible_cols_to_width(
                        &mut sht.grid, &mut new_ixs, cursor.col, data_width_cb,
                    );
                }
                let rec = app.core.workbook.active_sheet().clone();
                let g = &rec.grid;
                let mc = g.main_cols();
                let new_layout: Vec<(u32, u32, String)> = new_ixs
                    .iter()
                    .map(|&c| {
                        let w = g.col_width(c).max(1);
                        let label = crate::addr::ui_column_fragment(c, mc);
                        (c as u32, w as u32, label)
                    })
                    .collect();
                spreadsheet_set_column_layout(sid, new_layout);
                col_ixs_cb = new_ixs;
                // Keep the widget's margin_cols and main_cols in sync
                spreadsheet_set_grid_config(sid, lm as u32, mc as u32);
                // Repopulate cells with updated column layout after growth
                let dr: Vec<usize> = display_rows_for_cb.borrow().clone();
                let new_col_widths: HashMap<usize, usize> = col_ixs_cb.iter()
                    .map(|&c| (c, g.col_width(c).max(1)))
                    .collect();
                let new_row_agg = compute::compute_row_agg_func(g, &dr, hr_cb, mr);
                fill_cells(
                    &sheet_cb, &dr, &col_ixs_cb, &new_col_widths,
                    g, hr_cb, mr, mc, MARGIN_COLS, data_width_cb,
                    cursor.row, cursor.col, &new_row_agg,
                );
            } else if !col_ixs_cb.contains(&(_display_col as usize)) {
                // Update column viewport when cursor column moves outside the
                // currently visible range (matching ratatui's per-frame recompute).
                let rec = app.core.workbook.active_sheet().clone();
                let cursor = app.core.cursor;
                let (mut new_ixs, _) =
                    crate::ui_core::visible_col_indices(&rec, cursor, data_cols_cb, 0);
                // Trim columns to fit (matching ratatui: no proportional refit).
                {
                    let sht = app.core.workbook.active_sheet_mut();
                    crate::ui_core::trim_visible_cols_to_width(
                        &mut sht.grid, &mut new_ixs, cursor.col, data_width_cb,
                    );
                }
                let rec = app.core.workbook.active_sheet().clone();
                let g = &rec.grid;
                let mc = g.main_cols();
                let new_layout: Vec<(u32, u32, String)> = new_ixs
                    .iter()
                    .map(|&c| {
                        let w = g.col_width(c).max(1);
                        let label = crate::addr::ui_column_fragment(c, mc);
                        (c as u32, w as u32, label)
                    })
                    .collect();
                spreadsheet_set_column_layout(sid, new_layout);
                col_ixs_cb = new_ixs;
                // Repopulate cells with updated column viewport
                let dr: Vec<usize> = display_rows_for_cb.borrow().clone();
                let new_col_widths: HashMap<usize, usize> = col_ixs_cb.iter()
                    .map(|&c| (c, g.col_width(c).max(1)))
                    .collect();
                let new_row_agg = compute::compute_row_agg_func(g, &dr, hr_cb, mc);
                fill_cells(
                    &sheet_cb, &dr, &col_ixs_cb, &new_col_widths,
                    g, hr_cb, g.main_rows(), mc, MARGIN_COLS, data_width_cb,
                    cursor.row, cursor.col, &new_row_agg,
                );
            } else {
                // Cursor moved within the current viewport — refresh cells to ensure
                // formatted display values are used (commits overwrite cells with raw values).
                let rec = app.core.workbook.active_sheet().clone();
                let g = &rec.grid;
                let mc = g.main_cols();
                let cursor = app.core.cursor;
                let dr: Vec<usize> = display_rows_for_cb.borrow().clone();
                let new_col_widths: HashMap<usize, usize> = col_ixs_cb.iter()
                    .map(|&c| (c, g.col_width(c).max(1)))
                    .collect();
                let new_row_agg = compute::compute_row_agg_func(g, &dr, hr_cb, g.main_rows());
                fill_cells(
                    &sheet_cb, &dr, &col_ixs_cb, &new_col_widths,
                    g, hr_cb, g.main_rows(), mc, MARGIN_COLS, data_width_cb,
                    cursor.row, cursor.col, &new_row_agg,
                );
            }
        }
    });

    // ── Commit edit callback: persist cell edits to workbook ──────────
    // After the commit, re-align the committed cell so the grid shows
    // the correctly-aligned display text (spreadsheet_commit_edit stores
    // the RAW value in the cells HashMap, but display text must be aligned).
    let commit_sheet = spreadsheet.clone();
    let app_ptr_ce = app_ptr;
    add_commit_edit_callback(move |display_row, col, value| {
        let app = unsafe { &mut *app_ptr_ce };
        let dr = display_rows_for_ce.borrow();
        let logical_row = dr.get(display_row as usize).copied().unwrap_or(0);
        let main_row = logical_row.saturating_sub(hr_ce);
        let main_col = col.saturating_sub(lm_ce as u32);
        let addr = CellAddr::Main { row: main_row as u32, col: main_col as u32 };
        crate::debug_log::log(&format!(
            "COMMIT_CB display_row={} col={} logical_row={} hr_ce={} lm_ce={} main_row={} main_col={} addr={:?} value={:?} app_cursor_row={} app_cursor_col={}",
            display_row, col, logical_row, hr_ce, lm_ce, main_row, main_col, addr, value, app.core.cursor.row, app.core.cursor.col
        ));
        let sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
        let op = Op::SetCell { addr, value: value.clone() };
        let wbo = WorkbookOp::SheetOp { sheet_id, op };
        if let Some(ref p) = app.core.path.clone() {
            let mut active_sheet = sheet_id;
            let _ = crate::io::commit_workbook_op(
                p,
                &mut app.core.offset,
                &mut app.core.workbook,
                &mut active_sheet,
                &wbo,
            );
            app.core.ops_applied = app.core.ops_applied.saturating_add(1);
        } else {
            // No live file: apply the value to the in-memory grid directly so the
            // committed text is visible (commit_workbook_op is skipped without a
            // path, which previously made the text vanish after Enter).
            app.core.workbook.active_sheet_mut().grid.set(&addr, value);
            app.core.ops_applied = app.core.ops_applied.saturating_add(1);
        }
        // Re-align the committed cell's display text (spreadsheet_commit_edit
        // stored the raw value, but we need the aligned version).
        let rec = app.core.workbook.active_sheet().clone();
        let g = &rec.grid;
        if let Some(effective) = logical_row.checked_sub(hr_ce).and_then(|mr| {
            g.get(&CellAddr::Main { row: mr as u32, col: main_col as u32 })
        }) {
            let formatted = crate::ui_core::format_cell_display(g, &addr, effective);
            let fw = formatted.width();
            let cw = g.col_width(col as usize).max(1);
            let align = crate::ui_core::effective_cell_align(g, &addr, &formatted);
            let aligned = if fw > cw
                && (align.is_none() || align == Some(crate::grid::TextAlign::Left))
            {
                // Text would spill into adjacent columns — keep the full text
                // (matching fill_cells spill logic).
                formatted
            } else {
                crate::ui_core::align_cell_display(formatted, cw, align)
            };
            commit_sheet.set_cell(display_row, col, &aligned);
        }
    });

    // ── Go-to (Ctrl+G) callback ───────────────────────────────────────────
    // The ratatui reference jump target for Ctrl+G is A1000. This callback
    // recomputes the visible viewport around that cell, repopulates the
    // widget, and positions the widget cursor there so the formula bar shows
    // the target address (e.g. `A1000`).
    let goto_sheet = spreadsheet.clone();
    let app_ptr_goto = app_ptr;
    let hr_goto = hr;
    let lm_goto = MARGIN_COLS;
    let sid_goto = sid;
    add_goto_callback(move || {
        let app = unsafe { &mut *app_ptr_goto };
        let cursor = crate::grid::SheetCursor {
            row: crate::grid::HEADER_ROWS + 999,
            col: lm_goto,
        };
        // Ensure the jump target (A1000) actually exists in the grid so it is
        // treated as a main row (label "1000") rather than spilling into a
        // footer row. This mirrors ratatui's Go-To, which extends the grid to
        // the jump target; without it, on a small workbook the far cursor is
        // clamped to a footer label (e.g. `_996`) instead of `A1000`.
        {
            let sht = app.core.workbook.active_sheet_mut();
            let target_main_row = cursor.row - crate::grid::HEADER_ROWS + 1; // = 1000
            if sht.grid.main_rows() < target_main_row {
                sht.grid.set_main_size(target_main_row, sht.grid.main_cols());
            }
        }
        // Recompute the viewport around the target cell.
        let (new_display_rows, new_mr, new_mc, new_ixs) = {
            let rec = app.core.workbook.active_sheet().clone();
            let (ndr, _) = crate::ui_core::visible_row_indices(&rec, cursor, data_rows_cb, 0);
            let nmr = rec.grid.main_rows();
            let nmc = rec.grid.main_cols();
            let (mut nix, _) = crate::ui_core::visible_col_indices(&rec, cursor, data_cols_cb, 0);
            {
                let sht = app.core.workbook.active_sheet_mut();
                crate::ui_core::trim_visible_cols_to_width(
                    &mut sht.grid, &mut nix, cursor.col, data_width_cb,
                );
            }
            (ndr, nmr, nmc, nix)
        };
        let rec = app.core.workbook.active_sheet().clone();
        let target_logical = crate::grid::HEADER_ROWS + 999;
        let display_ri = new_display_rows
            .iter()
            .position(|&r| r == target_logical)
            .unwrap_or(0);
        app.core.cursor = cursor;
        let boundary_title = format!(
            "corro  {}r × {}c  ops {}",
            new_mr, new_mc, app.core.ops_applied
        );
        spreadsheet_set_border_title(sid_goto, &boundary_title);
        let new_labels: Vec<(u32, String)> = new_display_rows
            .iter()
            .enumerate()
            .map(|(idx, &r)| (idx as u32, crate::addr::ui_row_label(r, new_mr)))
            .collect();
        spreadsheet_set_row_labels(sid_goto, new_labels);
        let g = &rec.grid;
        let new_layout: Vec<(u32, u32, String)> = new_ixs
            .iter()
            .map(|&c| {
                let w = g.col_width(c).max(1);
                let label = crate::addr::ui_column_fragment(c, new_mc);
                (c as u32, w as u32, label)
            })
            .collect();
        spreadsheet_set_column_layout(sid_goto, new_layout);
        spreadsheet_set_grid_config(sid_goto, lm_goto as u32, new_mc as u32);
        let new_col_widths: HashMap<usize, usize> =
            new_ixs.iter().map(|&c| (c, rec.grid.col_width(c).max(1))).collect();
        let new_row_agg = compute::compute_row_agg_func(&rec.grid, &new_display_rows, hr_goto, new_mr);
        fill_cells(
            &goto_sheet,
            &new_display_rows,
            &new_ixs,
            &new_col_widths,
            &rec.grid,
            hr_goto,
            new_mr,
            new_mc,
            MARGIN_COLS,
            data_width_cb,
            cursor.row,
            cursor.col,
            &new_row_agg,
        );
        // Position the widget cursor on the target cell (A1000).
        goto_sheet.set_cursor(display_ri as u32, lm_goto as u32);
        let target_val = rec
            .grid
            .get(&CellAddr::Main { row: 999, col: 0 })
            .unwrap_or_default();
        goto_sheet.set_raw_cell(display_ri as u32, lm_goto as u32, &target_val);
        *dr_for_goto_cb.borrow_mut() = new_display_rows.clone();
        *dr_for_goto_ce.borrow_mut() = new_display_rows.clone();
    });

    _backend.run().map_err(|e| format!("pancurses error: {e}"))?;
    Ok(())
}




