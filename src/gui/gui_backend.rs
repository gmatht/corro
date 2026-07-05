use rustxwidgets::prelude::*;
use rustxwidgets::core::DrawContext;

use std::cell::{Cell, RefCell};
use std::collections::HashMap;
use std::rc::Rc;

use crate::grid::{CellAddr, SheetCursor, HEADER_ROWS, MARGIN_COLS};
use crate::ops::{Op, WorkbookOp};
use crate::ui_core;

use super::compute::{self, CellDisplayStyle};
use super::dialogs;
use super::render::{self, CellSink};

use rustxwidgets::core::key::{normalize, RETURN, ESCAPE, BACKSPACE, DELETE, LEFT, UP, RIGHT, DOWN, TAB, HOME, END, PAGE_UP, PAGE_DOWN, F1, F2, ALT_L, ALT_R};

const KEYLOG_PATH: &str = "/tmp/corro_keylog.txt";

fn key_name(keyval: u32) -> String {
    if keyval == 0 { return "MENU".into(); }
    match normalize(keyval) {
        RETURN    => "RETURN".into(),
        ESCAPE    => "ESCAPE".into(),
        BACKSPACE => "BACKSPACE".into(),
        DELETE    => "DELETE".into(),
        LEFT      => "LEFT".into(),
        UP        => "UP".into(),
        RIGHT     => "RIGHT".into(),
        DOWN      => "DOWN".into(),
        TAB       => "TAB".into(),
        HOME      => "HOME".into(),
        END       => "END".into(),
        PAGE_UP   => "PAGE_UP".into(),
        PAGE_DOWN => "PAGE_DOWN".into(),
        F1        => "F1".into(),
        F2        => "F2".into(),
        ALT_L     => "ALT_L".into(),
        ALT_R     => "ALT_R".into(),
        k if (32..=126).contains(&k) => format!("'{}'", char::from_u32(k).unwrap_or('?')),
        k => format!("0x{k:X}"),
    }
}

fn format_cell(state: &GuiState) -> String {
    format!("R{}C{}", state.last_row.get(), state.last_col.get())
}

fn append_keylog(msg: &str) {
    if let Ok(mut f) = std::fs::OpenOptions::new()
        .create(true).append(true)
        .open(KEYLOG_PATH)
    {
        use std::io::Write;
        let _ = write!(f, "{}", msg);
    }
}

fn log_ui_action(action: &str, detail: &str) {
    log_key_action(0, action, detail)
}

fn log_key_action(keyval: u32, action: &str, detail: &str) {
    if let Ok(mut f) = std::fs::OpenOptions::new()
        .create(true).append(true)
        .open(KEYLOG_PATH)
    {
        use std::io::Write;
        let _ = writeln!(f, "KEY: {}  ACTION: {action}  DETAIL: {detail}", key_name(keyval));
    }
}

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const FONT_SIZE: f64 = 12.0;
const ROW_H: f64 = 20.0;
const HEADER_H: f64 = 24.0;
const ROW_LABEL_W: f64 = 50.0;
const MAX_RENDER_ROWS: usize = 500;
const MAX_RENDER_COLS: usize = 50;
const CHAR_W: f64 = 7.2;

// ---------------------------------------------------------------------------
// Mode
// ---------------------------------------------------------------------------

#[derive(Clone, Copy, PartialEq)]
enum GuiMode {
    Normal,
    Help,
}

// ---------------------------------------------------------------------------
// CanvasSink
// ---------------------------------------------------------------------------

struct GuiCanvasSink {
    cells: RefCell<HashMap<(u32, u32), String>>,
    styles: RefCell<HashMap<(u32, u32), CellDisplayStyle>>,
    raw_values: RefCell<HashMap<(u32, u32), String>>,
    cursor_pos: Cell<Option<(u32, u32)>>,
}

impl GuiCanvasSink {
    fn new() -> Self {
        GuiCanvasSink {
            cells: RefCell::new(HashMap::new()),
            styles: RefCell::new(HashMap::new()),
            raw_values: RefCell::new(HashMap::new()),
            cursor_pos: Cell::new(None),
        }
    }

}

impl CellSink for GuiCanvasSink {
    fn set_cell(&mut self, row: u32, col: u32, text: &str) {
        self.cells.borrow_mut().insert((row, col), text.to_string());
    }
    fn set_cell_style(&mut self, row: u32, col: u32, style: CellDisplayStyle) {
        self.styles.borrow_mut().insert((row, col), style);
    }
    fn set_raw_cell(&mut self, row: u32, col: u32, text: &str) {
        self.raw_values.borrow_mut().insert((row, col), text.to_string());
    }
    fn set_cursor(&mut self, row: u32, col: u32) {
        self.cursor_pos.set(Some((row, col)));
    }
}

// ---------------------------------------------------------------------------
// Save-before-quit helper
// ---------------------------------------------------------------------------

fn save_before_quit(state: &GuiState) {
    let _ = std::fs::write("/tmp/corro_quit_called.txt", "save_before_quit called\n");
    eprintln!("DEBUG save_before_quit called");
    let fx_text = state.formula_entry.get_text().unwrap_or_default();
    let cell = format_cell(state);
    log_ui_action("quit", &format!("cell={} fx_textarea={:?}", cell, fx_text));
    log_ui_action("fx_text_on_exit", &format!("{:?}", fx_text));

    // Write GUI‑state snapshot so opencode can see the UI state at exit time.
    let snapshot = format!(
        "--- corro GUI snapshot (on exit) ---\n\
         cell={}\n\
         editing={}\n\
         mode={}\n\
         fx_text={}\n\
         edit_buf={}\n\
         key_counter={}\n\
         --- end snapshot ---\n",
         cell, state.editing.get(), match state.mode.get() { GuiMode::Normal => "Normal", GuiMode::Help => "Help" },
        fx_text, state.edit_buf.borrow(),
         state.key_counter.get(),
    );
    let _ = std::fs::write("/tmp/corro_gui_snapshot.txt", &snapshot);

    if state.editing.get() {
        commit_edit(state);
    }
    match state.rxapp.try_quit() {
        Ok(()) => std::process::exit(0),
        Err(e) => {
            eprintln!("try_quit failed (main loop not yet started?), force-exiting: {e}");
            std::process::exit(0);
        }
    }
}

// ---------------------------------------------------------------------------
// Shared state
// ---------------------------------------------------------------------------

struct GuiState {
    app: *mut super::App,
    rxapp: rustxwidgets::App,
    canvas: Canvas,
    formula_entry: Entry,
    addr_label: Label,
    status_label: Label,
    editing: Cell<bool>,
    edit_buf: RefCell<String>,
    mode: Cell<GuiMode>,
    last_row: Cell<usize>,
    last_col: Cell<usize>,
    data_rows: Cell<usize>,
    data_cols: Cell<usize>,
    last_key: Cell<u32>,
    key_counter: Cell<u64>,
    entry_processed_key: Cell<bool>,
}

// ---------------------------------------------------------------------------
// Rendering
// ---------------------------------------------------------------------------

fn render_to(
    sink: &GuiCanvasSink,
    dc: &mut dyn DrawContext,
    col_ixs: &[usize],
    col_widths: &HashMap<usize, usize>,
    display_rows: &[usize],
    _mr: usize,
    _mc: usize,
    cursor_row: usize,
    cursor_col: usize,
    is_editing: bool,
    edit_text: &str,
    selection_anchor: Option<(usize, usize)>,
) {
    let cells = sink.cells.borrow();
    let styles = sink.styles.borrow();

    for (ri, &logical_row) in display_rows.iter().enumerate().take(MAX_RENDER_ROWS) {
        let ry = HEADER_H + ri as f64 * ROW_H;
        let is_sel_row = selection_anchor.map_or(false, |(ar, ac)| {
            let r1 = ar.min(cursor_row);
            let r2 = ar.max(cursor_row);
            let _c1 = ac.min(cursor_col);
            let _c2 = ac.max(cursor_col);
            logical_row >= r1 && logical_row <= r2
        });

        for (ci, &c) in col_ixs.iter().enumerate().take(MAX_RENDER_COLS) {
            let cw = *col_widths.get(&c).unwrap_or(&8) as f64 * CHAR_W;
            let cx = ROW_LABEL_W + col_ixs.iter().take(ci).map(|&pc| *col_widths.get(&pc).unwrap_or(&8) as f64 * CHAR_W).sum::<f64>();

            let key = (ri as u32, c as u32);
            let raw_text = cells.get(&key).map(|s| s.as_str()).unwrap_or("");
            let style_key = (ri as u32, c as u32);
            let style = styles.get(&style_key).copied().unwrap_or(CellDisplayStyle::Default);
            let is_current = logical_row == cursor_row && c == cursor_col;

            let bg = if is_current {
                if is_editing { (1.0, 1.0, 0.8, 1.0) } else { (0.8, 0.9, 1.0, 1.0) }
            } else if is_sel_row && selection_anchor.is_some() {
                (0.9, 0.95, 1.0, 1.0)
            } else {
                (1.0, 1.0, 1.0, 1.0)
            };

            dc.fill_rect(cx, ry, cw, ROW_H, bg.0, bg.1, bg.2, bg.3);

            // Selection highlight
            if is_current && !is_editing {
                dc.stroke_rect(cx, ry, cw, ROW_H, 0.0, 0.4, 0.8, 1.0, 2.0);
            }

            // Grid lines
            dc.stroke_rect(cx, ry, cw, ROW_H, 0.8, 0.8, 0.8, 1.0, 0.5);

            if !raw_text.is_empty() {
                match style {
                    CellDisplayStyle::Default => {
                        dc.draw_text(cx + 2.0, ry + 2.0, raw_text, "monospace", FONT_SIZE, 0.0, 0.0, 0.0, 1.0);
                    }
                    CellDisplayStyle::Cursor | CellDisplayStyle::Selected => {
                        dc.draw_text(cx + 2.0, ry + 2.0, raw_text, "monospace", FONT_SIZE, 0.0, 0.0, 0.0, 1.0);
                    }
                    CellDisplayStyle::Aggregate | CellDisplayStyle::FooterAggregate => {
                        dc.draw_text(cx + 2.0, ry + 2.0, raw_text, "monospace", FONT_SIZE, 0.5, 0.5, 0.5, 1.0);
                    }
                    CellDisplayStyle::ActiveHeader | CellDisplayStyle::InactiveHeader => {
                        dc.draw_text(cx + 2.0, ry + 2.0, raw_text, "monospace", FONT_SIZE, 0.3, 0.3, 0.3, 1.0);
                    }
                }
            }
        }
    }

    // Edit overlay on cursor cell
    if is_editing && !edit_text.is_empty() {
        if let Some(pos) = col_ixs.iter().position(|&c| c == cursor_col) {
            let cw = *col_widths.get(&cursor_col).unwrap_or(&8) as f64 * CHAR_W;
            let cx = ROW_LABEL_W + col_ixs.iter().take(pos).map(|&pc| *col_widths.get(&pc).unwrap_or(&8) as f64 * CHAR_W).sum::<f64>();
            if let Some(pos_r) = display_rows.iter().position(|&r| r == cursor_row) {
                let ry = HEADER_H + pos_r as f64 * ROW_H;
                dc.fill_rect(cx, ry, cw, ROW_H, 1.0, 1.0, 0.8, 1.0);
                dc.draw_text(cx + 2.0, ry + 2.0, edit_text, "monospace", FONT_SIZE, 0.0, 0.0, 0.0, 1.0);
            }
        }
    }
}

fn render_grid(dc: &mut dyn DrawContext, state: &GuiState, w: i32, h: i32) {
    dc.clear(0.94, 0.94, 0.94, 1.0);
    dc.clip(0.0, 0.0, w as f64, h as f64);

    let app = unsafe { &*state.app };
    let hr = HEADER_ROWS;
    let lm = MARGIN_COLS;
    let cursor_row = state.last_row.get();
    let cursor_col = state.last_col.get();

    let display_rows: Vec<usize> = {
        let sheet = app.core.workbook.active_sheet();
        ui_core::visible_row_indices(sheet, app.core.cursor, state.data_rows.get(), 0).0
    };
    let col_ixs: Vec<usize> = {
        let sheet = app.core.workbook.active_sheet();
        ui_core::visible_col_indices(sheet, app.core.cursor, state.data_cols.get(), 0).0
    };
    let mr = app.core.workbook.active_sheet().grid.main_rows();
    let mc = app.core.workbook.active_sheet().grid.main_cols();

    // Row headers
    for (ri, &logical_row) in display_rows.iter().enumerate().take(MAX_RENDER_ROWS) {
        let ry = HEADER_H + ri as f64 * ROW_H;
        let label = crate::addr::ui_row_label(logical_row, mr);
        let (_, _, tw, _) = dc.text_extents(&label, "monospace", FONT_SIZE);
        dc.fill_rect(0.0, ry, ROW_LABEL_W, ROW_H, 0.9, 0.9, 0.9, 1.0);
        dc.draw_text(ROW_LABEL_W - tw - 4.0, ry + 2.0, &label, "monospace", FONT_SIZE, 0.3, 0.3, 0.3, 1.0);
    }

    // Column headers
    for (ci, &c) in col_ixs.iter().enumerate().take(MAX_RENDER_COLS) {
        let cw = sheet_rec_col_width(&app.core.workbook.active_sheet(), c) as f64 * CHAR_W;
        let cx = ROW_LABEL_W + col_ixs.iter().take(ci).map(|&pc| sheet_rec_col_width(&app.core.workbook.active_sheet(), pc) as f64 * CHAR_W).sum::<f64>();
        let col_name = crate::addr::ui_column_fragment(c, mc);
        dc.fill_rect(cx, 0.0, cw, HEADER_H, 0.9, 0.9, 0.9, 1.0);
        let (_, _, tw, _) = dc.text_extents(&col_name, "monospace", FONT_SIZE);
        dc.draw_text(cx + (cw - tw) / 2.0, (HEADER_H - FONT_SIZE * 1.2) / 2.0, &col_name, "monospace", FONT_SIZE, 0.3, 0.3, 0.3, 1.0);
    }

    let col_widths: HashMap<usize, usize> = col_ixs.iter()
        .map(|&c| (c, sheet_rec_col_width(&app.core.workbook.active_sheet(), c)))
        .collect();

    let row_agg_func = compute::compute_row_agg_func(
        &app.core.workbook.active_sheet().grid,
        &display_rows, hr, mr,
    );

    // Fill cells via the shared render pipeline
    let sink_snapshot: HashMap<(u32, u32), String>;
    {
        let mut sink = GuiCanvasSink::new();
        render::fill_cells(
            &mut sink, &display_rows, &col_ixs, &col_widths,
            &app.core.workbook.active_sheet().grid,
            hr, mr, mc,
            lm, state.data_cols.get(),
            cursor_row, cursor_col,
            &row_agg_func,
        );
        sink_snapshot = sink.cells.borrow().clone();
        render_to(
            &sink, dc, &col_ixs, &col_widths, &display_rows,
            mr, mc,
            cursor_row, cursor_col,
            state.editing.get(),
            &state.edit_buf.borrow(),
            app.core.anchor.map(|a| (a.row, a.col)),
        );
    }

    // Status line at bottom
    if h as f64 > HEADER_H + 20.0 {
        dc.fill_rect(0.0, h as f64 - 20.0, w as f64, 20.0, 0.9, 0.9, 0.9, 1.0);
    }

    // Diagnostic: read first data cell from grid + sink
    let first_row = hr;
    let first_col = lm;
    let main_row = first_row.saturating_sub(hr);
    let main_col = first_col.saturating_sub(lm);
    let addr = crate::grid::CellAddr::Main { row: main_row as u32, col: main_col as u32 };
    let cell_val = app.core.workbook.active_sheet().grid.get(&addr).unwrap_or_default();
    let is_editing = if state.editing.get() { "EDIT" } else { "NORM" };
    let eb = state.edit_buf.borrow().clone();
    let buf_display = if eb.is_empty() { "(empty)" } else { &eb };
    let sink_key_col = MARGIN_COLS as u32;
    let sink_first = sink_snapshot.get(&(0u32, sink_key_col)).cloned().unwrap_or_default();
    let cur = (state.last_row.get(), state.last_col.get());
    let lk = state.last_key.get();
    let kc = state.key_counter.get();
    dc.draw_text(100.0, h as f64 - 140.0, &format!("Grid(0,0)='{cell_val}' Sink(0,{sink_key_col})='{sink_first}' Cur=({},{})", cur.0, cur.1), "monospace", 12.0, 0.0, 0.5, 0.0, 1.0);
    dc.draw_text(100.0, h as f64 - 120.0, &format!("lastKey=0x{lk:04x} cnt={kc}", ), "monospace", 12.0, 1.0, 0.0, 0.0, 1.0);
    dc.draw_text(100.0, h as f64 - 100.0, &format!("Mode:{is_editing} Buf:'{buf_display}'"), "monospace", 14.0, 0.5, 0.0, 0.5, 1.0);
}

fn sheet_rec_col_width(sheet: &crate::ops::SheetState, col: usize) -> usize {
    sheet.grid.col_width(col).max(1)
}

// ---------------------------------------------------------------------------
// Keyboard handling
// ---------------------------------------------------------------------------

fn handle_key(keyval: u32, state_rc: &Rc<GuiState>) -> bool {
    let state: &GuiState = &**state_rc;
    state.last_key.set(keyval);
    let app = unsafe { &mut *state.app };
    let key = normalize(keyval);

    match state.mode.get() {
        GuiMode::Help => {
            if key == ESCAPE {
                log_key_action(keyval, "help_exit", "");
                state.mode.set(GuiMode::Normal);
                state.canvas.queue_redraw();
                return true;
            }
            log_key_action(keyval, "help_ignore", "help mode blocks all keys except ESCAPE");
            return true;
        }
        _ => {}
    }

    if state.editing.get() {
        return handle_edit_key(key, state);
    }

    match key {
        F1 => {
            log_key_action(keyval, "help_mode", "");
            state.mode.set(GuiMode::Help);
            state.canvas.queue_redraw();
            true
        }
        F2 => {
            log_key_action(keyval, "start_edit", &format!("cell={}", format_cell(state)));
            start_edit(state);
            true
        }
        RETURN => {
            // When editing=false but the formula entry has non-empty text, the
            // user typed characters but editing state was not established yet
            // (GTK4 timing race during present()).  Start editing and commit
            // immediately to avoid losing the edit.
            if !state.editing.get() {
                if let Some(text) = state.formula_entry.get_text() {
                    if !text.is_empty() {
                        state.editing.set(true);
                        *state.edit_buf.borrow_mut() = text;
                        return handle_edit_key(key, state);
                    }
                }
                // When window CAPTURE consumed printable chars and pushed to
                // edit_buf directly (bypassing the entry widget), the entry
                // text is empty but edit_buf has content.  Commit from there.
                if !state.edit_buf.borrow().is_empty() {
                    state.editing.set(true);
                    return handle_edit_key(key, state);
                }
            }
            log_key_action(keyval, "move_cursor_down", &format!("cell={}", format_cell(state)));
            move_cursor(state, 1, 0);
            true
        }
        TAB => {
            log_key_action(keyval, "move_cursor_right", &format!("cell={}", format_cell(state)));
            move_cursor(state, 0, 1);
            true
        }
        ESCAPE => {
            log_key_action(keyval, "cancel_nav", "");
            app.core.anchor = Some(SheetCursor {
                row: state.last_row.get(),
                col: state.last_col.get(),
            });
            state.canvas.queue_redraw();
            true
        }
        LEFT => {
            log_key_action(keyval, "move_cursor_left", &format!("cell={}", format_cell(state)));
            move_cursor(state, 0, -1);
            true
        }
        RIGHT => {
            log_key_action(keyval, "move_cursor_right", &format!("cell={}", format_cell(state)));
            move_cursor(state, 0, 1);
            true
        }
        UP => {
            log_key_action(keyval, "move_cursor_up", &format!("cell={}", format_cell(state)));
            if state.last_row.get() > HEADER_ROWS {
                move_cursor(state, -1, 0);
            }
            true
        }
        DOWN => {
            log_key_action(keyval, "move_cursor_down", &format!("cell={}", format_cell(state)));
            move_cursor(state, 1, 0);
            true
        }
        HOME => {
            log_key_action(keyval, "move_cursor_home", &format!("cell={}", format_cell(state)));
            state.last_col.set(MARGIN_COLS);
            update_state_cursor(state, state.last_row.get(), MARGIN_COLS);
            true
        }
        END => {
            log_key_action(keyval, "move_cursor_end", &format!("cell={}", format_cell(state)));
            state.last_col.set(state.last_col.get() + 10);
            update_state_cursor(state, state.last_row.get(), state.last_col.get());
            true
        }
        PAGE_UP => {
            log_key_action(keyval, "move_cursor_page_up", &format!("cell={}", format_cell(state)));
            let dr = state.data_rows.get();
            let new_row = state.last_row.get().saturating_sub(dr);
            state.last_row.set(new_row.max(HEADER_ROWS));
            update_state_cursor(state, state.last_row.get(), state.last_col.get());
            true
        }
        PAGE_DOWN => {
            log_key_action(keyval, "move_cursor_page_down", &format!("cell={}", format_cell(state)));
            let dr = state.data_rows.get();
            state.last_row.set(state.last_row.get() + dr);
            update_state_cursor(state, state.last_row.get(), state.last_col.get());
            true
        }
        DELETE => {
            log_key_action(keyval, "delete_cell", &format!("cell={}", format_cell(state)));
            handle_delete(state);
            true
        }
        BACKSPACE => {
            log_key_action(keyval, "delete_cell", &format!("cell={}", format_cell(state)));
            handle_delete(state);
            true
        }
        _ if (32..=126).contains(&key) => {
            let ch = char::from_u32(key).unwrap_or('?');
            log_key_action(keyval, "start_edit_with", &format!("char={ch} cell={}", format_cell(state)));
            start_edit_with(state, ch);
            true
        }
        _ => false,
    }
}

fn handle_edit_key(key: u32, state: &GuiState) -> bool {
    match key {
        RETURN => {
            log_key_action(key, "commit_edit", &format!("cell={} mode=edit", format_cell(state)));
            commit_edit(state);
            move_cursor(state, 1, 0);
            true
        }
        ESCAPE => {
            log_key_action(key, "cancel_edit", &format!("cell={} mode=edit", format_cell(state)));
            state.editing.set(false);
            state.edit_buf.borrow_mut().clear();
            state.mode.set(GuiMode::Normal);
            update_formula_bar(state, state.last_row.get(), state.last_col.get());
            state.canvas.queue_redraw();
            true
        }
        TAB => {
            log_key_action(key, "commit_edit_tab", &format!("cell={} mode=edit", format_cell(state)));
            commit_edit(state);
            move_cursor(state, 0, 1);
            true
        }
        BACKSPACE => {
            log_key_action(key, "edit_backspace", &format!("cell={} mode=edit", format_cell(state)));
            state.edit_buf.borrow_mut().pop();
            state.canvas.queue_redraw();
            true
        }
        DELETE => {
            log_key_action(key, "edit_clear", &format!("cell={} mode=edit", format_cell(state)));
            state.edit_buf.borrow_mut().clear();
            state.canvas.queue_redraw();
            true
        }
        LEFT => {
            log_key_action(key, "commit_edit_left", &format!("cell={} mode=edit", format_cell(state)));
            commit_edit(state);
            let c = state.last_col.get();
            if c > 0 {
                state.last_col.set(c - 1);
                update_state_cursor(state, state.last_row.get(), state.last_col.get());
            }
            true
        }
        RIGHT => {
            log_key_action(key, "commit_edit_right", &format!("cell={} mode=edit", format_cell(state)));
            commit_edit(state);
            state.last_col.set(state.last_col.get() + 1);
            update_state_cursor(state, state.last_row.get(), state.last_col.get());
            true
        }
        UP => {
            log_key_action(key, "commit_edit_up", &format!("cell={} mode=edit", format_cell(state)));
            commit_edit(state);
            if state.last_row.get() > HEADER_ROWS {
                state.last_row.set(state.last_row.get() - 1);
                update_state_cursor(state, state.last_row.get(), state.last_col.get());
            }
            true
        }
        DOWN => {
            log_key_action(key, "commit_edit_down", &format!("cell={} mode=edit", format_cell(state)));
            commit_edit(state);
            state.last_row.set(state.last_row.get() + 1);
            update_state_cursor(state, state.last_row.get(), state.last_col.get());
            true
        }
        _ if (32..=126).contains(&key) => {
            let ch = char::from_u32(key).unwrap_or('?');
            log_key_action(key, "edit_insert", &format!("char={ch} cell={} mode=edit", format_cell(state)));
            state.edit_buf.borrow_mut().push(ch);
            state.canvas.queue_redraw();
            true
        }
        _ => true,
    }
}

// ---------------------------------------------------------------------------
// Edit operations
// ---------------------------------------------------------------------------

fn start_edit(state: &GuiState) {
    state.editing.set(true);
    state.edit_buf.borrow_mut().clear();
    state.formula_entry.set_text("");
    state.formula_entry.grab_focus();
    state.canvas.queue_redraw();
}

fn start_edit_with(state: &GuiState, ch: char) {
    let already_editing = state.editing.get();
    state.editing.set(true);
    let s = ch.to_string();
    if state.edit_buf.borrow().is_empty() {
        state.edit_buf.borrow_mut().push_str(&s);
    } else {
        let mut buf = state.edit_buf.borrow_mut();
        buf.push_str(&s);
    }
    if !already_editing {
        // Set the entry text to the typed character so it is visible in the
        // formula bar.  Using set_text(&s) instead of set_text("") ensures
        // the entry displays the first character (important when the window
        // CAPTURE-phase controller handles the key and stops propagation,
        // preventing the entry's default handler from inserting the char).
        //
        // set_text() triggers connect_changed -> on_formula_entry_changed,
        // which would overwrite edit_buf with the entry text.  The starts_with
        // guard in on_formula_entry_changed accepts this because the entry
        // text ("4") is a forward extension of current edit_buf (""), and
        // after restoring the saved value the result is identical.
        let saved = state.edit_buf.borrow().clone();
        state.formula_entry.set_text(&s);
        *state.edit_buf.borrow_mut() = saved;
        state.formula_entry.grab_focus();
    }
    state.canvas.queue_redraw();
}

fn commit_edit(state: &GuiState) {
    state.editing.set(false);
    state.mode.set(GuiMode::Normal);
    let val = state.edit_buf.borrow().clone();
    if !val.is_empty() {
        let app = unsafe { &mut *state.app };
        let row = state.last_row.get();
        let col = state.last_col.get();
        let main_row = row.saturating_sub(HEADER_ROWS);
        let main_col = col.saturating_sub(MARGIN_COLS);
        let addr = CellAddr::Main { row: main_row as u32, col: main_col as u32 };
        app.core.workbook.active_sheet_mut().grid.set(&addr, val.clone());
        let sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
        let op = Op::SetCell { addr, value: val };
        let wbo = WorkbookOp::SheetOp { sheet_id, op };
        if let Some(ref p) = app.core.path.clone() {
            let mut active_sheet = sheet_id;
            if let Err(e) = crate::io::commit_workbook_op(
                p, &mut app.core.offset, &mut app.core.workbook,
                &mut active_sheet, &wbo,
            ) {
                eprintln!("ERROR: commit_workbook_op failed: {e} (path={})", p.display());
            } else {
                app.core.ops_applied = app.core.ops_applied.saturating_add(1);
            }
        }
        let main_rows = crate::addr::MainRows(app.core.workbook.active_sheet().grid.main_rows());
        let main_cols = crate::addr::MainCols(app.core.workbook.active_sheet().grid.main_cols());
        app.core.status = format!("Set cell {}", crate::addr::sheet_cursor_to_addr(
            crate::addr::LogicalRow(row),
            crate::addr::GlobalCol(col),
            main_rows,
            main_cols,
        ));
        recompute_viewport(state);
    }
    state.edit_buf.borrow_mut().clear();
    state.canvas.queue_redraw();
}

fn handle_delete(state: &GuiState) {
    let app = unsafe { &mut *state.app };
    let row = state.last_row.get();
    let col = state.last_col.get();
    let main_row = row.saturating_sub(HEADER_ROWS);
    let main_col = col.saturating_sub(MARGIN_COLS);
    if let Some(anchor) = app.core.anchor {
        let r1 = anchor.row.min(row);
        let r2 = anchor.row.max(row);
        let c1 = anchor.col.min(col);
        let c2 = anchor.col.max(col);
        let ro = r2 - r1 + 1;
        let co = c2 - c1 + 1;
        if ro > 1 || co > 1 {
            for r in r1..=r2 {
                for c in c1..=c2 {
                    let main_r = r.saturating_sub(HEADER_ROWS);
                    let main_c = c.saturating_sub(MARGIN_COLS);
                    let addr = CellAddr::Main { row: main_r as u32, col: main_c as u32 };
                    app.core.workbook.active_sheet_mut().grid.set(&addr, String::new());
                }
            }
            let sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
            let addr = CellAddr::Main { row: main_row as u32, col: main_col as u32 };
            let op = Op::SetCell { addr, value: String::new() };
            let wbo = WorkbookOp::SheetOp { sheet_id, op };
            if let Some(ref p) = app.core.path.clone() {
                let mut active_sheet = sheet_id;
                if let Err(e) = crate::io::commit_workbook_op(
                    p, &mut app.core.offset, &mut app.core.workbook,
                    &mut active_sheet, &wbo,
                ) {
                    eprintln!("ERROR: commit_workbook_op failed: {e} (path={})", p.display());
                } else {
                    app.core.ops_applied = app.core.ops_applied.saturating_add(1);
                }
            }
            app.core.status = "Cleared selection".into();
            recompute_viewport(state);
            state.canvas.queue_redraw();
            return;
        }
    }
    let addr = CellAddr::Main { row: main_row as u32, col: main_col as u32 };
    app.core.workbook.active_sheet_mut().grid.set(&addr, String::new());
    let sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
    let op = Op::SetCell { addr, value: String::new() };
    let wbo = WorkbookOp::SheetOp { sheet_id, op };
    if let Some(ref p) = app.core.path.clone() {
        let mut active_sheet = sheet_id;
        if let Err(e) = crate::io::commit_workbook_op(
            p, &mut app.core.offset, &mut app.core.workbook,
            &mut active_sheet, &wbo,
        ) {
            eprintln!("ERROR: commit_workbook_op failed: {e} (path={})", p.display());
        } else {
            app.core.ops_applied = app.core.ops_applied.saturating_add(1);
        }
    }
    recompute_viewport(state);
    state.canvas.queue_redraw();
}

fn recompute_viewport(state: &GuiState) {
    let app = unsafe { &*state.app };
    let hr = HEADER_ROWS;
    let cursor_row = state.last_row.get();
    let cursor_col = state.last_col.get();
    let sheet = app.core.workbook.active_sheet();
    let (display_rows, _) = ui_core::visible_row_indices(sheet, app.core.cursor, state.data_rows.get(), 0);
    let (col_ixs, _) = ui_core::visible_col_indices(sheet, app.core.cursor, state.data_cols.get(), 0);
    if !display_rows.contains(&cursor_row) || !col_ixs.contains(&cursor_col) {
        let new_row = if cursor_row > display_rows.last().copied().unwrap_or(hr) {
            display_rows.first().copied().unwrap_or(hr)
        } else {
            cursor_row
        };
        let new_col = if cursor_col > col_ixs.last().copied().unwrap_or(MARGIN_COLS) {
            col_ixs.first().copied().unwrap_or(MARGIN_COLS)
        } else {
            cursor_col
        };
        update_state_cursor(state, new_row, new_col);
    }
}

fn move_cursor(state: &GuiState, dr: isize, dc: isize) {
    let row = state.last_row.get();
    let col = state.last_col.get();
    let app = unsafe { &mut *state.app };
    let mr = app.core.workbook.active_sheet().grid.main_rows();
    let mc = app.core.workbook.active_sheet().grid.main_cols() + MARGIN_COLS;
    let new_row = (row as isize + dr).max(HEADER_ROWS as isize).min((HEADER_ROWS + mr).max(HEADER_ROWS) as isize) as usize;
    let new_col = (col as isize + dc).max(MARGIN_COLS as isize).min(mc as isize - 1).max(MARGIN_COLS as isize) as usize;
    update_state_cursor(state, new_row, new_col);
}

fn update_state_cursor(state: &GuiState, row: usize, col: usize) {
    state.last_row.set(row);
    state.last_col.set(col);
    let app = unsafe { &mut *state.app };
    app.core.cursor.row = row;
    app.core.cursor.col = col;
    update_formula_bar(state, row, col);
    state.canvas.queue_redraw();
}

fn update_formula_bar(state: &GuiState, row: usize, col: usize) {
    let app = unsafe { &*state.app };
    let main_row = row.saturating_sub(HEADER_ROWS);
    let main_col = col.saturating_sub(MARGIN_COLS);
    let addr_str = crate::addr::sheet_cursor_to_addr(
        crate::addr::LogicalRow(row),
        crate::addr::GlobalCol(col),
        crate::addr::MainRows(app.core.workbook.active_sheet().grid.main_rows()),
        crate::addr::MainCols(app.core.workbook.active_sheet().grid.main_cols()),
    );
    state.addr_label.set_text(&addr_str.to_string());
    let addr = crate::grid::CellAddr::Main { row: main_row as u32, col: main_col as u32 };
    let val = app.core.workbook.active_sheet().grid.get(&addr).unwrap_or_default();
    state.formula_entry.set_text(&val);
    state.status_label.set_text(&app.core.status);
    if state.status_label.raw_handle().is_null() {
        // status label will be updated; no-op
    }
}

// ---------------------------------------------------------------------------
// Click handling
// ---------------------------------------------------------------------------

fn handle_click(x: f64, y: f64, state_rc: &Rc<GuiState>) {
    let state: &GuiState = &**state_rc;
    let app = unsafe { &mut *state.app };
    if x < ROW_LABEL_W || y < HEADER_H {
        return;
    }
    let col_ixs: Vec<usize> = {
        let sheet = app.core.workbook.active_sheet();
        ui_core::visible_col_indices(sheet, app.core.cursor, state.data_cols.get(), 0).0
    };
    let mut cx = ROW_LABEL_W;
    for &c in &col_ixs {
        let cw = sheet_rec_col_width(&app.core.workbook.active_sheet(), c) as f64 * CHAR_W;
        if x >= cx && x < cx + cw {
            let ri = ((y - HEADER_H) / ROW_H) as usize;
            let display_rows: Vec<usize> = {
                let sheet = app.core.workbook.active_sheet();
                ui_core::visible_row_indices(sheet, app.core.cursor, state.data_rows.get(), 0).0
            };
            if ri < display_rows.len() {
                let logical_row = display_rows[ri];
                state.last_row.set(logical_row);
                state.last_col.set(c);
                app.core.cursor.row = logical_row;
                app.core.cursor.col = c;
                if app.core.anchor.is_none() {
                    app.core.anchor = Some(SheetCursor { row: logical_row, col: c });
                }
                update_formula_bar(state, logical_row, c);
                start_edit(state);
                state.canvas.queue_redraw();
            }
            return;
        }
        cx += cw;
    }
}

// ---------------------------------------------------------------------------
// Menu building
// ---------------------------------------------------------------------------

fn build_menu(rxapp: &rustxwidgets::App, win: &Window, state: &Rc<GuiState>) -> Result<MenuBar, Box<dyn std::error::Error>> {
    use crate::gui::menu;

    let action_group = rxapp.ensure_action_group()?;

    // Build the full menu tree from shared definitions (rustxwidgets).
    // Prefix submenu labels with "_" so GTK4 assigns mnemonic accelerators
    // (ALT+F for File, ALT+E for Edit, etc.).
    let menubar_model = rxapp.build_menu_model(&menu::all_submenus(), "_")?;

    // Register action callbacks with state access
    let s = state.clone();
    for &items in &[menu::FILE_MENU, menu::EDIT_MENU, menu::VIEW_MENU, menu::INSERT_MENU, menu::FORMAT_MENU, menu::SHEET_MENU, menu::DATA_MENU, menu::HELP_MENU] {
        for item in items {
            let name = menu::action_kind_to_name(item.action);
            let name_owned = name.to_string();
            let state_cb = s.clone();
            menu::register_action(rxapp, name, move || handle_menu_action(&name_owned, &state_cb))?;
        }
    }

    let menubar = rxapp.new_menubar(&menubar_model, action_group)?;
    // Insert the action group on both the window and the menubar widget.
    // The menubar insertion is critical: when the popover surface processes
    // mnemonic key presses, gtk_widget_activate_action walks the popover's
    // parent chain (popover -> bar_item -> menubar -> ... -> window).  Without
    // the action group on the menubar, the walk may stop before reaching the
    // window on some GTK4 versions/configurations.
    win.insert_action_group("app", action_group);
    unsafe { menubar.insert_action_group("app", action_group); }

    Ok(menubar)
}

fn handle_menu_action(name: &str, state: &GuiState) {
    let app = unsafe { &mut *state.app };
    log_ui_action("menu_action", name);
    match name {
        "open" => {
            if let Some(path) = dialogs::file_open_dialog() {
                match crate::io::load_workbook_snapshot(&path) {
                    Ok(snapshot) => {
                        app.core.workbook = crate::ops::WorkbookState::from_snapshot(&snapshot);
                        app.core.offset = 0;
                        app.core.ops_applied = 0;
                        app.core.path = Some(path);
                        app.core.status = "Opened file".into();
                        recompute_viewport(state);
                        state.canvas.queue_redraw();
                    }
                    Err(e) => app.core.status = format!("Open error: {e}"),
                }
            }
        }
        "save" => {
            if let Some(ref p) = app.core.path.clone() {
                let snapshot = crate::ops::WorkbookSnapshot::from_workbook(&app.core.workbook);
                match crate::io::save_workbook(p, &snapshot) {
                    Ok(()) => app.core.status = "Saved".into(),
                    Err(e) => app.core.status = format!("Save error: {e}"),
                }
            }
        }
        "save_as" => {
            if let Some(path) = dialogs::file_save_dialog() {
                app.core.path = Some(path.clone());
                let snapshot = crate::ops::WorkbookSnapshot::from_workbook(&app.core.workbook);
                match crate::io::save_workbook(&path, &snapshot) {
                    Ok(()) => app.core.status = format!("Saved to {}", path.display()),
                    Err(e) => app.core.status = format!("Save error: {e}"),
                }
            }
        }
        "corro_quit" => {
            eprintln!("DEBUG handle_menu_action: corro_quit activated");
            save_before_quit(state);
        }
        "find" => dialogs::find_dialog(|result| {
            if let Some(text) = result {
                app.core.status = format!("Find: {text}");
            }
        }),
        "replace" => dialogs::replace_dialog(|result| {
            if let Some((find, replace)) = result {
                app.core.status = format!("Replace: '{find}' with '{replace}'");
            }
        }),
        "sort_asc" => {
            let wb = crate::ops::WorkbookState::default();
            dialogs::sort_dialog(&wb, |result| {
                if let Some((col, asc)) = result {
                    app.core.status = format!("Sort col {col} asc: {asc}");
                }
            });
        }
        "sort_desc" => {
            let wb = crate::ops::WorkbookState::default();
            dialogs::sort_dialog(&wb, |result| {
                if let Some((col, asc)) = result {
                    app.core.status = format!("Sort col {col} desc: {}", !asc);
                }
            });
        }
        "balance_books" => dialogs::balance_dialog(|result| {
            if let Some(col) = result {
                app.core.status = format!("Balance col: {col}");
            }
        }),
        "about" => {
            use std::sync::atomic::{AtomicUsize, Ordering};
            static ABOUT_COUNT: AtomicUsize = AtomicUsize::new(0);
            let n = ABOUT_COUNT.fetch_add(1, Ordering::SeqCst) + 1;
            let _ = std::fs::write("/tmp/corro_about_count.txt", format!("about called: {n}\n"));
            dialogs::show_about_dialog();
        }
        "help_keybinds" => {
            use std::sync::atomic::{AtomicUsize, Ordering};
            static KB_COUNT: AtomicUsize = AtomicUsize::new(0);
            let n = KB_COUNT.fetch_add(1, Ordering::SeqCst) + 1;
            let _ = std::fs::write("/tmp/corro_kb_count.txt", format!("help_keybinds called: {n}\n"));
            dialogs::show_keybinds_help();
        }
        "rename_sheet" => dialogs::find_dialog(|result| {
            if let Some(name) = result {
                app.core.status = format!("Rename sheet to: {name}");
            }
        }),
        "undo" => {
            app.core.status = "Undo not yet implemented".into();
            state.canvas.queue_redraw();
        }
        "redo" => {
            app.core.status = "Redo not yet implemented".into();
            state.canvas.queue_redraw();
        }
        "cut" => {
            app.core.status = "Cut not yet implemented".into();
        }
        "copy" => {
            app.core.status = "Copy not yet implemented".into();
        }
        "paste" => {
            app.core.status = "Paste not yet implemented".into();
        }
        "delete_cell" => {
            handle_delete(state);
        }
        "select_all" => {
            app.core.status = "Select All".into();
            app.core.anchor = None;
            state.canvas.queue_redraw();
        }
        "toggle_headers" => {
            app.core.status = "Toggle headers not yet implemented".into();
        }
        "toggle_margins" => {
            app.core.status = "Toggle margins not yet implemented".into();
        }
        "new_sheet" => {
            app.core.status = "New sheet not yet implemented".into();
        }
        "delete_sheet" => {
            app.core.status = "Delete sheet not yet implemented".into();
        }
        "export_tsv" => {
            if let Some(path) = dialogs::file_save_dialog() {
                app.core.status = format!("Exporting TSV to {}", path.display());
            }
        }
        "export_csv" => {
            if let Some(path) = dialogs::file_save_dialog() {
                app.core.status = format!("Exporting CSV to {}", path.display());
            }
        }
        "export_ods" => {
            if let Some(path) = dialogs::file_save_dialog() {
                app.core.status = format!("Exporting ODS to {}", path.display());
            }
        }
        "export_ascii" => {
            if let Some(path) = dialogs::file_save_dialog() {
                app.core.status = format!("Exporting ASCII to {}", path.display());
            }
        }
        "insert_rows" => {
            app.core.status = "Insert rows not yet implemented".into();
        }
        "insert_mitosis_row" => {
            app.core.status = "Insert mitosis row not yet implemented".into();
        }
        "insert_mitosis_col" => {
            app.core.status = "Insert mitosis col not yet implemented".into();
        }
        "insert_cols" => {
            app.core.status = "Insert cols not yet implemented".into();
        }
        "insert_special_chars" => {
            app.core.status = "Insert special chars not yet implemented".into();
        }
        "insert_date" => {
            app.core.status = "Insert date not yet implemented".into();
        }
        "insert_time" => {
            app.core.status = "Insert time not yet implemented".into();
        }
        "insert_hyperlink" => {
            app.core.status = "Insert hyperlink not yet implemented".into();
        }
        "format_apply_all" | "format_apply_full_column" | "format_apply_data"
        | "format_apply_special" | "format_apply_cell" | "format_apply_selection" => {
            app.core.status = format!("Format scope: {name}");
        }
        "format_decimal_generic" | "format_currency" | "format_rational"
        | "format_fixed_0" | "format_fixed_1" | "format_fixed_2" | "format_fixed_custom" => {
            app.core.status = format!("Format number: {name}");
        }
        "format_align_left" | "format_align_center" | "format_align_right"
        | "format_align_default" => {
            app.core.status = format!("Format align: {name}");
        }
        "format_reset" => {
            app.core.status = "Format reset".into();
        }
        _ => {
            app.core.status = format!("Menu action: {name}");
        }
    }
    update_formula_bar(state, state.last_row.get(), state.last_col.get());
}

// ---------------------------------------------------------------------------
// Formula entry change callback
// ---------------------------------------------------------------------------

fn on_formula_entry_changed(state: &GuiState) {
    if !state.editing.get() {
        if let Some(text) = state.formula_entry.get_text() {
            if !text.is_empty() {
                state.editing.set(true);
                state.mode.set(GuiMode::Normal);
                *state.edit_buf.borrow_mut() = text;
                state.canvas.queue_redraw();
            }
        }
        return;
    }
    if let Some(text) = state.formula_entry.get_text() {
        let current = state.edit_buf.borrow();
        // Safety check: only overwrite edit_buf from entry text when the
        // entry text is a forward or backward extension of the current
        // edit_buf.  This prevents data corruption when keystrokes from
        // the window-level handler (start_edit_with/handle_edit_key) race
        // with the entry's "changed" signal — a scenario where edit_buf
        // contains "4" (from window handler) but the entry text is "2"
        // (just arrived via entry default handler after grab_focus took
        // effect).  Without this guard, edit_buf would be overwritten to
        // "2", silently dropping the "4".
        if text.starts_with(&*current) || current.starts_with(&text) {
            drop(current);
            *state.edit_buf.borrow_mut() = text;
        }
        state.canvas.queue_redraw();
    }
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

pub fn run_gui(corro_app: &mut super::App) -> Result<(), Box<dyn std::error::Error>> {
    rustxwidgets::core::install_debug_crash_handlers();
    let rxapp = rustxwidgets::App::init()
        .map_err(|e| format!("GUI init failed: {e}"))?;

    let win = rxapp.new_window()?;
    win.set_title(&format!("corro {}", env!("CARGO_PKG_VERSION")));
    win.set_default_size(1200, 800);

    let mut vbox = rxapp.new_box(Orientation::Vertical, 0)?;

    // Fit column widths to rendered content
    corro_app.fit_main_columns_to_max_width();

    let hr = HEADER_ROWS;
    let lm = MARGIN_COLS;
    let cursor_row = hr;
    let cursor_col = lm;
    corro_app.core.cursor.row = cursor_row;
    corro_app.core.cursor.col = cursor_col;
    corro_app.core.anchor = Some(SheetCursor { row: hr, col: lm });

    let data_rows = 30usize;
    let data_cols = 12usize;

    // Formula bar
    let mut formula_bar = rxapp.new_box(Orientation::Horizontal, 2)?;
    let addr_label = rxapp.new_label("A1")?;
    let f_label = rxapp.new_label("  fx  ")?;
    let formula_entry = rxapp.new_entry()?;
    formula_entry.set_hexpand(true);
    formula_bar.append(&addr_label);
    formula_bar.append(&f_label);
    formula_bar.append(&formula_entry);
    formula_bar.set_child_hexpand(&formula_entry, true);

    // Canvas
    let canvas = rxapp.new_canvas()?;
    canvas.set_size_request(800, 600);
    // Ensure the canvas can receive keyboard focus (needed after commit_edit
    // to return focus — GtkDrawingArea does not accept focus by default).
    canvas.set_can_focus(true);

    // Status label
    let status_label = rxapp.new_label("Ready")?;

    let shared = Rc::new(GuiState {
        app: corro_app as *mut super::App,
        rxapp: rxapp.clone(),
        canvas: canvas.clone(),
        formula_entry: formula_entry.clone(),
        addr_label: addr_label.clone(),
        status_label: status_label.clone(),
        editing: Cell::new(false),
        edit_buf: RefCell::new(String::new()),
        mode: Cell::new(GuiMode::Normal),
        last_row: Cell::new(cursor_row),
        last_col: Cell::new(cursor_col),
        data_rows: Cell::new(data_rows),
        data_cols: Cell::new(data_cols),
        last_key: Cell::new(0),
        key_counter: Cell::new(0),
        entry_processed_key: Cell::new(false),
    });

    // Build menu
    let menubar = build_menu(&rxapp, &win, &shared)?;
    let menubar_cb = menubar.clone();
    vbox.append(&menubar);

    // Keyboard: canvas.on_key, win.on_event_key, etc.
    let shared_key = shared.clone();
    canvas.on_key(Box::new(move |keyval: u32| -> bool {
        handle_key(keyval, &shared_key)
    }));

    // Click
    let shared_click = shared.clone();
    canvas.on_click(Box::new(move |x: f64, y: f64| { handle_click(x, y, &shared_click); }));

    // Formula entry change
    let shared_entry = shared.clone();
    formula_entry.connect_changed(move || { on_formula_entry_changed(&shared_entry); })?;

    // Direct RETURN handling via connect_activate.  On GTK4 the entry's
    // internal CAPTURE-phase EventControllerKey consumes RETURN before
    // our BUBBLE-phase EventControllerKey (registered by on_key_raw)
    // ever fires.  The "activate" signal is the only path RETURN reaches
    // our code.  This handler provides a second, independent path that
    // does NOT go through the shared-callback RefCell in the GTK adapter's
    // on_key_raw, making it robust against any RefCell-borrow failures in
    // that path.  On NWG this is a no-op (connect_activate returns Ok(0)).
    let shared_act = shared.clone();
    formula_entry.connect_activate(Box::new(move |_entry: *mut std::os::raw::c_void| {
        handle_key(0xFF0D, &shared_act);
    }))?;

    // Window-level event interception: fallback for keys that escape the
    // focused widget.  The EventControllerKey is now stored in Window's
    // _controllers field so it stays alive.
    //
    // On GTK4 the entry's BUBBLE controller fires before the window's
    // BUBBLE controller.  For printable chars the entry marks
    // entry_processed_key; the window checks it here to avoid duplicating
    // the handle_key call (which would double characters in edit_buf).
    // On NWG the window's raw handler fires independently of the entry's
    // (only when no child has focus) so no doubling occurs.
    {
        let state_w = shared.clone();
        win.on_event_key(Box::new(move |keyval: u32, state: u32| -> i32 {
            let s: &GuiState = &*state_w;
            s.key_counter.set(s.key_counter.get() + 1);

            append_keylog(&format!("WINDOW_KEY: keyval={keyval} state={state} key_counter={}\n",
                s.key_counter.get()));

            // Propagate ALT key itself so GTK shows mnemonic hints
            if keyval == ALT_L || keyval == ALT_R {
                append_keylog("ALT alone, propagating\n");
                return 0;
            }

            let ch = char::from_u32(keyval).unwrap_or('\0').to_ascii_lowercase();
            let alt_held = (state & 0x8) != 0;
            let ctrl_held = (state & 0x4) != 0;

            // Ctrl+Q: quit
            if ctrl_held && ch == 'q' {
                eprintln!("DEBUG window key handler: Ctrl+Q detected");
                save_before_quit(s);
                return 1;
            }

            // ALT+letter: manually activate the submenu popover by calling
            // gtk_widget_activate on the corresponding PopoverMenuBarItem.
            // GTK's native mnemonic accelerator does NOT fire when a
            // CAPTURE-phase controller is present on the window, so we must
            // do it ourselves.
            if alt_held && (32..=126).contains(&keyval) {
                let ok = menubar_cb.activate_submenu_by_mnemonic(keyval);
                append_keylog(&format!("Alt+letter keyval={keyval} activate_submenu={ok}\n"));
                if ok {
                    return 1;
                }
                return 0;
            }

            // Check whether the entry already processed this printable
            // character.  On GTK4 BUBBLE phase, the entry's on_key_raw
            // fires first (line ~1405), calls handle_key, and sets
            // entry_processed_key=true.  Without this guard the window
            // handler would call handle_key again, doubling the character
            // in edit_buf ("4422" instead of "42").  This is the fix
            // described in Attempt 195 of the idea log.
            let nk = normalize(keyval);
            if (32..=126).contains(&nk) && s.entry_processed_key.get() {
                s.entry_processed_key.set(false);
                return 0;
            }

            // Safety net for RETURN: if editing is false but the formula entry
            // has text or edit_buf has content (e.g., from CAPTURE-phase key
            // processing on GTK where the entry widget never received the key),
            // set editing=true before delegating to handle_key so the edit is
            // committed instead of moving the cursor.
            if nk == RETURN && !s.editing.get() {
                let text = s.formula_entry.get_text().unwrap_or_default();
                if !text.is_empty() || !s.edit_buf.borrow().is_empty() {
                    s.editing.set(true);
                    if !text.is_empty() {
                        *s.edit_buf.borrow_mut() = text;
                    }
                }
            }
    let hk = handle_key(keyval, &state_w);
    append_keylog(&format!("handle_key={hk}\n"));
    if hk { 
        if normalize(keyval) == RETURN {
            append_keylog(&format!("window handler returned 1 for RETURN, editing={} text={:?} edit_buf={:?}\n",
                s.editing.get(),
                s.formula_entry.get_text().unwrap_or_default(),
                *s.edit_buf.borrow()));
        }
        1 
    } else { 0 }
        }));
    }

    // Register save-before-quit on close
    {
        let state_c = shared.clone();
        win.on_close(Box::new(move || { save_before_quit(&*state_c); }));
    }

    // Intercept Enter/Escape from formula entry
    {
        let shared_k = shared.clone();
        let shared_k_cnt = shared.clone();
        formula_entry.on_key_raw(Box::new(move |keyval: u32, state: u32| -> bool {
            shared_k_cnt.last_key.set(keyval);
            let k = normalize(keyval);
            match k {
                RETURN | ESCAPE | TAB | LEFT | RIGHT | UP | DOWN | HOME | END | PAGE_UP | PAGE_DOWN => {
                    handle_key(keyval, &shared_k);
                    true
                }
                _ if (32..=126).contains(&k) => {
                    // Process the key to update edit_buf (via handle_key →
                    // handle_edit_key or start_edit_with), then let the event
                    // propagate so the entry's default handler inserts the
                    // character and fires "changed" → on_formula_entry_changed.
                    // The starts_with guard in on_formula_entry_changed prevents
                    // the (already-correct) edit_buf from being overwritten by
                    // stale entry text in the mixed-flow scenario (first char
                    // via window handler, subsequent chars via entry handler
                    // after grab_focus).
                    //
                    // For keys with Ctrl (0x4) or Alt (0x8) modifiers, do NOT
                    // claim the key so the event bubbles to the window BUBBLE
                    // handler which processes Ctrl+Q.
                    if (state & (0x4 | 0x8)) != 0 {
                        return false;
                    }
                    handle_key(keyval, &shared_k);
                    shared_k.entry_processed_key.set(true);
                    false
                }
                _ => false,
            }
        }));
    }

// Assemble layout
    vbox.append(&formula_bar);
    vbox.append(&canvas);
    vbox.set_child_vexpand(&canvas, true);
    vbox.append(&status_label);

    // Register the draw callback BEFORE present() so the extensive event
    // pumping inside present() — which waits for the frame clock to fire
    // its first tick (up to 500+500 blocking iterations) — actually calls
    // our draw function.  When set_draw_callback was placed AFTER present(),
    // the initial 1050+ iterations did nothing because no draw function
    // was registered yet.  On virtual displays (WSLg, Xvfb) the next
    // frame clock tick may be delayed enough for a replayer to find the
    // window visible but blank ("WINDOW_DRAWN=false (blank window)").
    //
    // gtk_drawing_area_set_draw_func stores the callback in widget instance
    // data and works correctly whether the DrawingArea is realized or not;
    // the callback is invoked on the first frame clock tick after realization.
    let shared_draw = shared.clone();
    eprintln!("PHASE: before_set_draw_callback");
    let _ = std::fs::write("/tmp/gui_setup_phase1.txt", "before_set_draw_callback\n");
    canvas.set_draw_callback(Box::new(move |dc: &mut dyn DrawContext, w: i32, h: i32| {
        eprintln!("DRAW_CALLBACK called: w={} h={}", w, h);
        let _ = std::fs::write("/tmp/dim.txt", format!("{} {}\n", w, h));
        render_grid(dc, &shared_draw, w, h);
        // Test marker: 8x8 square of 0xFEEDBE at top-left, drawn AFTER render_grid
        // so it appears on top of the grid background and is visible in screenshots.
        dc.fill_rect(0.0, 0.0, 8.0, 8.0, 254.0/255.0, 237.0/255.0, 190.0/255.0, 1.0);
    }));
    eprintln!("PHASE: after_set_draw_callback");
    let _ = std::fs::write("/tmp/gui_setup_phase2.txt", "after_set_draw_callback\n");

    log_ui_action("gui_started", &format!("title={}", env!("CARGO_PKG_VERSION")));

    win.set_child_box(&vbox);
    // Grab focus on the formula entry BEFORE present() so the entry receives
    // initial keyboard focus when the window is mapped.  This ensures that
    // keystrokes from the external replayer (which detects the window during
    // present()'s event pumping) go through the entry's CAPTURE-phase
    // EventControllerKey, where printable characters are handled by
    // on_formula_entry_changed and RETURN is intercepted directly.  Without
    // this, the entry may not have focus during present(), causing keystrokes
    // to be processed by the window-level BUBBLE handler — which can race
    // with a later grab_focus() call and produce a mixed-flow data corruption
    // where edit_buf gets overwritten by incomplete entry text.
    formula_entry.grab_focus();
    eprintln!("PHASE: about_to_present");
    win.present();
    eprintln!("PHASE: after_present");
    let _ = std::fs::write("/tmp/gui_setup_phase3.txt", "after_present\n");

    // Queue an explicit redraw on the toplevel window: on GTK4 with the
    // Cairo GSK renderer, a canvas-only queue_draw may not cascade to
    // the toplevel's frame clock.  Marking the window dirty ensures the
    // frame clock is armed before the start_edit() canvas queue_redraw.
    win.queue_redraw();

    // Start editing at A1: grab_focus on the formula entry.
    // The draw callback was already registered before present(),
    // so the initial frame clock tick inside present() draws the grid.
    //
    // Keystrokes from the external replayer may arrive during present()'s
    // event pumping, before start_edit() is called.  Those keystrokes
    // processed via handle_key -> start_edit_with or handle_key ->
    // handle_edit_key establish editing state (editing=true, edit_buf
    // non-empty).  If we unconditionally call start_edit() here, it
    // clears edit_buf and destroys the in-flight edit — causing
    // OUTPUT_EXISTS=false when a subsequent RETURN commits an empty
    // buffer.  Guard the call: if editing is already in progress with
    // content, just grab focus and redraw without clearing.
    if shared.editing.get() && !shared.edit_buf.borrow().is_empty() {
        shared.formula_entry.grab_focus();
        shared.canvas.queue_redraw();
    } else {
        start_edit(&shared);
    }

    // Pump events after start_edit() to ensure the frame clock processes
    // the pending redraw (from queue_redraw/queue_draw) and the focus
    // change (from grab_focus) BEFORE the main loop starts.  This
    // prevents WINDOW_DRAWN=false on slow virtual displays where the
    // frame clock timer hasn't fired yet.

    // Pre-create the main loop so quit_main_loop finds a valid pointer
    // even if the user clicks Quit during the warm-up phase below.
    // The GtkApp::run() creates a fresh loop, so this pre-created loop
    // is only used if quit happens before run() starts.
    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    if let Some(loader) = rustxwidgets::backends::gtk::loader() {
        if let Some(loop_new) = loader.symbols.g_main_loop_new {
            let early_loop = unsafe { loop_new(std::ptr::null_mut(), 0) };
            if !early_loop.is_null() {
                *loader.main_loop.lock().unwrap() = early_loop as usize;
            }
        }
    }

    rxapp.pump_events(500);

    // Second safety net: queue another redraw and pump again.  Some
    // virtual displays (WSLg, Xvfb) need multiple pump cycles before
    // the frame clock tick is dispatched, even with all-blocking
    // pumping inside present().
    win.queue_redraw();
    canvas.queue_redraw();
    rxapp.pump_events(500);

    // Final fallback: force an immediate draw directly to the GdkSurface,
    // bypassing the frame clock entirely.  This is only needed when the
    // frame clock timer never fires (some WSLg/Xvfb configurations with
    // GSK_RENDERER=cairo).  The 1200x800 fallback dimensions match the
    // window default size, used when the surface reports zero size.
    canvas.force_draw(win.hwnd(), 1200, 800);

    // Move rxapp.run() earlier — before pump_events — so the main loop
    // pointer is available for quit_main_loop before any user interaction.
    rxapp.run()?;
    Ok(())
}
