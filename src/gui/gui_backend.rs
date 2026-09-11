use rswidgets::prelude::*;
use rswidgets::core::DrawContext;

use std::cell::{Cell, RefCell};
use std::collections::HashMap;
use std::rc::Rc;

use super::actions::run_prompt_action;
use super::actions::{dispatch_menu_action, MenuDispatch};
use super::extrapolate;

use crate::grid::{CellAddr, SheetCursor, HEADER_ROWS, MARGIN_COLS};
use crate::ops::{Op, WorkbookOp};
use crate::ui_core;

use super::compute::{self, CellDisplayStyle};
use super::dialogs;
use super::render::{self, CellSink};

use rswidgets::core::key::{normalize, RETURN, ESCAPE, BACKSPACE, DELETE, LEFT, UP, RIGHT, DOWN, TAB, HOME, END, PAGE_UP, PAGE_DOWN, F1, F2, ALT_L, ALT_R};

const KEYLOG_PATH: &str = "/tmp/corro_keylog.txt";
/// Modifier bit for Shift in key-event state masks. Shared by GDK
/// (GDK_SHIFT_MASK) and the nwg adapter (Win32 shift bit).
const MOD_SHIFT: u32 = 0x1;

fn key_name(keyval: u32) -> String {
    if keyval == 0 { return "MENU".into(); }
    #[allow(unreachable_patterns)] // ALT_L/ALT_R are the same value on Windows
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
        ALT_L | ALT_R => "ALT".into(),
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
    rxapp: rswidgets::App,
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
    last_alt_keyval: Cell<u32>,
    // Tracks whether the entry widget's key handler has ever fired. Gates the
    // entry_processed_key protocol below: that flag is only meaningful when
    // both the window and the entry can observe the same event (double-fire
    // setups). On streams where the entry never fires, setting the flag would
    // poison the next keypress (it would wrongly skip handle_key).
    entry_seen: Cell<bool>,
    // Shared-dispatch state for menu actions: clipboard for cut/copy/paste
    // and the pending format scope. The pancurses backend passes these as
    // dispatch arguments; the GUI backend keeps them here so every menu item
    // behaves identically across backends instead of drifting into stubs.
    clipboard: RefCell<String>,
    pending_scope: Cell<u8>,
    // Scrollbar sync: native scrollbars around the sheet (thumb tracks the
    // cursor; dragging/clicking moves the cursor, so the selection is always
    // visible). Guard against reentrancy between programmatic sets and the
    // value-changed notification.
    scrolled: rswidgets::common::ScrolledWindow,
    syncing_scroll: Cell<bool>,
    // Press/release dedup trackers. GTK4-only (feature = "gtk4"): only there
    // can release events arrive as same-keyval callbacks. Everywhere else
    // (GTK3 presses-only, nwg WM_KEYDOWN-only, wasm keydown-only) every key
    // event is a genuine press and deduping would swallow genuine repeats.
    #[cfg(feature = "gtk4")]
    last_dedup_key: Cell<u32>,
    #[cfg(feature = "gtk4")]
    dedup_count: Cell<u32>,
    // Prevents the RETURN safety net (line ~1375) from re-entering edit mode
    // on the release event of a RETURN press that already committed an edit.
    // Set after handle_key processes RETURN; checked by the safety net to
    // distinguish between a legitimate RETURN press (editing=false, text
    // non-empty during present() race) and a release event following a
    // normal RETURN press that committed an edit and re-displayed the
    // new cell's value in the formula entry.
    return_pressed: Cell<bool>,
    // Press/release dedup bookkeeping. GTK4-only (see entry_seen above):
    // EventControllerKey::key-pressed can fire for both GDK_KEY_PRESS and
    // GDK_KEY_RELEASE on some versions/display servers.  When the same
    // canonical keyval arrives twice consecutively, the second event is
    // a release and should be skipped.  Set at each return point where
    // a key was actually processed; cleared on skip so the next different
    // key is not affected.
    #[cfg(feature = "gtk4")]
    last_keyval_dedup: Cell<u32>,
}

impl GuiState {
    /// Borrow the host [`App`](super::App) mutably. Centralised here so the
    /// raw pointer is dereferenced in exactly one place and call sites stay
    /// `unsafe`-free.
    ///
    /// # Contract (not machine-checked)
    /// The `'a` lifetime is intentionally free: the pointer is trusted, so
    /// the compiler cannot prevent two live borrows — only discipline can.
    /// Callers must observe two rules:
    /// 1. **LIFO nesting only.** A borrow may overlap an outer borrow only
    ///    while the outer borrow is untouched (as in `move_cursor` calling
    ///    `update_state_cursor` and never touching its own borrow after).
    ///    Never use an outer borrow after an inner one was taken.
    /// 2. **No cross-frame borrows.** Never capture a borrow in a `'static`
    ///    dialog/event callback. Capture an `Rc<GuiState>` clone instead
    ///    and borrow inside the callback at fire time, so each borrow's
    ///    dynamic extent lies within a single sequential dispatch.
    /// The remaining premise — the pointer itself cannot dangle — holds
    /// because the caller keeps `corro_app` alive across the blocking
    /// [`run_gui`](run_gui) event loop.
    #[allow(clippy::needless_lifetimes)]
    fn app_mut<'a>(&self) -> &'a mut super::App {
        unsafe { &mut *self.app }
    }

    /// Shared-borrow variant of [`app_mut`](GuiState::app_mut) for read-only
    /// sites. Same contract; shared borrows compose freely with the other
    /// `&self` uses at those sites.
    fn app_ref(&self) -> &super::App {
        unsafe { &*self.app }
    }
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

    // Selection rectangle (anchor..cursor, rows AND columns), computed once.
    // Plain navigation collapses the anchor, so no band is painted then —
    // only explicit selections (Shift+arrows, select-all) highlight.
    let sel_rect: Option<(usize, usize, usize, usize)> = selection_anchor.map(|(ar, ac)| {
        let (r1, r2) = if ar <= cursor_row { (ar, cursor_row) } else { (cursor_row, ar) };
        let (c1, c2) = if ac <= cursor_col { (ac, cursor_col) } else { (cursor_col, ac) };
        (r1, r2, c1, c2)
    });
    for (ri, &logical_row) in display_rows.iter().enumerate().take(MAX_RENDER_ROWS) {
        let ry = HEADER_H + ri as f64 * ROW_H;

        for (ci, &c) in col_ixs.iter().enumerate().take(MAX_RENDER_COLS) {
            let cw = *col_widths.get(&c).unwrap_or(&8) as f64 * CHAR_W;
            let cx = ROW_LABEL_W + col_ixs.iter().take(ci).map(|&pc| *col_widths.get(&pc).unwrap_or(&8) as f64 * CHAR_W).sum::<f64>();

            let key = (ri as u32, c as u32);
            let raw_text = cells.get(&key).map(|s| s.as_str()).unwrap_or("");
            let style_key = (ri as u32, c as u32);
            let style = styles.get(&style_key).copied().unwrap_or(CellDisplayStyle::Default);
            let is_current = logical_row == cursor_row && c == cursor_col;
            let in_selection = sel_rect.map_or(false, |(r1, r2, c1, c2)| {
                logical_row >= r1 && logical_row <= r2 && c >= c1 && c <= c2
            });

            let bg = if is_current {
                if is_editing { (1.0, 1.0, 0.8, 1.0) } else { (0.8, 0.9, 1.0, 1.0) }
            } else if in_selection {
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

/// How many columns are needed so the fitted column widths cover `avail_px`
/// pixels (matching render_grid's width accumulation). The GUI canvas can be
/// any size, so the count is derived from the live canvas width every frame
/// (like ratatui sizes from the terminal each frame) instead of a hardcoded
/// constant — otherwise the sheet stops early and leaves a huge blank area.
fn cols_to_fill_px(app: &super::App, cursor: SheetCursor, avail_px: i32) -> usize {
    let sheet = app.core.workbook.active_sheet();
    let mut dim = 1usize;
    loop {
        let (cols, _) = ui_core::visible_col_indices(sheet, cursor, dim, 0);
        let used: f64 = cols
            .iter()
            .map(|&c| sheet_rec_col_width(sheet, c) as f64 * CHAR_W)
            .sum();
        if used >= avail_px as f64 || dim >= 2048 || cols.len() < dim {
            return dim.max(1);
        }
        dim += 8;
    }
}

/// How many rows are needed to cover a canvas `h` pixels tall. Reserves the
/// 20px in-canvas status strip (drawn over the bottom) so the last row —
/// often the cursor — stays fully visible instead of sliding underneath it.
fn rows_to_fill_px(h: i32) -> usize {
    (((h as f64 - HEADER_H - 20.0) / ROW_H + 1.0).max(1.0)) as usize
}

fn render_grid(dc: &mut dyn DrawContext, state: &GuiState, w: i32, h: i32) {
    dc.clear(0.94, 0.94, 0.94, 1.0);
    dc.clip(0.0, 0.0, w as f64, h as f64);

    let app = state.app_ref();
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

    // Selection rectangle (anchor..cursor, rows AND columns). None while
    // navigating plainly — only explicit selections highlight.

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

    // Diagnostic overlay (cell/sink/cursor/key state) painted over the grid.
    // Opt-in via CORRO_DEBUG_OVERLAY (any value): off by default so normal
    // runs render a clean sheet. Previously this always drew, obscuring cells
    // and breaking pixel-level assertions about grid content.
    if std::env::var_os("CORRO_DEBUG_OVERLAY").is_some() {
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
}

fn sheet_rec_col_width(sheet: &crate::ops::SheetState, col: usize) -> usize {
    sheet.grid.col_width(col).max(1)
}

// ---------------------------------------------------------------------------
// Keyboard handling
// ---------------------------------------------------------------------------

/// Handle a key while the interactive extrapolate modal is active. Mirrors the
/// ratatui reference `Mode::Extrapolate`: arrows/navigation extend the selection
/// (the anchor stays put), Enter commits, Esc cancels. The modal state lives in
/// `extrapolate.rs` and is shared by all GUI backends.
fn handle_extrapolate_key(key: u32, state: &GuiState) -> bool {
    match key {
        ESCAPE => {
            log_key_action(state.last_key.get(), "extrapolate_cancel", "");
            super::extrapolate::cancel(state.app_mut());
            update_state_cursor(state, state.last_row.get(), state.last_col.get());
            state.canvas.queue_redraw();
            true
        }
        RETURN => {
            log_key_action(state.last_key.get(), "extrapolate_commit", &format!("cell={}", format_cell(state)));
            super::extrapolate::commit(state.app_mut());
            update_state_cursor(state, state.last_row.get(), state.last_col.get());
            state.canvas.queue_redraw();
            true
        }
        LEFT | RIGHT | UP | DOWN => {
            let (dr, dc) = match key {
                LEFT => (0, -1),
                RIGHT => (0, 1),
                UP => (-1, 0),
                _ => (1, 0),
            };
            // Preserve the anchor so the selection extends (move_cursor clears it).
            let anchor = state.app_ref().core.anchor;
            if anchor.is_none() {
                state.app_mut().core.anchor = Some(state.app_ref().core.cursor);
            }
            move_cursor(state, dr, dc);
            if let Some(a) = anchor {
                state.app_mut().core.anchor = Some(a);
            }
            super::extrapolate::refresh_preview(state.app_mut());
            state.canvas.queue_redraw();
            true
        }
        _ => false,
    }
}

fn handle_key(keyval: u32, state_rc: &Rc<GuiState>, mods: u32) -> bool {
    let state: &GuiState = &**state_rc;
    state.last_key.set(keyval);
    let app = state.app_mut();
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

    // Interactive extrapolate modal (mirrors ratatui's Mode::Extrapolate):
    // arrows extend the selection (anchor stays put), Enter commits, Esc
    // cancels. Intercepted here before normal cursor/enter handling.
    if state.app_ref().extrapolate.is_some() {
        return handle_extrapolate_key(key, state);
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
            // Collapse any selection (plain navigation state).
            app.core.anchor = None;
            state.canvas.queue_redraw();
            true
        }
        LEFT => {
            log_key_action(keyval, "move_cursor_left", &format!("cell={}", format_cell(state)));
            if mods & MOD_SHIFT != 0 {
                extend_selection(state, 0, -1);
            } else {
                move_cursor(state, 0, -1);
            }
            true
        }
        RIGHT => {
            log_key_action(keyval, "move_cursor_right", &format!("cell={}", format_cell(state)));
            if mods & MOD_SHIFT != 0 {
                extend_selection(state, 0, 1);
            } else {
                move_cursor(state, 0, 1);
            }
            true
        }
        UP => {
            log_key_action(keyval, "move_cursor_up", &format!("cell={}", format_cell(state)));
            // Plain Up walks into the header band (matching ratatui);
            // move_cursor itself floors at row 0.
            if state.last_row.get() > 0 {
                if mods & MOD_SHIFT != 0 {
                    extend_selection(state, -1, 0);
                } else {
                    move_cursor(state, -1, 0);
                }
            }
            true
        }
        DOWN => {
            log_key_action(keyval, "move_cursor_down", &format!("cell={}", format_cell(state)));
            if mods & MOD_SHIFT != 0 {
                extend_selection(state, 1, 0);
            } else {
                move_cursor(state, 1, 0);
            }
            true
        }
        HOME => {
            log_key_action(keyval, "move_cursor_home", &format!("cell={}", format_cell(state)));
            // Jump to the leftmost non-blank cell in the row (matching ratatui).
            if let Some((leftmost, _)) = row_nonblank_extremes(state, state.last_row.get()) {
                state.last_col.set(leftmost);
                update_state_cursor(state, state.last_row.get(), leftmost);
            }
            true
        }
        END => {
            log_key_action(keyval, "move_cursor_end", &format!("cell={}", format_cell(state)));
            // Jump to the rightmost non-blank cell in the row (matching ratatui).
            if let Some((_, rightmost)) = row_nonblank_extremes(state, state.last_row.get()) {
                state.last_col.set(rightmost);
                update_state_cursor(state, state.last_row.get(), rightmost);
            }
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
            // Prime the press/release tracker so a following release event is
            // skipped by handle_edit_key. GTK4-only: only there can releases
            // arrive as same-keyval callbacks; elsewhere priming would cause
            // handle_edit_key to skip the next identical char.
            #[cfg(feature = "gtk4")]
            {
                let dk = ch.to_ascii_lowercase() as u32;
                state.last_dedup_key.set(dk);
                state.dedup_count.set(1);
            }
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
            sync_entry_to_buf(state);
            state.canvas.queue_redraw();
            true
        }
        DELETE => {
            log_key_action(key, "edit_clear", &format!("cell={} mode=edit", format_cell(state)));
            state.edit_buf.borrow_mut().clear();
            sync_entry_to_buf(state);
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
            // Route through move_cursor (not a direct +1) so the grid grows
            // at the boundary exactly like plain Right does — matching
            // ratatui, where commit-then-Right grows on the just-committed
            // content (e.g. A,Right,Right reaches C1, not the margin).
            move_cursor(state, 0, 1);
            true
        }
        UP => {
            log_key_action(key, "commit_edit_up", &format!("cell={} mode=edit", format_cell(state)));
            commit_edit(state);
            // Commit-then-Up walks the header band too (matching ratatui).
            if state.last_row.get() > 0 {
                state.last_row.set(state.last_row.get() - 1);
                update_state_cursor(state, state.last_row.get(), state.last_col.get());
            }
            true
        }
        DOWN => {
            log_key_action(key, "commit_edit_down", &format!("cell={} mode=edit", format_cell(state)));
            commit_edit(state);
            // Same growth routing as RIGHT above (ratatui grows here too).
            move_cursor(state, 1, 0);
            true
        }
        _ if (32..=126).contains(&key) => {
            // Press/release dedup. GTK4-only (see GuiState): elsewhere every
            // key event is a genuine press (releases are filtered at the
            // source or never hooked), so deduping would swallow genuine
            // repeats ("HELLO" -> "HELO").
            #[cfg(feature = "gtk4")]
            {
                let dedup_key = char::from_u32(key).map(|c| c.to_ascii_lowercase() as u32).unwrap_or(key);
                if dedup_key == state.last_dedup_key.get() {
                    let cnt = state.dedup_count.get() + 1;
                    state.dedup_count.set(cnt);
                    // Skip every even occurrence (the release event)
                    if cnt % 2 == 0 {
                        return true;
                    }
                } else {
                    state.last_dedup_key.set(dedup_key);
                    state.dedup_count.set(1);
                }
            }
            let ch = char::from_u32(key).unwrap_or('?');
            log_key_action(key, "edit_insert", &format!("char={ch} cell={} mode=edit", format_cell(state)));
            state.edit_buf.borrow_mut().push(ch);
            sync_entry_to_buf(state);
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
    state.edit_buf.borrow_mut().push(ch);
    // Keep the widget identical to edit_buf (see sync_entry_to_buf): the
    // widget text is what the user sees, edit_buf is what gets committed.
    sync_entry_to_buf(state);
    if !already_editing {
        state.formula_entry.grab_focus();
    }
    state.canvas.queue_redraw();
}

/// Keep the formula entry widget text identical to edit_buf (the commit
/// source of truth). Without this the two diverge: the native widget inserts
/// typed chars itself on top of handle_key's push (doubling, "AA"), and it
/// never sees Backspace/Delete (handle_key consumes those), leaving stale
/// text on screen. Syncing here covers every key path (window + entry, all
/// GUI backends); on_formula_entry_changed accepts identical text as a no-op.
fn sync_entry_to_buf(state: &GuiState) {
    let buf = state.edit_buf.borrow().clone();
    if state.formula_entry.get_text().as_deref() != Some(buf.as_str()) {
        state.formula_entry.set_text(&buf);
    }
}

fn commit_edit(state: &GuiState) {
    state.editing.set(false);
    state.mode.set(GuiMode::Normal);
    let val = state.edit_buf.borrow().clone();
    if !val.is_empty() {
        let app = state.app_mut();
        let row = state.last_row.get();
        let col = state.last_col.get();
        // Resolve the TRUE cell address (header/margin/footer included),
        // matching ratatui's commit_edit_buffer which commits to the edit
        // target address. Building CellAddr::Main unconditionally misroutes
        // header/margin/footer edits into main cells.
        let addr = crate::addr::sheet_cursor_to_addr(
            crate::addr::LogicalRow(row),
            crate::addr::GlobalCol(col),
            crate::addr::MainRows(app.core.workbook.active_sheet().grid.main_rows()),
            crate::addr::MainCols(app.core.workbook.active_sheet().grid.main_cols()),
        );
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
    let app = state.app_mut();
    let row = state.last_row.get();
    let col = state.last_col.get();
    // Resolve the TRUE cell address (header/margin/footer included),
    // matching ratatui: clearing a margin/header cell must clear that cell,
    // not the clamped main cell.
    let addr_of = |app: &mut super::App, r: usize, c: usize| crate::addr::sheet_cursor_to_addr(
        crate::addr::LogicalRow(r),
        crate::addr::GlobalCol(c),
        crate::addr::MainRows(app.core.workbook.active_sheet().grid.main_rows()),
        crate::addr::MainCols(app.core.workbook.active_sheet().grid.main_cols()),
    );
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
                    let addr = addr_of(app, r, c);
                    app.core.workbook.active_sheet_mut().grid.set(&addr, String::new());
                }
            }
            let sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
            let addr = addr_of(app, row, col);
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
            // Clearing consumes the selection.
            app.core.anchor = None;
            state.canvas.queue_redraw();
            return;
        }
    }
    let addr = addr_of(app, row, col);
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
    // Clearing consumes the selection (nothing remains selected).
    state.app_mut().core.anchor = None;
    state.canvas.queue_redraw();
}

fn recompute_viewport(state: &GuiState) {
    let app = state.app_ref();
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

/// Leftmost/rightmost non-blank global column in `row` (global row space),
/// matching ratatui's row_nonblank_horizontal_extremes. None when the row has
/// no non-blank cells.
fn row_nonblank_extremes(state: &GuiState, row: usize) -> Option<(usize, usize)> {
    let app = state.app_ref();
    let grid = &app.core.workbook.active_sheet().grid;
    let main_cols = grid.main_cols();
    let mut first: Option<usize> = None;
    let mut last: Option<usize> = None;
    for (addr, val) in grid.iter_nonempty() {
        if val.trim().is_empty() {
            continue;
        }
        let (r, c) = match addr {
            CellAddr::Main { row: r, col: c } => (HEADER_ROWS + r as usize, MARGIN_COLS + c as usize),
            CellAddr::Left { row: r, .. } => (HEADER_ROWS + r as usize, 0),
            CellAddr::Right { row: r, .. } => (HEADER_ROWS + r as usize, MARGIN_COLS + main_cols),
            CellAddr::Header { row: r, .. } | CellAddr::Footer { row: r, .. } => (r as usize, 0),
        };
        if r == row {
            if first.is_none() {
                first = Some(c);
            }
            last = Some(c);
        }
    }
    match (first, last) {
        (Some(a), Some(b)) => Some((a, b)),
        _ => None,
    }
}

fn move_cursor(state: &GuiState, dr: isize, dc: isize) {
    let hr = HEADER_ROWS;
    let lm = MARGIN_COLS;
    let (mut row, mut col) = (state.last_row.get(), state.last_col.get());
    if dc > 0 {
        // Grow the grid when stepping right off the last main column with
        // few trailing blanks (matching ratatui's move_cursor_one_col_horizontal).
        // Clamping here instead strands the cursor (Right does nothing) and
        // even walks it backwards when the clamp bound sits below the cursor.
        let mc = state.app_ref().core.workbook.active_sheet().grid.main_cols();
        if col == lm + mc.saturating_sub(1) {
            let app = state.app_mut();
            let sheet = app.core.workbook.active_sheet_mut();
            if compute::trailing_blank_main_cols(&sheet.grid) < ui_core::NAV_BLANK_COLS {
                sheet.grid.grow_main_col_at_right();
            }
        }
        col = col.saturating_add(1);
    } else if dc < 0 {
        // Plain Left walks into the left margin (floor 0), matching
        // ratatui's move_cursor_one_col_horizontal (Shift+Left is the one
        // that stays in main — see extend_selection).
        col = col.saturating_sub(1);
    } else if dr > 0 {
        // Grow the grid when stepping down off the last main row
        // (matching ratatui's move_cursor_one_row_vertical).
        let mr = state.app_ref().core.workbook.active_sheet().grid.main_rows();
        if row == hr + mr.saturating_sub(1) {
            let app = state.app_mut();
            let sheet = app.core.workbook.active_sheet_mut();
            if compute::trailing_blank_main_rows(&sheet.grid) < ui_core::NAV_BLANK_ROWS {
                sheet.grid.grow_main_row_at_bottom();
            }
        }
        row = row.saturating_add(1);
    } else if dr < 0 {
        // Plain Up walks into the header band (floor 0), matching ratatui's
        // move_cursor_one_row_vertical (Shift+Up is the one that stays in
        // main — see extend_selection).
        row = row.saturating_sub(1);
    }
    // Clamp into the grid extent and ensure it (matching ratatui).
    {
        let app = state.app_mut();
        let grid = &mut app.core.workbook.active_sheet_mut().grid;
        let mut cursor = SheetCursor { row, col };
        cursor.clamp(grid);
        grid.ensure_extent_for_cursor(cursor.row, cursor.col);
        row = cursor.row;
        col = cursor.col;
    }
    update_state_cursor(state, row, col);
}

fn update_state_cursor(state: &GuiState, row: usize, col: usize) {
    state.last_row.set(row);
    state.last_col.set(col);
    let app = state.app_mut();
    // Plain navigation collapses any selection (matching ratatui, where the
    // anchor exists only transiently). Without this the anchor sticks and
    // every move paints a phantom band over all rows above the cursor.
    // (extend_selection saves and restores the anchor around moves.)
    app.core.anchor = None;
    app.core.cursor.row = row;
    app.core.cursor.col = col;
    update_formula_bar(state, row, col);
    state.canvas.queue_redraw();
}

/// Extend the selection (Shift+arrows, matching ratatui): anchor at the
/// pre-move cursor if unset, then move (with grid growth). Refuses to leave
/// the main area, like the reference. move_cursor collapses the anchor, so
/// any previous selection is saved and restored around the move.
fn extend_selection(state: &GuiState, dr: isize, dc: isize) {
    let (row, col) = (state.last_row.get(), state.last_col.get());
    if dc < 0 && col <= MARGIN_COLS {
        return;
    }
    if dr < 0 && row <= HEADER_ROWS {
        return;
    }
    if dr == 0 && dc == 0 {
        return;
    }
    let prev = state.app_ref().core.anchor;
    move_cursor(state, dr, dc);
    state.app_mut().core.anchor = prev.or(Some(SheetCursor { row, col }));
}

/// Scrollbar domain (upper bounds) for (rows, cols): content plus the
/// visible page plus footer padding, so upper > page always holds (a
/// degenerate upper <= page disables the bar). The native bars need a finite
/// domain; the sheet's margin/header bands are astronomically large.
fn scroll_domain(state: &GuiState) -> (usize, usize) {
    let app = state.app_ref();
    let grid = &app.core.workbook.active_sheet().grid;
    let ru = grid.main_rows() + state.data_rows.get().max(1) + 10;
    let cu = grid.main_cols() + state.data_cols.get().max(1) + 10;
    (ru.max(2), cu.max(2))
}

/// Push cursor position and domain into the native scrollbars (thumb tracks
/// the selection). Guarded against reentrancy with the value-changed
/// notification below. Domain and page satisfy upper > page so the bars
/// stay live: upper covers content plus a viewport plus footer padding.
fn sync_scrollbars(state: &GuiState) {
    if state.syncing_scroll.get() {
        return;
    }
    state.syncing_scroll.set(true);
    let (ru, cu) = scroll_domain(state);
    let vv = state.last_row.get().saturating_sub(HEADER_ROWS).min(ru.saturating_sub(1));
    let hv = state.last_col.get().saturating_sub(MARGIN_COLS).min(cu.saturating_sub(1));
    state.scrolled.scroll_to(
        hv as f64, cu as f64, state.data_cols.get().max(1) as f64,
        vv as f64, ru as f64, state.data_rows.get().max(1) as f64,
    );
    state.syncing_scroll.set(false);
}

/// Scrollbar interaction moves the cursor (selection) to the thumb-indicated
/// cell, clamped into the domain (no grid growth from scrollbars). The
/// per-frame viewport recompute then keeps it visible. Plain navigation, so
/// any selection collapses via update_state_cursor.
fn scroll_to_cursor(state: &GuiState, vertical: bool, value: f64) {
    if state.syncing_scroll.get() {
        return;
    }
    let (ru, cu) = scroll_domain(state);
    if vertical {
        let row = (HEADER_ROWS as f64 + value).max(HEADER_ROWS as f64) as usize;
        let row = row.min(HEADER_ROWS + ru.saturating_sub(1));
        update_state_cursor(state, row, state.last_col.get());
    } else {
        let col = (MARGIN_COLS as f64 + value).max(MARGIN_COLS as f64) as usize;
        let col = col.min(MARGIN_COLS + cu.saturating_sub(1));
        update_state_cursor(state, state.last_row.get(), col);
    }
}

fn update_formula_bar(state: &GuiState, row: usize, col: usize) {
    let app = state.app_ref();
    let addr_str = crate::addr::sheet_cursor_to_addr(
        crate::addr::LogicalRow(row),
        crate::addr::GlobalCol(col),
        crate::addr::MainRows(app.core.workbook.active_sheet().grid.main_rows()),
        crate::addr::MainCols(app.core.workbook.active_sheet().grid.main_cols()),
    );
    state.addr_label.set_text(&addr_str.to_string());
    // Look up the entry value at the TRUE address (a header/margin cursor
    // shows that cell's value, not the clamped main cell's).
    let addr = crate::addr::sheet_cursor_to_addr(
        crate::addr::LogicalRow(row),
        crate::addr::GlobalCol(col),
        crate::addr::MainRows(app.core.workbook.active_sheet().grid.main_rows()),
        crate::addr::MainCols(app.core.workbook.active_sheet().grid.main_cols()),
    );
    let val = app.core.workbook.active_sheet().grid.get(&addr).unwrap_or_default();
    // While an edit is in progress the entry widget belongs to the edit
    // buffer (e.g. an Insert Date/Time preset), not the grid cell: overwriting
    // it here would wipe the preset (and the entry's change handler could
    // then eat edit_buf too). Address/status labels always update.
    if !state.editing.get() {
        state.formula_entry.set_text(&val);
    }
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
    let app = state.app_mut();
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
                // Plain click collapses any selection (fresh single-cell focus).
                app.core.anchor = None;
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

fn build_menu(rxapp: &rswidgets::App, win: &Window, state: &Rc<GuiState>) -> Result<MenuBar, Box<dyn std::error::Error>> {
    use crate::gui::menu;

    let action_group = rxapp.ensure_action_group()?;

    // Build the full menu tree from the shared definition (menu::menu_bar),
    // the same tree the ratatui reference and the pancurses backend build
    // from, so the GTK menus can never drift from them.
    let bar = menu::menu_bar();
    let mut menubar_model = rxapp.new_menu()?;
    for root in &bar {
        let sub = menu::build_common_menu(rxapp, root.submenu.as_deref().unwrap_or(&[]), "app")?;
        // Top-level items carry no shortcut: first-character mnemonic.
        menubar_model.append_submenu(&menu::mnemonic_label(root.label, root.shortcut), &sub);
    }

    // Register action callbacks with state access (walk the whole tree).
    fn register_actions(
        rxapp: &rswidgets::App,
        items: &[menu::MenuAction],
        s: &Rc<GuiState>,
    ) -> Result<(), Box<dyn std::error::Error>> {
        for item in items {
            if let Some(sub) = item.submenu.as_deref() {
                register_actions(rxapp, sub, s)?;
            } else {
                let name = menu::action_kind_to_name(item.action);
                let name_owned = name.to_string();
                let state_cb = s.clone();
                menu::register_action(rxapp, name, move || handle_menu_action(&name_owned, &state_cb))?;
            }
        }
        Ok(())
    }
    let s = state.clone();
    for root in &bar {
        register_actions(rxapp, root.submenu.as_deref().unwrap_or(&[]), &s)?;
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

/// Start editing with a full preset string (used by menu actions such as
/// Insert > Date/Time, which arrive via MenuDispatch::Edit). Mirrors
/// start_edit_with but takes &str so multi-char values need no loop.
fn start_edit_with_text(state: &GuiState, text: &str) {
    state.editing.set(true);
    *state.edit_buf.borrow_mut() = text.to_string();
    // Keep the widget identical to edit_buf (see sync_entry_to_buf).
    sync_entry_to_buf(state);
    state.formula_entry.grab_focus();
    state.canvas.queue_redraw();
}

/// Refresh viewport, formula bar, and canvas after a menu action mutated the
/// workbook or cursor outside the normal key path (dialog callbacks, shared
/// dispatch). Without this the grid shows stale cells after e.g. Find moves
/// the cursor or Replace edits cells.
fn refresh_after_dialog(state: &Rc<GuiState>) {
    recompute_viewport(state);
    update_formula_bar(state, state.last_row.get(), state.last_col.get());
    state.canvas.queue_redraw();
}

/// Route a menu action through the shared `actions::dispatch_menu_action`
/// implementation (the same code the pancurses backend uses), then apply the
/// result to GUI state. This is what keeps GUI menu behavior identical to the
/// other backends instead of drifting into per-backend stubs: any menu item
/// handled here behaves exactly as it does under pancurses/ratatui.
fn delegate_shared_action(name: &str, state: &Rc<GuiState>) {
    let mut scope = state.pending_scope.get();
    let mut cb = state.clipboard.borrow().clone();
    let result = dispatch_menu_action(state.app_mut(), name, &mut scope, &mut cb);
    state.pending_scope.set(scope);
    *state.clipboard.borrow_mut() = cb;
    match result {
        MenuDispatch::Status(s) => {
            // Empty means "keep current status" (copy/paste/single-sheet nav).
            if !s.is_empty() {
                state.app_mut().core.status = s;
            }
        }
        MenuDispatch::Edit { value } => {
            // Insert > Date/Time: preset the edit buffer; the user commits
            // with Enter exactly like the other backends.
            start_edit_with_text(state, &value);
        }
        MenuDispatch::Prompt(_label, action) => {
            // Only "save" (no path yet) reaches here; open a save dialog.
            if action == "save_as" {
                if let Some(path) = dialogs::file_save_dialog() {
                    run_prompt_action(state.app_mut(), action, &path.display().to_string());
                }
            }
        }
        MenuDispatch::About { status } => {
            state.app_mut().core.status = status;
            dialogs::show_about_dialog();
        }
        MenuDispatch::HelpFull { .. } => {
            dialogs::show_keybinds_help();
        }
        MenuDispatch::HelpKeybinds { .. } => {
            dialogs::show_keybinds_help();
        }
    }
    // Shared ops may move the cursor (select_all, mitosis, go_to); sync the
    // widget cursor cells before redrawing so the formula bar follows.
    let (cr, cc) = {
        let app = state.app_ref();
        (app.core.cursor.row, app.core.cursor.col)
    };
    state.last_row.set(cr);
    state.last_col.set(cc);
    refresh_after_dialog(state);
}

fn handle_menu_action(name: &str, state: &Rc<GuiState>) {
    let app = state.app_mut();
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
            } else if let Some(path) = dialogs::file_save_dialog() {
                // No path yet: fall back to Save As (same as pancurses).
                run_prompt_action(app, "save_as", &path.display().to_string());
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
        "quit" => {
            eprintln!("DEBUG handle_menu_action: quit activated");
            save_before_quit(state);
        }
        "find" | "replace" | "balance_books" | "rename_sheet" => {
            // Shared search/mutate logic (same as pancurses/ratatui); the
            // dialog only supplies the input. Previously these set a
            // status string without doing anything.
            let st = state.clone();
            let action = name.to_string();
            dialogs::find_dialog(move |result| {
                if let Some(text) = result {
                    run_prompt_action(st.app_mut(), &action, &text);
                    refresh_after_dialog(&st);
                }
            });
        }
        "sort_asc" => {
            let wb = crate::ops::WorkbookState::default();
            let st = state.clone();
            dialogs::sort_dialog(&wb, move |result| {
                if let Some((col, asc)) = result {
                    st.app_mut().core.status = format!("Sort col {col} asc: {asc}");
                }
            });
        }
        "sort_desc" => {
            let wb = crate::ops::WorkbookState::default();
            let st = state.clone();
            dialogs::sort_dialog(&wb, move |result| {
                if let Some((col, asc)) = result {
                    st.app_mut().core.status = format!("Sort col {col} desc: {}", !asc);
                }
            });
        }
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
        "undo" | "redo" | "cut" | "copy" | "paste" => {
            // Shared clipboard/history logic (same as pancurses/ratatui).
            delegate_shared_action(name, state);
        }
        "delete_cell" => {
            handle_delete(state);
        }
        "select_all" | "toggle_headers" | "toggle_margins" | "new_sheet" => {
            // Shared selection/chrome/sheet logic (same as pancurses/ratatui).
            delegate_shared_action(name, state);
        }
        "delete_sheet" => {
            // Shared logic needs the sheet name: prompt, then delegate.
            let st = state.clone();
            dialogs::find_dialog(move |result| {
                if let Some(text) = result {
                    run_prompt_action(st.app_mut(), "delete_sheet", &text);
                    refresh_after_dialog(&st);
                }
            });
        }
        "export_tsv" | "export_csv" | "export_ods" | "export_ascii" | "export_all" => {
            // Shared export logic writes the file (same as pancurses/ratatui);
            // the save dialog only supplies the destination path.
            let st = state.clone();
            let action = name.to_string();
            if let Some(path) = dialogs::file_save_dialog() {
                run_prompt_action(st.app_mut(), &action, &path.display().to_string());
                refresh_after_dialog(&st);
            }
        }
        "insert_rows" | "insert_mitosis_row" | "insert_mitosis_col" | "insert_cols"
        | "insert_date" | "insert_time" => {
            // Shared insert logic (same as pancurses/ratatui). Date/Time
            // arrive as Edit{value} and preset the edit buffer for Enter.
            delegate_shared_action(name, state);
        }
        "insert_special_chars" | "insert_hyperlink" | "sort_view" | "persist_sort" => {
            // Shared logic needs prompt input: ask, then delegate.
            let st = state.clone();
            let action = name.to_string();
            dialogs::find_dialog(move |result| {
                if let Some(text) = result {
                    run_prompt_action(st.app_mut(), &action, &text);
                    refresh_after_dialog(&st);
                }
            });
        }
        "format_apply_all" | "format_apply_full_column" | "format_apply_data"
        | "format_apply_special" | "format_apply_cell" | "format_apply_selection"
        | "format_decimal_generic" | "format_currency" | "format_rational"
        | "format_fixed_0" | "format_fixed_1" | "format_fixed_2" | "format_fixed_custom"
        | "format_align_left" | "format_align_center" | "format_align_right"
        | "format_align_default" | "format_reset" => {
            // Shared format logic with the pending scope (same as pancurses).
            delegate_shared_action(name, state);
        }
        // Ratatui-parity menu actions without dedicated GTK widgets yet.
        // Each arm records an honest status (never a silent no-op) so menu
        // activation is observable; the pancurses backend (`actions.rs`)
        // carries the fully-wired implementations.
        "submenu" => {
            app.core.status = "Menu action: submenu placeholder (never dispatched)".into();
        }
        "replay" | "duplicate" | "sheet_prev" | "sheet_next" | "move_sheet" => {
            // Shared workbook logic (same as pancurses/ratatui).
            delegate_shared_action(name, state);
        }
        "set_max_col_width" | "set_col_width" => {
            let st = state.clone();
            let action = name.to_string();
            dialogs::find_dialog(move |result| {
                if let Some(text) = result {
                    run_prompt_action(st.app_mut(), &action, &text);
                    refresh_after_dialog(&st);
                }
            });
        }
        "extrapolate" => {
            // Enter interactive extrapolate modal (mirrors ratatui).
            extrapolate::enter(state.app_mut());
            extrapolate::refresh_preview(state.app_mut());
            update_state_cursor(state, state.last_row.get(), state.last_col.get());
            // Return keyboard focus to the canvas so subsequent arrow/Enter keys
            // reach the grid (the menu action left focus on the menu bar).
            state.canvas.grab_focus();
            state.canvas.queue_redraw();
        }
        "copy_sheet" => {
            let st = state.clone();
            dialogs::find_dialog(move |result| {
                if let Some(text) = result {
                    run_prompt_action(st.app_mut(), "copy_sheet", &text);
                    refresh_after_dialog(&st);
                }
            });
        }
        "go_to_cell" => {
            let st = state.clone();
            dialogs::find_dialog(move |result| {
                if let Some(text) = result {
                    run_prompt_action(st.app_mut(), "go_to_cell", &text);
                    refresh_after_dialog(&st);
                }
            });
        }
        "help_rows" => {
            app.core.status = "Row ops: select full rows, then move to target row".into();
        }
        "help_cols" => {
            app.core.status = "Col ops: select full columns, then move to target column".into();
        }
        "help_full" => {
            dialogs::show_keybinds_help();
        }
        _ => {
            app.core.status = format!("Menu action: {name}");
        }
    }
    // Refresh the formula bar — unless an Edit{value} action just preset an
    // edit buffer (Insert Date/Time). update_formula_bar copies the grid cell
    // into the entry widget, which would wipe the preset (and the entry's
    // change handler could then eat edit_buf too). The status label still
    // needs the new status text.
    if state.editing.get() {
        state.status_label.set_text(&state.app_ref().core.status);
    } else {
        update_formula_bar(state, state.last_row.get(), state.last_col.get());
    }
}

// ---------------------------------------------------------------------------
// Formula entry change callback
// ---------------------------------------------------------------------------

fn on_formula_entry_changed(state: &GuiState) {
    if !state.editing.get() {
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
    rswidgets::core::install_debug_crash_handlers();
    let rxapp = rswidgets::App::init()
        .map_err(|e| format!("GUI init failed: {e}"))?;

    let win = rxapp.new_window()?;
    win.set_title(&format!("corro {}", env!("CARGO_PKG_VERSION")));
    win.set_default_size(1200, 800);

    let vbox = rxapp.new_box(Orientation::Vertical, 0)?;

    // Fit column widths to rendered content
    corro_app.fit_main_columns_to_max_width();

    let hr = HEADER_ROWS;
    let lm = MARGIN_COLS;
    let cursor_row = hr;
    let cursor_col = lm;
    corro_app.core.cursor.row = cursor_row;
    corro_app.core.cursor.col = cursor_col;
    // No selection at rest (matching ratatui, where the anchor only exists
    // transiently for shift-extend/select-all/format-selection).
    corro_app.core.anchor = None;

    let data_rows = 30usize;
    let data_cols = 12usize;

    // Formula bar
    let formula_bar = rxapp.new_box(Orientation::Horizontal, 2)?;
    let addr_label = rxapp.new_label("A1")?;
    let f_label = rxapp.new_label("  fx  ")?;
    let formula_entry = rxapp.new_entry()?;
    formula_entry.set_hexpand(true);
    formula_bar.append(&addr_label);
    formula_bar.append(&f_label);
    formula_bar.append(&formula_entry);
    formula_bar.set_child_hexpand(&formula_entry, true);

    // Canvas inside native scrollbars: the thumb tracks the cursor and
    // scrollbar interaction moves the cursor (selection), so the selected
    // cell is always visible. Expand flags go on BEFORE append (GTK3 freezes
    // pack params at append time). Policy 0 = always show (GtkPolicyType).
    let canvas = rxapp.new_canvas()?;
    canvas.set_size_request(1, 1);
    // Ensure the canvas can receive keyboard focus (needed after commit_edit
    // to return focus — GtkDrawingArea does not accept focus by default).
    canvas.set_can_focus(true);
    let scrolled = rxapp.new_scrolled_window()?;
    scrolled.set_policy(0, 0);
    scrolled.set_child(&canvas);
    scrolled.set_vexpand(true);

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
        last_alt_keyval: Cell::new(0),
        entry_seen: Cell::new(false),
        clipboard: RefCell::new(String::new()),
        pending_scope: Cell::new(0),
        #[cfg(feature = "gtk4")]
        last_dedup_key: Cell::new(0),
        #[cfg(feature = "gtk4")]
        dedup_count: Cell::new(0),
        return_pressed: Cell::new(false),
        #[cfg(feature = "gtk4")]
        last_keyval_dedup: Cell::new(0),
        scrolled: scrolled.clone(),
        syncing_scroll: Cell::new(false),
    });

    // Scrollbar interaction moves the cursor (selection); the per-frame
    // viewport recompute then keeps it visible. Reentrancy-safe: syncs
    // from sync_scrollbars (below) set the guard.
    {
        let shared_scroll = shared.clone();
        scrolled.on_scroll(Box::new(move |vertical: bool, value: f64| {
            scroll_to_cursor(&shared_scroll, vertical, value);
        }));
    }

    // Build menu
    let menubar = build_menu(&rxapp, &win, &shared)?;
    let menubar_cb = menubar.clone();
    vbox.append(&menubar);

    // Keyboard: canvas.on_key_raw, win.on_event_key, etc.
    let shared_key = shared.clone();
    canvas.on_key_raw(Box::new(move |keyval: u32, state: u32| -> bool {
        // Alt+letter is reserved for menu mnemonics (handled natively by GTK
        // after propagation): never start a grid edit from it.  Without this
        // guard a canvas-focused Alt+F would type 'f' into the grid, since
        // this path (unlike entry/window) otherwise drops modifier state.
        if (state & 0x8) != 0 {
            return false;
        }
        // Modifier state arrives here (unlike common::Canvas::on_key, which
        // drops it), but only Alt is inspected: Shift+arrows from a
        // canvas-focused keypress still move plainly; entry/window paths
        // below carry Shift (bit 0x1, GDK_SHIFT_MASK / Win32-shift bit).
        handle_key(keyval, &shared_key, 0)
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
        handle_key(0xFF0D, &shared_act, 0);
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

            // Clear the return_pressed flag when a non-RETURN key arrives.
            // This prevents the flag from persisting across unrelated key
            // sequences (e.g., RETURN press with no release event, followed
            // by a legitimate later RETURN press that should be handled).
            let nk = normalize(keyval);
            if nk != RETURN {
                s.return_pressed.set(false);
            }

            let ch = char::from_u32(keyval).unwrap_or('\0').to_ascii_lowercase();
            let alt_held = (state & 0x8) != 0;
            let ctrl_held = (state & 0x4) != 0;

            // General press/release dedup. GTK4-only (see GuiState): elsewhere
            // every key event is genuine (releases are filtered at the source
            // or never hooked), so skipping repeats would drop real input
            // (e.g. Right,Right).
            #[cfg(feature = "gtk4")]
            if nk != 0 && nk == s.last_keyval_dedup.get() {
                s.last_keyval_dedup.set(0);
                append_keylog(&format!("dedup: skipping keyval={nk} release\n"));
                return 1;
            }

            // Ctrl+Q: quit
            if ctrl_held && ch == 'q' {
                eprintln!("DEBUG window key handler: Ctrl+Q detected");
                save_before_quit(s);
                return 1;
            }

            // ALT+letter: let GTK open the submenu natively via the underscore
            // mnemonics baked into the GTK3 menu labels (build_gtk3).  Do NOT
            // engage the Rust-side handle_menu_key fallback here: it consumes
            // the key and only sets internal flags without displaying any
            // popup, so the user sees nothing happen.  Native handling owns the
            // whole interaction (open, navigate, dismiss); returning 0
            // propagates the event to the toplevel default handler which
            // activates the mnemonic.  (GTK4's adapter ignores Alt+letter and
            // already propagated, so this is a no-op there.)
            if alt_held && (32..=126).contains(&keyval) {
                append_keylog(&format!("Alt+letter keyval={keyval} propagating to GTK mnemonics\n"));
                return 0;
            }

            // Check whether the entry already processed this printable
            // character.  On GTK4 BUBBLE phase, the entry's on_key_raw
            // fires first (line ~1405), calls handle_key, and sets
            // entry_processed_key=true.  Without this guard the window
            // handler would call handle_key again, doubling the character
            // in edit_buf ("4422" instead of "42").  This is the fix
            // described in Attempt 195 of the idea log.
            if (32..=126).contains(&nk) && s.entry_processed_key.get() {
                s.entry_processed_key.set(false);
                return 0;
            }

            // When the keyboard menu is active (Alt+letter opened a submenu
            // via handle_menu_key's Rust-side fallback), route unmodified
            // printable keys through handle_menu_key which calls
            // handle_mnemonic_key to select items by mnemonic.  This avoids
            // the character being typed into the formula entry instead.
            //
            // Skip if the key matches the Alt+letter that opened the menu:
            // on GTK, both key-press and key-release events fire the window
            // callback.  After Alt+S opens the Sheet submenu, the subsequent
            // 's' release event (or a second press from xdotool) arrives with
            // state=0 and must not be routed as a menu mnemonic — doing so
            // causes handle_mnemonic_key to globally search all submenus,
            // find "Save" (File → Save, mnemonic _s), and close the menu
            // before the intended mnemonic (e.g., 'b' for Balance Books)
            // arrives.
            if menubar_cb.menu_active() && !alt_held && !ctrl_held
                && (32..=126).contains(&nk)
            {
                if nk == s.last_alt_keyval.get() {
                    return 1;
                }
                let consumed = menubar_cb.handle_menu_key(keyval, state);
                if consumed {
                    // Mark this printable key as processed so the BUBBLE-phase
                    // fallthrough (handle_key) doesn't re-process it.  Without
                    // this guard, if a second key-press event for the same
                    // character arrives (autorepeat or BUBBLE re-entry), it
                    // would fall through to handle_key and start editing with
                    // that character, corrupting the test.
                    s.entry_processed_key.set(true);
                    return 1;
                }
            }

            // Escape when keyboard menu is active: close it via handle_menu_key.
            if nk == ESCAPE && menubar_cb.menu_active() {
                menubar_cb.handle_menu_key(keyval, state);
                return 1;
            }

            // Safety net for RETURN: if editing is false but the formula entry
            // has text or edit_buf has content (e.g., from CAPTURE-phase key
            // processing on GTK where the entry widget never received the key),
            // set editing=true before delegating to handle_key so the edit is
            // committed instead of moving the cursor.
            //
            // On GTK, the CAPTURE-phase EventControllerKey fires for both
            // GDK_KEY_PRESS and GDK_KEY_RELEASE of the same physical key.
            // The press event commits the edit (editing=true → handle_key →
            // commit_edit → editing=false).  After commit_edit, move_cursor
            // calls update_formula_bar which sets the formula entry text to
            // the new cell's value (e.g., "2").  The release event then
            // arrives with editing=false, sees text="2", and re-enters edit
            // mode via this safety net — committing "2" instead of "Hello".
            // The return_pressed flag is set after handle_key processes a
            // RETURN (below).  When the safety net fires on the release
            // event, return_pressed is true and we skip to prevent the
            // spurious second commit.
            if nk == RETURN && !s.editing.get() {
                if s.return_pressed.get() {
                    s.return_pressed.set(false);
                    append_keylog("return_pressed: skipping RETURN release\n");
                    return 1;
                }
                let text = s.formula_entry.get_text().unwrap_or_default();
                if !text.is_empty() || !s.edit_buf.borrow().is_empty() {
                    s.editing.set(true);
                    if !text.is_empty() {
                        *s.edit_buf.borrow_mut() = text;
                    }
                }
            }
    let hk = handle_key(keyval, &state_w, state);
    append_keylog(&format!("handle_key={hk}\n"));
    if hk {
        // Same-event double-fire guard: on setups where the entry observes
        // the same event after the window, it clears this flag and skips its
        // duplicate handle_key call (prevents window+entry doubling).
        // Only arm it once the entry has proven it can observe events
        // (entry_seen); otherwise — streams where the entry never fires —
        // the flag would linger and wrongly skip the NEXT keypress (every
        // second typed char silently lost). Never armed on Windows: the
        // window never observes entry-focused keys there, so the flag could
        // never be cleared same-event. Always armed on GTK4 (legacy).
        if (32..=126).contains(&nk)
            && (cfg!(feature = "gtk4") || (!cfg!(windows) && s.entry_seen.get()))
        {
            s.entry_processed_key.set(true);
        }
        if nk == RETURN {
            s.return_pressed.set(true);
            append_keylog(&format!("window handler returned 1 for RETURN, editing={} text={:?} edit_buf={:?}\n",
                s.editing.get(),
                s.formula_entry.get_text().unwrap_or_default(),
                *s.edit_buf.borrow()));
        }
        // GTK4-only release bookkeeping (see GuiState); elsewhere the field
        // does not exist.
        #[cfg(feature = "gtk4")]
        s.last_keyval_dedup.set(nk);
        1 
    } else {
        #[cfg(feature = "gtk4")]
        s.last_keyval_dedup.set(0);
        0 
    }
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
            // Capability witness for the entry_processed_key protocol (see the
            // window handler): once this has fired, both layers can observe
            // the same event, so the window arms the same-event guard.
            shared_k_cnt.entry_seen.set(true);
            let k = normalize(keyval);
            match k {
                RETURN | ESCAPE | TAB | LEFT | RIGHT | UP | DOWN | HOME | END | PAGE_UP | PAGE_DOWN
                | BACKSPACE | DELETE => {
                    // `state` carries the modifier mask (bit 0x1 = Shift on
                    // both GDK and the nwg adapter); Shift+arrows extend the
                    // selection via handle_key.
                    handle_key(keyval, &shared_k, state);
                    true
                }
                _ if (32..=126).contains(&k) => {
                    // When the window CAPTURE controller (which fires before
                    // this entry CAPTURE controller) already processed this key
                    // and set entry_processed_key, skip the duplicate.  On some
                    // GTK versions GDK_EVENT_STOP from CAPTURE doesn't stop
                    // propagation to child widgets, so both the window CAPTURE
                    // and the entry CAPTURE fire for the same key event.
                    if shared_k.entry_processed_key.get() {
                        shared_k.entry_processed_key.set(false);
                        return false;
                    }
                    // Process the key to update edit_buf (via handle_key →
                    // handle_edit_key or start_edit_with), which also syncs the
                    // widget text to edit_buf (sync_entry_to_buf), then consume
                    // the event so the native widget does NOT insert the char
                    // a second time (doubling, "AA") — the widget already
                    // shows exactly what will be committed.
                    //
                    // For keys with Ctrl (0x4) or Alt (0x8) modifiers, do NOT
                    // claim the key so the event bubbles to the window BUBBLE
                    // handler which processes Ctrl+Q.
                    if (state & (0x4 | 0x8)) != 0 {
                        return false;
                    }
                    handle_key(keyval, &shared_k, state);
                    // NOTE: do NOT set entry_processed_key here. That flag is
                    // strictly window-arms / entry-clears for the SAME event
                    // (double-fire setups). If the entry set it after handling
                    // a key, the window would skip a LATER genuine key thinking
                    // the entry already handled it — silently dropping every
                    // second char on streams where the window never observes
                    // entry keys.
                    true
                }
                _ => false,
            }
        }));
    }

// Assemble layout
    vbox.append(&formula_bar);
    // NOTE (GTK3): expand must be set BEFORE append — pack_start freezes the
    // expand/fill params at append time, so setting vexpand after appending
    // has no effect and the scrolled sheet would never grow vertically.
    scrolled.set_vexpand(true);
    vbox.set_child_vexpand(&scrolled, true);
    vbox.append(&scrolled);
    // The nwg manual box layout looks the child up by handle, so it needs
    // the flag set AFTER append as well (a no-op repeat everywhere else).
    vbox.set_child_vexpand(&scrolled, true);
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
        // Size the viewport from the live canvas every frame (cheap: a few
        // visible_col_indices passes) so the sheet always fills the canvas
        // after menus/chrome, at any window size. Hardcoded counts leave a
        // huge blank area whenever the canvas outgrows them.
        {
            let app = shared_draw.app_ref();
            shared_draw.data_cols.set(cols_to_fill_px(
                app,
                app.core.cursor,
                (w as f64 - ROW_LABEL_W).max(0.0) as i32,
            ));
            shared_draw.data_rows.set(rows_to_fill_px(h));
            // Keep the scrollbar thumb on the cursor (ranges track grid
            // growth here too).
            sync_scrollbars(&shared_draw);
        }
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
    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    if let Some(loader) = rswidgets::backends::gtk::loader() {
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
