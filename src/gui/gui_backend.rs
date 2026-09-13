use rswidgets::prelude::*;
use rswidgets::core::DrawContext;

use std::cell::{Cell, RefCell};
use std::collections::HashMap;
use std::rc::Rc;

use super::actions::run_prompt_action;
use super::actions::{dispatch_menu_action, menu_action_needs_prompt, MenuDispatch};
use super::extrapolate;

use crate::grid::{CellAddr, GridBox, SheetCursor, HEADER_ROWS, MARGIN_COLS};
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
    window: Window,
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
    // Starts true: the formula entry grabs focus at setup, so it observes
    // from the first keypress — and the very first keypress needs the guard
    // (a single Left from A1 must land [A1, not two columns deep). Never
    // armed on Windows (gate below): the window never observes entry-focused
    // keys there, so the flag could never be cleared same-event.
    entry_seen: Cell<bool>,
    // Shared-dispatch state for menu actions: clipboard for cut/copy/paste
    // and the pending format scope. The pancurses backend passes these as
    // dispatch arguments; the GUI backend keeps them here so every menu item
    // behaves identically across backends instead of drifting into stubs.
    clipboard: RefCell<String>,
    pending_scope: Cell<u8>,
    // Row/column pins (padlock feature): logical rows / global cols pinned
    // visible while scrolling. Session-only (not persisted); cleared lazily
    // when the active sheet changes (see pinned_sets). Hit rects painted
    // last frame, used for click toggling.
    pinned_rows: RefCell<std::collections::BTreeSet<usize>>,
    pinned_cols: RefCell<std::collections::BTreeSet<usize>>,
    pinned_sheet: Cell<u32>,
    padlocks: RefCell<Vec<GutterPadlock>>,
    // Sheet tab bar (below the grid): one tab per sheet, visible only when
    // the workbook has 2+ sheets. Hit rects painted last frame, used for
    // click-to-switch. Cached visibility avoids redundant set_visible calls.
    tabbar: Canvas,
    tab_hits: RefCell<Vec<TabHit>>,
    tabbar_visible: Cell<bool>,
    // Scrollbar sync: native scrollbars around the sheet (thumb tracks the
    // viewport; dragging/clicking moves the cursor, which pulls the
    // viewport along, so the selection stays visible). Guard against reentrancy between programmatic sets and the
    // value-changed notification.
    scrolled: rswidgets::common::ScrolledWindow,
    syncing_scroll: Cell<bool>,
    // Last pushed scrollbar state (hval, hupper, hpage, vval, vupper, vpage).
    // sync_scrollbars only configures an adjustment when its desired state
    // differs — re-pushing identical values every draw is not just waste:
    // each configure revalidates the range (async value-changed echoes),
    // and a push mid-drag snaps a user-moved thumb back to the stale
    // cursor-implied position. -1 forces the first push.
    sb_push: Cell<(f64, f64, f64, f64, f64, f64)>,
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

/// True for data cells outside the main body: left/right margin columns,
/// header rows, footer rows. Margin zones render dimmer than body content.
fn is_margin_cell(
    logical_row: usize,
    global_col: usize,
    hr: usize,
    mr: usize,
    lm: usize,
    mc: usize,
) -> bool {
    !(logical_row >= hr
        && logical_row < hr + mr
        && global_col >= lm
        && global_col < lm + mc)
}

/// The one-past-main ring (row hr+mr across body cols, col lm+mc across
/// body rows, plus their corner): the clickable trailing blank. It renders
/// body-white (not margin gray) so the data area visibly includes it on
/// startup (B row / column B on an empty sheet), WITHOUT extending main
/// extent — footer addressing and ratatui parity depend on main staying
/// put until the blank is actually entered (click promotion grows it then).
fn is_trailing_blank_cell(
    logical_row: usize,
    global_col: usize,
    hr: usize,
    mr: usize,
    lm: usize,
    mc: usize,
) -> bool {
    (logical_row == hr + mr && global_col >= lm && global_col <= lm + mc)
        || (global_col == lm + mc && logical_row >= hr && logical_row <= hr + mr)
}

fn render_to(
    sink: &GuiCanvasSink,
    dc: &mut dyn DrawContext,
    col_ixs: &[usize],
    col_widths: &HashMap<usize, usize>,
    display_rows: &[usize],
    hr: usize,
    mr: usize,
    mc: usize,
    lm: usize,
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
            } else if is_trailing_blank_cell(logical_row, c, hr, mr, lm, mc) {
                // Clickable trailing blank: body-white so the data area
                // reads as data (see is_trailing_blank_cell).
                (1.0, 1.0, 1.0, 1.0)
            } else if is_margin_cell(logical_row, c, hr, mr, lm, mc) {
                // Margin-zone data cells render at 75% background brightness
                // so margins read as subordinate to body content.
                (0.75, 0.75, 0.75, 1.0)
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
    // The grid is exhausted exactly when every column is returned (column
    // indices live in [0, total), so len >= total means full coverage).
    let total = MARGIN_COLS + sheet.grid.main_cols() + MARGIN_COLS;
    let mc = sheet.grid.main_cols();
    let mut dim = 1usize;
    loop {
        let (cols, _) = ui_core::visible_col_indices(sheet, cursor, dim, 0);
        let used: f64 = cols
            .iter()
            .map(|&c| display_col_width(sheet, c, mc) as f64 * CHAR_W)
            .sum();
        // Exit when covered, capped, or exhausted. Never exit on
        // `cols.len() < dim`: visible_col_indices may legitimately return
        // fewer than dim (e.g. dim-1 with a left-margin cursor, or equal
        // consecutive counts that still grow later), and bailing there
        // strands the sheet narrow with a huge blank area.
        if used >= avail_px as f64 || dim >= 2048 || cols.len() >= total {
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

/// Paint the row-label gutter. Labels render bold (weight 1), matching the
/// ratatui reference (which bolds the active row and footer rows) and plain
/// spreadsheet convention. Pure apart from the draw calls, so unit tests
/// drive it headlessly with the recording context.
fn paint_row_headers(
    dc: &mut dyn DrawContext,
    display_rows: &[usize],
    mr: usize,
    pinned: &std::collections::BTreeSet<usize>,
    out_padlocks: &mut Vec<GutterPadlock>,
) {
    for (ri, &logical_row) in display_rows.iter().enumerate().take(MAX_RENDER_ROWS) {
        let ry = HEADER_H + ri as f64 * ROW_H;
        let label = crate::addr::ui_row_label(logical_row, mr);
        let (_, _, tw, _) = dc.text_extents_styled(&label, "monospace", FONT_SIZE, 0, 1);
        dc.fill_rect(0.0, ry, ROW_LABEL_W, ROW_H, 0.9, 0.9, 0.9, 1.0);
        // Row numbers sit 6px off the gutter's right gridline so glyphs
        // never touch it.
        dc.draw_text_styled(ROW_LABEL_W - tw - 6.0, ry + 2.0, &label, "monospace", FONT_SIZE, 0.3, 0.3, 0.3, 1.0, 0, 1);
        // Padlock at the gutter's left edge (short labels only): the label
        // is right-aligned, so the left side always has room.
        if wants_padlock(&label) {
            let locked = pinned.contains(&logical_row);
            let (px, py) = (2.0, ry + (ROW_H - PADLOCK_H) / 2.0);
            paint_padlock(dc, px, py, locked);
            out_padlocks.push(GutterPadlock {
                x: px,
                y: py,
                w: PADLOCK_W,
                h: PADLOCK_H,
                is_row: true,
                index: logical_row,
                locked,
            });
        }
    }
}

/// Paint the column-label header strip. Labels render bold (weight 1),
/// matching the ratatui reference (all column headers bold). See
/// [`paint_row_headers`] for testability notes.
fn paint_col_headers(
    dc: &mut dyn DrawContext,
    col_ixs: &[usize],
    col_widths: &HashMap<usize, usize>,
    mc: usize,
    pinned: &std::collections::BTreeSet<usize>,
    out_padlocks: &mut Vec<GutterPadlock>,
) {
    for (ci, &c) in col_ixs.iter().enumerate().take(MAX_RENDER_COLS) {
        let cw = *col_widths.get(&c).unwrap_or(&8) as f64 * CHAR_W;
        let cx = ROW_LABEL_W + col_ixs.iter().take(ci).map(|&pc| *col_widths.get(&pc).unwrap_or(&8) as f64 * CHAR_W).sum::<f64>();
        let col_name = crate::addr::ui_column_fragment(c, mc);
        dc.fill_rect(cx, 0.0, cw, HEADER_H, 0.9, 0.9, 0.9, 1.0);
        let (_, _, tw, _) = dc.text_extents_styled(&col_name, "monospace", FONT_SIZE, 0, 1);
        // Label and padlock center as a unit, so the icon fits inside its
        // own column instead of dangling past the gridline: the column is
        // one character wider than recorded (see display_col_width)
        // precisely to hold this group.
        let lock = wants_padlock(&col_name);
        let group = tw + if lock { 2.0 + PADLOCK_W } else { 0.0 };
        let tx = cx + (cw - group) / 2.0;
        dc.draw_text_styled(tx, (HEADER_H - FONT_SIZE * 1.2) / 2.0, &col_name, "monospace", FONT_SIZE, 0.3, 0.3, 0.3, 1.0, 0, 1);
        // Padlock right after the centered text, inside its own column: the
        // column is one character wider than recorded (see display_col_width)
        // precisely so this icon fits without spilling over the neighbor.
        if lock {
            let locked = pinned.contains(&c);
            let (px, py) = (tx + tw + 2.0, (HEADER_H - PADLOCK_H) / 2.0);
            paint_padlock(dc, px, py, locked);
            out_padlocks.push(GutterPadlock {
                x: px,
                y: py,
                w: PADLOCK_W,
                h: PADLOCK_H,
                is_row: false,
                index: c,
                locked,
            });
        }
    }
}

/// Padlock affordance geometry (device px): small enough for 20px rows and
/// the 24px header strip, big enough to click and to read at a glance.
const PADLOCK_W: f64 = 10.0;
const PADLOCK_H: f64 = 12.0;
/// Unlocked padlock slate (115): distinct from header gray (77), grid lines
/// (204), backgrounds (191/229/255) and cursor/selection blues.
const PADLOCK_OPEN: (f64, f64, f64) = (0.45, 0.45, 0.45);
/// Locked padlock slate (51): darker than any chrome gray.
const PADLOCK_SHUT: (f64, f64, f64) = (0.2, 0.2, 0.2);

/// A painted padlock hit target from the last frame: gutter position plus
/// which logical row (is_row) or global column it pins.
#[derive(Clone, Copy, Debug, PartialEq)]
struct GutterPadlock {
    x: f64,
    y: f64,
    w: f64,
    h: f64,
    is_row: bool,
    index: usize,
    locked: bool,
}

/// Padlock eligibility: gutter labels with short text only (dual-character
/// or less, non-empty) get the affordance — the auto-generated margin
/// labels (`1`, `_1`, `[A`, `AA`, ...) rather than long content.
fn wants_padlock(label: &str) -> bool {
    !label.is_empty() && label.chars().count() <= 2
}

/// Toggle a pin in a set; returns true when the index ends up pinned.
fn toggle_pin_in(set: &mut std::collections::BTreeSet<usize>, idx: usize) -> bool {
    if set.contains(&idx) {
        set.remove(&idx);
        false
    } else {
        set.insert(idx);
        true
    }
}

/// Merge pinned rows/cols ahead of the normal display window (frozen at the
/// top/left, like frozen panes). `pinned` must be ascending (BTreeSet order
/// qualifies). Entries already in `display` are not duplicated.
fn union_pinned(display: &[usize], pinned: &[usize]) -> Vec<usize> {
    let mut out: Vec<usize> = pinned.to_vec();
    out.extend(display.iter().filter(|r| !pinned.contains(r)).copied());
    out
}

/// The right shackle bar connects to the body when locked and floats with a
/// gap when unlocked. Vector rects only (no font/emoji dependency), so both
/// backends and screenshots render it identically.
fn paint_padlock(dc: &mut dyn DrawContext, ox: f64, oy: f64, locked: bool) {
    let (r, g, b) = if locked { PADLOCK_SHUT } else { PADLOCK_OPEN };
    if locked {
        dc.fill_rect(ox + 1.0, oy + 6.0, 8.0, 5.0, r, g, b, 1.0);
    } else {
        dc.stroke_rect(ox + 1.0, oy + 6.0, 8.0, 5.0, r, g, b, 1.0, 1.0);
    }
    dc.fill_rect(ox + 2.0, oy + 1.0, 2.0, 6.0, r, g, b, 1.0);
    dc.fill_rect(ox + 2.0, oy + 1.0, 6.0, 2.0, r, g, b, 1.0);
    if locked {
        dc.fill_rect(ox + 6.0, oy + 1.0, 2.0, 6.0, r, g, b, 1.0);
    } else {
        dc.fill_rect(ox + 6.0, oy + 2.0, 2.0, 3.0, r, g, b, 1.0);
    }
}

/// Current pin sets, clearing them lazily when the active sheet changed
/// (pins are per-sheet session state, never persisted to the log).
fn pinned_sets(state: &GuiState) -> (Vec<usize>, Vec<usize>) {
    let sid = {
        let app = state.app_ref();
        app.core.workbook.sheet_id(app.core.workbook.active_sheet)
    };
    if state.pinned_sheet.get() != sid {
        state.pinned_rows.borrow_mut().clear();
        state.pinned_cols.borrow_mut().clear();
        state.pinned_sheet.set(sid);
    }
    let rows: Vec<usize> = state.pinned_rows.borrow().iter().copied().collect();
    let cols: Vec<usize> = state.pinned_cols.borrow().iter().copied().collect();
    (rows, cols)
}

/// Display rows with pins merged in (frozen first), for rendering and click
/// mapping alike — both must agree or clicks land on the wrong cells.
fn displayed_rows(state: &GuiState) -> Vec<usize> {
    let (display, _) = {
        let app = state.app_ref();
        let sheet = app.core.workbook.active_sheet();
        ui_core::visible_row_indices(sheet, app.core.cursor, state.data_rows.get(), 0)
    };
    let (pinned, _) = pinned_sets(state);
    union_pinned(&display, &pinned)
}

/// Display columns with pins merged in (frozen first). See [`displayed_rows`].
fn displayed_cols(state: &GuiState) -> Vec<usize> {
    let (col_ixs, _) = {
        let app = state.app_ref();
        let sheet = app.core.workbook.active_sheet();
        ui_core::visible_col_indices(sheet, app.core.cursor, state.data_cols.get(), 0)
    };
    let (_, pinned) = pinned_sets(state);
    union_pinned(&col_ixs, &pinned)
}

/// First viewport (non-pinned) row/col of the current display: what the
/// scrollbar thumbs track. The thumb follows the viewport (like Excel),
/// not the cursor — in-viewport arrow steps leave it alone; it moves only
/// when the viewport itself scrolls. Pinned rows/cols render frozen first
/// and must not nail the thumb, so they are skipped (the fallback is the
/// body origin when everything displayed is pinned).
fn viewport_origin(state: &GuiState) -> (usize, usize) {
    let (pinned_rows, pinned_cols) = pinned_sets(state);
    let rows = displayed_rows(state);
    let cols = displayed_cols(state);
    let r = rows
        .iter()
        .copied()
        .find(|r| !pinned_rows.contains(r))
        .unwrap_or(HEADER_ROWS);
    let c = cols
        .iter()
        .copied()
        .find(|c| !pinned_cols.contains(c))
        .unwrap_or(MARGIN_COLS);
    (r, c)
}

/// Toggle the pin hit-tested from the last frame's padlocks; returns true
/// when the row/column ends up pinned.
fn toggle_pin(state: &GuiState, is_row: bool, idx: usize) -> bool {
    let _ = pinned_sets(state);
    if is_row {
        toggle_pin_in(&mut state.pinned_rows.borrow_mut(), idx)
    } else {
        toggle_pin_in(&mut state.pinned_cols.borrow_mut(), idx)
    }
}

// ---------------------------------------------------------------------------
// Sheet tab bar
// ---------------------------------------------------------------------------

/// Height of the sheet tab strip (device px): matches the 24px column-header
/// strip so the chrome reads as one family.
const TAB_H: f64 = 24.0;
/// Active tab fill: the terminal reference paints the active sheet tab
/// black-on-yellow bold (ratatui `draw_visual`); the GUI uses a softer
/// yellow that stays distinct from its blue selection/cursor language.
const TAB_ACTIVE_BG: (f64, f64, f64) = (1.0, 1.0, 0.6);
/// Inactive tab fill: same gray as the row/column gutters.
const TAB_IDLE_BG: (f64, f64, f64) = (0.9, 0.9, 0.9);
/// Tab divider lines.
const TAB_DIV: (f64, f64, f64) = (0.55, 0.55, 0.55);
/// Horizontal padding inside each tab and gap between tabs (device px).
const TAB_PAD_X: f64 = 10.0;
const TAB_GAP: f64 = 6.0;

/// A painted sheet tab from the last frame: strip position plus which sheet
/// index a click switches to.
#[derive(Clone, Copy, Debug, PartialEq)]
struct TabHit {
    x0: f64,
    x1: f64,
    index: usize,
}

/// Lay out sheet tabs left to right from x=2: each tab pads its measured
/// title by TAB_PAD_X on both sides, TAB_GAP separates tabs. The active tab
/// measures bold (weight 1), like it paints. Pure apart from measuring, so
/// unit tests drive it with a stub measure closure.
fn tab_layout(
    titles: &[String],
    active: usize,
    measure: &dyn Fn(&str, i32) -> f64,
) -> Vec<TabHit> {
    let mut hits = Vec::with_capacity(titles.len());
    let mut x = 2.0;
    for (index, title) in titles.iter().enumerate() {
        let weight = if index == active { 1 } else { 0 };
        let tw = measure(title, weight);
        let x0 = x;
        let x1 = x0 + TAB_PAD_X + tw + TAB_PAD_X;
        hits.push(TabHit { x0, x1, index });
        x = x1 + TAB_GAP;
    }
    hits
}

/// Show the tab strip iff the workbook has 2+ sheets, then repaint it.
/// Called from refresh_after_dialog (every menu/keyboard action) — never
/// driven from the draw callback alone, because a hidden widget never
/// draws and could never re-show itself.
fn sync_tabbar(state: &Rc<GuiState>) {
    let show = state.app_ref().core.workbook.sheet_count() >= 2;
    if show != state.tabbar_visible.get() {
        state.tabbar.set_visible(show);
        state.tabbar_visible.set(show);
    }
    state.tabbar.queue_redraw();
}

/// Paint the sheet tab strip: one tab per sheet title, the active sheet
/// bold on yellow (mirroring the terminal's black-on-yellow active tab).
/// Reads live workbook state every frame, so created/renamed/deleted/copied
/// sheets repaint with only a tabbar redraw (see refresh_after_dialog).
/// When below 2 sheets the bar hides itself on every draw unconditionally:
/// present() runs gtk_widget_show_all inside a seconds-long event pump,
/// resurrecting setup-hidden widgets while screenshots and input are already
/// live — a change-guarded hide would never fire (cache already says hidden).
/// Once hidden GTK stops drawing, so the unconditional hide costs nothing
/// steady-state.
fn render_tabbar(dc: &mut dyn DrawContext, state: &GuiState, w: i32, h: i32) {
    let (titles, active) = {
        let app = state.app_ref();
        let wb = &app.core.workbook;
        let titles: Vec<String> =
            (0..wb.sheet_count()).map(|i| wb.sheet_title(i).to_string()).collect();
        (titles, wb.active_sheet)
    };
    let show = titles.len() >= 2;
    if !show {
        // Unconditional (see doc comment): heals show_all resurrection.
        state.tabbar.set_visible(false);
        state.tabbar_visible.set(false);
        state.tab_hits.borrow_mut().clear();
        return;
    }
    if !state.tabbar_visible.get() {
        // Normally sync_tabbar already showed it; heal any drift.
        state.tabbar.set_visible(true);
        state.tabbar_visible.set(true);
    }
    dc.clear(0.94, 0.94, 0.94, 1.0);
    dc.clip(0.0, 0.0, w as f64, h as f64);
    let hits = {
        let measure = |t: &str, weight: i32| {
            dc.text_extents_styled(t, "monospace", FONT_SIZE, 0, weight).2
        };
        tab_layout(&titles, active, &measure)
    };
    for hit in &hits {
        let is_active = hit.index == active;
        let (r, g, b) = if is_active { TAB_ACTIVE_BG } else { TAB_IDLE_BG };
        dc.fill_rect(hit.x0, 2.0, hit.x1 - hit.x0, TAB_H - 4.0, r, g, b, 1.0);
        // Divider at the tab's right edge (also the click target's edge).
        dc.fill_rect(hit.x1, 2.0, 1.0, TAB_H - 4.0, TAB_DIV.0, TAB_DIV.1, TAB_DIV.2, 1.0);
        let (tr, tg, tb) = if is_active { (0.0, 0.0, 0.0) } else { (0.3, 0.3, 0.3) };
        let weight = if is_active { 1 } else { 0 };
        dc.draw_text_styled(
            hit.x0 + TAB_PAD_X,
            (TAB_H - FONT_SIZE * 1.2) / 2.0,
            &titles[hit.index],
            "monospace",
            FONT_SIZE,
            tr,
            tg,
            tb,
            1.0,
            0,
            weight,
        );
    }
    *state.tab_hits.borrow_mut() = hits;
}

/// Switch to the tabbed sheet (mirrors sheet_prev/sheet_next arrival:
/// cursor parks at A1, status names the sheet). No-op for the active tab
/// or an out-of-range index.
fn switch_to_sheet(state_rc: &Rc<GuiState>, index: usize) {
    let state: &GuiState = &**state_rc;
    let app = state.app_mut();
    let n = app.core.workbook.sheet_count();
    if index >= n || index == app.core.workbook.active_sheet {
        return;
    }
    app.core.workbook.active_sheet = index;
    app.core.view_sheet_id = app.core.workbook.sheet_id(index);
    app.core.cursor = SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS };
    app.core.anchor = None;
    app.core.status = format!("Sheet {} of {}", index + 1, n);
    state.last_row.set(HEADER_ROWS);
    state.last_col.set(MARGIN_COLS);
    // A fresh sheet still gets its clickable trailing blank. No pointer
    // growth here: landing on the last cell of a 1x1 sheet must not grow
    // (that would shift footer addressing and break ratatui parity) — the
    // ring renders white regardless (see is_trailing_blank_cell).
    maintain_extent(state, false);
    update_formula_bar(state, HEADER_ROWS, MARGIN_COLS);
    state.canvas.queue_redraw();
    state.tabbar.queue_redraw();
}

/// Click on the tab strip: a tab hit switches sheets; anything else is
/// ignored and must never move the grid cursor.
fn handle_tab_click(x: f64, state_rc: &Rc<GuiState>) {
    let state: &GuiState = &**state_rc;
    let hit = state.tab_hits.borrow().iter().find(|h| x >= h.x0 && x < h.x1).copied();
    if let Some(hit) = hit {
        switch_to_sheet(state_rc, hit.index);
    }
}

fn render_grid(dc: &mut dyn DrawContext, state: &GuiState, w: i32, h: i32) {
    dc.clear(0.94, 0.94, 0.94, 1.0);
    dc.clip(0.0, 0.0, w as f64, h as f64);

    let app = state.app_ref();
    let hr = HEADER_ROWS;
    let lm = MARGIN_COLS;
    let cursor_row = state.last_row.get();
    let cursor_col = state.last_col.get();

    let display_rows: Vec<usize> = displayed_rows(state);
    let col_ixs: Vec<usize> = displayed_cols(state);
    let mr = app.core.workbook.active_sheet().grid.main_rows();
    let mc = app.core.workbook.active_sheet().grid.main_cols();

    let col_widths: HashMap<usize, usize> = col_ixs.iter()
        .map(|&c| (c, display_col_width(&app.core.workbook.active_sheet(), c, mc)))
        .collect();

    // Row headers (padlock hit rects refresh every frame for click mapping).
    let pinned_rows: std::collections::BTreeSet<usize> =
        state.pinned_rows.borrow().iter().copied().collect();
    let pinned_cols: std::collections::BTreeSet<usize> =
        state.pinned_cols.borrow().iter().copied().collect();
    let mut padlocks: Vec<GutterPadlock> = Vec::new();
    paint_row_headers(dc, &display_rows, mr, &pinned_rows, &mut padlocks);

    // Selection rectangle (anchor..cursor, rows AND columns). None while
    // navigating plainly — only explicit selections highlight.

    // Column headers
    paint_col_headers(dc, &col_ixs, &col_widths, mc, &pinned_cols, &mut padlocks);
    *state.padlocks.borrow_mut() = padlocks;

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
            hr, mr, mc, lm,
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

/// On-screen width of a column in characters: the recorded width plus one
/// spare character when the column header carries a padlock, so the icon
/// fits inside its own column instead of overlapping the neighbor's
/// gridline. Single source of truth for every width accumulation (render
/// headers, render cells, click mapping, viewport sizing) — they must all
/// agree or columns misalign.
fn display_col_width(sheet: &crate::ops::SheetState, c: usize, mc: usize) -> usize {
    sheet_rec_col_width(sheet, c)
        + usize::from(wants_padlock(&crate::addr::ui_column_fragment(c, mc)))
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
        // header/margin/footer edits into main cells. The first margin
        // row/col addresses (and saves) as footer/margin (_1/]A), same as
        // ratatui.
        let addr = crate::addr::sheet_cursor_to_addr(
            crate::addr::LogicalRow(row),
            crate::addr::GlobalCol(col),
            crate::addr::MainRows(app.core.workbook.active_sheet().grid.main_rows()),
            crate::addr::MainCols(app.core.workbook.active_sheet().grid.main_cols()),
        );
        app.core.workbook.active_sheet_mut().grid.set(&addr, val.clone());
        let sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
        let op = Op::SetCell { addr: addr.clone(), value: val };
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
        app.core.status = format!("Set cell {}", crate::addr::cell_ref_text(&addr, app.core.workbook.active_sheet().grid.main_cols()));
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

/// Whether the body extent needs one more row/column so a blank stays past
/// the last non-blank index. Pure (unit-tested): `extent` is main_rows or
/// main_cols, `trailing_content_blanks` comes from the same trailing-blank
/// counters keyboard navigation uses, and `cursor_main` is the cursor's
/// 0-based main index when it sits inside the body (None in the margins).
/// The cursor counts as non-blank: parking on the last body row/column
/// opens one more beyond it, like an infinite canvas.
fn grow_for_trailing_blank(
    extent: usize,
    trailing_content_blanks: usize,
    cursor_main: Option<usize>,
) -> bool {
    let last_content = extent.checked_sub(trailing_content_blanks + 1);
    let last = match (last_content, cursor_main) {
        (Some(a), Some(b)) => Some(a.max(b)),
        (Some(a), None) => Some(a),
        (None, Some(b)) => Some(b),
        (None, None) => None,
    };
    match last {
        Some(l) => l + 1 >= extent,
        None => false,
    }
}

/// Keep one blank body row below (and one blank body column right of) the
/// last non-blank one (content-counted only). This is what lets the mouse
/// select a fresh data row/column at load: without it the cells beyond the
/// content are margins, so clicking can never start new data. Growth is
/// ephemeral like keyboard growth (never logged).
///
/// Intentional GUI-only divergence from ratatui (which shows no trailing
/// blank at load): the terminal has no mouse, so it needs no clickable
/// blank row. Keyboard behavior is unchanged and still matches ratatui.
///
/// NB: deliberately NOT cursor-counted. Counting the cursor as non-blank
/// here treadmills keyboard navigation: every arrow step onto the trailing
/// blank grows a fresh one ahead, so the cursor never leaves main (never
/// reaching margins/footers, breaking ratatui parity). Pointer arrivals
/// (clicks, scrollbar drags, sheet switches) grow explicitly via
/// grow_blank_past_cursor below; keyboard growth stays with the classic
/// NAV_BLANK step-off logic in move_cursor.
/// Converge the body extent on need: content + one blank, floored by the
/// cursor and by 1x1. Grows toward the target and silently shrinks
/// abandoned growth back toward it (via the cursor floor + the grid's
/// silent shrink_to_content) — this is what lets the sheet shrink again
/// when blank rows/cols are no longer needed (navigate back up, delete
/// content). Target per axis = max(content+1, cursor_idx+1, 1); the +1
/// keeps the clickable trailing blank, the cursor floor keeps the
/// selection addressable, and 1x1 keeps empty sheets footer-addressable
/// (ratatui parity: footers count from main end). All silent (no ops) —
/// file/ratatui parity preserved; only the rendered extent tightens.
/// Pointer arrivals (clicks, scrollbar drags, sheet switches) add
/// grow_blank_past_cursor after this when landing on the last row/col.
/// When `allow_shrink` is false only growth runs (fresh loads/switches must
/// not second-guess the file's stored extent); navigation and post-delete
/// refresh pass true so abandoned growth trims back.
fn maintain_extent(state: &GuiState, allow_shrink: bool) {
    let app = state.app_mut();
    let cursor = app.core.cursor;
    let sheet = app.core.workbook.active_sheet_mut();
    let grid = &mut sheet.grid;
    let (hr, lm) = (HEADER_ROWS, MARGIN_COLS);
    // Content blank via the tested one-blank rule (the cursor never counts
    // here — treadmill). Each iteration re-scans; converges once the blank
    // exists (empty sheets: tb == extent fails the rule, so no growth).
    while grow_for_trailing_blank(grid.main_rows(), compute::trailing_blank_main_rows(grid), None) {
        grid.grow_main_row_at_bottom();
    }
    while grow_for_trailing_blank(grid.main_cols(), compute::trailing_blank_main_cols(grid), None) {
        grid.grow_main_col_at_right();
    }
    if !allow_shrink {
        return;
    }
    // Shrink abandoned growth back toward need (see doc comment): recompute
    // post-growth, then keep max(content+1, cursor floor, 1x1).
    let (mr, mc) = (grid.main_rows(), grid.main_cols());
    let tb_r = compute::trailing_blank_main_rows(grid).min(mr);
    let tb_c = compute::trailing_blank_main_cols(grid).min(mc);
    let content_r = mr - tb_r;
    let content_c = mc - tb_c;
    // Floor: cursor-relative index + 1 while inside the body; the CURRENT
    // extent while outside (margins/headers/footers). Shrinking under an
    // out-of-body cursor re-addresses it (its margin/footer label counts
    // from main end), so the extent is left alone there.
    let floor_r = if cursor.row >= hr && cursor.row < hr + mr {
        cursor.row - hr + 1
    } else {
        mr
    };
    let floor_c = if cursor.col >= lm && cursor.col < lm + mc {
        cursor.col - lm + 1
    } else {
        mc
    };
    let target_r = (content_r + 1).max(floor_r).max(1);
    let target_c = (content_c + 1).max(floor_c).max(1);
    grid.set_min_extent(target_r as u32, target_c as u32);
    grid.shrink_to_content();
}

/// Open one blank body row/column beyond the cursor when it sits exactly
/// on the last body row/column. Pointer arrivals only (clicks, scrollbar
/// drags — see maintain_extent for why keyboard moves must not call this).
/// Growth is ephemeral (never logged).
fn grow_blank_past_cursor(state: &GuiState) {
    let app = state.app_mut();
    let cursor = app.core.cursor;
    let sheet = app.core.workbook.active_sheet_mut();
    let grid = &mut sheet.grid;
    let (hr, lm) = (HEADER_ROWS, MARGIN_COLS);
    let mr = grid.main_rows();
    if mr > 0 && cursor.row >= hr && cursor.row - hr + 1 == mr {
        grid.grow_main_row_at_bottom();
    }
    let mc = grid.main_cols();
    if mc > 0 && cursor.col >= lm && cursor.col - lm + 1 == mc {
        grid.grow_main_col_at_right();
    }
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
    // Converge the extent on need (grows and shrinks; see maintain_extent).
    maintain_extent(state, true);
    update_formula_bar(state, row, col);
    state.canvas.queue_redraw();
    // Window-level cascade: a canvas-only queue_draw may not trigger the
    // toplevel's frame clock (same reason setup calls win.queue_redraw()),
    // leaving pure cursor moves invisible — the state advances but the grid
    // keeps showing the old highlight. Marking the window dirty every move
    // keeps the view honest on every backend.
    state.window.queue_redraw();
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

/// Push viewport position and domain into the native scrollbars (thumb
/// tracks the viewport, like Excel — not the cursor, so in-viewport arrow
/// steps leave it alone and it moves only when the viewport scrolls).
/// Guarded against reentrancy with the value-changed notification below.
/// Domain and page satisfy upper > page so the bars stay live: upper
/// covers content plus a viewport plus footer padding.
fn sync_scrollbars(state: &GuiState) {
    if state.syncing_scroll.get() {
        return;
    }
    let (ru, cu) = scroll_domain(state);
    let (orow, ocol) = viewport_origin(state);
    let vv = orow.saturating_sub(HEADER_ROWS).min(ru.saturating_sub(1));
    let hv = ocol.saturating_sub(MARGIN_COLS).min(cu.saturating_sub(1));
    let page_h = state.data_cols.get().max(1) as f64;
    let page_v = state.data_rows.get().max(1) as f64;
    let want = (hv as f64, cu as f64, page_h, vv as f64, ru as f64, page_v);
    // Push only on genuine change: re-configuring identical values every
    // draw revalidates the ranges (spurious async value-changed echoes)
    // and a push landing mid-drag snaps the thumb back to the stale
    // position. (Exact ints as f64 — equality is exact.)
    if want == state.sb_push.get() {
        return;
    }
    state.syncing_scroll.set(true);
    state.scrolled.scroll_to(
        hv as f64, cu as f64, page_h,
        vv as f64, ru as f64, page_v,
    );
    state.sb_push.set(want);
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
    // The thumb addresses the body viewport only, so scrollbar input is
    // ignored while the cursor sits in chrome (margins/headers): a stale
    // re-emission would otherwise drag such a cursor back into the body
    // (e.g. horizontal 0 resetting [A1 to A1, or vertical 0 resetting A~1
    // to A1). The per-frame viewport recompute still keeps it visible.
    // (Footer rows are expressible — row-hr needs no floor — so they
    // stay live here.)
    if vertical {
        if state.last_row.get() < HEADER_ROWS {
            return;
        }
    } else {
        let lm = MARGIN_COLS;
        let mc = state
            .app_ref()
            .core
            .workbook
            .active_sheet()
            .grid
            .main_cols();
        let c = state.last_col.get();
        if c < lm || c >= lm + mc {
            return;
        }
    }
    let (ru, cu) = scroll_domain(state);
    if vertical {
        let row = (HEADER_ROWS as f64 + value).max(HEADER_ROWS as f64) as usize;
        let row = row.min(HEADER_ROWS + ru.saturating_sub(1));
        update_state_cursor(state, row, state.last_col.get());
        grow_blank_past_cursor(state);
    } else {
        let col = (MARGIN_COLS as f64 + value).max(MARGIN_COLS as f64) as usize;
        let col = col.min(MARGIN_COLS + cu.saturating_sub(1));
        update_state_cursor(state, state.last_row.get(), col);
        grow_blank_past_cursor(state);
    }
}

/// User-facing address text for the formula bar (left of `fx`), matching the
/// ratatui reference (`addr_label`/`cell_ref_text`): `A1`, `[A1`, `A~1`, ...
/// Never `CellAddr`'s internal rendering (`(0, 0)`, `<701>(0)`, ...).
fn formula_addr_label(row: usize, col: usize, grid: &GridBox) -> String {
    let addr = crate::addr::sheet_cursor_to_addr(
        crate::addr::LogicalRow(row),
        crate::addr::GlobalCol(col),
        crate::addr::MainRows(grid.main_rows()),
        crate::addr::MainCols(grid.main_cols()),
    );
    crate::addr::cell_ref_text(&addr, grid.main_cols())
}

fn update_formula_bar(state: &GuiState, row: usize, col: usize) {
    let app = state.app_ref();
    let grid = &app.core.workbook.active_sheet().grid;
    state.addr_label.set_text(&formula_addr_label(row, col, grid));
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
    // Padlock hits first: padlocks live in the gutter chrome that plain
    // clicks ignore, and toggling a pin must not move the cursor, collapse
    // the selection, or start editing.
    let hit = state
        .padlocks
        .borrow()
        .iter()
        .find(|h| x >= h.x && x < h.x + h.w && y >= h.y && y < h.y + h.h)
        .copied();
    if let Some(hit) = hit {
        toggle_pin(state, hit.is_row, hit.index);
        state.canvas.queue_redraw();
        return;
    }
    let app = state.app_mut();
    if x < ROW_LABEL_W || y < HEADER_H {
        return;
    }
    let col_ixs: Vec<usize> = displayed_cols(state);
    let mc = app.core.workbook.active_sheet().grid.main_cols();
    let mut cx = ROW_LABEL_W;
    for &c in &col_ixs {
        let cw = display_col_width(&app.core.workbook.active_sheet(), c, mc) as f64 * CHAR_W;
        if x >= cx && x < cx + cw {
            let ri = ((y - HEADER_H) / ROW_H) as usize;
            // Same pinned-first display set the renderer uses, or clicks
            // land on the wrong rows once pins are active.
            let display_rows: Vec<usize> = displayed_rows(state);
            if ri < display_rows.len() {
                let logical_row = display_rows[ri];
                state.last_row.set(logical_row);
                state.last_col.set(c);
                app.core.cursor.row = logical_row;
                app.core.cursor.col = c;
                // Plain click collapses any selection (fresh single-cell focus).
                app.core.anchor = None;
                // Clicking the trailing blank opens one more beyond it
                // (pointer arrival — see grow_blank_past_cursor).
                maintain_extent(state, false);
                grow_blank_past_cursor(state);
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
    // Menu ops can change content (delete/insert) without moving the cursor:
    // re-converge the extent (trims post-delete bloat, keeps the blank).
    maintain_extent(state, true);
    update_formula_bar(state, state.last_row.get(), state.last_col.get());
    state.canvas.queue_redraw();
    // Created/renamed/deleted/copied sheets change tab titles or the bar's
    // visibility — sync and repaint the strip with every refresh.
    sync_tabbar(state);
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

/// Dialog chrome (window title, OK button label, initial entry text) for a
/// prompt-gated menu action. Pure so unit tests pin every action's labels:
/// a mislabeled dialog (Rename Sheet showing "Find") fails here instead of
/// reaching users. The initial text pre-fills the entry for edit-in-place
/// actions (rename); everything else starts empty.
fn prompt_chrome(action: &str, current_sheet_title: &str) -> (String, String, String) {
    match action {
        "rename_sheet" => ("Rename sheet".into(), "Rename".into(), current_sheet_title.into()),
        "copy_sheet" => ("Copy sheet".into(), "Copy".into(), String::new()),
        "delete_sheet" => ("Delete sheet".into(), "Delete".into(), String::new()),
        "go_to_cell" => ("Go to cell".into(), "Go".into(), String::new()),
        "set_col_width" => ("Column width".into(), "Set".into(), String::new()),
        "set_max_col_width" => ("Default width".into(), "Set".into(), String::new()),
        "find" => ("Find".into(), "Find".into(), String::new()),
        "insert_special_chars" => ("Insert special char".into(), "Insert".into(), String::new()),
        "insert_hyperlink" => ("Insert hyperlink".into(), "Insert".into(), String::new()),
        "sort_view" => ("Sort view".into(), "Sort".into(), String::new()),
        "persist_sort" => ("Persist sort".into(), "Sort".into(), String::new()),
        "balance_books" => ("Balance books".into(), "Balance".into(), String::new()),
        _ => ("Prompt".into(), "OK".into(), String::new()),
    }
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
            // Shared save logic: saves directly with no dialog when a path
            // exists, otherwise prompts Save As (same as pancurses).
            delegate_shared_action(name, state);
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

        _ if menu_action_needs_prompt(name).is_some() => {
            // Every other prompt-gated action funnels through the shared
            // prompt logic; the dialog only supplies the input. Routing by
            // the shared gate (not a hardcoded list) means a newly-added
            // prompt action can never silently fall to the stub below.
            // Replace has two fields, so it gets its dedicated dialog (a
            // single Find box could never supply "find|replacement");
            // everything else gets a correctly labeled single-entry prompt
            // (never a recycled "Find" dialog — see prompt_chrome).
            let st = state.clone();
            let action = name.to_string();
            if action == "replace" {
                dialogs::replace_dialog(move |result| {
                    if let Some((find, repl)) = result {
                        run_prompt_action(st.app_mut(), &action, &format!("{find}|{repl}"));
                        refresh_after_dialog(&st);
                    }
                });
            } else {
                let (title, ok, initial) = {
                    let app = st.app_ref();
                    let wb = &app.core.workbook;
                    prompt_chrome(&action, wb.sheet_title(wb.active_sheet))
                };
                dialogs::prompt_dialog(&title, &ok, &initial, move |result| {
                    if let Some(text) = result {
                        run_prompt_action(st.app_mut(), &action, &text);
                        refresh_after_dialog(&st);
                    }
                });
            }
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

    // Canvas inside native scrollbars: the thumb tracks the viewport and
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

    // Sheet tab strip (below the grid): one tab per sheet once the workbook
    // has 2+ sheets. Hidden until then — the first New sheet creates the
    // bar, later sheets just append tabs. Fixed strip height; a vertical box
    // stretches children across the full width on every backend.
    let tabbar = rxapp.new_canvas()?;
    tabbar.set_size_request(1, TAB_H as i32);
    tabbar.set_visible(false);

    let shared = Rc::new(GuiState {
        app: corro_app as *mut super::App,
        rxapp: rxapp.clone(),
        window: win.clone(),
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
        entry_seen: Cell::new(true),
        clipboard: RefCell::new(String::new()),
        pending_scope: Cell::new(0),
        pinned_rows: RefCell::new(std::collections::BTreeSet::new()),
        pinned_cols: RefCell::new(std::collections::BTreeSet::new()),
        pinned_sheet: Cell::new(u32::MAX),
        padlocks: RefCell::new(Vec::new()),
        tabbar: tabbar.clone(),
        tab_hits: RefCell::new(Vec::new()),
        tabbar_visible: Cell::new(false),
        #[cfg(feature = "gtk4")]
        last_dedup_key: Cell::new(0),
        #[cfg(feature = "gtk4")]
        dedup_count: Cell::new(0),
        return_pressed: Cell::new(false),
        #[cfg(feature = "gtk4")]
        last_keyval_dedup: Cell::new(0),
        scrolled: scrolled.clone(),
        syncing_scroll: Cell::new(false),
        sb_push: Cell::new((-1.0, -1.0, -1.0, -1.0, -1.0, -1.0)),
    });

    // A fresh load still opens with a clickable trailing blank data row/col.
    maintain_extent(&shared, false);

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
        // Same-event duplicate guard (mirrors the entry handler): the window
        // handler fires first and arms entry_processed_key; without this
        // check a canvas-focused keypress moves twice (Left from A1 lands
        // two columns deep instead of [A1). Arm after handling for the
        // reverse order (canvas first, window second); the window's own
        // guard below consumes it.
        if shared_key.entry_processed_key.get() {
            shared_key.entry_processed_key.set(false);
            return false;
        }
        // Modifier state arrives here (unlike common::Canvas::on_key, which
        // drops it), but only Alt is inspected: Shift+arrows from a
        // canvas-focused keypress still move plainly; entry/window paths
        // below carry Shift (bit 0x1, GDK_SHIFT_MASK / Win32-shift bit).
        let r = handle_key(keyval, &shared_key, 0);
        if !cfg!(windows) && shared_key.entry_seen.get() {
            shared_key.entry_processed_key.set(true);
        }
        r
    }));

    // Click
    let shared_click = shared.clone();
    canvas.on_click(Box::new(move |x: f64, y: f64| { handle_click(x, y, &shared_click); }));

    // Sheet tabs: paint the strip and switch sheets on click. Registered
    // before present() like the grid canvas so the first frame draws tabs.
    let shared_tabdraw = shared.clone();
    tabbar.set_draw_callback(Box::new(move |dc: &mut dyn DrawContext, w: i32, h: i32| {
        render_tabbar(dc, &shared_tabdraw, w, h);
    }));
    let shared_tabclick = shared.clone();
    tabbar.on_click(Box::new(move |x: f64, _y: f64| {
        handle_tab_click(x, &shared_tabclick);
    }));

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
            // Skip keys the entry/canvas already handled first (same-event
            // flag): printables (native insertion continues below) and pure
            // movement keys (native caret nudge is harmless). RETURN/ESCAPE/
            // TAB/BACKSPACE/DELETE keep their dedicated flows below.
            let skip_dup = (32..=126).contains(&nk)
                || matches!(
                    nk,
                    LEFT | RIGHT | UP | DOWN | HOME | END | PAGE_UP | PAGE_DOWN
                );
            if skip_dup && s.entry_processed_key.get() {
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

            // Navigation/edit keys are also observed by the entry's own
            // handler, which skips when this handler arms entry_processed_key
            // below — see the arming comment for the full protocol.
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
        // Armed for every key the entry also observes (printables plus the
        // navigation/edit keys in its explicit arm): without the
        // non-printable half, one arrow press moves twice whenever
        // GDK_EVENT_STOP fails to propagate (Left from A1 lands two columns
        // deep instead of [A1).
        // Gated on entry_seen AND live focus: arming when the entry cannot
        // observe would linger uncleared and wrongly skip a later genuine
        // keypress there.
        let entry_also_observes = (32..=126).contains(&nk)
            || matches!(
                nk,
                RETURN | ESCAPE | TAB | LEFT | RIGHT | UP | DOWN | HOME | END | PAGE_UP | PAGE_DOWN | BACKSPACE | DELETE
            );
        if entry_also_observes
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
                    // Same-event duplicate guard (mirrors the printable arm
                    // below): the window handler fires first and arms
                    // entry_processed_key; without this check a second
                    // handle_key call moves twice (Left from A1 lands two
                    // columns deep instead of [A1) whenever GDK_EVENT_STOP
                    // fails to propagate between the layers.
                    if shared_k.entry_processed_key.get() {
                        shared_k.entry_processed_key.set(false);
                        return false;
                    }
                    // `state` carries the modifier mask (bit 0x1 = Shift on
                    // both GDK and the nwg adapter); Shift+arrows extend the
                    // selection via handle_key.
                    handle_key(keyval, &shared_k, state);
                    // Arm for the reverse order (entry first, window second);
                    // the window's guard above consumes it. Unreachable when
                    // the window fired first (skipped above).
                    if !cfg!(windows) && shared_k.entry_seen.get() {
                        shared_k.entry_processed_key.set(true);
                    }
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
    // Sheet tabs sit below the grid (above the status line), like the
    // terminal reference rendering tabs in its bottom row.
    vbox.append(&tabbar);
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
            // Keep the scrollbar thumb on the viewport (ranges track grid
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

#[cfg(test)]
mod gutter_tests {
    use super::*;
    use rswidgets::backends::headless::{DrawOp, RecordingDrawContext};
    use std::collections::HashMap;

    fn styled_texts(dc: &RecordingDrawContext) -> Vec<(String, i32, (f64, f64, f64, f64))> {
        dc.ops
            .iter()
            .filter_map(|op| match op {
                DrawOp::StyledText { text, weight, rgba, .. } => {
                    Some((text.clone(), *weight, *rgba))
                }
                _ => None,
            })
            .collect()
    }

    /// Row gutter labels must paint bold (weight 1), with the row's label.
    /// Short labels (<=2 chars) additionally record a padlock hit rect.
    #[test]
    fn row_headers_paint_bold() {
        use std::collections::BTreeSet;
        let mut dc = RecordingDrawContext::new();
        // Logical rows: main row 0 then footer rows (mr=1).
        let mut hits = Vec::new();
        paint_row_headers(&mut dc, &[HEADER_ROWS, HEADER_ROWS + 1], 1, &BTreeSet::new(), &mut hits);
        let texts = styled_texts(&dc);
        assert_eq!(texts.len(), 2, "one label per row, got {texts:?}");
        assert_eq!(texts[0].0, "1", "main row label text, got {:?}", texts[0].0);
        for (text, weight, _) in &texts {
            assert_eq!(
                *weight, 1,
                "gutter label {text:?} must be bold (weight 1), got weight {weight}"
            );
        }
        // Both labels ("1", "_1") are short: two padlock hits recorded.
        assert_eq!(hits.len(), 2, "short row labels need padlocks, got {hits:?}");
        assert!(hits.iter().all(|h| h.is_row && !h.locked));
        assert_eq!((hits[0].x, hits[0].w), (2.0, PADLOCK_W));
        // Row numbers sit 6px off the gutter's right gridline (headless
        // measures "1" at 8px wide: 50 - 8 - 6 = 36), never touching it.
        let xs: Vec<f64> = dc
            .ops
            .iter()
            .filter_map(|op| match op {
                DrawOp::StyledText { text, x, .. } if text == "1" => Some(*x),
                _ => None,
            })
            .collect();
        assert_eq!(xs, vec![ROW_LABEL_W - 8.0 - 6.0], "row number x, got {xs:?}");
    }

    /// Column gutter labels must paint bold (weight 1), with the column name.
    #[test]
    fn col_headers_paint_bold() {
        use std::collections::BTreeSet;
        let mut dc = RecordingDrawContext::new();
        // Margin col, main col A (mc=1).
        let col_ixs = vec![MARGIN_COLS - 1, MARGIN_COLS];
        let col_widths: HashMap<usize, usize> =
            col_ixs.iter().map(|&c| (c, 8)).collect();
        let mut hits = Vec::new();
        paint_col_headers(&mut dc, &col_ixs, &col_widths, 1, &BTreeSet::new(), &mut hits);
        let texts = styled_texts(&dc);
        assert_eq!(texts.len(), 2, "one label per column, got {texts:?}");
        assert_eq!(texts[1].0, "A", "main column label text, got {:?}", texts[1].0);
        for (text, weight, _) in &texts {
            assert_eq!(
                *weight, 1,
                "gutter label {text:?} must be bold (weight 1), got weight {weight}"
            );
        }
        // "[A" and "A" are short: two padlock hits recorded.
        assert_eq!(hits.len(), 2, "short col labels need padlocks, got {hits:?}");
        assert!(hits.iter().all(|h| !h.is_row && !h.locked));
    }

    /// Padlock eligibility: short (dual-character or less), non-empty gutter
    /// labels get the affordance; longer content does not.
    #[test]
    fn padlock_eligibility_rule() {
        assert!(!wants_padlock(""));
        assert!(wants_padlock("A"));
        assert!(wants_padlock("1"));
        assert!(wants_padlock("AB"));
        assert!(wants_padlock("[A"));
        assert!(wants_padlock("]B"));
        assert!(wants_padlock("~1"));
        assert!(wants_padlock("_1"));
        assert!(wants_padlock("10"));
        assert!(!wants_padlock("ABC"));
        assert!(!wants_padlock("_12"));
        assert!(!wants_padlock("Hello"));
    }

    /// Pinned rows/cols merge ahead of the display window (frozen first);
    /// entries already displayed are not duplicated.
    #[test]
    fn pinned_union_freezes_first() {
        assert_eq!(union_pinned(&[10, 11, 12], &[]), vec![10, 11, 12]);
        assert_eq!(union_pinned(&[10, 11, 12], &[1]), vec![1, 10, 11, 12]);
        assert_eq!(union_pinned(&[10, 11, 12], &[11]), vec![11, 10, 12]);
        assert_eq!(union_pinned(&[], &[3]), vec![3]);
    }

    /// Pin toggle flips membership and reports the end state.
    #[test]
    fn pin_toggle_flips() {
        use std::collections::BTreeSet;
        let mut set = BTreeSet::new();
        assert!(toggle_pin_in(&mut set, 7));
        assert!(set.contains(&7));
        assert!(!toggle_pin_in(&mut set, 7));
        assert!(!set.contains(&7));
    }

    /// Locked padlocks paint filled bodies in the dark slate; unlocked paint
    /// outlines in the lighter slate. Exact op assertions (geometry + color).
    #[test]
    fn padlock_paint_states_differ() {
        // Locked: filled body + connected shackle, all dark slate.
        let mut locked_dc = RecordingDrawContext::new();
        paint_padlock(&mut locked_dc, 0.0, 0.0, true);
        let fills: Vec<_> = locked_dc
            .ops
            .iter()
            .filter_map(|op| match op {
                DrawOp::FillRect { x, y, w, h, rgba } => Some((*x, *y, *w, *h, *rgba)),
                _ => None,
            })
            .collect();
        assert!(
            fills.iter().any(|&(x, y, w, h, c)| (x, y, w, h) == (1.0, 6.0, 8.0, 5.0)
                && c == (0.2, 0.2, 0.2, 1.0)),
            "locked padlock needs a filled dark body, got {fills:?}"
        );
        // Unlocked: stroked (outline) body in lighter slate, never filled.
        let mut open_dc = RecordingDrawContext::new();
        paint_padlock(&mut open_dc, 0.0, 0.0, false);
        let open_fills: Vec<_> = open_dc
            .ops
            .iter()
            .filter_map(|op| match op {
                DrawOp::FillRect { x, y, w, h, rgba } => Some((*x, *y, *w, *h, *rgba)),
                _ => None,
            })
            .collect();
        assert!(
            !open_fills.iter().any(|&(x, y, w, h, _)| (x, y, w, h) == (1.0, 6.0, 8.0, 5.0)),
            "unlocked padlock body must be outline-only, got fills {open_fills:?}"
        );
        let strokes: Vec<_> = open_dc
            .ops
            .iter()
            .filter_map(|op| match op {
                DrawOp::StrokeRect { x, y, w, h, rgba, .. } => Some((*x, *y, *w, *h, *rgba)),
                _ => None,
            })
            .collect();
        assert!(
            strokes.iter().any(|&(x, y, w, h, c)| (x, y, w, h) == (1.0, 6.0, 8.0, 5.0)
                && c == (0.45, 0.45, 0.45, 1.0)),
            "unlocked padlock needs an outlined lighter body, got {strokes:?}"
        );
    }
}

#[cfg(test)]
mod prompt_tests {
    use super::*;

    /// Regression for "Rename Sheet opens Find": every prompt-gated action
    /// gets its own correctly labeled dialog chrome — titles and buttons
    /// pinned here, not eyeballed in screenshots.
    #[test]
    fn prompt_chrome_labels_every_action() {
        let cases: &[(&str, &str, &str)] = &[
            ("rename_sheet", "Rename sheet", "Rename"),
            ("copy_sheet", "Copy sheet", "Copy"),
            ("delete_sheet", "Delete sheet", "Delete"),
            ("go_to_cell", "Go to cell", "Go"),
            ("set_col_width", "Column width", "Set"),
            ("set_max_col_width", "Default width", "Set"),
            ("find", "Find", "Find"),
            ("insert_special_chars", "Insert special char", "Insert"),
            ("insert_hyperlink", "Insert hyperlink", "Insert"),
            ("sort_view", "Sort view", "Sort"),
            ("persist_sort", "Persist sort", "Sort"),
            ("balance_books", "Balance books", "Balance"),
        ];
        for (action, title, ok) in cases {
            let (t, o, initial) = prompt_chrome(action, "Sheet2");
            assert_eq!(&t, title, "dialog title for {action}");
            assert_eq!(&o, ok, "dialog button for {action}");
            if *action != "find" {
                assert_ne!(&t, "Find", "{action} must not recycle the Find dialog");
            }
            let _ = initial;
        }
    }

    /// Rename pre-fills the entry with the current title (edit-in-place);
    /// creation-style prompts start empty so stale text can never leak in.
    #[test]
    fn prompt_chrome_initial_text() {
        assert_eq!(
            prompt_chrome("rename_sheet", "Budget"),
            ("Rename sheet".into(), "Rename".into(), "Budget".into())
        );
        for action in ["copy_sheet", "delete_sheet", "go_to_cell", "find"] {
            assert_eq!(
                prompt_chrome(action, "Budget").2,
                String::new(),
                "{action} must start with an empty entry"
            );
        }
    }
}

#[cfg(test)]
mod tab_tests {
    use super::*;

    /// Tabs lay out left to right from x=2, padded by TAB_PAD_X on both
    /// sides with TAB_GAP between. "Sheet1" at the stub 8px/char measures
    /// 48px, so tab 1 spans 2..70 and tab 2 starts at 76.
    #[test]
    fn tab_layout_pads_and_gaps_tabs() {
        let titles = vec!["Sheet1".to_string(), "Sheet2".to_string()];
        let hits = tab_layout(&titles, 1, &|t: &str, _w: i32| t.chars().count() as f64 * 8.0);
        assert_eq!(hits.len(), 2, "one hit per sheet, got {hits:?}");
        assert_eq!((hits[0].x0, hits[0].x1), (2.0, 70.0));
        assert_eq!(hits[0].index, 0);
        assert_eq!((hits[1].x0, hits[1].x1), (76.0, 144.0));
        assert_eq!(hits[1].index, 1);
    }

    /// No titles, no hits (single-sheet workbooks hide the bar anyway).
    #[test]
    fn tab_layout_empty_without_titles() {
        assert!(tab_layout(&[], 0, &|_: &str, _: i32| 0.0).is_empty());
        assert!(tab_layout(&["Only".to_string()], 0, &|_: &str, _: i32| 0.0).len() == 1);
    }

    /// The active tab measures bold (weight 1): the measure closure must
    /// observe weight 1 exactly for the active index and 0 elsewhere, or
    /// bold titles would misalign their hit rects.
    #[test]
    fn tab_layout_measures_active_bold() {
        let titles = vec!["A".to_string(), "B".to_string(), "C".to_string()];
        let seen = std::cell::RefCell::new(Vec::new());
        let hits = tab_layout(&titles, 2, &|t: &str, w: i32| {
            seen.borrow_mut().push((t.to_string(), w));
            10.0
        });
        assert_eq!(seen.borrow().clone(), vec![("A".into(), 0), ("B".into(), 0), ("C".into(), 1)]);
        assert_eq!(hits.len(), 3);
        // Each tab is 10 + 10 + 10 = 30 wide with 6px gaps: tab 0 spans
        // 2..32, tab 1 spans 38..68, tab 2 spans 74..104.
        assert_eq!((hits[2].x0, hits[2].x1), (74.0, 104.0));
    }
}

#[cfg(test)]
mod brightness_tests {
    use super::*;
    use rswidgets::backends::headless::RecordingDrawContext;
    use std::collections::HashMap;

    fn two_by_two_grid() -> GridBox {
        crate::grid::Grid::new(2, 2).into()
    }

    /// Margin-zone data cells must render at 75% background brightness
    /// (0.75 gray) while main cells stay white. Regression: every data cell
    /// painted white, so margins were indistinguishable from body content.
    #[test]
    fn margin_cells_render_dimmer_than_main_cells() {
        let grid = two_by_two_grid();
        let hr = HEADER_ROWS;
        let lm = MARGIN_COLS;
        // One main row, two columns: left-margin col then main col A.
        let display_rows = vec![hr];
        let col_ixs = vec![lm - 1, lm];
        let col_widths: HashMap<usize, usize> =
            col_ixs.iter().map(|&c| (c, sheet_rec_col_width_for_test(&grid, c))).collect();
        let row_agg = compute::compute_row_agg_func(&grid, &display_rows, hr, 2);
        let mut sink = GuiCanvasSink::new();
        render::fill_cells(
            &mut sink, &display_rows, &col_ixs, &col_widths, &grid,
            hr, 2, 2, lm, col_ixs.len(), hr, lm, &row_agg,
        );
        let mut dc = RecordingDrawContext::new();
        // Cursor parked off the displayed row so neither sample cell takes
        // the cursor highlight (which would mask the base backgrounds).
        render_to(
            &sink, &mut dc, &col_ixs, &col_widths, &display_rows,
            hr, 2, 2, lm, hr + 1, lm, false, "", None,
        );
        // Cell background rects in the grid band, left to right.
        let mut bgs: Vec<((f64, f64), (f64, f64, f64, f64))> = dc
            .fill_rects()
            .into_iter()
            .filter(|r| r.y >= HEADER_H - 0.5 && r.y < HEADER_H + ROW_H)
            .map(|r| ((r.x, r.w), r.rgba))
            .collect();
        bgs.sort_by(|a, b| a.0 .0.partial_cmp(&b.0 .0).unwrap());
        assert!(
            bgs.len() >= 2,
            "expected margin + main background rects, got {bgs:?}"
        );
        let margin_bg = bgs[0].1;
        let main_bg = bgs[1].1;
        assert!(
            (margin_bg.0 - 0.75).abs() < 0.01
                && (margin_bg.1 - 0.75).abs() < 0.01
                && (margin_bg.2 - 0.75).abs() < 0.01,
            "margin cell background must be 75% brightness, got {margin_bg:?}"
        );
        assert_eq!(
            main_bg,
            (1.0, 1.0, 1.0, 1.0),
            "main cell background must stay white, got {main_bg:?}"
        );
    }

    /// Test-only width lookup mirroring sheet_rec_col_width without an App.
    fn sheet_rec_col_width_for_test(grid: &GridBox, col: usize) -> usize {
        grid.col_width(col).max(1)
    }
}

#[cfg(test)]
mod trailing_tests {
    use super::*;

    /// grow_for_trailing_blank keeps exactly one blank past the last
    /// non-blank index (content or cursor).
    #[test]
    fn trailing_blank_rule() {
        // Content fills the extent, cursor inside: grow (no blank at all).
        assert!(grow_for_trailing_blank(2, 0, Some(0)));
        // One blank already, cursor inside content: hold.
        assert!(!grow_for_trailing_blank(3, 1, Some(0)));
        // Cursor parked on the last (blank) row: it counts as non-blank,
        // so one more opens beyond it.
        assert!(grow_for_trailing_blank(3, 1, Some(2)));
        // Cursor beyond the blank (can happen transiently): still grow, the
        // blank must sit past the cursor, not just past content.
        assert!(grow_for_trailing_blank(3, 1, Some(3)));
        // Empty extent with cursor on row 0: grow the first blank.
        assert!(grow_for_trailing_blank(1, 1, Some(0)));
        // Cursor out in the margins (None): content rule alone; all blank
        // with no cursor anchor needs nothing.
        assert!(!grow_for_trailing_blank(2, 2, None));
        // Content on the last row, cursor elsewhere inside: grow.
        assert!(grow_for_trailing_blank(2, 0, Some(1)));
        // Two blanks, cursor inside: hold (keyboard keeps two; the GUI
        // floor is one, and one already exceeds it).
        assert!(!grow_for_trailing_blank(4, 2, Some(0)));
    }
}

#[cfg(test)]
mod formula_tests {
    use super::*;

    fn empty_grid() -> GridBox {
        crate::grid::Grid::new(1, 1).into()
    }

    /// The formula-bar address label must render user-facing cell names for
    /// every zone (matching ratatui's addr_label/cell_ref_text), never
    /// CellAddr's internal rendering (`(0, 0)`, `<701>(0)`, ...).
    #[test]
    fn addr_label_main_cell() {
        let grid = empty_grid();
        assert_eq!(formula_addr_label(HEADER_ROWS, MARGIN_COLS, &grid), "A1");
    }

    #[test]
    fn addr_label_left_margin() {
        let grid = empty_grid();
        assert_eq!(
            formula_addr_label(HEADER_ROWS, MARGIN_COLS - 1, &grid),
            "[A1"
        );
    }

    #[test]
    fn addr_label_right_margin() {
        let grid = empty_grid();
        // First margin col addresses (and saves) as margin, same as
        // ratatui: ]A1, then ]B1 past it.
        assert_eq!(
            formula_addr_label(HEADER_ROWS, MARGIN_COLS + 1, &grid),
            "]A1"
        );
        assert_eq!(
            formula_addr_label(HEADER_ROWS, MARGIN_COLS + 2, &grid),
            "]B1"
        );
    }

    #[test]
    fn addr_label_header_row() {
        let grid = empty_grid();
        assert_eq!(
            formula_addr_label(HEADER_ROWS - 1, MARGIN_COLS, &grid),
            "A~1"
        );
    }

    #[test]
    fn addr_label_footer_row() {
        let grid = empty_grid();
        // First margin row addresses (and saves) as footer, same as
        // ratatui: A_1, then A_2 past it.
        assert_eq!(
            formula_addr_label(HEADER_ROWS + 1, MARGIN_COLS, &grid),
            "A_1"
        );
        assert_eq!(
            formula_addr_label(HEADER_ROWS + 2, MARGIN_COLS, &grid),
            "A_2"
        );
    }

    /// First margin row/col addresses as margin everywhere (formula,
    /// gutters, commits): _1 / ]A, same as ratatui. The trailing ring
    /// renders white but keeps margin names and margin storage.
    #[test]
    fn first_margin_row_col_address_as_margin() {
        let grid = empty_grid(); // 1x1: first margin is row hr+1 / col lm+1
        // Formula bar.
        assert_eq!(
            formula_addr_label(HEADER_ROWS + 1, MARGIN_COLS + 1, &grid),
            "]A_1"
        );
        // Gutter labels (shared ui fns, same as ratatui).
        assert_eq!(crate::addr::ui_row_label(HEADER_ROWS + 1, 1), "_1");
        assert_eq!(crate::addr::ui_row_label(HEADER_ROWS + 2, 1), "_2");
        assert_eq!(crate::addr::ui_column_fragment(MARGIN_COLS + 1, 1), "]A");
        assert_eq!(crate::addr::ui_column_fragment(MARGIN_COLS + 2, 1), "]B");
        // Ring predicate itself (body intersections only, not margins).
        let (hr, lm) = (HEADER_ROWS, MARGIN_COLS);
        assert!(is_trailing_blank_cell(hr + 1, lm, hr, 1, lm, 1));
        assert!(is_trailing_blank_cell(hr, lm + 1, hr, 1, lm, 1));
        assert!(is_trailing_blank_cell(hr + 1, lm + 1, hr, 1, lm, 1));
        assert!(!is_trailing_blank_cell(hr + 2, lm, hr, 1, lm, 1));
        assert!(!is_trailing_blank_cell(hr, lm + 2, hr, 1, lm, 1));
        assert!(!is_trailing_blank_cell(hr + 1, lm - 1, hr, 1, lm, 1));
    }
}

#[cfg(test)]
mod fill_tests {
    use super::*;
    use std::path::PathBuf;

    fn overflow_app() -> crate::gui::App {
        let mut dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
        dir.push("docs/tests/overflow.corro");
        let mut app = crate::gui::App::new_with_paths(vec![dir]);
        app.load_initial().unwrap();
        app
    }

    /// cols_to_fill_px must size the viewport to cover the available width
    /// at EVERY margin depth 1..=10, not just one spot. Regression: with a
    /// margin cursor, visible_col_indices returns dim-1 columns, so the
    /// `cols.len() < dim` exit fired after two iterations (data_cols=9) and
    /// the sheet stopped ~460px into a 1280px window, leaving a huge blank.
    #[test]
    fn cols_to_fill_px_covers_viewport_with_margin_cursor() {
        let app = overflow_app();
        for depth in 1..=10usize {
            let cursor = SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS - depth };
            let dim = cols_to_fill_px(&app, cursor, 1220);
            let sheet = app.core.workbook.active_sheet();
            let mc = sheet.grid.main_cols();
            let (cols, _) = ui_core::visible_col_indices(sheet, cursor, dim, 0);
            // Same display widths the renderer uses (padlock columns are one
            // char wider); recomputing with recorded widths would undercount.
            let used_px: usize = cols
                .iter()
                .map(|&c| display_col_width(sheet, c, mc))
                .sum::<usize>()
                * CHAR_W as usize;
            assert!(
                used_px >= 1220,
                "viewport must cover 1220px at margin depth {depth} (dim={dim}, cols={}, used~{used_px})",
                cols.len()
            );
        }
    }

    /// display_col_width spares exactly one extra character for padlock
    /// columns (short gutter labels) and none otherwise — the +1 the
    /// renderer relies on so the icon fits inside its own column.
    #[test]
    fn display_col_width_spares_a_char_for_padlock_columns() {
        let app = overflow_app();
        let sheet = app.core.workbook.active_sheet();
        let mc = sheet.grid.main_cols();
        let mut saw_pad = false;
        let mut saw_plain = false;
        // Deep margin (long labels) and the A1 neighborhood (short labels).
        for c in (0..8).chain(MARGIN_COLS - 4..MARGIN_COLS + 8) {
            let label = crate::addr::ui_column_fragment(c, mc);
            let extra = display_col_width(sheet, c, mc) - sheet_rec_col_width(sheet, c);
            assert_eq!(
                extra,
                usize::from(wants_padlock(&label)),
                "col {c} ({label:?}): extra width must match padlock eligibility"
            );
            saw_pad |= extra == 1;
            saw_plain |= extra == 0;
        }
        assert!(saw_pad, "expected padlock columns near A1");
        assert!(saw_plain, "expected plain columns in the deep margin");
    }

    /// Every padlock painted with production widths must end inside its own
    /// column (no spill over the neighbor's gridline). Regression: the icon
    /// hung past narrow margin columns.
    #[test]
    fn padlocks_fit_inside_their_widened_columns() {
        use rswidgets::backends::headless::RecordingDrawContext;
        use std::collections::BTreeSet;
        let app = overflow_app();
        let sheet = app.core.workbook.active_sheet();
        let mc = sheet.grid.main_cols();
        let cursor = SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS };
        let (cols, _) = ui_core::visible_col_indices(sheet, cursor, 12, 0);
        let widths: HashMap<usize, usize> = cols
            .iter()
            .map(|&c| (c, display_col_width(sheet, c, mc)))
            .collect();
        let mut dc = RecordingDrawContext::new();
        let mut hits = Vec::new();
        paint_col_headers(&mut dc, &cols, &widths, mc, &BTreeSet::new(), &mut hits);
        assert!(!hits.is_empty(), "expected padlock columns in the A1 view");
        // Replay the paint's x-accumulation; each padlock's right edge must
        // not pass its own column's right edge.
        let mut cx = ROW_LABEL_W;
        let mut hi = 0usize;
        for &c in &cols {
            let cw = *widths.get(&c).unwrap() as f64 * CHAR_W;
            if wants_padlock(&crate::addr::ui_column_fragment(c, mc)) {
                let h = &hits[hi];
                hi += 1;
                assert!(
                    h.x + h.w <= cx + cw + 1e-9,
                    "padlock for col {c} spills past its column: right {} vs edge {}",
                    h.x + h.w,
                    cx + cw
                );
            }
            cx += cw;
        }
        assert_eq!(hi, hits.len(), "hit/column mapping drifted");
    }

    /// Same guarantee with the cursor on A1 (the pre-existing passing case,
    /// pinned so the fix cannot regress the common path).
    #[test]
    fn cols_to_fill_px_covers_viewport_on_main_cursor() {
        let app = overflow_app();
        let cursor = SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS };
        let dim = cols_to_fill_px(&app, cursor, 1220);
        let sheet = app.core.workbook.active_sheet();
        let mc = sheet.grid.main_cols();
        let (cols, _) = ui_core::visible_col_indices(sheet, cursor, dim, 0);
        // Same display widths the renderer uses (see margin-cursor case).
        let used_px: usize = cols
            .iter()
            .map(|&c| display_col_width(sheet, c, mc))
            .sum::<usize>()
            * CHAR_W as usize;
        assert!(
            used_px >= 1220,
            "viewport must cover 1220px on A1 (dim={dim}, cols={}, used~{used_px})",
            cols.len()
        );
    }
}
