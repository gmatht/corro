// Combined-gui builds flip the rswidgets root prelude to pancurses-adapter
// types; the native backend always needs the common wrappers, so on Linux
// it names them explicitly. Other platforms keep the prelude (unchanged).
#[cfg(not(any(target_os = "linux", target_os = "windows", target_os = "ios", target_os = "android")))]
use rswidgets::prelude::*;
#[cfg(target_os = "linux")]
use rswidgets::common::{Canvas, Entry, Label, MenuBar, Orientation, Window};
// iOS, Android and macOS name the common wrappers explicitly for the same
// reason Windows does: their prelude is the platform adapter's, and mixing the
// two would silently rebind `Window`/`Canvas`/... to adapter-local handles.
#[cfg(any(target_os = "ios", target_os = "android", target_os = "macos"))]
use rswidgets::common::{Canvas, Entry, Label, MenuBar, Orientation, Window};
// Windows uses the same common wrappers explicitly: under pancurses the
// root prelude flips to pancurses-adapter types, so the glob alone would
// silently rebind these names there.
#[cfg(target_os = "windows")]
use rswidgets::common::{Canvas, Entry, Label, MenuBar, Orientation, Window};
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

use rswidgets::core::key::{normalize, RETURN, ESCAPE, BACKSPACE, DELETE, LEFT, UP, RIGHT, DOWN, TAB, HOME, END, PAGE_UP, PAGE_DOWN, F1, F2, F3, ALT_L, ALT_R};

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
//
// These are *desktop* pixel metrics: 12px text on a 20px row. That is the
// right size on a monitor at 96dpi, but the Android backend draws into raw
// device pixels, where a 420dpi phone renders 12px at ~4.6dp — about a third
// of Android's 14sp body-text floor, with rows far under the 48dp touch
// target.
//
// `metrics_scale()` therefore multiplies every metric by the host's density
// factor on Android (2.625x on a 420dpi device, giving ~31px text and ~52px
// rows) and is exactly 1.0 everywhere else, so desktop GTK/nwg/pancurses and
// the movie capture are untouched. The metrics are functions rather than
// consts so there is a single place to apply that scale.
// ---------------------------------------------------------------------------

/// Base (desktop) metrics, before any host density scaling.
pub(crate) const FONT_SIZE_BASE: f64 = 12.0;
pub(crate) const ROW_H_BASE: f64 = 20.0;
pub(crate) const HEADER_H_BASE: f64 = 24.0;
pub(crate) const ROW_LABEL_W_BASE: f64 = 50.0;
pub(crate) const CHAR_W_BASE: f64 = 7.2;

/// Density multiplier for the pixel metrics.
///
/// 1.0 on every desktop backend. On Android it is the display density
/// (`DisplayMetrics.density`), so the grid is legible on a phone without
/// changing a single desktop call site. Resolved once and cached: it is read
/// on every frame and cannot change while the app runs.
#[cfg(target_os = "android")]
pub(crate) fn metrics_scale() -> f64 {
    use std::sync::OnceLock;
    static SCALE: OnceLock<f64> = OnceLock::new();
    *SCALE.get_or_init(|| {
        rswidgets::backends_android_adapter::display_density()
            .unwrap_or(1.0)
            .max(1.0)
    })
}

/// iOS: `UIScreen.scale` (1.0 non-Retina, 2.0 Retina, 3.0 Plus/X era). Same
/// purpose as the Android density above — a phone needs bigger pixels than a
/// desktop monitor for the same number to be legible — and resolved once
/// (it cannot change while the app runs).
#[cfg(target_os = "ios")]
pub(crate) fn metrics_scale() -> f64 {
    use std::sync::OnceLock;
    static SCALE: OnceLock<f64> = OnceLock::new();
    *SCALE.get_or_init(|| {
        rswidgets::backends_ios_adapter::display_density()
            .unwrap_or(1.0)
            .max(1.0)
    })
}

/// macOS: `NSScreen.backingScaleFactor` (1.0 non-Retina, 2.0 Retina). Same
/// purpose as the iOS/Android density — a Retina panel needs bigger *points*
/// for the same nominal pixel number to be legible — and resolved once (it
/// cannot change while the app runs).
///
/// Note there is no macOS touch-target floor: the AppKit adapter uses compact
/// desktop control metrics, so this scale only affects corro's own chrome.
#[cfg(target_os = "macos")]
pub(crate) fn metrics_scale() -> f64 {
    use std::sync::OnceLock;
    static SCALE: OnceLock<f64> = OnceLock::new();
    *SCALE.get_or_init(|| {
        rswidgets::backends_macos_adapter::display_density()
            .unwrap_or(1.0)
            .max(1.0)
    })
}

#[cfg(not(any(target_os = "android", target_os = "ios", target_os = "macos")))]
pub(crate) fn metrics_scale() -> f64 {
    1.0
}

/// Startup phase marker for the mobile backends.
///
/// A panic inside the host's `extern "C"` entry point aborts the process, and
/// the release build's optimisations can leave the backtrace unreadable — so on
/// iOS/Android the only reliable way to locate a startup failure is to say how
/// far the GUI got. On every desktop backend this compiles to nothing, so no
/// existing behaviour or output changes.
#[inline(always)]
pub(crate) fn phase(marker: &str) {
    // Android and iOS need the same startup breadcrumb, but the loggers are
    // different modules and each exists only on its own target: `log_ios`
    // lives behind `cfg(target_os = "ios")`, so calling it unconditionally
    // from an android|ios gate failed to compile on Android. Route each to
    // its own sink - logcat on Android (via the helper below), the UIKit
    // logger on iOS.
    #[cfg(target_os = "android")]
    {
        super::android_backend::logcat(marker);
    }
    #[cfg(target_os = "ios")]
    {
        rswidgets::backends::ios::log_ios(marker);
    }
    #[cfg(target_os = "macos")]
    {
        rswidgets::backends::macos::log_macos(marker);
    }
    #[cfg(not(any(target_os = "android", target_os = "ios", target_os = "macos")))]
    {
        let _ = marker;
    }
}

pub(crate) fn font_size() -> f64 { FONT_SIZE_BASE * metrics_scale() }
pub(crate) fn row_h() -> f64 { ROW_H_BASE * metrics_scale() }
pub(crate) fn header_h() -> f64 { HEADER_H_BASE * metrics_scale() }
pub(crate) fn row_label_w() -> f64 { ROW_LABEL_W_BASE * metrics_scale() }
pub(crate) fn char_w() -> f64 { CHAR_W_BASE * metrics_scale() }

/// Grid metrics for the Android touch-scroll path.
///
/// `SheetView` converts a pixel drag into whole rows/columns, so it needs the
/// same numbers the renderer used — a constant hardcoded in Java would drift
/// from `row_h()`/`char_w()` as soon as the density or font metrics change,
/// and scrolling would scale wrong by exactly that factor.
#[cfg(target_os = "android")]
pub(crate) fn touch_row_h() -> f64 { row_h() }

/// Default column advance in pixels (the width of one character cell; a grid
/// column is a whole number of these). See [`touch_row_h`].
#[cfg(target_os = "android")]
pub(crate) fn touch_col_w() -> f64 { char_w() }

const MAX_RENDER_ROWS: usize = 500;
const MAX_RENDER_COLS: usize = 50;

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

pub(crate) struct GuiState {
    app: *mut super::App,
    pub(crate) rxapp: rswidgets::App,
    pub(crate) window: Window,
    /// Menu bar handle: lets modal flows (e.g. the special-char picker)
    /// dismiss an open menu grab that would otherwise swallow all later
    /// keys (an open keyboard menu routes everything past the window and
    /// entry handlers, so a staged edit could never commit). Set once
    /// after build_menu (which needs the state for callbacks).
    menubar: std::cell::OnceCell<MenuBar>,
    canvas: Canvas,
    formula_entry: Entry,
    addr_label: Label,
    /// Bottom strip: the shared hints line (ratatui parity — ratatui's
    /// bottom row always shows hints, never status). Renamed from
    /// status_label when status moved to the formula row.
    hints_label: Label,
    /// Formula-row `· status` suffix (ratatui parity): a SEPARATE label
    /// after the entry, never entry text, so status can never leak into
    /// the edit buffer. Visible only when status is non-empty and no
    /// edit/dialog owns the formula row (ratatui's Edit/input arms show
    /// the buffer with no status span).
    formula_status: Label,
    editing: Cell<bool>,
    edit_buf: RefCell<String>,
    /// Caret position within `edit_buf`, as a **character** index. The native
    /// Entry exposes no caret API, so the host owns the caret the same way
    /// ratatui's `Mode::Edit` does: typed chars insert here, Backspace/Delete
    /// act here, and Left/Right move here (committing and moving the *cell*
    /// cursor only at the buffer edges).
    edit_caret: Cell<usize>,
    /// In-grid aggregate dropdown state (painted on the canvas, not a native
    /// widget — see `AggDrop`).
    agg_drop: RefCell<Option<AggDrop>>,
    /// Last canvas size seen by the draw callback, for hit-testing the
    /// canvas-painted dropdown.
    canvas_size: Cell<(i32, i32)>,
    /// Formula text displayed in the entry the last time the formula bar
    /// was refreshed (the adopt source for click-to-edit repair). Refreshed
    /// on every bar update, so it can never go stale.
    entry_snapshot: RefCell<String>,
    /// Set by a pointer click into the formula entry, consumed by the next
    /// entry keystroke (repair) or cleared by navigation/commit/cancel
    /// (then typing starts a fresh value, ratatui parity).
    entry_clicked: Cell<bool>,
    /// Synthetic pointer for `--movie`: `(x, y, pressed)` in canvas pixels.
    ///
    /// X11 does not composite the cursor into an `x11grab` capture (verified: a
    /// moved pointer changes no pixels in the grab), so a recording cannot show
    /// the real mouse. The movie driver instead animates this position and the
    /// canvas paints it, which is what makes a menu tour read as a pointer
    /// travelling to an item and clicking it. `None` means nothing is drawn, so
    /// ordinary interactive use is untouched.
    movie_pointer: Cell<Option<(f64, f64, bool)>>,
    /// Cell address last shown in the formula bar. A bar refresh for a
    /// different cell ends click-edit intent (navigation resets to
    /// ratatui-style replace for the next keystroke).
    entry_shown: Cell<(usize, usize)>,
    mode: Cell<GuiMode>,
    last_row: Cell<usize>,
    last_col: Cell<usize>,
    data_rows: Cell<usize>,
    data_cols: Cell<usize>,
    last_key: Cell<u32>,
    key_counter: Cell<u64>,
    /// Set when any pointer click has been handled. `present()` pumps the
    /// GTK main loop for several hundred iterations before the real main
    /// loop starts, so a click can arrive while `run_gui`'s setup is still
    /// running. The post-`present` `start_edit_keep_display` would then
    /// clobber that selection (moving the yellow edit highlight to the
    /// startup cursor and hiding any dropdown the click opened). Setup
    /// checks this flag and leaves the user's click alone.
    clicked: Cell<bool>,
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
    /// Viewport scroll offset (rows, cols) as a *display index*, not a cell
    /// address. `visible_row_indices`/`visible_col_indices` take it as
    /// `prev_start` and keep it unless the cursor would fall outside the
    /// window, so persisting it lets the viewport stay put during ordinary
    /// navigation and lets a click re-centre deliberately.
    ///
    /// `None` means "no anchor yet": the viewport is derived from the cursor
    /// alone, which is the original behaviour.
    viewport_anchor: Cell<Option<(usize, usize)>>,
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
        let ry = header_h() + ri as f64 * row_h();

        for (ci, &c) in col_ixs.iter().enumerate().take(MAX_RENDER_COLS) {
            let cw = *col_widths.get(&c).unwrap_or(&8) as f64 * char_w();
            let cx = row_label_w() + col_ixs.iter().take(ci).map(|&pc| *col_widths.get(&pc).unwrap_or(&8) as f64 * char_w()).sum::<f64>();

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
            } else if is_margin_cell(logical_row, c, hr, mr, lm, mc) {
                // Margin-zone cells render at 75% background brightness
                // so margins read as subordinate to body content —
                // including the first margin row/col (_1/]A), same gray
                // as the rest of the margin.
                (0.75, 0.75, 0.75, 1.0)
            } else {
                (1.0, 1.0, 1.0, 1.0)
            };

            dc.fill_rect(cx, ry, cw, row_h(), bg.0, bg.1, bg.2, bg.3);

            // Selection highlight
            if is_current && !is_editing {
                dc.stroke_rect(cx, ry, cw, row_h(), 0.0, 0.4, 0.8, 1.0, 2.0);
            }

            // Grid lines
            dc.stroke_rect(cx, ry, cw, row_h(), 0.8, 0.8, 0.8, 1.0, 0.5);

            if !raw_text.is_empty() {
                match style {
                    CellDisplayStyle::Default => {
                        dc.draw_text(cx + 2.0, ry + 2.0, raw_text, "monospace", font_size(), 0.0, 0.0, 0.0, 1.0);
                    }
                    CellDisplayStyle::Cursor | CellDisplayStyle::Selected => {
                        dc.draw_text(cx + 2.0, ry + 2.0, raw_text, "monospace", font_size(), 0.0, 0.0, 0.0, 1.0);
                    }
                    CellDisplayStyle::Aggregate | CellDisplayStyle::FooterAggregate => {
                        dc.draw_text(cx + 2.0, ry + 2.0, raw_text, "monospace", font_size(), 0.5, 0.5, 0.5, 1.0);
                    }
                    CellDisplayStyle::ActiveHeader | CellDisplayStyle::InactiveHeader => {
                        dc.draw_text(cx + 2.0, ry + 2.0, raw_text, "monospace", font_size(), 0.3, 0.3, 0.3, 1.0);
                    }
                    CellDisplayStyle::Hyperlink => {
                        // Hyperlinks render blue and underlined by default
                        // (same rule as the terminal backends); cursor and
                        // selection paints above already won for this cell.
                        dc.draw_text(cx + 2.0, ry + 2.0, raw_text, "monospace", font_size(), 0.0, 0.0, 0.9, 1.0);
                        let (tw, _, _, _) = dc.text_extents(raw_text, "monospace", font_size());
                        if tw > 0.0 {
                            dc.fill_rect(cx + 2.0, ry + 2.0 + font_size() + 1.0, tw, 1.0, 0.0, 0.0, 0.9, 1.0);
                        }
                    }
                }
            }
        }
    }

    // Edit overlay on cursor cell
    if is_editing && !edit_text.is_empty() {
        if let Some(pos) = col_ixs.iter().position(|&c| c == cursor_col) {
            let cw = *col_widths.get(&cursor_col).unwrap_or(&8) as f64 * char_w();
            let cx = row_label_w() + col_ixs.iter().take(pos).map(|&pc| *col_widths.get(&pc).unwrap_or(&8) as f64 * char_w()).sum::<f64>();
            if let Some(pos_r) = display_rows.iter().position(|&r| r == cursor_row) {
                let ry = header_h() + pos_r as f64 * row_h();
                dc.fill_rect(cx, ry, cw, row_h(), 1.0, 1.0, 0.8, 1.0);
                dc.draw_text(cx + 2.0, ry + 2.0, edit_text, "monospace", font_size(), 0.0, 0.0, 0.0, 1.0);
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
            .map(|&c| display_col_width(sheet, c, mc) as f64 * char_w())
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
    (((h as f64 - header_h() - 20.0) / row_h() + 1.0).max(1.0)) as usize
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
    hl_rows: Option<(usize, usize)>,
) {
    for (ri, &logical_row) in display_rows.iter().enumerate().take(MAX_RENDER_ROWS) {
        let ry = header_h() + ri as f64 * row_h();
        let label = crate::addr::ui_row_label(logical_row, mr);
        let (_, _, tw, _) = dc.text_extents_styled(&label, "monospace", font_size(), 0, 1);
        // Covered rows (anchor↔cursor selection) use the body selection
        // fill so the gutter mirrors the selected band; plain rows keep the
        // neutral header gray. With no selection nothing changes.
        let hl = hl_rows.is_some_and(|(r0, r1)| logical_row >= r0 && logical_row <= r1);
        let fill = if hl { (0.9, 0.95, 1.0, 1.0) } else { (0.9, 0.9, 0.9, 1.0) };
        dc.fill_rect(0.0, ry, row_label_w(), row_h(), fill.0, fill.1, fill.2, fill.3);
        // Row numbers sit 6px off the gutter's right gridline so glyphs
        // never touch it.
        dc.draw_text_styled(row_label_w() - tw - 6.0, ry + 2.0, &label, "monospace", font_size(), 0.3, 0.3, 0.3, 1.0, 0, 1);
        // Padlock at the gutter's left edge (short labels only): the label
        // is right-aligned, so the left side always has room.
        if wants_padlock(&label) {
            let locked = pinned.contains(&logical_row);
            let (px, py) = (2.0, ry + (row_h() - padlock_h()) / 2.0);
            paint_padlock(dc, px, py, locked);
            out_padlocks.push(GutterPadlock {
                x: px,
                y: py,
                w: padlock_w(),
                h: padlock_h(),
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
    hl_cols: Option<(usize, usize)>,
) {
    for (ci, &c) in col_ixs.iter().enumerate().take(MAX_RENDER_COLS) {
        let cw = *col_widths.get(&c).unwrap_or(&8) as f64 * char_w();
        let cx = row_label_w() + col_ixs.iter().take(ci).map(|&pc| *col_widths.get(&pc).unwrap_or(&8) as f64 * char_w()).sum::<f64>();
        let col_name = crate::addr::ui_column_fragment(c, mc);
        // Same selection fill as covered row headers (see paint_row_headers).
        let hl = hl_cols.is_some_and(|(c0, c1)| c >= c0 && c <= c1);
        let fill = if hl { (0.9, 0.95, 1.0, 1.0) } else { (0.9, 0.9, 0.9, 1.0) };
        dc.fill_rect(cx, 0.0, cw, header_h(), fill.0, fill.1, fill.2, fill.3);
        let (_, _, tw, _) = dc.text_extents_styled(&col_name, "monospace", font_size(), 0, 1);
        // Label and padlock center as a unit, so the icon fits inside its
        // own column instead of dangling past the gridline: the column is
        // one character wider than recorded (see display_col_width)
        // precisely to hold this group.
        let lock = wants_padlock(&col_name);
        let group = tw + if lock { 2.0 + padlock_w() } else { 0.0 };
        let tx = cx + (cw - group) / 2.0;
        dc.draw_text_styled(tx, (header_h() - font_size() * 1.2) / 2.0, &col_name, "monospace", font_size(), 0.3, 0.3, 0.3, 1.0, 0, 1);
        // Padlock right after the centered text, inside its own column: the
        // column is one character wider than recorded (see display_col_width)
        // precisely so this icon fits without spilling over the neighbor.
        if lock {
            let locked = pinned.contains(&c);
            let (px, py) = (tx + tw + 2.0, (header_h() - padlock_h()) / 2.0);
            paint_padlock(dc, px, py, locked);
            out_padlocks.push(GutterPadlock {
                x: px,
                y: py,
                w: padlock_w(),
                h: padlock_h(),
                is_row: false,
                index: c,
                locked,
            });
        }
    }
}

/// Padlock affordance geometry (base device px, scaled by the host density
/// like the row/header metrics): small enough for a 20px row and the 24px
/// header strip on a desktop, big enough to click and read there, and grown
/// proportionally on a phone where 10px would be a ~4dp speck.
const PADLOCK_W_BASE: f64 = 10.0;
const PADLOCK_H_BASE: f64 = 12.0;
/// Base glyph inset inside the padlock box (scaled with it).
const PADLOCK_INSET_BASE: f64 = 1.0;
/// Minimum hit target for a pin toggle, in dp. Android's guidance is 48dp;
/// the padlock is drawn inside a 20dp row/column header, so the *drawn* icon
/// stays row-sized and only the hit test is widened to this, which keeps the
/// tap usable without the icon overlapping neighbouring headers.
const PADLOCK_HIT_DP: f64 = 44.0;

pub(crate) fn padlock_w() -> f64 { PADLOCK_W_BASE * metrics_scale() }
pub(crate) fn padlock_h() -> f64 { PADLOCK_H_BASE * metrics_scale() }
fn padlock_inset() -> f64 { PADLOCK_INSET_BASE * metrics_scale() }
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

/// Hit rectangle for a painted padlock: the drawn box, widened on Android to
/// a finger-sized target centred on it.
///
/// The glyph must stay header-sized when drawn (a bigger icon would overlap
/// the neighbouring header and hide its label), so only the hit test grows.
/// Growth is clamped to half a header per axis, which keeps the target
/// inside the cell the padlock belongs to.
fn padlock_hit_rect(h: &GutterPadlock) -> (f64, f64, f64, f64) {
    let (mut w, mut hh) = (h.w, h.h);
    if cfg!(any(target_os = "android", target_os = "ios")) {
        let target = PADLOCK_HIT_DP * metrics_scale();
        w = w.max(target);
        hh = hh.max(target);
        // Do not let the grown rect reach past the neighbouring header: at
        // most one header's worth of extra width/height overall.
        let max_w = if h.is_row { row_label_w() } else { char_w() * 12.0 };
        let max_h = if h.is_row { row_h() } else { header_h() };
        w = w.min(max_w.max(h.w));
        hh = hh.min(max_h.max(h.h));
    }
    let cx = h.x + h.w / 2.0;
    let cy = h.y + h.h / 2.0;
    (cx - w / 2.0, cy - hh / 2.0, w, hh)
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
/// Paint a padlock glyph whose geometry is derived from [`padlock_w`] /
/// [`padlock_h`], so it scales with the host density instead of staying a
/// 10x12px speck on a phone. All offsets are fractions of the box.
fn paint_padlock(dc: &mut dyn DrawContext, ox: f64, oy: f64, locked: bool) {
    let (r, g, b) = if locked { PADLOCK_SHUT } else { PADLOCK_OPEN };
    let w = padlock_w();
    let h = padlock_h();
    let inset = padlock_inset();
    // Body occupies the lower ~45% of the box, the shackle the upper part.
    let body_y = oy + h * 0.5;
    let body_h = (h * 0.5 - inset).max(1.0);
    let body_w = (w - 2.0 * inset).max(1.0);
    if locked {
        dc.fill_rect(ox + inset, body_y, body_w, body_h, r, g, b, 1.0);
    } else {
        dc.stroke_rect(ox + inset, body_y, body_w, body_h, r, g, b, 1.0, 1.0);
    }
    // Shackle: two uprights joined by a top bar, inside the upper half.
    let bar_h = (h * 0.18).max(1.0);
    let leg_w = (w * 0.2).max(1.0);
    let leg_h = (body_y - oy - bar_h).max(1.0);
    let left_x = ox + w * 0.22;
    let right_x = ox + w - inset - leg_w;
    dc.fill_rect(left_x, oy + inset, leg_w, leg_h, r, g, b, 1.0);
    dc.fill_rect(left_x, oy + inset, (right_x + leg_w - left_x).max(1.0), bar_h, r, g, b, 1.0);
    if locked {
        dc.fill_rect(right_x, oy + inset, leg_w, leg_h, r, g, b, 1.0);
    } else {
        // Open: the right upright stops short, leaving the shackle ajar.
        dc.fill_rect(right_x, oy + inset, leg_w, (leg_h * 0.55).max(1.0), r, g, b, 1.0);
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
    // `prev_start` keeps the viewport where it is while the cursor moves
    // inside it; the helper only re-anchors when the cursor would fall
    // outside. Passing the remembered anchor is what lets a click choose the
    // position (see `centre_on_cursor`); 0 is the "derive from cursor" default.
    let prev = state.viewport_anchor.get().map(|(r, _)| r).unwrap_or(0);
    // Only the window is needed here; the start offset is for callers that
    // aim the viewport (see `centre_on_cursor`, which computes it itself).
    let (display, _) = {
        let app = state.app_ref();
        let sheet = app.core.workbook.active_sheet();
        ui_core::visible_row_indices(sheet, app.core.cursor, state.data_rows.get(), prev)
    };

    let (pinned, _) = pinned_sets(state);
    union_pinned(&display, &pinned)
}

/// Display columns with pins merged in (frozen first). See [`displayed_rows`].
fn displayed_cols(state: &GuiState) -> Vec<usize> {
    let prev = state.viewport_anchor.get().map(|(_, c)| c).unwrap_or(0);
    let (col_ixs, _) = {
        let app = state.app_ref();
        let sheet = app.core.workbook.active_sheet();
        ui_core::visible_col_indices(sheet, app.core.cursor, state.data_cols.get(), prev)
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

/// Height of the sheet tab strip (base device px, density-scaled): matches
/// the column-header strip so the chrome reads as one family.
const TAB_H_BASE: f64 = 24.0;
pub(crate) fn tab_h() -> f64 { TAB_H_BASE * metrics_scale() }
/// Active tab fill: the terminal reference paints the active sheet tab
/// black-on-yellow bold (ratatui `draw_visual`); the GUI uses a softer
/// yellow that stays distinct from its blue selection/cursor language.
const TAB_ACTIVE_BG: (f64, f64, f64) = (1.0, 1.0, 0.6);
/// Inactive tab fill: same gray as the row/column gutters.
const TAB_IDLE_BG: (f64, f64, f64) = (0.9, 0.9, 0.9);
/// Tab divider lines.
const TAB_DIV: (f64, f64, f64) = (0.55, 0.55, 0.55);
/// Horizontal padding inside each tab and gap between tabs (base device px,
/// density-scaled).
const TAB_PAD_X_BASE: f64 = 10.0;
const TAB_GAP_BASE: f64 = 6.0;
pub(crate) fn tab_pad_x() -> f64 { TAB_PAD_X_BASE * metrics_scale() }
pub(crate) fn tab_gap() -> f64 { TAB_GAP_BASE * metrics_scale() }

/// A painted sheet tab from the last frame: strip position plus which sheet
/// index a click switches to.
#[derive(Clone, Copy, Debug, PartialEq)]
struct TabHit {
    x0: f64,
    x1: f64,
    index: usize,
}

/// Lay out sheet tabs left to right from x=2: each tab pads its measured
/// title by tab_pad_x() on both sides, tab_gap() separates tabs. The active tab
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
        let x1 = x0 + tab_pad_x() + tw + tab_pad_x();
        hits.push(TabHit { x0, x1, index });
        x = x1 + tab_gap();
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
            dc.text_extents_styled(t, "monospace", font_size(), 0, weight).2
        };
        tab_layout(&titles, active, &measure)
    };
    for hit in &hits {
        let is_active = hit.index == active;
        let (r, g, b) = if is_active { TAB_ACTIVE_BG } else { TAB_IDLE_BG };
        dc.fill_rect(hit.x0, 2.0, hit.x1 - hit.x0, tab_h() - 4.0, r, g, b, 1.0);
        // Divider at the tab's right edge (also the click target's edge).
        dc.fill_rect(hit.x1, 2.0, 1.0, tab_h() - 4.0, TAB_DIV.0, TAB_DIV.1, TAB_DIV.2, 1.0);
        let (tr, tg, tb) = if is_active { (0.0, 0.0, 0.0) } else { (0.3, 0.3, 0.3) };
        let weight = if is_active { 1 } else { 0 };
        dc.draw_text_styled(
            hit.x0 + tab_pad_x(),
            (tab_h() - font_size() * 1.2) / 2.0,
            &titles[hit.index],
            "monospace",
            font_size(),
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
    // (that would shift footer addressing and break ratatui parity).
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
    let app = state.app_ref();
    let cursor_row = state.last_row.get();
    let cursor_col = state.last_col.get();
    let display_rows: Vec<usize> = displayed_rows(state);
    let col_ixs: Vec<usize> = displayed_cols(state);
    let mc = app.core.workbook.active_sheet().grid.main_cols();
    let col_widths: HashMap<usize, usize> = col_ixs
        .iter()
        .map(|&c| (c, display_col_width(&app.core.workbook.active_sheet(), c, mc)))
        .collect();
    let pinned_rows: std::collections::BTreeSet<usize> =
        state.pinned_rows.borrow().iter().copied().collect();
    let pinned_cols: std::collections::BTreeSet<usize> =
        state.pinned_cols.borrow().iter().copied().collect();
    let vp = viewport_from_parts(&display_rows, &col_ixs, &col_widths, app);

    dc.clear(0.94, 0.94, 0.94, 1.0);
    let mut padlocks: Vec<GutterPadlock> = Vec::new();
    render_grid_body_inner(
        dc,
        app,
        &vp,
        &col_widths,
        cursor_row,
        cursor_col,
        w,
        h,
        state.editing.get(),
        &state.edit_buf.borrow(),
        true,
        &pinned_rows,
        &pinned_cols,
        &mut padlocks,
    );
    *state.padlocks.borrow_mut() = padlocks;
}

/// Draw the synthetic movie pointer, if one is armed.
///
/// Drawn rather than relying on the X cursor, which `x11grab` does not capture
/// (a moved pointer changes no pixels in the grab, verified) — so an
/// unassisted recording cannot show the mouse at all. The arrow is outlined in
/// white and filled in a dark colour so it stays legible over the grid, the
/// menu bar and an open menu alike; `pressed` turns it red so a click reads in
/// the recording.
fn paint_movie_pointer(dc: &mut dyn DrawContext, state: &GuiState) {
    // `(x, y)` is in **canvas** coordinates, not window coordinates: the canvas
    // widget sits below the menu bar and the formula bar, so its origin is
    // `2 * header_h()` down from the window top (48px at the default metrics).
    // Anything computing a window-relative position for the pointer has to
    // subtract that, or the arrow draws ~48px below where it was aimed.
    let Some((x, y, pressed)) = state.movie_pointer.get() else {
        return;
    };
    // Classic arrow: tip at (x, y), a tall leading edge and a notched tail.
    const BODY: &[(f64, f64)] = &[
        (0.0, 0.0),
        (0.0, 17.0),
        (4.5, 12.5),
        (7.5, 19.0),
        (10.5, 17.5),
        (7.5, 11.0),
        (13.5, 11.0),
    ];
    // Outline first, then the fill, so the arrow stays legible whether it sits
    // over the pale grid or a dark menu surface. `pressed` turns it red, which
    // is what makes a click read in a recording.
    for i in 0..BODY.len() {
        let (ax, ay) = BODY[i];
        let (bx, by) = BODY[(i + 1) % BODY.len()];
        stroke_segment(dc, x + ax, y + ay, x + bx, y + by, 3.5, 1.0, 1.0, 1.0);
    }
    let (r, g, b) = if pressed { (0.85, 0.1, 0.1) } else { (0.05, 0.05, 0.15) };
    for row in 0..17 {
        let t = row as f64 / 17.0;
        stroke_segment(dc, x + 1.0, y + t * 17.0, x + 1.0 + 5.0 * (1.0 - t), y + t * 17.0, 2.0, r, g, b);
    }
    for row in 0..7 {
        stroke_segment(dc, x + 7.0, y + 11.5 + row as f64, x + 10.0, y + 11.5 + row as f64, 3.0, r, g, b);
    }
}

/// Draw a line of arbitrary direction as a thin filled rectangle.
fn stroke_segment(dc: &mut dyn DrawContext, x0: f64, y0: f64, x1: f64, y1: f64, w: f64, r: f64, g: f64, b: f64) {
    let steps = ((x1 - x0).abs().max((y1 - y0).abs())).ceil().max(1.0) as i32;
    for i in 0..=steps {
        let t = i as f64 / steps as f64;
        let px = x0 + t * (x1 - x0);
        let py = y0 + t * (y1 - y0);
        dc.fill_rect(px - w / 2.0, py - w / 2.0, w, w, r, g, b, 1.0);
    }
}

/// Assemble a full [`Viewport`] from display rows/columns the caller already
/// resolved. Keeps the row labels, column layout and row-aggregate derivation
/// in one place for callers that compute their own visible sets (the live
/// canvas state and the movie painter).
fn viewport_from_parts(
    display_rows: &[usize],
    col_ixs: &[usize],
    col_widths: &HashMap<usize, usize>,
    app: &super::App,
) -> crate::gui::viewport::Viewport {
    let g = &app.core.workbook.active_sheet().grid;
    let mr = g.main_rows();
    let mc = g.main_cols();
    let row_labels: Vec<(u32, String)> = display_rows
        .iter()
        .enumerate()
        .map(|(idx, &r)| (idx as u32, crate::addr::ui_row_label(r, mr)))
        .collect();
    let column_layout: Vec<(u32, u32, String)> = col_ixs
        .iter()
        .map(|&c| {
            let w = *col_widths.get(&c).unwrap_or(&1);
            (c as u32, w as u32, crate::addr::ui_column_fragment(c, mc))
        })
        .collect();
    let row_agg_func = compute::compute_row_agg_func(g, display_rows, HEADER_ROWS, mr);
    crate::gui::viewport::Viewport {
        display_rows: display_rows.to_vec(),
        col_ixs: col_ixs.to_vec(),
        col_widths: col_widths.clone(),
        row_labels,
        column_layout,
        row_agg_func,
        mr,
        mc,
        data_width: col_widths.values().copied().sum(),
    }
}

/// Paint a complete sheet body (gutter, column headers, cells, selection) for
/// an already-computed [`Viewport`].
///
/// This is the single sheet renderer: the interactive canvas calls it with its
/// live state, and the `--movie` painter calls it with a frame's viewport, so a
/// recorded frame is pixel-identical to the window (margin shading included)
/// instead of being a second, drifting implementation.
#[allow(clippy::too_many_arguments)]
pub(crate) fn render_grid_body(
    dc: &mut dyn DrawContext,
    app: &super::App,
    vp: &crate::gui::viewport::Viewport,
    col_widths: &HashMap<usize, usize>,
    cursor_row: usize,
    cursor_col: usize,
    w: i32,
    h: i32,
) {
    let mut padlocks: Vec<GutterPadlock> = Vec::new();
    let (pinned_rows, pinned_cols) = (Default::default(), Default::default());
    render_grid_body_inner(
        dc,
        app,
        vp,
        col_widths,
        cursor_row,
        cursor_col,
        w,
        h,
        false,
        "",
        true,
        &pinned_rows,
        &pinned_cols,
        &mut padlocks,
    );
}

#[allow(clippy::too_many_arguments)]
fn render_grid_body_inner(
    dc: &mut dyn DrawContext,
    app: &super::App,
    vp: &crate::gui::viewport::Viewport,
    col_widths: &HashMap<usize, usize>,
    cursor_row: usize,
    cursor_col: usize,
    w: i32,
    h: i32,
    is_editing: bool,
    edit_buf: &str,
    publish_pins: bool,
    pinned_rows: &std::collections::BTreeSet<usize>,
    pinned_cols: &std::collections::BTreeSet<usize>,
    padlocks: &mut Vec<GutterPadlock>,
) {
    let display_rows = &vp.display_rows;
    let col_ixs = &vp.col_ixs;
    let hr = HEADER_ROWS;
    let lm = MARGIN_COLS;
    let mr = vp.mr;
    let mc = vp.mc;
    dc.clip(0.0, 0.0, w as f64, h as f64);

    // Header coverage mirrors the body's selection rectangle (anchor↔cursor on
    // both axes — the GUI has no Rows/Cols-only modes). None while navigating
    // plainly, so unselected chrome renders exactly as before.
    let cover: Option<((usize, usize), (usize, usize))> = app.core.anchor.map(|a| {
        let (r0, r1) = (a.row.min(cursor_row), a.row.max(cursor_row));
        let (c0, c1) = (a.col.min(cursor_col), a.col.max(cursor_col));
        ((r0, r1), (c0, c1))
    });
    paint_row_headers(dc, display_rows, mr, pinned_rows, padlocks, cover.map(|(r, _)| r));
    paint_col_headers(dc, col_ixs, col_widths, mc, pinned_cols, padlocks, cover.map(|(_, c)| c));

    // Fill cells via the shared render pipeline.
    let mut sink = GuiCanvasSink::new();
    render::fill_cells(
        &mut sink,
        display_rows,
        col_ixs,
        col_widths,
        &app.core.workbook.active_sheet().grid,
        hr,
        mr,
        mc,
        lm,
        vp.data_width(),
        cursor_row,
        cursor_col,
        &vp.row_agg_func,
    );
    render_to(
        &sink,
        dc,
        col_ixs,
        col_widths,
        display_rows,
        hr,
        mr,
        mc,
        lm,
        cursor_row,
        cursor_col,
        is_editing,
        edit_buf,
        app.core.anchor.map(|a| (a.row, a.col)),
    );

    // Status line at the bottom of the body.
    if h as f64 > header_h() + 20.0 {
        dc.fill_rect(0.0, h as f64 - 20.0, w as f64, 20.0, 0.9, 0.9, 0.9, 1.0);
    }

    // Diagnostic overlay (cell/sink/cursor/key state) painted over the grid.
    // Opt-in via CORRO_DEBUG_OVERLAY (any value): off by default so normal runs
    // render a clean sheet.
    if std::env::var_os("CORRO_DEBUG_OVERLAY").is_some() && publish_pins {
        let addr = crate::grid::CellAddr::Main { row: 0, col: 0 };
        let cell_val = app.core.workbook.active_sheet().grid.get(&addr).unwrap_or_default();
        let mode = if is_editing { "EDIT" } else { "NORM" };
        dc.draw_text(
            100.0,
            h as f64 - 100.0,
            &format!("Cell A1='{cell_val}' Mode:{mode} Buf:'{edit_buf}'"),
            "monospace",
            12.0,
            0.5,
            0.0,
            0.5,
            1.0,
        );
    }
}
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
    // Tail external changes on every keystroke (parity with the TUI loop,
    // which polls sync_external each iteration): another window's committed
    // cells land here without a reload. Commits append to the log
    // immediately, so no save step is needed on either side. Redraw only
    // when something actually arrived; tail errors become status text,
    // never a lost keystroke.
    let tailed = {
        let app = state.app_mut();
        match app.core.poll_log_tail() {
            Ok(changed) => changed,
            Err(e) => {
                app.core.status = format!("Sync error: {e}");
                true
            }
        }
    };
    if tailed {
        update_formula_bar(state, HEADER_ROWS, MARGIN_COLS);
        sync_chrome_labels(state);
        state.canvas.queue_redraw();
    }
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

    // The in-grid aggregate dropdown owns the keyboard while open (it is
    // painted on the canvas, so there is no native widget to route to).
    if agg_dropdown_open(state) {
        match key {
            ESCAPE => {
                log_key_action(keyval, "agg_drop_cancel", "");
                hide_agg_dropdown(state);
                super::agg_picker::close(state.app_mut());
                state.canvas.queue_redraw();
                return true;
            }
            RETURN => {
                log_key_action(keyval, "agg_drop_commit", "");
                agg_dropdown_commit(state);
                return true;
            }
            UP | LEFT => {
                log_key_action(keyval, "agg_drop_step_up", "");
                agg_dropdown_step(state, -1);
                return true;
            }
            DOWN | RIGHT | TAB => {
                log_key_action(keyval, "agg_drop_step_down", "");
                agg_dropdown_step(state, 1);
                return true;
            }
            // Digit hotkeys commit the numbered row directly (ratatui
            // parity: `1`..=`7`). `1`..=`9` are exactly the digit keyvals.
            k if (49..=57).contains(&k) => {
                let ch = char::from_u32(k).unwrap_or('0');
                if let Some(idx) = crate::ui_core::agg_choice_index_for_digit(ch) {
                    log_key_action(keyval, "agg_drop_digit", &format!("digit={ch} idx={idx}"));
                    // The dropdown's own `sel` is what commit reads.
                    if let Some(d) = state.agg_drop.borrow_mut().as_mut() {
                        d.sel = idx;
                    }
                    super::agg_picker::set(state.app_mut(), idx);
                    agg_dropdown_commit(state);
                }
                state.canvas.queue_redraw();
                return true;
            }
            _ => {}
        }
    }

    if state.editing.get() {
        return handle_edit_key(key, state, mods);
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
            start_edit_seeded_from_cell(state);
            true
        }
        F3 => {
            // Keyboard route to the margin-aggregate picker. Resolves a key
            // from *any* cursor position (on a key cell -> that cell; on a
            // data cell -> the margin key governing it) so the picker is
            // reachable without hunting for the key cell.
            if open_agg_picker_inline(state_rc) {
                log_key_action(keyval, "agg_picker", &format!("cell={}", format_cell(state)));
            } else {
                log_key_action(keyval, "agg_picker_noop", &format!("cell={}", format_cell(state)));
                state.app_mut().core.status =
                    "Aggregate: no margin TOTAL/MAX/… key for this cell".into();
                state.canvas.queue_redraw();
            }
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
                        return handle_edit_key(key, state, mods);
                    }
                }
                // When window CAPTURE consumed printable chars and pushed to
                // edit_buf directly (bypassing the entry widget), the entry
                // text is empty but edit_buf has content.  Commit from there.
                if !state.edit_buf.borrow().is_empty() {
                    state.editing.set(true);
                    return handle_edit_key(key, state, mods);
                }
            }
            // Focus can drift off the entry (setup grab_focus races
            // present()'s pump): a window-observed Return with a non-empty
            // buffer is a commit, not navigation — the entry's CAPTURE
            // intercept only fires when focused. Ratatui parity: Edit +
            // Return commits; bare Return with an empty buffer still moves.
            if state.editing.get() && !state.edit_buf.borrow().is_empty() {
                return handle_edit_key(key, state, mods);
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

fn handle_edit_key(key: u32, state: &GuiState, mods: u32) -> bool {
    // `mods` is unused today (Shift+arrows never reach here — the caller
    // routes selection separately) but kept in the signature so the call
    // sites stay uniform with the other key paths.
    let _ = mods;
    match key {
        // F2 re-seeds the edit with the cell's current value. The GUI idles
        // with `editing=true` (type-first), so without this arm F2 would be
        // swallowed here and the test's "F2 seeds the cell" flow could never
        // run — the buffer would stay empty and the next keystroke would
        // replace the value instead of appending to it.
        F2 => {
            log_key_action(key, "start_edit", &format!("cell={} mode=edit", format_cell(state)));
            start_edit_seeded_from_cell(state);
            true
        }
        RETURN => {
            log_key_action(key, "commit_edit", &format!("cell={} mode=edit", format_cell(state)));
            commit_edit(state);
            move_cursor(state, 1, 0);
            true
        }
        ESCAPE => {
            log_key_action(key, "cancel_edit", &format!("cell={} mode=edit", format_cell(state)));
            state.editing.set(false);
            state.entry_clicked.set(false);
            state.edit_buf.borrow_mut().clear();
            state.edit_caret.set(0);
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
            // No-op on an empty buffer (and never resync): selections show
            // the cell value while editing with nothing typed yet, and a
            // sync here would blank that display without changing anything.
            if !state.edit_buf.borrow().is_empty() {
                // Caret-aware: removes the char BEFORE the caret.
                edit_backspace(state);
                sync_entry_to_buf(state);
                state.canvas.queue_redraw();
            }
            true
        }
        DELETE => {
            log_key_action(key, "edit_clear", &format!("cell={} mode=edit", format_cell(state)));
            // Same empty-buffer rule as Backspace above.
            if !state.edit_buf.borrow().is_empty() {
                if crate::formula::is_formula(&state.edit_buf.borrow()) {
                    // Formula editing: Delete removes the char AT the caret
                    // (ratatui parity — `edit_delete_forward`).
                    edit_delete_forward(state);
                } else {
                    // Plain value: Delete clears the whole buffer.
                    state.edit_buf.borrow_mut().clear();
                    state.edit_caret.set(0);
                }
                sync_entry_to_buf(state);
                state.canvas.queue_redraw();
            }
            true
        }
        LEFT => {
            // Caret-left inside the buffer (ratatui `Mode::Edit` parity): the
            // caret moves without ending the edit. Only at the buffer start
            // does Left commit and step to the previous cell. Committing on
            // every Left made typing insert at the end and moved the cell
            // cursor instead of the caret.
            let mut caret = state.edit_caret.get();
            let outcome = text_edit::left(&state.edit_buf.borrow(), &mut caret);
            if outcome == text_edit::KeyOutcome::Edited {
                state.edit_caret.set(caret);
                push_caret_to_entry(state);
                state.canvas.queue_redraw();
                return true;
            }
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
            // Mirror of LEFT: move the caret to the buffer end, then commit
            // and step right.
            let mut caret = state.edit_caret.get();
            let outcome = text_edit::right(&state.edit_buf.borrow(), &mut caret);
            if outcome == text_edit::KeyOutcome::Edited {
                state.edit_caret.set(caret);
                push_caret_to_entry(state);
                state.canvas.queue_redraw();
                return true;
            }
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
            // Adopt the displayed value when the buffer is empty: the edit
            // session is armed with the cell's text on screen (idle
            // type-first state, or a click into the formula bar). Pushing
            // onto the empty buffer would replace that text with this one
            // char; restore it first, keeping the caret the click placed (so
            // typing inserts at the clicked position, not always at the end).
            if state.edit_buf.borrow().is_empty() {
                if let Some(shown) = state.formula_entry.get_text() {
                    if !shown.is_empty() {
                        *state.edit_buf.borrow_mut() = shown;
                    }
                }
                state.entry_clicked.set(false);
                pull_caret_from_entry(state);
                state.edit_caret.set(state.edit_caret.get().min(edit_caret_len(state)));
            }
            // Caret-aware insert, so typing mid-buffer inserts at the caret.
            edit_insert_char(state, ch);
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

/// Fresh empty edit (F2). Retained for callers that want a blank buffer;
/// the interactive F2 path uses [`start_edit_with_text`] so the cell's
/// current value is seeded rather than discarded.
#[allow(dead_code)]
fn start_edit(state: &GuiState) {
    state.editing.set(true);
    state.edit_buf.borrow_mut().clear();
    state.edit_caret.set(0);
    state.formula_entry.set_text("");
    state.formula_entry.grab_focus();
    state.canvas.queue_redraw();
}

/// F2: edit the cursor cell, seeded with its current value (LibreOffice
/// parity — typing appends to the existing content instead of replacing it,
/// so a keystroke can never silently discard a value the user never saw).
///
/// Reads the **cell** (not the formula bar, which may not have refreshed
/// yet for the current cursor) so the seed is always the real content.
fn start_edit_seeded_from_cell(state: &GuiState) {
    // Editing supersedes the aggregate dropdown.
    hide_agg_dropdown(state);
    state.editing.set(true);
    state.entry_clicked.set(false);
    let raw = {
        let app = state.app_ref();
        let grid = &app.core.workbook.active_sheet().grid;
        let addr = crate::addr::sheet_cursor_to_addr(
            crate::addr::LogicalRow(state.last_row.get()),
            crate::addr::GlobalCol(state.last_col.get()),
            crate::addr::MainRows(grid.main_rows()),
            crate::addr::MainCols(grid.main_cols()),
        );
        grid.get(&addr).unwrap_or_default()
    };
    *state.edit_buf.borrow_mut() = raw;
    sync_entry_to_buf(state);
    state.edit_caret.set(edit_caret_len(state));
    push_caret_to_entry(state);
    state.formula_entry.grab_focus();
    state.canvas.queue_redraw();
}

/// What run_gui's setup should do after `present()` returns, chosen from
/// state that may already have been changed by events delivered *during*
/// `present()`'s main-loop pump. Pure so the race is unit-testable without
/// widgets.
///
/// `present()` pumps the GTK main loop for hundreds of iterations before the
/// real loop runs, and the canvas click/key handlers are registered before
/// it. A click arriving in that window has already selected its cell and may
/// have opened the aggregate dropdown; running the default "select A1" step
/// afterwards would (a) move the yellow edit highlight to A1 and (b)
/// `hide_agg_dropdown`, closing the list the click just opened. Hence a
/// click wins over the default selection.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum StartupEdit {
    /// An edit with typed content is already in flight: keep it, just focus.
    KeepInFlight,
    /// A click already chose the cell: do not clobber it with A1.
    KeepClicked,
    /// Fresh start: select A1 for typing.
    SelectA1,
}

fn startup_edit_action(editing_with_content: bool, clicked: bool) -> StartupEdit {
    if editing_with_content {
        StartupEdit::KeepInFlight
    } else if clicked {
        StartupEdit::KeepClicked
    } else {
        StartupEdit::SelectA1
    }
}

/// Keeps the formula bar showing the cell's value
/// (grid clicks and setup select; they must display, not blank). The buffer
/// is still cleared, so typing replaces (ratatui parity) and committing an
/// untouched selection stays a quiet no-op via the empty check in
/// [`commit_edit`]. F2 seeds via [`start_edit_with_text`] instead.
fn start_edit_keep_display(state: &GuiState) {
    // Typing into the cell supersedes the aggregate dropdown.
    hide_agg_dropdown(state);
    state.editing.set(true);
    state.edit_buf.borrow_mut().clear();
    // Empty buffer: the caret belongs at the start.
    state.edit_caret.set(0);
    push_caret_to_entry(state);
    state.formula_entry.grab_focus();
    state.canvas.queue_redraw();
    // Window-level cascade as well: a canvas-only queue_draw may not arm the
    // toplevel's frame clock, so the repaint never happens (the cursor
    // appears not to follow navigation). update_state_cursor does the same.
    state.window.queue_redraw();
}

fn start_edit_with(state: &GuiState, ch: char) {
    let already_editing = state.editing.get();
    state.editing.set(true);
    // Click-to-edit adopt: the user clicked into the formula bar (which
    // displays the cell's formula) and is now typing into it. Pushing onto
    // the empty buffer would wipe the formula to this one char; restore the
    // displayed text first. The caret is whatever the click put in the
    // widget (adopted by `pull_caret_from_entry` on button-press), so
    // typing inserts *at the clicked position* rather than always appending.
    // Adopt whenever the entry is displaying a value and the buffer is
    // empty: the displayed text is the cell's current content, and starting
    // from it (rather than from nothing) is what makes typing EDIT the cell
    // instead of silently replacing its value. A stale empty buffer is never
    // a reason to discard what the user can see.
    if state.edit_buf.borrow().is_empty() {
        let shown = state.formula_entry.get_text().unwrap_or_default();
        if !shown.is_empty() {
            *state.edit_buf.borrow_mut() = shown;
            // Only a genuine click places a caret worth honouring; a
            // function-bar refresh or the type-first idle state leaves the
            // widget's position at 0, which would prepend instead of
            // appending. Default to the end (append) in that case — the
            // least destructive reading of "the user is editing this value".
            if state.entry_clicked.get() {
                pull_caret_from_entry(state);
            } else {
                state.edit_caret.set(edit_caret_len(state));
            }
            state.edit_caret.set(state.edit_caret.get().min(edit_caret_len(state)));
        }
        state.entry_clicked.set(false);
    }
    // Caret-aware insert (never append): the same `text_edit` surface
    // Left/Right and Backspace drive, so typing mid-buffer inserts there.
    edit_insert_char(state, ch);
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
/// Caret-aware text editing for the formula buffer.
///
/// The native Entry exposes no caret API, so the host owns the caret exactly
/// as ratatui's `Mode::Edit` does. Pure `(buffer, caret)` operations so the
/// semantics are unit-testable without a window.
mod text_edit {
    /// What an edit-mode key did, so the caller knows whether the edit
    /// continues or the cell cursor should move.
    #[derive(Clone, Copy, Debug, Eq, PartialEq)]
    pub enum KeyOutcome {
        Edited,
        CommitAndMoveCell,
    }

    /// Insert `ch` at the caret and advance it.
    pub fn insert_char(buf: &mut String, caret: &mut usize, ch: char) {
        let mut chars: Vec<char> = buf.chars().collect();
        let pos = (*caret).min(chars.len());
        chars.insert(pos, ch);
        *buf = chars.into_iter().collect();
        *caret = pos + 1;
    }

    /// Remove the char before the caret (no-op at position 0).
    pub fn backspace(buf: &mut String, caret: &mut usize) {
        let mut chars: Vec<char> = buf.chars().collect();
        let pos = (*caret).min(chars.len());
        if pos > 0 {
            chars.remove(pos - 1);
            *buf = chars.into_iter().collect();
            *caret = pos - 1;
        }
    }

    /// Remove the char *at* the caret (no-op at end).
    pub fn delete_forward(buf: &mut String, caret: &mut usize) {
        let mut chars: Vec<char> = buf.chars().collect();
        let pos = (*caret).min(chars.len());
        if pos < chars.len() {
            chars.remove(pos);
            *buf = chars.into_iter().collect();
        }
    }

    /// Caret within `buf` in characters.
    pub fn caret_len(buf: &str) -> usize {
        buf.chars().count()
    }

    /// Move the caret left, or report commit-and-move-cell at the start.
    pub fn left(buf: &str, caret: &mut usize) -> KeyOutcome {
        if *caret > 0 {
            *caret -= 1;
            KeyOutcome::Edited
        } else {
            let _ = buf;
            KeyOutcome::CommitAndMoveCell
        }
    }

    /// Mirror of [`left`] at the end of the buffer.
    pub fn right(buf: &str, caret: &mut usize) -> KeyOutcome {
        if *caret < caret_len(buf) {
            *caret += 1;
            KeyOutcome::Edited
        } else {
            KeyOutcome::CommitAndMoveCell
        }
    }
}

/// Push the app's caret into the widget so the visible cursor matches what
/// the next edit will use (a `set_text` otherwise resets it to the end).
fn push_caret_to_entry(state: &GuiState) {
    state.formula_entry.set_position(state.edit_caret.get());
}

/// Pull the widget's caret into the app before an edit: the widget owns the
/// authoritative cursor once the user has clicked or arrowed inside it.
fn pull_caret_from_entry(state: &GuiState) {
    if let Some(pos) = state.formula_entry.get_position() {
        state.edit_caret.set(pos);
    }
}

/// Insert `ch` at the caret (never appended).
fn edit_insert_char(state: &GuiState, ch: char) {
    let mut caret = state.edit_caret.get();
    text_edit::insert_char(&mut state.edit_buf.borrow_mut(), &mut caret, ch);
    state.edit_caret.set(caret);
}

/// Backspace at the caret (no-op at position 0).
fn edit_backspace(state: &GuiState) {
    let mut caret = state.edit_caret.get();
    text_edit::backspace(&mut state.edit_buf.borrow_mut(), &mut caret);
    state.edit_caret.set(caret);
}

/// Delete the char at the caret (no-op at end).
fn edit_delete_forward(state: &GuiState) {
    let mut caret = state.edit_caret.get();
    text_edit::delete_forward(&mut state.edit_buf.borrow_mut(), &mut caret);
    state.edit_caret.set(caret);
}

/// Caret length of the buffer in characters.
fn edit_caret_len(state: &GuiState) -> usize {
    text_edit::caret_len(&state.edit_buf.borrow())
}

fn sync_entry_to_buf(state: &GuiState) {
    let buf = state.edit_buf.borrow().clone();
    if state.formula_entry.get_text().as_deref() != Some(buf.as_str()) {
        state.formula_entry.set_text(&buf);
    }
    push_caret_to_entry(state);
}

fn commit_edit(state: &GuiState) {
    state.editing.set(false);
    // A commit is a fresh boundary (ratatui parity): later typing starts a
    // new value, it never repairs into the committed formula.
    state.entry_clicked.set(false);
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
        // Refresh chrome here, not just the canvas: commits arriving via
        // the window key path (unfocused Return) otherwise leave the
        // formula row and hints showing pre-commit text until the next
        // cursor event. Idempotent for paths that refresh separately.
        update_formula_bar(state, state.last_row.get(), state.last_col.get());
    }
    state.edit_buf.borrow_mut().clear();
    state.edit_caret.set(0);
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
    // Use the same anchor the renderer uses. Computing with `prev_start = 0`
    // here disagreed with `displayed_rows`/`displayed_cols` whenever an anchor
    // was set: the zero-anchored window need not contain the cursor, so this
    // "corrected" the cursor on the next frame and undid the centring a click
    // had just performed.
    let (prev_r, prev_c) = state.viewport_anchor.get().unwrap_or((0, 0));
    let (display_rows, _) = ui_core::visible_row_indices(sheet, app.core.cursor, state.data_rows.get(), prev_r);
    let (col_ixs, _) = ui_core::visible_col_indices(sheet, app.core.cursor, state.data_cols.get(), prev_c);
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
    // Genuine navigation dismisses the in-grid dropdown: it is anchored to
    // one cell, and leaving it floating while the cursor moves would put it
    // over an unrelated cell. (Done here, not in `update_state_cursor`,
    // which also runs from scroll callbacks.)
    hide_agg_dropdown(state);
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
/// cursor and by a minimal 2x2 body (so startup always shows data row 2
/// and column B). Grows toward the target and silently shrinks abandoned
/// growth back toward it (via the cursor floor + the grid's silent
/// shrink_to_content) — this is what lets the sheet shrink again when
/// blank rows/cols are no longer needed (navigate back up, delete
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
    // Converge on need below (see doc comment): recompute post-growth,
    // then grow toward target always but shrink only when allowed (fresh
    // loads/switches must not second-guess stored extents).
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
    // The mobile block below raises these to cover the viewport; on desktop
    // they are never reassigned, so silence the `mut` warning there rather
    // than leaving one that fires on every desktop build.
    #[cfg_attr(not(any(target_os = "android", target_os = "ios")), allow(unused_mut))]
    let mut target_r = (content_r + 1).max(floor_r).max(2);
    #[cfg_attr(not(any(target_os = "android", target_os = "ios")), allow(unused_mut))]
    let mut target_c = (content_c + 1).max(floor_c).max(2);
    // Android: a phone screen shows ~38 rows and ~17 columns at once, but the
    // minimal 2x2 body above leaves all the rest as header/footer rows that
    // render as margin grey - a blank-looking sheet with two usable cells. The
    // desktop rule (start small, grow as you type) assumes a pointer and a
    // large screen; on a touch device the body should fill what the user can
    // actually see. Growth only, and capped by the viewport, so it cannot run
    // away or shrink a stored extent.
    #[cfg(any(target_os = "android", target_os = "ios", target_os = "macos"))]
    {
        let visible_rows = state.data_rows.get().max(1);
        let visible_cols = state.data_cols.get().max(1);
        // The viewport spends a column on the left margin border and one on
        // the right, so a body sized to exactly `visible_cols` still leaves
        // the far edge showing margin grey — on a phone an empty sheet then
        // read as a white patch in a grey field. Ask for the border's worth
        // extra so the body spans what the user can see.
        const MARGIN_BORDER_COLS: usize = 2;
        // A sheet grown to exactly the viewport has NO scroll range: the
        // display list comes back the same length as the window, so
        // `visible_row_indices` clamps `start` to 0 and the viewport can never
        // move. That also made aiming the viewport at a tapped cell a no-op —
        // there was nothing to scroll. A spreadsheet is expected to extend past
        // what is on screen, so keep a screenful of headroom below the body:
        // enough to scroll somewhere, bounded so the sheet cannot run away.
        const SCROLL_HEADROOM_ROWS: usize = 1;
        let headroom = visible_rows * (SCROLL_HEADROOM_ROWS + 1);
        target_r = target_r.max(headroom.min(MAX_RENDER_ROWS));
        target_c = target_c.max((visible_cols + MARGIN_BORDER_COLS).min(MAX_RENDER_COLS));
    }
    // Grow toward target (covers the minimal 2x2 body on empty sheets and
    // any cursor floor above current extent).
    while grid.main_rows() < target_r {
        grid.grow_main_row_at_bottom();
    }
    while grid.main_cols() < target_c {
        grid.grow_main_col_at_right();
    }
    if !allow_shrink {
        return;
    }
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
    // NOTE: this is a *state sync* helper, also reached from scroll callbacks
    // (e.g. `canvas.grab_focus()` synchronously scrolls to the cursor). It
    // deliberately does NOT dismiss the in-grid dropdown: doing so made
    // opening the dropdown hide it again via that re-entrant path. Genuine
    // navigation dismisses it explicitly in `move_cursor`.
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
    // then eat edit_buf too). Address/status labels always update. An empty
    // buffer holds no in-flight edit even when the editing flag is on
    // (post-click/post-setup selection state), so selections and cursor
    // moves always display the new cell instead of going stale or blank.
    if !state.editing.get() || state.edit_buf.borrow().is_empty() {
        state.formula_entry.set_text(&val);
    }
    // Click-to-edit bookkeeping (see entry_clicked): a refresh for a
    // different cell ends click-edit intent; the snapshot always tracks the
    // displayed grid truth so a repair can never use stale text.
    if (row, col) != state.entry_shown.get() {
        state.entry_clicked.set(false);
        state.entry_shown.set((row, col));
    }
    *state.entry_snapshot.borrow_mut() = val;
    sync_chrome_labels(state);
}

/// Push ratatui-parity chrome labels: bottom strip always shows the shared
/// hints line; the formula row shows `· status` trailing only when status
/// is non-empty and no edit/dialog owns the row. The edit gate is the
/// BUFFER, not the `editing` flag: the GUI idles with `editing=true`
/// (entry focused, ready to type — its type-first design), so gating on
/// the flag would hide every at-rest status. A non-empty buffer means a
/// genuine in-progress edit (ratatui's Edit arm shows the buffer with no
/// status span); Help mode hides it like ratatui's About arm. Plain text
/// on every backend (no markup: nwg/wasm render markup tags literally,
/// so shared code must not emit any — the DarkGray hint tint stays
/// terminal-only).
fn sync_chrome_labels(state: &GuiState) {
    let app = state.app_ref();
    // Same text as the ratatui bottom row (never status).
    state.hints_label.set_text(&crate::core::state::normal_hints(
        app.core.anchor.is_some(),
        !app.core.op_history.is_empty(),
        !app.core.redo_history.is_empty(),
        app.core.path.is_some(),
    ));
    let show = state.edit_buf.borrow().is_empty()
        && matches!(state.mode.get(), GuiMode::Normal)
        && !app.core.status.is_empty();
    if show {
        state
            .formula_status
            .set_text(&format!("   ·  {}", app.core.status));
    } else {
        // Clear as well as hide: a backend that still measures hidden
        // children must allocate nothing for the suffix.
        state.formula_status.set_text("");
    }
    state.formula_status.set_visible(show);
}

// ---------------------------------------------------------------------------
// Click handling
// ---------------------------------------------------------------------------

/// Centre the viewport on the current cursor.
///
/// The viewport is otherwise cursor-derived but *sticky*: `visible_*_indices`
/// keeps its `prev_start` unless the cursor would leave the window, so a click
/// near an edge leaves the cell sitting at that edge. Selecting a cell should
/// instead put it in the middle of the screen.
///
/// `prev_start` is an index into the *full* display list, while
/// `displayed_rows`/`displayed_cols` return only the current window — so the
/// cursor's absolute position has to come from the full list, computed here
/// with the same helper the viewport uses (pass `prev_start = 0` and read the
/// returned offset). Only an anchor is stored; the clamping that keeps a
/// cursor near the sheet's start or end inside the window lives in
/// `visible_*_indices`, so it does not need duplicating here.
fn centre_on_cursor(state: &GuiState, app: &super::App) {
    let sheet = app.core.workbook.active_sheet();
    let cursor = app.core.cursor;
    let dim_r = state.data_rows.get().max(1);
    let dim_c = state.data_cols.get().max(1);

    // Rows. `visible_row_indices` returns the window AND the index it started
    // at. Asking with `prev_start = 0` makes it anchor the window on the
    // cursor whenever it has to scroll, so the returned `start` *is* the
    // cursor's absolute index in the display list — the unit `prev_start`
    // is expressed in. (Reading the position out of the returned window
    // instead would be wrong: that list is already trimmed to `dim`, and it
    // has pinned rows merged in, so its offsets are not the anchor's.)
    let row_pos = ui_core::cursor_display_row_index(sheet, cursor);
    let centred_r = row_pos.saturating_sub(dim_r / 2);

    let centred_c = centre_col_anchor(state, sheet, cursor, dim_c);

    state.viewport_anchor.set(Some((centred_r, centred_c)));
}

/// Column anchor (`prev_start` for `visible_col_indices`) that centres the
/// cursor.
///
/// The cursor's position comes from `ui_core::cursor_display_col_index` rather
/// than being recomputed here: a column's `prev_start` indexes that function's
/// internal `filtered` list (every column except the reserved ones), which
/// spans the whole sheet — not the ~17-wide window `displayed_cols` returns.
/// Deriving it locally put the anchor hundreds of columns off and scrolled the
/// body off-screen, so the one canonical computation lives next to the code it
/// has to agree with.
fn centre_col_anchor(
    state: &GuiState,
    sheet: &crate::ops::SheetState,
    cursor: SheetCursor,
    dim: usize,
) -> usize {
    let lm = MARGIN_COLS;
    let pos = ui_core::cursor_display_col_index(sheet, cursor);

    // Never anchor before the body: centring pulls the window `dim / 2`
    // columns left of the cursor, and for a cursor near the body's left edge
    // that lands among the left-margin columns, pushing the body off the right
    // of the screen and filling the viewport with margin grey.
    let body_origin_pos = if state.last_col.get() < lm {
        0
    } else {
        // Columns before the body, minus those the helper treats as reserved.
        ui_core::cursor_display_col_index(sheet, SheetCursor { col: lm, ..cursor })
    };
    pos.saturating_sub(dim / 2).max(body_origin_pos)
}

fn handle_click(x: f64, y: f64, state_rc: &Rc<GuiState>) {
    let state: &GuiState = &**state_rc;
    // Mark that the user has clicked: clicks can arrive during `present()`'s
    // startup pump (see `GuiState::clicked`), and run_gui's post-present
    // selection must not clobber them.
    state.clicked.set(true);
    // Padlock hits first: padlocks live in the gutter chrome that plain
    // clicks ignore, and toggling a pin must not move the cursor, collapse
    // the selection, or start editing.
    //
    // The drawn glyph is only as big as a row/column header (it must not
    // overlap its neighbours), which on a phone is a small finger target. On
    // Android the hit rect is grown to a finger-sized square centred on the
    // glyph, clamped so it cannot spill into the next header cell and steal
    // its clicks: at most one header's worth of padding per axis.
    let hit = {
        let padlocks = state.padlocks.borrow();
        padlocks
            .iter()
            .find(|h| {
                let (hx, hy, hw, hh) = padlock_hit_rect(h);
                x >= hx && x < hx + hw && y >= hy && y < hy + hh
            })
            .copied()
    };
    if let Some(hit) = hit {
        toggle_pin(state, hit.is_row, hit.index);
        state.canvas.queue_redraw();
        return;
    }
    // While the in-grid dropdown is open it owns pointer input: a click on a
    // row commits it, a click anywhere else dismisses it (standard dropdown
    // behaviour) and the click is consumed.
    if agg_dropdown_open(state) {
        let (cw, ch) = state.canvas_size.get();
        match agg_drop_row_at(state, x, y, cw as f64, ch as f64) {
            Some(row) => {
                if let Some(d) = state.agg_drop.borrow_mut().as_mut() {
                    d.sel = row;
                }
                agg_dropdown_commit(state);
            }
            None => {
                hide_agg_dropdown(state);
                super::agg_picker::close(state.app_mut());
            }
        }
        state.canvas.queue_redraw();
        return;
    }

    let app = state.app_mut();
    if x < row_label_w() || y < header_h() {
        return;
    }
    let col_ixs: Vec<usize> = displayed_cols(state);
    let mc = app.core.workbook.active_sheet().grid.main_cols();
    let mut cx = row_label_w();
    for &c in &col_ixs {
        let cw = display_col_width(&app.core.workbook.active_sheet(), c, mc) as f64 * char_w();
        if x >= cx && x < cx + cw {
            let ri = ((y - header_h()) / row_h()) as usize;
            // Same pinned-first display set the renderer uses, or clicks
            // land on the wrong rows once pins are active.
            let display_rows: Vec<usize> = displayed_rows(state);
            if ri < display_rows.len() {
                let logical_row = display_rows[ri];
                // Clicking away from an in-progress edit commits it to the
                // cell it was entered in: `last_row`/`last_col` still point
                // there, and the `start_edit` below would otherwise clear
                // `edit_buf` and silently discard the typed value (Esc is the
                // cancel path — it remembers the text via `pending_lost_edit`,
                // which this path never did).
                let same_cell =
                    logical_row == state.last_row.get() && c == state.last_col.get();
                if state.editing.get() && !state.edit_buf.borrow().is_empty() {
                    if same_cell {
                        // Re-clicking the cell being edited keeps the text
                        // in flight rather than committing and restarting.
                        state.formula_entry.grab_focus();
                        state.canvas.queue_redraw();
                        return;
                    }
                    commit_edit(state);
                }
                state.last_row.set(logical_row);
                state.last_col.set(c);
                app.core.cursor.row = logical_row;
                app.core.cursor.col = c;
                // Clicking a margin aggregate key (the seeded TOTAL in `]A~1`
                // / `[A_1`, or any `==MAX` directive) offers the function
                // picker as an in-grid dropdown anchored to that cell: the
                // click both selects the cell and opens the list, so choosing
                // MAX/AVERAGE/… never requires knowing the vocabulary.
                // Ordinary margin text is not a key and clicks normally.
                {
                    let grid = &app.core.workbook.active_sheet().grid;
                    let hit = crate::addr::sheet_cursor_to_addr(
                        crate::addr::LogicalRow(logical_row),
                        crate::addr::GlobalCol(c),
                        crate::addr::MainRows(grid.main_rows()),
                        crate::addr::MainCols(grid.main_cols()),
                    );
                    if crate::gui::agg_picker::is_agg_key_cell(app, &hit)
                        && crate::gui::agg_picker::open_for(app, &hit)
                    {
                        app.core.anchor = None;
                        show_agg_dropdown(state_rc, &hit);
                        return;
                    }
                }
                // Plain click collapses any selection (fresh single-cell focus).
                app.core.anchor = None;
                // Clicking the trailing blank opens one more beyond it
                // (pointer arrival — see grow_blank_past_cursor). Shrink
                // allowed too (mouse users prune like keyboard users).
                maintain_extent(state, true);
                grow_blank_past_cursor(state);
                update_formula_bar(state, logical_row, c);
                // Bring the clicked cell to the middle of the viewport. Done
                // after `maintain_extent`/`grow_blank_past_cursor` because
                // those can change the display lists the centring indexes
                // into.
                centre_on_cursor(state, app);
                // Select (display the value), don't blank: typing replaces
                // via the cleared buffer, and an untouched commit stays
                // quiet (see `start_edit_keep_display`).
                start_edit_keep_display(state);
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
    let bar = super::menu::menu_bar();
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
    // Preset inserted whole: the caret follows it.
    state.edit_caret.set(text.chars().count());
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

/// Open the Insert > Special Char picker dialog over shared picker state
/// (ratatui/pancurses parity: same 10 items, order, arrows, digits,
/// Enter, Esc — see dialogs::special_char_dialog). Confirming splices the
/// choice into the edit: fresh cells snapshot the visible cell text first
/// (caret at end), mid-edit picks append to the buffer end — no
/// widget-caret API exists on every backend, so the ratatui mid-caret
/// splice degrades to append here. Cancel closes picker state and stages
/// nothing.
/// Return keyboard focus to the formula entry once the picker dialog has
/// fully torn down.
///
/// The dialog closes asynchronously (delete-event) and its own focus restore
/// runs *after* a synchronous grab, stealing focus back — the staged edit then
/// sits visible in the formula bar while every later key (including the Return
/// that should commit it) is dropped, so the picker looks like a no-op.
/// Deferring past teardown on GTK fixes it; backends with modeless dialogs
/// grab immediately.
/// The in-grid aggregate dropdown: drawn **into the grid canvas** rather
/// than a native child widget.
///
/// A native combo proved unusable: on GTK3 a child added to an overlay is
/// allocated 1x1 regardless of its size request (verified with
/// `gtk_widget_get_allocated_width`), so it was placed correctly but painted
/// nothing. Painting it on the canvas works identically on every backend
/// (GTK, Win32, any canvas), needs no reparenting, and can never desync from
/// the cell it is anchored to.
struct AggDrop {
    /// Cell the picker edits; committed on selection.
    target: crate::grid::CellAddr,
    /// True while the list is shown.
    open: bool,
    /// Highlighted row.
    sel: usize,
}

/// Rows the dropdown offers (shared vocabulary with the other backends).
fn agg_drop_rows() -> Vec<String> {
    crate::ui_core::agg_labelled_choices()
}

/// Pixel rectangle of the clicked margin-key cell, in **canvas-local**
/// coordinates, or `None` when it is scrolled out of view (the list must not
/// float over an unrelated cell).
/// Map an aggregate-key address to its `(display_row, global_col)`, the
/// inverse of the click mapping in [`handle_click`]. Pure so it is unit
/// testable without a live window (`GuiState` needs real widgets).
///
/// Every arm matters: `show_agg_dropdown` bails when this returns `None`, so
/// a missing variant silently kills the dropdown for that whole key class.
/// In particular **footer** keys (`[A_1`, `[A_2`, `[A_3`, … — the `[A_n`
/// aggregate key column, see `docs/tests/subtotal.corro`) live at
/// `HEADER_ROWS + main_rows + footer_row`; omitting them made `[A_2`/`[A_3`
/// resolve as keys but never pop the list.
fn agg_key_display_rc(
    addr: &crate::grid::CellAddr,
    main_cols: usize,
    main_rows: usize,
) -> Option<(usize, usize)> {
    use crate::grid::CellAddr;
    let rc = match addr {
        CellAddr::Header { row, col } => (*row as usize, col.to_global(main_cols)),
        CellAddr::Footer { row, col } => (
            HEADER_ROWS + main_rows + *row as usize,
            col.to_global(main_cols),
        ),
        CellAddr::Main { row, col } => (
            HEADER_ROWS + *row as usize,
            MARGIN_COLS + *col as usize,
        ),
        CellAddr::Left { row, col } => (
            HEADER_ROWS + *row as usize,
            *col,
        ),
        CellAddr::Right { row, col } => (
            HEADER_ROWS + *row as usize,
            MARGIN_COLS + main_cols + col,
        ),
    };
    Some(rc)
}

fn cell_rect(state: &GuiState, addr: &crate::grid::CellAddr) -> Option<(i32, i32, i32, i32)> {
    let app = state.app_ref();
    let grid = &app.core.workbook.active_sheet().grid;
    let mc = grid.main_cols();

    let (target_row, target_col) = agg_key_display_rc(addr, mc, grid.main_rows())?;

    let display_rows = displayed_rows(state);
    let ri = display_rows.iter().position(|&r| r == target_row)?;
    let col_ixs = displayed_cols(state);
    let ci = col_ixs.iter().position(|&c| c == target_col)?;

    let mut x = row_label_w();
    for &c in &col_ixs[..ci] {
        x += display_col_width(&app.core.workbook.active_sheet(), c, mc) as f64 * char_w();
    }
    let w = display_col_width(&app.core.workbook.active_sheet(), target_col, mc) as f64 * char_w();
    let y = header_h() + ri as f64 * row_h();
    Some((x as i32, y as i32, w as i32, row_h() as i32))
}

/// Geometry of the open dropdown: `(box_x, box_y, box_w, box_h, row_h)`,
/// anchored under the cell and clamped into the canvas.
fn agg_drop_layout(
    state: &GuiState,
    target: &crate::grid::CellAddr,
    canvas_w: f64,
    canvas_h: f64,
) -> Option<(f64, f64, f64, f64, f64)> {
    let (cx, cy, cw, _ch) = cell_rect(state, target)?;
    Some(agg_drop_layout_for_cell(cx, cy, cw, canvas_w, canvas_h))
}

/// Pure placement math for the dropdown panel, given the anchored cell's
/// canvas-local rectangle. Split out so the geometry is testable without a
/// live window (a `GuiState` needs real widgets).
fn agg_drop_layout_for_cell(
    cx: i32,
    cy: i32,
    cw: i32,
    canvas_w: f64,
    canvas_h: f64,
) -> (f64, f64, f64, f64, f64) {
    let rh = row_h();
    let rows = agg_drop_rows().len() as f64;
    let list_w = (cw as f64).max(150.0);
    let list_h = rows * rh + 2.0;
    // Open downward from the cell; flip above it when that would overflow,
    // then clamp the whole box inside the canvas so it is never partly
    // off-screen (a tall list on a short canvas scrolls less than it shows).
    let mut y = (cy as f64) + rh;
    if y + list_h > canvas_h {
        y = (cy as f64) - list_h;
    }
    let max_y = (canvas_h - list_h).max(0.0);
    let y = y.clamp(0.0, max_y);
    let max_x = (canvas_w - list_w).max(0.0);
    let x = (cx as f64).clamp(0.0, max_x);
    (x, y, list_w, list_h, rh)
}

/// The row under a canvas-local point, while the dropdown is open.
fn agg_drop_row_at(state: &GuiState, x: f64, y: f64, canvas_w: f64, canvas_h: f64) -> Option<usize> {
    let target = {
        let slot = state.agg_drop.borrow();
        match slot.as_ref() {
            Some(d) if d.open => d.target.clone(),
            _ => return None,
        }
    };
    let (bx, by, bw, bh, row_h) = agg_drop_layout(state, &target, canvas_w, canvas_h)?;
    if x < bx || x >= bx + bw || y < by || y >= by + bh {
        return None;
    }
    let idx = ((y - by) / row_h) as usize;
    (idx < agg_drop_rows().len()).then_some(idx)
}

/// Show the in-grid dropdown over `addr`. Returns false when the cell is not
/// currently visible (caller falls back to the dialog so the picker stays
/// reachable either way).
fn show_agg_dropdown(state: &Rc<GuiState>, addr: &crate::grid::CellAddr) -> bool {
    if cell_rect(state, addr).is_none() {
        return false;
    }
    let sel = super::agg_picker::index(state.app_ref()).unwrap_or(0);
    *state.agg_drop.borrow_mut() = Some(AggDrop {
        target: addr.clone(),
        open: true,
        sel,
    });
    state.canvas.queue_redraw();
    // The grid paints the list, so keyboard focus belongs to the canvas for
    // arrow/Enter handling.
    state.canvas.grab_focus();
    true
}

/// Hide the in-grid dropdown without committing.
fn hide_agg_dropdown(state: &GuiState) {
    if let Some(d) = state.agg_drop.borrow_mut().as_mut() {
        d.open = false;
    }
    state.canvas.queue_redraw();
}

/// Whether the in-grid dropdown is currently shown.
fn agg_dropdown_open(state: &GuiState) -> bool {
    state.agg_drop.borrow().as_ref().is_some_and(|d| d.open)
}

/// Move the dropdown highlight by `delta`, clamped to the row range.
fn agg_dropdown_step(state: &GuiState, delta: i32) {
    if let Some(d) = state.agg_drop.borrow_mut().as_mut() {
        if d.open {
            d.sel = crate::ui_core::agg_step_index(d.sel, delta);
        }
    }
    state.canvas.queue_redraw();
}

/// Commit the dropdown's highlighted row to its target cell and hide.
fn agg_dropdown_commit(state: &GuiState) {
    let (target, idx) = {
        let slot = state.agg_drop.borrow();
        let Some(d) = slot.as_ref() else { return };
        if !d.open {
            return;
        }
        (d.target.clone(), d.sel)
    };
    if let Some(directive) = crate::ui_core::agg_choice_directive(idx) {
        let app = state.app_mut();
        super::actions::commit_cell(app, target.clone(), directive.to_string());
        app.core.status = format!(
            "{} = {}",
            crate::addr::cell_ref_text(&target, app.core.workbook.active_sheet().grid.main_cols()),
            crate::agg::cell_display(&app.core.workbook.active_sheet().grid, &target)
        );
    }
    hide_agg_dropdown(state);
    super::agg_picker::close(state.app_mut());
    state.canvas.grab_focus();
}

/// Open the shared aggregate-picker state for the cursor and show it as the
/// in-grid dropdown. Returns false when nothing governs the cursor.
fn open_agg_picker_inline(state: &Rc<GuiState>) -> bool {
    let opened = {
        let app = state.app_mut();
        let cursor = app.core.cursor;
        super::agg_picker::open_for_cursor(app, &cursor)
    };
    if !opened {
        return false;
    }
    let target = state.app_ref().agg_picker_target.clone();
    match target {
        Some(addr) => show_agg_dropdown(state, &addr),
        None => false,
    }
}

/// Paint the open dropdown over the grid. Called at the end of
/// [`render_grid`] so it always sits above the cells.
fn render_agg_dropdown(dc: &mut dyn DrawContext, state: &GuiState, w: i32, h: i32) {
    let (target, sel) = {
        let slot = state.agg_drop.borrow();
        match slot.as_ref() {
            Some(d) if d.open => (d.target.clone(), d.sel),
            _ => return,
        }
    };
    let Some((bx, by, bw, bh, row_h)) = agg_drop_layout(state, &target, w as f64, h as f64) else {
        return;
    };

    // Panel: white fill, dark border, accent line along the top edge.
    dc.fill_rect(bx, by, bw, bh, 1.0, 1.0, 1.0, 1.0);
    dc.stroke_rect(bx, by, bw, bh, 0.25, 0.25, 0.3, 1.0, 1.0);
    dc.fill_rect(bx, by, bw, 1.0, 0.25, 0.25, 0.3, 1.0);

    for (i, row) in agg_drop_rows().iter().enumerate() {
        let ry = by + 1.0 + i as f64 * row_h;
        if i == sel {
            // Highlight the active row so it reads as the current choice.
            dc.fill_rect(bx + 1.0, ry, bw - 2.0, row_h, 0.82, 0.89, 0.98, 1.0);
        }
        dc.draw_text(bx + 6.0, ry + 3.0, row, "monospace", font_size(), 0.05, 0.05, 0.1, 1.0);
    }
}

fn restore_editor_focus(state: &Rc<GuiState>) {
    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork"), not(feature = "gtk4-rs")))]
    {
        let entry = state.formula_entry.clone();
        let window = state.window.clone();
        let _ = rswidgets::backends::gtk::timeout_add_once(
            150,
            Box::new(move || {
                // The toplevel itself lost activation to the dialog; without
                // re-presenting it, keys sent to the window are ignored even
                // though the entry is focused.
                window.present();
                entry.grab_focus();
            }),
        );
    }
    #[cfg(not(all(feature = "gtk", target_os = "linux", not(feature = "zork"), not(feature = "gtk4-rs"))))]
    {
        state.formula_entry.grab_focus();
    }
}

fn open_special_char_picker(state: &Rc<GuiState>) {
    // Dismiss any open menu grab FIRST: Alt+I,s leaves the Insert menu
    // open behind the dialog, and its grab would swallow every later key
    // (including the main-window Return that commits the staged edit).
    // Same as the user pressing Esc on the menu: closes the menu only.
    if state.menubar.get().map(|mb| mb.menu_active()).unwrap_or(false) {
        if let Some(mb) = state.menubar.get() {
            mb.handle_menu_key(ESCAPE, 0);
        }
    }
    let items: Vec<String> = super::special_picker::items().into_iter().collect();
    let initial = super::special_picker::index(state.app_ref()).unwrap_or(0);
    let shared = state.clone();
    dialogs::special_char_dialog(&items, initial, move |result| {
        let app = shared.app_mut();
        match result {
            Some(idx) => {
                super::special_picker::set(app, idx);
                if let Some(choice) = super::special_picker::take(&mut *app) {
                    if shared.editing.get() {
                        shared.edit_buf.borrow_mut().push_str(&choice);
                        sync_entry_to_buf(&shared);
                        restore_editor_focus(&shared);
                    } else {
                        // Snapshot the visible cell text first so the staged
                        // edit starts from what the user sees.
                        let grid = &app.core.workbook.active_sheet().grid;
                        let addr = crate::addr::sheet_cursor_to_addr(
                            crate::addr::LogicalRow(shared.last_row.get()),
                            crate::addr::GlobalCol(shared.last_col.get()),
                            crate::addr::MainRows(grid.main_rows()),
                            crate::addr::MainCols(grid.main_cols()),
                        );
                        let cur = grid.get(&addr).unwrap_or_default();
                        start_edit_with_text(&shared, &format!("{cur}{choice}"));
                        restore_editor_focus(&shared);
                    }
                }
            }
            None => {
                super::special_picker::close(app);
            }
        }
        refresh_after_dialog(&shared);
    });
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
        MenuDispatch::SpecialPicker => {
            open_special_char_picker(state);
        }
        MenuDispatch::AggregatePicker => {
            // The menu action already resolved and opened the picker state;
            // show it as the in-grid dropdown over the resolved key.
            let target = state.app_ref().agg_picker_target.clone();
            match target {
                Some(addr) if show_agg_dropdown(state, &addr) => {}
                _ => {
                    super::agg_picker::close(state.app_mut());
                    state.app_mut().core.status =
                        "Aggregate: the key is not visible; scroll to it first".into();
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

// The live GUI state, reachable from an `app.<name>` string.
//
// The Android menu strip and the iOS/macOS menu bars dispatch action names
// across an FFI boundary (neither platform can hand a Rust closure to a native
// menu), so the live state has to be reachable from a plain function. `run_gui`
// publishes its `Rc<GuiState>` here once the state exists; the menu UI is only
// built after that, and both platforms drive the UI on one thread, so a
// process-wide slot is sufficient.
//
// thread_local, not a static Mutex: GuiState holds Rc/Cell/RefCell and is
// deliberately !Send (the whole GUI runs on one thread), so a global would
// need an unsafe Send impl. The platform's UI thread is the only caller.
// Compiled on every backend: the mobile platform paths are the only callers,
// but a desktop build also compiles `gui::ios_backend` (the menu *model* is
// plain data — see the note there), and keeping one definition avoids a
// second cfg-gated copy of the dispatch table.
thread_local! {
    static MOBILE_MENU_STATE: std::cell::RefCell<Option<std::rc::Rc<GuiState>>> =
        const { std::cell::RefCell::new(None) };
}

/// Publish the live state for [`dispatch_mobile_menu_action`]. Called once
/// per `run_gui`; a second call (a second window) replaces the first, which
/// matches the single-activity/single-scene model on both platforms.
// Called only from the mobile/macOS menu blocks below (each cfg-gated), so
// gate the definition to match: a plain desktop build compiled it and then
// warned that nothing used it.
#[cfg(any(target_os = "android", target_os = "ios", target_os = "macos"))]
pub(crate) fn publish_mobile_menu_state(state: &Rc<GuiState>) {
    MOBILE_MENU_STATE.with(|s| *s.borrow_mut() = Some(state.clone()));
}

/// Dispatch an `app.<name>` action from the platform's menu UI. Unknown
/// names are ignored (the menu may be built from a newer menu tree than the
/// running handler knows).
pub(crate) fn dispatch_mobile_menu_action(action: &str) {
    let name = action.strip_prefix("app.").unwrap_or(action);
    let state = MOBILE_MENU_STATE.with(|s| s.borrow().clone());
    match state {
        Some(state) => handle_menu_action(name, &state),
        None => eprintln!("mobile menu action before state published: {name}"),
    }
}

/// Android-facing alias for [`dispatch_mobile_menu_action`] (see above).
#[cfg(target_os = "android")]
pub(crate) fn dispatch_android_menu_action(action: &str) {
    dispatch_mobile_menu_action(action);
}

/// Scroll the sheet viewport by whole rows/columns (touch drag on mobile).
///
/// There is no independent scroll offset in this GUI: `displayed_rows` /
/// `displayed_cols` derive the viewport purely from `app.core.cursor`
/// (`prev_start` is always 0), so the cursor *is* the viewport origin. A
/// touch drag therefore moves the cursor, exactly as a scrollbar drag does
/// through [`scroll_to_cursor`] — one code path for "viewport moved", and
/// the selected cell stays visible for free.
///
/// `d_rows`/`d_cols` are whole-cell counts; the caller accumulates the
/// sub-cell remainder so a slow drag still scrolls. Movement is clamped into
/// the same domain the scrollbars address, so a drag can never fling the
/// cursor past the end of the sheet. `grow_blank_past_cursor` is
/// deliberately NOT called: dragging is navigation over existing content,
/// and growing the grid on every drag frame would extend the sheet without
/// bound (the scrollbar path grows only on an explicit thumb drag).
// Reached only through `android_backend::scroll_viewport` (cfg-gated on
// Android), where the touch-drag path lives.
#[cfg(target_os = "android")]
pub(crate) fn scroll_viewport_by_cells(d_rows: i32, d_cols: i32) {
    let state = match MOBILE_MENU_STATE.with(|s| s.borrow().clone()) {
        Some(state) => state,
        None => {
            eprintln!("mobile scroll before state published");
            return;
        }
    };
    if d_rows == 0 && d_cols == 0 {
        return;
    }
    // An in-progress cell edit owns the selection; scrolling under it would
    // move the cursor away from the edit target. Commit first (the guard
    // every other navigation path uses) so a drag behaves like any other
    // navigation that interrupts typing.
    if state.editing.get() {
        commit_edit(&state);
    }
    let (ru, cu) = scroll_domain(&state);
    let row0 = state.last_row.get();
    let col0 = state.last_col.get();
    let mut row = row0;
    let mut col = col0;
    if d_rows != 0 {
        let max_row = HEADER_ROWS + ru.saturating_sub(1);
        row = (row0 as i64 + d_rows as i64).clamp(HEADER_ROWS as i64, max_row as i64) as usize;
    }
    if d_cols != 0 {
        let max_col = MARGIN_COLS + cu.saturating_sub(1);
        col = (col0 as i64 + d_cols as i64).clamp(MARGIN_COLS as i64, max_col as i64) as usize;
    }
    if row == row0 && col == col0 {
        return; // already at the edge: no redraw churn per drag event
    }
    update_state_cursor(&state, row, col);
}

/// Where the drawn pointer should sit for a menu tour stop, in **canvas**
/// coordinates (see `paint_movie_pointer` for why that is not the same as
/// window coordinates).
///
/// The menu bar is a real GTK widget above the canvas, so the pointer cannot be
/// drawn over it; the arrow is aimed at the top edge of the canvas beneath the
/// section's button, which reads as pointing at the open menu.
///
/// The horizontal position is estimated from each label's rendered width. The
/// menu bar lays its buttons out left to right in `menu::menu_bar()` order with
/// a little padding, which is stable enough for a demo pointer; a drift of a
/// few pixels under a button is not noticeable, and the item activation does
/// not depend on the pointer's position at all.
fn menu_section_pointer_target(section: &str, item_index: u32) -> (f64, f64) {
    // Approximate width of a menu label at the default metrics.
    let label_w = |s: &str| (s.chars().count() as f64 + 2.0) * char_w();
    let mut x = 0.0;
    for root in super::menu::menu_bar() {
        let w = label_w(root.label);
        if root.label == section {
            // Down into the popover far enough to sit beside the item row that
            // is about to activate (rows are about one row-height tall).
            let y = (item_index as f64 + 0.5) * row_h() - 2.0 * header_h();
            return (x + w / 2.0, y.max(4.0));
        }
        x += w;
    }
    // Unknown section: park the pointer at the left of the bar.
    (20.0, 4.0)
}

/// Open a menu-bar section and activate one of its items, for the movie tour.
///
/// The popover is opened through the real GTK mnemonic path — the same thing a
/// user pressing Alt+letter triggers — because that is the one input route that
/// provably reaches this widget tree under a synthetic driver: keys arrive,
/// while synthetic X11 pointer events do not reach GTK4's gesture handling at
/// all (verified by capture; see `movie_pointer`). The item itself is then
/// dispatched by name through [`handle_menu_action`], the same entry point the
/// menu's own action callbacks use, so the tour exercises the production path
/// rather than a demo-only shortcut.
///
/// A missing section or item is reported in the status line instead of being
/// silently ignored, so a typo in a tour script is visible in the recording.
fn run_menu_tour_stop(state: &Rc<GuiState>, section: &str, item: &str) {
    // Resolve the item to its action name from the shared menu definition (the
    // same tree every backend builds from, so the tour cannot drift from the
    // real menus).
    let bar = super::menu::menu_bar();
    let Some(root) = bar.iter().find(|r| r.label == section) else {
        state.app_mut().core.status = format!("Menu tour: no section {section:?}");
        sync_chrome_labels(state);
        return;
    };
    // Search nested submenus too (`File > Width > Default width`,
    // `Format > Scope > All`), so a tour can address items at any depth.
    fn find_action(items: &[super::menu::MenuAction], label: &str) -> Option<&'static str> {
        for i in items {
            if i.label == label && i.submenu.is_none() {
                return Some(super::menu::action_kind_to_name(i.action));
            }
            if let Some(sub) = i.submenu.as_deref() {
                if let Some(found) = find_action(sub, label) {
                    return Some(found);
                }
            }
        }
        None
    }
    let action = find_action(root.submenu.as_deref().unwrap_or(&[]), item);
    let Some(action) = action else {
        state.app_mut().core.status = format!("Menu tour: no item {section} ▸ {item}");
        sync_chrome_labels(state);
        return;
    };

    // Open the section's popover for real, so the recording shows the menu.
    if let Some(menubar) = state.menubar.get() {
        if let Some(c) = root.label.chars().next() {
            let keyval = c.to_ascii_lowercase() as u32;
            let _ = menubar.activate_submenu_by_mnemonic(keyval);
        }
    }

    // The status write has to be scoped: `app_mut()` is a `RefCell` borrow and
    // `sync_chrome_labels`/`handle_menu_action` take it again internally, so
    // holding it across those calls panics inside the draw callback.
    {
        let app = state.app_mut();
        app.core.status = format!("Menu tour: {section} ▸ {item}");
    }
    sync_chrome_labels(state);
    handle_menu_action(action, state);

    // Close so the next stop starts from a clean menu state.
    if let Some(menubar) = state.menubar.get() {
        menubar.menu_close();
    }
    state.canvas.queue_redraw();
    state.window.queue_redraw();
}

fn handle_menu_action(name: &str, state: &Rc<GuiState>) {
    let app = state.app_mut();
    log_ui_action("menu_action", name);
    match name {
        "open" => {
            if let Some(path) = dialogs::file_open_dialog() {
                match crate::io::load_workbook_file(&path) {
                    Ok(workbook) => {
                        app.core.workbook = workbook;
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
                match crate::io::write_workbook_log(
                    &path,
                    &app.core.workbook,
                    &app.core.persisted_view_sort_cols,
                ) {
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
        "follow_hyperlink" => {
            // Shared follow-link logic (same status texts as
            // pancurses/ratatui): opens the cursor cell's hyperlink.
            delegate_shared_action(name, state);
        }
        "edit_workbook_external" => {
            // Shared external-workbook logic: dispatch launches the GUI
            // editor detached on Backend::Gui (probed lightest-first) and
            // the save is picked up by the log-tail poll/tick, like another
            // window's Save. (Cell-text "edit_external" stays unwired: its
            // blocking $EDITOR roundtrip cannot run without a terminal.)
            delegate_shared_action(name, state);
        }
        "delete_cell" => {
            handle_delete(state);
        }
        "select_all" | "toggle_headers" | "toggle_margins" | "new_sheet" => {
            // Shared selection/chrome/sheet logic (same as pancurses/ratatui).
            delegate_shared_action(name, state);
        }
        "new_file" => {
            // Shared blank-document logic (same as pancurses/ratatui).
            // The trailing refresh recomputes the viewport for the new
            // workbook dims and syncs tabs, formula bar, and chrome.
            delegate_shared_action(name, state);
        }
        "export_tsv" | "export_csv" | "export_ods" | "export_ascii" | "export_all" => {
            // Shared export logic writes the file (same as pancurses/ratatui);
            // the export dialog supplies a type-filtered destination path
            // (format extension appended when the user types a bare name).
            let st = state.clone();
            let action = name.to_string();
            if let Some(path) = dialogs::file_export_dialog(&action) {
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
        "insert_special_chars" => {
            // 10-choice picker dialog over shared picker state (same as
            // pancurses/ratatui); arrives as MenuDispatch::SpecialPicker.
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
        // Non-prompt leaves delegate to shared dispatch (which is total over
        // the menu tree, so a newly-wired dispatch action works here with no
        // per-backend arm — Edit ▸ Workbook (External) shipped broken until
        // exactly this fallback existed). Prompt-gated leaves stay loud:
        // shared dispatch only answers their prompts via dialogs, so
        // delegating them would drop the prompt silently.
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
            if menu_action_needs_prompt(name).is_some() {
                app.core.status = format!("Menu action: {name}");
            } else {
                delegate_shared_action(name, state);
            }
        }
    }
    // Refresh the formula bar — unless an Edit{value} action just preset an
    // edit buffer (Insert Date/Time). update_formula_bar copies the grid cell
    // into the entry widget, which would wipe the preset (and the entry's
    // change handler could then eat edit_buf too). The chrome labels
    // (hints + formula status suffix) still need the new text.
    if state.editing.get() {
        sync_chrome_labels(state);
    } else {
        update_formula_bar(state, state.last_row.get(), state.last_col.get());
    }
}

// ---------------------------------------------------------------------------
// Formula entry change callback
// ---------------------------------------------------------------------------

fn on_formula_entry_changed(state: &GuiState) {
    // Android (TextWatcher) and iOS (UITextField `editingChanged`) have no
    // key-event path for soft-keyboard typing: the text-changed shim is the
    // ONLY signal that the user typed. Desktop sets
    // `editing` from its key handlers before the entry text changes, so a
    // non-editing change there is always programmatic (formula refresh)
    // and must be ignored. On Android, entry text that differs from the
    // committed cell value can only be user input, so adopt it as a fresh
    // edit (mirrors start_edit_with on first keystroke).
    #[cfg(any(target_os = "android", target_os = "ios", target_os = "macos"))]
    if !state.editing.get() {
        if let Some(text) = state.formula_entry.get_text() {
            let app = state.app_ref();
            let grid = &app.core.workbook.active_sheet().grid;
            let addr = crate::addr::sheet_cursor_to_addr(
                crate::addr::LogicalRow(state.last_row.get()),
                crate::addr::GlobalCol(state.last_col.get()),
                crate::addr::MainRows(grid.main_rows()),
                crate::addr::MainCols(grid.main_cols()),
            );
            let committed = grid.get(&addr).unwrap_or_default();
            if !text.is_empty() && text != committed {
                state.editing.set(true);
            } else {
                return;
            }
        } else {
            return;
        }
    }
    #[cfg(not(any(target_os = "android", target_os = "ios", target_os = "macos")))]
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

/// TEMPORARY Win95 diagnosis: append bytes to c:\gcorro.log via raw
/// CreateFileA (std::fs is broken on 9x: CreateFileW stub, error 120).
#[cfg(all(target_family = "rust9x", target_env = "msvc"))]
unsafe fn mark95(s: &[u8]) {
    use std::os::raw::c_void;
    unsafe extern "system" {
        fn CreateFileA(name: *const u8, access: u32, share: u32, sa: *mut c_void,
            disp: u32, flags: u32, tmpl: *mut c_void) -> *mut c_void;
        fn SetFilePointer(h: *mut c_void, lo: i32, hi: *mut i32, how: u32) -> u32;
        fn WriteFile(h: *mut c_void, buf: *const u8, len: u32, w: *mut u32, ov: *mut c_void) -> i32;
        fn CloseHandle(h: *mut c_void) -> i32;
    }
    let h = CreateFileA(b"c:\\gcorro.log\0".as_ptr(), 0x4000_0000, 1,
        std::ptr::null_mut(), 4, 0x80, std::ptr::null_mut());
    if h.is_null() || h as isize == -1 {
        return;
    }
    SetFilePointer(h, 0, std::ptr::null_mut(), 2);
    let mut w = 0u32;
    WriteFile(h, s.as_ptr(), s.len() as u32, &mut w, std::ptr::null_mut());
    CloseHandle(h);
}

/// TEMPORARY Win95 diagnosis: log parent/class/rect/visible/text-len of one
/// hwnd into c:\gcorro.log. `tag` is exactly 5 bytes (e.g. *b"main").
#[cfg(all(target_family = "rust9x", target_env = "msvc"))]
unsafe fn probe95(hwnd: *mut std::os::raw::c_void, tag: [u8; 5]) {
    use std::os::raw::c_void;
    unsafe extern "system" {
        fn GetParent(h: *mut c_void) -> *mut c_void;
        fn GetClassNameA(h: *mut c_void, buf: *mut u8, max: i32) -> i32;
        fn GetWindowRect(h: *mut c_void, r: *mut [i32; 4]) -> i32;
        fn IsWindowVisible(h: *mut c_void) -> i32;
        fn GetWindowTextLengthA(h: *mut c_void) -> i32;
    }
    let hx = b"0123456789abcdef";
    let mut msg = [0u8; 110];
    let mut p = 0;
    for i in 0..5 { msg[p] = tag[i]; p += 1; }
    msg[p] = b' '; p += 1;
    let par = GetParent(hwnd);
    for v in [hwnd as usize, par as usize] {
        for sh in [28u32, 24, 20, 16] {
            msg[p] = hx[((v >> sh) & 0xf) as usize]; p += 1;
        }
        msg[p] = b' '; p += 1;
    }
    let mut cls = [0u8; 24];
    let cl = GetClassNameA(hwnd, cls.as_mut_ptr(), 24);
    let mut i = 0;
    while i < cl && p < 70 { msg[p] = cls[i as usize]; p += 1; i += 1; }
    msg[p] = b' '; p += 1;
    msg[p] = if IsWindowVisible(hwnd) != 0 { b'V' } else { b'h' }; p += 1;
    msg[p] = b' '; p += 1;
    let mut r = [0i32; 4];
    GetWindowRect(hwnd, &mut r);
    for v in [r[0], r[1], r[2], r[3]] {
        let v = v as u32;
        for sh in [12u32, 8, 4, 0] {
            msg[p] = hx[((v >> sh) & 0xf) as usize]; p += 1;
        }
        msg[p] = b','; p += 1;
    }
    msg[p] = b'\n'; p += 1;
    mark95(&msg[..p]);
}

pub fn run_gui(corro_app: &mut super::App) -> Result<(), Box<dyn std::error::Error>> {
    run_gui_with_movie(corro_app, None)
}

/// Run the GUI, optionally replaying a `--movie` through the live window.
///
/// `--movie` is not a separate rendering path: it is this same window, same
/// widget tree, same draw callbacks, driven by a timer that applies one movie
/// step at a time. The only difference from an interactive session is where
/// the state changes come from — a log instead of a keyboard.
pub fn run_gui_with_movie(
    corro_app: &mut super::App,
    mut movie: Option<super::movie::GuiMovie>,
) -> Result<(), Box<dyn std::error::Error>> {
    rswidgets::core::install_debug_crash_handlers();
    phase("run_gui: crash handlers installed");
    // TEMPORARY Win95 diagnosis: startup progression (see mark95 below).
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe {
        mark95(b"nwgpre\n");
    }
    phase("run_gui: App::init");
    let rxapp = rswidgets::App::init()
        .map_err(|e| format!("GUI init failed: {e}"))?;
    phase("run_gui: App::init ok");
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe {
        mark95(b"nwgpost\n");
    }

    phase("run_gui: new_window");
    let win = rxapp.new_window()?;
    phase("run_gui: new_window ok");
    // TEMPORARY Win95 diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe {
        mark95(b"winpost\n");
    }
    win.set_title(&format!("corro {}", env!("CARGO_PKG_VERSION")));
    // TEMPORARY ReactOS diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"m-title\n"); }
    // TEMPORARY Win95 diagnosis: fit the 640x480 VM screen (release keeps
    // 1200x800 for real displays).
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    win.set_default_size(620, 420);
    #[cfg(not(all(target_family = "rust9x", target_env = "msvc")))]
    win.set_default_size(1200, 800);

    // TEMPORARY ReactOS diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"m-size\n"); }
    phase("run_gui: new_box");
    let vbox = rxapp.new_box(Orientation::Vertical, 0)?;
    phase("run_gui: new_box ok");
    // TEMPORARY ReactOS diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"m-box\n"); }

    // Fit column widths to rendered content
    corro_app.fit_main_columns_to_max_width();
    // TEMPORARY ReactOS diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"m-fit\n"); }

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
    // TEMPORARY ReactOS diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"m-fbar\n"); }
    let addr_label = rxapp.new_label("A1")?;
    // TEMPORARY ReactOS diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"m-alab\n"); }
    let f_label = rxapp.new_label("  fx  ")?;
    // TEMPORARY ReactOS diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"m-flab\n"); }
    let formula_entry = rxapp.new_entry()?;
    // TEMPORARY ReactOS diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"m-fentry\n"); }
    formula_entry.set_hexpand(true);
    formula_bar.append(&addr_label);
    formula_bar.append(&f_label);
    formula_bar.append(&formula_entry);
    formula_bar.set_child_hexpand(&formula_entry, true);
    let formula_status = rxapp.new_label("")?;
    formula_status.set_visible(false);
    // TEMPORARY Win95 diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"m-entry\n"); }
    formula_bar.append(&formula_status);

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

    // Bottom strip: the shared hints line (ratatui parity).
    // TEMPORARY Win95 diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"m-canvas\n"); }
    let hints_label = rxapp.new_label("Ready")?;

    // Sheet tab strip (below the grid): one tab per sheet once the workbook
    // has 2+ sheets. Hidden until then — the first New sheet creates the
    // bar, later sheets just append tabs. Fixed strip height; a vertical box
    // stretches children across the full width on every backend.
    let tabbar = rxapp.new_canvas()?;
    tabbar.set_size_request(1, tab_h() as i32);
    tabbar.set_visible(false);

    let shared = Rc::new(GuiState {
        app: corro_app as *mut super::App,
        rxapp: rxapp.clone(),
        window: win.clone(),
        menubar: std::cell::OnceCell::new(),
        canvas: canvas.clone(),
        formula_entry: formula_entry.clone(),
        addr_label: addr_label.clone(),
        hints_label: hints_label.clone(),
        formula_status: formula_status.clone(),
        editing: Cell::new(false),
        edit_buf: RefCell::new(String::new()),
        edit_caret: Cell::new(0),
        agg_drop: RefCell::new(None),
        canvas_size: Cell::new((0, 0)),
        entry_snapshot: RefCell::new(String::new()),
        entry_clicked: Cell::new(false),
        movie_pointer: Cell::new(None),
        entry_shown: Cell::new((usize::MAX, usize::MAX)),
        mode: Cell::new(GuiMode::Normal),
        last_row: Cell::new(cursor_row),
        last_col: Cell::new(cursor_col),
        data_rows: Cell::new(data_rows),
        data_cols: Cell::new(data_cols),
        last_key: Cell::new(0),
        key_counter: Cell::new(0),
        clicked: Cell::new(false),
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
        viewport_anchor: Cell::new(None),
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
    // TEMPORARY Win95 diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"m-menu\n"); }
    let _ = shared.menubar.set(menubar.clone());
    let menubar_cb = menubar.clone();
    // Android: the GTK menubar above is a model only (the Android adapter's
    // create_menubar has no view), so a phone would show no menus at all.
    // Build the real strip: a top bar with quick actions and an overflow
    // popup carrying every top-level menu, dispatching the same `app.*`
    // action names the desktop build registers.
    #[cfg(target_os = "android")]
    {
        // Publish the state before building the strip: the strip's buttons
        // dispatch back into handle_menu_action by name.
        publish_mobile_menu_state(&shared);
        // A missing/mismatched MenuStrip class must not take the whole UI
        // down: the sheet is still usable without menus.
        if let Err(e) = super::android_backend::install_menu_strip() {
            eprintln!("android menu strip unavailable: {e}");
        }
    }
    // iOS: same story with UIKit spelling — publish the state, then let the
    // host's menu bar (a `UIBarButtonItem`/`UIMenu` the Swift side builds
    // from this model) dispatch the same `app.*` names back.
    #[cfg(target_os = "ios")]
    {
        publish_mobile_menu_state(&shared);
        if let Err(e) = super::ios_backend::install_menu_model(&rxapp, &shared) {
            eprintln!("ios menu model unavailable: {e}");
        }
    }
    // macOS: a real `NSMenu` can be built from the same model, so the host
    // gets the model and renders a genuine menubar (unlike iOS's overflow
    // item). Same publication, same `app.*` names.
    #[cfg(target_os = "macos")]
    {
        publish_mobile_menu_state(&shared);
        if let Err(e) = super::macos_backend::install_menu_model(&rxapp, &shared) {
            eprintln!("macos menu model unavailable: {e}");
        }
    }
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

    // Pointer click into the formula entry arms click-to-edit repair: the
    // next entry keystroke restores the displayed formula around a fresh
    // single-char push instead of losing it. The click still focuses and
    // places the caret natively; nothing is consumed here. No-op on
    // backends without entry click support.
    let shared_click = shared.clone();
    // Arm click-to-edit when the entry gains focus. A pointer press into the
    // entry focuses it, and `focus-in-event` is a plain GtkEntry signal that
    // is actually delivered on GTK3 — unlike `button-press-event`, which
    // needs a GDK event mask on the entry's own GdkWindow and never fires
    // here. A click still places the caret natively; we only need to notice
    // that the user is now editing this text.
    formula_entry.connect_button_press(move || {
        // The user pressed the pointer in the formula entry: adopt the caret
        // the click placed so a following insert lands at that position
        // rather than appending.
        pull_caret_from_entry(&shared_click);
        shared_click.entry_clicked.set(true);
    })?;


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
                    // Counted already (window handled it, or a menu consumed
                    // it): skip, exactly once. The window layer owns all
                    // buffer updates (including click-to-edit adopt), so the
                    // entry never reconciles here — it only avoids doubling.
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
    vbox.append(&scrolled);
    // The nwg manual box layout looks the child up by handle, so it needs
    // the flag set AFTER append as well (a no-op repeat everywhere else).
    vbox.set_child_vexpand(&scrolled, true);
    // Sheet tabs sit below the grid (above the status line), like the
    // terminal reference rendering tabs in its bottom row.
    vbox.append(&tabbar);
    vbox.append(&hints_label);

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
                (w as f64 - row_label_w()).max(0.0) as i32,
            ));
            shared_draw.data_rows.set(rows_to_fill_px(h));
            // Keep the scrollbar thumb on the viewport (ranges track grid
            // growth here too).
            sync_scrollbars(&shared_draw);
        }
        shared_draw.canvas_size.set((w, h));
        render_grid(dc, &shared_draw, w, h);
        // The in-grid aggregate dropdown paints last so it sits above the
        // cells (and above the header/margin chrome).
        render_agg_dropdown(dc, &shared_draw, w, h);
        // The synthetic movie pointer goes last of all: it has to be visible
        // over the grid, the chrome and any open dropdown.
        dc.clip(0.0, 0.0, w as f64, h as f64);
        paint_movie_pointer(dc, &shared_draw);
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
    // Initial chrome sync (ratatui parity): event paths refresh the formula
    // row and tab strip, but nothing runs between construction and the
    // first key/click — without this the labels keep constructor text
    // ("Ready", "A1") on the first paint and any startup status (e.g. a
    // template note) is invisible until the user acts. Runs BEFORE
    // present(): present() pumps events for a very long time on slow
    // displays (so long that tests only ever observe the pumped phase),
    // and anything placed after it never runs there. GtkLabel applies
    // set_text/visibility on realize, so pre-present sync paints correctly.
    // Mirror other windows (registered BEFORE present(): present() pumps
    // events for a very long time on slow displays, and anything placed after
    // it may never run at all — which is exactly why this tick never fired). `handle_key` also polls, but that only runs when
    // a key arrives here: two windows editing the same file did not update
    // each other until you pressed a key in the one you were watching. Poll
    // on a timer instead, matching the TUI loop (which calls `sync_external`
    // every iteration). Commits append to the log immediately, so no save
    // step is needed on either side.
    {
        let tick_state = shared.clone();
        match rswidgets::add_periodic_tick(
            &win,
            250,
            Box::new(move || {
                let changed = matches!(
                    tick_state.app_mut().core.poll_log_tail(),
                    Ok(true)
                );
                if changed {
                    update_formula_bar(&tick_state, HEADER_ROWS, MARGIN_COLS);
                    sync_chrome_labels(&tick_state);
                    sync_tabbar(&tick_state);
                    tick_state.canvas.queue_redraw();
                }
                true // keep ticking for the window's lifetime
            }),
        ) {
            Ok(()) => {}
            // No timer on this backend: the window still mirrors on
            // keystrokes (handle_key), just not while idle.
            Err(e) => eprintln!("periodic log-tail poll unavailable: {e}"),
        }
    }

    update_formula_bar(&shared, shared.last_row.get(), shared.last_col.get());
    sync_tabbar(&shared);
    formula_entry.grab_focus();
    eprintln!("PHASE: about_to_present");
    win.present();
    // TEMPORARY Win95 diagnosis: probe each known window (parent/class/
    // rect/visible) to find where the controls really live.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe {
        probe95(win.hwnd(), *b"main ");
        probe95(*formula_entry.inner.as_ref(), *b"entry");
        probe95(*canvas.inner.as_ref(), *b"canv ");
        probe95(*addr_label.inner.as_ref(), *b"label");
    }
    // `--movie`: replay through this very window. The timer applies one step per
    // tick and queues a redraw, so every frame comes from the normal draw
    // callbacks and a recording cannot drift from the app. Armed here, before
    // the blocking pump loops below, so replay starts as soon as the window is
    // mapped instead of after several seconds of event pumping.
    if let Some(movie) = movie.take() {
        arm_movie_driver(&shared, movie);
    } else {
        // A plain `--gui` session can also be scripted, so a demo can show two
        // windows *editing* one file rather than replaying it.
        arm_edit_script(&shared);
    }

    eprintln!("PHASE: after_present");
    let _ = std::fs::write("/tmp/gui_setup_phase3.txt", "after_present\n");

    // Queue an explicit redraw on the toplevel window: on GTK4 with the
    // Cairo GSK renderer, a canvas-only queue_draw may not cascade to
    // the toplevel's frame clock.  Marking the window dirty ensures the
    // frame clock is armed before the start_edit() canvas queue_redraw.
    win.queue_redraw();
    // TEMPORARY Win95 diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"m-queued\n"); }

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
    match startup_edit_action(
        shared.editing.get() && !shared.edit_buf.borrow().is_empty(),
        shared.clicked.get(),
    ) {
        StartupEdit::KeepInFlight => {
            shared.formula_entry.grab_focus();
            shared.canvas.queue_redraw();
        }
        StartupEdit::KeepClicked => {
            // A click during present()'s pump already chose a cell. If it
            // opened the aggregate dropdown, leave everything alone — running
            // the A1 selection here would move the yellow edit highlight and
            // `hide_agg_dropdown` would close the list the click just opened
            // (the "clicked [A_n but [A1 turned yellow" bug).
            if agg_dropdown_open(&shared) {
                shared.canvas.queue_redraw();
            } else {
                // The click selected a cell (or landed in the formula entry).
                // Still open the edit session so the first keystroke edits the
                // selected cell rather than starting a fresh replace — this is
                // what the startup path is for, and skipping it left
                // `editing=false`, so typing replaced the displayed value.
                start_edit_keep_display(&shared);
            }
        }
        StartupEdit::SelectA1 => {
            // Select A1 for typing (display its value); see the click path.
            start_edit_keep_display(&shared);
        }
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
    // TEMPORARY Win95 diagnosis (bindings done, entering warm-up).
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"m-bound\n"); }

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
    // TEMPORARY Win95 diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"pre-force\n"); }

    // Move rxapp.run() earlier — before pump_events — so the main loop
    // pointer is available for quit_main_loop before any user interaction.
    // TEMPORARY Win95 diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"pre-run\n"); }
    rxapp.run()?;
    // TEMPORARY Win95 diagnosis (unreachable if run() loops until quit).
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95(b"post-run\n"); }
    Ok(())
}

// ---------------------------------------------------------------------------
// `--movie` driver
// ---------------------------------------------------------------------------

/// Drive a scripted *editing* session through the live window.
///
/// Unlike `arm_movie_driver` (which replays a finished log), this makes the
/// window *produce* edits via the ordinary commit path, so a second window
/// watching the same file sees them arrive. Used by the two-window demo.
fn arm_edit_script(state: &Rc<GuiState>) {
    let steps = super::movie::edit_script_from_env();
    if steps.is_empty() {
        return;
    }
    let step_count = steps.len();
    let state_for_tick = state.clone();
    let idx = std::cell::Cell::new(0usize);
    let start = std::time::Instant::now();

    // Reveal each value one character at a time, then commit it — the same
    // shape the `--movie` driver uses. Committing the whole string in one tick
    // (the earlier behaviour) made a two-window recording show values simply
    // appearing, so neither window ever looked like it was being typed into.
    // The per-character delay reuses `--movie-typing-cps`; the tick is fixed and
    // the elapsed time decides when each character is due, so the animation
    // cannot drift when a tick is late.
    let options = super::movie::GuiMovieOptions::from_env();
    let char_ms = options.char_delay_ms().max(1.0);
    let tick_ms = (char_ms / 4.0).round().max(1.0) as u32;
    eprintln!("[corro] edit-script typing {} cps ({char_ms}ms/char)", options.typing_cps);

    let tick = move || -> bool {
        let i = idx.get();
        let Some(step) = steps.get(i) else {
            return false;
        };
        let due = |at: f64| start.elapsed().as_secs_f64() * 1000.0 >= at;
        if !due(step.at_ms as f64) {
            return true; // not due yet
        }
        let addr = crate::grid::CellAddr::Main { row: step.row, col: step.col };
        let chars: Vec<char> = step.value.chars().collect();
        // How far the typing should have got by now. The step's own timestamp is
        // when it starts; characters follow at the configured rate.
        let elapsed = start.elapsed().as_secs_f64() * 1000.0 - step.at_ms as f64;
        let shown_n = ((elapsed / char_ms).floor() as usize + 1).min(chars.len());
        let shown: String = chars[..shown_n].iter().collect();

        {
            let app = state_for_tick.app_mut();
            app.core.cursor = crate::grid::SheetCursor {
                row: crate::grid::HEADER_ROWS + step.row as usize,
                col: crate::grid::MARGIN_COLS + step.col as usize,
            };
            let grid = &app.core.workbook.active_sheet().grid;
            app.core.cursor.clamp(grid);
            app.core.status = format!(
                "{} = {shown}",
                crate::addr::cell_ref_text(&addr, grid.main_cols()),
            );
            // The in-progress text lives in the entry buffer, which is what the
            // formula bar and the grid's edit overlay both paint from.
            state_for_tick.editing.set(true);
            *state_for_tick.edit_buf.borrow_mut() = shown.clone();
            sync_entry_to_buf(&state_for_tick);
            update_state_cursor(&state_for_tick, app.core.cursor.row, app.core.cursor.col);
            sync_chrome_labels(&state_for_tick);
            state_for_tick.canvas.queue_redraw();
            state_for_tick.window.queue_redraw();
        }

        if shown_n >= chars.len() {
            // Commit once the last character has landed. This is the real
            // commit path: it appends to the log, which is what makes the other
            // window's tail pick the value up.
            let app = state_for_tick.app_mut();
            super::actions::commit_cell(app, addr.clone(), step.value.clone());
            app.core.state = app.core.workbook.active_sheet().clone();
            state_for_tick.editing.set(false);
            state_for_tick.edit_buf.borrow_mut().clear();
            update_state_cursor(&state_for_tick, app.core.cursor.row, app.core.cursor.col);
            sync_chrome_labels(&state_for_tick);
            state_for_tick.canvas.queue_redraw();
            state_for_tick.window.queue_redraw();
            idx.set(i + 1);
        }
        true
    };
    match rswidgets::add_periodic_tick(&state.window, tick_ms, Box::new(tick)) {
        Ok(()) => eprintln!("[corro] edit script armed ({step_count} steps, {tick_ms}ms tick)"),
        Err(e) => eprintln!("[corro] edit-script timer unavailable: {e}"),
    }
}

/// Drive a `--movie` replay through the live window.
///
/// One step per tick, so the window paints each intermediate state exactly as
/// it does for a keypress. `step_ms` is derived from the movie's own pacing
/// (`--movie-confirm-ms`), which is what sets the tempo of a recording.
fn arm_movie_driver(state: &Rc<GuiState>, movie: super::movie::GuiMovie) {
    let options = super::movie::GuiMovieOptions::from_env();
    let state_for_tick = state.clone();
    let movie = std::cell::RefCell::new(movie);
    let index = std::cell::Cell::new(0usize);
    let finished = std::cell::Cell::new(false);

    // The movie animates one *character* per tick while a cell is being typed,
    // then holds the finished cell for `confirm_ms` — the same shape as the TUI
    // replayer. Advancing a whole step per tick instead (the earlier
    // behaviour) made every value appear instantly and then sit motionless for
    // the whole hold, so a long margin label like `--- Belmont ---` occupied
    // the formula bar for ten seconds with nothing happening.
    let char_ms = options.char_delay_ms().max(1.0);
    let tick_ms = char_ms.round().max(1.0) as u32;
    let hold_ticks = ((options.confirm_delay_ms as f64 / char_ms).round() as u32).max(1);
    // Ticks to show the *empty* sheet before the first line replays. The driver
    // is armed before the window is mapped, so without a lead-in the first
    // steps are applied while the window is still being built and the recording
    // opens on a sheet that already has content — indistinguishable from the
    // replay skipping its first lines. One second of blank sheet is enough for a
    // viewer to read "this is an empty spreadsheet" before the typing starts.
    let lead_in_ticks = (1000.0 / char_ms).round() as u32;

    // What the current step is doing: `Typing` reveals `value` one character at
    // a time; `Holding` shows the completed step. The edit itself is applied at
    // the END of the typing animation, so the grid changes exactly when the
    // last character lands (matching what a user typing the value would see).
    #[derive(Clone)]
    enum Phase {
        LeadIn(u32),
        Idle,
        Moving { step: usize, value: String, hold: u32 },
        Typing { step: usize, typed: usize, value: String, hold: u32 },
        Holding { step: usize, hold: u32 },
        /// Ease the drawn pointer toward a target, then run `then`.
        ///
        /// The pointer is painted by the app (X11 does not composite the cursor
        /// into an `x11grab` capture, so the real one is invisible in a
        /// recording). Motion is eased so it reads as a hand moving rather than
        /// a jump, and `total` frames gives the travel a duration.
        PointerMove {
            from: (f64, f64),
            to: (f64, f64),
            frame: u32,
            total: u32,
            /// Menu stop to run once the press completes, if this is a tour move.
            press_after: Option<(String, String)>,
        },
        /// Hold the pressed state briefly so a click is legible.
        PointerPress { until: u32, then_menu: Option<(String, String)> },
    }
    let phase = std::cell::RefCell::new(Phase::LeadIn(lead_in_ticks.max(1)));

    // An optional menu tour, run after the replay finishes: "the user opens
    // each menu and picks an item". Kept out of the `.corro` log because a tour
    // is a presentation concern — expressing it as an op would mean a new kind
    // in the log format and in every backend's parser.
    let tour = std::cell::RefCell::new(crate::ui_core::menu_tour_from_env());
    let tour_idx = std::cell::Cell::new(0usize);
    let tour_start = std::cell::Cell::new(None::<std::time::Instant>);
    // Ticks to keep the window alive after the last stop starts, so its
    // pointer/click phases can finish before the replay quits.
    let tour_settle_frames = std::cell::Cell::new(0u32);

    let tick = move || -> bool {
        if finished.get() {
            return false;
        }
        let mut current = phase.borrow_mut();
        match current.clone() {
            Phase::LeadIn(remaining) => {
                // Show the untouched sheet. `App::reset_workbook_for_movie`
                // cleared the workbook, so nothing has to be undone to paint
                // the blank state — no step has run yet.
                sync_chrome_labels(&state_for_tick);
                state_for_tick.canvas.queue_redraw();
                state_for_tick.window.queue_redraw();
                if remaining > 1 {
                    *current = Phase::LeadIn(remaining - 1);
                } else {
                    *current = Phase::Idle;
                }
                true
            }
            Phase::PointerMove { from, to, frame, total, press_after } => {
                let f = (frame + 1).min(total);
                let t = f as f64 / total.max(1) as f64;
                // Ease-in-out: a hand accelerates away and settles, rather than
                // travelling at a constant rate. Cubic on both halves is enough
                // to read as deliberate movement at 16fps.
                let e = if t < 0.5 {
                    4.0 * t * t * t
                } else {
                    1.0 - (-2.0 * t + 2.0).powi(3) / 2.0
                };
                let x = from.0 + (to.0 - from.0) * e;
                let y = from.1 + (to.1 - from.1) * e;
                state_for_tick.movie_pointer.set(Some((x, y, false)));
                state_for_tick.canvas.queue_redraw();
                state_for_tick.window.queue_redraw();
                if f >= total {
                    state_for_tick.movie_pointer.set(Some((to.0, to.1, false)));
                    match press_after {
                        Some((section, item)) => {
                            state_for_tick.movie_pointer.set(Some((to.0, to.1, true)));
                            *current = Phase::PointerPress {
                                until: 2,
                                then_menu: Some((section, item)),
                            };
                        }
                        None => *current = Phase::Idle,
                    }
                } else {
                    *current = Phase::PointerMove { from, to, frame: f, total, press_after };
                }
                true
            }
            Phase::PointerPress { until, then_menu } => {
                if until > 1 {
                    *current = Phase::PointerPress { until: until - 1, then_menu };
                    return true;
                }
                // Release: the pointer returns to its normal colour and the
                // queued action runs.
                if let Some((x, y, _)) = state_for_tick.movie_pointer.get() {
                    state_for_tick.movie_pointer.set(Some((x, y, false)));
                }
                state_for_tick.canvas.queue_redraw();
                state_for_tick.window.queue_redraw();
                if let Some((section, item)) = then_menu {
                    run_menu_tour_stop(&state_for_tick, &section, &item);
                }
                *current = Phase::Idle;
                true
            }
            Phase::Idle => {
                let i = index.get();
                let mut movie = movie.borrow_mut();
                if i >= movie.len() {
                    // Run any scripted menu tour before the window closes.
                    //
                    // The tour owns the driver until it is finished: starting a
                    // stop must NOT fall through to the quit below, or the
                    // window closes while the pointer is still travelling and
                    // the recording ends before the menu is even opened.
                    let steps = tour.borrow();
                    let ti = tour_idx.get();
                    if ti < steps.len() {
                        if tour_start.get().is_none() {
                            tour_start.set(Some(std::time::Instant::now()));
                        }
                        let elapsed = tour_start.get().unwrap().elapsed().as_millis() as u64;
                        // Keep the window open for as long as any stop is still
                        // pending, and refresh the settle budget each tick: the
                        // final stop's pointer/click phases need it, and a later
                        // stop can be scheduled seconds away.
                        tour_settle_frames.set(8);
                        if steps[ti].at_ms <= elapsed {
                            let step = steps[ti].clone();
                            // Aim at the section's menu button (the bar runs
                            // along the top of the window) and press there.
                            let target = menu_section_pointer_target(&step.section, step.item_index);
                            let from = state_for_tick
                                .movie_pointer
                                .get()
                                .map(|(x, y, _)| (x, y))
                                .unwrap_or((target.0 - 120.0, target.1 + 180.0));
                            *current = Phase::PointerMove {
                                from,
                                to: target,
                                frame: 0,
                                total: 12,
                                press_after: Some((step.section.clone(), step.item.clone())),
                            };
                            tour_idx.set(ti + 1);
                        }
                        drop(steps);
                        return true;
                    }
                    // Every stop has been started, but the last one may still be
                    // animating: wait for its phases to finish before quitting,
                    // or the window closes mid-click.
                    let left = tour_settle_frames.get();
                    if left > 0 {
                        tour_settle_frames.set(left - 1);
                        return true;
                    }
                    drop(steps);
                    finished.set(true);
                    state_for_tick.app_mut().core.status =
                        format!("Movie complete: {} lines", movie.applied);
                    sync_chrome_labels(&state_for_tick);
                    state_for_tick.canvas.queue_redraw();
                    state_for_tick.window.queue_redraw();
                    // A movie is a script, not an interactive session: when it
                    // ends the window closes (the TUI replayer quits the same
                    // way), so a recording finishes on its own instead of
                    // leaving the app open forever. `save_before_quit` is the
                    // same path the File ▸ Quit menu takes.
                    save_before_quit(&state_for_tick);
                    return false;
                }
                // What will this step write, if it is a plain cell value?
                let typed = movie.step_typed_text(i);
                // Show the step's menu flash first, if it has one.
                let menu = movie.step_menu(i);
                if let Some((section, item)) = menu {
                    let app = state_for_tick.app_mut();
                    app.core.status = format!(
                        "Movie {}/{}  {} ▸ {}",
                        i + 1, movie.len(), section, item
                    );
                    sync_chrome_labels(&state_for_tick);
                    state_for_tick.canvas.queue_redraw();
                    state_for_tick.window.queue_redraw();
                    *current = Phase::Holding { step: i, hold: hold_ticks };
                    return true;
                }
                match typed {
                    Some(value) => {
                        // Park the cursor on the cell this step writes, before a
                        // single character appears. The grid paints its edit
                        // preview on the cursor cell, so a cursor left over from
                        // the previous step would show the typing happening in
                        // the wrong cell — the value would then "jump" to its
                        // real home only at the commit.
                        //
                        // `step_cursor` resolves the step's address with the
                        // *live* workbook, exactly as the step itself does. That
                        // matters twice over: the margin cell `[A1` is a
                        // different cell from main `A1`, and the `_N`/`~1` forms
                        // (`[A_2`, `]A~1`) are relative to the sheet's current
                        // extent, which grows as the replay runs. Resolving
                        // against a fresh 1x1 workbook put the cursor up to
                        // eight rows or a column away from the cell it typed
                        // into.
                        let cursor = {
                            let app = state_for_tick.app_ref();
                            movie.step_cursor(i, &app.core.workbook)
                        };
                        if let Some(cursor) = cursor {
                            let app = state_for_tick.app_mut();
                            app.core.cursor = cursor;
                            let grid = &app.core.workbook.active_sheet().grid;
                            app.core.cursor.clamp(grid);
                            let row = app.core.cursor.row;
                            let col = app.core.cursor.col;
                            update_state_cursor(&state_for_tick, row, col);
                        }
                        // The cursor move is its own frame: it must be on screen
                        // before the first character lands, otherwise the first
                        // keystroke appears on the old cell.
                        *current = Phase::Moving { step: i, value, hold: hold_ticks };
                        true
                    }
                    None => {
                        apply_movie_step(&state_for_tick, &mut movie, i);
                        *current = Phase::Holding { step: i, hold: hold_ticks };
                        true
                    }
                }
            }
            Phase::Moving { step, value, hold } => {
                // The cursor landed on the previous tick; start revealing the
                // characters from here.
                *current = Phase::Typing { step, typed: 0, value, hold };
                true
            }
            Phase::Typing { step, typed, value, hold } => {
                let chars: Vec<char> = value.chars().collect();
                let next = (typed + 1).min(chars.len());
                let shown: String = chars[..next].iter().collect();
                let done = next >= chars.len();
                {
                    // Put the partial value in the formula entry so the bar
                    // shows the text growing, exactly as the TUI does by
                    // re-entering edit mode with the partial buffer each
                    // character. Without this the animation was invisible: the
                    // status line counted up while the bar stayed empty, and the
                    // value only appeared at the commit.
                    state_for_tick.editing.set(true);
                    *state_for_tick.edit_buf.borrow_mut() = shown.clone();
                    sync_entry_to_buf(&state_for_tick);
                    let app = state_for_tick.app_mut();
                    app.core.status = format!(
                        "Movie {}/{}  typing: {shown}",
                        step + 1, movie.borrow().len()
                    );
                    sync_chrome_labels(&state_for_tick);
                    state_for_tick.canvas.queue_redraw();
                    state_for_tick.window.queue_redraw();
                }
                if done {
                    // Commit exactly when the last character lands, so the grid
                    // changes with the typing rather than before it.
                    let mut movie = movie.borrow_mut();
                    apply_movie_step(&state_for_tick, &mut movie, step);
                    index.set(step + 1);
                    *current = Phase::Holding { step, hold };
                } else {
                    *current = Phase::Typing { step, typed: next, value, hold };
                }
                true
            }
            Phase::Holding { step, hold } => {
                if hold > 1 {
                    *current = Phase::Holding { step, hold: hold - 1 };
                    return true;
                }
                let mut movie = movie.borrow_mut();
                if index.get() <= step {
                    index.set(step + 1);
                }
                let _ = &mut movie;
                *current = Phase::Idle;
                true
            }
        }
    };

    match rswidgets::add_periodic_tick(&state.window, tick_ms, Box::new(tick)) {
        Ok(()) => eprintln!(
            "[corro] movie driver armed ({tick_ms}ms/char, {hold_ticks} ticks per step hold)"
        ),
        Err(e) => eprintln!("[corro] movie timer unavailable: {e}"),
    }
}

/// Apply one movie step to the app and refresh the chrome.
fn apply_movie_step(
    state: &Rc<GuiState>,
    movie: &mut super::movie::GuiMovie,
    i: usize,
) {
    let app = state.app_mut();
    match movie.apply_step(&mut app.core.workbook, &mut app.core.view_sheet_id, i) {
        Ok(frame) => {
            app.core.state = app.core.workbook.active_sheet().clone();
            app.core.ops_applied = movie.applied;
            if let Some(addr) = frame.cursor.as_ref() {
                app.core.cursor = super::movie::cursor_of(addr, &app.core.workbook);
                let grid = &app.core.workbook.active_sheet().grid;
                app.core.cursor.clamp(grid);
            }
            // The typing animation left a partial buffer in the entry; drop it
            // before the chrome refresh so the bar shows the freshly committed
            // *cell* value rather than the last typed character run.
            state.editing.set(false);
            state.edit_buf.borrow_mut().clear();
            // The same chrome refresh a keypress performs.
            update_state_cursor(state, app.core.cursor.row, app.core.cursor.col);
            let caption = match frame.menu.as_ref() {
                Some((section, item)) => format!(
                    "Movie {}/{}  {} ▸ {}  ·  {}",
                    frame.progress.0, frame.progress.1, section, item, frame.status
                ),
                None => format!(
                    "Movie {}/{}  {}",
                    frame.progress.0, frame.progress.1, frame.status
                ),
            };
            app.core.status = caption;
            sync_chrome_labels(state);
            state.canvas.queue_redraw();
            state.window.queue_redraw();
        }
        Err(e) => {
            state.app_mut().core.status = format!("Movie error: {e}");
            sync_chrome_labels(state);
        }
    }
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

    /// Covered row headers use the body selection fill; uncovered rows keep
    /// the neutral header gray (and `None` keeps everything gray, i.e. the
    /// pre-existing unselected rendering).
    #[test]
    fn covered_row_headers_use_selection_fill() {
        use std::collections::BTreeSet;
        let mut dc = RecordingDrawContext::new();
        let mut hits = Vec::new();
        paint_row_headers(
            &mut dc,
            &[HEADER_ROWS, HEADER_ROWS + 1],
            1,
            &BTreeSet::new(),
            &mut hits,
            Some((HEADER_ROWS, HEADER_ROWS)),
        );
        let fills: Vec<(f64, f64, f64, f64)> = dc
            .ops
            .iter()
            .filter_map(|op| match op {
                DrawOp::FillRect { x, w, rgba, .. }
                    if *x == 0.0 && *w == row_label_w() =>
                {
                    Some(*rgba)
                }
                _ => None,
            })
            .collect();
        assert_eq!(
            fills,
            vec![(0.9, 0.95, 1.0, 1.0), (0.9, 0.9, 0.9, 1.0)],
            "covered row must use the body selection fill, uncovered row header gray, got {fills:?}"
        );
    }

    /// Row gutter labels must paint bold (weight 1), with the row's label.
    /// Short labels (<=2 chars) additionally record a padlock hit rect.
    #[test]
    fn row_headers_paint_bold() {
        use std::collections::BTreeSet;
        let mut dc = RecordingDrawContext::new();
        // Logical rows: main row 0 then footer rows (mr=1).
        let mut hits = Vec::new();
        paint_row_headers(&mut dc, &[HEADER_ROWS, HEADER_ROWS + 1], 1, &BTreeSet::new(), &mut hits, None);
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
        assert_eq!((hits[0].x, hits[0].w), (2.0, padlock_w()));
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
        assert_eq!(xs, vec![row_label_w() - 8.0 - 6.0], "row number x, got {xs:?}");
    }

    /// Column gutter labels must paint bold (weight 1), with the column name.
    #[test]
    fn covered_col_headers_use_selection_fill() {
        use std::collections::BTreeSet;
        let mut dc = RecordingDrawContext::new();
        // Margin col, main col A (mc=1); cover only the main column.
        let col_ixs = vec![MARGIN_COLS - 1, MARGIN_COLS];
        let col_widths: HashMap<usize, usize> =
            col_ixs.iter().map(|&c| (c, 8)).collect();
        let mut hits = Vec::new();
        paint_col_headers(
            &mut dc,
            &col_ixs,
            &col_widths,
            1,
            &BTreeSet::new(),
            &mut hits,
            Some((MARGIN_COLS, MARGIN_COLS)),
        );
        let fills: Vec<(f64, f64, f64, f64)> = dc
            .ops
            .iter()
            .filter_map(|op| match op {
                DrawOp::FillRect { y, h, rgba, .. }
                    if *y == 0.0 && *h == header_h() =>
                {
                    Some(*rgba)
                }
                _ => None,
            })
            .collect();
        assert_eq!(
            fills,
            vec![(0.9, 0.9, 0.9, 1.0), (0.9, 0.95, 1.0, 1.0)],
            "margin col keeps header gray, covered col uses selection fill, got {fills:?}"
        );
    }
    #[test]
    fn col_headers_paint_bold() {
        use std::collections::BTreeSet;
        let mut dc = RecordingDrawContext::new();
        // Margin col, main col A (mc=1).
        let col_ixs = vec![MARGIN_COLS - 1, MARGIN_COLS];
        let col_widths: HashMap<usize, usize> =
            col_ixs.iter().map(|&c| (c, 8)).collect();
        let mut hits = Vec::new();
        paint_col_headers(&mut dc, &col_ixs, &col_widths, 1, &BTreeSet::new(), &mut hits, None);
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

    /// Tabs lay out left to right from x=2, padded by tab_pad_x() on both
    /// sides with tab_gap() between. "Sheet1" at the stub 8px/char measures
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

fn sheet_rec_col_width(sheet: &crate::ops::SheetState, col: usize) -> usize {
    sheet.grid.col_width(col).max(1)
}

/// On-screen width of a column in characters: the recorded width plus one
/// spare character when the column header carries a padlock, so the icon fits
/// inside its own column instead of overlapping the neighbor's gridline.
/// Single source of truth for every width accumulation (render headers, render
/// cells, click mapping, viewport sizing) — they must all agree or columns
/// misalign.
fn display_col_width(sheet: &crate::ops::SheetState, c: usize, mc: usize) -> usize {
    sheet_rec_col_width(sheet, c)
        + usize::from(wants_padlock(&crate::addr::ui_column_fragment(c, mc)))
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
            .filter(|r| r.y >= header_h() - 0.5 && r.y < header_h() + row_h())
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
                * char_w() as usize;
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
        paint_col_headers(&mut dc, &cols, &widths, mc, &BTreeSet::new(), &mut hits, None);
        assert!(!hits.is_empty(), "expected padlock columns in the A1 view");
        // Replay the paint's x-accumulation; each padlock's right edge must
        // not pass its own column's right edge.
        let mut cx = row_label_w();
        let mut hi = 0usize;
        for &c in &cols {
            let cw = *widths.get(&c).unwrap() as f64 * char_w();
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
            * char_w() as usize;
        assert!(
            used_px >= 1220,
            "viewport must cover 1220px on A1 (dim={dim}, cols={}, used~{used_px})",
            cols.len()
        );
    }


    /// `cursor_display_row_index` must return a position in the display list,
    /// not the row's address.
    ///
    /// Regression (Android): the first version returned `cursor.row` for body
    /// rows. Body addresses are `HEADER_ROWS + index` and `HEADER_ROWS` is
    /// 999_999_999, so the "position" came back as ~1e9. Used as a viewport
    /// anchor that scrolled the sheet millions of rows past the cursor, which
    /// then rendered nowhere — a tap appeared to select nothing at all.
    #[test]
    fn cursor_display_row_index_is_a_list_position_not_a_row_address() {
        let app = overflow_app();
        let sheet = app.core.workbook.active_sheet();
        let hr = HEADER_ROWS;
        // The first body row is a legitimate, common cursor position.
        let cursor = SheetCursor { row: hr, col: MARGIN_COLS };
        let idx = ui_core::cursor_display_row_index(sheet, cursor);
        assert!(
            idx < 10_000,
            "index {idx} looks like a row address, not a list position \
             (HEADER_ROWS is {hr})"
        );
    }
}

#[cfg(test)]
mod text_edit_tests {
    use super::text_edit::{self, KeyOutcome};

    /// A minimal model of the formula edit session: the buffer and the caret
    /// the host owns — the same `text_edit` surface `handle_edit_key` drives,
    /// so these test production logic rather than a paraphrase.
    struct Session {
        buf: String,
        caret: usize,
    }
    impl Session {
        fn type_char(&mut self, ch: char) {
            text_edit::insert_char(&mut self.buf, &mut self.caret, ch);
        }
        fn backspace(&mut self) {
            text_edit::backspace(&mut self.buf, &mut self.caret);
        }
        fn delete(&mut self) -> bool {
            if self.caret < text_edit::caret_len(&self.buf) {
                text_edit::delete_forward(&mut self.buf, &mut self.caret);
                true
            } else {
                false
            }
        }
        fn left(&mut self) -> KeyOutcome {
            text_edit::left(&self.buf, &mut self.caret)
        }
        fn right(&mut self) -> KeyOutcome {
            text_edit::right(&self.buf, &mut self.caret)
        }
    }
    fn session(seed: &str, caret: usize) -> Session {
        Session { buf: seed.to_string(), caret }
    }

    /// Regression (gcorro.exe): typing must insert at the caret, not always
    /// append. "=1+2" with the caret on '+' plus '9' is "=19+2", not "=1+29".
    #[test]
    fn typing_inserts_at_caret_not_at_end() {
        let mut s = session("=1+2", 2);
        assert_eq!(s.buf.chars().nth(s.caret), Some('+'));
        s.type_char('9');
        assert_eq!(s.buf, "=19+2", "must insert at the caret, not append");
        assert_ne!(s.buf, "=1+29", "regression: appended to the end");
        assert_eq!(s.caret, 3);
    }

    #[test]
    fn type_left_type_inserts_in_the_middle() {
        let mut s = session("", 0);
        for ch in "=1+2".chars() {
            s.type_char(ch);
        }
        assert_eq!(s.left(), KeyOutcome::Edited);
        assert_eq!(s.left(), KeyOutcome::Edited);
        assert_eq!(s.caret, 2);
        s.type_char('9');
        assert_eq!(s.buf, "=19+2");
    }

    #[test]
    fn arrows_move_caret_and_only_commit_at_edges() {
        let mut s = session("ABCD", 2);
        assert_eq!(s.left(), KeyOutcome::Edited);
        assert_eq!(s.left(), KeyOutcome::Edited);
        assert_eq!(s.caret, 0);
        assert_eq!(s.left(), KeyOutcome::CommitAndMoveCell);
        assert_eq!(s.caret, 0, "edge Left must not underflow");

        let mut s = session("ABCD", 2);
        assert_eq!(s.right(), KeyOutcome::Edited);
        assert_eq!(s.right(), KeyOutcome::Edited);
        assert_eq!(s.caret, 4);
        assert_eq!(s.right(), KeyOutcome::CommitAndMoveCell);
        assert_eq!(s.caret, 4, "edge Right must not run past the buffer");
    }

    #[test]
    fn backspace_and_delete_act_at_the_caret() {
        let mut s = session("ABC", 2);
        s.backspace();
        assert_eq!(s.buf, "AC");
        assert_eq!(s.caret, 1);

        let mut s = session("AC", 0);
        s.backspace();
        assert_eq!(s.buf, "AC", "no-op at position 0");

        let mut s = session("ABC", 1);
        assert!(s.delete());
        assert_eq!(s.buf, "AC");
        assert_eq!(s.caret, 1);

        let mut s = session("AC", 2);
        assert!(!s.delete(), "no-op at end");
    }

    /// The caret is a char index, so multibyte input can never split a
    /// codepoint.
    #[test]
    fn multibyte_insert_and_delete_are_char_safe() {
        let mut s = session("aé", 1);
        s.type_char('ß');
        assert_eq!(s.buf, "aßé");
        s.backspace();
        assert_eq!(s.buf, "aé");
        s.delete();
        assert_eq!(s.buf, "a");
    }
}

#[cfg(test)]
mod agg_drop_tests {
    use super::*;

    /// The dropdown's rows are the shared vocabulary, in navigation order —
    /// the same list every other backend offers.
    #[test]
    fn rows_match_the_shared_vocabulary() {
        let rows = agg_drop_rows();
        assert_eq!(rows, crate::ui_core::agg_labelled_choices());
        assert_eq!(rows[0], "1: TOTAL");
        assert_eq!(rows[5], "6: MEDIAN");
    }

    /// The list hangs under the anchored cell, is at least readable width,
    /// and flips above the cell when it would overflow the canvas bottom.
    #[test]
    fn layout_hangs_under_the_cell_and_flips_on_overflow() {
        let (x, y, w, h, rh) = agg_drop_layout_for_cell(158, 24, 50, 1200.0, 800.0);
        assert_eq!(rh, row_h());
        assert_eq!(y, 24.0 + row_h(), "opens downward from the cell");
        assert!(w >= 150.0, "at least wide enough to read a label");
        let rows = crate::ui_core::AGG_CHOICES.len() as f64;
        assert_eq!(h, rows * row_h() + 2.0, "one row per choice plus the border");
        assert_eq!(x, 158.0);

        // Near the right edge the box is pulled back inside the canvas.
        let (x_edge, _, w_edge, _, _) =
            agg_drop_layout_for_cell(1180, 24, 50, 1200.0, 800.0);
        assert!(
            x_edge + w_edge <= 1200.0,
            "must not spill past the right edge (x={x_edge}, w={w_edge})"
        );

        // A canvas taller than the cell but shorter than the list: the box
        // flips above the cell and stays pinned to the top edge (it cannot
        // fit, but it must not start off-screen).
        let (_, y_flip, _, _, _) = agg_drop_layout_for_cell(158, 400, 50, 1200.0, 500.0);
        assert!(
            y_flip < 400.0,
            "must flip above the cell when it would overflow (y={y_flip})"
        );
        // A very short canvas pins the box to the top instead of going
        // negative (the remaining rows are simply clipped by the canvas).
        let (_, y_small, _, _, _) = agg_drop_layout_for_cell(158, 24, 50, 1200.0, 60.0);
        assert!(y_small >= 0.0, "must never start above the canvas");
    }

    /// Every row's band maps to that row; points outside the box hit nothing.
    #[test]
    fn row_bands_are_contiguous_and_bounded() {
        let (bx, by, _bw, _bh, row_h) = agg_drop_layout_for_cell(158, 24, 50, 1200.0, 800.0);
        let rows = agg_drop_rows().len();
        for i in 0..rows {
            let top = by + 1.0 + i as f64 * row_h;
            let mid = top + row_h / 2.0;
            let idx = ((mid - by) / row_h) as usize;
            assert_eq!(idx, i, "row {i} band must map to itself");
        }
        // The band below the last row is out of range.
        let below = by + 1.0 + rows as f64 * row_h + row_h / 2.0;
        assert!(((below - by) / row_h) as usize >= rows);
    }

    /// The highlight clamps to the choice range at both ends.
    #[test]
    fn step_clamps_within_the_choice_range() {
        assert_eq!(crate::ui_core::agg_step_index(0, -1), 0);
        assert_eq!(
            crate::ui_core::agg_step_index(0, 99),
            crate::ui_core::AGG_CHOICES.len() - 1
        );
    }

    /// REPRO/fix: the address → display `(row, col)` mapping must resolve
    /// **every** key class, especially the `[A_n` footer key column
    /// (`[A_1`, `[A_2`, `[A_3`, …). A `None` here makes `cell_rect` return
    /// `None`, which makes `show_agg_dropdown` bail — so clicking such a key
    /// silently does nothing (no dropdown). This is the bug where only
    /// `[A_1`-style keys popped the list.
    #[test]
    fn footer_keys_have_a_display_rect_like_every_other_key() {
        use crate::grid::{CellAddr, ColumnAddr};
        let mc = 3;
        let mr = 4;

        // The `[A_n` key column: left margin, last left column, **every**
        // footer row. All must map to a footer display row past the body.
        for footer_row in 0..4u32 {
            let addr = CellAddr::Footer { row: footer_row, col: ColumnAddr::Left(MARGIN_COLS - 1) };
            let (r, c) = agg_key_display_rc(&addr, mc, mr)
                .unwrap_or_else(|| panic!("[A_{} must have a display rect", footer_row + 1));
            assert_eq!(r, HEADER_ROWS + mr + footer_row as usize, "footer row past the body");
            assert_eq!(c, MARGIN_COLS - 1, "left-margin key column");
        }

        // Regression guard: the classes that already worked keep working.
        assert!(agg_key_display_rc(
            &CellAddr::Header { row: 0, col: ColumnAddr::Right(0) }, mc, mr
        )
        .is_some());
        assert!(agg_key_display_rc(
            &CellAddr::Right { row: 0, col: 0 }, mc, mr
        )
        .is_some());
        assert!(agg_key_display_rc(
            &CellAddr::Main { row: 0, col: 0 }, mc, mr
        )
        .is_some());
        assert!(agg_key_display_rc(
            &CellAddr::Left { row: 0, col: 0 }, mc, mr
        )
        .is_some());

        // The footer key and its inverse (click mapping) must agree, or the
        // list anchors on a different row than the one the user clicked.
        let addr = CellAddr::Footer { row: 2, col: ColumnAddr::Left(MARGIN_COLS - 1) };
        let (r, c) = agg_key_display_rc(&addr, mc, mr).unwrap();
        let back = crate::addr::sheet_cursor_to_addr(
            crate::addr::LogicalRow(r),
            crate::addr::GlobalCol(c),
            crate::addr::MainRows(mr),
            crate::addr::MainCols(mc),
        );
        assert_eq!(back, addr, "display rect must round-trip to the same key");
    }
}

#[cfg(test)]
mod startup_race_tests {
    use super::*;

    /// REPRO: clicking a footer key `[A_n` during present()'s event pump
    /// selected the wrong cell and turned it yellow (edit state) instead of
    /// leaving the clicked cell blue with its dropdown open. The click sets
    /// `clicked`, and setup must then refuse to reselect A1.
    #[test]
    fn a_click_during_startup_pump_wins_over_selecting_a1() {
        assert_eq!(
            startup_edit_action(false, true),
            StartupEdit::KeepClicked,
            "a click during present() must not be clobbered by the A1 selection"
        );
    }

    /// With no interaction the default still selects A1 (unchanged behaviour).
    #[test]
    fn no_interaction_still_selects_a1() {
        assert_eq!(startup_edit_action(false, false), StartupEdit::SelectA1);
    }

    /// An in-flight edit with typed content outranks even a click: the
    /// replayer may have typed before present() returned.
    #[test]
    fn in_flight_edit_outranks_a_click() {
        assert_eq!(
            startup_edit_action(true, true),
            StartupEdit::KeepInFlight
        );
        assert_eq!(
            startup_edit_action(true, false),
            StartupEdit::KeepInFlight
        );
    }
}
