use crate::grid::{CellAddr, ColumnAddr, SheetCursor, HEADER_ROWS, MARGIN_COLS};
use crate::ui_core;
use std::collections::HashMap;
use rswidgets::backends_pancurses_adapter::*;


use unicode_width::UnicodeWidthStr;

use super::actions::{apply_format, commit_cell, dispatch_menu_action, main_addr_label, menu_action_needs_prompt, run_prompt_action, sort_sheet, MenuDispatch};
use super::viewport::Viewport;
use super::compute;
use super::render::{self, CellSink};

/// Pancurses adapter: wraps a `Spreadsheet` ref as a `CellSink` for the
/// backend-agnostic `render::fill_cells`.
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



/// Show a modal info dialog (About / Help) in the pancurses UI.  The backend
/// draws it via SGR on top of the spreadsheet output (ncurses widgets get
/// overwritten by the spreadsheet's direct SGR writes).  Without this, Help ->
/// About just set a status string and no dialog ever appeared.
fn show_info_dialog(title: &str, text: &str) {
    rswidgets::backends::pancurses::show_dialog(title, text);
}

/// Borrow the host [`App`](super::App) behind a raw UI-thread pointer.
/// Centralised here so the raw dereference happens in exactly one place and
/// call sites stay `unsafe`-free.
///
/// # Contract (not machine-checked)
/// The `'a` lifetime is intentionally free: the pointer is trusted, so the
/// compiler cannot prevent two live borrows — only discipline can. Callers
/// must observe two rules (see `GuiState::app_mut` for the full rationale):
/// 1. **LIFO nesting only** — never use an outer borrow after an inner one
///    was taken.
/// 2. **No cross-frame borrows** — each closure below binds its borrow
///    locally and drops it before returning, so every borrow's dynamic
///    extent lies within a single sequential dispatch.
/// The pointer itself cannot dangle: the caller holds the `App` across the
/// blocking event loop.
#[allow(clippy::needless_lifetimes)]
fn app_from_raw<'a>(app_ptr: *mut super::App) -> &'a mut super::App {
    unsafe { &mut *app_ptr }
}

/// Re-run the full viewport computation and re-fill the spreadsheet widget's
/// cells from the workbook.  Menu actions and prompt submissions mutate the
/// workbook (Insert Date, Cut/Paste, New sheet, Open file, ...) but, unlike a
/// cursor move, nothing re-filled the widget's own cell buffers — so the grid
/// kept showing the stale values until the next key press (e.g. Insert Date
/// only appeared after arrowing away).  Every mutation entry point calls this
/// so the grid reflects the change immediately.
fn refresh_viewport_after_action(
    app: &mut super::App,
    ss: &Spreadsheet,
    sid: usize,
    display_rows: &std::rc::Rc<std::cell::RefCell<Vec<usize>>>,
    data_rows: usize,
    data_cols: usize,
    data_width: usize,
    hr: usize,
) {
    let cursor = app.core.cursor;
    let vp = Viewport::recompute(app, cursor, data_rows, data_cols, data_width, hr, MARGIN_COLS);
    let rec = app.core.workbook.active_sheet().clone();
    spreadsheet_set_border_title(sid, &vp.border_title(app.core.ops_applied));
    spreadsheet_set_row_labels(sid, vp.row_labels.clone());
    spreadsheet_set_column_layout(sid, vp.column_layout.clone());
    *display_rows.borrow_mut() = vp.display_rows.clone();
    spreadsheet_set_grid_config(sid, MARGIN_COLS as u32, vp.mc as u32);
    // Sync the tab bar (New/Copy/Rename/Move sheet change the workbook's sheet
    // list; without this the tab bar stays stale after a menu action).
    if app.core.workbook.sheet_count() > 1 {
        let titles: Vec<String> = app.core.workbook.sheets.iter()
            .map(|s| s.title.clone())
            .collect();
        let active = app.core.workbook.active_sheet;
        ss.set_tab_data(&titles, active);
    } else {
        ss.set_tab_data(&[], 0);
    }
    vp.refill(&mut SpreadsheetSink::new(ss), &rec.grid, hr, MARGIN_COLS, data_width, cursor.row, cursor.col);
    // The cursor cell is rendered from the raw (unformatted) value, matching
    // the cursor-move callback's post-refill raw-cell update.  Also sync the
    // WIDGET cursor to the (possibly moved) app cursor so the formula bar
    // address follows — mitosis moves the cursor onto the duplicate row/col.
    if let Some(new_display_ri) = vp.display_rows.iter().position(|&r| r == cursor.row) {
        ss.set_cursor(new_display_ri as u32, cursor.col as u32);
        let cursor_addr = crate::addr::sheet_cursor_to_addr(
            crate::addr::LogicalRow(cursor.row),
            crate::addr::GlobalCol(cursor.col),
            crate::addr::MainRows(vp.mr),
            crate::addr::MainCols(vp.mc),
        );
        let raw = rec.grid.get(&cursor_addr).unwrap_or_default();
        ss.set_raw_cell(new_display_ri as u32, cursor.col as u32, &raw);
    }
}

/// The About-dialog body text, sourced from the ratatui reference
/// (`crate::ui::App::about_page_body`) so the pancurses dialog renders the SAME
/// content as the ratatui backend (render parity).
fn about_body() -> String {
    crate::ui_core::about_page_body()
}

/// The Full-help dialog body text, sourced from the ratatui reference
/// (`crate::ui::App::help_page_body`) so the pancurses dialog renders the SAME
/// content as the ratatui backend (render parity).
fn help_body() -> String {
    crate::ui_core::help_page_body()
}

/// Append `bytes` to `path` (create if needed) via raw CreateFileA — std::fs
/// is broken on Win9x (CreateFileW is an unimplemented stub, OS error 120).
#[cfg(windows)]
fn append_marker_raw(path: &str, bytes: &[u8]) {
    unsafe {
        use std::os::raw::c_void;
        unsafe extern "system" {
            fn CreateFileA(
                name: *const u8,
                access: u32,
                share: u32,
                sa: *mut c_void,
                disp: u32,
                flags: u32,
                tmpl: *mut c_void,
            ) -> *mut c_void;
            fn WriteFile(h: *mut c_void, buf: *const u8, len: u32, written: *mut u32, ov: *mut c_void) -> i32;
            fn CloseHandle(h: *mut c_void) -> i32;
        }
        let mut pathz = path.as_bytes().to_vec();
        pathz.push(0);
        let h = CreateFileA(
            pathz.as_ptr(),
            0x4000_0000, // GENERIC_WRITE
            1,           // FILE_SHARE_READ
            std::ptr::null_mut(),
            4,           // OPEN_ALWAYS
            0x80,        // FILE_ATTRIBUTE_NORMAL
            std::ptr::null_mut(),
        );
        if h.is_null() || h as isize == -1 {
            return;
        }
        let mut written = 0u32;
        WriteFile(h, bytes.as_ptr(), bytes.len() as u32, &mut written, std::ptr::null_mut());
        CloseHandle(h);
    }
}

/// Non-Windows append: plain std::fs (the CreateFileW-stub issue is Win9x-only).
#[cfg(not(windows))]
fn append_marker_raw(path: &str, bytes: &[u8]) {
    use std::io::Write as _;
    if let Ok(mut f) = std::fs::OpenOptions::new().create(true).append(true).open(path) {
        let _ = f.write_all(bytes);
    }
}

pub fn run_pancurses(app: &mut super::App) -> Result<(), Box<dyn std::error::Error>> {    // Win9x: environment variables do not propagate to Win32 processes (neither
    // DOS-box `set` nor AUTOEXEC.BAT), so fall back to fixed diagnostic paths
    // when built for the rust9x-msvc (Win95) target.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    let trace_fallback = {
        rswidgets::backends::pancurses::set_input_trace_file("c:\\corro.keys");
        true
    };
    #[cfg(not(all(target_family = "rust9x", target_env = "msvc")))]
    let trace_fallback = false;
    let _ = trace_fallback;

    let _ = std::env::var("INPUT_TRACE_FILE").inspect(|v| {
        eprintln!("[corro] input trace file: {v}");
    });

    // TEMPORARY Win95 diagnosis: run the probe95 input sequence (open CONIN$,
    // SetConsoleMode, poll GetNumberOfConsoleInputEvents + ReadConsoleInputA)
    // BEFORE pancurses init, to A/B against the post-init state.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe fn input_probe_pre() {
        use std::os::raw::c_void;
        unsafe extern "system" {
            fn CreateFileA(
                name: *const u8, access: u32, share: u32, sa: *mut c_void,
                disp: u32, flags: u32, tmpl: *mut c_void) -> *mut c_void;
            fn SetConsoleMode(h: *mut c_void, mode: u32) -> i32;
            fn GetNumberOfConsoleInputEvents(h: *mut c_void, n: *mut u32) -> i32;
            fn ReadConsoleInputA(h: *mut c_void, rec: *mut c_void, len: u32, read: *mut u32) -> i32;
            fn CloseHandle(h: *mut c_void) -> i32;
            fn Sleep(ms: u32);
        }
        let h = CreateFileA(
            b"CONIN$\0".as_ptr(), 0xC000_0000, 3, std::ptr::null_mut(),
            3, 0x80, std::ptr::null_mut());
        let mut out = String::new();
        if h.is_null() || h as isize == -1 {
            out.push_str("open=fail\n");
        } else {
            let m = SetConsoleMode(h, 0x18);
            out.push_str(&format!("open=ok mode0x18={m}\n"));
            for _ in 0..20 {
                let mut n = 0u32;
                GetNumberOfConsoleInputEvents(h, &mut n);
                if n > 0 {
                    let mut buf = [0u32; 5];
                    let mut r = 0u32;
                    let ok = ReadConsoleInputA(h, buf.as_mut_ptr() as *mut c_void, 1, &mut r);
                    let et = buf[0];
                    out.push_str(&format!("ev ok={ok} et={et}\n"));
                } else {
                    out.push_str(&format!("n={n}\n"));
                    Sleep(150);
                }
            }
            CloseHandle(h);
        }
        append_marker_raw("c:\\corro.inq", out.as_bytes());
    }
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { input_probe_pre() };

    /// TEMPORARY Win95 diagnosis: after pancurses init + first frame, poll the
    /// console input queue and log whether events arrive.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe fn input_probe_post() {
        use std::os::raw::c_void;
        unsafe extern "system" {
            fn CreateFileA(
                name: *const u8, access: u32, share: u32, sa: *mut c_void,
                disp: u32, flags: u32, tmpl: *mut c_void) -> *mut c_void;
            fn SetConsoleMode(h: *mut c_void, mode: u32) -> i32;
            fn GetNumberOfConsoleInputEvents(h: *mut c_void, n: *mut u32) -> i32;
            fn ReadConsoleInputA(h: *mut c_void, rec: *mut c_void, len: u32, read: *mut u32) -> i32;
            fn CloseHandle(h: *mut c_void) -> i32;
            fn Sleep(ms: u32);
        }
        let h = CreateFileA(
            b"CONIN$\0".as_ptr(), 0xC000_0000, 3, std::ptr::null_mut(),
            3, 0x80, std::ptr::null_mut());
        let mut out = String::new();
        if h.is_null() || h as isize == -1 {
            out.push_str("post open=fail\n");
        } else {
            let m = SetConsoleMode(h, 0x18);
            out.push_str(&format!("post open=ok mode0x18={m}\n"));
            for _ in 0..10 {
                let mut n = 0u32;
                GetNumberOfConsoleInputEvents(h, &mut n);
                if n > 0 {
                    let mut buf = [0u32; 5];
                    let mut r = 0u32;
                    let ok = ReadConsoleInputA(h, buf.as_mut_ptr() as *mut c_void, 1, &mut r);
                    let et = buf[0];
                    out.push_str(&format!("post ev ok={ok} et={et}\n"));
                } else {
                    out.push_str(&format!("post n={n}\n"));
                    Sleep(150);
                }
            }
            CloseHandle(h);
        }
        append_marker_raw("c:\\corro.inq", out.as_bytes());
    }
    let _backend = rswidgets::backends::pancurses::init()
        .map_err(|e| format!("pancurses init failed: {e}"))?;

    // Test-harness idle marker: the toolkit exposes a generic after-redraw
    // callback; corro wires it to append to the CORRO_IDLE_MARKER file so a
    // test can detect when a frame is fully flushed (instead of sleeping).
    // Win95 DOS boxes cannot type `_` reliably (keyboard-layout mismatch), so
    // the underscore-free alias CORROIDLEMARKER is also accepted.
    let idle_path = match std::env::var("CORRO_IDLE_MARKER")
        .or_else(|_| std::env::var("CORROIDLEMARKER"))
    {
        Ok(p) => Some(p),
        Err(_) if cfg!(all(target_family = "rust9x", target_env = "msvc")) => {
            // Win9x: env vars unavailable; fixed diagnostic path.
            Some("c:\\corro.idle".to_string())
        }
        Err(_) => None,
    };
    if let Some(path) = idle_path {
        eprintln!("[corro] idle marker file: {path}");
        // TEMPORARY Win95 diagnosis: on the FIRST after-redraw (post-init, post
        // first frame), poll the input queue and log whether events arrive.
        #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
        {
            use std::sync::atomic::{AtomicBool, Ordering};
            static FIRST: AtomicBool = AtomicBool::new(true);
            rswidgets::backends::pancurses::set_after_redraw_callback(Box::new(move || {
                if FIRST.swap(false, Ordering::SeqCst) {
                    unsafe { input_probe_post() };
                }
                append_marker_raw(&path, b"idle\n");
            }));
        }
        #[cfg(not(all(target_family = "rust9x", target_env = "msvc")))]
        rswidgets::backends::pancurses::set_after_redraw_callback(Box::new(move || {
            append_marker_raw(&path, b"idle\n");
        }));
    }

    // Alt+letter shortcuts matching the ratatui reference: Alt+O/T/W/A/X open
    // specific File items/submenus.  The toolkit's generic alt-key callback
    // lets the app decide; the backend itself knows nothing about corro's menus.
    rswidgets::backends::pancurses::set_alt_key_callback(Box::new(|ch: char| {
        match ch.to_ascii_lowercase() {
            'o' => { rswidgets::backends::pancurses::open_menu(0, vec![], 0); true } // File -> Open file
            't' => { rswidgets::backends::pancurses::open_menu(0, vec![2], 0); true } // File -> Export
            'w' => { rswidgets::backends::pancurses::open_menu(0, vec![3], 0); true } // File -> Width
            'a' => { rswidgets::backends::pancurses::open_menu(0, vec![2], 2); true } // File -> Export -> ASCII table
            'x' => { rswidgets::backends::pancurses::open_menu(0, vec![3], 1); true } // File -> Width -> Column width
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
    render::fill_cells(&mut SpreadsheetSink::new(&spreadsheet), &display_rows, &col_ixs, &col_widths,
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
    // same Menu type is used by every rswidgets backend, so the menu
    // definitions in crate::gui::menu are not tied to the pancurses backend.
    let menubar_model = rswidgets::backends_pancurses_adapter::create_menu()?;
    for root in crate::gui::menu::menu_bar() {
        let sub = rswidgets::backends_pancurses_adapter::create_menu()?;
        crate::gui::menu::build_menu_model(&sub, root.submenu.as_deref().unwrap_or(&[]));
        menubar_model.append_submenu(root.label, &sub);
    }
    let _menubar = rswidgets::backends_pancurses_adapter::create_menubar(&menubar_model, std::ptr::null_mut())?;

    // Formula bar trailing: show app status text (matches ratatui's
    // mode_prompt_widget which appends "   ·  {status}" after the cell value).
    if !app.core.status.is_empty() {
        spreadsheet.set_formula_bar_trailing(&format!("   ·  {}", app.core.status));
    } else {
        spreadsheet.set_formula_bar_trailing("");
    }

    win.set_child(&spreadsheet);
    rswidgets::backends::pancurses::set_focus(spreadsheet.id());
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
    // One borrow flag per callback chain (all start cleared): each chain's
    // AppBorrow guard is independent, so sequential dispatch never trips
    // another chain's flag. See [`AppBorrow`](super::AppBorrow).

    // Wire menu item activation: when a menu item is chosen in the pancurses UI,
    // dispatch its action here.  "quit" is handled by the backend itself
    // (sets running=false); everything else performs the corresponding op on the
    // App or records status so the user sees the action fired.  Previously the
    // pancurses Enter handler just closed the menu and did nothing.
    let menu_ss = spreadsheet.clone();
    let display_rows_menu = display_rows_for_cb.clone();
    // Persist the pending format target across the Scope/Number/Align picks,
    // mirroring the ratatui menu (scope is chosen first, then the format).
    let pending_scope: std::rc::Rc<std::cell::RefCell<u8>> = std::rc::Rc::new(std::cell::RefCell::new(0));
    // Simple session clipboard for Copy/Cut/Paste.
    let clipboard: std::rc::Rc<std::cell::RefCell<String>> = std::rc::Rc::new(std::cell::RefCell::new(String::new()));
    // Clone for the Ctrl+C handler (the menu callback below moves `clipboard`).
    let ctrlc_clip = clipboard.clone();
    rswidgets::backends::pancurses::set_menu_action_callback(Box::new(move |name: String| {
        let app = app_from_raw(app_ptr);
        if let Some(label) = menu_action_needs_prompt(&name) {
            rswidgets::backends::pancurses::set_prompt(label, &name);
            return;
        }
        let result = dispatch_menu_action(
            app,
            &name,
            &mut *pending_scope.borrow_mut(),
            &mut *clipboard.borrow_mut(),
        );
        let mut apply_status = |s: &str| {
            if !s.is_empty() {
                app.core.status = s.to_string();
                menu_ss.set_formula_bar_trailing(&format!("   ·  {}", s));
            }
        };
        match result {
            MenuDispatch::Status(s) => apply_status(&s),
            MenuDispatch::Prompt(label, action) => {
                rswidgets::backends::pancurses::set_prompt(label, action);
            }
            MenuDispatch::About { status } => {
                show_info_dialog(" About ", &about_body());
                apply_status(&status);
            }
            MenuDispatch::HelpFull { status } => {
                show_info_dialog(" Help ", &help_body());
                apply_status(&status);
            }
            MenuDispatch::HelpKeybinds { status } => {
                show_info_dialog(
                    "Keybindings",
                    "F2 edit\narrows move\nEnter commit\nCtrl+G go-to\nCtrl+Q quit",
                );
                apply_status(&status);
            }
        }
        // The action may have mutated the workbook (Insert Date, Cut/Paste,
        // New/Rename sheet, sort, ...).  Re-fill the widget's cells from the
        // workbook so the grid reflects the change immediately instead of
        // staying stale until the next cursor move.
        refresh_viewport_after_action(
            app, &menu_ss, sid, &display_rows_menu,
            data_rows, data_cols, data_width, HEADER_ROWS,
        );
    }));

    // Ctrl+C copies the cursor cell (standard terminal copy; it does NOT quit).
    // The value goes to the in-app clipboard (for Paste) and, via OSC 52, to the
    // terminal's system clipboard so it can be pasted elsewhere.
    let ctrlc_ss = spreadsheet.clone();
    rswidgets::backends::pancurses::add_key_callback('\x03', Box::new(move || {
        let app = app_from_raw(app_ptr);
        let hr = HEADER_ROWS;
        let lm = MARGIN_COLS;
        let main_row = app.core.cursor.row.saturating_sub(hr) as u32;
        let main_col = app.core.cursor.col.saturating_sub(lm) as u32;
        let addr = CellAddr::Main { row: main_row, col: main_col };
        let val = app.core.workbook.active_sheet().grid.get(&addr).unwrap_or_default();
        if !val.is_empty() {
            *ctrlc_clip.borrow_mut() = val.clone();
            rswidgets::backends::pancurses::set_clipboard_text(&val);
            app.core.status = format!("Copied {} ({})", main_addr_label(main_row, main_col), val);
            ctrlc_ss.set_formula_bar_trailing(&format!("   ·  Copied {}", main_addr_label(main_row, main_col)));
        } else {
            app.core.status = format!("Nothing to copy at {}", main_addr_label(main_row, main_col));
        }
    }));

    // Prompt callback: perform the real file operation for path/name actions
    // (Open/Save As/Export) submitted via the TUI text prompt.
    let prompt_ss = spreadsheet.clone();
    let sid_prompt = sid;
    let display_rows_prompt = display_rows_for_cb.clone();
    rswidgets::backends::pancurses::set_prompt_callback(Box::new(move |action: String, text: String| {
        let app = app_from_raw(app_ptr);
        run_prompt_action(app, &action, &text);
        prompt_ss.set_formula_bar_trailing(&format!("   ·  {}", app.core.status));
        // A prompt action can replace the whole workbook (Open) or write cells
        // (Save As records state) — re-fill the widget from the workbook so the
        // grid reflects it immediately.
        refresh_viewport_after_action(
            app, &prompt_ss, sid_prompt, &display_rows_prompt,
            data_rows, data_cols, data_width, HEADER_ROWS,
        );
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
        let app = app_from_raw(app_ptr);

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
            let vp = Viewport::recompute(app, cursor, data_rows_cb, data_cols_cb, data_width_cb, hr_cb, MARGIN_COLS);
            // Re-read the sheet after width adjustments.
            let rec = app.core.workbook.active_sheet().clone();
            spreadsheet_set_border_title(sid, &vp.border_title(app.core.ops_applied));
            spreadsheet_set_row_labels(sid, vp.row_labels.clone());
            spreadsheet_set_column_layout(sid, vp.column_layout.clone());
            col_ixs_cb = vp.col_ixs.clone();
            spreadsheet_set_grid_config(sid, MARGIN_COLS as u32, vp.mc as u32);
            vp.refill(&mut SpreadsheetSink::new(&sheet_cb), &rec.grid, hr_cb, MARGIN_COLS, data_width_cb, cursor.row, cursor.col);
            if let Some(new_display_ri) = vp.display_rows.iter().position(|&r| r == cursor.row) {
                let cursor_addr = crate::addr::sheet_cursor_to_addr(
                    crate::addr::LogicalRow(cursor.row),
                    crate::addr::GlobalCol(cursor.col),
                    crate::addr::MainRows(vp.mr),
                    crate::addr::MainCols(vp.mc),
                );
                if let Some(raw_val) = rec.grid.get(&cursor_addr) {
                    sheet_cb.set_raw_cell(new_display_ri as u32, cursor.col as u32, &raw_val);
                } else {
                    sheet_cb.set_raw_cell(new_display_ri as u32, cursor.col as u32, "");
                }
                sheet_cb.set_cursor(new_display_ri as u32, cursor.col as u32);
            }
            *display_rows_for_cb.borrow_mut() = vp.display_rows.clone();
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
                // Grid grew — recompute visible columns and repaint (shared).
                let cursor = app.core.cursor;
                let vp = Viewport::recompute_columns(app, &display_rows_for_cb.borrow(), cursor, data_cols_cb, data_width_cb, hr_cb, MARGIN_COLS);
                let rec = app.core.workbook.active_sheet().clone();
                spreadsheet_set_border_title(sid, &vp.border_title(app.core.ops_applied));
                spreadsheet_set_row_labels(sid, vp.row_labels.clone());
                spreadsheet_set_column_layout(sid, vp.column_layout.clone());
                col_ixs_cb = vp.col_ixs.clone();
                spreadsheet_set_grid_config(sid, lm as u32, vp.mc as u32);
                vp.refill(&mut SpreadsheetSink::new(&sheet_cb), &rec.grid, hr_cb, MARGIN_COLS, data_width_cb, cursor.row, cursor.col);
            } else if !col_ixs_cb.contains(&(_display_col as usize)) {
                // Update column viewport when cursor column moves outside the
                // currently visible range (matching ratatui's per-frame recompute).
                let cursor = app.core.cursor;
                let vp = Viewport::recompute_columns(app, &display_rows_for_cb.borrow(), cursor, data_cols_cb, data_width_cb, hr_cb, MARGIN_COLS);
                let rec = app.core.workbook.active_sheet().clone();
                spreadsheet_set_border_title(sid, &vp.border_title(app.core.ops_applied));
                spreadsheet_set_column_layout(sid, vp.column_layout.clone());
                col_ixs_cb = vp.col_ixs.clone();
                vp.refill(&mut SpreadsheetSink::new(&sheet_cb), &rec.grid, hr_cb, MARGIN_COLS, data_width_cb, cursor.row, cursor.col);
            } else {
                // Cursor moved within the current viewport — refresh cells to ensure
                // formatted display values are used (commits overwrite cells with raw values).
                let cursor = app.core.cursor;
                let dr: Vec<usize> = display_rows_for_cb.borrow().clone();
                let rec = app.core.workbook.active_sheet().clone();
                let vp = Viewport::snapshot(app, &dr, &col_ixs_cb, hr_cb, MARGIN_COLS);
                // Keep the border op-count current: the backend fires this
                // callback after every Enter-commit, so this is where a
                // just-committed op first becomes visible. The commit path
                // itself must stay refill-free here: calling the full
                // refresh helper from inside the deferred commit drain
                // panics with "RefCell already borrowed" (observed).
                spreadsheet_set_border_title(sid, &vp.border_title(app.core.ops_applied));
                vp.refill(&mut SpreadsheetSink::new(&sheet_cb), &rec.grid, hr_cb, MARGIN_COLS, data_width_cb, cursor.row, cursor.col);
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
        let app = app_from_raw(app_ptr_ce);
        let dr = display_rows_for_ce.borrow();
        let logical_row = dr.get(display_row as usize).copied().unwrap_or(0);
        let main_row = logical_row.saturating_sub(hr_ce);
        let main_col = col.saturating_sub(lm_ce as u32);
        let addr = CellAddr::Main { row: main_row as u32, col: main_col as u32 };
        crate::debug_log::log(&format!(
            "COMMIT_CB display_row={} col={} logical_row={} hr_ce={} lm_ce={} main_row={} main_col={} addr={:?} value={:?} app_cursor_row={} app_cursor_col={}",
            display_row, col, logical_row, hr_ce, lm_ce, main_row, main_col, addr, value, app.core.cursor.row, app.core.cursor.col
        ));
        // Commit the edited value to the workbook via the shared helper
        // (logs to the live .corro file when one is open, else applies in-memory).
        commit_cell(app, addr, value);
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
        let app = app_from_raw(app_ptr_goto);
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
        // Recompute the viewport around the target cell (shared controller).
        let vp = Viewport::recompute(app, cursor, data_rows_cb, data_cols_cb, data_width_cb, hr_goto, MARGIN_COLS);
        let rec = app.core.workbook.active_sheet().clone();
        let target_logical = crate::grid::HEADER_ROWS + 999;
        let display_ri = vp.display_rows
            .iter()
            .position(|&r| r == target_logical)
            .unwrap_or(0);
        app.core.cursor = cursor;
        spreadsheet_set_border_title(sid_goto, &vp.border_title(app.core.ops_applied));
        spreadsheet_set_row_labels(sid_goto, vp.row_labels.clone());
        spreadsheet_set_column_layout(sid_goto, vp.column_layout.clone());
        spreadsheet_set_grid_config(sid_goto, lm_goto as u32, vp.mc as u32);
        vp.refill(&mut SpreadsheetSink::new(&goto_sheet), &rec.grid, hr_goto, MARGIN_COLS, data_width_cb, cursor.row, cursor.col);
        // Position the widget cursor on the target cell (A1000).
        goto_sheet.set_cursor(display_ri as u32, lm_goto as u32);
        let target_val = rec
            .grid
            .get(&CellAddr::Main { row: 999, col: 0 })
            .unwrap_or_default();
        goto_sheet.set_raw_cell(display_ri as u32, lm_goto as u32, &target_val);
        *dr_for_goto_cb.borrow_mut() = vp.display_rows.clone();
        *dr_for_goto_ce.borrow_mut() = vp.display_rows.clone();
    });

    _backend.run().map_err(|e| format!("pancurses error: {e}"))?;
    Ok(())
}




