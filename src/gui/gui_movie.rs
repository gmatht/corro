//! `--movie` replay for the rustxWidgets GUI backend.
//!
//! Split out of `gui_backend.rs` (which is already ~5k lines of interactive
//! widget wiring) because the movie path is a *driver*: it builds no widgets.
//! It renders each replayed step with the same backend-agnostic pipeline the
//! interactive canvas uses ([`crate::gui::compute`] + [`crate::gui::render`])
//! against a raster surface, so a recorded movie is pixel-identical to what
//! the interactive window paints for the same workbook state.
//!
//! Two entry points:
//!
//! * [`run_replay`] — the pure driver: apply every step of a
//!   [`GuiMovie`](crate::gui::movie::GuiMovie), paint each frame through a
//!   [`FramePainter`], and sleep the movie's pacing between them. No toolkit,
//!   no event loop, so it runs headless and is unit-tested.
//! * [`run_gui_movie_capture`] — the same driver with a PNG-writing painter,
//!   used by `--gui --movie --movie-frames <dir>` to record a video without an
//!   X server (see `scripts/gui_movie.py`, which encodes the frames with
//!   ffmpeg).
//!
//! The interactive GTK/nwg window is untouched: launching `--movie` without
//! `--movie-frames` falls back to the same capture path with a null painter,
//! which still prints progress but writes nothing.

use super::movie::{GuiMovie, GuiMovieOptions};
use crate::grid::{HEADER_ROWS, MARGIN_COLS};
use std::collections::HashMap;

/// Canvas geometry used to render movie frames, matching the interactive
/// GUI's pixel contract (see `gui_backend`: `FONT_SIZE`, `ROW_H`, ...).
const FRAME_W: i32 = 1200;
const FRAME_H: i32 = 800;
const FONT_SIZE: f64 = 12.0;
/// Height of the window chrome above the grid: menu bar + formula bar, the
/// same two strips the interactive window stacks over the canvas.
const CHROME_H: f64 = 56.0;
/// Height of the column-label strip inside the grid.
const HEADER_H: f64 = 24.0;
const ROW_H: f64 = 20.0;
const ROW_LABEL_W: f64 = 50.0;
const CHAR_W: f64 = 7.2;
/// Height of the status band painted over the bottom of a frame.
const STATUS_H: f64 = 30.0;

/// A raster surface a movie frame can be painted into.
pub trait FramePainter {
    /// Width/height in device pixels.
    fn size(&self) -> (i32, i32);
    /// True when this painter repaints a live surface (a window) rather than
    /// writing one file per frame. The driver only re-paints while holding a
    /// delay for live painters, so a capture does not duplicate frames.
    fn is_live(&self) -> bool {
        false
    }

    /// Paint one frame of `app`'s active sheet and save it as frame number
    /// `n` (0-based). `caption` is a human line for the frame (step status),
    /// which painters may draw or ignore.
    fn paint(&mut self, app: &crate::gui::App, frame: &MovieFrameView, n: usize) -> Result<(), String>;
}

/// What the driver knows about the step being painted.
#[derive(Clone, Debug)]
pub struct MovieFrameView {
    /// 1-based step number.
    pub step: usize,
    pub total: usize,
    pub status: String,
    /// Menu path a user would have taken for this step (`("Sheet",
    /// "Balance books")`), shown while the step's "menu moment" holds.
    pub menu: Option<(String, String)>,
    /// A cell the step wrote, which mouse-style painters may highlight.
    pub cursor: Option<crate::grid::CellAddr>,
}

/// Apply every step of `movie` to `app`, painting each frame and holding it
/// for the movie's pacing. Returns the number of frames painted.
///
/// This is the whole replay loop: the interactive backend and the capture
/// backend differ only in their [`FramePainter`], so both replay identically.
pub fn run_replay(
    app: &mut crate::gui::App,
    movie: &mut GuiMovie,
    options: GuiMovieOptions,
    painter: &mut dyn FramePainter,
) -> Result<usize, String> {
    let char_delay = options.char_delay();
    let confirm = options.confirm_delay();
    let menu_hold = options.menu_hold();
    let mut frames = 0usize;
    let total = movie.len();

    for i in 0..total {
        let step = movie.steps[i].clone();
        let sheet_before = app.core.workbook.active_sheet;
        let frame = movie
            .apply_step(&mut app.core.workbook, &mut app.core.view_sheet_id, i)
            .map_err(|e| format!("movie step {} failed: {e}", i + 1))?;

        // Keep the visible sheet in step with the workbook: a step may create,
        // activate or delete sheets, and the canvas reads active_sheet.
        if app.core.workbook.active_sheet != sheet_before {
            app.core.workbook.ensure_active_sheet();
        }
        app.core.ops_applied = movie.applied;
        app.core.state = app.core.workbook.active_sheet().clone();

        // Cursor: the cell the step wrote to, so the grid highlight animates
        // along with the replay instead of sitting at A1.
        if let Some(addr) = frame.cursor.as_ref() {
            app.core.cursor =
                super::movie::cursor_of(addr, &app.core.workbook);
            app.core.cursor.clamp(&app.core.workbook.active_sheet().grid);
        }
        let typed = typed_text_for(&step.line);
        let view = MovieFrameView {
            step: i + 1,
            total,
            status: frame.status.clone(),
            menu: frame.menu.clone(),
            cursor: frame.cursor.clone(),
        };

        // Menu moment: hold the menu flash on screen before the change lands,
        // exactly like the TUI's `movie_show_menu` (a step that opens a menu is
        // legible only if the menu is visible while nothing has changed yet).
        if let Some((section, item)) = view.menu.clone() {
            let mut held = view.clone();
            held.status = format!("{} ▸ {}", section, item);
            hold(&held, menu_hold, painter, app, &mut frames)?;
        }

        // Typing animation: for a `SET` line, hold one frame per typed
        // character with the partial text visible in the formula bar, which is
        // what makes a recorded movie look like someone using the app.
        if let Some(text) = typed {
            let mut partial = String::new();
            for ch in text.chars() {
                partial.push(ch);
                let mut typing = view.clone();
                typing.status = format!("{} = {}", addr_label(app, &view.cursor), partial);
                hold(&typing, char_delay, painter, app, &mut frames)?;
            }
        }

        hold(&view, confirm, painter, app, &mut frames)?;
    }

    // Final "movie complete" frame, matching the TUI's closing status.
    let done = MovieFrameView {
        step: total,
        total,
        status: format!("Movie complete: {} lines from {}", movie.applied, movie.path.display()),
        menu: None,
        cursor: None,
    };
    hold(&done, confirm, painter, app, &mut frames)?;
    Ok(frames)
}

/// Paint one frame and hold it on screen for `d`.
///
/// The hold is a single sleep, not a repaint loop: one frame *is* the delay,
/// and the frame rate of the recording comes from `--fps` at encode time. A
/// live painter additionally re-paints once after the sleep so a windowed
/// replay keeps showing the frame while the replay pauses — a capture painter
/// would only write the same image twice, so it skips that.
fn hold(
    view: &MovieFrameView,
    d: std::time::Duration,
    painter: &mut dyn FramePainter,
    app: &crate::gui::App,
    frames: &mut usize,
) -> Result<(), String> {
    painter.paint(app, view, *frames)?;
    *frames += 1;
    // A live painter (a window) needs the frame repainted after the pause so
    // it keeps showing it; a capture painter would only write the same image
    // twice. Zero delay means the frame is already the whole step.
    if d > std::time::Duration::ZERO && painter.is_live() {
        std::thread::sleep(d);
        painter.paint(app, view, *frames)?;
        *frames += 1;
    }
    Ok(())
}

/// Longest value worth animating character by character. Past this the typing
/// effect stops being legible (the frame rate can no longer keep up, so the
/// characters would not even be seen) and the recording balloons: a movie of
/// deliberately-huge overflow strings otherwise costs thousands of frames for
/// one cell. Longer values still replay — they just land in one frame.
const MAX_TYPED_CHARS: usize = 48;

/// The literal text a `SET` line types (the part after the address), so the
/// typing animation has something to animate. `None` for non-SET lines.
fn typed_text_for(line: &str) -> Option<String> {
    let rest = line.strip_prefix("SET ")?.trim_start();
    // `<sheet>:` prefix is optional.
    let rest = match rest.find(':') {
        Some(colon) if rest[..colon].trim_start().starts_with('$') => rest[colon + 1..].trim_start(),
        _ => rest,
    };
    let (addr, value) = rest.split_once(' ')?;
    if addr.contains('=') || addr.contains("FILL") {
        return None;
    }
    let value = value.trim_end();
    if value.chars().count() > MAX_TYPED_CHARS {
        return None;
    }
    Some(value.to_string())
}

fn addr_label(app: &crate::gui::App, cursor: &Option<crate::grid::CellAddr>) -> String {
    match cursor {
        Some(addr) => crate::addr::cell_ref_text(addr, app.core.workbook.active_sheet().grid.main_cols()),
        None => "?".into(),
    }
}

/// A painter that does nothing but count frames: the headless default, so
/// `--gui --movie` can still replay and report progress without a display.
#[derive(Default)]
pub struct NullPainter {
    pub painted: usize,
}

impl FramePainter for NullPainter {
    fn size(&self) -> (i32, i32) {
        (FRAME_W, FRAME_H)
    }
    fn paint(
        &mut self,
        _app: &crate::gui::App,
        _frame: &MovieFrameView,
        _n: usize,
    ) -> Result<(), String> {
        self.painted += 1;
        Ok(())
    }
}

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------

/// `--gui --movie` entry point: replay headlessly, optionally writing one PNG
/// per frame into `frames_dir` (which is how a video gets made without a
/// display server — see `scripts/gui_movie.py`).
pub fn run_gui_movie(
    app: &mut crate::gui::App,
    movie: GuiMovie,
    options: GuiMovieOptions,
) -> Result<(), Box<dyn std::error::Error>> {
    let frames_dir = std::env::var("CORRO_MOVIE_FRAMES").ok().map(std::path::PathBuf::from);
    let _ = &app;
    let mut painter: Box<dyn FramePainter> = match frames_dir {
        Some(dir) => Box::new(PngFramePainter::new(dir)),
        None => Box::new(NullPainter::default()),
    };
    let size = painter.size();
    let mut movie = movie;
    let painted = run_replay(app, &mut movie, options, painter.as_mut())?;
    eprintln!(
        "[corro] movie: {} steps, {} frames rendered at {}x{}",
        movie.len(),
        painted,
        size.0,
        size.1
    );
    Ok(())
}

/// A capturing "live" painter used by interactive GUI movie mode: it renders
/// each frame through the same headless pipeline the capture path uses, so a
/// windowed movie shows exactly the pixels a captured movie contains.
pub struct HeadlessCanvasPainter {
    /// Directory to write `frame-NNNNN.png` into (already created).
    pub dir: std::path::PathBuf,
}

impl FramePainter for HeadlessCanvasPainter {
    fn size(&self) -> (i32, i32) {
        (FRAME_W, FRAME_H)
    }
    fn paint(
        &mut self,
        app: &crate::gui::App,
        frame: &MovieFrameView,
        n: usize,
    ) -> Result<(), String> {
        paint_frame_png(app, frame, n, FRAME_W, FRAME_H, &self.dir)
    }
    fn is_live(&self) -> bool {
        true
    }
}

/// Writes one image per painted frame.
pub struct PngFramePainter {
    dir: std::path::PathBuf,
}

impl PngFramePainter {
    pub fn new(dir: std::path::PathBuf) -> Self {
        let _ = std::fs::create_dir_all(&dir);
        PngFramePainter { dir }
    }
}

impl FramePainter for PngFramePainter {
    fn size(&self) -> (i32, i32) {
        (FRAME_W, FRAME_H)
    }
    fn paint(
        &mut self,
        app: &crate::gui::App,
        frame: &MovieFrameView,
        n: usize,
    ) -> Result<(), String> {
        paint_frame_png(app, frame, n, FRAME_W, FRAME_H, &self.dir)
    }
}

/// Render one movie frame to `dir/frame-NNNNN.ppm`.
///
/// A binary PPM is used instead of PNG so the capture path needs no image
/// crate: it is the same lossless raster ffmpeg (and PIL, and ImageMagick)
/// read directly, and it is trivial to write from here.
fn paint_frame_png(
    app: &crate::gui::App,
    frame: &MovieFrameView,
    n: usize,
    w: i32,
    h: i32,
    dir: &std::path::Path,
) -> Result<(), String> {
    let mut dc = RasterDrawContext::new(w, h);
    paint_sheet(&mut dc, app, frame, w, h);
    let path = dir.join(format!("frame-{n:05}.ppm"));
    std::fs::write(&path, dc.into_ppm()).map_err(|e| format!("write {}: {e}", path.display()))
}

/// Paint the sheet the way the GUI canvas does — same viewport computation,
/// same `render::fill_cells` pipeline — onto a raster context.
fn paint_sheet(
    dc: &mut RasterDrawContext,
    app: &crate::gui::App,
    frame: &MovieFrameView,
    w: i32,
    h: i32,
) {
    use crate::gui::viewport::Viewport;

    // Background (mirrors render_grid's clear).
    dc.clear(0.94, 0.94, 0.94, 1.0);

    let cursor = app.core.cursor;
    // Sizing follows the interactive canvas: rows/columns derived from the
    // pixel extent so the movie frame is framed like the window.
    let data_rows = rows_for_height(h);
    // `data_cols` bounds the viewport search; `data_width` is the character
    // budget that decides how much of it survives trimming. Both are derived
    // from the frame size the way the live canvas derives them.
    let data_width = data_width_for(w);
    let data_cols = data_width / 2;
    // `Viewport::recompute` takes `&mut App` because trimming column widths
    // adjusts the grid; the movie must not mutate the workbook it is
    // replaying, so the viewport is computed against a private copy of the
    // state and only the resulting layout is used.
    let vp = {
        let mut vp_app = app.copy_for_layout();
        Viewport::recompute(
            &mut vp_app,
            cursor,
            data_rows,
            data_cols,
            data_width,
            HEADER_ROWS,
            MARGIN_COLS,
        )
    };

    let sheet = app.core.workbook.active_sheet();
    let grid = &sheet.grid;
    let mr = grid.main_rows();
    let mc = grid.main_cols();

    // Stretch the visible columns across the frame, the way the window does
    // when the sheet is narrower than the canvas. Without this a small
    // workbook renders as a narrow ribbon with an empty right half — the
    // "huge empty spaces" the window had before it sized the viewport from
    // its live canvas size.
    let stretched = stretch_columns_to_width(&vp, w);
    let col_widths = if stretched.is_empty() { vp.col_widths.clone() } else { stretched };

    // Window chrome: menu bar, then formula bar (address + value), matching
    // the interactive window's layout (see `run_gui`: vbox = menu, formula
    // bar, canvas, hints).
    paint_menu_bar(dc, w);
    paint_formula_bar(dc, app, frame, w);

    // Row gutter, column header strip and the cell body.
    paint_gutter(dc, &vp.display_rows, mr);
    paint_headers(dc, &vp.column_layout, &col_widths);
    let mut sink = MovieSink::new();
    vp.refill(
        &mut sink,
        grid,
        HEADER_ROWS,
        MARGIN_COLS,
        data_width,
        cursor.row,
        cursor.col,
    );
    paint_cells(dc, &sink, &vp, &col_widths, cursor.row, cursor.col);

    // Status band: the movie caption (what is being replayed right now).
    let caption = format!(
        "corro movie  ·  step {}/{}  ·  {}",
        frame.step, frame.total, frame.status
    );
    paint_status(dc, &caption, frame, w, h);
}

/// Rows needed to cover a canvas `h` pixels tall. Mirrors the live canvas
/// (`gui_backend::rows_to_fill_px`): the grid starts below the chrome and the
/// column strip, and the status band sits over the bottom of the frame.
fn rows_for_height(h: i32) -> usize {
    let usable = h as f64 - CHROME_H - HEADER_H - STATUS_H;
    ((usable / ROW_H + 1.0).max(1.0)) as usize
}

/// Widen the visible columns so the sheet spans the frame.
///
/// The window sizes its viewport from the live canvas and lets columns use the
/// whole width; a movie frame must do the same or a small workbook renders as
/// a narrow ribbon with a blank right half (the "huge empty spaces" bug).
/// Only the *main* columns are stretched — margin columns keep their natural
/// width, since they are chrome, not data.
fn stretch_columns_to_width(
    vp: &crate::gui::viewport::Viewport,
    w: i32,
) -> HashMap<usize, usize> {
    let mut widths: HashMap<usize, usize> = vp.col_widths.clone();
    let main: Vec<usize> = vp
        .col_ixs
        .iter()
        .copied()
        .filter(|c| *c >= MARGIN_COLS && *c < MARGIN_COLS + vp.mc)
        .collect();
    if main.is_empty() {
        return widths;
    }
    let total_px: f64 = vp
        .col_ixs
        .iter()
        .map(|c| col_px(*widths.get(c).unwrap_or(&4)))
        .sum();
    let spare = (w as f64 - ROW_LABEL_W) - total_px;
    if spare <= 0.0 {
        return widths;
    }
    // Spread the spare width evenly, in whole characters so text still aligns
    // to the grid.
    let per_col_chars = (spare / (main.len() as f64 * CHAR_W)).floor().max(0.0) as usize;
    if per_col_chars == 0 {
        return widths;
    }
    for c in main {
        let cur = *widths.get(&c).unwrap_or(&4);
        widths.insert(c, cur + per_col_chars);
    }
    widths
}

/// Character width available to the sheet (row gutter + right margin
/// reserved), the unit `ui_core::trim_visible_cols_to_width` expects. This is
/// a *character* count, not a column count — passing a column count here
/// trims the sheet down to the margin columns and leaves the rest of the
/// frame blank.
fn data_width_for(w: i32) -> usize {
    let usable = w as f64 - ROW_LABEL_W;
    (usable / CHAR_W).max(1.0) as usize
}

/// Menu bar strip: the same labels `crate::gui::menu::menu_bar_text()` renders
/// in the window, with the movie's current menu highlighted.
fn paint_menu_bar(dc: &mut RasterDrawContext, w: i32) {
    dc.fill_rect(0.0, 0.0, w as f64, 26.0, 0.95, 0.95, 0.95, 1.0);
    dc.fill_rect(0.0, 26.0, w as f64, 1.0, 0.78, 0.78, 0.78, 1.0);
    dc.draw_text(8.0, 8.0, &crate::gui::menu::menu_bar_text(), FONT_SIZE, 0.15, 0.15, 0.18, 1.0);
}

/// Formula bar strip: the cursor's address and raw value, plus the movie
/// caption. Mirrors `run_gui`'s address label + `fx` entry + status suffix.
fn paint_formula_bar(
    dc: &mut RasterDrawContext,
    app: &crate::gui::App,
    frame: &MovieFrameView,
    w: i32,
) {
    dc.fill_rect(0.0, 27.0, w as f64, 28.0, 1.0, 1.0, 1.0, 1.0);
    dc.fill_rect(0.0, 55.0, w as f64, 1.0, 0.78, 0.78, 0.78, 1.0);

    let grid = &app.core.workbook.active_sheet().grid;
    let addr = crate::addr::sheet_cursor_to_addr(
        crate::addr::LogicalRow(app.core.cursor.row),
        crate::addr::GlobalCol(app.core.cursor.col),
        crate::addr::MainRows(grid.main_rows()),
        crate::addr::MainCols(grid.main_cols()),
    );
    let label = crate::addr::cell_ref_text(&addr, grid.main_cols());
    let value = grid.get(&addr).unwrap_or_default();

    dc.draw_text(8.0, 35.0, &label, FONT_SIZE, 0.15, 0.15, 0.18, 1.0);
    dc.draw_text(70.0, 35.0, "fx", FONT_SIZE, 0.45, 0.45, 0.5, 1.0);
    let shown: String = value.chars().take(80).collect();
    dc.draw_text(92.0, 35.0, &shown, FONT_SIZE, 0.0, 0.0, 0.0, 1.0);

    // Status suffix (the movie caption) on the right of the formula row, the
    // same place the window puts its `· status` label.
    let caption = format!("· {}", frame.status);
    let caption: String = caption.chars().take(60).collect();
    let cx = (w as f64 - 8.0 - caption.chars().count() as f64 * glyph_advance(2.0)).max(200.0);
    dc.draw_text(cx, 35.0, &caption, FONT_SIZE, 0.3, 0.3, 0.45, 1.0);
}

fn paint_gutter(dc: &mut RasterDrawContext, display_rows: &[usize], mr: usize) {
    const ROW_H: f64 = 20.0;
    const ROW_LABEL_W: f64 = 50.0;
    for (ri, &logical_row) in display_rows.iter().enumerate() {
        let y = CHROME_H + HEADER_H + ri as f64 * ROW_H;
        dc.fill_rect(0.0, y, ROW_LABEL_W, ROW_H, 0.9, 0.9, 0.9, 1.0);
        let label = crate::addr::ui_row_label(logical_row, mr);
        dc.draw_text(6.0, y + 3.0, &label, FONT_SIZE, 0.3, 0.3, 0.3, 1.0);
    }
}

fn paint_headers(dc: &mut RasterDrawContext, layout: &[(u32, u32, String)], col_widths: &HashMap<usize, usize>) {
    let mut x = ROW_LABEL_W;
    for (c, width, label) in layout {
        let cw = col_px(*col_widths.get(&(*c as usize)).unwrap_or(&(*width as usize)));
        dc.fill_rect(x, CHROME_H, cw, HEADER_H, 0.9, 0.9, 0.9, 1.0);
        dc.draw_text(x + 2.0, CHROME_H + 5.0, label, FONT_SIZE, 0.3, 0.3, 0.3, 1.0);
        x += cw;
    }
}

fn paint_cells(
    dc: &mut RasterDrawContext,
    sink: &MovieSink,
    vp: &crate::gui::viewport::Viewport,
    col_widths: &HashMap<usize, usize>,
    cursor_row: usize,
    cursor_col: usize,
) {
    for (ri, &logical_row) in vp.display_rows.iter().enumerate() {
        let y = CHROME_H + HEADER_H + ri as f64 * ROW_H;
        let mut x = ROW_LABEL_W;
        for &c in vp.col_ixs.iter() {
            let cw = col_px(*col_widths.get(&c).unwrap_or(&8));
            let is_cursor = logical_row == cursor_row && c == cursor_col;
            let key = (ri as u32, c as u32);
            let text = sink.cells.get(&key).cloned().unwrap_or_default();
            let bg = if is_cursor {
                (0.8, 0.9, 1.0, 1.0)
            } else {
                (1.0, 1.0, 1.0, 1.0)
            };
            dc.fill_rect(x, y, cw, ROW_H, bg.0, bg.1, bg.2, bg.3);
            dc.stroke_rect(x, y, cw, ROW_H, 0.78, 0.78, 0.78, 1.0, 1.0);
            if !text.trim().is_empty() {
                // Cells hold at most `col_width` characters (that is what the
                // grid guarantees and what the live canvas relies on); drawing
                // the whole string would run it across the neighbouring cells.
                let shown: String = text.chars().take(cols_in(cw)).collect();
                dc.draw_text(x + 2.0, y + 3.0, &shown, FONT_SIZE, 0.0, 0.0, 0.0, 1.0);
            }
            x += cw;
        }
    }
}

/// Pixel width of a grid column. One layout character (`CHAR_W`) per character
/// plus a small gutter, so a column that holds N characters is wide enough to
/// draw them with the bitmap font's advance.
fn col_px(width_chars: usize) -> f64 {
    (width_chars as f64 * CHAR_W) + 4.0
}

/// How many characters fit in a column of `px` pixels (the inverse of
/// [`col_px`], using the text advance rather than the layout metric).
fn cols_in(px: f64) -> usize {
    (((px - 4.0) / glyph_advance(2.0)).floor() as usize).max(1)
}

fn paint_status(dc: &mut RasterDrawContext, caption: &str, frame: &MovieFrameView, w: i32, h: i32) {
    let y = h as f64 - STATUS_H;
    dc.fill_rect(0.0, y, w as f64, STATUS_H, 0.12, 0.14, 0.18, 1.0);
    dc.draw_text(8.0, y + 7.0, caption, FONT_SIZE, 1.0, 1.0, 1.0, 1.0);
    // Progress bar: replay position as a fraction of the movie.
    if frame.total > 0 {
        let frac = frame.step as f64 / frame.total as f64;
        let bar_w = (w as f64 - 16.0) * frac;
        dc.fill_rect(8.0, y + STATUS_H - 5.0, bar_w, 3.0, 0.25, 0.6, 0.95, 1.0);
    }
    if let Some((section, item)) = frame.menu.as_ref() {
        // Menu flash: a small popup under the top-left corner, like an open
        // menu, so menu-driven steps are visible in the recording.
        let label = format!("{} ▸ {}", section, item);
        let box_w = (label.chars().count() as f64 * 7.2) + 18.0;
        dc.fill_rect(50.0, 24.0, box_w, 24.0, 0.16, 0.18, 0.22, 0.95);
        dc.stroke_rect(50.0, 24.0, box_w, 24.0, 0.35, 0.55, 0.85, 1.0, 1.0);
        dc.draw_text(58.0, 30.0, &label, FONT_SIZE, 1.0, 1.0, 1.0, 1.0);
    }
}

/// A `CellSink` for the movie painter: keeps the formatted text of each
/// visible cell (the same values the interactive canvas draws).
struct MovieSink {
    cells: std::collections::HashMap<(u32, u32), String>,
}

impl MovieSink {
    fn new() -> Self {
        MovieSink {
            cells: std::collections::HashMap::new(),
        }
    }
}

impl crate::gui::render::CellSink for MovieSink {
    fn set_cell(&mut self, row: u32, col: u32, text: &str) {
        self.cells.insert((row, col), text.to_string());
    }
    fn set_cell_style(
        &mut self,
        _row: u32,
        _col: u32,
        _style: crate::gui::compute::CellDisplayStyle,
    ) {
    }
    fn set_raw_cell(&mut self, _row: u32, _col: u32, _text: &str) {}
    fn set_cursor(&mut self, _row: u32, _col: u32) {}
}

// ---------------------------------------------------------------------------
// Raster surface
// ---------------------------------------------------------------------------

/// A minimal RGB raster with just the primitives the sheet painter needs
/// (filled/stroked rectangles and text). Text advances by a fixed 7.2px per
/// character — the same monospace metric the interactive canvas assumes — so
/// the raster matches the widget layout without font measurement.
pub struct RasterDrawContext {
    w: i32,
    h: i32,
    px: Vec<u8>,
}

impl RasterDrawContext {
    pub fn new(w: i32, h: i32) -> Self {
        let w = w.max(1);
        let h = h.max(1);
        RasterDrawContext {
            w,
            h,
            px: vec![255u8; (w as usize) * (h as usize) * 3],
        }
    }

    pub fn size(&self) -> (i32, i32) {
        (self.w, self.h)
    }

    pub fn clear(&mut self, r: f64, g: f64, b: f64, _a: f64) {
        let (r, g, b) = (chan(r), chan(g), chan(b));
        for i in (0..self.px.len()).step_by(3) {
            self.px[i] = r;
            self.px[i + 1] = g;
            self.px[i + 2] = b;
        }
    }

    pub fn fill_rect(&mut self, x: f64, y: f64, w: f64, h: f64, r: f64, g: f64, b: f64, a: f64) {
        let (r0, g0, b0, a0) = (chan(r), chan(g), chan(b), a.clamp(0.0, 1.0));
        let x0 = x.floor() as i32;
        let y0 = y.floor() as i32;
        let x1 = (x + w).ceil() as i32;
        let y1 = (y + h).ceil() as i32;
        for py in y0.max(0)..y1.min(self.h) {
            for pxx in x0.max(0)..x1.min(self.w) {
                self.blend(pxx, py, r0, g0, b0, a0);
            }
        }
    }

    pub fn stroke_rect(
        &mut self,
        x: f64,
        y: f64,
        w: f64,
        h: f64,
        r: f64,
        g: f64,
        b: f64,
        a: f64,
        lw: f64,
    ) {
        let t = lw.max(0.5);
        self.fill_rect(x, y, w, t, r, g, b, a);
        self.fill_rect(x, y + h - t, w, t, r, g, b, a);
        self.fill_rect(x, y, t, h, r, g, b, a);
        self.fill_rect(x + w - t, y, t, h, r, g, b, a);
    }

    /// Draw single-line text with a fixed monospace advance. Uses the built-in
    /// 5x7 bitmap font so the capture path needs no font library.
    ///
    /// The advance is the glyph width plus a gap — `CHAR_W` (7.2px) is the
    /// *layout* metric that decides how many characters fit in a column, but
    /// drawing at that pitch makes 2x glyphs overlap into an unreadable smear.
    pub fn draw_text(
        &mut self,
        x: f64,
        y: f64,
        text: &str,
        size: f64,
        r: f64,
        g: f64,
        b: f64,
        a: f64,
    ) {
        let scale = if size >= 11.0 { 2.0 } else { 1.0 };
        let advance = glyph_advance(scale);
        let mut cx = x;
        for ch in text.chars() {
            if ch != ' ' {
                if let Some(glyph) = glyph_for(ch) {
                    for (row, bits) in glyph.iter().enumerate() {
                        for col in 0..5u8 {
                            if bits & (1 << (4 - col)) != 0 {
                                let gx = cx + col as f64 * scale;
                                let gy = y + row as f64 * scale;
                                self.fill_rect(gx, gy, scale, scale, r, g, b, a);
                            }
                        }
                    }
                }
            }
            cx += advance;
        }
    }

    fn blend(&mut self, x: i32, y: i32, r: u8, g: u8, b: u8, a: f64) {
        if x < 0 || y < 0 || x >= self.w || y >= self.h {
            return;
        }
        let i = ((y as usize) * (self.w as usize) + x as usize) * 3;
        let mix = |old: u8, new: u8| -> u8 {
            if a >= 1.0 {
                new
            } else {
                (old as f64 * (1.0 - a) + new as f64 * a).round() as u8
            }
        };
        self.px[i] = mix(self.px[i], r);
        self.px[i + 1] = mix(self.px[i + 1], g);
        self.px[i + 2] = mix(self.px[i + 2], b);
    }

    /// Serialize as a binary PPM (P6) — lossless and read by ffmpeg, PIL and
    /// ImageMagick without any image crate in the build.
    pub fn into_ppm(self) -> Vec<u8> {
        let header = format!("P6\n{} {}\n255\n", self.w, self.h);
        let mut out = Vec::with_capacity(header.len() + self.px.len());
        out.extend_from_slice(header.as_bytes());
        out.extend_from_slice(&self.px);
        out
    }
}

fn chan(v: f64) -> u8 {
    (v.clamp(0.0, 1.0) * 255.0).round() as u8
}

/// Horizontal pitch for one bitmap character at `scale`: the 5px glyph plus a
/// 2px gap, so neighbouring glyphs stay distinct.
fn glyph_advance(scale: f64) -> f64 {
    (5.0 + 2.0) * scale
}

/// 5x7 bitmap glyphs for the characters a movie frame needs (ASCII subset).
/// Returns row bitmaps, MSB = leftmost column.
fn glyph_for(ch: char) -> Option<[u8; 7]> {
    let g = match ch.to_ascii_uppercase() {
        'A' => [0x0E, 0x11, 0x11, 0x1F, 0x11, 0x11, 0x11],
        'B' => [0x1E, 0x11, 0x11, 0x1E, 0x11, 0x11, 0x1E],
        'C' => [0x0E, 0x11, 0x10, 0x10, 0x10, 0x11, 0x0E],
        'D' => [0x1E, 0x11, 0x11, 0x11, 0x11, 0x11, 0x1E],
        'E' => [0x1F, 0x10, 0x10, 0x1E, 0x10, 0x10, 0x1F],
        'F' => [0x1F, 0x10, 0x10, 0x1E, 0x10, 0x10, 0x10],
        'G' => [0x0E, 0x11, 0x10, 0x17, 0x11, 0x11, 0x0F],
        'H' => [0x11, 0x11, 0x11, 0x1F, 0x11, 0x11, 0x11],
        'I' => [0x0E, 0x04, 0x04, 0x04, 0x04, 0x04, 0x0E],
        'J' => [0x07, 0x02, 0x02, 0x02, 0x02, 0x12, 0x0C],
        'K' => [0x11, 0x12, 0x14, 0x18, 0x14, 0x12, 0x11],
        'L' => [0x10, 0x10, 0x10, 0x10, 0x10, 0x10, 0x1F],
        'M' => [0x11, 0x1B, 0x15, 0x15, 0x11, 0x11, 0x11],
        'N' => [0x11, 0x19, 0x15, 0x13, 0x11, 0x11, 0x11],
        'O' => [0x0E, 0x11, 0x11, 0x11, 0x11, 0x11, 0x0E],
        'P' => [0x1E, 0x11, 0x11, 0x1E, 0x10, 0x10, 0x10],
        'Q' => [0x0E, 0x11, 0x11, 0x11, 0x15, 0x12, 0x0D],
        'R' => [0x1E, 0x11, 0x11, 0x1E, 0x14, 0x12, 0x11],
        'S' => [0x0F, 0x10, 0x10, 0x0E, 0x01, 0x01, 0x1E],
        'T' => [0x1F, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04],
        'U' => [0x11, 0x11, 0x11, 0x11, 0x11, 0x11, 0x0E],
        'V' => [0x11, 0x11, 0x11, 0x11, 0x11, 0x0A, 0x04],
        'W' => [0x11, 0x11, 0x11, 0x15, 0x15, 0x15, 0x0A],
        'X' => [0x11, 0x11, 0x0A, 0x04, 0x0A, 0x11, 0x11],
        'Y' => [0x11, 0x11, 0x0A, 0x04, 0x04, 0x04, 0x04],
        'Z' => [0x1F, 0x01, 0x02, 0x04, 0x08, 0x10, 0x1F],
        '0' => [0x0E, 0x11, 0x13, 0x15, 0x19, 0x11, 0x0E],
        '1' => [0x04, 0x0C, 0x04, 0x04, 0x04, 0x04, 0x0E],
        '2' => [0x0E, 0x11, 0x01, 0x02, 0x04, 0x08, 0x1F],
        '3' => [0x1F, 0x02, 0x04, 0x02, 0x01, 0x11, 0x0E],
        '4' => [0x02, 0x06, 0x0A, 0x12, 0x1F, 0x02, 0x02],
        '5' => [0x1F, 0x10, 0x1E, 0x01, 0x01, 0x11, 0x0E],
        '6' => [0x06, 0x08, 0x10, 0x1E, 0x11, 0x11, 0x0E],
        '7' => [0x1F, 0x01, 0x02, 0x04, 0x08, 0x08, 0x08],
        '8' => [0x0E, 0x11, 0x11, 0x0E, 0x11, 0x11, 0x0E],
        '9' => [0x0E, 0x11, 0x11, 0x0F, 0x01, 0x02, 0x0C],
        '.' => [0x00, 0x00, 0x00, 0x00, 0x00, 0x0C, 0x0C],
        ',' => [0x00, 0x00, 0x00, 0x00, 0x0C, 0x04, 0x08],
        ':' => [0x00, 0x0C, 0x0C, 0x00, 0x0C, 0x0C, 0x00],
        ';' => [0x00, 0x0C, 0x0C, 0x00, 0x0C, 0x04, 0x08],
        '-' => [0x00, 0x00, 0x00, 0x1F, 0x00, 0x00, 0x00],
        '+' => [0x00, 0x04, 0x04, 0x1F, 0x04, 0x04, 0x00],
        '=' => [0x00, 0x00, 0x1F, 0x00, 0x1F, 0x00, 0x00],
        '*' => [0x00, 0x0A, 0x04, 0x1F, 0x04, 0x0A, 0x00],
        '/' => [0x01, 0x02, 0x02, 0x04, 0x08, 0x08, 0x10],
        '\\' => [0x10, 0x08, 0x08, 0x04, 0x02, 0x02, 0x01],
        '(' => [0x02, 0x04, 0x08, 0x08, 0x08, 0x04, 0x02],
        ')' => [0x08, 0x04, 0x02, 0x02, 0x02, 0x04, 0x08],
        '[' => [0x0E, 0x08, 0x08, 0x08, 0x08, 0x08, 0x0E],
        ']' => [0x0E, 0x02, 0x02, 0x02, 0x02, 0x02, 0x0E],
        '_' => [0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x1F],
        '~' => [0x00, 0x00, 0x0A, 0x15, 0x00, 0x00, 0x00],
        '|' => [0x04, 0x04, 0x04, 0x04, 0x04, 0x04, 0x04],
        '"' => [0x0A, 0x0A, 0x00, 0x00, 0x00, 0x00, 0x00],
        '\'' => [0x04, 0x04, 0x00, 0x00, 0x00, 0x00, 0x00],
        '!' => [0x04, 0x04, 0x04, 0x04, 0x04, 0x00, 0x04],
        '?' => [0x0E, 0x11, 0x01, 0x02, 0x04, 0x00, 0x04],
        '#' => [0x0A, 0x1F, 0x0A, 0x0A, 0x1F, 0x0A, 0x00],
        '$' => [0x04, 0x0F, 0x14, 0x0E, 0x05, 0x1E, 0x04],
        '%' => [0x18, 0x19, 0x02, 0x04, 0x08, 0x13, 0x03],
        '&' => [0x0C, 0x12, 0x14, 0x08, 0x15, 0x12, 0x0D],
        '<' => [0x02, 0x04, 0x08, 0x10, 0x08, 0x04, 0x02],
        '>' => [0x08, 0x04, 0x02, 0x01, 0x02, 0x04, 0x08],
        '^' => [0x04, 0x0A, 0x11, 0x00, 0x00, 0x00, 0x00],
        _ => return None,
    };
    Some(g)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::gui::movie::{GuiMovie, GuiMovieOptions};

    fn write_movie(text: &str, tag: &str) -> std::path::PathBuf {
        let path = std::env::temp_dir().join(format!("corro-gui-movie-{tag}-{}.corro", std::process::id()));
        std::fs::write(&path, text).unwrap();
        path
    }

    #[test]
    fn replay_paints_one_frame_per_step_without_a_display() {
        let path = write_movie("SET $1:A1 42\nSET $1:A2 7\n", "basic");
        let mut movie = GuiMovie::new(&path).unwrap();
        let mut app = crate::gui::App::new_with_paths(vec![path.clone()]);
        let mut painter = NullPainter::default();
        let opts = GuiMovieOptions {
            typing_cps: 1000.0,
            confirm_delay_ms: 0,
            menu_hold_ms: 0,
        };
        let frames = run_replay(&mut app, &mut movie, opts, &mut painter).unwrap();
        // Two steps (each: optional menu frame + typing frames + confirm) plus
        // the closing "movie complete" frame.
        assert!(frames >= 3, "expected at least one frame per step, got {frames}");
        assert_eq!(painter.painted, frames);
        // The workbook really changed.
        assert_eq!(
            app.core.workbook.active_sheet().grid.get(&crate::grid::CellAddr::Main { row: 0, col: 0 }),
            Some("42".to_string())
        );
        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn frame_painter_writes_lossless_ppm_frames() {
        let path = write_movie("SET $1:A1 1\n", "ppm");
        let mut movie = GuiMovie::new(&path).unwrap();
        let mut app = crate::gui::App::new_with_paths(vec![path.clone()]);
        let dir = std::env::temp_dir().join(format!("corro-gui-movie-frames-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).unwrap();
        let mut painter = PngFramePainter::new(dir.clone());
        let size = painter.size();
        let opts = GuiMovieOptions {
            typing_cps: 1000.0,
            confirm_delay_ms: 0,
            menu_hold_ms: 0,
        };
        let frames = run_replay(&mut app, &mut movie, opts, &mut painter).unwrap();
        let written: Vec<_> = std::fs::read_dir(&dir)
            .unwrap()
            .filter_map(|e| e.ok())
            .map(|e| e.file_name().to_string_lossy().into_owned())
            .collect();
        assert_eq!(written.len(), frames, "one file per painted frame");
        let first = std::fs::read(dir.join("frame-00000.ppm")).unwrap();
        assert_eq!(&first[..2], b"P6", "frames are binary PPMs");
        assert!(
            first.len() > (size.0 * size.1) as usize,
            "frame carries pixel data"
        );
        let _ = std::fs::remove_dir_all(dir);
        let _ = std::fs::remove_file(path);
    }

    #[test]
    fn raster_draw_context_serializes_ppm_header() {
        let dc = RasterDrawContext::new(4, 2);
        let ppm = dc.into_ppm();
        let header = String::from_utf8_lossy(&ppm[..ppm.len() - 24]);
        assert!(header.starts_with("P6\n4 2\n255\n"), "header was {header:?}");
    }

    #[test]
    fn typed_text_extraction_handles_sheet_prefixes_and_fills() {
        assert_eq!(typed_text_for("SET $1:A1 42").as_deref(), Some("42"));
        assert_eq!(typed_text_for("SET A1 hello world").as_deref(), Some("hello world"));
        assert_eq!(typed_text_for("FILL $1:A1:B2 3"), None);
        assert_eq!(typed_text_for("NEW_SHEET Budget"), None);
    }

    /// A value longer than `MAX_TYPED_CHARS` replays in one frame instead of
    /// one frame per character: the animation would not be legible at that
    /// length and the recording would grow without bound.
    #[test]
    fn very_long_values_are_not_animated_per_character() {
        let long = "x".repeat(MAX_TYPED_CHARS + 1);
        assert_eq!(typed_text_for(&format!("SET A1 {long}")), None);
        let at_limit = "y".repeat(MAX_TYPED_CHARS);
        assert!(typed_text_for(&format!("SET A1 {at_limit}")).is_some());
    }
}
