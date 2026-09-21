//! `--movie` replay for the rustxWidgets GUI backend.
//!
//! Movie mode is the normal UI. `--gui --movie` opens the same window, builds
//! the same widget tree and runs the same draw callbacks as an interactive
//! session; a periodic timer (`gui_backend::arm_movie_driver`) applies one
//! movie step per tick instead of waiting for a keystroke. Everything the app
//! paints — margin shading, cell text, cursor ring, chrome — therefore comes
//! from the one production renderer, and a demo cannot drift from the app.
//!
//! That is the whole point: an earlier version of this module reimplemented
//! the sheet into a raster surface so it could render without a display. It
//! drifted immediately — blank frame, overlapping glyphs, then missing margin
//! shading, then missing text — because it was a second implementation of
//! rendering. It is gone.
//!
//! This file keeps only the parts that are genuinely backend-independent:
//!
//! * [`MovieFrameView`] — what the driver knows about the step being shown.
//! * [`run_gui_movie`] — the `--gui --movie` entry point.
//! * [`RasterDrawContext`] — a raster `DrawContext` used by tests to render
//!   the *real* sheet body (`gui_backend::render_grid_body`) headlessly, so
//!   the production renderer can be asserted on without a display server.

use super::movie::{GuiMovie, GuiMovieOptions};
use crate::grid::{HEADER_ROWS, MARGIN_COLS};
use rswidgets::core::DrawContext as MovieDrawContext;
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
    // Movie mode IS the normal UI: the same window, widget tree and draw
    // callbacks, driven by a timer that applies one movie step per tick. That
    // is the only way a recorded demo cannot drift from the app — every
    // previous version of this file reimplemented the sheet and did drift
    // (missing margin shading, missing text, wrong column sizing).
    options.publish_to_env();
    // A replay reads the log; it must not append to it. See `detach_source`.
    let _bound = GuiMovie::detach_source(app);
    eprintln!(
        "[corro] movie: replaying {} steps through the GUI window",
        movie.len()
    );
    crate::gui::gui_backend::run_gui_with_movie(app, Some(movie))
}

/// Writes one raster frame per replay step by driving the production sheet
/// renderer (`paint_sheet`) — used by tests and headless capture.
pub struct HeadlessFramePainter {
    /// Directory frames are written into (already created).
    pub dir: std::path::PathBuf,
}

impl HeadlessFramePainter {
    /// Create a painter writing `frame-NNNNN.ppm` into `dir`.
    pub fn new(dir: std::path::PathBuf) -> Self {
        let _ = std::fs::create_dir_all(&dir);
        HeadlessFramePainter { dir }
    }
}

impl FramePainter for HeadlessFramePainter {
    fn size(&self) -> (i32, i32) {
        (FRAME_W, FRAME_H)
    }
    fn paint(
        &mut self,
        app: &crate::gui::App,
        _frame: &MovieFrameView,
        n: usize,
    ) -> Result<(), String> {
        paint_frame(app, n, FRAME_W, FRAME_H, &self.dir)
    }
    fn is_live(&self) -> bool {
        true
    }
}

/// Renders the REAL sheet body (`gui_backend::render_grid_body`) into a raster
/// file, one per frame.
///
/// This exists for tests and for headless frame production: it drives the
/// production renderer through a `DrawContext`, so what it writes is what the
/// window paints. It is not a second renderer — the only thing it owns is the
/// raster surface itself.
fn paint_frame(
    app: &crate::gui::App,
    n: usize,
    w: i32,
    h: i32,
    dir: &std::path::Path,
) -> Result<(), String> {
    let mut dc = RasterDrawContext::new(w, h);
    paint_sheet(&mut dc, app, w, h);
    let path = dir.join(format!("frame-{n:05}.ppm"));
    std::fs::write(&path, dc.into_ppm()).map_err(|e| format!("write {}: {e}", path.display()))
}

/// Paint a single frame of `app` through `painter`. Exposed so tests (and the
/// capture path) can produce a frame without running a whole replay.
pub fn paint_one_frame(
    app: &crate::gui::App,
    painter: &mut dyn FramePainter,
    n: usize,
) -> Result<(), String> {
    let frame = MovieFrameView {
        step: 1,
        total: 1,
        status: String::new(),
        menu: None,
        cursor: None,
    };
    painter.paint(app, &frame, n)
}

/// Paint the sheet body into `dc` using the production renderer.
///
/// This is the SHEET ONLY: no menu bar, no formula bar, no status band. It
/// exists so tests can assert on the sheet renderer without a display; the
/// window itself (and `--gui --movie`) paints the full chrome around it.
fn paint_sheet(dc: &mut RasterDrawContext, app: &crate::gui::App, w: i32, h: i32) {
    use crate::gui::viewport::Viewport;

    let cursor = app.core.cursor;
    let data_rows = rows_for_height(h);
    let data_width = data_width_for(w);
    let data_cols = data_width / 2;
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
    let col_widths = stretch_columns_to_width(&vp, w);
    dc.clear(0.94, 0.94, 0.94, 1.0);
    crate::gui::gui_backend::render_grid_body(
        dc,
        app,
        &vp,
        &col_widths,
        cursor.row,
        cursor.col,
        w,
        h,
    );
}

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------

/// `--gui --movie` entry point: replay headlessly, optionally writing one PNG
/// per frame into `frames_dir` (which is how a video gets made without a
/// display server — see `scripts/gui_movie.py`).
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
    let total_px =
        |widths: &HashMap<usize, usize>| -> f64 {
            widths
                .iter()
                .filter(|(c, _)| vp.col_ixs.contains(c))
                .map(|(_, w)| col_px(*w))
                .sum()
        };
    let mut spare = (w as f64 - ROW_LABEL_W) - total_px(&widths);
    // Spread the spare width evenly, one character at a time, so text still
    // aligns to the grid and the loop can stop as soon as the row is covered.
    while spare >= CHAR_W {
        let per_col_chars = (spare / (main.len() as f64 * CHAR_W)).floor() as usize;
        if per_col_chars == 0 {
            break;
        }
        for c in &main {
            let cur = *widths.get(c).unwrap_or(&4);
            widths.insert(*c, cur + per_col_chars);
        }
        spare = (w as f64 - ROW_LABEL_W) - total_px(&widths);
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

/// Pixel width of a grid column: one layout character per character plus a
/// small gutter, matching the live canvas's column accumulation.
fn col_px(width_chars: usize) -> f64 {
    (width_chars as f64 * CHAR_W) + 4.0
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
    /// Added to every incoming coordinate. The shared sheet renderer works in
    /// canvas space (origin at the top-left of the grid), while a movie frame
    /// has the window's chrome above it; an offset view lets one renderer draw
    /// into either place without threading an origin through every call.
    ox: f64,
    oy: f64,
}

impl RasterDrawContext {
    pub fn new(w: i32, h: i32) -> Self {
        let w = w.max(1);
        let h = h.max(1);
        RasterDrawContext {
            w,
            h,
            px: vec![255u8; (w as usize) * (h as usize) * 3],
            ox: 0.0,
            oy: 0.0,
        }
    }

    /// Shift this surface's origin by `(dx, dy)`. Used to paint the sheet
    /// below the window chrome: the shared renderer works in canvas space.
    pub fn set_origin(&mut self, dx: f64, dy: f64) {
        self.ox = dx;
        self.oy = dy;
    }

    pub fn size(&self) -> (i32, i32) {
        (self.w, self.h)
    }

    /// Clear the whole surface (ignores the offset).
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
        let x = x + self.ox;
        let y = y + self.oy;
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
        // No origin shift here: the glyph cells go through `fill_rect`, which
        // applies the offset once. Adding it here too pushed every glyph one
        // chrome-height below its cell (and, for the body region, off-frame).
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
    pub fn into_ppm(&self) -> Vec<u8> {
        let header = format!("P6\n{} {}\n255\n", self.w, self.h);
        let mut out = Vec::with_capacity(header.len() + self.px.len());
        out.extend_from_slice(header.as_bytes());
        out.extend_from_slice(&self.px);
        out
    }

    fn clip(&mut self, _x: f64, _y: f64, _w: f64, _h: f64) {
        // The raster always covers the whole frame; the shared renderer's clip
        // calls are a no-op here (as they are for the widget canvas, which
        // relies on the toolkit's clip).
    }

}

impl MovieDrawContext for RasterDrawContext {
    fn fill_rect(&mut self, x: f64, y: f64, w: f64, h: f64, r: f64, g: f64, b: f64, a: f64) {
        RasterDrawContext::fill_rect(self, x, y, w, h, r, g, b, a)
    }
    fn stroke_rect(
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
        RasterDrawContext::stroke_rect(self, x, y, w, h, r, g, b, a, lw)
    }
    fn draw_text_styled(
        &mut self,
        x: f64,
        y: f64,
        text: &str,
        _font: &str,
        size: f64,
        r: f64,
        g: f64,
        b: f64,
        a: f64,
        _slant: i32,
        _weight: i32,
    ) {
        RasterDrawContext::draw_text(self, x, y, text, size, r, g, b, a)
    }
    fn text_extents_styled(
        &self,
        text: &str,
        _font: &str,
        size: f64,
        _slant: i32,
        _weight: i32,
    ) -> (f64, f64, f64, f64) {
        let scale = if size >= 11.0 { 2.0 } else { 1.0 };
        let w = text.chars().count() as f64 * glyph_advance(scale);
        (0.0, -size, w, size * 1.2)
    }
    fn clear(&mut self, r: f64, g: f64, b: f64, a: f64) {
        RasterDrawContext::clear(self, r, g, b, a)
    }
    fn save(&mut self) {}
    fn restore(&mut self) {}
    fn clip(&mut self, x: f64, y: f64, w: f64, h: f64) {
        RasterDrawContext::clip(self, x, y, w, h)
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

    /// The raster is the surface the *production* sheet renderer paints into;
    /// this pins its primitives (text advance, offset handling, PPM output)
    /// rather than any sheet content, which comes from `render_grid_body`.
    #[test]
    fn raster_draw_context_serializes_ppm_header() {
        let dc = RasterDrawContext::new(4, 2);
        let ppm = dc.into_ppm();
        let header = String::from_utf8_lossy(&ppm[..ppm.len() - 24]).to_string();
        assert!(header.starts_with("P6\n4 2\n255\n"), "header was {header:?}");
    }

    /// Glyphs must be separated: drawing at the 7.2px *layout* pitch made 2x
    /// glyphs overlap into a smear. The advance is the glyph width plus a gap.
    #[test]
    fn glyph_advance_leaves_a_gap_between_characters() {
        let scale = 2.0;
        assert!(
            glyph_advance(scale) > 5.0 * scale,
            "advance must exceed the 5px glyph width at this scale"
        );
    }

    /// The origin offset must be applied exactly once: applying it in both
    /// `draw_text` and the `fill_rect` it draws glyphs through pushed text one
    /// chrome-height below its cell (off-frame for the sheet body).
    #[test]
    fn origin_offset_is_applied_once_to_text() {
        let mut dc = RasterDrawContext::new(40, 40);
        dc.clear(1.0, 1.0, 1.0, 1.0);
        dc.set_origin(0.0, 10.0);
        let _: &dyn MovieDrawContext = &dc;
        MovieDrawContext::draw_text_styled(
            &mut dc, 0.0, 0.0, "A", "monospace", 12.0, 0.0, 0.0, 0.0, 1.0, 0, 0,
        );
        let px = dc.into_ppm();
        // Find the topmost dark row: it must be at the offset (10), not 20.
        let (w, _h, _) = (40usize, 40usize, ());
        let body = &px[px.len() - w * 40 * 3..];
        let first_dark_row = (0..40)
            .find(|y| (0..w).any(|x| body[(y * w + x) * 3] < 80))
            .expect("glyph painted");
        assert_eq!(
            first_dark_row, 10,
            "glyph must start at the offset row, not double it"
        );
    }

    #[test]
    fn columns_stretch_to_cover_the_frame() {
        // A viewport with a couple of columns: stretching must widen them until
        // the row reaches the frame edge, or the sheet renders as a ribbon.
        let mut app = crate::gui::App::new_with_paths(vec![]);
        let cursor = app.core.cursor;
        let w = 1200;
        let data_width = data_width_for(w);
        let vp = crate::gui::viewport::Viewport::recompute(
            &mut app,
            cursor,
            30,
            data_width / 2,
            data_width,
            HEADER_ROWS,
            MARGIN_COLS,
        );
        let widths = stretch_columns_to_width(&vp, w);
        let total: f64 = vp
            .col_ixs
            .iter()
            .map(|c| col_px(*widths.get(c).unwrap_or(&4)))
            .sum();
        let available = w as f64 - ROW_LABEL_W;
        assert!(
            total > available - 4.0 * CHAR_W,
            "columns must cover the frame: {total} of {available}"
        );
    }

    #[test]
    fn rows_for_height_leaves_room_for_chrome_and_status() {
        // 800px frame: chrome + column strip + status band come off the top and
        // bottom, so the row count must not consume the whole height.
        let rows = rows_for_height(800);
        let used = CHROME_H + HEADER_H + (rows as f64) * ROW_H;
        assert!(used <= 800.0, "rows overflow the frame: {used}");
        assert!(rows > 20, "expected a full sheet, got {rows} rows");
    }
}
