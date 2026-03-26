//! Ratatui front-end: sheet viewport, editing, aggregates, move, file sync.

use crate::agg::{cell_display, compute_aggregate};
use crate::grid::{
    addr_logical_col, addr_logical_row, CellAddr, Grid, FOOTER_ROWS, HEADER_ROWS, MARGIN_COLS,
};
use crate::grid::MainRange;
use crate::io::{commit_op, load_full, tail_apply, IoError, LogWatcher};
use crate::ops::{AggFunc, AggregateDef, Op, SheetState};
use crossterm::event::{self, Event, KeyCode, KeyEvent, KeyEventKind, KeyModifiers};
use crossterm::execute;
use crossterm::terminal::{
    disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen,
};
use ratatui::prelude::*;
use ratatui::widgets::{Block, Borders, Paragraph};
use std::io::{self, stdout};
use std::path::{Path, PathBuf};
use thiserror::Error;

/// Width of the row-label gutter (`^A`, ` 1 `, `_B`).
const ROW_LABEL_CHARS: usize = 5;
/// Fixed cell display width in terminal columns.
const CELL_W: usize = 12;
/// Keep at most this many blank lines/cols around the active main data window.
const DISPLAY_EDGE_BLANK: usize = 1;
/// Trailing blank main rows/cols past the last data cell allowed before Down/Right transitions
/// into the footer / right-margin instead of growing the main extent further.
const NAV_BLANK: usize = 1;

#[derive(Debug, Error)]
pub enum RunError {
    #[error("I/O: {0}")]
    Io(#[from] IoError),
    #[error("Terminal: {0}")]
    Term(#[from] io::Error),
}

/// Logical cursor position across header+main+footer rows × total global columns.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct SheetCursor {
    pub row: usize,
    pub col: usize,
}

impl SheetCursor {
    fn clamp(&mut self, grid: &Grid) {
        let rows = HEADER_ROWS + grid.main_rows() + FOOTER_ROWS;
        let cols = grid.total_cols();
        if rows > 0 {
            self.row = self.row.min(rows - 1);
        }
        if cols > 0 {
            self.col = self.col.min(cols - 1);
        }
    }

    fn to_addr(self, grid: &Grid) -> CellAddr {
        let hr = HEADER_ROWS;
        let mr = grid.main_rows();
        let mc = grid.main_cols();
        if self.row < hr {
            CellAddr::Header {
                row: self.row as u8,
                col: self.col as u32,
            }
        } else if self.row < hr + mr {
            let mri = self.row - hr;
            let mcc = self.col;
            if mcc < MARGIN_COLS {
                CellAddr::Left {
                    col: mcc as u8,
                    row: mri as u32,
                }
            } else if mcc < MARGIN_COLS + mc {
                CellAddr::Main {
                    row: mri as u32,
                    col: (mcc - MARGIN_COLS) as u32,
                }
            } else {
                CellAddr::Right {
                    col: (mcc - MARGIN_COLS - mc) as u8,
                    row: mri as u32,
                }
            }
        } else {
            CellAddr::Footer {
                row: (self.row - hr - mr) as u8,
                col: self.col as u32,
            }
        }
    }
}

#[derive(Clone, Debug)]
pub enum Mode {
    Normal,
    Edit { buffer: String },
    OpenPath { buffer: String },
    Help,
    AggPick { idx: usize },
    /// Alt-activated menu bar; letter shortcuts execute actions.
    Menu,
}

// ── Viewport helpers ──────────────────────────────────────────────────────────

fn main_row_window(state: &SheetState, cursor: SheetCursor) -> (usize, usize) {
    let g = &state.grid;
    let hr = HEADER_ROWS;
    let mr = g.main_rows();
    if mr == 0 {
        return (0, 0);
    }

    let mut lo = usize::MAX;
    let mut hi = 0usize;

    for r in 0..mr {
        if g.logical_row_has_content(hr + r) {
            lo = lo.min(r);
            hi = hi.max(r);
        }
    }
    for a in state.aggregates.keys() {
        let lr = addr_logical_row(a, g);
        if lr >= hr && lr < hr + mr {
            let ri = lr - hr;
            lo = lo.min(ri);
            hi = hi.max(ri);
        }
    }
    if cursor.row >= hr && cursor.row < hr + mr {
        let ri = cursor.row - hr;
        lo = lo.min(ri);
        hi = hi.max(ri);
    }
    if lo == usize::MAX {
        lo = 0;
        hi = 0;
    }

    lo = lo.saturating_sub(DISPLAY_EDGE_BLANK);
    hi = hi.saturating_add(DISPLAY_EDGE_BLANK).min(mr.saturating_sub(1));
    (lo, hi)
}

fn main_col_window(state: &SheetState, cursor: SheetCursor) -> (usize, usize) {
    let g = &state.grid;
    let lm = MARGIN_COLS;
    let mc = g.main_cols();
    if mc == 0 {
        return (0, 0);
    }

    let mut lo = usize::MAX;
    let mut hi = 0usize;

    for c in 0..mc {
        if g.logical_col_has_content(lm + c) {
            lo = lo.min(c);
            hi = hi.max(c);
        }
    }
    for a in state.aggregates.keys() {
        let lc = addr_logical_col(a, g);
        if lc >= lm && lc < lm + mc {
            let ci = lc - lm;
            lo = lo.min(ci);
            hi = hi.max(ci);
        }
    }
    if cursor.col >= lm && cursor.col < lm + mc {
        let ci = cursor.col - lm;
        lo = lo.min(ci);
        hi = hi.max(ci);
    }
    if lo == usize::MAX {
        lo = 0;
        hi = 0;
    }

    lo = lo.saturating_sub(DISPLAY_EDGE_BLANK);
    hi = hi.saturating_add(DISPLAY_EDGE_BLANK).min(mc.saturating_sub(1));
    (lo, hi)
}

/// Row viewport with pinned totals and a stable fit-on-screen band.
fn visible_row_indices(state: &SheetState, cursor: SheetCursor, dim: usize) -> Vec<usize> {
    let g = &state.grid;
    let hr = HEADER_ROWS;
    let mr = g.main_rows();
    let fr = FOOTER_ROWS;
    let total = hr + mr + fr;
    let dim = dim.max(1).min(total.max(1));
    let cur = cursor.row.min(total.saturating_sub(1));

    // If everything fits, show all logical rows so section moves don't remap.
    if total <= dim {
        return (0..total).collect();
    }

    // Stable compact band: ^A + focused main rows (+1 blank edge) + _A.
    let (main_lo, main_hi) = main_row_window(state, cursor);
    let main_span = main_hi.saturating_sub(main_lo) + 1;
    let mut stable_band = Vec::with_capacity(main_span + 2);
    if hr > 0 {
        stable_band.push(hr - 1);
    }
    stable_band.extend((main_lo..=main_hi).map(|ri| hr + ri));
    if fr > 0 {
        stable_band.push(hr + mr);
    }
    if stable_band.len() <= dim && stable_band.contains(&cur) {
        return stable_band;
    }

    // Freeze panes: keep _A pinned; keep ^A pinned when room allows.
    let mut reserved: Vec<usize> = Vec::new();
    if fr > 0 {
        reserved.push(hr + mr); // _A totals row (adjacent footer row)
    }
    if hr > 0 && dim > reserved.len() + 1 {
        reserved.push(hr - 1); // ^A orientation row
    }
    reserved.sort_unstable();
    reserved.dedup();

    let available = dim.saturating_sub(reserved.len()).max(1);
    let filtered: Vec<usize> = (0..total)
        .filter(|r| !reserved.iter().any(|p| p == r))
        .collect();
    if filtered.is_empty() {
        return reserved;
    }

    let cur_pos = match filtered.binary_search(&cur) {
        Ok(i) => i,
        Err(i) => i.min(filtered.len().saturating_sub(1)),
    };
    let mut start = cur_pos.saturating_sub(available / 2);
    if start + available > filtered.len() {
        start = filtered.len().saturating_sub(available);
    }
    let end = (start + available).min(filtered.len());

    let mut out = filtered[start..end].to_vec();
    out.extend(reserved);
    out.sort_unstable();
    out.truncate(dim);
    out
}

/// Column viewport with pinned totals and a stable fit-on-screen band.
fn visible_col_indices(state: &SheetState, cursor: SheetCursor, dim: usize) -> Vec<usize> {
    let g = &state.grid;
    let lm = MARGIN_COLS; // left margin width (= 10)
    let mc = g.main_cols();
    let rm = MARGIN_COLS; // right margin width
    let total = lm + mc + rm;
    let dim = dim.max(1).min(total.max(1));
    let cur = cursor.col.min(total.saturating_sub(1));

    // If everything fits, show all columns so section moves don't remap.
    if total <= dim {
        return (0..total).collect();
    }

    // Stable compact band: <0 + focused main cols (+1 blank edge) + >0.
    let (main_lo, main_hi) = main_col_window(state, cursor);
    let main_span = main_hi.saturating_sub(main_lo) + 1;
    let mut stable_band = Vec::with_capacity(main_span + 2);
    if lm > 0 {
        stable_band.push(lm - 1);
    }
    stable_band.extend((main_lo..=main_hi).map(|ci| lm + ci));
    if rm > 0 {
        stable_band.push(lm + mc);
    }
    if stable_band.len() <= dim && stable_band.contains(&cur) {
        return stable_band;
    }

    // Freeze panes: keep >0 pinned; keep <0 pinned when room allows.
    let mut reserved: Vec<usize> = Vec::new();
    if rm > 0 {
        reserved.push(lm + mc); // >0 totals col (adjacent right-margin col)
    }
    if lm > 0 && dim > reserved.len() + 1 {
        reserved.push(lm - 1); // <0 orientation col
    }
    reserved.sort_unstable();
    reserved.dedup();

    let available = dim.saturating_sub(reserved.len()).max(1);
    let filtered: Vec<usize> = (0..total)
        .filter(|c| !reserved.iter().any(|p| p == c))
        .collect();
    if filtered.is_empty() {
        return reserved;
    }

    let cur_pos = match filtered.binary_search(&cur) {
        Ok(i) => i,
        Err(i) => i.min(filtered.len().saturating_sub(1)),
    };
    let mut start = cur_pos.saturating_sub(available / 2);
    if start + available > filtered.len() {
        start = filtered.len().saturating_sub(available);
    }
    let end = (start + available).min(filtered.len());

    let mut out = filtered[start..end].to_vec();
    out.extend(reserved);
    out.sort_unstable();
    out.truncate(dim);
    out
}

// ── Navigation helpers ────────────────────────────────────────────────────────

/// Number of completely blank main rows at the bottom of the current extent
/// (rows after the last row that contains any content).
fn trailing_blank_main_rows(state: &SheetState) -> usize {
    let g = &state.grid;
    let hr = HEADER_ROWS;
    let mr = g.main_rows();
    match (0..mr).rev().find(|&r| g.logical_row_has_content(hr + r)) {
        None => mr,
        Some(last) => mr.saturating_sub(last + 1),
    }
}

/// Number of completely blank main columns at the right of the current extent.
fn trailing_blank_main_cols(state: &SheetState) -> usize {
    let g = &state.grid;
    let lm = MARGIN_COLS;
    let mc = g.main_cols();
    match (0..mc).rev().find(|&c| g.logical_col_has_content(lm + c)) {
        None => mc,
        Some(last) => mc.saturating_sub(last + 1),
    }
}

// ── <0-column footer aggregates ──────────────────────────────────────────────

/// If the `<0` cell of `footer_row_idx` holds a recognised aggregate keyword,
/// return the corresponding function so every main-data cell in that footer row
/// can be rendered as the column-wide aggregate instead of its stored value.
///
/// `<0` is the innermost left-margin column: global col `MARGIN_COLS - 1`.
fn footer_row_agg_func(grid: &Grid, footer_row_idx: usize) -> Option<AggFunc> {
    let key_col = (MARGIN_COLS - 1) as u32;
    let val = grid.get(&CellAddr::Footer {
        row: footer_row_idx as u8,
        col: key_col,
    })?;
    match val.trim().to_uppercase().as_str() {
        "TOTAL" | "SUM"              => Some(AggFunc::Sum),
        "MEAN" | "AVERAGE" | "AVG"  => Some(AggFunc::Mean),
        "MEDIAN"                     => Some(AggFunc::Median),
        "MIN" | "MINIMUM"            => Some(AggFunc::Min),
        "MAX" | "MAXIMUM"            => Some(AggFunc::Max),
        "COUNT"                      => Some(AggFunc::Count),
        _                            => None,
    }
}

/// If the `^A` cell of `global_col` holds a recognised aggregate keyword,
/// return the function so every main-data cell in that right-margin column
/// can be rendered as the row-wide aggregate instead of its stored value.
///
/// `^A` is the last header row (nearest to main), used as a column-header
/// label: `^A,>0 = TOTAL` means "this column shows row totals".
fn right_col_agg_func(grid: &Grid, global_col: usize) -> Option<AggFunc> {
    let val = grid.get(&CellAddr::Header {
        row: (HEADER_ROWS - 1) as u8,
        col: global_col as u32,
    })?;
    match val.trim().to_uppercase().as_str() {
        "TOTAL" | "SUM"              => Some(AggFunc::Sum),
        "MEAN" | "AVERAGE" | "AVG"  => Some(AggFunc::Mean),
        "MEDIAN"                     => Some(AggFunc::Median),
        "MIN" | "MINIMUM"            => Some(AggFunc::Min),
        "MAX" | "MAXIMUM"            => Some(AggFunc::Max),
        "COUNT"                      => Some(AggFunc::Count),
        _                            => None,
    }
}

fn parse_num(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    t.parse::<f64>().ok()
}

fn fold_numbers(func: AggFunc, xs: &[f64]) -> String {
    if xs.is_empty() {
        return String::new();
    }
    match func {
        AggFunc::Sum => format!("{}", xs.iter().sum::<f64>()),
        AggFunc::Mean => format!("{}", xs.iter().sum::<f64>() / xs.len() as f64),
        AggFunc::Median => {
            let mut ys = xs.to_vec();
            ys.sort_by(|a, b| a.partial_cmp(b).unwrap());
            let n = ys.len();
            let m = if n % 2 == 1 {
                ys[n / 2]
            } else {
                (ys[n / 2 - 1] + ys[n / 2]) / 2.0
            };
            format!("{m}")
        }
        AggFunc::Min => xs
            .iter()
            .copied()
            .min_by(|a, b| a.partial_cmp(b).unwrap())
            .map(|v| format!("{v}"))
            .unwrap_or_default(),
        AggFunc::Max => xs
            .iter()
            .copied()
            .max_by(|a, b| a.partial_cmp(b).unwrap())
            .map(|v| format!("{v}"))
            .unwrap_or_default(),
        AggFunc::Count => format!("{}", xs.len()),
    }
}

fn footer_special_col_aggregate(
    grid: &Grid,
    footer_func: AggFunc,
    global_col: usize,
    main_rows: usize,
    main_cols: usize,
) -> Option<String> {
    let right_func = right_col_agg_func(grid, global_col)?;
    let mut samples: Vec<f64> = Vec::new();
    for r in 0..main_rows {
        let row_val = compute_aggregate(
            grid,
            &AggregateDef {
                func: right_func,
                source: MainRange {
                    row_start: r as u32,
                    row_end: r as u32 + 1,
                    col_start: 0,
                    col_end: main_cols as u32,
                },
            },
        );
        if let Some(n) = parse_num(&row_val) {
            samples.push(n);
        }
    }
    Some(fold_numbers(footer_func, &samples))
}

// ── App ───────────────────────────────────────────────────────────────────────

pub struct App {
    pub path: Option<PathBuf>,
    pub offset: u64,
    pub state: SheetState,
    pub cursor: SheetCursor,
    pub anchor: Option<SheetCursor>,
    pub mode: Mode,
    pub watcher: Option<LogWatcher>,
    pub status: String,
    pub ops_applied: usize,
}

impl App {
    pub fn new(path: Option<PathBuf>) -> Self {
        App {
            path,
            offset: 0,
            state: SheetState::new(1, 1),
            cursor: SheetCursor {
                row: HEADER_ROWS,
                col: MARGIN_COLS,
            },
            anchor: None,
            mode: Mode::Normal,
            watcher: None,
            status: String::new(),
            ops_applied: 0,
        }
    }

    pub fn load_initial(&mut self) -> Result<(), IoError> {
        if let Some(ref p) = self.path.clone() {
            if Path::new(p).exists() {
                let (off, n) = load_full(p, &mut self.state)?;
                self.offset = off;
                self.ops_applied = n;
                self.watcher = Some(LogWatcher::new(p.clone())?);
                self.status = format!("Loaded {}", p.display());
            } else {
                self.watcher = Some(LogWatcher::new(p.clone())?);
                self.status = format!("New file {}", p.display());
            }
        } else {
            self.status = "No file — press o to set path".into();
        }
        self.cursor.clamp(&self.state.grid);
        Ok(())
    }

    fn sync_external(&mut self) -> Result<(), IoError> {
        if let Some(w) = &self.watcher {
            if w.poll_dirty() {
                if let Some(ref p) = self.path {
                    match tail_apply(p, self.offset, &mut self.state) {
                        Ok(new_off) => {
                            self.offset = new_off;
                            self.status = "External change applied".into();
                        }
                        Err(_) => {
                            self.state = SheetState::new(1, 1);
                            let (off, n) = load_full(p, &mut self.state)?;
                            self.offset = off;
                            self.ops_applied = n;
                            self.status = "File reset; full reload".into();
                        }
                    }
                }
            }
        }
        Ok(())
    }

    fn default_main_range(&self) -> MainRange {
        MainRange {
            row_start: 0,
            row_end: self.state.grid.main_rows() as u32,
            col_start: 0,
            col_end: self.state.grid.main_cols() as u32,
        }
    }

    fn selection_main_row_range(&self) -> Option<(u32, u32)> {
        let a = self.anchor?;
        let b = self.cursor;
        let hr = HEADER_ROWS;
        let r0 = a.row.min(b.row);
        let r1 = a.row.max(b.row);
        let c0 = a.col.min(b.col);
        let c1 = a.col.max(b.col);
        let left = MARGIN_COLS;
        let right = MARGIN_COLS + self.state.grid.main_cols();
        if r0 < hr || r1 >= hr + self.state.grid.main_rows() {
            return None;
        }
        if c0 != left || c1 != right.saturating_sub(1) {
            return None;
        }
        Some(((r0 - hr) as u32, (r1 - hr) as u32))
    }

    fn selection_main_col_range(&self) -> Option<(u32, u32)> {
        let a = self.anchor?;
        let b = self.cursor;
        let hr = HEADER_ROWS;
        let r0 = a.row.min(b.row);
        let r1 = a.row.max(b.row);
        let c0 = a.col.min(b.col);
        let c1 = a.col.max(b.col);
        let left = MARGIN_COLS;
        let right = MARGIN_COLS + self.state.grid.main_cols();
        if c0 < left || c1 >= right {
            return None;
        }
        let last_main = hr + self.state.grid.main_rows().saturating_sub(1);
        if r0 != hr || r1 != last_main {
            return None;
        }
        Some(((c0 - left) as u32, (c1 - left) as u32))
    }

    pub fn run(&mut self) -> Result<(), RunError> {
        enable_raw_mode()?;
        let mut stdout = stdout();
        execute!(stdout, EnterAlternateScreen)?;
        let backend = CrosstermBackend::new(stdout);
        let mut terminal = Terminal::new(backend)?;

        loop {
            self.sync_external()?;
            terminal.draw(|f| self.draw(f))?;

            if !event::poll(std::time::Duration::from_millis(200))? {
                continue;
            }
            if let Event::Key(key) = event::read()? {
                if self.handle_key(key)? {
                    break;
                }
            }
        }

        disable_raw_mode()?;
        execute!(terminal.backend_mut(), LeaveAlternateScreen)?;
        Ok(())
    }

    fn draw(&mut self, f: &mut Frame) {
        // Layout: formula bar (1) | grid (fills terminal) | hints bar (1)
        let layout = Layout::default()
            .direction(Direction::Vertical)
            .constraints([
                Constraint::Length(1),
                Constraint::Min(3),
                Constraint::Length(1),
            ])
            .split(f.area());
        let formula_area = layout[0];
        let grid_area = layout[1];
        let hints_area = layout[2];

        // Compute inner area first (block borders don't depend on content).
        let sentinel = Block::default().borders(Borders::ALL);
        let inner = sentinel.inner(grid_area);
        let inner_h = inner.height as usize;
        let inner_w = inner.width as usize;

        // Rows available for data (minus column-header line).
        let data_rows = inner_h.saturating_sub(1).max(1);
        // Columns that fit after the row-label gutter.
        let data_cols = inner_w
            .saturating_sub(ROW_LABEL_CHARS)
            .checked_div(CELL_W)
            .unwrap_or(1)
            .max(1);

        // Avoid aggressive auto-growth in draw(); growth happens via navigation
        // and edits so small sheets can stay fully visible and stable on screen.

        // ── Block title ───────────────────────────────────────────────────────
        let grid = &self.state.grid;
        let title_str = {
            let raw = format!(
                " corro  {}r × {}c  ops {}",
                grid.main_rows(),
                grid.main_cols(),
                self.ops_applied
            );
            let max_w = (grid_area.width.saturating_sub(4) as usize).max(8);
            if raw.chars().count() > max_w {
                format!(
                    "{}…",
                    raw.chars().take(max_w.saturating_sub(1)).collect::<String>()
                )
            } else {
                raw
            }
        };

        let block = Block::default()
            .borders(Borders::ALL)
            .title(Span::styled(
                title_str,
                Style::default().add_modifier(Modifier::BOLD),
            ));

        // ── Viewport ──────────────────────────────────────────────────────────
        let row_ixs = visible_row_indices(&self.state, self.cursor, data_rows);
        let col_ixs = visible_col_indices(&self.state, self.cursor, data_cols);

        // ── Formula bar ───────────────────────────────────────────────────────
        let addr = self.cursor.to_addr(grid);
        let addr_str = addr_label(&addr, grid.main_cols());
        let (formula_text, formula_style) = match &self.mode {
            Mode::Edit { buffer } => (
                format!(" {addr_str}  {buffer}_"),
                Style::default()
                    .fg(Color::White)
                    .bg(Color::DarkGray)
                    .add_modifier(Modifier::BOLD),
            ),
            Mode::OpenPath { buffer } => (
                format!(" open: {buffer}_"),
                Style::default().fg(Color::White).bg(Color::DarkGray),
            ),
            Mode::Help => (
                " ^A/_A and <0/>0 stay visible  |  e·edit  v·select  a·agg  r/c·move  o·open  q·quit".into(),
                Style::default().fg(Color::White).bg(Color::Blue),
            ),
            Mode::Menu => (
                "  F·open-file   R·row-ops   C·col-ops   A·aggregate   ?·help   Esc·close".into(),
                Style::default().fg(Color::Black).bg(Color::Cyan),
            ),
            _ => {
                let val = cell_display(grid, &self.state.aggregates, &addr);
                let base = format!(" {addr_str}  {val}");
                let text = if self.status.is_empty() {
                    base
                } else {
                    format!("{base}   ·  {}", self.status)
                };
                (text, Style::default().fg(Color::Cyan))
            }
        };
        f.render_widget(
            Paragraph::new(formula_text).style(formula_style),
            formula_area,
        );

        // ── Grid ──────────────────────────────────────────────────────────────
        let mut lines: Vec<Line> = Vec::new();

        // Column header row
        {
            let mut spans: Vec<Span> = vec![Span::styled(
                format!("{:>width$}", "", width = ROW_LABEL_CHARS),
                Style::default().add_modifier(Modifier::BOLD),
            )];
            for &c in &col_ixs {
                let name = col_header_label(c, grid.main_cols());
                let active_col = c == self.cursor.col;
                let style = if active_col {
                    Style::default()
                        .fg(Color::Black)
                        .bg(Color::Yellow)
                        .add_modifier(Modifier::BOLD)
                } else {
                    Style::default().fg(Color::Cyan).add_modifier(Modifier::BOLD)
                };
                spans.push(Span::styled(
                    format!("{:>w$}", name, w = CELL_W),
                    style,
                ));
            }
            lines.push(Line::from(spans));
        }

        let hr = HEADER_ROWS;
        let mr = grid.main_rows();
        let lm = MARGIN_COLS;
        let mc = grid.main_cols();
        let max_data_lines = inner_h.saturating_sub(1);
        for &r in row_ixs.iter().take(max_data_lines) {
            let active_row = r == self.cursor.row;
            let row_label_style = if active_row {
                Style::default()
                    .fg(Color::Black)
                    .bg(Color::Yellow)
                    .add_modifier(Modifier::BOLD)
            } else {
                Style::default().fg(Color::Yellow)
            };
            let mut spans: Vec<Span> = vec![Span::styled(
                format!("{:>4} ", sheet_row_label(r, grid.main_rows())),
                row_label_style,
            )];
            // If this is a footer row whose <0 cell holds an aggregate keyword,
            // pre-compute the function so every main-data cell in the row can
            // display the column-wide aggregate instead of its stored value.
            let footer_agg = if r >= hr + mr {
                footer_row_agg_func(grid, r - hr - mr)
            } else {
                None
            };
            for &c in &col_ixs {
                let cur = SheetCursor { row: r, col: c };
                let cell_addr = cur.to_addr(grid);
                let text = if let Some(func) = footer_agg {
                    // <0-driven: footer row → show column aggregate for main cols.
                    if c >= lm && c < lm + mc {
                        let main_col = (c - lm) as u32;
                        compute_aggregate(grid, &AggregateDef {
                            func,
                            source: MainRange {
                                row_start: 0,
                                row_end: mr as u32,
                                col_start: main_col,
                                col_end: main_col + 1,
                            },
                        })
                    } else if c >= lm + mc {
                        footer_special_col_aggregate(grid, func, c, mr, mc)
                            .unwrap_or_else(|| cell_display(grid, &self.state.aggregates, &cell_addr))
                    } else {
                        cell_display(grid, &self.state.aggregates, &cell_addr)
                    }
                } else if r >= hr && r < hr + mr && c >= lm + mc {
                    // _A-driven: right-margin col → show row aggregate for main rows.
                    if let Some(func) = right_col_agg_func(grid, c) {
                        let main_row = (r - hr) as u32;
                        compute_aggregate(grid, &AggregateDef {
                            func,
                            source: MainRange {
                                row_start: main_row,
                                row_end: main_row + 1,
                                col_start: 0,
                                col_end: mc as u32,
                            },
                        })
                    } else {
                        cell_display(grid, &self.state.aggregates, &cell_addr)
                    }
                } else {
                    cell_display(grid, &self.state.aggregates, &cell_addr)
                };
                let disp = if text.chars().count() > CELL_W {
                    format!("{}…", text.chars().take(CELL_W).collect::<String>())
                } else {
                    format!("{:w$}", text, w = CELL_W)
                };
                let sel = self.anchor.is_some_and(|a| {
                    let r0 = a.row.min(self.cursor.row);
                    let r1 = a.row.max(self.cursor.row);
                    let c0 = a.col.min(self.cursor.col);
                    let c1 = a.col.max(self.cursor.col);
                    r >= r0 && r <= r1 && c >= c0 && c <= c1
                });
                let is_cur = r == self.cursor.row && c == self.cursor.col;
                let st = if is_cur {
                    Style::default().bg(Color::DarkGray)
                } else if sel {
                    Style::default().bg(Color::Blue)
                } else {
                    Style::default()
                };
                spans.push(Span::styled(disp, st));
            }
            lines.push(Line::from(spans));
        }

        let n = lines.len().min(inner_h);
        if n > 0 {
            let mut constraints: Vec<Constraint> =
                (0..n).map(|_| Constraint::Length(1)).collect();
            if inner.height > n as u16 {
                constraints.push(Constraint::Min(0));
            }
            let row_areas = Layout::default()
                .direction(Direction::Vertical)
                .constraints(constraints)
                .split(inner);
            for i in 0..n {
                f.render_widget(
                    Paragraph::new(lines[i].clone()).left_aligned(),
                    row_areas[i],
                );
            }
        }

        f.render_widget(block, grid_area);

        // ── Context-sensitive hints bar ───────────────────────────────────────
        let hints = self.hints_line();
        f.render_widget(
            Paragraph::new(hints).style(Style::default().fg(Color::DarkGray)),
            hints_area,
        );
    }

    fn hints_line(&self) -> String {
        match &self.mode {
            Mode::Normal => {
                if self.anchor.is_some() {
                    "  r·move-rows   c·move-cols   a·agg   v·deselect   Esc·cancel".into()
                } else {
                    "  e·edit   v·select   a·agg   o·open   hjkl/arrows·nav   totals pinned   q·quit   Alt·menu   ?·help".into()
                }
            }
            Mode::Edit { .. } => "  type to edit   Enter·confirm   Esc·discard".into(),
            Mode::OpenPath { .. } => "  type file path   Enter·open   Esc·cancel".into(),
            Mode::Help => "  Esc / q / ?  ·  close help".into(),
            Mode::AggPick { .. } => {
                "  1·SUM  2·MEAN  3·MEDIAN  4·MIN  5·MAX  6·COUNT   Enter·set   Esc·cancel".into()
            }
            Mode::Menu => "  press a letter shortcut   Esc·close".into(),
        }
    }

    fn handle_key(&mut self, key: KeyEvent) -> Result<bool, RunError> {
        if key.kind == KeyEventKind::Release {
            return Ok(false);
        }

        // Alt: toggle menu or execute a shortcut directly from Normal/Menu.
        if key.modifiers.contains(KeyModifiers::ALT) {
            if matches!(self.mode, Mode::Normal | Mode::Menu) {
                match key.code {
                    KeyCode::Char('f') | KeyCode::Char('F') => {
                        self.mode = Mode::OpenPath {
                            buffer: self
                                .path
                                .as_ref()
                                .map(|p| p.to_string_lossy().into_owned())
                                .unwrap_or_default(),
                        };
                    }
                    KeyCode::Char('a') | KeyCode::Char('A') => {
                        self.mode = Mode::AggPick { idx: 0 };
                    }
                    KeyCode::Char('?') | KeyCode::Char('h') | KeyCode::Char('H') => {
                        self.mode = Mode::Help;
                    }
                    _ => {
                        self.mode = if matches!(self.mode, Mode::Menu) {
                            Mode::Normal
                        } else {
                            Mode::Menu
                        };
                    }
                }
            }
            return Ok(false);
        }

        match &mut self.mode {
            Mode::Menu => {
                match key.code {
                    KeyCode::Esc => self.mode = Mode::Normal,
                    KeyCode::Char('f') | KeyCode::Char('F') => {
                        self.mode = Mode::OpenPath {
                            buffer: self
                                .path
                                .as_ref()
                                .map(|p| p.to_string_lossy().into_owned())
                                .unwrap_or_default(),
                        };
                    }
                    KeyCode::Char('r') | KeyCode::Char('R') => {
                        self.mode = Mode::Normal;
                        self.status =
                            "Row ops: v·select full rows, then r·move to target row".into();
                    }
                    KeyCode::Char('c') | KeyCode::Char('C') => {
                        self.mode = Mode::Normal;
                        self.status =
                            "Col ops: v·select full columns, then c·move to target column".into();
                    }
                    KeyCode::Char('a') | KeyCode::Char('A') => {
                        self.mode = Mode::AggPick { idx: 0 };
                    }
                    KeyCode::Char('?') => self.mode = Mode::Help,
                    _ => {}
                }
                return Ok(false);
            }
            Mode::Help => {
                if matches!(
                    key.code,
                    KeyCode::Esc | KeyCode::Char('q') | KeyCode::Char('?')
                ) {
                    self.mode = Mode::Normal;
                }
                return Ok(false);
            }
            Mode::AggPick { idx } => {
                match key.code {
                    KeyCode::Esc => self.mode = Mode::Normal,
                    KeyCode::Char(c) if c.is_ascii_digit() => {
                        let d = c.to_digit(10).unwrap_or(1).saturating_sub(1).min(5);
                        *idx = d as usize;
                    }
                    KeyCode::Enter => {
                        let funcs = [
                            AggFunc::Sum,
                            AggFunc::Mean,
                            AggFunc::Median,
                            AggFunc::Min,
                            AggFunc::Max,
                            AggFunc::Count,
                        ];
                        let f = funcs[*idx % funcs.len()];
                        let addr = self.cursor.to_addr(&self.state.grid);
                        let source = if let Some(a) = self.anchor {
                            let r0 = a.row.min(self.cursor.row);
                            let r1 = a.row.max(self.cursor.row);
                            let c0 = a.col.min(self.cursor.col);
                            let c1 = a.col.max(self.cursor.col);
                            let hr = HEADER_ROWS;
                            if r0 >= hr
                                && r1 < hr + self.state.grid.main_rows()
                                && c0 >= MARGIN_COLS
                                && c1 < MARGIN_COLS + self.state.grid.main_cols()
                            {
                                MainRange {
                                    row_start: (r0 - hr) as u32,
                                    row_end: (r1 - hr) as u32 + 1,
                                    col_start: (c0 - MARGIN_COLS) as u32,
                                    col_end: (c1 - MARGIN_COLS) as u32 + 1,
                                }
                            } else {
                                self.default_main_range()
                            }
                        } else {
                            self.default_main_range()
                        };
                        let op = Op::SetAggregate {
                            addr,
                            def: AggregateDef { func: f, source },
                        };
                        if let Some(ref p) = self.path.clone() {
                            commit_op(p, &mut self.offset, &mut self.state, &op)?;
                            self.ops_applied = self.ops_applied.saturating_add(1);
                        } else {
                            op.apply(&mut self.state);
                            self.status = "No file — aggregate in memory only".into();
                        }
                        self.mode = Mode::Normal;
                    }
                    _ => {}
                }
                return Ok(false);
            }
            Mode::OpenPath { buffer } => {
                match key.code {
                    KeyCode::Enter => {
                        let path = PathBuf::from(buffer.trim());
                        self.path = Some(path.clone());
                        self.offset = 0;
                        self.state = SheetState::new(1, 1);
                        if path.exists() {
                            let (off, n) = load_full(&path, &mut self.state)?;
                            self.offset = off;
                            self.ops_applied = n;
                        }
                        self.watcher =
                            Some(LogWatcher::new(path.clone()).map_err(IoError::from)?);
                        self.cursor = SheetCursor {
                            row: HEADER_ROWS,
                            col: MARGIN_COLS,
                        };
                        self.mode = Mode::Normal;
                        self.status = format!("Opened {}", path.display());
                    }
                    KeyCode::Esc => self.mode = Mode::Normal,
                    KeyCode::Char(c) => buffer.push(c),
                    KeyCode::Backspace => {
                        buffer.pop();
                    }
                    _ => {}
                }
                return Ok(false);
            }
            Mode::Edit { buffer } => {
                match key.code {
                    KeyCode::Enter => {
                        let addr = self.cursor.to_addr(&self.state.grid);
                        let op = Op::SetCell {
                            addr,
                            value: buffer.clone(),
                        };
                        if let Some(ref p) = self.path.clone() {
                            commit_op(p, &mut self.offset, &mut self.state, &op)?;
                            self.ops_applied = self.ops_applied.saturating_add(1);
                        } else {
                            op.apply(&mut self.state);
                            self.status = "No file — edit in memory only".into();
                        }
                        self.mode = Mode::Normal;
                    }
                    KeyCode::Esc => self.mode = Mode::Normal,
                    KeyCode::Char(c) => buffer.push(c),
                    KeyCode::Backspace => {
                        buffer.pop();
                    }
                    _ => {}
                }
                return Ok(false);
            }
            Mode::Normal => {}
        }

        if key.modifiers.contains(KeyModifiers::CONTROL) && key.code == KeyCode::Char('q') {
            return Ok(true);
        }

        match key.code {
            KeyCode::Esc => {
                self.anchor = None;
            }
            KeyCode::Char('?') => self.mode = Mode::Help,
            KeyCode::Char('o') => {
                self.mode = Mode::OpenPath {
                    buffer: self
                        .path
                        .as_ref()
                        .map(|p| p.to_string_lossy().into_owned())
                        .unwrap_or_default(),
                };
            }
            KeyCode::Char('e') | KeyCode::Enter => {
                let addr = self.cursor.to_addr(&self.state.grid);
                let cur = cell_display(&self.state.grid, &self.state.aggregates, &addr);
                self.mode = Mode::Edit { buffer: cur };
            }
            KeyCode::Char('v') => {
                self.anchor = if self.anchor.is_none() {
                    Some(self.cursor)
                } else {
                    None
                };
            }
            KeyCode::Char('a') => self.mode = Mode::AggPick { idx: 0 },
            KeyCode::Char('r') => {
                if let Some((mr0, mr1)) = self.selection_main_row_range() {
                    let hr = HEADER_ROWS;
                    if self.cursor.row < hr
                        || self.cursor.row >= hr + self.state.grid.main_rows()
                    {
                        self.status =
                            "Place cursor on a main row as move target, then press r".into();
                        return Ok(false);
                    }
                    let count = mr1 - mr0 + 1;
                    let to = (self.cursor.row - hr) as u32;
                    let op = Op::MoveRowRange {
                        from: mr0,
                        count,
                        to,
                    };
                    if let Some(ref p) = self.path.clone() {
                        commit_op(p, &mut self.offset, &mut self.state, &op)?;
                        self.ops_applied = self.ops_applied.saturating_add(1);
                    } else {
                        op.apply(&mut self.state);
                    }
                    self.anchor = None;
                    self.status = format!("Moved rows {mr0}..{} → before row {to}", mr0 + count);
                } else {
                    self.status = "Select full main rows first (v), then r to move".into();
                }
            }
            KeyCode::Char('c') => {
                if let Some((mc0, mc1)) = self.selection_main_col_range() {
                    let left = MARGIN_COLS;
                    let right = MARGIN_COLS + self.state.grid.main_cols();
                    if self.cursor.col < left || self.cursor.col >= right {
                        self.status =
                            "Place cursor on a main column as move target, then press c".into();
                        return Ok(false);
                    }
                    let count = mc1 - mc0 + 1;
                    let to = (self.cursor.col - left) as u32;
                    let op = Op::MoveColRange {
                        from: mc0,
                        count,
                        to,
                    };
                    if let Some(ref p) = self.path.clone() {
                        commit_op(p, &mut self.offset, &mut self.state, &op)?;
                        self.ops_applied = self.ops_applied.saturating_add(1);
                    } else {
                        op.apply(&mut self.state);
                    }
                    self.anchor = None;
                    self.status = format!("Moved cols {mc0}..{} → before col {to}", mc0 + count);
                } else {
                    self.status = "Select full main columns first (v), then c to move".into();
                }
            }
            KeyCode::Left | KeyCode::Char('h') => {
                self.cursor.col = self.cursor.col.saturating_sub(1);
                self.cursor.clamp(&self.state.grid);
                self.state
                    .grid
                    .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
            }
            KeyCode::Right | KeyCode::Char('l') => {
                // Grow main only when stepping off the last main column *and* there are
                // fewer than NAV_BLANK trailing blank main columns.  Once there are enough
                // blank columns the cursor naturally enters the right margin (>0 …).
                let lm = MARGIN_COLS;
                let mc = self.state.grid.main_cols();
                if self.cursor.col == lm + mc.saturating_sub(1)
                    && trailing_blank_main_cols(&self.state) < NAV_BLANK
                {
                    self.state.grid.grow_main_col_at_right();
                }
                self.cursor.col = self.cursor.col.saturating_add(1);
                self.cursor.clamp(&self.state.grid);
                self.state
                    .grid
                    .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
            }
            KeyCode::Up | KeyCode::Char('k') => {
                self.cursor.row = self.cursor.row.saturating_sub(1);
                self.cursor.clamp(&self.state.grid);
                self.state
                    .grid
                    .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
            }
            KeyCode::Down | KeyCode::Char('j') => {
                // Grow main only when stepping off the last main row *and* there are
                // fewer than NAV_BLANK trailing blank main rows.  Once there are enough
                // blank rows the cursor naturally falls into the footer section (_Z …).
                let hr = HEADER_ROWS;
                let last_main = hr + self.state.grid.main_rows().saturating_sub(1);
                if self.cursor.row == last_main
                    && trailing_blank_main_rows(&self.state) < NAV_BLANK
                {
                    self.state.grid.grow_main_row_at_bottom();
                }
                self.cursor.row = self.cursor.row.saturating_add(1);
                self.cursor.clamp(&self.state.grid);
                self.state
                    .grid
                    .ensure_extent_for_cursor(self.cursor.row, self.cursor.col);
            }
            KeyCode::Char('q') => return Ok(true),
            _ => {}
        }

        Ok(false)
    }
}

// ── Display helpers ───────────────────────────────────────────────────────────

/// Compact address label for the formula bar (e.g. `A1`, `^Z:A`, `_A:<0`).
fn addr_label(addr: &CellAddr, main_cols: usize) -> String {
    match addr {
        CellAddr::Header { row, col } => format!(
            "^{}:{}",
            (b'Z' - *row) as char,
            col_header_label(*col as usize, main_cols)
        ),
        CellAddr::Footer { row, col } => format!(
            "_{}:{}",
            (b'A' + *row) as char,
            col_header_label(*col as usize, main_cols)
        ),
        CellAddr::Main { row, col } => {
            format!("{}{}", excel_column_name(*col as usize), row + 1)
        }
        CellAddr::Left { col, row } => format!("<{}:{}", MARGIN_COLS - 1 - (*col as usize), row + 1),
        CellAddr::Right { col, row } => format!(">{}:{}", col, row + 1),
    }
}

/// Left gutter label: `^Z`–`^A` (header), `1`–`n` (main), `_Z`–`_A` (footer).
fn sheet_row_label(logical_row: usize, main_rows: usize) -> String {
    let hr = HEADER_ROWS;
    if logical_row < hr {
        format!("^{}", (b'Z' - logical_row as u8) as char)
    } else if logical_row < hr + main_rows {
        format!("{}", logical_row - hr + 1)
    } else {
        let fr = logical_row - hr - main_rows;
        format!("_{}", (b'A' + fr as u8) as char)
    }
}

/// Top header label: `<9`–`<0` (left margin, outermost→innermost), Excel letters (main),
/// `>0`–`>9` (right margin, innermost→outermost).
fn col_header_label(global_col: usize, main_cols: usize) -> String {
    let m = MARGIN_COLS;
    if global_col < m {
        format!("<{}", m - 1 - global_col)
    } else if global_col < m + main_cols {
        excel_column_name(global_col - m)
    } else {
        format!(">{}", global_col - m - main_cols)
    }
}

/// Excel-style column name: 0 → `A`, 25 → `Z`, 26 → `AA`, …
fn excel_column_name(main_col_index: usize) -> String {
    let mut n = main_col_index + 1;
    let mut s = String::new();
    while n > 0 {
        n -= 1;
        s.push((b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    s.chars().rev().collect()
}
