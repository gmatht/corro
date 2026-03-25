//! Ratatui front-end: sheet viewport, editing, aggregates, move, file sync.

use crate::agg::cell_display;
use crate::grid::{CellAddr, Grid, FOOTER_ROWS, HEADER_ROWS, MARGIN_COLS};
use crate::io::{commit_op, load_full, tail_apply, IoError, LogWatcher};
use crate::ops::{AggFunc, AggregateDef, Op, SheetState};
use crate::grid::MainRange;
use crossterm::event::{self, Event, KeyCode, KeyEvent, KeyModifiers};
use crossterm::execute;
use crossterm::terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen};
use ratatui::prelude::*;
use ratatui::widgets::{Block, Borders, Paragraph};
use std::io::{self, stdout};
use std::path::{Path, PathBuf};
use thiserror::Error;

/// Width of the row-label column (`^A`, ` 1 `, `_B`).
const ROW_LABEL_CHARS: usize = 5;
/// Fixed cell width so rows stay one terminal line (no Paragraph wrap).
const CELL_W: usize = 12;

#[derive(Debug, Error)]
pub enum RunError {
    #[error("I/O: {0}")]
    Io(#[from] IoError),
    #[error("Terminal: {0}")]
    Term(#[from] io::Error),
}

/// Logical sheet cursor: row across header + main + footer, col global width.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct SheetCursor {
    pub row: usize,
    pub col: usize,
}

impl SheetCursor {
    fn clamp(&mut self, grid: &Grid) {
        let rows = HEADER_ROWS + grid.main_rows + FOOTER_ROWS;
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
        if self.row < hr {
            CellAddr::Header {
                row: self.row as u8,
                col: self.col as u32,
            }
        } else if self.row < hr + grid.main_rows {
            let mr = self.row - hr;
            let mc = self.col;
            if mc < MARGIN_COLS {
                CellAddr::Left {
                    col: mc as u8,
                    row: mr as u32,
                }
            } else if mc < MARGIN_COLS + grid.main_cols {
                CellAddr::Main {
                    row: mr as u32,
                    col: (mc - MARGIN_COLS) as u32,
                }
            } else {
                CellAddr::Right {
                    col: (mc - MARGIN_COLS - grid.main_cols) as u8,
                    row: mr as u32,
                }
            }
        } else {
            CellAddr::Footer {
                row: (self.row - hr - grid.main_rows) as u8,
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
}

pub struct App {
    pub path: Option<PathBuf>,
    pub offset: u64,
    pub state: SheetState,
    pub cursor: SheetCursor,
    pub anchor: Option<SheetCursor>,
    pub mode: Mode,
    pub watcher: Option<LogWatcher>,
    pub scroll_row: usize,
    pub scroll_col: usize,
    pub status: String,
    pub ops_applied: usize,
}

impl App {
    pub fn new(path: Option<PathBuf>) -> Self {
        App {
            path,
            offset: 0,
            state: SheetState::new(4, 6),
            cursor: SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS },
            anchor: None,
            mode: Mode::Normal,
            watcher: None,
            scroll_row: 0,
            scroll_col: 0,
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
                self.status = format!("Loaded {} (append-only; last writer wins per cell)", p.display());
            } else {
                self.watcher = Some(LogWatcher::new(p.clone())?);
                self.status = format!("New file {}; changes append on write", p.display());
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
            row_end: self.state.grid.main_rows as u32,
            col_start: 0,
            col_end: self.state.grid.main_cols as u32,
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
        let right = MARGIN_COLS + self.state.grid.main_cols;
        if r0 < hr || r1 >= hr + self.state.grid.main_rows {
            return None;
        }
        if c0 != left || c1 != right.saturating_sub(1) {
            return None;
        }
        let mr0 = (r0 - hr) as u32;
        let mr1 = (r1 - hr) as u32;
        Some((mr0, mr1))
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
        let right = MARGIN_COLS + self.state.grid.main_cols;
        if c0 < left || c1 >= right {
            return None;
        }
        let last_main = hr + self.state.grid.main_rows.saturating_sub(1);
        if r0 != hr || r1 != last_main {
            return None;
        }
        let mc0 = (c0 - left) as u32;
        let mc1 = (c1 - left) as u32;
        Some((mc0, mc1))
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
        let main = Layout::default()
            .direction(Direction::Vertical)
            .constraints([Constraint::Min(3), Constraint::Length(2)])
            .split(f.area());

        let grid_area = main[0];
        let status_area = main[1];

        let grid = &self.state.grid;
        let total_rows = HEADER_ROWS + grid.main_rows + FOOTER_ROWS;
        let total_cols = grid.total_cols();

        // Inner height: borders eat 2 rows; one line is the column-header row.
        let inner_h = grid_area.height.saturating_sub(2) as usize;
        let data_view_rows = inner_h.saturating_sub(1);
        // Approximate how many cell columns fit (row label + N × cell width).
        let view_cols = (grid_area.width as usize)
            .saturating_sub(2)
            .saturating_sub(ROW_LABEL_CHARS)
            / CELL_W.max(1);

        if self.cursor.row < self.scroll_row {
            self.scroll_row = self.cursor.row;
        }
        if data_view_rows > 0 && self.cursor.row >= self.scroll_row + data_view_rows {
            self.scroll_row = self
                .cursor
                .row
                .saturating_sub(data_view_rows.saturating_sub(1));
        }
        if self.cursor.col < self.scroll_col {
            self.scroll_col = self.cursor.col;
        }
        if view_cols > 0 && self.cursor.col >= self.scroll_col + view_cols {
            self.scroll_col = self
                .cursor
                .col
                .saturating_sub(view_cols.saturating_sub(1));
        }

        let mut lines: Vec<Line> = Vec::new();

        // Column names row (scrolls horizontally with data).
        {
            let mut spans: Vec<Span> = Vec::new();
            spans.push(Span::styled(
                format!("{:>width$}", "", width = ROW_LABEL_CHARS),
                Style::default().add_modifier(Modifier::BOLD),
            ));
            for vc in 0..view_cols.min(total_cols.saturating_sub(self.scroll_col)) {
                let c = self.scroll_col + vc;
                let name = col_label(c, grid.main_cols);
                let disp = format!("{:>w$}", name, w = CELL_W);
                spans.push(Span::styled(
                    disp,
                    Style::default()
                        .fg(Color::Cyan)
                        .add_modifier(Modifier::BOLD),
                ));
            }
            lines.push(Line::from(spans));
        }

        for vr in 0..data_view_rows.min(total_rows.saturating_sub(self.scroll_row)) {
            let r = self.scroll_row + vr;
            let mut spans: Vec<Span> = Vec::new();
            let label = row_label(r, grid.main_rows);
            spans.push(Span::styled(
                format!("{label:>4} "),
                Style::default().fg(Color::Yellow),
            ));

            for vc in 0..view_cols.min(total_cols.saturating_sub(self.scroll_col)) {
                let c = self.scroll_col + vc;
                let cur = SheetCursor { row: r, col: c };
                let addr = cur.to_addr(grid);
                let text = cell_display(grid, &self.state.aggregates, &addr);
                let disp = if text.len() > CELL_W {
                    format!("{}…", &text[..CELL_W])
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
                let mut st = Style::default();
                if is_cur {
                    st = st.bg(Color::DarkGray);
                } else if sel {
                    st = st.bg(Color::Blue);
                }
                spans.push(Span::styled(disp, st));
            }
            lines.push(Line::from(spans));
        }

        let block = Paragraph::new(lines).block(
            Block::default()
                .borders(Borders::ALL)
                .title(Line::from(vec![
                    Span::styled(" sheet ", Style::default().add_modifier(Modifier::BOLD)),
                    Span::styled(
                        format!(" {total_rows}×{total_cols} ops {} ", self.ops_applied),
                        Style::default().fg(Color::DarkGray),
                    ),
                ])),
        );
        f.render_widget(block, grid_area);

        let mode_line = match &self.mode {
            Mode::Normal => format!("NORMAL | {}", self.status),
            Mode::Edit { buffer } => format!("EDIT | {buffer}_"),
            Mode::OpenPath { buffer } => format!("OPEN | {buffer}_"),
            Mode::Help => "HELP | q/? exit help | arrows move | e edit | v visual | a aggregate | r row move | c col move | o open".into(),
            Mode::AggPick { idx } => format!("AGG | pick 1-6 idx={idx} SUM MEAN MEDIAN MIN MAX COUNT | Enter confirm Esc cancel"),
        };
        let status = Paragraph::new(mode_line).block(Block::default().borders(Borders::ALL));
        f.render_widget(status, status_area);
    }

    fn handle_key(&mut self, key: KeyEvent) -> Result<bool, RunError> {
        match &mut self.mode {
            Mode::Help => {
                if matches!(key.code, KeyCode::Esc | KeyCode::Char('q') | KeyCode::Char('?')) {
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
                                && r1 < hr + self.state.grid.main_rows
                                && c0 >= MARGIN_COLS
                                && c1 < MARGIN_COLS + self.state.grid.main_cols
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
            Mode::OpenPath { buffer } => match key.code {
                KeyCode::Enter => {
                    let path = PathBuf::from(buffer.trim());
                    self.path = Some(path.clone());
                    self.offset = 0;
                    self.state = SheetState::new(4, 6);
                    if path.exists() {
                        let (off, n) = load_full(&path, &mut self.state)?;
                        self.offset = off;
                        self.ops_applied = n;
                    }
                    self.watcher = Some(LogWatcher::new(path.clone()).map_err(IoError::from)?);
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
            },
            Mode::Edit { buffer } => match key.code {
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
            },
            Mode::Normal => {}
        }

        if matches!(self.mode, Mode::OpenPath { .. } | Mode::Edit { .. } | Mode::AggPick { .. }) {
            return Ok(false);
        }

        if key.modifiers.contains(KeyModifiers::CONTROL) && key.code == KeyCode::Char('q') {
            return Ok(true);
        }

        match key.code {
            KeyCode::Char('?') => {
                self.mode = Mode::Help;
            }
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
                if self.anchor.is_none() {
                    self.anchor = Some(self.cursor);
                } else {
                    self.anchor = None;
                }
            }
            KeyCode::Char('a') => {
                self.mode = Mode::AggPick { idx: 0 };
            }
            KeyCode::Char('r') => {
                let hr = HEADER_ROWS;
                if let Some((mr0, mr1)) = self.selection_main_row_range() {
                    if self.cursor.row < hr
                        || self.cursor.row >= hr + self.state.grid.main_rows
                    {
                        self.status =
                            "Place cursor on a main row (1..N) as move target before pressing r"
                                .into();
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
                    self.status = format!("Moved rows {mr0}..{} to before row {to}", mr0 + count);
                } else {
                    self.status = "Select full main rows (visual v) for row move".into();
                }
            }
            KeyCode::Char('c') => {
                let left = MARGIN_COLS;
                let right = MARGIN_COLS + self.state.grid.main_cols;
                if let Some((mc0, mc1)) = self.selection_main_col_range() {
                    if self.cursor.col < left || self.cursor.col >= right {
                        self.status =
                            "Place cursor on a main column as move target before pressing c".into();
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
                    self.status = format!("Moved cols {mc0}..{} to before col {to}", mc0 + count);
                } else {
                    self.status = "Select full main columns (visual v) for col move".into();
                }
            }
            KeyCode::Left => {
                self.cursor.col = self.cursor.col.saturating_sub(1);
                self.cursor.clamp(&self.state.grid);
            }
            KeyCode::Right => {
                self.cursor.col = self.cursor.col.saturating_add(1);
                self.cursor.clamp(&self.state.grid);
            }
            KeyCode::Up => {
                self.cursor.row = self.cursor.row.saturating_sub(1);
                self.cursor.clamp(&self.state.grid);
            }
            KeyCode::Down => {
                self.cursor.row = self.cursor.row.saturating_add(1);
                self.cursor.clamp(&self.state.grid);
            }
            KeyCode::Char('q') => return Ok(true),
            _ => {}
        }

        Ok(false)
    }
}

fn row_label(r: usize, main_rows: usize) -> String {
    let hr = HEADER_ROWS;
    if r < hr {
        format!("^{}", (b'A' + r as u8) as char)
    } else if r < hr + main_rows {
        format!("{}", r - hr + 1)
    } else {
        format!("_{}", (b'A' + (r - hr - main_rows) as u8) as char)
    }
}

/// Column title for global column index: `<0`..`<9`, Excel-style `A`.. for main, `>0`..`>9` right.
fn col_label(global_col: usize, main_cols: usize) -> String {
    let m = MARGIN_COLS;
    if global_col < m {
        format!("<{}", global_col)
    } else if global_col < m + main_cols {
        excel_column_name(global_col - m)
    } else {
        format!(">{}", global_col - m - main_cols)
    }
}

/// Spreadsheet-style column letters (1-based: A, B, … Z, AA, AA, …).
fn excel_column_name(col_idx: usize) -> String {
    let mut n = col_idx + 1;
    let mut s = String::new();
    while n > 0 {
        n -= 1;
        s.push((b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    s.chars().rev().collect()
}
