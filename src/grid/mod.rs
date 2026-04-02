//! Five-region sheet layout: headers ^A–^Z, footers _A–_Z, margins <0–<9 and >0–>9, and main data.
//! Main and margin cells use sparse storage for unbounded logical size.

use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fmt;

pub const HEADER_ROWS: usize = 26;
pub const FOOTER_ROWS: usize = 26;
pub const MARGIN_COLS: usize = 10;

/// Logical cell address (stable across main resize where possible).
#[derive(Clone, Debug, Eq, PartialEq, Hash, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum CellAddr {
    /// `^` row: `row` 0 = ^A … 25 = ^Z; `col` is global column index.
    Header { row: u8, col: u32 },
    /// `_` row: same indexing as headers.
    Footer { row: u8, col: u32 },
    /// Main grid.
    Main { row: u32, col: u32 },
    /// Left margin `<0`..`<9`: `col` 0–9, `row` is main row index.
    Left { col: u8, row: u32 },
    /// Right margin `>0`..`>9`.
    Right { col: u8, row: u32 },
}

impl fmt::Display for CellAddr {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CellAddr::Header { row, col } => {
                write!(f, "^{}(col {})", header_label(*row), col)
            }
            CellAddr::Footer { row, col } => {
                write!(f, "_{}(col {})", header_label(*row), col)
            }
            CellAddr::Main { row, col } => write!(f, "({}, {})", row, col),
            CellAddr::Left { col, row } => write!(f, "<{}>({})", col, row),
            CellAddr::Right { col, row } => write!(f, ">{}>({})", col, row),
        }
    }
}

fn header_label(row: u8) -> char {
    (b'A' + row.min(25)) as char
}

/// Inclusive-exclusive range in the **main** region (for aggregates).
#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct MainRange {
    pub row_start: u32,
    pub row_end: u32,
    pub col_start: u32,
    pub col_end: u32,
}

impl MainRange {
    pub fn is_empty(&self) -> bool {
        self.row_start >= self.row_end || self.col_start >= self.col_end
    }
}

/// Full sheet: dense header/footer; sparse main + left/right margins (infinite logical extent).
#[derive(Clone, Debug)]
pub struct Grid {
    /// Main cells; absent key = empty.
    pub main_cells: HashMap<(u32, u32), String>,
    /// Logical main size: at least 1×1; grows with data/cursor.
    pub extent_main_rows: u32,
    pub extent_main_cols: u32,
    /// Left margin: (main_row, margin_col 0–9).
    pub left: HashMap<(u32, u8), String>,
    /// Right margin: (main_row, margin_col 0–9).
    pub right: HashMap<(u32, u8), String>,
    /// Default display width cap for columns.
    pub max_col_width: usize,
    /// Optional per-global-column display width overrides.
    pub col_width_overrides: HashMap<usize, usize>,
    /// Optional sorted main-column view order.
    pub view_sort_cols: Vec<usize>,
    pub header: Vec<Vec<String>>,
    pub footer: Vec<Vec<String>>,
}

impl Default for Grid {
    fn default() -> Self {
        Self::new(1, 1)
    }
}

impl Grid {
    pub fn new(main_rows: u32, main_cols: u32) -> Self {
        let mut g = Grid {
            main_cells: HashMap::new(),
            extent_main_rows: main_rows.max(1),
            extent_main_cols: main_cols.max(1),
            left: HashMap::new(),
            right: HashMap::new(),
            max_col_width: 20,
            col_width_overrides: HashMap::new(),
            view_sort_cols: Vec::new(),
            header: Vec::new(),
            footer: Vec::new(),
        };
        g.resize_header_footer_width();
        g
    }

    /// One new main row at the bottom (cursor moving down from the last main row).
    pub fn grow_main_row_at_bottom(&mut self) {
        self.extent_main_rows = self.extent_main_rows.saturating_add(1);
        self.resize_header_footer_width();
    }

    /// One new main column at the right (cursor moving right from the last sheet column).
    pub fn grow_main_col_at_right(&mut self) {
        self.extent_main_cols = self.extent_main_cols.saturating_add(1);
        self.resize_header_footer_width();
    }

    /// Back-compat: logical main row count.
    #[inline]
    pub fn main_rows(&self) -> usize {
        self.extent_main_rows as usize
    }

    /// Back-compat: logical main column count.
    #[inline]
    pub fn main_cols(&self) -> usize {
        self.extent_main_cols as usize
    }

    pub fn total_cols(&self) -> usize {
        MARGIN_COLS + self.extent_main_cols as usize + MARGIN_COLS
    }

    pub fn total_logical_rows(&self) -> usize {
        HEADER_ROWS + self.extent_main_rows as usize + FOOTER_ROWS
    }

    /// Grow extent so cursor (logical row/col) is addressable in main/margins.
    /// Returns true if the extent was actually grown (for UI feedback).
    pub fn ensure_extent_for_cursor(&mut self, row: usize, col: usize) -> bool {
        let hr = HEADER_ROWS;
        let m = MARGIN_COLS;
        let main_end = m + self.extent_main_cols as usize;
        let mut grown = false;
        if (hr..hr + self.extent_main_rows as usize).contains(&row) && (m..main_end).contains(&col)
        {
            let mr = (row - hr) as u32;
            let mc = (col - m) as u32;
            if mr + 1 > self.extent_main_rows {
                self.extent_main_rows = mr + 1;
                grown = true;
            }
            if mc + 1 > self.extent_main_cols {
                self.extent_main_cols = mc + 1;
                grown = true;
            }
        } else if (hr..hr + self.extent_main_rows as usize).contains(&row) && (0..m).contains(&col)
        {
            let mr = (row - hr) as u32;
            if mr + 1 > self.extent_main_rows {
                self.extent_main_rows = mr + 1;
                grown = true;
            }
        } else if (hr..hr + self.extent_main_rows as usize).contains(&row)
            && (main_end..main_end + MARGIN_COLS).contains(&col)
        {
            let mr = (row - hr) as u32;
            if mr + 1 > self.extent_main_rows {
                self.extent_main_rows = mr + 1;
                grown = true;
            }
        }
        if grown {
            self.resize_header_footer_width();
        }
        grown
    }

    pub fn logical_row_has_content(&self, r: usize) -> bool {
        let hr = HEADER_ROWS;
        if r < hr {
            return self
                .header
                .get(r)
                .is_some_and(|row| row.iter().any(|s| !s.is_empty()));
        }
        if r < hr + self.extent_main_rows as usize {
            let mr = r - hr;
            let mru = mr as u32;
            return self.main_cells.keys().any(|(row, _)| *row == mru)
                || self.left.keys().any(|(row, _)| *row == mru)
                || self.right.keys().any(|(row, _)| *row == mru);
        }
        let fr = r - hr - self.extent_main_rows as usize;
        self.footer
            .get(fr)
            .is_some_and(|row| row.iter().any(|s| !s.is_empty()))
    }

    pub fn logical_col_has_content(&self, c: usize) -> bool {
        let tc = self.total_cols();
        if c >= tc {
            return false;
        }
        for r in 0..HEADER_ROWS {
            if self.header[r].get(c).is_some_and(|s| !s.is_empty()) {
                return true;
            }
        }
        let m = MARGIN_COLS;
        let me = m + self.extent_main_cols as usize;
        if c < m {
            return self.left.keys().any(|(_, mc)| *mc as usize == c);
        }
        if c < me {
            let mc = (c - m) as u32;
            return self.main_cells.keys().any(|(_, col)| *col == mc);
        }
        if c < me + MARGIN_COLS {
            let mc = (c - me) as u8;
            return self.right.keys().any(|(_, rc)| *rc == mc);
        }
        for fr in 0..FOOTER_ROWS {
            if self.footer[fr].get(c).is_some_and(|s| !s.is_empty()) {
                return true;
            }
        }
        false
    }

    fn resize_header_footer_width(&mut self) {
        let w = self.total_cols();
        self.header.resize(HEADER_ROWS, Vec::new());
        for row in &mut self.header {
            row.resize(w, String::new());
        }
        self.footer.resize(FOOTER_ROWS, Vec::new());
        for row in &mut self.footer {
            row.resize(w, String::new());
        }
    }

    pub fn set_main_size(&mut self, main_rows: usize, main_cols: usize) {
        self.extent_main_rows = main_rows.max(1) as u32;
        self.extent_main_cols = main_cols.max(1) as u32;
        self.main_cells
            .retain(|&(r, c), _| r < self.extent_main_rows && c < self.extent_main_cols);
        self.left.retain(|&(r, _), _| r < self.extent_main_rows);
        self.right.retain(|&(r, _), _| r < self.extent_main_rows);
        self.resize_header_footer_width();
    }

    pub fn col_width(&self, global_col: usize) -> usize {
        self.col_width_overrides
            .get(&global_col)
            .copied()
            .unwrap_or(self.max_col_width)
            .max(1)
    }

    pub fn set_max_col_width(&mut self, width: usize) {
        self.max_col_width = width.max(1);
    }

    pub fn set_col_width(&mut self, global_col: usize, width: Option<usize>) {
        match width {
            Some(w) => {
                self.col_width_overrides.insert(global_col, w.max(1));
            }
            None => {
                self.col_width_overrides.remove(&global_col);
            }
        }
    }

    pub fn set_view_sort_cols(&mut self, cols: Vec<usize>) {
        self.view_sort_cols = cols;
    }

    /// Logical main-row order for the current view sort.
    pub fn sorted_main_rows(&self) -> Vec<usize> {
        let mut rows: Vec<usize> = (0..self.extent_main_rows as usize).collect();
        if self.view_sort_cols.is_empty() {
            return rows;
        }

        rows.sort_by(|a, b| {
            for &global_col in &self.view_sort_cols {
                if global_col < MARGIN_COLS || global_col >= MARGIN_COLS + self.main_cols() {
                    continue;
                }
                let col = (global_col - MARGIN_COLS) as u32;
                let va = self
                    .get(&CellAddr::Main {
                        row: *a as u32,
                        col,
                    })
                    .unwrap_or("");
                let vb = self
                    .get(&CellAddr::Main {
                        row: *b as u32,
                        col,
                    })
                    .unwrap_or("");
                let ord = va.cmp(vb);
                if ord != std::cmp::Ordering::Equal {
                    return ord;
                }
            }
            a.cmp(b)
        });
        rows
    }

    pub fn get(&self, addr: &CellAddr) -> Option<&str> {
        match addr {
            CellAddr::Header { row, col } => {
                let r = *row as usize;
                let c = *col as usize;
                self.header
                    .get(r)
                    .and_then(|row| row.get(c))
                    .map(|s| s.as_str())
            }
            CellAddr::Footer { row, col } => {
                let r = *row as usize;
                let c = *col as usize;
                self.footer
                    .get(r)
                    .and_then(|row| row.get(c))
                    .map(|s| s.as_str())
            }
            CellAddr::Main { row, col } => self.main_cells.get(&(*row, *col)).map(|s| s.as_str()),
            CellAddr::Left { col, row } => self.left.get(&(*row, *col)).map(|s| s.as_str()),
            CellAddr::Right { col, row } => self.right.get(&(*row, *col)).map(|s| s.as_str()),
        }
    }

    pub fn set(&mut self, addr: &CellAddr, value: String) {
        match addr {
            CellAddr::Header { row, col } => {
                let r = *row as usize;
                let c = *col as usize;
                if r < HEADER_ROWS && c < self.total_cols() {
                    self.header[r][c] = value;
                }
            }
            CellAddr::Footer { row, col } => {
                let r = *row as usize;
                let c = *col as usize;
                if r < FOOTER_ROWS && c < self.total_cols() {
                    self.footer[r][c] = value;
                }
            }
            CellAddr::Main { row, col } => {
                let r = *row;
                let c = *col;
                if value.is_empty() {
                    self.main_cells.remove(&(r, c));
                } else {
                    self.extent_main_rows = self.extent_main_rows.max(r + 1);
                    self.extent_main_cols = self.extent_main_cols.max(c + 1);
                    self.resize_header_footer_width();
                    self.main_cells.insert((r, c), value);
                }
            }
            CellAddr::Left { col, row } => {
                let mc = *col;
                let r = *row;
                if (mc as usize) < MARGIN_COLS {
                    if value.is_empty() {
                        self.left.remove(&(r, mc));
                    } else {
                        self.extent_main_rows = self.extent_main_rows.max(r + 1);
                        self.resize_header_footer_width();
                        self.left.insert((r, mc), value);
                    }
                }
            }
            CellAddr::Right { col, row } => {
                let mc = *col;
                let r = *row;
                if (mc as usize) < MARGIN_COLS {
                    if value.is_empty() {
                        self.right.remove(&(r, mc));
                    } else {
                        self.extent_main_rows = self.extent_main_rows.max(r + 1);
                        self.resize_header_footer_width();
                        self.right.insert((r, mc), value);
                    }
                }
            }
        }
    }

    pub fn move_main_rows(&mut self, from: usize, count: usize, to: usize) {
        let er = self.extent_main_rows as usize;
        if count == 0 || from + count > er || to > er {
            return;
        }
        if to > from && to < from + count {
            return;
        }
        let insert_at = if to > from { to - count } else { to };

        let mut order: Vec<u32> = (0..self.extent_main_rows).collect();
        let taken: Vec<u32> = order.drain(from..from + count).collect();
        order.splice(insert_at..insert_at, taken);

        let mut new_main = HashMap::new();
        for (new_pos, &old_r) in order.iter().enumerate() {
            let old_r = old_r;
            for c in 0..self.extent_main_cols {
                if let Some(v) = self.main_cells.get(&(old_r, c)).cloned() {
                    new_main.insert((new_pos as u32, c), v);
                }
            }
        }
        self.main_cells = new_main;

        let mut new_left = HashMap::new();
        for (new_pos, &old_r) in order.iter().enumerate() {
            for mc in 0..MARGIN_COLS as u8 {
                if let Some(v) = self.left.get(&(old_r, mc)).cloned() {
                    new_left.insert((new_pos as u32, mc), v);
                }
            }
        }
        self.left = new_left;

        let mut new_right = HashMap::new();
        for (new_pos, &old_r) in order.iter().enumerate() {
            for mc in 0..MARGIN_COLS as u8 {
                if let Some(v) = self.right.get(&(old_r, mc)).cloned() {
                    new_right.insert((new_pos as u32, mc), v);
                }
            }
        }
        self.right = new_right;

        self.extent_main_rows = order.len() as u32;
    }

    pub fn move_main_cols(&mut self, from: usize, count: usize, to: usize) {
        let ec = self.extent_main_cols as usize;
        if count == 0 || from + count > ec || to > ec {
            return;
        }
        if to > from && to < from + count {
            return;
        }
        let insert_at = if to > from { to - count } else { to };

        let mut order: Vec<u32> = (0..self.extent_main_cols).collect();
        let taken: Vec<u32> = order.drain(from..from + count).collect();
        order.splice(insert_at..insert_at, taken);

        let mut new_main = HashMap::new();
        for r in 0..self.extent_main_rows {
            for (new_pos, &old_c) in order.iter().enumerate() {
                if let Some(v) = self.main_cells.get(&(r, old_c)).cloned() {
                    new_main.insert((r, new_pos as u32), v);
                }
            }
        }
        self.main_cells = new_main;

        let g_from = MARGIN_COLS + from;
        let insert_global = MARGIN_COLS + insert_at;

        for row in &mut self.header {
            if g_from + count <= row.len() && insert_global + count <= row.len() {
                let block: Vec<String> = row.drain(g_from..g_from + count).collect();
                row.splice(insert_global..insert_global, block);
            }
        }
        for row in &mut self.footer {
            if g_from + count <= row.len() && insert_global + count <= row.len() {
                let block: Vec<String> = row.drain(g_from..g_from + count).collect();
                row.splice(insert_global..insert_global, block);
            }
        }

        self.extent_main_cols = order.len() as u32;
    }
}

/// Logical sheet row index (0 = top header row) for addressing.
pub fn addr_logical_row(addr: &CellAddr, grid: &Grid) -> usize {
    let hr = HEADER_ROWS;
    match addr {
        CellAddr::Header { row, .. } => *row as usize,
        CellAddr::Main { row, .. } => hr + *row as usize,
        CellAddr::Left { row, .. } | CellAddr::Right { row, .. } => hr + *row as usize,
        CellAddr::Footer { row, .. } => hr + grid.extent_main_rows as usize + *row as usize,
    }
}

/// Global column index for addressing.
pub fn addr_logical_col(addr: &CellAddr, grid: &Grid) -> usize {
    match addr {
        CellAddr::Header { col, .. }
        | CellAddr::Footer { col, .. }
        | CellAddr::Main { col, .. } => *col as usize,
        CellAddr::Left { col, .. } => *col as usize,
        CellAddr::Right { col, .. } => MARGIN_COLS + grid.extent_main_cols as usize + *col as usize,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn move_rows_sparse() {
        let mut g = Grid::new(4, 2);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        g.set(&CellAddr::Main { row: 1, col: 0 }, "b".into());
        g.set(&CellAddr::Main { row: 2, col: 0 }, "c".into());
        g.set(&CellAddr::Main { row: 3, col: 0 }, "d".into());
        g.move_main_rows(0, 2, 4);
        assert_eq!(g.get(&CellAddr::Main { row: 0, col: 0 }), Some("c"));
        assert_eq!(g.get(&CellAddr::Main { row: 1, col: 0 }), Some("d"));
        assert_eq!(g.get(&CellAddr::Main { row: 2, col: 0 }), Some("a"));
        assert_eq!(g.get(&CellAddr::Main { row: 3, col: 0 }), Some("b"));
    }

    #[test]
    fn move_cols_sparse() {
        let mut g = Grid::new(2, 4);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "b".into());
        g.set(&CellAddr::Main { row: 0, col: 2 }, "c".into());
        g.set(&CellAddr::Main { row: 0, col: 3 }, "d".into());
        g.move_main_cols(0, 2, 4);
        assert_eq!(g.get(&CellAddr::Main { row: 0, col: 0 }), Some("c"));
        assert_eq!(g.get(&CellAddr::Main { row: 0, col: 1 }), Some("d"));
        assert_eq!(g.get(&CellAddr::Main { row: 0, col: 2 }), Some("a"));
        assert_eq!(g.get(&CellAddr::Main { row: 0, col: 3 }), Some("b"));
    }
}
