//! Five-region sheet layout: headers ^A–^Z, footers _A–_Z, margins <0–<9 and >0–>9, and main data.

use serde::{Deserialize, Serialize};
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

/// Full sheet grid: dense storage for fast row/column moves.
#[derive(Clone, Debug, Default)]
pub struct Grid {
    pub main_rows: usize,
    pub main_cols: usize,
    /// Main[row][col]
    pub main: Vec<Vec<String>>,
    /// Left[main_row][margin_col 0..10]
    pub left: Vec<Vec<String>>,
    /// Right[main_row][margin_col 0..10]
    pub right: Vec<Vec<String>>,
    /// Header[header_row][global_col]
    pub header: Vec<Vec<String>>,
    /// Footer[footer_row][global_col]
    pub footer: Vec<Vec<String>>,
}

impl Grid {
    pub fn new(main_rows: usize, main_cols: usize) -> Self {
        let mut g = Grid {
            main_rows,
            main_cols,
            main: vec![vec![String::new(); main_cols]; main_rows],
            left: vec![vec![String::new(); MARGIN_COLS]; main_rows],
            right: vec![vec![String::new(); MARGIN_COLS]; main_rows],
            header: (0..HEADER_ROWS).map(|_| Vec::new()).collect(),
            footer: (0..FOOTER_ROWS).map(|_| Vec::new()).collect(),
        };
        g.resize_header_footer_width();
        g
    }

    pub fn total_cols(&self) -> usize {
        MARGIN_COLS + self.main_cols + MARGIN_COLS
    }

    fn resize_header_footer_width(&mut self) {
        let w = self.total_cols();
        for row in &mut self.header {
            row.resize(w, String::new());
        }
        for row in &mut self.footer {
            row.resize(w, String::new());
        }
    }

    pub fn set_main_size(&mut self, main_rows: usize, main_cols: usize) {
        self.main_rows = main_rows.max(1);
        self.main_cols = main_cols.max(1);
        self.main = vec![vec![String::new(); self.main_cols]; self.main_rows];
        self.left = vec![vec![String::new(); MARGIN_COLS]; self.main_rows];
        self.right = vec![vec![String::new(); MARGIN_COLS]; self.main_rows];
        self.resize_header_footer_width();
    }

    pub fn get(&self, addr: &CellAddr) -> Option<&str> {
        match addr {
            CellAddr::Header { row, col } => {
                let r = *row as usize;
                let c = *col as usize;
                self.header.get(r).and_then(|row| row.get(c)).map(|s| s.as_str())
            }
            CellAddr::Footer { row, col } => {
                let r = *row as usize;
                let c = *col as usize;
                self.footer.get(r).and_then(|row| row.get(c)).map(|s| s.as_str())
            }
            CellAddr::Main { row, col } => {
                let r = *row as usize;
                let c = *col as usize;
                self.main.get(r).and_then(|row| row.get(c)).map(|s| s.as_str())
            }
            CellAddr::Left { col, row } => {
                let mc = *col as usize;
                let r = *row as usize;
                self.left.get(r).and_then(|row| row.get(mc)).map(|s| s.as_str())
            }
            CellAddr::Right { col, row } => {
                let mc = *col as usize;
                let r = *row as usize;
                self.right.get(r).and_then(|row| row.get(mc)).map(|s| s.as_str())
            }
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
                let r = *row as usize;
                let c = *col as usize;
                if r < self.main_rows && c < self.main_cols {
                    self.main[r][c] = value;
                }
            }
            CellAddr::Left { col, row } => {
                let mc = *col as usize;
                let r = *row as usize;
                if mc < MARGIN_COLS && r < self.main_rows {
                    self.left[r][mc] = value;
                }
            }
            CellAddr::Right { col, row } => {
                let mc = *col as usize;
                let r = *row as usize;
                if mc < MARGIN_COLS && r < self.main_rows {
                    self.right[r][mc] = value;
                }
            }
        }
    }

    /// Move `count` contiguous main rows starting at `from` to destination index `to` (in the
    /// **original** row order before the move; may be `main_rows` to append after the last row).
    pub fn move_main_rows(&mut self, from: usize, count: usize, to: usize) {
        if count == 0 || from + count > self.main_rows || to > self.main_rows {
            return;
        }
        if to > from && to < from + count {
            return;
        }
        let insert_at = if to > from { to - count } else { to };

        let extract = |v: &mut Vec<Vec<String>>| {
            let taken: Vec<Vec<String>> = v.drain(from..from + count).collect();
            v.splice(insert_at..insert_at, taken);
        };

        extract(&mut self.main);
        extract(&mut self.left);
        extract(&mut self.right);
        self.main_rows = self.main.len();
    }

    /// Move `count` contiguous main columns starting at `from` to destination index `to` (same
    /// semantics as [`Self::move_main_rows`]).
    pub fn move_main_cols(&mut self, from: usize, count: usize, to: usize) {
        if count == 0 || from + count > self.main_cols || to > self.main_cols {
            return;
        }
        if to > from && to < from + count {
            return;
        }
        let insert_at = if to > from { to - count } else { to };

        for row in &mut self.main {
            let block: Vec<String> = row.drain(from..from + count).collect();
            row.splice(insert_at..insert_at, block);
        }
        let g_from = MARGIN_COLS + from;
        let insert_global = MARGIN_COLS + insert_at;

        for row in &mut self.header {
            let block: Vec<String> = row.drain(g_from..g_from + count).collect();
            row.splice(insert_global..insert_global, block);
        }
        for row in &mut self.footer {
            let block: Vec<String> = row.drain(g_from..g_from + count).collect();
            row.splice(insert_global..insert_global, block);
        }

        self.main_cols = self.main.first().map(|r| r.len()).unwrap_or(1);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn move_rows_to_end() {
        let mut g = Grid::new(4, 2);
        g.main[0][0] = "a".into();
        g.main[1][0] = "b".into();
        g.main[2][0] = "c".into();
        g.main[3][0] = "d".into();
        g.move_main_rows(0, 2, 4);
        assert_eq!(g.main[0][0], "c");
        assert_eq!(g.main[1][0], "d");
        assert_eq!(g.main[2][0], "a");
        assert_eq!(g.main[3][0], "b");
    }

    #[test]
    fn move_cols() {
        let mut g = Grid::new(2, 4);
        g.main[0][0] = "a".into();
        g.main[0][1] = "b".into();
        g.main[0][2] = "c".into();
        g.main[0][3] = "d".into();
        g.move_main_cols(0, 2, 4);
        assert_eq!(g.main[0][0], "c");
        assert_eq!(g.main[0][1], "d");
        assert_eq!(g.main[0][2], "a");
        assert_eq!(g.main[0][3], "b");
    }
}
