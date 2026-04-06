//! Five-region sheet layout: headers ~A–~Z, footers _A–_Z, margins <0–<9 and >0–>9, and main data.
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
    /// `~` row: `row` 0 = ~A … 25 = ~Z; `col` is global column index.
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
                write!(f, "~{}(col {})", header_label(*row), col)
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

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct SortSpec {
    pub col: usize,
    pub desc: bool,
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
    pub view_sort_cols: Vec<SortSpec>,
    pub header: Vec<Vec<String>>,
    pub footer: Vec<Vec<String>>,
    pub(crate) spill_followers: HashMap<CellAddr, String>,
    pub(crate) spill_errors: HashMap<CellAddr, &'static str>,
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
            spill_followers: HashMap::new(),
            spill_errors: HashMap::new(),
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
        let old_main_cols = self.extent_main_cols as usize;
        let new_main_cols = old_main_cols.saturating_add(1);
        self.remap_main_col_layout_for_resize(old_main_cols, new_main_cols);
        self.extent_main_cols = new_main_cols as u32;
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
                let old_main_cols = self.extent_main_cols as usize;
                let new_main_cols = mc as usize + 1;
                self.remap_main_col_layout_for_resize(old_main_cols, new_main_cols);
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
        let old_main_cols = self.extent_main_cols as usize;
        let new_main_cols = main_cols.max(1);
        self.remap_main_col_layout_for_resize(old_main_cols, new_main_cols);
        self.extent_main_rows = main_rows.max(1) as u32;
        self.extent_main_cols = new_main_cols as u32;
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

    fn content_width_for_column(&self, global_col: usize) -> Option<usize> {
        let mut maxw = 0usize;
        let mut saw_content = false;
        let main_cols = self.main_cols();

        for r in 0..HEADER_ROWS {
            if let Some(val) = self.header.get(r).and_then(|row| row.get(global_col)) {
                if !val.is_empty() {
                    saw_content = true;
                    maxw = maxw.max(val.chars().count() + 1);
                }
            }
        }
        for r in 0..FOOTER_ROWS {
            if let Some(val) = self.footer.get(r).and_then(|row| row.get(global_col)) {
                if !val.is_empty() {
                    saw_content = true;
                    maxw = maxw.max(val.chars().count() + 1);
                }
            }
        }
        for r in 0..self.extent_main_rows as usize {
            if global_col < MARGIN_COLS {
                if let Some(val) = self.left.get(&(r as u32, global_col as u8)) {
                    if !val.is_empty() {
                        saw_content = true;
                        maxw = maxw.max(val.chars().count() + 1);
                    }
                }
            } else if global_col < MARGIN_COLS + main_cols {
                let mc = global_col - MARGIN_COLS;
                if let Some(val) = self.main_cells.get(&(r as u32, mc as u32)) {
                    if !val.is_empty() {
                        saw_content = true;
                        maxw = maxw.max(val.chars().count() + 1);
                    }
                }
            } else {
                let rc = global_col - MARGIN_COLS - main_cols;
                if let Some(val) = self.right.get(&(r as u32, rc as u8)) {
                    if !val.is_empty() {
                        saw_content = true;
                        maxw = maxw.max(val.chars().count() + 1);
                    }
                }
            }
        }

        saw_content.then_some(maxw.max(4))
    }

    pub fn auto_fit_column(&mut self, global_col: usize) {
        if let Some(maxw) = self.content_width_for_column(global_col) {
            if maxw > self.max_col_width {
                self.col_width_overrides.insert(global_col, maxw);
            }
        }
    }

    pub fn fit_column_to_content(&mut self, global_col: usize) {
        if let Some(maxw) = self.content_width_for_column(global_col) {
            self.col_width_overrides
                .insert(global_col, maxw.min(self.max_col_width));
        } else {
            self.col_width_overrides.remove(&global_col);
        }
    }

    fn remap_main_col_layout_for_resize(&mut self, old_main_cols: usize, new_main_cols: usize) {
        if old_main_cols == new_main_cols {
            return;
        }

        let old_total = MARGIN_COLS + old_main_cols + MARGIN_COLS;
        let new_total = MARGIN_COLS + new_main_cols + MARGIN_COLS;
        let old_right_start = MARGIN_COLS + old_main_cols;
        let new_right_start = MARGIN_COLS + new_main_cols;

        fn remap_col(
            col: usize,
            new_main_cols: usize,
            old_right_start: usize,
            new_right_start: usize,
        ) -> Option<usize> {
            if col < MARGIN_COLS {
                Some(col)
            } else if col < old_right_start {
                let main_idx = col - MARGIN_COLS;
                (main_idx < new_main_cols).then_some(MARGIN_COLS + main_idx)
            } else {
                let right_idx = col - old_right_start;
                Some(new_right_start + right_idx)
            }
        }

        for row in &mut self.header {
            let old_row = std::mem::take(row);
            let mut new_row = vec![String::new(); new_total];
            for (col, value) in old_row.into_iter().enumerate().take(old_total) {
                let new_col = remap_col(col, new_main_cols, old_right_start, new_right_start);
                if let Some(new_col) = new_col {
                    if new_col < new_total {
                        new_row[new_col] = value;
                    }
                }
            }
            *row = new_row;
        }

        for row in &mut self.footer {
            let old_row = std::mem::take(row);
            let mut new_row = vec![String::new(); new_total];
            for (col, value) in old_row.into_iter().enumerate().take(old_total) {
                let new_col = remap_col(col, new_main_cols, old_right_start, new_right_start);
                if let Some(new_col) = new_col {
                    if new_col < new_total {
                        new_row[new_col] = value;
                    }
                }
            }
            *row = new_row;
        }

        let mut remapped = HashMap::new();
        for (col, width) in self.col_width_overrides.drain() {
            let new_col = remap_col(col, new_main_cols, old_right_start, new_right_start);
            if let Some(new_col) = new_col {
                remapped.insert(new_col, width);
            }
        }
        self.col_width_overrides = remapped;
    }

    fn remap_main_col_width_overrides_for_order(&mut self, order: &[u32]) {
        let old_main_cols = order.len();
        if old_main_cols == 0 {
            return;
        }

        let mut old_to_new = vec![0usize; old_main_cols];
        for (new_pos, &old_pos) in order.iter().enumerate() {
            old_to_new[old_pos as usize] = new_pos;
        }

        let mut remapped = HashMap::new();
        for (col, width) in self.col_width_overrides.drain() {
            if col < MARGIN_COLS || col >= MARGIN_COLS + old_main_cols {
                remapped.insert(col, width);
            } else {
                let old_pos = col - MARGIN_COLS;
                remapped.insert(MARGIN_COLS + old_to_new[old_pos], width);
            }
        }
        self.col_width_overrides = remapped;
    }

    pub fn set_view_sort_cols(&mut self, cols: Vec<SortSpec>) {
        self.view_sort_cols = cols;
    }

    /// Logical main-row order for the current view sort.
    pub fn sorted_main_rows(&self) -> Vec<usize> {
        let mut rows: Vec<usize> = (0..self.extent_main_rows as usize).collect();
        if self.view_sort_cols.is_empty() {
            return rows;
        }

        rows.sort_by(|a, b| {
            for spec in &self.view_sort_cols {
                let global_col = spec.col;
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
                let ord = compare_sort_values(va, vb, spec.desc);
                if ord != std::cmp::Ordering::Equal {
                    return ord;
                }
            }
            a.cmp(b)
        });
        rows
    }

    pub fn sort_specs_to_log(cols: &[SortSpec]) -> String {
        cols.iter()
            .map(|spec| {
                let name = crate::addr::excel_column_name(spec.col.saturating_sub(MARGIN_COLS));
                if spec.desc {
                    format!("!{name}")
                } else {
                    name
                }
            })
            .collect::<Vec<_>>()
            .join(" ")
    }

    pub fn get(&self, addr: &CellAddr) -> Option<&str> {
        if let Some(v) = self.spill_followers.get(addr) {
            return Some(v.as_str());
        }
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

    pub(crate) fn spill_error(&self, addr: &CellAddr) -> Option<&'static str> {
        self.spill_errors.get(addr).copied()
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
                    self.main_cells.insert((r, c), value);
                    self.auto_fit_column(MARGIN_COLS + c as usize);
                    self.resize_header_footer_width();
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
                        self.left.insert((r, mc), value);
                        self.auto_fit_column(mc as usize);
                        self.resize_header_footer_width();
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
                        self.right.insert((r, mc), value);
                        self.auto_fit_column(
                            MARGIN_COLS + self.extent_main_cols as usize + mc as usize,
                        );
                        self.resize_header_footer_width();
                    }
                }
            }
        }
    }

    pub(crate) fn clear_spills(&mut self) {
        self.spill_followers.clear();
        self.spill_errors.clear();
    }

    pub(crate) fn set_spill_value(&mut self, addr: CellAddr, value: String) {
        self.spill_followers.insert(addr, value);
    }

    pub(crate) fn set_spill_error(&mut self, addr: CellAddr, err: &'static str) {
        self.spill_errors.insert(addr, err);
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

        self.remap_main_col_width_overrides_for_order(&order);

        self.extent_main_cols = order.len() as u32;
    }
}

#[derive(Debug, PartialEq)]
enum SortKey<'a> {
    Blank,
    Text(&'a str),
    Number(f64),
}

impl<'a> Eq for SortKey<'a> {}

fn sort_key(value: &str) -> SortKey<'_> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        SortKey::Blank
    } else if let Ok(n) = trimmed.parse::<f64>() {
        SortKey::Number(n)
    } else {
        SortKey::Text(trimmed)
    }
}

fn compare_sort_values(va: &str, vb: &str, desc: bool) -> std::cmp::Ordering {
    use std::cmp::Ordering;

    match (sort_key(va), sort_key(vb)) {
        (SortKey::Blank, SortKey::Blank) => Ordering::Equal,
        (SortKey::Blank, _) => Ordering::Greater,
        (_, SortKey::Blank) => Ordering::Less,
        (SortKey::Text(a), SortKey::Text(b)) => {
            if desc {
                b.cmp(a)
            } else {
                a.cmp(b)
            }
        }
        (SortKey::Number(a), SortKey::Number(b)) => {
            if desc {
                b.partial_cmp(&a).unwrap_or(Ordering::Equal)
            } else {
                a.partial_cmp(&b).unwrap_or(Ordering::Equal)
            }
        }
        (SortKey::Text(_), SortKey::Number(_)) => {
            if desc {
                Ordering::Greater
            } else {
                Ordering::Less
            }
        }
        (SortKey::Number(_), SortKey::Text(_)) => {
            if desc {
                Ordering::Less
            } else {
                Ordering::Greater
            }
        }
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

    #[test]
    fn sorted_rows_put_text_before_numbers() {
        let mut g = Grid::new(3, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "2".into());
        g.set(&CellAddr::Main { row: 1, col: 0 }, "apple".into());
        g.set(&CellAddr::Main { row: 2, col: 0 }, "10".into());
        g.set_view_sort_cols(vec![SortSpec {
            col: MARGIN_COLS,
            desc: false,
        }]);

        assert_eq!(g.sorted_main_rows(), vec![1, 0, 2]);
    }

    #[test]
    fn auto_fit_only_grows_touched_column() {
        let mut g = Grid::new(1, 2);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "short".into());
        g.set(
            &CellAddr::Main { row: 0, col: 1 },
            "abcdefghijklmnopqrstuvwx".into(),
        );

        assert_eq!(g.col_width(MARGIN_COLS), 20);
        assert!(g.col_width(MARGIN_COLS + 1) >= 24);
    }

    #[test]
    fn widths_shift_when_main_cols_grow() {
        let mut g = Grid::new(1, 1);
        g.set_col_width(MARGIN_COLS + 1, Some(24));

        g.grow_main_col_at_right();

        assert_eq!(g.col_width(MARGIN_COLS + 1), 20);
        assert_eq!(g.col_width(MARGIN_COLS + 2), 24);
    }

    #[test]
    fn widths_follow_moved_main_columns() {
        let mut g = Grid::new(1, 3);
        g.set_col_width(MARGIN_COLS + 1, Some(24));

        g.move_main_cols(1, 1, 3);

        assert_eq!(g.col_width(MARGIN_COLS + 1), 20);
        assert_eq!(g.col_width(MARGIN_COLS + 2), 24);
    }
}
