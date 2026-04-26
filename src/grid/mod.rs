//! Five-region sheet layout: headers `~N`, footers `_N`, margins, and main data.
//! Main and margin cells use sparse storage for unbounded logical size.

use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fmt;

pub const HEADER_ROWS: usize = 999_999_999;
pub const FOOTER_ROWS: usize = 999_999_999;
/// Number of margin columns on each side. Expanded to support multi-letter
/// mirror names (e.g. A..ZZ). Use usize for indexes.
pub const MARGIN_COLS: usize = 26 * 27; // A..ZZ inclusive

/// Type alias for margin column indices to make it easy to widen the type in
/// one place if needed.
pub type MarginIndex = usize;

/// Logical cell address (stable across main resize where possible).
#[derive(Clone, Debug, Eq, PartialEq, Hash, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum CellAddr {
    /// `~` row: `row` 0 = the top header row; `col` is global column index.
    Header { row: u32, col: u32 },
    /// `_` row: same indexing as headers.
    Footer { row: u32, col: u32 },
    /// Main grid.
    Main { row: u32, col: u32 },
    /// Left margin: `col` is a MarginIndex (usize), `row` is main row index.
    Left { col: MarginIndex, row: u32 },
    /// Right margin: `col` is a MarginIndex (usize).
    Right { col: MarginIndex, row: u32 },
}

impl fmt::Display for CellAddr {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            CellAddr::Header { row, col } => {
                write!(f, "~{}(col {})", HEADER_ROWS as u32 - row, col)
            }
            CellAddr::Footer { row, col } => write!(f, "_{}(col {})", row + 1, col),
            CellAddr::Main { row, col } => write!(f, "({}, {})", row, col),
            CellAddr::Left { col, row } => write!(f, "<{}>({})", col, row),
            CellAddr::Right { col, row } => write!(f, ">{}>({})", col, row),
        }
    }
}

// Abstraction trait for Grid implementations.
// Methods return owned Strings where necessary to keep the trait object-safe.
pub trait GridImpl {
    // Basic size/query
    fn main_rows(&self) -> usize;
    fn main_cols(&self) -> usize;
    fn total_cols(&self) -> usize;
    fn total_logical_rows(&self) -> usize;

    // Cell access (owned returns keep the trait object-safe)
    fn get_owned(&self, addr: &CellAddr) -> Option<String>;
    fn text(&self, addr: &CellAddr) -> String;
    fn set_owned(&mut self, addr: &CellAddr, value: String);
    fn set(&mut self, addr: &CellAddr, value: String);

    // Layout / extent
    fn set_main_size(&mut self, main_rows: usize, main_cols: usize);
    fn ensure_extent_for_cursor(&mut self, row: usize, col: usize) -> bool;
    fn grow_main_row_at_bottom(&mut self);
    fn grow_main_col_at_right(&mut self);
    fn move_main_rows(&mut self, from: usize, count: usize, to: usize);
    fn move_main_cols(&mut self, from: usize, count: usize, to: usize);

    // Column sizing and widths
    fn max_col_width(&self) -> usize;
    fn col_width(&self, global_col: usize) -> usize;
    fn get_col_width_override(&self, global_col: usize) -> Option<usize>;
    fn content_width_for_column(&self, global_col: usize) -> Option<usize>;
    fn set_max_col_width(&mut self, width: usize);
    fn set_col_width(&mut self, global_col: usize, width: Option<usize>);
    fn auto_fit_column(&mut self, global_col: usize);
    fn fit_column_to_content(&mut self, global_col: usize);
    fn col_width_overrides(&self) -> Vec<(usize, usize)>;

    // Clear main and margin cells (used by callers that need a fresh grid).
    fn clear_cells(&mut self);

    // Replace the entire set of column width overrides.
    fn set_col_width_overrides(&mut self, overrides: Vec<(usize, usize)>);

    // Formatting
    fn set_view_sort_cols(&mut self, cols: Vec<SortSpec>);
    fn view_sort_cols(&self) -> Vec<SortSpec>;
    fn sorted_main_rows(&self) -> Vec<usize>;
    fn set_column_format(&mut self, scope: FormatScope, col: usize, format: CellFormat);
    fn set_cell_format(&mut self, addr: CellAddr, format: CellFormat);
    fn format_for_addr(&self, addr: &CellAddr) -> CellFormat;
    fn format_for_global_col(&self, scope: FormatScope, col: usize) -> CellFormat;
    fn col_all_formats(&self) -> Vec<(usize, CellFormat)>;
    fn col_data_formats(&self) -> Vec<(usize, CellFormat)>;
    fn col_special_formats(&self) -> Vec<(usize, CellFormat)>;
    fn cell_formats(&self) -> Vec<(CellAddr, CellFormat)>;

    // Spills / volatile
    fn clear_spills(&mut self);
    fn set_spill_value(&mut self, addr: CellAddr, value: String);
    fn set_spill_error(&mut self, addr: CellAddr, err: &'static str);
    fn spill_error(&self, addr: &CellAddr) -> Option<&'static str>;
    // Return current spill follower mappings (addr -> value).
    fn spill_followers(&self) -> Vec<(CellAddr, String)>;
    // Return current spill error mappings (addr -> static error tag).
    fn spill_errors(&self) -> Vec<(CellAddr, &'static str)>;
    fn bump_volatile_seed(&mut self);
    fn volatile_seed(&self) -> u64;
    fn set_volatile_seed(&mut self, seed: u64);

    // Logical content queries
    fn logical_row_has_content(&self, r: usize) -> bool;
    fn logical_col_has_content(&self, c: usize) -> bool;

    // Iteration
    fn iter_nonempty(&self) -> Box<dyn Iterator<Item = (CellAddr, String)> + '_>;

    // Clone trait-object helper
    fn clone_box(&self) -> Box<dyn GridImpl>;
}

/// A boxed handle to an abstract Grid implementation.
pub struct GridBox {
    pub inner: Box<dyn GridImpl>,
}

impl GridBox {
    pub fn new<G: GridImpl + 'static>(g: G) -> Self {
        Self { inner: Box::new(g) }
    }

    pub fn main_rows(&self) -> usize {
        self.inner.main_rows()
    }

    pub fn main_cols(&self) -> usize {
        self.inner.main_cols()
    }

    pub fn total_cols(&self) -> usize {
        self.inner.total_cols()
    }

    pub fn total_logical_rows(&self) -> usize {
        self.inner.total_logical_rows()
    }

    pub fn get_owned(&self, addr: &CellAddr) -> Option<String> {
        self.inner.get_owned(addr)
    }

    /// Convenience owned-get that mirrors the old Grid::get (returns owned String)
    pub fn get(&self, addr: &CellAddr) -> Option<String> {
        self.inner.get_owned(addr)
    }

    pub fn text(&self, addr: &CellAddr) -> String {
        self.inner.text(addr)
    }

    pub fn set_owned(&mut self, addr: &CellAddr, value: String) {
        self.inner.set_owned(addr, value)
    }

    pub fn set(&mut self, addr: &CellAddr, value: String) {
        self.inner.set(addr, value)
    }

    pub fn set_main_size(&mut self, r: usize, c: usize) {
        self.inner.set_main_size(r, c)
    }

    pub fn ensure_extent_for_cursor(&mut self, row: usize, col: usize) -> bool {
        self.inner.ensure_extent_for_cursor(row, col)
    }

    pub fn grow_main_row_at_bottom(&mut self) {
        self.inner.grow_main_row_at_bottom()
    }

    pub fn grow_main_col_at_right(&mut self) {
        self.inner.grow_main_col_at_right()
    }

    pub fn move_main_rows(&mut self, from: usize, count: usize, to: usize) {
        self.inner.move_main_rows(from, count, to)
    }

    pub fn move_main_cols(&mut self, from: usize, count: usize, to: usize) {
        self.inner.move_main_cols(from, count, to)
    }

    pub fn bump_volatile_seed(&mut self) {
        self.inner.bump_volatile_seed()
    }

    pub fn volatile_seed(&self) -> u64 {
        self.inner.volatile_seed()
    }

    pub fn set_volatile_seed(&mut self, seed: u64) {
        self.inner.set_volatile_seed(seed)
    }

    pub fn spill_followers(&self) -> Vec<(CellAddr, String)> {
        self.inner.spill_followers()
    }

    pub fn spill_errors(&self) -> Vec<(CellAddr, &'static str)> {
        self.inner.spill_errors()
    }

    pub fn max_col_width(&self) -> usize {
        self.inner.max_col_width()
    }

    pub fn col_width(&self, global_col: usize) -> usize {
        self.inner.col_width(global_col)
    }

    pub fn get_col_width_override(&self, global_col: usize) -> Option<usize> {
        self.inner.get_col_width_override(global_col)
    }

    pub fn content_width_for_column(&self, global_col: usize) -> Option<usize> {
        self.inner.content_width_for_column(global_col)
    }

    pub fn set_max_col_width(&mut self, width: usize) {
        self.inner.set_max_col_width(width)
    }

    pub fn set_col_width(&mut self, global_col: usize, width: Option<usize>) {
        self.inner.set_col_width(global_col, width)
    }

    pub fn auto_fit_column(&mut self, global_col: usize) {
        self.inner.auto_fit_column(global_col)
    }

    pub fn fit_column_to_content(&mut self, global_col: usize) {
        self.inner.fit_column_to_content(global_col)
    }

    pub fn col_width_overrides(&self) -> Vec<(usize, usize)> {
        self.inner.col_width_overrides()
    }

    pub fn clear_cells(&mut self) {
        self.inner.clear_cells()
    }

    pub fn set_col_width_overrides(&mut self, overrides: Vec<(usize, usize)>) {
        self.inner.set_col_width_overrides(overrides)
    }

    pub fn set_view_sort_cols(&mut self, cols: Vec<SortSpec>) {
        self.inner.set_view_sort_cols(cols)
    }

    pub fn view_sort_cols(&self) -> Vec<SortSpec> {
        self.inner.view_sort_cols()
    }

    pub fn sorted_main_rows(&self) -> Vec<usize> {
        self.inner.sorted_main_rows()
    }

    pub fn set_column_format(&mut self, scope: FormatScope, col: usize, format: CellFormat) {
        self.inner.set_column_format(scope, col, format)
    }

    pub fn set_cell_format(&mut self, addr: CellAddr, format: CellFormat) {
        self.inner.set_cell_format(addr, format)
    }

    pub fn format_for_addr(&self, addr: &CellAddr) -> CellFormat {
        self.inner.format_for_addr(addr)
    }

    pub fn format_for_global_col(&self, scope: FormatScope, col: usize) -> CellFormat {
        self.inner.format_for_global_col(scope, col)
    }

    pub fn clear_spills(&mut self) {
        self.inner.clear_spills()
    }

    pub fn set_spill_value(&mut self, addr: CellAddr, value: String) {
        self.inner.set_spill_value(addr, value)
    }

    pub fn set_spill_error(&mut self, addr: CellAddr, err: &'static str) {
        self.inner.set_spill_error(addr, err)
    }

    pub fn spill_error(&self, addr: &CellAddr) -> Option<&'static str> {
        self.inner.spill_error(addr)
    }

    pub fn logical_row_has_content(&self, r: usize) -> bool {
        self.inner.logical_row_has_content(r)
    }

    pub fn logical_col_has_content(&self, c: usize) -> bool {
        self.inner.logical_col_has_content(c)
    }

    pub fn col_all_formats(&self) -> Vec<(usize, CellFormat)> {
        self.inner.col_all_formats()
    }

    pub fn col_data_formats(&self) -> Vec<(usize, CellFormat)> {
        self.inner.col_data_formats()
    }

    pub fn col_special_formats(&self) -> Vec<(usize, CellFormat)> {
        self.inner.col_special_formats()
    }

    pub fn cell_formats(&self) -> Vec<(CellAddr, CellFormat)> {
        self.inner.cell_formats()
    }

    pub fn iter_nonempty(&self) -> Box<dyn Iterator<Item = (CellAddr, String)> + '_> {
        self.inner.iter_nonempty()
    }
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

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub enum NumberFormat {
    Currency { decimals: usize },
    Fixed { decimals: usize },
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub enum TextAlign {
    Left,
    Center,
    Right,
    Default,
}

#[derive(Clone, Copy, Debug, Default, Eq, PartialEq, Serialize, Deserialize)]
pub struct CellFormat {
    pub number: Option<NumberFormat>,
    pub align: Option<TextAlign>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub enum FormatScope {
    All,
    Data,
    Special,
}

/// Full sheet with sparse storage for each editable region.
#[derive(Clone, Debug)]
pub struct Grid {
    /// Main cells; absent key = empty.
    pub main_cells: HashMap<(u32, u32), String>,
    /// Logical main size: at least 1×1; grows with data/cursor.
    pub extent_main_rows: u32,
    pub extent_main_cols: u32,
    /// Left margin: (main_row, margin_col).
    pub left: HashMap<(u32, MarginIndex), String>,
    /// Right margin: (main_row, margin_col).
    pub right: HashMap<(u32, MarginIndex), String>,
    /// Default display width cap for columns.
    pub max_col_width: usize,
    /// Optional per-global-column display width overrides.
    pub col_width_overrides: HashMap<usize, usize>,
    /// Optional sorted main-column view order.
    pub view_sort_cols: Vec<SortSpec>,
    /// Column-wide format for all cells in a global column.
    pub col_all_formats: HashMap<usize, CellFormat>,
    /// Column-wide format for main-region cells in a global column.
    pub col_data_formats: HashMap<usize, CellFormat>,
    /// Column-wide format for header/footer/margin cells in a global column.
    pub col_special_formats: HashMap<usize, CellFormat>,
    /// Exact-cell overrides used for Cell/Selection formatting.
    pub cell_formats: HashMap<CellAddr, CellFormat>,
    pub header: HashMap<(u32, u32), String>,
    pub footer: HashMap<(u32, u32), String>,
    pub(crate) spill_followers: HashMap<CellAddr, String>,
    pub(crate) spill_errors: HashMap<CellAddr, &'static str>,
    pub(crate) volatile_seed: u64,
}

impl Default for Grid {
    fn default() -> Self {
        Self::new(1, 1)
    }
}

impl Grid {
    pub fn new(main_rows: u32, main_cols: u32) -> Self {
        let g = Grid {
            main_cells: HashMap::new(),
            extent_main_rows: main_rows.max(1),
            extent_main_cols: main_cols.max(1),
            left: HashMap::new(),
            right: HashMap::new(),
            max_col_width: 20,
            col_width_overrides: HashMap::new(),
            view_sort_cols: Vec::new(),
            col_all_formats: HashMap::new(),
            col_data_formats: HashMap::new(),
            col_special_formats: HashMap::new(),
            cell_formats: HashMap::new(),
            header: HashMap::new(),
            footer: HashMap::new(),
            spill_followers: HashMap::new(),
            spill_errors: HashMap::new(),
            volatile_seed: 0,
        };
        g
    }

    /// One new main row at the bottom (cursor moving down from the last main row).
    pub fn grow_main_row_at_bottom(&mut self) {
        self.extent_main_rows = self.extent_main_rows.saturating_add(1);
    }

    /// One new main column at the right (cursor moving right from the last sheet column).
    pub fn grow_main_col_at_right(&mut self) {
        let old_main_cols = self.extent_main_cols as usize;
        let new_main_cols = old_main_cols.saturating_add(1);
        self.remap_main_col_layout_for_resize(old_main_cols, new_main_cols);
        self.remap_formats_for_resize(old_main_cols, new_main_cols);
        self.extent_main_cols = new_main_cols as u32;
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
                self.remap_formats_for_resize(old_main_cols, new_main_cols);
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
        grown
    }

    pub fn logical_row_has_content(&self, r: usize) -> bool {
        let hr = HEADER_ROWS;
        if r < hr {
            let row = r as u32;
            return self.header.keys().any(|&(stored_row, _)| stored_row == row);
        }
        if r < hr + self.extent_main_rows as usize {
            let mr = r - hr;
            let mru = mr as u32;
            return self.main_cells.keys().any(|(row, _)| *row == mru)
                || self.left.keys().any(|(row, _)| *row == mru)
                || self.right.keys().any(|(row, _)| *row == mru);
        }
        let fr = r - hr - self.extent_main_rows as usize;
        let fr = fr as u32;
        self.footer.keys().any(|&(stored_row, _)| stored_row == fr)
    }

    pub fn logical_col_has_content(&self, c: usize) -> bool {
        let tc = self.total_cols();
        if c >= tc {
            return false;
        }
        let c_u32 = c as u32;
        if self.header.keys().any(|&(_, col)| col == c_u32) {
            return true;
        }
        let m = MARGIN_COLS;
        let me = m + self.extent_main_cols as usize;
        let data_region_has_content = if c < m {
            self.left.keys().any(|(_, mc)| *mc == c)
        } else if c < me {
            let mc = (c - m) as u32;
            self.main_cells.keys().any(|(_, col)| *col == mc)
        } else if c < me + MARGIN_COLS {
            let mc = c - me;
            self.right.keys().any(|(_, rc)| *rc == mc)
        } else {
            false
        };
        if data_region_has_content {
            return true;
        }
        self.footer.keys().any(|&(_, col)| col == c_u32)
    }

    fn resize_header_footer_width(&mut self) {
        let total_cols = self.total_cols() as u32;
        self.header.retain(|&(row, col), value| {
            row < HEADER_ROWS as u32 && col < total_cols && !value.is_empty()
        });
        self.footer.retain(|&(row, col), value| {
            row < FOOTER_ROWS as u32 && col < total_cols && !value.is_empty()
        });
    }

    pub fn set_main_size(&mut self, main_rows: usize, main_cols: usize) {
        let old_main_cols = self.extent_main_cols as usize;
        let new_main_cols = main_cols.max(1);
        self.remap_main_col_layout_for_resize(old_main_cols, new_main_cols);
        self.remap_formats_for_resize(old_main_cols, new_main_cols);
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

    pub fn content_width_for_column(&self, global_col: usize) -> Option<usize> {
        let mut maxw = 0usize;
        let mut saw_content = false;
        let main_cols = self.main_cols();

        let global_col_u32 = global_col as u32;
        for (&(_, col), val) in &self.header {
            if col == global_col_u32 {
                saw_content = true;
                maxw = maxw.max(val.chars().count() + 1);
            }
        }
        for (&(_, col), val) in &self.footer {
            if col == global_col_u32 {
                saw_content = true;
                maxw = maxw.max(val.chars().count() + 1);
            }
        }
        for r in 0..self.extent_main_rows as usize {
            if global_col < MARGIN_COLS {
                if let Some(val) = self.left.get(&(r as u32, global_col as usize)) {
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
                if let Some(val) = self.right.get(&(r as u32, rc as usize)) {
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

        fn remap_sparse_rows(
            cells: &mut HashMap<(u32, u32), String>,
            new_main_cols: usize,
            old_right_start: usize,
            new_right_start: usize,
            old_total: usize,
            new_total: usize,
        ) {
            let mut remapped = HashMap::new();
            for ((row, col), value) in cells.drain() {
                let old_col = col as usize;
                if old_col >= old_total {
                    continue;
                }
                let Some(new_col) =
                    remap_col(old_col, new_main_cols, old_right_start, new_right_start)
                else {
                    continue;
                };
                if new_col < new_total && !value.is_empty() {
                    remapped.insert((row, new_col as u32), value);
                }
            }
            *cells = remapped;
        }

        remap_sparse_rows(
            &mut self.header,
            new_main_cols,
            old_right_start,
            new_right_start,
            old_total,
            new_total,
        );
        remap_sparse_rows(
            &mut self.footer,
            new_main_cols,
            old_right_start,
            new_right_start,
            old_total,
            new_total,
        );

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

    fn merge_format(base: CellFormat, overlay: CellFormat) -> CellFormat {
        CellFormat {
            number: overlay.number.or(base.number),
            align: overlay.align.or(base.align),
        }
    }

    fn set_scoped_column_format(
        map: &mut HashMap<usize, CellFormat>,
        col: usize,
        format: CellFormat,
    ) {
        if format == CellFormat::default() {
            map.remove(&col);
        } else {
            map.insert(col, format);
        }
    }

    pub fn set_column_format(&mut self, scope: FormatScope, col: usize, format: CellFormat) {
        match scope {
            FormatScope::All => {
                Self::set_scoped_column_format(&mut self.col_all_formats, col, format)
            }
            FormatScope::Data => {
                Self::set_scoped_column_format(&mut self.col_data_formats, col, format)
            }
            FormatScope::Special => {
                Self::set_scoped_column_format(&mut self.col_special_formats, col, format)
            }
        }
    }

    pub fn set_cell_format(&mut self, addr: CellAddr, format: CellFormat) {
        if format == CellFormat::default() {
            self.cell_formats.remove(&addr);
        } else {
            self.cell_formats.insert(addr, format);
        }
    }

    pub fn format_for_addr(&self, addr: &CellAddr) -> CellFormat {
        let global_col = addr_logical_col(addr, self);
        let base = *self
            .col_all_formats
            .get(&global_col)
            .unwrap_or(&CellFormat::default());
        let region = match addr {
            CellAddr::Main { .. } => self.col_data_formats.get(&global_col).copied(),
            _ => self.col_special_formats.get(&global_col).copied(),
        }
        .unwrap_or_default();
        let exact = self.cell_formats.get(addr).copied().unwrap_or_default();
        Self::merge_format(Self::merge_format(base, region), exact)
    }

    pub fn format_for_global_col(&self, scope: FormatScope, col: usize) -> CellFormat {
        match scope {
            FormatScope::All => self.col_all_formats.get(&col).copied().unwrap_or_default(),
            FormatScope::Data => self.col_data_formats.get(&col).copied().unwrap_or_default(),
            FormatScope::Special => self
                .col_special_formats
                .get(&col)
                .copied()
                .unwrap_or_default(),
        }
    }

    pub fn remap_formats_for_resize(&mut self, old_main_cols: usize, new_main_cols: usize) {
        fn remap_map(
            map: &mut HashMap<usize, CellFormat>,
            old_main_cols: usize,
            new_main_cols: usize,
        ) {
            let old_total = MARGIN_COLS + old_main_cols + MARGIN_COLS;
            let new_total = MARGIN_COLS + new_main_cols + MARGIN_COLS;
            let old_right_start = MARGIN_COLS + old_main_cols;
            let new_right_start = MARGIN_COLS + new_main_cols;
            let mut remapped = HashMap::new();
            for (col, fmt) in map.drain() {
                let new_col = if col < MARGIN_COLS {
                    Some(col)
                } else if col < old_right_start {
                    let main_idx = col - MARGIN_COLS;
                    (main_idx < new_main_cols).then_some(MARGIN_COLS + main_idx)
                } else {
                    let right_idx = col - old_right_start;
                    Some(new_right_start + right_idx)
                };
                if let Some(new_col) = new_col {
                    if new_col < new_total && col < old_total {
                        remapped.insert(new_col, fmt);
                    }
                }
            }
            *map = remapped;
        }

        remap_map(&mut self.col_all_formats, old_main_cols, new_main_cols);
        remap_map(&mut self.col_data_formats, old_main_cols, new_main_cols);
        remap_map(&mut self.col_special_formats, old_main_cols, new_main_cols);
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
            CellAddr::Header { row, col } => self.header.get(&(*row, *col)).map(|s| s.as_str()),
            CellAddr::Footer { row, col } => self.footer.get(&(*row, *col)).map(|s| s.as_str()),
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
                let c = *col as usize;
                if (*row as usize) < HEADER_ROWS && c < self.total_cols() {
                    if value.is_empty() {
                        self.header.remove(&(*row, *col));
                    } else {
                        self.header.insert((*row, *col), value);
                        self.auto_fit_column(c);
                    }
                }
            }
            CellAddr::Footer { row, col } => {
                let c = *col as usize;
                if (*row as usize) < FOOTER_ROWS && c < self.total_cols() {
                    if value.is_empty() {
                        self.footer.remove(&(*row, *col));
                    } else {
                        self.footer.insert((*row, *col), value);
                        self.auto_fit_column(c);
                    }
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
                if mc < MARGIN_COLS {
                    if value.is_empty() {
                        self.left.remove(&(r, mc));
                    } else {
                        self.extent_main_rows = self.extent_main_rows.max(r + 1);
                        self.left.insert((r, mc), value);
                        self.auto_fit_column(mc);
                        self.resize_header_footer_width();
                    }
                }
            }
            CellAddr::Right { col, row } => {
                let mc = *col;
                let r = *row;
                if mc < MARGIN_COLS {
                    if value.is_empty() {
                        self.right.remove(&(r, mc));
                    } else {
                        self.extent_main_rows = self.extent_main_rows.max(r + 1);
                        self.right.insert((r, mc), value);
                        self.auto_fit_column(MARGIN_COLS + self.extent_main_cols as usize + mc);
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

    pub(crate) fn bump_volatile_seed(&mut self) {
        self.volatile_seed = self.volatile_seed.wrapping_add(1);
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
            for mc in 0..MARGIN_COLS as usize {
                if let Some(v) = self.left.get(&(old_r, mc)).cloned() {
                    new_left.insert((new_pos as u32, mc), v);
                }
            }
        }
        self.left = new_left;

        let mut new_right = HashMap::new();
        for (new_pos, &old_r) in order.iter().enumerate() {
            for mc in 0..MARGIN_COLS as usize {
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

        fn remap_sparse_main_cols(
            cells: &mut HashMap<(u32, u32), String>,
            order: &[u32],
            old_main_cols: usize,
        ) {
            let mut old_to_new = vec![0usize; old_main_cols];
            for (new_pos, &old_pos) in order.iter().enumerate() {
                old_to_new[old_pos as usize] = new_pos;
            }

            let mut remapped = HashMap::new();
            for ((row, col), value) in cells.drain() {
                let col_usize = col as usize;
                let new_col = if col_usize < MARGIN_COLS || col_usize >= MARGIN_COLS + old_main_cols
                {
                    col_usize
                } else {
                    MARGIN_COLS + old_to_new[col_usize - MARGIN_COLS]
                };
                remapped.insert((row, new_col as u32), value);
            }
            *cells = remapped;
        }

        remap_sparse_main_cols(&mut self.header, &order, ec);
        remap_sparse_main_cols(&mut self.footer, &order, ec);

        self.remap_main_col_width_overrides_for_order(&order);

        self.extent_main_cols = order.len() as u32;
    }
}

// Implement GridImpl for the existing Grid so we can use Grid via GridBox.
impl GridImpl for Grid {
    fn main_rows(&self) -> usize {
        self.main_rows()
    }

    fn main_cols(&self) -> usize {
        self.main_cols()
    }

    fn total_cols(&self) -> usize {
        self.total_cols()
    }

    fn get_owned(&self, addr: &CellAddr) -> Option<String> {
        self.get(addr).map(|s| s.to_string())
    }

    fn set_owned(&mut self, addr: &CellAddr, value: String) {
        self.set(addr, value)
    }

    fn set_main_size(&mut self, main_rows: usize, main_cols: usize) {
        self.set_main_size(main_rows, main_cols)
    }

    fn bump_volatile_seed(&mut self) {
        self.bump_volatile_seed()
    }

    fn iter_nonempty(&self) -> Box<dyn Iterator<Item = (CellAddr, String)> + '_> {
        // Build a vec of non-empty cells across regions and return an iterator.
        let mut v: Vec<(CellAddr, String)> = Vec::new();
        for (&(r, c), val) in &self.header {
            v.push((CellAddr::Header { row: r, col: c }, val.clone()));
        }
        for (&(r, c), val) in &self.footer {
            v.push((CellAddr::Footer { row: r, col: c }, val.clone()));
        }
        for (&(r, c), val) in &self.main_cells {
            v.push((CellAddr::Main { row: r, col: c }, val.clone()));
        }
        for (&(r, mc), val) in &self.left {
            v.push((CellAddr::Left { col: mc, row: r }, val.clone()));
        }
        for (&(r, mc), val) in &self.right {
            v.push((CellAddr::Right { col: mc, row: r }, val.clone()));
        }
        Box::new(v.into_iter())
    }

    fn total_logical_rows(&self) -> usize {
        self.total_logical_rows()
    }

    fn text(&self, addr: &CellAddr) -> String {
        self.get(addr).unwrap_or("").to_string()
    }

    fn set(&mut self, addr: &CellAddr, value: String) {
        self.set(addr, value)
    }

    fn ensure_extent_for_cursor(&mut self, row: usize, col: usize) -> bool {
        self.ensure_extent_for_cursor(row, col)
    }

    fn grow_main_row_at_bottom(&mut self) {
        self.grow_main_row_at_bottom()
    }

    fn grow_main_col_at_right(&mut self) {
        self.grow_main_col_at_right()
    }

    fn move_main_rows(&mut self, from: usize, count: usize, to: usize) {
        self.move_main_rows(from, count, to)
    }

    fn move_main_cols(&mut self, from: usize, count: usize, to: usize) {
        self.move_main_cols(from, count, to)
    }

    fn max_col_width(&self) -> usize {
        self.max_col_width
    }

    fn col_width(&self, global_col: usize) -> usize {
        self.col_width(global_col)
    }

    fn get_col_width_override(&self, global_col: usize) -> Option<usize> {
        self.col_width_overrides.get(&global_col).copied()
    }

    fn content_width_for_column(&self, global_col: usize) -> Option<usize> {
        self.content_width_for_column(global_col)
    }

    fn set_max_col_width(&mut self, width: usize) {
        self.set_max_col_width(width)
    }

    fn set_col_width(&mut self, global_col: usize, width: Option<usize>) {
        self.set_col_width(global_col, width)
    }

    fn auto_fit_column(&mut self, global_col: usize) {
        self.auto_fit_column(global_col)
    }

    fn fit_column_to_content(&mut self, global_col: usize) {
        self.fit_column_to_content(global_col)
    }

    fn col_width_overrides(&self) -> Vec<(usize, usize)> {
        self.col_width_overrides
            .iter()
            .map(|(&c, &w)| (c, w))
            .collect()
    }

    fn set_view_sort_cols(&mut self, cols: Vec<SortSpec>) {
        self.set_view_sort_cols(cols)
    }

    fn view_sort_cols(&self) -> Vec<SortSpec> {
        self.view_sort_cols.clone()
    }

    fn sorted_main_rows(&self) -> Vec<usize> {
        self.sorted_main_rows()
    }

    fn set_column_format(&mut self, scope: FormatScope, col: usize, format: CellFormat) {
        self.set_column_format(scope, col, format)
    }

    fn set_cell_format(&mut self, addr: CellAddr, format: CellFormat) {
        self.set_cell_format(addr, format)
    }

    fn format_for_addr(&self, addr: &CellAddr) -> CellFormat {
        self.format_for_addr(addr)
    }

    fn format_for_global_col(&self, scope: FormatScope, col: usize) -> CellFormat {
        self.format_for_global_col(scope, col)
    }

    fn col_all_formats(&self) -> Vec<(usize, CellFormat)> {
        self.col_all_formats.iter().map(|(&c, &f)| (c, f)).collect()
    }

    fn col_data_formats(&self) -> Vec<(usize, CellFormat)> {
        self.col_data_formats
            .iter()
            .map(|(&c, &f)| (c, f))
            .collect()
    }

    fn col_special_formats(&self) -> Vec<(usize, CellFormat)> {
        self.col_special_formats
            .iter()
            .map(|(&c, &f)| (c, f))
            .collect()
    }

    fn cell_formats(&self) -> Vec<(CellAddr, CellFormat)> {
        self.cell_formats
            .iter()
            .map(|(k, &v)| (k.clone(), v))
            .collect()
    }

    fn clear_spills(&mut self) {
        self.clear_spills()
    }

    fn spill_followers(&self) -> Vec<(CellAddr, String)> {
        self.spill_followers
            .iter()
            .map(|(k, v)| (k.clone(), v.clone()))
            .collect()
    }

    fn spill_errors(&self) -> Vec<(CellAddr, &'static str)> {
        self.spill_errors
            .iter()
            .map(|(k, &v)| (k.clone(), v))
            .collect()
    }

    fn clear_cells(&mut self) {
        self.main_cells.clear();
        self.left.clear();
        self.right.clear();
    }

    fn set_col_width_overrides(&mut self, overrides: Vec<(usize, usize)>) {
        self.col_width_overrides = overrides.into_iter().collect();
    }

    fn set_spill_value(&mut self, addr: CellAddr, value: String) {
        self.set_spill_value(addr, value)
    }

    fn set_spill_error(&mut self, addr: CellAddr, err: &'static str) {
        self.set_spill_error(addr, err)
    }

    fn spill_error(&self, addr: &CellAddr) -> Option<&'static str> {
        self.spill_error(addr)
    }

    fn volatile_seed(&self) -> u64 {
        self.volatile_seed
    }

    fn set_volatile_seed(&mut self, seed: u64) {
        self.volatile_seed = seed;
    }

    fn logical_row_has_content(&self, r: usize) -> bool {
        self.logical_row_has_content(r)
    }

    fn logical_col_has_content(&self, c: usize) -> bool {
        self.logical_col_has_content(c)
    }

    fn clone_box(&self) -> Box<dyn GridImpl> {
        Box::new(self.clone())
    }
}

impl From<Grid> for GridBox {
    fn from(g: Grid) -> Self {
        GridBox::new(g)
    }
}

impl Clone for GridBox {
    fn clone(&self) -> Self {
        GridBox {
            inner: self.inner.clone_box(),
        }
    }
}

impl std::fmt::Debug for GridBox {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("GridBox").finish()
    }
}

impl Default for GridBox {
    fn default() -> Self {
        GridBox::new(Grid::new(1, 1))
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
        CellAddr::Header { col, .. } | CellAddr::Footer { col, .. } => *col as usize,
        CellAddr::Main { col, .. } => MARGIN_COLS + *col as usize,
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

    #[test]
    fn header_footer_rows_are_sparse_at_high_limits() {
        let mut g = Grid::new(1, 1);
        let header = CellAddr::Header {
            row: 0,
            col: MARGIN_COLS as u32,
        };
        let footer = CellAddr::Footer {
            row: (FOOTER_ROWS - 1) as u32,
            col: MARGIN_COLS as u32,
        };

        g.set(&header, "top".into());
        g.set(&footer, "bottom".into());

        assert_eq!(g.header.len(), 1);
        assert_eq!(g.footer.len(), 1);
        assert_eq!(g.get(&header), Some("top"));
        assert_eq!(g.get(&footer), Some("bottom"));
        assert!(g.logical_row_has_content(0));
        assert!(g.logical_row_has_content(HEADER_ROWS + g.main_rows() + FOOTER_ROWS - 1));

        g.set(&header, String::new());
        assert!(g.header.is_empty());
    }

    #[test]
    fn format_scope_merges_by_region() {
        let mut g = Grid::new(1, 1);
        g.set_column_format(
            FormatScope::All,
            MARGIN_COLS,
            CellFormat {
                number: Some(NumberFormat::Fixed { decimals: 2 }),
                align: None,
            },
        );
        g.set_column_format(
            FormatScope::Data,
            MARGIN_COLS,
            CellFormat {
                number: None,
                align: Some(TextAlign::Right),
            },
        );
        g.set_cell_format(
            CellAddr::Main { row: 0, col: 0 },
            CellFormat {
                number: None,
                align: Some(TextAlign::Center),
            },
        );

        let fmt = g.format_for_addr(&CellAddr::Main { row: 0, col: 0 });
        assert_eq!(fmt.number, Some(NumberFormat::Fixed { decimals: 2 }));
        assert_eq!(fmt.align, Some(TextAlign::Center));
    }
}
