//! Shared row/column freeze ("padlock") state.
//!
//! One [`LockState`] per app so the mouse, the menu, and every terminal
//! backend read the same sets instead of each keeping its own copy. The GUI
//! used to hold these on its widget state and toggle them from a click
//! handler only; moving them here is what lets `Sheet ▸ Col-Lock` and
//! `Sheet ▸ Row-Lock` do the same thing without a mouse, and lets the TUIs
//! draw the same affordance the GTK backend draws.
//!
//! Session-only, never persisted: the sets are view state like the scroll
//! offsets, not document content, so they are cleared when the active sheet
//! changes ([`LockState::ensure_sheet`]) and never written to the log.

use crate::grid::{GridBox, MARGIN_COLS};
use std::collections::BTreeSet;

/// Frozen rows/columns for the active sheet.
#[derive(Clone, Debug, Default)]
pub struct LockState {
    /// Logical main rows pinned visible while scrolling.
    pub rows: BTreeSet<usize>,
    /// *Global* column indices pinned visible while scrolling. Global rather
    /// than main-relative because that is what the gutter headers address:
    /// the `[A` / `]A` margin columns are lockable too.
    pub cols: BTreeSet<usize>,
    /// Sheet id the sets belong to. Compared on every read so switching sheets
    /// drops the old pins lazily instead of each call site having to remember.
    pub sheet_id: u32,
}

impl LockState {
    /// Fresh, empty state with no sheet claimed yet.
    ///
    /// `sheet_id` starts at 0, which is deliberately *not* a valid sheet id, so
    /// the first [`LockState::ensure_sheet`] against a real sheet always takes
    /// the clearing branch. That is safe (the sets are empty anyway) but it
    /// means a caller must `ensure_sheet` BEFORE the first
    /// [`LockState::toggle_col`] / [`LockState::toggle_row`]: toggling first
    /// and reading later has the pin wiped by that first read, so the lock
    /// silently does nothing.
    pub fn new() -> Self {
        Self::default()
    }

    /// Drop the pins if the active sheet changed, then return the sets.
    ///
    /// Callers read pins through this rather than touching the fields so a
    /// sheet switch can never leave one backend showing the previous sheet's
    /// frozen columns.
    pub fn ensure_sheet(&mut self, sheet_id: u32) -> (&BTreeSet<usize>, &BTreeSet<usize>) {
        if self.sheet_id != sheet_id {
            self.rows.clear();
            self.cols.clear();
            self.sheet_id = sheet_id;
        }
        (&self.rows, &self.cols)
    }

    /// Toggle `global_col`; returns true when it ends up locked.
    pub fn toggle_col(&mut self, global_col: usize) -> bool {
        toggle_in(&mut self.cols, global_col)
    }

    /// Toggle `logical_row`; returns true when it ends up locked.
    pub fn toggle_row(&mut self, logical_row: usize) -> bool {
        toggle_in(&mut self.rows, logical_row)
    }

    pub fn col_locked(&self, global_col: usize) -> bool {
        self.cols.contains(&global_col)
    }

    pub fn row_locked(&self, logical_row: usize) -> bool {
        self.rows.contains(&logical_row)
    }

    /// Whether anything is frozen, so a backend can skip the whole padlock
    /// path on a freshly opened document.
    pub fn is_empty(&self) -> bool {
        self.rows.is_empty() && self.cols.is_empty()
    }

    /// Pinned rows in ascending order (rendering order).
    pub fn pinned_rows(&self) -> Vec<usize> {
        self.rows.iter().copied().collect()
    }

    /// Pinned global columns in ascending order (rendering order).
    pub fn pinned_cols(&self) -> Vec<usize> {
        self.cols.iter().copied().collect()
    }
}

/// Toggle `idx` in a set; returns true when the index ends up pinned.
pub fn toggle_in(set: &mut BTreeSet<usize>, idx: usize) -> bool {
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
pub fn union_pinned(display: &[usize], pinned: &[usize]) -> Vec<usize> {
    let mut out: Vec<usize> = pinned.to_vec();
    out.extend(display.iter().filter(|r| !pinned.contains(r)).copied());
    out
}

/// Padlock eligibility: gutter labels with short text only (dual-character
/// or less, non-empty) get the affordance — the auto-generated margin
/// labels (`1`, `_1`, `[A`, `AA`, ...) rather than long content.
pub fn wants_padlock(label: &str) -> bool {
    !label.is_empty() && label.chars().count() <= 2
}

/// The (un)locked padlock a terminal backend draws, right-aligned in the
/// gutter header.
///
/// One function so ratatui and pancurses cannot disagree on the character.
/// Both are East-Asian *Wide*, so they occupy two cells — see
/// [`lock_glyph_width`].
pub const LOCKED_GLYPH: char = '\u{1F512}';
/// See [`LOCKED_GLYPH`].
pub const UNLOCKED_GLYPH: char = '\u{1F513}';

/// The padlock character for a locked/unlocked gutter, so a column always
/// carries the affordance and only its *state* changes — locking never
/// changes a header's width, and the two hosts read the same glyph.
pub fn glyph(locked: bool) -> char {
    if locked { LOCKED_GLYPH } else { UNLOCKED_GLYPH }
}

/// Display cells the padlock reserves in a gutter header.
///
/// One cell, not the two `UnicodeWidthChar` reports for the emoji: a terminal
/// renders it in a single cell when it can, and reserving two would silently
/// steal a character of every data cell in that column. This is the same
/// one-cell reservation the GTK backend makes for its vector padlock.
pub const LOCK_GLYPH_CELLS: usize = 1;

/// Columns whose header shows a padlock, by the gutter label rule. Mirrors
/// the GTK eligibility check so both hosts offer a lock on exactly the same
/// columns.
pub fn lock_glyph_width(label: &str) -> usize {
    usize::from(wants_padlock(label))
}

/// Rendered width of a column in the terminal backends: the stored content
/// width plus the padlock cell, exactly as the GTK backend's
/// `display_col_width` does.
///
/// Every terminal layout site must go through this. The header, the separator
/// rule, and each data row derive their geometry from the column width
/// independently, so reserving the padlock in only some of them fans the
/// columns apart.
pub fn tui_col_width(grid: &GridBox, global_col: usize) -> usize {
    let label = crate::addr::ui_column_fragment(global_col, grid.main_cols());
    grid.col_width(global_col).max(1) + lock_glyph_width(&label)
}

/// Whether `global_col` is in the left or right margin, where the gutter
/// labels are the bracketed mirror names and a padlock is still meaningful.
pub fn is_margin_col(global_col: usize) -> bool {
    global_col < MARGIN_COLS
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn toggle_reports_final_state() {
        let mut set = BTreeSet::new();
        assert!(toggle_in(&mut set, 3));
        assert!(set.contains(&3));
        assert!(!toggle_in(&mut set, 3));
        assert!(!set.contains(&3));
    }

    #[test]
    fn ensure_sheet_clears_on_change_but_not_on_repeat() {
        let mut st = LockState::new();
        st.sheet_id = 7;
        st.toggle_col(2);
        st.toggle_row(1);
        let (rows, cols) = st.ensure_sheet(7);
        assert_eq!(cols.len(), 1);
        assert_eq!(rows.len(), 1);
        let (rows, cols) = st.ensure_sheet(8);
        assert!(cols.is_empty() && rows.is_empty());
    }

    #[test]
    fn union_pinned_freezes_ahead_and_dedups() {
        // Column 1 is inside the window already: it must not appear twice.
        let out = union_pinned(&[0, 1, 2, 3], &[1, 9]);
        assert_eq!(out, vec![1, 9, 0, 2, 3]);
    }

    #[test]
    fn union_pinned_preserves_order_with_no_pins() {
        let out = union_pinned(&[0, 1, 2], &[]);
        assert_eq!(out, vec![0, 1, 2]);
    }

    #[test]
    fn wants_padlock_covers_short_labels_only() {
        assert!(wants_padlock("A"));
        assert!(wants_padlock("AA"));
        assert!(wants_padlock("[A"));
        assert!(!wants_padlock(""));
        assert!(!wants_padlock("ABC"));
        assert!(!wants_padlock("Total"));
    }

    #[test]
    fn glyph_flips_between_locked_and_open() {
        assert_eq!(glyph(true), LOCKED_GLYPH);
        assert_eq!(glyph(false), UNLOCKED_GLYPH);
        assert_ne!(LOCKED_GLYPH, UNLOCKED_GLYPH);
    }
}
