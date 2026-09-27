//! Shared Insert > Special Char picker state (GUI + pancurses backends).
//!
//! The ratatui reference owns an equivalent picker inline (`ui::App`);
//! this module is the backend-agnostic selection machine every other picker
//! UI drives, so arrow/digit/Enter/Esc semantics can never drift between
//! backends. Backends are thin renderers: they display
//! [`items`], read [`index`], and call [`open`] / [`step`] / [`set`] /
//! [`close`] / [`take`] for keys and dialog buttons. Committing returns the
//! bare choice string; splicing it into edit state is the backend's job
//! (ratatui splices at its edit caret; the GUI appends to its formula-bar
//! buffer; pancurses splices at its widget caret).
//!
//! Pure logic over [`crate::gui::App`]; unit-tested below without any
//! display.

use super::App;
use crate::ui_core::{
    special_choice_index_for_digit, special_step_index, SPECIAL_VALUE_CHOICES,
};

/// List-widget rows every picker renders (`"1: ∞"` … `"0: θ"`).
pub fn items() -> [String; 10] {
    crate::ui_core::special_labelled_choices()
}

/// Current selection index (`None` = picker closed).
pub fn index(app: &App) -> Option<usize> {
    app.special_picker
}

/// Open the picker on the first choice (Down*n lands on the nth item).
pub fn open(app: &mut App) {
    app.special_picker = Some(0);
}

/// Close without committing.
pub fn close(app: &mut App) {
    app.special_picker = None;
}

/// One arrow step (`delta` +1 for Down/Right, -1 for Up/Left), clamped to
/// the choice range. No-op while closed.
pub fn step(app: &mut App, delta: i32) {
    if let Some(idx) = app.special_picker {
        app.special_picker = Some(special_step_index(idx, delta));
    }
}

/// Absolute selection (widget → state sync), clamped. No-op while closed.
pub fn set(app: &mut App, idx: usize) {
    if app.special_picker.is_some() {
        app.special_picker = Some(idx.min(SPECIAL_VALUE_CHOICES.len() - 1));
    }
}

/// Commit the current selection: returns the choice and closes. `None`
/// while closed.
pub fn take(app: &mut App) -> Option<String> {
    let idx = app.special_picker.take()?;
    Some(SPECIAL_VALUE_CHOICES[idx].to_string())
}

/// Digit hotkey → choice index (`1`..=`9` → 0..=8, `0` → 9).
pub fn index_for_digit(digit: char) -> Option<usize> {
    special_choice_index_for_digit(digit)
}

/// Where a picked special character should be spliced, as the *resulting*
/// `(text, caret)` a backend then writes into its own editor widget.
///
/// The picker never commits the edit: it splices the choice into the
/// in-progress text and leaves the user to commit. Both GUIs agree on the
/// two cases, which differ only in where the base text comes from:
///
/// * **Already editing** — splice at the caret inside `edit_text` (the
///   widget's live buffer), keeping the caret after the choice.
/// * **Not editing** — start from the cursor cell's display text and append
///   the choice (caret at the end), i.e. a seeded edit of `cell + choice`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SpecialSplice {
    /// The full text after splicing (what the editor widget should show).
    pub text: String,
    /// Caret position after the splice, as a **char** index.
    pub caret: usize,
}

impl SpecialSplice {
    /// Splice `choice` into the editor's live buffer at `caret`
    /// (already-editing case). `caret`/`edit_text` are char-based, matching
    /// the shared text-edit model, so multibyte choices never split a
    /// codepoint.
    pub fn into_edit(edit_text: &str, caret: usize, choice: &str) -> Self {
        let mut text = edit_text.to_string();
        let caret = crate::ui_core::insert_str_at_char(&mut text, caret, choice);
        SpecialSplice { text, caret }
    }

    /// Seed a fresh edit from the cursor cell's display `cell_text` and
    /// append `choice` (not-editing case).
    pub fn into_cell(cell_text: &str, choice: &str) -> Self {
        let mut text = cell_text.to_string();
        let end = text.chars().count();
        let caret = crate::ui_core::insert_str_at_char(&mut text, end, choice);
        SpecialSplice { text, caret }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::gui::App as GuiApp;

    fn app() -> GuiApp {
        GuiApp::new_with_paths(vec![])
    }

    #[test]
    fn rows_match_ratatui_labels_in_order() {
        let rows = items();
        assert_eq!(rows.len(), 10);
        assert_eq!(rows[0], "1: ∞");
        assert_eq!(rows[2], "3: Ω");
        assert_eq!(rows[9], "0: θ");
    }

    #[test]
    fn open_step_take_roundtrip() {
        let mut a = app();
        assert_eq!(index(&a), None);
        open(&mut a);
        assert_eq!(index(&a), Some(0));
        step(&mut a, 1);
        step(&mut a, 1);
        assert_eq!(index(&a), Some(2));
        assert_eq!(take(&mut a), Some("Ω".to_string()));
        assert_eq!(index(&a), None, "take closes the picker");
    }

    #[test]
    fn steps_clamp_at_both_ends() {
        let mut a = app();
        open(&mut a);
        step(&mut a, -1);
        assert_eq!(index(&a), Some(0));
        step(&mut a, 25);
        assert_eq!(index(&a), Some(9));
        step(&mut a, 1);
        assert_eq!(index(&a), Some(9));
        step(&mut a, -25);
        assert_eq!(index(&a), Some(0));
    }

    #[test]
    fn closed_picker_ignores_keys() {
        let mut a = app();
        step(&mut a, 1);
        set(&mut a, 5);
        assert_eq!(take(&mut a), None);
        assert_eq!(index(&a), None);
    }

    #[test]
    fn set_clamps_and_digits_map() {
        let mut a = app();
        open(&mut a);
        set(&mut a, 99);
        assert_eq!(index(&a), Some(9));
        assert_eq!(index_for_digit('1'), Some(0));
        assert_eq!(index_for_digit('3'), Some(2));
        assert_eq!(index_for_digit('0'), Some(9));
        assert_eq!(index_for_digit('x'), None);
        close(&mut a);
        assert_eq!(index(&a), None);
    }

    /// Splicing into an in-progress edit inserts at the caret, not at the
    /// end — the same caret-awareness typing has.
    #[test]
    fn splice_into_edit_inserts_at_the_caret() {
        // "ab" with the caret between a and b, splicing "X" -> "aXb", caret 2.
        let s = SpecialSplice::into_edit("ab", 1, "X");
        assert_eq!(s.text, "aXb");
        assert_eq!(s.caret, 2);
        // The caret is a *char* index: a multibyte choice must not split.
        let s = SpecialSplice::into_edit("aθb", 2, "Ω");
        assert_eq!(s.text, "aθΩb");
        assert_eq!(s.caret, 3);
        assert!(s.text.chars().count() == 4);
        // A caret past the end clamps (append), never panics.
        let s = SpecialSplice::into_edit("ab", 99, "X");
        assert_eq!(s.text, "abX");
        assert_eq!(s.caret, 3);
    }

    /// Not editing: the edit is seeded from the cell's text plus the choice,
    /// with the caret after it.
    #[test]
    fn splice_into_cell_appends_the_choice() {
        let s = SpecialSplice::into_cell("=1+2", "Ω");
        assert_eq!(s.text, "=1+2Ω");
        assert_eq!(s.caret, 5);
        let s = SpecialSplice::into_cell("", "∞");
        assert_eq!(s.text, "∞");
        assert_eq!(s.caret, 1);
    }
}
