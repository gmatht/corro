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
}
