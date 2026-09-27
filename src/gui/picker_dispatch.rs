//! One navigation contract for the GUI list pickers (aggregate, special
//! char, Balance Books).
//!
//! The three pickers are separate state machines (different rows, different
//! commit payloads), but every backend drives them through the *same* key
//! vocabulary: Up/Left and Down/Right step, digits jump-and-commit, Enter
//! commits, Esc cancels, anything else is ignored. Writing that mapping once
//! here means a backend's key hook is a thin `match` over the picker
//! currently open, and the arrow/digit semantics cannot drift between the
//! pancurses and GTK backends or between the three pickers.
//!
//! The *commit* action stays with the caller: each picker commits a
//! different payload through a different action (a cell directive, a splice
//! into the editor, a Balance Books run), and on pancurses the commit also
//! closes the toolkit's list widget. So this module owns only the selection
//! machine and the key→outcome mapping; the caller supplies what to do on a
//! commit.
//!
//! Pure over [`super::App`]: no widget, no toolkit types, so it unit-tests
//! without a display.

use super::App;

/// A picker whose selection is an index into a fixed list of rows.
pub trait IndexedPicker {
    /// Current selection index (`None` = the picker is closed).
    fn index(app: &App) -> Option<usize>;
    /// One clamped arrow step (`delta` +1 / -1). No-op while closed.
    fn step(app: &mut App, delta: i32);
    /// Absolute selection (widget → state sync), clamped.
    fn set(app: &mut App, idx: usize);
    /// Close without committing.
    fn close(app: &mut App);
    /// Digit hotkey → row index, or `None` when the digit selects nothing.
    fn index_for_digit(digit: char) -> Option<usize>;
}

/// Which picker is currently open.
///
/// Exactly one is ever open (they are never nested — see the pnc/GTK key
/// hooks), so a single first-match lookup routes a key.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum OpenPicker {
    Agg,
    Special,
    Balance,
}

/// The picker currently open, if any.
pub fn open_picker(app: &App) -> Option<OpenPicker> {
    if app.agg_picker.is_some() {
        Some(OpenPicker::Agg)
    } else if app.special_picker.is_some() {
        Some(OpenPicker::Special)
    } else if app.balance_picker.is_some() {
        Some(OpenPicker::Balance)
    } else {
        None
    }
}

/// A toolkit-agnostic key as the pickers understand it.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum PickerKeyInput {
    /// Up/Left: step toward the top.
    Prev,
    /// Down/Right: step toward the bottom.
    Next,
    /// Enter.
    Enter,
    /// Esc.
    Escape,
    /// A character; digits are the picker hotkeys.
    Char(char),
    /// Any other key (ignored by the pickers).
    Other,
}

/// What a key did to a picker, so a backend knows whether to swallow it and
/// whether to close its list widget.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum PickerKey {
    /// The key was not for this picker; the backend should not swallow it.
    Ignored,
    /// The selection moved; repaint the list (still open).
    Stepped,
    /// A digit jumped to a row and the selection should be committed.
    Committed,
    /// Enter: commit the current selection.
    Commit,
    /// Esc: cancel and close (already closed in state).
    Cancel,
}

/// Dispatch `key` to `P`, returning what happened.
///
/// `Committed`/`Commit` mean the caller should now commit the selection
/// (each picker's `take` closes its own state); `Cancel`/`Committed`/`Commit`
/// also mean the caller should dismiss its list widget. `Ignored` means no
/// picker state was touched.
pub fn handle_picker_key<P: IndexedPicker>(app: &mut App, key: PickerKeyInput) -> PickerKey {
    match key {
        PickerKeyInput::Prev => {
            P::step(app, -1);
            PickerKey::Stepped
        }
        PickerKeyInput::Next => {
            P::step(app, 1);
            PickerKey::Stepped
        }
        PickerKeyInput::Enter => PickerKey::Commit,
        PickerKeyInput::Escape => {
            P::close(app);
            PickerKey::Cancel
        }
        PickerKeyInput::Char(c) => match P::index_for_digit(c) {
            Some(idx) => {
                P::set(app, idx);
                PickerKey::Committed
            }
            None => PickerKey::Ignored,
        },
        PickerKeyInput::Other => PickerKey::Ignored,
    }
}

/// Marker types for the three pickers, so a caller can select one without
/// naming a module path in the dispatch call.
pub struct AggPicker;
pub struct SpecialPicker;
pub struct BalancePicker;

impl IndexedPicker for AggPicker {
    fn index(app: &App) -> Option<usize> {
        super::agg_picker::index(app)
    }
    fn step(app: &mut App, delta: i32) {
        super::agg_picker::step(app, delta)
    }
    fn set(app: &mut App, idx: usize) {
        super::agg_picker::set(app, idx)
    }
    fn close(app: &mut App) {
        super::agg_picker::close(app)
    }
    fn index_for_digit(digit: char) -> Option<usize> {
        super::agg_picker::index_for_digit(digit)
    }
}

impl IndexedPicker for SpecialPicker {
    fn index(app: &App) -> Option<usize> {
        super::special_picker::index(app)
    }
    fn step(app: &mut App, delta: i32) {
        super::special_picker::step(app, delta)
    }
    fn set(app: &mut App, idx: usize) {
        super::special_picker::set(app, idx)
    }
    fn close(app: &mut App) {
        super::special_picker::close(app)
    }
    fn index_for_digit(digit: char) -> Option<usize> {
        super::special_picker::index_for_digit(digit)
    }
}

impl IndexedPicker for BalancePicker {
    fn index(app: &App) -> Option<usize> {
        super::balance_picker::index(app)
    }
    fn step(app: &mut App, delta: i32) {
        super::balance_picker::step(app, delta)
    }
    fn set(app: &mut App, idx: usize) {
        super::balance_picker::set(app, idx)
    }
    fn close(app: &mut App) {
        super::balance_picker::close(app)
    }
    fn index_for_digit(digit: char) -> Option<usize> {
        super::balance_picker::index_for_digit(digit)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    

    fn app_with_agg() -> App {
        let mut app = App::new_with_paths(vec![]);
        crate::gui::agg_picker::open_for_cursor(
            &mut app,
            &crate::grid::SheetCursor {
                row: crate::grid::HEADER_ROWS,
                col: crate::grid::MARGIN_COLS,
            },
        );
        app
    }

    #[test]
    fn open_picker_reports_the_one_that_is_open() {
        let mut app = App::new_with_paths(vec![]);
        assert_eq!(open_picker(&app), None);
        crate::gui::special_picker::open(&mut app);
        assert_eq!(open_picker(&app), Some(OpenPicker::Special));
        crate::gui::special_picker::close(&mut app);
        crate::gui::balance_picker::open(&mut app);
        assert_eq!(open_picker(&app), Some(OpenPicker::Balance));
    }

    #[test]
    fn arrows_step_and_clamp_through_the_shared_mapping() {
        let mut app = app_with_agg();
        let start = AggPicker::index(&app).expect("open");
        assert_eq!(
            handle_picker_key::<AggPicker>(&mut app, PickerKeyInput::Next),
            PickerKey::Stepped
        );
        assert_eq!(AggPicker::index(&app), Some(start + 1));
        // Up from the top clamps rather than wrapping or erroring.
        for _ in 0..10 {
            handle_picker_key::<AggPicker>(&mut app, PickerKeyInput::Prev);
        }
        assert_eq!(AggPicker::index(&app), Some(0));
    }

    #[test]
    fn digits_jump_and_ask_for_a_commit() {
        let mut app = app_with_agg();
        assert_eq!(
            handle_picker_key::<AggPicker>(&mut app, PickerKeyInput::Char('3')),
            PickerKey::Committed
        );
        assert_eq!(AggPicker::index(&app), Some(2));
        // A digit with no row is ignored, leaving the selection alone.
        assert_eq!(
            handle_picker_key::<AggPicker>(&mut app, PickerKeyInput::Char('9')),
            PickerKey::Ignored
        );
        assert_eq!(AggPicker::index(&app), Some(2));
    }

    #[test]
    fn enter_asks_for_a_commit_and_escape_cancels() {
        let mut app = app_with_agg();
        assert_eq!(
            handle_picker_key::<AggPicker>(&mut app, PickerKeyInput::Enter),
            PickerKey::Commit
        );
        // Enter must NOT close: the caller's commit path does the take.
        assert!(AggPicker::index(&app).is_some());
        assert_eq!(
            handle_picker_key::<AggPicker>(&mut app, PickerKeyInput::Escape),
            PickerKey::Cancel
        );
        assert_eq!(AggPicker::index(&app), None, "Esc closes the picker");
    }

    #[test]
    fn other_keys_are_ignored() {
        let mut app = app_with_agg();
        assert_eq!(
            handle_picker_key::<AggPicker>(&mut app, PickerKeyInput::Other),
            PickerKey::Ignored
        );
        assert_eq!(
            handle_picker_key::<AggPicker>(&mut app, PickerKeyInput::Char('x')),
            PickerKey::Ignored
        );
    }
}
