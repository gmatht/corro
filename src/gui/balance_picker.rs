//! Shared Sheet > Balance Books picker state (GUI + pancurses backends).
//!
//! The ratatui reference owns an equivalent three-field dialog inline
//! (`ui::App`, `Mode::BalanceBooks`): a column, a report type (view-only vs
//! persisted) and a direction. The single-field text prompt the non-ratatui
//! backends used to offer silently fixed the last two (`PosToNeg`,
//! persisted), so those choices were unreachable off the terminal.
//!
//! This module is the backend-agnostic decision machine: a flat list of the
//! report-type/direction combinations the dialog exposes. Backends render
//! [`items`], read [`index`] and call [`open`] / [`step`] / [`set`] /
//! [`close`] / [`take`] for keys and dialog buttons, then hand the chosen
//! [`BalancePick`] to [`super::actions::run_balance_books`]. The column is
//! taken from the caller (cursor cell / auto-detected numeric column), the
//! same resolution the old prompt used.
//!
//! Pure logic over [`crate::gui::App`]; unit-tested below without any
//! display.

use super::App;
use crate::balance::BalanceDirection;

/// One row of the picker: a (report type, direction) pairing, mirroring the
/// two independent toggles in the ratatui dialog (`persist` × `direction`).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct BalancePick {
    pub persist: bool,
    pub direction: BalanceDirection,
}

/// The picker rows in display order. View-only is the TUI default
/// (`persist: false`) and `PosToNeg` is the default direction, so the
/// defaults come first.
pub const PICKS: [BalancePick; 4] = [
    BalancePick { persist: false, direction: BalanceDirection::PosToNeg },
    BalancePick { persist: false, direction: BalanceDirection::NegToPos },
    BalancePick { persist: true, direction: BalanceDirection::PosToNeg },
    BalancePick { persist: true, direction: BalanceDirection::NegToPos },
];

/// List-widget rows every picker renders.
pub fn items() -> [String; 4] {
    [
        "View only, match +ve with multiple -ve".to_string(),
        "View only, match -ve with multiple +ve".to_string(),
        "Persisted report, match +ve with multiple -ve".to_string(),
        "Persisted report, match -ve with multiple +ve".to_string(),
    ]
}

/// Current selection index (`None` = picker closed).
pub fn index(app: &App) -> Option<usize> {
    app.balance_picker
}

/// Open the picker on the default choice (view-only, `PosToNeg`) — the same
/// defaults the ratatui dialog opens with.
pub fn open(app: &mut App) {
    app.balance_picker = Some(0);
}

/// Close without committing.
pub fn close(app: &mut App) {
    app.balance_picker = None;
}

/// One arrow step (`delta` +1 for Down/Right, -1 for Up/Left), clamped to
/// the row range. No-op while closed.
pub fn step(app: &mut App, delta: i32) {
    if let Some(idx) = app.balance_picker {
        let last = PICKS.len() - 1;
        let next = if delta < 0 {
            idx.saturating_sub(delta.unsigned_abs() as usize)
        } else {
            (idx + delta as usize).min(last)
        };
        app.balance_picker = Some(next);
    }
}

/// Absolute selection (widget → state sync), clamped. No-op while closed.
pub fn set(app: &mut App, idx: usize) {
    if app.balance_picker.is_some() {
        app.balance_picker = Some(idx.min(PICKS.len() - 1));
    }
}

/// Commit the current selection: returns the chosen combination and closes.
/// `None` while closed.
pub fn take(app: &mut App) -> Option<BalancePick> {
    let idx = app.balance_picker.take()?;
    Some(PICKS[idx])
}

/// Digit hotkey → row index (`1`..=`4` → 0..=3).
pub fn index_for_digit(digit: char) -> Option<usize> {
    digit.to_digit(10).and_then(|d| {
        if (1..=4).contains(&d) {
            Some((d - 1) as usize)
        } else {
            None
        }
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::gui::App as GuiApp;

    fn app() -> GuiApp {
        GuiApp::new_with_paths(vec![])
    }

    #[test]
    fn rows_and_picks_agree() {
        assert_eq!(items().len(), PICKS.len());
        assert_eq!(PICKS[0].persist, false);
        assert_eq!(PICKS[0].direction, BalanceDirection::PosToNeg);
        assert_eq!(PICKS[3].persist, true);
        assert_eq!(PICKS[3].direction, BalanceDirection::NegToPos);
    }

    #[test]
    fn open_step_take_roundtrip() {
        let mut a = app();
        assert_eq!(index(&a), None);
        open(&mut a);
        assert_eq!(index(&a), Some(0));
        step(&mut a, 1);
        assert_eq!(index(&a), Some(1));
        assert_eq!(
            take(&mut a),
            Some(BalancePick { persist: false, direction: BalanceDirection::NegToPos })
        );
        assert_eq!(index(&a), None, "take closes the picker");
    }

    #[test]
    fn steps_clamp_at_both_ends() {
        let mut a = app();
        open(&mut a);
        step(&mut a, -1);
        assert_eq!(index(&a), Some(0));
        step(&mut a, 25);
        assert_eq!(index(&a), Some(3));
        step(&mut a, 1);
        assert_eq!(index(&a), Some(3));
        step(&mut a, -25);
        assert_eq!(index(&a), Some(0));
    }

    #[test]
    fn closed_picker_ignores_keys() {
        let mut a = app();
        step(&mut a, 1);
        set(&mut a, 3);
        assert_eq!(take(&mut a), None);
        assert_eq!(index(&a), None);
    }

    #[test]
    fn set_clamps_and_digits_map() {
        let mut a = app();
        open(&mut a);
        set(&mut a, 99);
        assert_eq!(index(&a), Some(3));
        assert_eq!(index_for_digit('1'), Some(0));
        assert_eq!(index_for_digit('4'), Some(3));
        assert_eq!(index_for_digit('0'), None);
        assert_eq!(index_for_digit('5'), None);
        assert_eq!(index_for_digit('x'), None);
        close(&mut a);
        assert_eq!(index(&a), None);
    }
}
