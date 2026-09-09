use crate::core::action::Action;
use rswidgets::core::key::{LEFT, RIGHT, UP, DOWN, RETURN, ENTER, ESCAPE, TAB,
    HOME, END, PAGE_UP, PAGE_DOWN, DELETE, F2};

pub fn keyval_to_action(keyval: u32, _state: &KeyState) -> Option<Action> {

    match keyval {
        LEFT => Some(Action::MoveLeft),
        RIGHT => Some(Action::MoveRight),
        UP => Some(Action::MoveUp),
        DOWN => Some(Action::MoveDown),
        HOME => Some(Action::MoveHome),
        END => Some(Action::MoveEnd),
        PAGE_UP => Some(Action::MovePageUp),
        PAGE_DOWN => Some(Action::MovePageDown),
        #[allow(unreachable_patterns)]
        RETURN | ENTER => Some(Action::StartEdit),
        ESCAPE => Some(Action::CancelEdit),
        TAB => Some(Action::MoveRight),
        DELETE => Some(Action::DeleteSelection),
        F2 => Some(Action::StartEdit),
        _ => None,
    }
}

pub struct KeyState {
    pub shift: bool,
    pub ctrl: bool,
    pub alt: bool,
}

impl KeyState {
    pub fn new() -> Self {
        KeyState { shift: false, ctrl: false, alt: false }
    }
}
