use crate::gui::sheet::SharedState;

pub use rustxwidgets::core::key::{
    RETURN as KEY_RETURN,
    ESCAPE as KEY_ESC,
    BACKSPACE as KEY_BACKSPACE,
    DELETE as KEY_DELETE,
    LEFT as KEY_LEFT,
    UP as KEY_UP,
    RIGHT as KEY_RIGHT,
    DOWN as KEY_DOWN,
    TAB as KEY_TAB,
    HOME as KEY_HOME,
    END as KEY_END,
    PAGE_UP as KEY_PAGE_UP,
    PAGE_DOWN as KEY_PAGE_DOWN,
    F1 as KEY_F1,
    F2 as KEY_F2,
};
pub use rustxwidgets::core::key::ENTER as KEY_ENTER;

pub enum EditAction {
    Commit(String),
    Cancel,
    Continue,
}

pub fn handle_edit_input(keyval: u32, shared: &SharedState, redraw: &dyn Fn()) -> EditAction {
    match keyval {
        #[allow(unreachable_patterns)]
        KEY_RETURN | KEY_ENTER => {
            let text = shared.edit_buf.borrow().clone();
            shared.editing.set(false);
            shared.edit_buf.borrow_mut().clear();
            EditAction::Commit(text)
        }
        KEY_ESC => {
            shared.editing.set(false);
            shared.edit_buf.borrow_mut().clear();
            EditAction::Cancel
        }
        KEY_BACKSPACE => {
            shared.edit_buf.borrow_mut().pop();
            redraw();
            EditAction::Continue
        }
        KEY_LEFT | KEY_RIGHT | KEY_UP | KEY_DOWN | KEY_TAB | KEY_HOME | KEY_END |
        KEY_PAGE_UP | KEY_PAGE_DOWN | KEY_DELETE | KEY_F2 => EditAction::Continue,
        _ if keyval >= 32 && keyval <= 126 => {
            shared.edit_buf.borrow_mut().push(char::from_u32(keyval).unwrap_or('?'));
            redraw();
            EditAction::Continue
        }
        _ => EditAction::Continue,
    }
}
