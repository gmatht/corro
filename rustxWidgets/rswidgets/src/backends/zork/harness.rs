//! Typed, synchronous in-process test harness over the
//! [`crate::backends::zork::model`].
//!
//! This is the recommended replacement for a JSON test backend. It drives the
//! exact same model the REPL does, but in plain Rust: widgets are returned as
//! typed, `Clone`-able handles, actions (`click`, `type_into`, `toggle`,
//! `select`) are synchronous and fire callbacks immediately, and state is
//! asserted via getters or [`ZorkState::snapshot`].
//!
//! Each widget handle owns a shared `Rc<RefCell<ZorkState>>`, so widget methods
//! (e.g. [`Label::set_text`]) do not require a `&Harness` reference and closures
//! can capture handles without borrowing the harness.
//!
//! ```
//! use std::rc::Rc;
//! use std::cell::Cell;
//! use rswidgets::backends::zork::harness::Harness;
//!
//! let h = Harness::new();
//! let label = h.create_label("count: 0");
//! let btn = h.create_button("inc");
//! let counter = Rc::new(Cell::new(0));
//! {
//!     let c = counter.clone();
//!     let l = label.clone();
//!     btn.on_click(move || {
//!         let n = c.get() + 1;
//!         c.set(n);
//!         l.set_text(&format!("count: {}", n));
//!     });
//! }
//! h.click(&btn);
//! h.click(&btn);
//! assert_eq!(h.label_text(&label), Some("count: 2".to_string()));
//! ```

use std::cell::RefCell;
use std::rc::Rc;

use crate::backends::headless::DrawOp;
use crate::backends::zork::model::{MenuItemData, ZorkProps, ZorkState};

/// Opaque, cloneable handle to a node in the harness model.
///
/// Cloning a handle is cheap: it shares the same underlying node id and the same
/// backing [`ZorkState`].
#[derive(Clone)]
pub struct Widget {
    pub(crate) id: usize,
    pub(crate) state: Rc<RefCell<ZorkState>>,
}

impl Widget {
    pub(crate) fn new(id: usize, state: Rc<RefCell<ZorkState>>) -> Self {
        Widget { id, state }
    }
}

/// Trait implemented by all typed widget handles so the harness can accept any
/// widget in action/getter methods.
pub trait AsId {
    fn id(&self) -> usize;
}

/// Typed test harness. Owns its own model (shared via [`Rc`] with the widget
/// handles it produces) so tests are isolated from the thread-local singleton and
/// from each other.
pub struct Harness {
    state: Rc<RefCell<ZorkState>>,
}

impl Default for Harness {
    fn default() -> Self {
        Self::new()
    }
}

impl Harness {
    pub fn new() -> Self {
        Harness {
            state: Rc::new(RefCell::new(ZorkState::new())),
        }
    }

    /// Borrow the underlying model for direct assertions (e.g. via
    /// [`ZorkState::snapshot`]).
    pub fn with_state<R>(&self, f: impl FnOnce(&ZorkState) -> R) -> R {
        f(&self.state.borrow())
    }

    /// Mutate the underlying model. Needed for the few operations that have no
    /// dedicated action (e.g. registering a draw callback).
    pub fn with_state_mut<R>(&self, f: impl FnOnce(&mut ZorkState) -> R) -> R {
        f(&mut self.state.borrow_mut())
    }

    // -- creation --

    pub fn create_window(&self) -> Window {
        let id = self.state.borrow_mut().create_window();
        Window(Widget::new(id, self.state.clone()))
    }
    pub fn create_dialog(&self) -> Dialog {
        let id = self.state.borrow_mut().create_dialog();
        Dialog(Widget::new(id, self.state.clone()))
    }
    pub fn create_button(&self, label: &str) -> Button {
        let id = self.state.borrow_mut().create_button(label);
        Button(Widget::new(id, self.state.clone()))
    }
    pub fn create_label(&self, text: &str) -> Label {
        let id = self.state.borrow_mut().create_label(text);
        Label(Widget::new(id, self.state.clone()))
    }
    pub fn create_box(&self, horizontal: bool, spacing: i32) -> BoxWidget {
        let id = self.state.borrow_mut().create_box(horizontal, spacing);
        BoxWidget(Widget::new(id, self.state.clone()))
    }
    pub fn create_grid(&self) -> Grid {
        let id = self.state.borrow_mut().create_grid();
        Grid(Widget::new(id, self.state.clone()))
    }
    pub fn create_entry(&self) -> Entry {
        let id = self.state.borrow_mut().create_entry();
        Entry(Widget::new(id, self.state.clone()))
    }
    pub fn create_textview(&self) -> TextView {
        let id = self.state.borrow_mut().create_textview();
        TextView(Widget::new(id, self.state.clone()))
    }
    pub fn create_dropdown(&self, items: &[&str]) -> DropDown {
        let id = self.state.borrow_mut().create_dropdown(items);
        DropDown(Widget::new(id, self.state.clone()))
    }
    pub fn create_checkbutton(&self, label: &str) -> CheckButton {
        let id = self.state.borrow_mut().create_checkbutton(label);
        CheckButton(Widget::new(id, self.state.clone()))
    }
    /// Create a radio button, optionally in `group`'s group.
    ///
    /// Passing a seed joins its group. An ungrouped seed is promoted to be its
    /// own group, so the seed and every member created from it stay mutually
    /// exclusive. With no seed the button is independent. This mirrors
    /// `backends_zork_adapter::create_radiobutton`.
    pub fn create_radiobutton(&self, group: Option<&RadioButton>, label: &str) -> RadioButton {
        let gid = group.and_then(|r| r.0.state.borrow().radiobutton_group(r.0.id)).filter(|g| *g != 0);
        let id = match (gid, group) {
            (Some(g), _) => self.state.borrow_mut().create_radiobutton(Some(g), label),
            (None, Some(seed)) => {
                let seed_id = seed.0.id;
                self.state.borrow_mut().set_radiobutton_group(seed_id, seed_id);
                self.state.borrow_mut().create_radiobutton(Some(seed_id), label)
            }
            (None, None) => self.state.borrow_mut().create_radiobutton(None, label),
        };
        RadioButton(Widget::new(id, self.state.clone()))
    }
    pub fn create_menu(&self) -> Menu {
        let id = self.state.borrow_mut().create_menu();
        Menu(Widget::new(id, self.state.clone()))
    }
    pub fn create_canvas(&self) -> Canvas {
        let id = self.state.borrow_mut().create_canvas();
        Canvas(Widget::new(id, self.state.clone()))
    }
    pub fn create_overlay(&self) -> Overlay {
        let id = self.state.borrow_mut().create_overlay();
        Overlay(Widget::new(id, self.state.clone()))
    }
    pub fn create_scrolled_window(&self) -> ScrolledWindow {
        let id = self.state.borrow_mut().create_scrolled_window();
        ScrolledWindow(Widget::new(id, self.state.clone()))
    }
    pub fn create_fixed(&self) -> Fixed {
        let id = self.state.borrow_mut().create_fixed();
        Fixed(Widget::new(id, self.state.clone()))
    }
    pub fn create_spreadsheet(&self) -> Spreadsheet {
        let id = self.state.borrow_mut().create_spreadsheet();
        Spreadsheet(Widget::new(id, self.state.clone()))
    }
    /// Register a named action so a menu item naming it can be dispatched.
    pub fn create_simple_action(&self, name: &str) -> SimpleAction {
        let id = self.state.borrow_mut().create_simple_action(name);
        SimpleAction(Widget::new(id, self.state.clone()))
    }
    /// Build a `MenuBar` from a `Menu` model.
    pub fn create_menubar(&self, model: &Menu) -> MenuBar {
        let id = self.state.borrow_mut().create_menubar(model.0.id, std::ptr::null_mut());
        MenuBar(Widget::new(id, self.state.clone()))
    }

    // -- actions (operate on the shared state) --

    /// Fire the callbacks registered on a widget. For `Button` this is a click;
    /// for `CheckButton`/`RadioButton` it re-fires without toggling (use
    /// [`Self::toggle`] to toggle).
    ///
    /// Callbacks are taken out of the shared state and fired *after* releasing the
    /// borrow, so a callback may freely mutate the model (re-entrancy safe).
    pub fn click(&self, w: &impl AsId) {
        let id = w.id();
        let mut cbs = {
            let mut state = self.state.borrow_mut();
            match state.node_mut(id) {
                Some(n) => std::mem::take(&mut n.callbacks),
                None => return,
            }
        };
        for cb in cbs.iter_mut() {
            cb();
        }
        if let Some(n) = self.state.borrow_mut().node_mut(id) {
            n.callbacks = cbs;
        }
    }

    /// Take the callbacks off a node and fire them outside any borrow of the
    /// shared state (re-entrancy safe).
    fn fire_callbacks(&self, id: usize) {
        let mut cbs = {
            let mut state = self.state.borrow_mut();
            match state.node_mut(id) {
                Some(n) => std::mem::take(&mut n.callbacks),
                None => return,
            }
        };
        for cb in cbs.iter_mut() {
            cb();
        }
        if let Some(n) = self.state.borrow_mut().node_mut(id) {
            n.callbacks = cbs;
        }
    }

    /// Type text into an `Entry` and fire its `changed` callbacks.
    pub fn type_into(&self, w: &impl AsId, text: &str) {
        let id = w.id();
        self.state.borrow_mut().set_entry_text(id, text);
        self.fire_callbacks(id);
    }

    /// Toggle a `CheckButton` or `RadioButton` and fire its callbacks.
    pub fn toggle(&self, w: &impl AsId) {
        let id = w.id();
        self.state.borrow_mut().toggle(id);
        self.fire_callbacks(id);
    }

    /// Select a menu item (by 1-based index) on a `Menu`/`MenuBar` and fire its
    /// callbacks. Out-of-range indices are ignored.
    pub fn select(&self, w: &impl AsId, index: usize) {
        let id = w.id();
        let len = self.state.borrow().menu_items.get(&id).map(|i| i.len()).unwrap_or(0);
        if index == 0 || index > len {
            return;
        }
        self.fire_callbacks(id);
    }

    /// Select a menu item (0-based, as the model indexes them) and dispatch
    /// the `SimpleAction` it names. Returns the item's label, or `None` when
    /// the index is out of range or the item is a separator/section.
    ///
    /// This is [`Self::select`]'s replacement: `select` fired the *menu node's*
    /// own callbacks, so the action name on each item was decorative. A test
    /// that wants "pressing Open runs app.open" wants this.
    pub fn select_item(&self, w: &impl AsId, index: usize) -> Option<String> {
        self.state.borrow_mut().menu_select(w.id(), index)
    }

    /// Fire a registered action by name. Returns whether it resolved.
    pub fn activate(&self, action_name: &str) -> bool {
        self.state.borrow_mut().activate_action(action_name)
    }

    /// Deliver a dialog response. Only an id registered with
    /// [`Dialog::add_button`] reaches the handler.
    pub fn respond(&self, w: &impl AsId, response_id: i32) {
        self.state.borrow_mut().dialog_respond(w.id(), response_id);
    }

    // -- pointer / key actions --

    /// Click a `Canvas` at model coordinates, firing its click hooks and then
    /// its generic callbacks.
    pub fn click_at(&self, w: &impl AsId, x: f64, y: f64) {
        self.state.borrow_mut().pointer_click(w.id(), x, y);
    }

    /// Click with a button/modifier state. Fires the `click_button` hooks first,
    /// then the plain click hooks (the adapter chain order).
    pub fn click_button_at(&self, w: &impl AsId, x: f64, y: f64, button: u32, state: u32) {
        self.state.borrow_mut().pointer_click_button(w.id(), x, y, button, state);
    }

    /// Deliver a pointer-motion event.
    pub fn motion(&self, w: &impl AsId, x: f64, y: f64, state: u32) {
        self.state.borrow_mut().pointer_motion(w.id(), x, y, state);
    }

    /// Deliver a pointer release.
    pub fn release(&self, w: &impl AsId, x: f64, y: f64, button: u32, state: u32) {
        self.state.borrow_mut().pointer_release(w.id(), x, y, button, state);
    }

    /// Deliver a key. Returns whether a handler consumed it; a consumed key
    /// skips the `key_raw` chain, matching the adapter order.
    pub fn key(&self, w: &impl AsId, keyval: u32) -> bool {
        self.state.borrow_mut().key(w.id(), keyval)
    }

    /// Run a `Canvas`'s draw callback against a recording context and return
    /// the recorded ops — the headless equivalent of a paint.
    pub fn draw(&self, w: &impl AsId) -> Vec<DrawOp> {
        self.state.borrow_mut().draw_canvas(w.id())
    }

    /// Drive a `ScrolledWindow`'s scroll handler as if the user had scrolled.
    /// A headless viewport produces no real scroll events, so this is how the
    /// handler gets exercised.
    pub fn scroll(&self, w: &impl AsId, vertical: bool, pos: f64) {
        self.state.borrow_mut().scroll(w.id(), vertical, pos);
    }

    /// Position a `BoxWidget`'s children along its packing axis.
    pub fn layout_box(&self, w: &impl AsId, x: i32, y: i32, width: i32, height: i32) {
        self.state.borrow_mut().layout_box(w.id(), x, y, width, height);
    }

    /// Attach a child to a `Grid` cell.
    pub fn attach(&self, grid: &impl AsId, child: &impl AsId, left: i32, top: i32, width: i32, height: i32) {
        self.state.borrow_mut().grid_attach(grid.id(), child.id(), left, top, width, height);
    }

    /// Stack a child on an `Overlay` above its base child.
    pub fn add_overlay(&self, overlay: &impl AsId, child: &impl AsId) {
        self.state.borrow_mut().overlay_add(overlay.id(), child.id());
    }

    /// Set a child of a container widget (Window/Box/Grid/Dialog), replacing any
    /// existing parent.
    pub fn set_child(&self, parent: &impl AsId, child: &impl AsId) {
        let pid = parent.id();
        let cid = child.id();
        self.state.borrow_mut().set_child(pid, cid);
    }

    /// Append a child to a box/grid container.
    pub fn append(&self, parent: &impl AsId, child: &impl AsId) {
        let pid = parent.id();
        let cid = child.id();
        self.state.borrow_mut().append_child(pid, cid);
    }

    // -- getters --

    pub fn label_text(&self, w: &impl AsId) -> Option<String> {
        self.state.borrow().get_label_text(w.id())
    }
    pub fn entry_text(&self, w: &impl AsId) -> Option<String> {
        self.state.borrow().get_entry_text(w.id())
    }
    pub fn textview_text(&self, w: &impl AsId) -> Option<String> {
        self.state.borrow().get_textview_text(w.id())
    }
    pub fn checkbutton_checked(&self, w: &impl AsId) -> bool {
        self.state.borrow().get_checkbutton_checked(w.id())
    }
    pub fn radiobutton_checked(&self, w: &impl AsId) -> bool {
        self.state.borrow().get_radiobutton_checked(w.id())
    }
    pub fn dropdown_selected(&self, w: &impl AsId) -> i32 {
        self.state.borrow().get_dropdown_selected(w.id())
    }
    pub fn menu_items(&self, w: &impl AsId) -> Vec<MenuItemData> {
        self.state.borrow().menu_items.get(&w.id()).cloned().unwrap_or_default()
    }

    // -- property getters --
    //
    // These read the same `ZorkProps` bag every property setter writes, so a
    // test can assert any of geometry/visibility/expansion/margins/classes/
    // alignment/scroll without going through the JSON snapshot.

    /// The widget's full property bag. A cheap way to assert several
    /// properties at once.
    pub fn props(&self, w: &impl AsId) -> ZorkProps {
        self.state
            .borrow()
            .node(w.id())
            .map(|n| n.props.clone())
            .unwrap_or_default()
    }

    pub fn visible(&self, w: &impl AsId) -> bool {
        self.state.borrow().get_visible(w.id())
    }
    pub fn offset(&self, w: &impl AsId) -> Option<(Option<i32>, Option<i32>)> {
        self.state.borrow().get_offset(w.id())
    }
    pub fn size_request(&self, w: &impl AsId) -> Option<(Option<i32>, Option<i32>)> {
        self.state.borrow().get_size_request(w.id())
    }
    pub fn classes(&self, w: &impl AsId) -> Vec<String> {
        self.state.borrow().classes(w.id())
    }
    pub fn has_class(&self, w: &impl AsId, class_name: &str) -> bool {
        self.state.borrow().has_class(w.id(), class_name)
    }
    pub fn hexpand(&self, w: &impl AsId) -> bool {
        self.props(w).hexpand
    }
    pub fn vexpand(&self, w: &impl AsId) -> bool {
        self.props(w).vexpand
    }
    pub fn margin_start(&self, w: &impl AsId) -> i32 {
        self.props(w).margin_start
    }
    pub fn margin_top(&self, w: &impl AsId) -> i32 {
        self.props(w).margin_top
    }
    pub fn entry_position(&self, w: &impl AsId) -> Option<usize> {
        self.state.borrow().get_entry_position(w.id())
    }
    pub fn editable(&self, w: &impl AsId) -> bool {
        self.state.borrow().get_editable(w.id())
    }
    /// The current `(h, v)` scroll offsets of a `ScrolledWindow`.
    pub fn scroll_offsets(&self, w: &impl AsId) -> (f64, f64) {
        self.state.borrow().get_scroll(w.id())
    }
    pub fn focused(&self) -> Option<usize> {
        self.state.borrow().get_focus()
    }
    pub fn has_focus(&self, w: &impl AsId) -> bool {
        self.state.borrow().has_focus(w.id())
    }
    pub fn button_label(&self, w: &impl AsId) -> Option<String> {
        self.state.borrow().get_button_label(w.id())
    }
    pub fn window_title(&self, w: &impl AsId) -> Option<String> {
        self.state.borrow().get_window_title(w.id())
    }
    pub fn redraws(&self, w: &impl AsId) -> u32 {
        self.state.borrow().redraw_count(w.id())
    }
    pub fn canvas_size(&self, w: &impl AsId) -> (i32, i32) {
        self.state.borrow().canvas_size(w.id())
    }
    /// The widget's children, in order.
    pub fn children(&self, w: &impl AsId) -> Vec<usize> {
        self.state.borrow().node(w.id()).map(|n| n.children.clone()).unwrap_or_default()
    }
    pub fn parent(&self, w: &impl AsId) -> Option<usize> {
        self.state.borrow().node(w.id()).and_then(|n| n.parent)
    }
    /// An `Overlay`'s stacked layers, base child excluded.
    pub fn overlay_layers(&self, w: &impl AsId) -> Vec<usize> {
        self.state.borrow().overlay_layers(w.id())
    }
    /// A `Grid`'s `(cols, rows)` as grown by `attach`.
    pub fn grid_dimensions(&self, w: &impl AsId) -> (usize, usize) {
        match self.state.borrow().node(w.id()).map(|n| &n.kind) {
            Some(crate::backends::zork::model::ZorkKind::Grid { cols, rows }) => (*cols, *rows),
            _ => (0, 0),
        }
    }
    /// A `Dialog`'s `(label, response_id)` buttons, in registration order.
    pub fn dialog_buttons(&self, w: &impl AsId) -> Vec<(String, i32)> {
        self.state.borrow().dialog_buttons(w.id())
    }
    pub fn dialog_transient_for(&self, w: &impl AsId) -> Option<usize> {
        self.state.borrow().dialog_transient_for(w.id())
    }
    pub fn is_destroyed(&self, w: &impl AsId) -> bool {
        self.state.borrow().node(w.id()).map(|n| n.destroyed).unwrap_or(false)
    }
    pub fn sheet_cell(&self, w: &impl AsId, row: u32, col: u32) -> Option<String> {
        self.state.borrow().sheet_get_cell(w.id(), row, col)
    }
    /// The size a `BoxWidget`'s children imply, or `None` for a non-box.
    pub fn measure_box(&self, w: &impl AsId) -> Option<(i32, i32)> {
        self.state.borrow().measure_box(w.id())
    }

    /// Serialize the current model to a JSON snapshot string.
    pub fn snapshot_json(&self) -> String {
        serde_json::to_string(&self.state.borrow().snapshot()).expect("snapshot serializes")
    }
}

macro_rules! widget_handle {
    ($name:ident) => {
        #[derive(Clone)]
        pub struct $name(Widget);
        impl AsId for $name {
            fn id(&self) -> usize {
                self.0.id
            }
        }
        impl $name {
            /// The internal node id (useful for diagnostics / snapshot lookups).
            pub fn id(&self) -> usize {
                self.0.id
            }
            /// Register a callback fired by [`Harness::click`].
            pub fn on_click(&self, f: impl FnMut() + 'static) {
                self.0.state.borrow_mut().add_callback(self.0.id, Box::new(f));
            }

            // -- type-independent properties --
            //
            // These mirror the `ZorkProps` fields, so a handle can carry every
            // property the model records regardless of its widget type.

            /// Show/hide the widget. A hidden widget takes no space in a box
            /// layout and cannot take focus.
            pub fn set_visible(&self, visible: bool) {
                self.0.state.borrow_mut().set_visible(self.0.id, visible);
            }
            pub fn set_size_request(&self, w: i32, h: i32) {
                self.0.state.borrow_mut().set_size_request(self.0.id, w, h);
            }
            pub fn set_offset(&self, x: i32, y: i32) {
                self.0.state.borrow_mut().set_offset(self.0.id, x, y);
            }
            pub fn set_hexpand(&self, expand: bool) {
                self.0.state.borrow_mut().set_hexpand(self.0.id, expand);
            }
            pub fn set_vexpand(&self, expand: bool) {
                self.0.state.borrow_mut().set_vexpand(self.0.id, expand);
            }
            pub fn set_margin_start(&self, px: i32) {
                self.0.state.borrow_mut().set_margin_start(self.0.id, px);
            }
            pub fn set_margin_top(&self, px: i32) {
                self.0.state.borrow_mut().set_margin_top(self.0.id, px);
            }
            pub fn add_class(&self, class_name: &str) {
                self.0.state.borrow_mut().add_class(self.0.id, class_name);
            }
            pub fn remove_class(&self, class_name: &str) {
                self.0.state.borrow_mut().remove_class(self.0.id, class_name);
            }
            pub fn set_xalign(&self, x: f32) {
                self.0.state.borrow_mut().set_xalign(self.0.id, x);
            }
            /// Pin the widget's width so sibling reflow cannot change it.
            /// `None` releases the pin.
            pub fn set_fixed_width(&self, w: Option<i32>) {
                self.0.state.borrow_mut().set_fixed_width(self.0.id, w);
            }
            /// Give the widget keyboard focus. Refused for a hidden or
            /// non-focusable widget.
            pub fn grab_focus(&self) {
                self.0.state.borrow_mut().set_focus(self.0.id);
            }
        }
    };
}

widget_handle!(Window);
widget_handle!(Dialog);
widget_handle!(Button);
widget_handle!(Label);
widget_handle!(BoxWidget);
widget_handle!(Grid);
widget_handle!(Entry);
widget_handle!(TextView);
widget_handle!(DropDown);
widget_handle!(CheckButton);
widget_handle!(RadioButton);
widget_handle!(Menu);
widget_handle!(Canvas);
widget_handle!(Overlay);
widget_handle!(ScrolledWindow);
widget_handle!(Fixed);
widget_handle!(Spreadsheet);
widget_handle!(SimpleAction);
widget_handle!(MenuBar);

impl Window {
    pub fn set_title(&self, title: &str) {
        self.0.state.borrow_mut().set_window_title(self.0.id, title);
    }
}

impl Label {
    pub fn set_text(&self, text: &str) {
        self.0.state.borrow_mut().set_label_text(self.0.id, text);
    }
}

impl Entry {
    pub fn set_text(&self, text: &str) {
        self.0.state.borrow_mut().set_entry_text(self.0.id, text);
    }
    pub fn connect_changed(&self, f: impl FnMut() + 'static) {
        self.0.state.borrow_mut().add_callback(self.0.id, Box::new(f));
    }
}

impl TextView {
    pub fn set_text(&self, text: &str) {
        self.0.state.borrow_mut().set_textview_text(self.0.id, text);
    }
}

impl CheckButton {
    pub fn set_checked(&self, checked: bool) {
        self.0.state.borrow_mut().set_checkbutton_checked(self.0.id, checked);
    }
    pub fn on_toggle(&self, f: impl FnMut() + 'static) {
        self.0.state.borrow_mut().add_callback(self.0.id, Box::new(f));
    }
}

impl RadioButton {
    /// The group this button belongs to, or `None` if it is not a radio.
    /// Group `0` means "ungrouped".
    pub fn group(&self) -> Option<usize> {
        self.0.state.borrow().radiobutton_group(self.0.id)
    }
    /// Join `group`, so a button created without a seed can be grouped later.
    pub fn set_group(&self, group_id: usize) {
        self.0.state.borrow_mut().set_radiobutton_group(self.0.id, group_id);
    }
    pub fn set_checked(&self, checked: bool) {
        self.0.state.borrow_mut().set_radiobutton_checked(self.0.id, checked);
    }
    pub fn on_toggle(&self, f: impl FnMut() + 'static) {
        self.0.state.borrow_mut().add_callback(self.0.id, Box::new(f));
    }
}

impl DropDown {
    pub fn set_items(&self, items: &[&str]) {
        self.0.state.borrow_mut().set_dropdown_items(self.0.id, items);
    }
    /// `None` selects nothing, matching the adapter's GTK/NWG-style signature.
    pub fn set_active(&self, index: Option<u32>) {
        let idx = match index {
            Some(i) => i as i32,
            None => -1,
        };
        self.0.state.borrow_mut().set_dropdown_selected(self.0.id, idx);
    }
    pub fn connect_changed(&self, f: impl FnMut() + 'static) {
        self.0.state.borrow_mut().add_callback(self.0.id, Box::new(f));
    }
}

impl Menu {
    pub fn append(&self, label: &str, action: &str) {
        self.0.state.borrow_mut().menu_append(self.0.id, label, action);
    }
    /// Append an item with a keyboard accelerator.
    pub fn append_with_shortcut(&self, label: &str, action: &str, accelerator: &str) {
        self.0.state.borrow_mut().menu_append_with_shortcut(self.0.id, label, action, accelerator);
    }
    /// Append a non-interactive section header.
    pub fn append_section(&self, label: &str) {
        self.0.state.borrow_mut().menu_append_section(self.0.id, label);
    }
    /// Append a non-interactive divider.
    pub fn append_separator(&self, label: &str) {
        self.0.state.borrow_mut().menu_append_separator(self.0.id, label);
    }
    /// Append a check item; returns its 0-based index for
    /// [`Self::set_item_checked`].
    pub fn append_check(&self, label: &str, action: &str, checked: bool) -> usize {
        self.0.state.borrow_mut().menu_append_check(self.0.id, label, action, checked)
    }
    /// Append a radio item in `group`; returns its 0-based index.
    pub fn append_radio(&self, label: &str, action: &str, group: usize, checked: bool) -> usize {
        self.0.state.borrow_mut().menu_append_radio(self.0.id, label, action, group, checked)
    }
    /// Set a check/radio item's state, keeping radio groups exclusive.
    pub fn set_item_checked(&self, index: usize, checked: bool) {
        self.0.state.borrow_mut().set_menu_item_checked(self.0.id, index, checked);
    }
    /// The menu's items, in order.
    pub fn items(&self) -> Vec<MenuItemData> {
        self.0.state.borrow().menu_items.get(&self.0.id).cloned().unwrap_or_default()
    }
}

impl SimpleAction {
    /// Fire the action's handlers. Returns whether the name resolved, so a
    /// test can assert a menu item actually reached its action.
    pub fn activate_named(&self, name: &str) -> bool {
        self.0.state.borrow_mut().activate_action(name)
    }
}

impl MenuBar {
    /// The bar's items, in order.
    pub fn items(&self) -> Vec<MenuItemData> {
        self.0.state.borrow().menu_items.get(&self.0.id).cloned().unwrap_or_default()
    }
}

impl Entry {
    /// The caret, as a character index.
    pub fn get_position(&self) -> Option<usize> {
        self.0.state.borrow().get_entry_position(self.0.id)
    }
    /// Move the caret, clamped to the buffer length.
    pub fn set_position(&self, pos: usize) {
        self.0.state.borrow_mut().set_entry_position(self.0.id, pos);
    }
}

impl Canvas {
    /// Install the draw callback. [`Harness::draw`] runs it against a
    /// recording context so a test can assert what it painted.
    pub fn set_draw_callback(&self, cb: Box<dyn FnMut(&mut dyn crate::core::DrawContext, i32, i32)>) {
        self.0.state.borrow_mut().set_draw_callback(self.0.id, cb);
    }
    pub fn clear_draw_callback(&self) {
        self.0.state.borrow_mut().clear_draw_callback(self.0.id);
    }
    /// The content size, also passed to the draw callback.
    pub fn set_content_size(&self, w: i32, h: i32) {
        self.0.state.borrow_mut().set_size_request(self.0.id, w, h);
    }
    /// Register a click handler, fired by [`Harness::click_at`].
    pub fn on_click_at(&self, cb: Box<dyn FnMut(f64, f64)>) {
        self.0.state.borrow_mut().add_click_hook(self.0.id, cb);
    }
    pub fn on_click_button(&self, cb: Box<dyn FnMut(f64, f64, u32, u32)>) {
        self.0.state.borrow_mut().add_click_button_hook(self.0.id, cb);
    }
    pub fn on_motion(&self, cb: Box<dyn FnMut(f64, f64, u32)>) {
        self.0.state.borrow_mut().add_motion_hook(self.0.id, cb);
    }
    pub fn on_release(&self, cb: Box<dyn FnMut(f64, f64, u32, u32)>) {
        self.0.state.borrow_mut().add_release_hook(self.0.id, cb);
    }
    /// Register a key handler. Returning `true` consumes the key, which skips
    /// the `on_key_raw` chain.
    pub fn on_key(&self, cb: Box<dyn FnMut(u32) -> bool>) {
        self.0.state.borrow_mut().add_key_hook(self.0.id, cb);
    }
    pub fn on_key_raw(&self, cb: Box<dyn FnMut(u32, u32) -> bool>) {
        self.0.state.borrow_mut().add_key_raw_hook(self.0.id, cb);
    }
    pub fn set_screen_origin(&self, origin: Option<(i32, i32)>) {
        self.0.state.borrow_mut().set_screen_origin(self.0.id, origin);
    }
    pub fn screen_origin(&self) -> Option<(i32, i32)> {
        self.0.state.borrow().node(self.0.id).and_then(|n| n.pointer.screen_origin)
    }
    pub fn queue_redraw(&self) {
        self.0.state.borrow_mut().queue_redraw(self.0.id);
    }
}

impl Overlay {
    /// The base child.
    pub fn set_child(&self, child: &impl AsId) {
        self.0.state.borrow_mut().set_child(self.0.id, child.id());
    }
    /// Set whether pointer events pass through a layer to the base.
    pub fn set_pass_through(&self, child: &impl AsId, pass: bool) {
        self.0.state.borrow_mut().overlay_set_pass_through(self.0.id, child.id(), pass);
    }
    /// The stacked layers, base child excluded.
    pub fn layers(&self) -> Vec<usize> {
        self.0.state.borrow().overlay_layers(self.0.id)
    }
    /// Make the overlay and every descendant visible.
    pub fn show_all(&self) {
        self.0.state.borrow_mut().show_all(self.0.id);
    }
}

impl ScrolledWindow {
    pub fn set_child(&self, child: &impl AsId) {
        self.0.state.borrow_mut().set_child(self.0.id, child.id());
    }
    /// Drive both scroll offsets. Values are clamped to `upper - page`.
    pub fn scroll_to(&self, hval: f64, hupper: f64, hpage: f64, vval: f64, vupper: f64, vpage: f64) {
        self.0.state.borrow_mut().set_scroll(self.0.id, hval, hupper, hpage, vval, vupper, vpage);
    }
    /// Register a scroll handler, fired by [`Harness::scroll`].
    pub fn on_scroll(&self, cb: Box<dyn FnMut(bool, f64)>) {
        self.0.state.borrow_mut().add_scroll_callback(self.0.id, cb);
    }
}

impl Dialog {
    /// Register a button; `response_id` is what [`Harness::respond`] delivers.
    pub fn add_button(&self, label: &str, response_id: i32) {
        self.0.state.borrow_mut().dialog_add_button(self.0.id, label, response_id);
    }
    pub fn set_default_response(&self, response_id: i32) {
        self.0.state.borrow_mut().dialog_set_default_response(self.0.id, response_id);
    }
    /// Parent this dialog to `parent`, so a test can assert the relation.
    pub fn set_transient_for(&self, parent: &impl AsId) {
        self.0.state.borrow_mut().dialog_set_transient_for(self.0.id, Some(parent.id()));
    }
    /// Register a response handler, called with the response id.
    pub fn connect_response(&self, f: impl FnMut(i32) + 'static) {
        self.0.state.borrow_mut().add_response_callback(self.0.id, Box::new(f));
    }
    /// Hide and mark the dialog destroyed.
    pub fn close(&self) {
        self.0.state.borrow_mut().dialog_mark_destroyed(self.0.id);
    }
}

impl Spreadsheet {
    pub fn set_cell(&self, row: u32, col: u32, text: &str) {
        self.0.state.borrow_mut().sheet_set_cell(self.0.id, row, col, text, 0);
    }
    pub fn set_raw_cell(&self, row: u32, col: u32, text: &str) {
        self.0.state.borrow_mut().sheet_set_raw_cell(self.0.id, row, col, text);
    }
    pub fn get_cell(&self, row: u32, col: u32) -> Option<String> {
        self.0.state.borrow().sheet_get_cell(self.0.id, row, col)
    }
    pub fn set_cell_style(&self, row: u32, col: u32, style: u8) {
        let text = self.get_cell(row, col).unwrap_or_default();
        self.0.state.borrow_mut().sheet_set_cell(self.0.id, row, col, &text, style);
    }
    pub fn set_border_title(&self, title: &str) {
        self.0.state.borrow_mut().sheet_set_border_title(self.0.id, title);
    }
    pub fn border_title(&self) -> Option<String> {
        self.0.state.borrow().sheet_border_title(self.0.id)
    }
}

impl BoxWidget {
    /// The size the box's visible children imply.
    pub fn measure(&self) -> Option<(i32, i32)> {
        self.0.state.borrow().measure_box(self.0.id)
    }
}

impl Grid {
    /// The grid's `(cols, rows)` as grown by [`Harness::attach`].
    pub fn dimensions(&self) -> (usize, usize) {
        match self.0.state.borrow().node(self.0.id).map(|n| &n.kind) {
            Some(crate::backends::zork::model::ZorkKind::Grid { cols, rows }) => (*cols, *rows),
            _ => (0, 0),
        }
    }
}
