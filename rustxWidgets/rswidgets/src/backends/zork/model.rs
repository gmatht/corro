//! Pure, I/O-free model for the headless `zork` backend.
//!
//! Everything here is plain data + operations on it: no terminal access, no
//! `rustyline`, and no `thread_local` global required. A [`ZorkState`] is just a
//! value you can own, clone, snapshot to JSON, and drive from either the
//! interactive REPL ([`crate::backends::zork::repl`]) or the typed test harness
//! ([`crate::backends::zork::harness`]).
//!
//! The free functions that the per-widget adapters call (e.g.
//! [`crate::backends::zork::facade::create_button`], `set_label_text`) operate on
//! a thread-local singleton so that the existing `backends_zork_adapter` shim
//! keeps working unchanged. That singleton is a thin wrapper over a `ZorkState`;
//! tests that want a clean, isolated instance should use [`ZorkState`] /
//! [`harness::Harness`] directly instead.
//!
//! # Model shape
//!
//! A node is *kind* + *properties*:
//!
//! * [`ZorkKind`] holds the type-specific data (a label's text, a grid's cell
//!   list, a menu's items).
//! * [`ZorkProps`] holds everything that applies to *every* widget regardless of
//!   type — geometry, visibility, expansion, margins, CSS classes, alignment,
//!   font, scroll, and the pointer/key callback lists. Keeping it in one
//!   serializable struct is what lets [`ZorkState::snapshot`] report the whole
//!   model and lets tests diff a property change without caring which widget it
//!   came from.

use std::cell::RefCell;
use std::collections::HashMap;

pub type Callback = Box<dyn FnMut()>;

/// A rectangle in model units (the same integer pixel contract the GUI
/// adapters use, so a caller can pass its values straight through).
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, serde::Serialize)]
pub struct Rect {
    pub x: i32,
    pub y: i32,
    pub w: i32,
    pub h: i32,
}

/// Per-widget properties that are independent of the widget's type.
///
/// Defaults are the "unset" state: an unconstrained rect, visible, not
/// expanded, no margins, no classes, default alignment and font.
#[derive(Clone, Debug, serde::Serialize)]
pub struct ZorkProps {
    /// Position/size set by `set_offset` / `set_size_request` / `layout` /
    /// `attach`. `None` fields mean "not set", so a caller can tell the
    /// difference between `set_size_request(0, 0)` and never calling it.
    pub offset_x: Option<i32>,
    pub offset_y: Option<i32>,
    pub width: Option<i32>,
    pub height: Option<i32>,
    /// Whether the widget is shown. A hidden widget is skipped by
    /// [`ZorkState::describe`]-style introspection and by focus traversal.
    pub visible: bool,
    pub hexpand: bool,
    pub vexpand: bool,
    pub margin_start: i32,
    pub margin_top: i32,
    /// CSS style classes, in the order they were added (duplicates ignored).
    pub classes: Vec<String>,
    /// 0.0 = start/left, 1.0 = end/right for the horizontal axis.
    pub xalign: f32,
    /// -1 = fill, 0 = start, 1 = end (GTK alignment convention).
    pub halign: i32,
    pub valign: i32,
    /// Pin the widget's width so sibling reflow cannot change it. `None`
    /// releases the pin.
    pub fixed_width: Option<i32>,
    /// `Some(..)` when the caller has set a font/size, else the backend default.
    pub font: Option<String>,
    pub font_size: f64,
    /// Whether a `TextView` accepts user input.
    pub editable: bool,
    /// Scroll offsets, set via [`ZorkState::set_scroll`].
    pub hscroll: f64,
    pub vscroll: f64,
}

impl Default for ZorkProps {
    fn default() -> Self {
        ZorkProps {
            offset_x: None,
            offset_y: None,
            width: None,
            height: None,
            visible: true,
            hexpand: false,
            vexpand: false,
            margin_start: 0,
            margin_top: 0,
            classes: Vec::new(),
            xalign: 0.0,
            halign: 0,
            valign: 0,
            fixed_width: None,
            font: None,
            font_size: 0.0,
            editable: true,
            hscroll: 0.0,
            vscroll: 0.0,
        }
    }
}

/// Font style bits, matching the GTK `FontStyle` decomposition used by
/// `set_font_style` on the other adapters.
pub mod font_style {
    pub const NORMAL: i32 = 0;
    pub const BOLD: i32 = 1;
    pub const ITALIC: i32 = 2;
    pub const STRIKETHROUGH: i32 = 4;
    pub const UNDERLINE: i32 = 8;
}

/// One menu item. Extends the original label/action/submenu triple with the
/// state GTK/NWG menus actually carry: a kind (so separators and check/radio
/// items are representable), an accelerator, a checked flag and `enabled`.
#[derive(Clone, Copy, Debug, PartialEq, Eq, serde::Serialize)]
pub enum MenuItemKind {
    Normal,
    /// A non-interactive divider. `label` is the visual text (often empty).
    Separator,
    /// A toggle item; `checked` is its state.
    Check,
    /// A radio item; `checked` is its state, `group` groups siblings.
    Radio,
    /// A non-interactive section header.
    Section,
}

#[derive(Clone, Debug, serde::Serialize)]
pub struct MenuItemData {
    pub label: String,
    pub action: String,
    pub submenu: Option<Vec<MenuItemData>>,
    /// What kind of item this is. Defaults to [`MenuItemKind::Normal`].
    pub kind: MenuItemKind,
    /// Keyboard shortcut, e.g. `"Ctrl+S"`. Empty when unset.
    pub accelerator: String,
    pub checked: bool,
    pub enabled: bool,
    /// Radio group id; `0` when the item is not a radio.
    pub group: usize,
}

impl MenuItemData {
    /// A plain, enabled, non-toggle item. The `menu_append` path produces
    /// exactly this so existing callers are unaffected by the new fields.
    pub fn normal(label: &str, action: &str) -> Self {
        MenuItemData {
            label: label.to_string(),
            action: action.to_string(),
            submenu: None,
            kind: MenuItemKind::Normal,
            accelerator: String::new(),
            checked: false,
            enabled: true,
            group: 0,
        }
    }
}

#[derive(Clone, Debug)]
pub enum ZorkKind {
    Window { title: String },
    Button { label: String },
    Label { text: String },
    BoxWidget { horizontal: bool, spacing: i32 },
    Grid { cols: usize, rows: usize },
    Entry { buffer: String, cursor: usize },
    CheckButton { label: String, checked: bool },
    RadioButton { label: String, checked: bool, group_id: usize },
    Dialog { title: String },
    Menu,
    MenuBar,
    SimpleAction,
    DropDown { items: Vec<String>, selected: Option<usize> },
    TextView { text: String },
    /// A custom drawing surface. The draw callback is a `Box<dyn FnMut>`, so it
    /// lives on the node rather than in the kind (which must stay `Clone`-able
    /// for `Snapshot`); [`ZorkState::draw_canvas`] invokes it against a
    /// recording context.
    Canvas { w: i32, h: i32 },
    /// A stack of children drawn over a base child; `set_child` is the base and
    /// `add_overlay` pushes additional layers.
    Overlay,
    /// A clipped, scrollable viewport around a child.
    ScrolledWindow,
    /// A child positioned at an explicit (x, y).
    Fixed,
    /// A registered, named action group. Kept in the model so `create_menubar`
    /// can resolve the action names a menu references.
    Application,
    /// A grid-of-cells widget with a border title.
    Spreadsheet { border_title: String },
}

/// A single cell of a [`ZorkKind::Spreadsheet`], addressed 1-based.
#[derive(Clone, Debug, Default, serde::Serialize)]
pub struct SheetCell {
    pub text: String,
    /// Coarse style id (0 = default). Enough for a model test to assert that a
    /// style was applied; rasterisation is the draw callback's job.
    pub style: u8,
    /// True when `text` is a literal the user typed rather than a formula.
    pub raw: bool,
}

pub struct ZorkNode {
    pub id: usize,
    pub kind: ZorkKind,
    pub parent: Option<usize>,
    pub children: Vec<usize>,
    pub callbacks: Vec<Callback>,
    /// Type-independent properties (see [`ZorkProps`]).
    pub props: ZorkProps,
    /// Whether this node can take keyboard focus.
    pub can_focus: bool,
    /// The canvas draw callback, set by `Canvas::set_draw_callback`. Stored
    /// out-of-band from `kind` so `kind` can stay `Clone`.
    pub draw: Option<Box<dyn FnMut(&mut dyn crate::core::DrawContext, i32, i32)>>,
    /// Pointer/key callbacks for a `Canvas`. Kept separate from `callbacks`
    /// (which is "the thing the harness `click` fires") so clicking a canvas
    /// does not also fire its generic action callbacks.
    pub pointer: PointerHooks,
    /// `ScrolledWindow` scroll-notification handlers, called as
    /// `cb(vertical, pos)`.
    pub scroll_cbs: Vec<Box<dyn FnMut(bool, f64)>>,
    /// `Dialog` response handlers, called as `cb(response_id)`. Separate from
    /// `callbacks` because a response must carry its id.
    pub response_cbs: Vec<Box<dyn FnMut(i32)>>,
    /// Spreadsheet cell contents, keyed 1-based by `(row, col)`.
    pub cells: HashMap<(u32, u32), SheetCell>,
    /// Dialog buttons as `(label, response_id)`, and the default response id.
    pub dialog_buttons: Vec<(String, i32)>,
    pub dialog_default_response: Option<i32>,
    pub dialog_transient_for: Option<usize>,
    /// Whether a dialog has been destroyed (`Dialog::close` / `mark_destroyed`).
    pub destroyed: bool,
    /// Overlay layers, in insertion order, and per-child pass-through flags.
    pub overlays: Vec<usize>,
    pub overlay_pass_through: HashMap<usize, bool>,
}

/// The pointer and key callbacks a `Canvas`/`Entry` can carry.
#[allow(clippy::type_complexity)]
#[derive(Default)]
pub struct PointerHooks {
    pub click: Vec<Box<dyn FnMut(f64, f64)>>,
    pub click_button: Vec<Box<dyn FnMut(f64, f64, u32, u32)>>,
    pub motion: Vec<Box<dyn FnMut(f64, f64, u32)>>,
    pub release: Vec<Box<dyn FnMut(f64, f64, u32, u32)>>,
    pub key: Vec<Box<dyn FnMut(u32) -> bool>>,
    pub key_raw: Vec<Box<dyn FnMut(u32, u32) -> bool>>,
    /// This canvas's top-left in "screen" coordinates, when known.
    pub screen_origin: Option<(i32, i32)>,
    /// Number of queued redraws, so a test can assert `queue_redraw` was called
    /// even though nothing is rasterised.
    pub redraws: u32,
}

impl std::fmt::Debug for PointerHooks {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("PointerHooks")
            .field("click", &self.click.len())
            .field("click_button", &self.click_button.len())
            .field("motion", &self.motion.len())
            .field("release", &self.release.len())
            .field("key", &self.key.len())
            .field("key_raw", &self.key_raw.len())
            .field("screen_origin", &self.screen_origin)
            .field("redraws", &self.redraws)
            .finish()
    }
}

impl ZorkNode {
    /// Run all callbacks currently stored on this node. Callers must invoke this
    /// *outside* of any borrow of the owning [`ZorkState`], because a callback
    /// may re-enter the model.
    pub fn fire_callbacks(&mut self) {
        // Take the callbacks out so the borrow on `self` is released before each
        // (potentially re-entrant) call.
        let mut cbs = std::mem::take(&mut self.callbacks);
        for cb in cbs.iter_mut() {
            cb();
        }
        self.callbacks = cbs;
    }
}

#[derive(Clone, Debug, serde::Serialize)]
pub struct SnapshotNode {
    pub id: usize,
    pub kind: String,
    pub parent: Option<usize>,
    pub children: Vec<usize>,
    pub title: Option<String>,
    pub text: Option<String>,
    pub label: Option<String>,
    pub checked: Option<bool>,
    pub selected: Option<usize>,
    pub items: Option<Vec<String>>,
    /// The full property bag, so a snapshot test can assert geometry,
    /// visibility, expansion, margins, classes and scroll.
    pub props: ZorkProps,
    pub can_focus: bool,
    /// Callback *counts*, never the closures themselves (not serializable).
    pub callbacks: usize,
    pub click_hooks: usize,
    pub redraws: u32,
    /// Spreadsheet cells as `(row, col, text, style, raw)`.
    pub cells: Vec<SheetCellEntry>,
    /// Dialog buttons as `(label, response_id)`.
    pub dialog_buttons: Vec<(String, i32)>,
    pub overlays: Vec<usize>,
}

#[derive(Clone, Debug, serde::Serialize)]
pub struct SheetCellEntry {
    pub row: u32,
    pub col: u32,
    pub text: String,
    pub style: u8,
    pub raw: bool,
}

#[derive(Clone, Debug, serde::Serialize)]
pub struct Snapshot {
    pub nodes: Vec<SnapshotNode>,
    pub menu_items: HashMap<usize, Vec<MenuItemData>>,
    pub current_id: usize,
    pub running: bool,
    /// The node id that holds keyboard focus, if any.
    pub focused: Option<usize>,
    /// Overlay child -> pass-through flag.
    pub overlay_pass_through: HashMap<usize, bool>,
    /// Spreadsheet cells keyed by `(row, col)`, flattened for easy lookup.
    pub cells: HashMap<(u32, u32), SheetCell>,
}

pub struct ZorkState {
    pub nodes: Vec<ZorkNode>,
    pub next_id: usize,
    pub running: bool,
    pub current_id: usize,
    pub prev_location: Option<usize>,
    /// Menu model items keyed by menu node id.
    pub menu_items: HashMap<usize, Vec<MenuItemData>>,
    /// The node that currently holds keyboard focus.
    pub focused: Option<usize>,
    /// Named actions, so `MenuBar` dispatch can resolve an action name to the
    /// `SimpleAction` node that registered it.
    pub action_names: HashMap<String, usize>,
}

impl Default for ZorkState {
    fn default() -> Self {
        Self::new()
    }
}

impl ZorkState {
    pub fn new() -> Self {
        ZorkState {
            nodes: Vec::new(),
            next_id: 1,
            running: true,
            current_id: 0,
            prev_location: None,
            menu_items: HashMap::new(),
            focused: None,
            action_names: HashMap::new(),
        }
    }

    pub fn alloc_id(&mut self) -> usize {
        let id = self.next_id;
        self.next_id += 1;
        id
    }

    pub fn add_node(&mut self, kind: ZorkKind, parent: Option<usize>) -> usize {
        let id = self.alloc_id();
        // If this is the first node (Window), set current_id
        if self.nodes.is_empty() {
            self.current_id = id;
        }
        let can_focus = matches!(
            kind,
            ZorkKind::Button { .. }
                | ZorkKind::Entry { .. }
                | ZorkKind::CheckButton { .. }
                | ZorkKind::RadioButton { .. }
                | ZorkKind::DropDown { .. }
                | ZorkKind::TextView { .. }
                | ZorkKind::Canvas { .. }
        );
        self.nodes.push(ZorkNode {
            id,
            kind,
            parent,
            children: Vec::new(),
            callbacks: Vec::new(),
            props: ZorkProps::default(),
            can_focus,
            draw: None,
            pointer: PointerHooks::default(),
            scroll_cbs: Vec::new(),
            response_cbs: Vec::new(),
            cells: HashMap::new(),
            dialog_buttons: Vec::new(),
            dialog_default_response: None,
            dialog_transient_for: None,
            destroyed: false,
            overlays: Vec::new(),
            overlay_pass_through: HashMap::new(),
        });
        if let Some(pid) = parent {
            if let Some(p) = self.nodes.iter_mut().find(|n| n.id == pid) {
                p.children.push(id);
            }
        }
        id
    }

    pub fn node_mut(&mut self, id: usize) -> Option<&mut ZorkNode> {
        self.nodes.iter_mut().find(|n| n.id == id)
    }

    pub fn node(&self, id: usize) -> Option<&ZorkNode> {
        self.nodes.iter().find(|n| n.id == id)
    }

    pub fn find_window_id(&self) -> Option<usize> {
        self.nodes
            .iter()
            .find(|n| matches!(n.kind, ZorkKind::Window { .. } | ZorkKind::Dialog { .. }))
            .map(|n| n.id)
    }

    // -- Factory functions (operate on this state) --

    pub fn create_window(&mut self) -> usize {
        self.add_node(ZorkKind::Window { title: String::new() }, None)
    }

    pub fn create_button(&mut self, label: &str) -> usize {
        self.add_node(ZorkKind::Button { label: label.to_string() }, self.find_window_id())
    }

    pub fn create_label(&mut self, text: &str) -> usize {
        self.add_node(ZorkKind::Label { text: text.to_string() }, self.find_window_id())
    }

    pub fn create_box(&mut self, horizontal: bool, spacing: i32) -> usize {
        self.add_node(ZorkKind::BoxWidget { horizontal, spacing }, self.find_window_id())
    }

    pub fn create_grid(&mut self) -> usize {
        self.add_node(ZorkKind::Grid { cols: 0, rows: 0 }, self.find_window_id())
    }

    pub fn create_entry(&mut self) -> usize {
        self.add_node(ZorkKind::Entry { buffer: String::new(), cursor: 0 }, self.find_window_id())
    }

    pub fn create_menu(&mut self) -> usize {
        let id = self.add_node(ZorkKind::Menu, self.find_window_id());
        self.menu_items.insert(id, Vec::new());
        id
    }

    pub fn create_simple_action(&mut self, name: &str) -> usize {
        let id = self.add_node(ZorkKind::SimpleAction, self.find_window_id());
        self.action_names.insert(name.to_string(), id);
        id
    }

    pub fn create_menubar(&mut self, model_id: usize, _action_group: *mut std::os::raw::c_void) -> usize {
        let bar_id = self.add_node(ZorkKind::MenuBar, self.find_window_id());
        let items = self.menu_items.get(&model_id).cloned().unwrap_or_default();
        self.menu_items.insert(bar_id, items);
        // Re-materialise one `Menu` node per submenu so the bar has navigable
        // children (the REPL walks the tree).
        let nested: Vec<Vec<MenuItemData>> = self
            .menu_items
            .get(&bar_id)
            .cloned()
            .unwrap_or_default()
            .into_iter()
            .filter_map(|i| i.submenu)
            .collect();
        for sub_items in nested {
            let sub_id = self.add_node(ZorkKind::Menu, Some(bar_id));
            self.menu_items.insert(sub_id, sub_items);
        }
        bar_id
    }

    pub fn create_dialog(&mut self) -> usize {
        self.add_node(ZorkKind::Dialog { title: String::new() }, self.find_window_id())
    }

    pub fn create_dropdown(&mut self, items: &[&str]) -> usize {
        let items_str: Vec<String> = items.iter().map(|s| s.to_string()).collect();
        self.add_node(ZorkKind::DropDown { items: items_str, selected: None }, self.find_window_id())
    }

    pub fn create_checkbutton(&mut self, label: &str) -> usize {
        self.add_node(ZorkKind::CheckButton { label: label.to_string(), checked: false }, self.find_window_id())
    }

    pub fn create_radiobutton(&mut self, group_id: Option<usize>, label: &str) -> usize {
        let gid = group_id.unwrap_or(0);
        self.add_node(ZorkKind::RadioButton { label: label.to_string(), checked: false, group_id: gid }, self.find_window_id())
    }

    pub fn create_textview(&mut self) -> usize {
        self.add_node(ZorkKind::TextView { text: String::new() }, self.find_window_id())
    }

    /// A custom drawing surface. Sizing goes through the same `props` bag as
    /// every other widget so `set_size_request` / `set_content_size` work
    /// uniformly.
    pub fn create_canvas(&mut self) -> usize {
        self.add_node(ZorkKind::Canvas { w: 0, h: 0 }, self.find_window_id())
    }

    /// An overlay stack. `set_child` is the base layer, `add_overlay` pushes
    /// further ones.
    pub fn create_overlay(&mut self) -> usize {
        self.add_node(ZorkKind::Overlay, self.find_window_id())
    }

    pub fn create_scrolled_window(&mut self) -> usize {
        self.add_node(ZorkKind::ScrolledWindow, self.find_window_id())
    }

    pub fn create_fixed(&mut self) -> usize {
        self.add_node(ZorkKind::Fixed, self.find_window_id())
    }

    pub fn create_application(&mut self) -> usize {
        self.add_node(ZorkKind::Application, None)
    }

    pub fn create_spreadsheet(&mut self) -> usize {
        self.add_node(ZorkKind::Spreadsheet { border_title: String::new() }, self.find_window_id())
    }

    // -- Menu model --

    pub fn menu_append(&mut self, menu_id: usize, label: &str, action: &str) {
        self.menu_items
            .entry(menu_id)
            .or_default()
            .push(MenuItemData::normal(label, action));
    }

    /// Append an item with a keyboard accelerator.
    pub fn menu_append_with_shortcut(&mut self, menu_id: usize, label: &str, action: &str, accelerator: &str) {
        let mut item = MenuItemData::normal(label, action);
        item.accelerator = accelerator.to_string();
        self.menu_items.entry(menu_id).or_default().push(item);
    }

    /// Append a non-interactive section header.
    pub fn menu_append_section(&mut self, menu_id: usize, label: &str) {
        let mut item = MenuItemData::normal(label, "");
        item.kind = MenuItemKind::Section;
        item.enabled = false;
        self.menu_items.entry(menu_id).or_default().push(item);
    }

    /// Append a divider. `label` is the visual text (usually empty).
    pub fn menu_append_separator(&mut self, menu_id: usize, label: &str) {
        let mut item = MenuItemData::normal(label, "");
        item.kind = MenuItemKind::Separator;
        item.enabled = false;
        self.menu_items.entry(menu_id).or_default().push(item);
    }

    /// Append a check item and return its index, so the caller can later
    /// [`Self::set_menu_item_checked`] it.
    pub fn menu_append_check(&mut self, menu_id: usize, label: &str, action: &str, checked: bool) -> usize {
        let mut item = MenuItemData::normal(label, action);
        item.kind = MenuItemKind::Check;
        item.checked = checked;
        let items = self.menu_items.entry(menu_id).or_default();
        items.push(item);
        items.len() - 1
    }

    /// Append a radio item in `group`, returning its index.
    pub fn menu_append_radio(&mut self, menu_id: usize, label: &str, action: &str, group: usize, checked: bool) -> usize {
        let mut item = MenuItemData::normal(label, action);
        item.kind = MenuItemKind::Radio;
        item.group = group;
        item.checked = checked;
        if checked {
            // Radio items are mutually exclusive: clear the rest of the group.
            if let Some(items) = self.menu_items.get_mut(&menu_id) {
                for other in items.iter_mut() {
                    if other.kind == MenuItemKind::Radio && other.group == group {
                        other.checked = false;
                    }
                }
            }
        }
        let items = self.menu_items.entry(menu_id).or_default();
        items.push(item);
        items.len() - 1
    }

    /// Set the `checked` state of a menu item, enforcing radio exclusivity.
    pub fn set_menu_item_checked(&mut self, menu_id: usize, index: usize, checked: bool) {
        let (kind, group) = match self.menu_items.get(&menu_id).and_then(|i| i.get(index)) {
            Some(it) => (it.kind, it.group),
            None => return,
        };
        let Some(items) = self.menu_items.get_mut(&menu_id) else { return };
        if let Some(it) = items.get_mut(index) {
            it.checked = checked;
        }
        if checked && kind == MenuItemKind::Radio {
            for (i, other) in items.iter_mut().enumerate() {
                if i == index {
                    continue;
                }
                if other.kind == MenuItemKind::Radio && other.group == group {
                    other.checked = false;
                }
            }
        }
    }

    /// Append `item` to a submenu without copying the submenu's items: the
    /// child `Menu` node is linked so later edits to it stay visible.
    pub fn menu_append_submenu(&mut self, menu_id: usize, label: &str, submenu_id: usize) {
        let mut item = MenuItemData::normal(label, "");
        item.submenu = self.menu_items.get(&submenu_id).cloned();
        if let Some(n) = self.node_mut(submenu_id) {
            n.parent = Some(menu_id);
        }
        if let Some(p) = self.node_mut(menu_id) {
            if !p.children.contains(&submenu_id) {
                p.children.push(submenu_id);
            }
        }
        self.menu_items.entry(menu_id).or_default().push(item);
    }

    /// Set the label of a `Button`/`CheckButton`/`RadioButton`.
    pub fn set_button_label(&mut self, id: usize, label: &str) {
        if let Some(n) = self.node_mut(id) {
            match &mut n.kind {
                ZorkKind::Button { label: l } => *l = label.to_string(),
                ZorkKind::CheckButton { label: l, .. } => *l = label.to_string(),
                ZorkKind::RadioButton { label: l, .. } => *l = label.to_string(),
                _ => {}
            }
        }
    }

    pub fn get_button_label(&self, id: usize) -> Option<String> {
        self.node(id).and_then(|n| match &n.kind {
            ZorkKind::Button { label }
            | ZorkKind::CheckButton { label, .. }
            | ZorkKind::RadioButton { label, .. } => Some(label.clone()),
            _ => None,
        })
    }

    // -- Setters / getters --

    pub fn set_window_title(&mut self, id: usize, title: &str) {
        if let Some(n) = self.node_mut(id) {
            if let ZorkKind::Window { title: ref mut t } | ZorkKind::Dialog { title: ref mut t } = n.kind {
                *t = title.to_string();
            }
        }
    }

    pub fn get_window_title(&self, id: usize) -> Option<String> {
        self.node(id).and_then(|n| match &n.kind {
            ZorkKind::Window { title } | ZorkKind::Dialog { title } => Some(title.clone()),
            _ => None,
        })
    }

    pub fn set_label_text(&mut self, id: usize, text: &str) {
        if let Some(n) = self.node_mut(id) {
            if let ZorkKind::Label { text: ref mut t } = n.kind {
                *t = text.to_string();
            }
        }
    }

    pub fn get_label_text(&self, id: usize) -> Option<String> {
        self.node(id).and_then(|n| {
            if let ZorkKind::Label { ref text } = n.kind {
                Some(text.clone())
            } else {
                None
            }
        })
    }

    pub fn add_callback(&mut self, id: usize, cb: Callback) {
        if let Some(n) = self.node_mut(id) {
            n.callbacks.push(cb);
        }
    }

    pub fn set_entry_text(&mut self, id: usize, text: &str) {
        if let Some(n) = self.node_mut(id) {
            if let ZorkKind::Entry { ref mut buffer, ref mut cursor } = n.kind {
                *buffer = text.to_string();
                *cursor = buffer.chars().count();
            }
        }
    }

    pub fn get_entry_text(&self, id: usize) -> Option<String> {
        self.node(id).and_then(|n| {
            if let ZorkKind::Entry { ref buffer, .. } = n.kind {
                Some(buffer.clone())
            } else {
                None
            }
        })
    }

    /// The entry caret, as a character index.
    pub fn get_entry_position(&self, id: usize) -> Option<usize> {
        self.node(id).and_then(|n| match &n.kind {
            ZorkKind::Entry { cursor, .. } => Some(*cursor),
            _ => None,
        })
    }

    /// Move the entry caret, clamped to the buffer length.
    pub fn set_entry_position(&mut self, id: usize, pos: usize) {
        if let Some(n) = self.node_mut(id) {
            if let ZorkKind::Entry { ref buffer, ref mut cursor } = n.kind {
                *cursor = pos.min(buffer.chars().count());
            }
        }
    }

    pub fn set_textview_text(&mut self, id: usize, text: &str) {
        if let Some(n) = self.node_mut(id) {
            if let ZorkKind::TextView { text: ref mut t } = n.kind {
                *t = text.to_string();
            }
        }
    }

    pub fn get_textview_text(&self, id: usize) -> Option<String> {
        self.node(id).and_then(|n| {
            if let ZorkKind::TextView { ref text } = n.kind {
                Some(text.clone())
            } else {
                None
            }
        })
    }

    /// Append to a `TextView` buffer. Unlike the adapters' read-modify-write
    /// fallback this is a real incremental insert, so a high-rate log pane
    /// stays O(appended) rather than O(document).
    pub fn append_textview_text(&mut self, id: usize, text: &str) {
        if let Some(n) = self.node_mut(id) {
            if let ZorkKind::TextView { text: ref mut t } = n.kind {
                t.push_str(text);
            }
        }
    }

    pub fn set_dropdown_items(&mut self, id: usize, items: &[&str]) {
        if let Some(n) = self.node_mut(id) {
            if let ZorkKind::DropDown { items: ref mut items_vec, ref mut selected } = n.kind {
                *items_vec = items.iter().map(|s| s.to_string()).collect();
                // Drop a now-out-of-range selection rather than reporting a
                // selection that points at nothing.
                if selected.is_some_and(|s| s >= items_vec.len()) {
                    *selected = None;
                }
            }
        }
    }

    pub fn set_dropdown_selected(&mut self, id: usize, idx: i32) {
        if let Some(n) = self.node_mut(id) {
            if let ZorkKind::DropDown { ref mut selected, ref items } = n.kind {
                *selected = if idx >= 0 && (idx as usize) < items.len() {
                    Some(idx as usize)
                } else {
                    None
                };
            }
        }
    }

    pub fn get_dropdown_selected(&self, id: usize) -> i32 {
        self.node(id)
            .and_then(|n| {
                if let ZorkKind::DropDown { ref selected, .. } = n.kind {
                    selected.map(|s| s as i32)
                } else {
                    None
                }
            })
            .unwrap_or(-1)
    }

    pub fn get_checkbutton_checked(&self, id: usize) -> bool {
        self.node(id)
            .map(|n| {
                if let ZorkKind::CheckButton { ref checked, .. } = n.kind {
                    *checked
                } else {
                    false
                }
            })
            .unwrap_or(false)
    }

    pub fn get_radiobutton_checked(&self, id: usize) -> bool {
        self.node(id)
            .map(|n| {
                if let ZorkKind::RadioButton { ref checked, .. } = n.kind {
                    *checked
                } else {
                    false
                }
            })
            .unwrap_or(false)
    }

    pub fn set_checkbutton_checked(&mut self, id: usize, checked: bool) {
        if let Some(n) = self.node_mut(id) {
            if let ZorkKind::CheckButton { checked: ref mut c, .. } = n.kind {
                *c = checked;
            }
        }
    }

    /// Check a radio button, clearing every other radio in the same group.
    /// A group id of `0` means "ungrouped" and is left alone.
    pub fn set_radiobutton_checked(&mut self, id: usize, checked: bool) {
        let group = match self.node(id) {
            Some(n) => match n.kind {
                ZorkKind::RadioButton { group_id, .. } => group_id,
                _ => return,
            },
            None => return,
        };
        if checked && group != 0 {
            // Collect the ids first so we are not mutating while iterating.
            let siblings: Vec<usize> = self
                .nodes
                .iter()
                .filter(|n| matches!(n.kind, ZorkKind::RadioButton { group_id: g, .. } if g == group))
                .map(|n| n.id)
                .collect();
            for sid in siblings {
                if sid == id {
                    continue;
                }
                if let Some(n) = self.node_mut(sid) {
                    if let ZorkKind::RadioButton { checked: ref mut c, .. } = n.kind {
                        *c = false;
                    }
                }
            }
        }
        if let Some(n) = self.node_mut(id) {
            if let ZorkKind::RadioButton { checked: ref mut c, .. } = n.kind {
                *c = checked;
            }
        }
    }

    /// Reparent `child_id` under `parent_id`, removing it from any previous
    /// parent. Unlike the old implementation this also drops the child's
    /// *own* `parent` pointer before re-linking, so a node cannot end up
    /// listed under two parents.
    pub fn set_child(&mut self, parent_id: usize, child_id: usize) {
        if parent_id == child_id {
            return;
        }
        // Detach from the current parent, if any.
        if let Some(old) = self.node(child_id).and_then(|n| n.parent) {
            if let Some(p) = self.node_mut(old) {
                p.children.retain(|c| *c != child_id);
                p.overlays.retain(|c| *c != child_id);
            }
        }
        // Drop the child's subtree from the global dedup so re-adding is clean.
        self.nodes.iter_mut().for_each(|n| {
            n.children.retain(|c| *c != child_id);
            n.overlays.retain(|c| *c != child_id);
        });
        if let Some(parent) = self.node_mut(parent_id) {
            if !parent.children.contains(&child_id) {
                parent.children.push(child_id);
            }
        }
        if let Some(child) = self.node_mut(child_id) {
            child.parent = Some(parent_id);
        }
    }

    /// Append a child to a container, keeping any existing children.
    ///
    /// This is genuinely an *append*: the old implementation delegated to
    /// `set_child`, which replaced the child list and so made
    /// `BoxWidget::append` twice a no-op for the first child.
    pub fn append_child(&mut self, parent_id: usize, child_id: usize) {
        if parent_id == child_id {
            return;
        }
        if let Some(old) = self.node(child_id).and_then(|n| n.parent) {
            if old == parent_id {
                return; // already a child here
            }
            if let Some(p) = self.node_mut(old) {
                p.children.retain(|c| *c != child_id);
                p.overlays.retain(|c| *c != child_id);
            }
        }
        self.nodes.iter_mut().for_each(|n| {
            n.children.retain(|c| *c != child_id);
            n.overlays.retain(|c| *c != child_id);
        });
        if let Some(parent) = self.node_mut(parent_id) {
            parent.children.push(child_id);
        }
        if let Some(child) = self.node_mut(child_id) {
            child.parent = Some(parent_id);
        }
    }

    /// Remove `child_id` from its parent entirely.
    pub fn remove_child(&mut self, parent_id: usize, child_id: usize) {
        if let Some(p) = self.node_mut(parent_id) {
            p.children.retain(|c| *c != child_id);
            p.overlays.retain(|c| *c != child_id);
        }
        if self.node(child_id).and_then(|n| n.parent) == Some(parent_id) {
            if let Some(c) = self.node_mut(child_id) {
                c.parent = None;
            }
        }
    }

    // -- Geometry / layout --

    pub fn set_offset(&mut self, id: usize, x: i32, y: i32) {
        if let Some(n) = self.node_mut(id) {
            n.props.offset_x = Some(x);
            n.props.offset_y = Some(y);
        }
    }

    pub fn set_size_request(&mut self, id: usize, w: i32, h: i32) {
        if let Some(n) = self.node_mut(id) {
            n.props.width = Some(w);
            n.props.height = Some(h);
        }
    }

    pub fn get_size_request(&self, id: usize) -> Option<(Option<i32>, Option<i32>)> {
        self.node(id).map(|n| (n.props.width, n.props.height))
    }

    pub fn get_offset(&self, id: usize) -> Option<(Option<i32>, Option<i32>)> {
        self.node(id).map(|n| (n.props.offset_x, n.props.offset_y))
    }

    /// Record a child's grid cell. Sets the position and size, and grows the
    /// grid's `cols`/`rows` to cover the cell.
    pub fn grid_attach(&mut self, grid_id: usize, child_id: usize, left: i32, top: i32, width: i32, height: i32) {
        self.set_offset(child_id, left, top);
        self.set_size_request(child_id, width, height);
        self.append_child(grid_id, child_id);
        // A cell spanning `width` columns starting at `left` ends at
        // `left + width`; a zero-width cell still occupies one column.
        let col = if width > 0 { left + width } else { left + 1 };
        let row = if height > 0 { top + height } else { top + 1 };
        let (col, row) = (col.max(0) as usize, row.max(0) as usize);
        if let Some(n) = self.node_mut(grid_id) {
            if let ZorkKind::Grid { cols: ref mut c, rows: ref mut r } = n.kind {
                *c = (*c).max(col);
                *r = (*r).max(row);
            }
        }
    }

<<<<<<< HEAD
    /// Compute the size of a `BoxWidget` from its children: the sum of the
    /// children along the packing axis (plus `spacing` between them) and the
    /// max on the cross axis. Returns `None` for a non-box node.
=======
    /// Compute the size a `BoxWidget` needs for its visible children: the sum
    /// of the children's extent along the *packing* axis (plus `spacing`
    /// between them) and the largest extent on the cross axis.
    ///
    /// A horizontal box packs along x, so its width is the sum of the child
    /// widths and its height the tallest child. Returns `None` for a non-box.
>>>>>>> 184a0b72 (feat(zork): make the zork backend build and cover the GTK/NWG surface)
    pub fn measure_box(&self, id: usize) -> Option<(i32, i32)> {
        let n = self.node(id)?;
        let (horizontal, spacing) = match n.kind {
            ZorkKind::BoxWidget { horizontal, spacing } => (horizontal, spacing),
            _ => return None,
        };
        let kids: Vec<&ZorkNode> = n
            .children
            .iter()
            .filter_map(|c| self.node(*c))
            .filter(|c| c.props.visible)
            .collect();
        if kids.is_empty() {
            return Some((0, 0));
        }
<<<<<<< HEAD
        let main: i32 = kids.iter().filter_map(|c| c.props.width).sum::<i32>()
            + spacing * (kids.len() as i32 - 1).max(0);
        let cross: i32 = kids
            .iter()
            .filter_map(|c| c.props.height)
            .max()
            .unwrap_or(0);
        Some(if horizontal { (main, cross) } else { (cross, main) })
    }

    /// Assign sequential positions to a box's visible children along the
    /// packing axis, honouring `spacing`, and record each child's size from
    /// its own size request (falling back to the measured box size).
    ///
    /// This is what makes `BoxWidget::layout(x, y, w, h)` a real operation
    /// rather than a discarded argument list.
=======
        // `pick` selects a child's extent on the requested axis.
        let main_of = |k: &ZorkNode| if horizontal { k.props.width } else { k.props.height };
        let cross_of = |k: &ZorkNode| if horizontal { k.props.height } else { k.props.width };
        let main: i32 = kids.iter().filter_map(|k| main_of(k)).sum::<i32>()
            + spacing * (kids.len() as i32 - 1).max(0);
        let cross: i32 = kids.iter().filter_map(|k| cross_of(k)).max().unwrap_or(0);
        Some(if horizontal { (main, cross) } else { (cross, main) })
    }

    /// Position a box's visible children along the packing axis, starting at
    /// `(x, y)` and honouring `spacing`, and record the box's own size request.
    ///
    /// A child advances the cursor by its own extent on the packing axis, so
    /// a child with no size request advances by nothing. Each child keeps its
    /// own size request — this is a positioning pass, not a resize pass; use
    /// [`Self::set_size_request`] to size children first.
>>>>>>> 184a0b72 (feat(zork): make the zork backend build and cover the GTK/NWG surface)
    pub fn layout_box(&mut self, id: usize, x: i32, y: i32, w: i32, h: i32) {
        let (horizontal, spacing) = match self.node(id) {
            Some(n) => match n.kind {
                ZorkKind::BoxWidget { horizontal, spacing } => (horizontal, spacing),
                _ => return,
            },
            None => return,
        };
        self.set_size_request(id, w, h);
        let kids: Vec<usize> = self
            .node(id)
            .map(|n| n.children.clone())
            .unwrap_or_default()
            .into_iter()
            .filter(|c| self.node(*c).is_some_and(|n| n.props.visible))
            .collect();
<<<<<<< HEAD
        if kids.is_empty() {
            return;
        }
        // Cross-axis extent available to each child.
        let cross_total = if horizontal { h } else { w };
        let mut main = x;
        let mut cross = y;
        for (i, kid) in kids.iter().enumerate() {
            let (kx, ky) = self
                .node(*kid)
                .map(|n| (n.props.offset_x.unwrap_or(0), n.props.offset_y.unwrap_or(0)))
                .unwrap_or((0, 0));
            if horizontal {
                self.set_offset(*kid, main + kx, y + ky);
                let ch = self.node(*kid).and_then(|n| n.props.height).unwrap_or(cross_total);
                main += ch.max(0) + spacing;
            } else {
                self.set_offset(*kid, x + kx, main + ky);
                let cw = self.node(*kid).and_then(|n| n.props.width).unwrap_or(cross_total);
                main += cw.max(0) + spacing;
            }
            if i == 0 && self.node(*kid).and_then(|n| n.props.width).is_none() {
                // Give the first child the whole cross extent so a single-child
                // box fills it.
                if horizontal {
                    self.set_size_request(*kid, self.node(*kid).and_then(|n| n.props.width).unwrap_or(0), cross_total);
                } else {
                    self.set_size_request(*kid, cross_total, self.node(*kid).and_then(|n| n.props.height).unwrap_or(0));
                }
            }
=======
        let mut main = if horizontal { x } else { y };
        for kid in kids {
            let extent = self
                .node(kid)
                .and_then(|n| if horizontal { n.props.width } else { n.props.height })
                .unwrap_or(0)
                .max(0);
            if horizontal {
                self.set_offset(kid, main, y);
            } else {
                self.set_offset(kid, x, main);
            }
            main += extent + spacing;
>>>>>>> 184a0b72 (feat(zork): make the zork backend build and cover the GTK/NWG surface)
        }
    }

    // -- Visibility / expansion / margins / classes / alignment --

    pub fn set_visible(&mut self, id: usize, visible: bool) {
        if let Some(n) = self.node_mut(id) {
            n.props.visible = visible;
        }
    }

    pub fn get_visible(&self, id: usize) -> bool {
        self.node(id).map(|n| n.props.visible).unwrap_or(false)
    }

    pub fn set_hexpand(&mut self, id: usize, v: bool) {
        if let Some(n) = self.node_mut(id) {
            n.props.hexpand = v;
        }
    }

    pub fn set_vexpand(&mut self, id: usize, v: bool) {
        if let Some(n) = self.node_mut(id) {
            n.props.vexpand = v;
        }
    }

    /// Set a child's expansion flag. GTK/NWG route this through the parent
    /// (`gtk_widget_set_hexpand(child, ...)`); the model stores it on the child
    /// so the flag survives without needing the parent to be resolved.
    pub fn set_child_hexpand(&mut self, _parent_id: usize, child_id: usize, v: bool) {
        self.set_hexpand(child_id, v);
    }

    pub fn set_child_vexpand(&mut self, _parent_id: usize, child_id: usize, v: bool) {
        self.set_vexpand(child_id, v);
    }

    pub fn set_margin_start(&mut self, id: usize, px: i32) {
        if let Some(n) = self.node_mut(id) {
            n.props.margin_start = px;
        }
    }

    pub fn set_margin_top(&mut self, id: usize, px: i32) {
        if let Some(n) = self.node_mut(id) {
            n.props.margin_top = px;
        }
    }

    pub fn add_class(&mut self, id: usize, class_name: &str) {
        if let Some(n) = self.node_mut(id) {
            if !n.props.classes.iter().any(|c| c == class_name) {
                n.props.classes.push(class_name.to_string());
            }
        }
    }

    pub fn remove_class(&mut self, id: usize, class_name: &str) {
        if let Some(n) = self.node_mut(id) {
            n.props.classes.retain(|c| c != class_name);
        }
    }

    pub fn has_class(&self, id: usize, class_name: &str) -> bool {
        self.node(id).map(|n| n.props.classes.iter().any(|c| c == class_name)).unwrap_or(false)
    }

    pub fn classes(&self, id: usize) -> Vec<String> {
        self.node(id).map(|n| n.props.classes.clone()).unwrap_or_default()
    }

    pub fn set_xalign(&mut self, id: usize, x: f32) {
        if let Some(n) = self.node_mut(id) {
            n.props.xalign = x;
        }
    }

    pub fn set_halign(&mut self, id: usize, align: i32) {
        if let Some(n) = self.node_mut(id) {
            n.props.halign = align;
        }
    }

    pub fn set_valign(&mut self, id: usize, align: i32) {
        if let Some(n) = self.node_mut(id) {
            n.props.valign = align;
        }
    }

    pub fn set_fixed_width(&mut self, id: usize, w: Option<i32>) {
        if let Some(n) = self.node_mut(id) {
            n.props.fixed_width = w;
            // A pinned width is a width request, so the two agree.
            if let Some(width) = w {
                n.props.width = Some(width);
            }
        }
    }

    pub fn set_font_style(&mut self, id: usize, font: Option<&str>, size: f64) {
        if let Some(n) = self.node_mut(id) {
            n.props.font = font.map(|f| f.to_string());
            n.props.font_size = size;
        }
    }

    /// Whether a font is bold. The model records the font *name* the caller
    /// passed, so this reads the `bold` bit out of the [`font_style`] flags the
    /// adapter decomposed the caller's style into.
    pub fn font_is_bold(&self, id: usize) -> bool {
        self.node(id)
            .map(|n| n.props.font.as_deref().is_some_and(|f| f.split_whitespace().any(|w| w.eq_ignore_ascii_case("bold"))))
            .unwrap_or(false)
    }

    pub fn set_editable(&mut self, id: usize, editable: bool) {
        if let Some(n) = self.node_mut(id) {
            n.props.editable = editable;
        }
    }

    pub fn get_editable(&self, id: usize) -> bool {
        self.node(id).map(|n| n.props.editable).unwrap_or(false)
    }

    // -- Focus --

    /// Move keyboard focus to `id`. Focus is cleared from the previous holder
    /// and refused for a hidden or non-focusable node.
    pub fn set_focus(&mut self, id: usize) {
        if self.node(id).is_none() {
            return;
        }
        if self.node(id).map(|n| n.props.visible && n.can_focus).unwrap_or(false) {
            self.focused = Some(id);
        }
    }

    pub fn get_focus(&self) -> Option<usize> {
        self.focused
    }

    pub fn has_focus(&self, id: usize) -> bool {
        self.focused == Some(id)
    }

    /// The focused node's id when it is focusable, else `None`. Used by
    /// `Entry::has_focus` and the adapter's `grab_focus`.
    pub fn focused_entry(&self) -> Option<usize> {
        self.focused
    }

    pub fn set_can_focus(&mut self, id: usize, can: bool) {
        if let Some(n) = self.node_mut(id) {
            n.can_focus = can;
        }
    }

    // -- Scroll --

    /// Record a scroll position on a `ScrolledWindow`. `hval`/`vval` are the
    /// values; the `*_upper`/`*_page` arguments are accepted and clamped so a
    /// caller can never scroll past the end of the document.
    pub fn set_scroll(&mut self, id: usize, hval: f64, hupper: f64, hpage: f64, vval: f64, vupper: f64, vpage: f64) {
        if let Some(n) = self.node_mut(id) {
            n.props.hscroll = clamp_scroll(hval, hupper, hpage);
            n.props.vscroll = clamp_scroll(vval, vupper, vpage);
        }
    }

    pub fn get_scroll(&self, id: usize) -> (f64, f64) {
        self.node(id).map(|n| (n.props.hscroll, n.props.vscroll)).unwrap_or((0.0, 0.0))
    }

    /// Register a scroll-notification handler, called as `cb(vertical, pos)`.
    pub fn add_scroll_callback(&mut self, id: usize, cb: Box<dyn FnMut(bool, f64)>) {
        if let Some(n) = self.node_mut(id) {
            n.scroll_cbs.push(cb);
        }
    }

    /// Drive a scroll handler as if the user had scrolled, and update the
    /// recorded offset. A headless viewport produces no real scroll events, so
    /// this is the only way to exercise the handler.
    pub fn scroll(&mut self, id: usize, vertical: bool, pos: f64) {
        if let Some(n) = self.node_mut(id) {
            if vertical {
                n.props.vscroll = pos.max(0.0);
            } else {
                n.props.hscroll = pos.max(0.0);
            }
        }
        let mut cbs = self.node_mut(id).map(|n| std::mem::take(&mut n.scroll_cbs)).unwrap_or_default();
        for cb in cbs.iter_mut() {
            cb(vertical, pos);
        }
        if let Some(n) = self.node_mut(id) {
            n.scroll_cbs = cbs;
        }
    }

    // -- Canvas / drawing --

    pub fn set_draw_callback(
        &mut self,
        id: usize,
        cb: Box<dyn FnMut(&mut dyn crate::core::DrawContext, i32, i32)>,
    ) {
        if let Some(n) = self.node_mut(id) {
            n.draw = Some(cb);
        }
    }

    pub fn clear_draw_callback(&mut self, id: usize) {
        if let Some(n) = self.node_mut(id) {
            n.draw = None;
        }
    }

    pub fn has_draw_callback(&self, id: usize) -> bool {
        self.node(id).map(|n| n.draw.is_some()).unwrap_or(false)
    }

    pub fn queue_redraw(&mut self, id: usize) {
        if let Some(n) = self.node_mut(id) {
            n.pointer.redraws += 1;
        }
    }

    pub fn redraw_count(&self, id: usize) -> u32 {
        self.node(id).map(|n| n.pointer.redraws).unwrap_or(0)
    }

    /// A canvas's drawable size, from `set_size_request` / `set_content_size`.
    pub fn canvas_size(&self, id: usize) -> (i32, i32) {
        self.node(id).map(|n| (n.props.width.unwrap_or(0), n.props.height.unwrap_or(0))).unwrap_or((0, 0))
    }

    /// Run a canvas's draw callback against a recording context and return the
    /// recorded ops. This is how a test asserts *what a draw callback paints*
    /// without a display server — the same trick
    /// `backends::pancurses_draw::render_model_to_grid` plays for the terminal.
    pub fn draw_canvas(&mut self, id: usize) -> Vec<crate::backends::headless::DrawOp> {
        let mut ctx = crate::backends::headless::RecordingDrawContext::new();
        let (w, h) = self.canvas_size(id);
        let mut cb = match self.node_mut(id) {
            Some(n) => n.draw.take(),
            None => return Vec::new(),
        };
        if let Some(cb) = cb.as_mut() {
            cb(&mut ctx, w, h);
        }
        // Put the callback back so the canvas stays drawable.
        if let Some(n) = self.node_mut(id) {
            n.draw = cb;
        }
        ctx.ops
    }

    // -- Pointer / key hooks --

    pub fn add_click_hook(&mut self, id: usize, cb: Box<dyn FnMut(f64, f64)>) {
        if let Some(n) = self.node_mut(id) {
            n.pointer.click.push(cb);
        }
    }

    pub fn add_click_button_hook(&mut self, id: usize, cb: Box<dyn FnMut(f64, f64, u32, u32)>) {
        if let Some(n) = self.node_mut(id) {
            n.pointer.click_button.push(cb);
        }
    }

    pub fn add_motion_hook(&mut self, id: usize, cb: Box<dyn FnMut(f64, f64, u32)>) {
        if let Some(n) = self.node_mut(id) {
            n.pointer.motion.push(cb);
        }
    }

    pub fn add_release_hook(&mut self, id: usize, cb: Box<dyn FnMut(f64, f64, u32, u32)>) {
        if let Some(n) = self.node_mut(id) {
            n.pointer.release.push(cb);
        }
    }

    pub fn add_key_hook(&mut self, id: usize, cb: Box<dyn FnMut(u32) -> bool>) {
        if let Some(n) = self.node_mut(id) {
            n.pointer.key.push(cb);
        }
    }

    pub fn add_key_raw_hook(&mut self, id: usize, cb: Box<dyn FnMut(u32, u32) -> bool>) {
        if let Some(n) = self.node_mut(id) {
            n.pointer.key_raw.push(cb);
        }
    }

    pub fn set_screen_origin(&mut self, id: usize, origin: Option<(i32, i32)>) {
        if let Some(n) = self.node_mut(id) {
            n.pointer.screen_origin = origin;
        }
    }

    /// Fire a click at model coordinates. Runs the node's pointer hooks and its
    /// generic callbacks, with the borrow released before each invocation so a
    /// hook may re-enter the model.
    pub fn pointer_click(&mut self, id: usize, x: f64, y: f64) {
        let mut cbs = self.node_mut(id).map(|n| std::mem::take(&mut n.pointer.click)).unwrap_or_default();
        for cb in cbs.iter_mut() {
            cb(x, y);
        }
        if let Some(n) = self.node_mut(id) {
            n.pointer.click = cbs;
        }
        self.click(id);
    }

    /// Fire a button/modifier-aware click, then the plain click hooks.
    pub fn pointer_click_button(&mut self, id: usize, x: f64, y: f64, button: u32, state: u32) {
        let mut cbs = self.node_mut(id).map(|n| std::mem::take(&mut n.pointer.click_button)).unwrap_or_default();
        for cb in cbs.iter_mut() {
            cb(x, y, button, state);
        }
        if let Some(n) = self.node_mut(id) {
            n.pointer.click_button = cbs;
        }
        self.pointer_click(id, x, y);
    }

    /// Fire a motion event. Returns whether any hook consumed it.
    pub fn pointer_motion(&mut self, id: usize, x: f64, y: f64, state: u32) {
        let mut cbs = self.node_mut(id).map(|n| std::mem::take(&mut n.pointer.motion)).unwrap_or_default();
        for cb in cbs.iter_mut() {
            cb(x, y, state);
        }
        if let Some(n) = self.node_mut(id) {
            n.pointer.motion = cbs;
        }
    }

    /// Fire a button release.
    pub fn pointer_release(&mut self, id: usize, x: f64, y: f64, button: u32, state: u32) {
        let mut cbs = self.node_mut(id).map(|n| std::mem::take(&mut n.pointer.release)).unwrap_or_default();
        for cb in cbs.iter_mut() {
            cb(x, y, button, state);
        }
        if let Some(n) = self.node_mut(id) {
            n.pointer.release = cbs;
        }
    }

    /// Deliver a key to the canvas. Returns true when a hook consumed it.
    /// The plain `key` hooks run first; if one returns true the `key_raw` hooks
    /// are skipped, matching the adapter chain order.
    pub fn key(&mut self, id: usize, keyval: u32) -> bool {
        let mut consumed = false;
        let mut cbs = self.node_mut(id).map(|n| std::mem::take(&mut n.pointer.key)).unwrap_or_default();
        for cb in cbs.iter_mut() {
            if cb(keyval) {
                consumed = true;
            }
        }
        if let Some(n) = self.node_mut(id) {
            n.pointer.key = cbs;
        }
        if consumed {
            return true;
        }
        let mut raw = self.node_mut(id).map(|n| std::mem::take(&mut n.pointer.key_raw)).unwrap_or_default();
        for cb in raw.iter_mut() {
            if cb(keyval, 0) {
                consumed = true;
            }
        }
        if let Some(n) = self.node_mut(id) {
            n.pointer.key_raw = raw;
        }
        consumed
    }

    // -- Overlay --

    /// Push a child onto an overlay's layer stack, above the base child.
    pub fn overlay_add(&mut self, overlay_id: usize, child_id: usize) {
        if overlay_id == child_id {
            return;
        }
        if let Some(old) = self.node(child_id).and_then(|n| n.parent) {
            if let Some(p) = self.node_mut(old) {
                p.children.retain(|c| *c != child_id);
            }
        }
        self.nodes.iter_mut().for_each(|n| {
            n.children.retain(|c| *c != child_id);
            n.overlays.retain(|c| *c != child_id);
        });
        if let Some(n) = self.node_mut(overlay_id) {
            if !n.overlays.contains(&child_id) {
                n.overlays.push(child_id);
            }
        }
        if let Some(c) = self.node_mut(child_id) {
            c.parent = Some(overlay_id);
        }
    }

    /// Set whether pointer events pass through an overlay layer to the base.
    pub fn overlay_set_pass_through(&mut self, overlay_id: usize, child_id: usize, pass: bool) {
        if let Some(n) = self.node_mut(overlay_id) {
            n.overlay_pass_through.insert(child_id, pass);
        }
    }

    /// The overlay's layers, base child first, then the stacked overlays.
    pub fn overlay_layers(&self, id: usize) -> Vec<usize> {
        self.node(id).map(|n| n.overlays.clone()).unwrap_or_default()
    }

    /// Make every descendant visible. Mirrors `GtkOverlay::show_all`.
    pub fn show_all(&mut self, id: usize) {
        let mut stack = vec![id];
        while let Some(cur) = stack.pop() {
            let kids = self.node(cur).map(|n| n.children.clone()).unwrap_or_default();
            if let Some(n) = self.node_mut(cur) {
                n.props.visible = true;
            }
            stack.extend(kids);
        }
    }

    // -- Dialog --

    pub fn dialog_add_button(&mut self, id: usize, label: &str, response_id: i32) {
        if let Some(n) = self.node_mut(id) {
            n.dialog_buttons.push((label.to_string(), response_id));
        }
    }

    pub fn dialog_set_default_response(&mut self, id: usize, response_id: i32) {
        if let Some(n) = self.node_mut(id) {
            n.dialog_default_response = Some(response_id);
        }
    }

    pub fn dialog_buttons(&self, id: usize) -> Vec<(String, i32)> {
        self.node(id).map(|n| n.dialog_buttons.clone()).unwrap_or_default()
    }

    pub fn dialog_set_transient_for(&mut self, id: usize, parent: Option<usize>) {
        if let Some(n) = self.node_mut(id) {
            n.dialog_transient_for = parent;
        }
    }

    pub fn dialog_transient_for(&self, id: usize) -> Option<usize> {
        self.node(id).and_then(|n| n.dialog_transient_for)
    }

    /// Mark a dialog as destroyed and hide it.
    pub fn dialog_mark_destroyed(&mut self, id: usize) {
        if let Some(n) = self.node_mut(id) {
            n.destroyed = true;
            n.props.visible = false;
        }
    }

    /// Register a dialog response handler, called as `cb(response_id)`.
    pub fn add_response_callback(&mut self, id: usize, cb: Box<dyn FnMut(i32)>) {
        if let Some(n) = self.node_mut(id) {
            n.response_cbs.push(cb);
        }
    }

    /// Deliver a dialog response. Only a response id registered with
    /// [`Self::dialog_add_button`] reaches the handlers — an unregistered id is
    /// ignored rather than silently firing, which is what the old
    /// `connect_response` no-op could not express.
    pub fn dialog_respond(&mut self, id: usize, response_id: i32) {
        let known = self
            .node(id)
            .map(|n| n.dialog_buttons.iter().any(|(_, r)| *r == response_id))
            .unwrap_or(false);
        if !known {
            return;
        }
        let mut cbs = self.node_mut(id).map(|n| std::mem::take(&mut n.response_cbs)).unwrap_or_default();
        for cb in cbs.iter_mut() {
            cb(response_id);
        }
        if let Some(n) = self.node_mut(id) {
            n.response_cbs = cbs;
        }
    }

    // -- Spreadsheet --

    pub fn sheet_set_cell(&mut self, id: usize, row: u32, col: u32, text: &str, style: u8) {
        if let Some(n) = self.node_mut(id) {
            n.cells.insert((row, col), SheetCell { text: text.to_string(), style, raw: false });
        }
    }

    pub fn sheet_set_raw_cell(&mut self, id: usize, row: u32, col: u32, text: &str) {
        if let Some(n) = self.node_mut(id) {
            n.cells.insert((row, col), SheetCell { text: text.to_string(), style: 0, raw: true });
        }
    }

    pub fn sheet_get_cell(&self, id: usize, row: u32, col: u32) -> Option<String> {
        self.node(id).and_then(|n| n.cells.get(&(row, col)).map(|c| c.text.clone()))
    }

    pub fn sheet_set_border_title(&mut self, id: usize, title: &str) {
        if let Some(n) = self.node_mut(id) {
            if let ZorkKind::Spreadsheet { border_title } = &mut n.kind {
                *border_title = title.to_string();
            }
        }
    }

    pub fn sheet_border_title(&self, id: usize) -> Option<String> {
        self.node(id).and_then(|n| match &n.kind {
            ZorkKind::Spreadsheet { border_title } => Some(border_title.clone()),
            _ => None,
        })
    }

    // -- Menu / action dispatch --

    /// Fire the `SimpleAction` a menu item names, if the action was created
    /// with [`Self::create_simple_action`]. Returns true when an action fired.
    ///
    /// The old `Harness::select` fired the *menu node's* own callbacks, which
    /// meant a menu item's action string was decorative. Resolving the name is
    /// what makes the menu model meaningful.
    pub fn activate_action(&mut self, action_name: &str) -> bool {
        let id = match self.action_names.get(action_name) {
            Some(id) => *id,
            None => return false,
        };
        self.click(id);
        true
    }

    /// Select menu item `index` (0-based) on `menu_id` and dispatch its action.
    ///
    /// Returns the item's label, or `None` when the index is out of range or
    /// the item is a separator/section (which are not selectable).
    pub fn menu_select(&mut self, menu_id: usize, index: usize) -> Option<String> {
        let item = self.menu_items.get(&menu_id).and_then(|i| i.get(index)).cloned()?;
        if item.enabled == false || matches!(item.kind, MenuItemKind::Separator | MenuItemKind::Section) {
            return None;
        }
        // A check/radio item flips before its action runs, so a handler that
        // reads the menu sees the new state.
        if matches!(item.kind, MenuItemKind::Check | MenuItemKind::Radio) {
            let new_state = !item.checked;
            self.set_menu_item_checked(menu_id, index, new_state);
        }
        if !item.action.is_empty() {
            self.activate_action(&item.action);
        }
        Some(item.label)
    }

    /// Whether a keyboard menu is open. The model has no popup concept, so this
    /// is always false — matching the NWG/pancurses no-op — but it exists so
    /// the adapter can answer honestly instead of omitting the method.
    pub fn menu_active(&self) -> bool {
        false
    }

    // -- Fire --

    /// Fire the callbacks registered on a node. Does nothing for a missing id.
    ///
    /// This must be called outside any borrow of `self` that the callback might
    /// re-take; the implementation releases its borrow before invoking each
    /// callback.
    pub fn click(&mut self, id: usize) {
        let mut cbs = match self.node_mut(id) {
            Some(n) => std::mem::take(&mut n.callbacks),
            None => return,
        };
        for cb in cbs.iter_mut() {
            cb();
        }
        if let Some(n) = self.node_mut(id) {
            n.callbacks = cbs;
        }
    }

    /// Convenience: type into an entry and fire `changed` callbacks.
    pub fn type_into(&mut self, id: usize, text: &str) {
        self.set_entry_text(id, text);
        self.click(id);
    }

    /// Convenience: toggle a check/radio button and fire its callbacks.
    pub fn toggle(&mut self, id: usize) {
        let became = match self.node_mut(id) {
            Some(n) => match &mut n.kind {
                ZorkKind::CheckButton { ref mut checked, .. } => {
                    *checked = !*checked;
                    Some(*checked)
                }
                ZorkKind::RadioButton { ref mut checked, .. } => {
                    *checked = !*checked;
                    Some(*checked)
                }
                _ => None,
            },
            None => None,
        };
        if became.is_some() {
            // Re-apply through the radio-aware setter so the group stays
            // mutually exclusive.
            if let Some(ZorkKind::RadioButton { .. }) = self.node(id).map(|n| &n.kind) {
                let now = self.get_radiobutton_checked(id);
                self.set_radiobutton_checked(id, now);
            }
            self.click(id);
        }
    }

    pub fn quit(&mut self) {
        self.running = false;
    }

    /// Produce a serialization-friendly snapshot of the current model.
    ///
    /// This is the *only* sanctioned JSON surface: it is used for snapshot
    /// regression tests and for external (non-Rust) drivers that want to inspect
    /// the model without parsing the REPL's prose. It deliberately does not
    /// include callbacks — only their counts.
    pub fn snapshot(&self) -> Snapshot {
        let nodes = self
            .nodes
            .iter()
            .map(|n| {
                let mut s = SnapshotNode {
                    id: n.id,
                    kind: match n.kind {
                        ZorkKind::Window { .. } => "Window",
                        ZorkKind::Button { .. } => "Button",
                        ZorkKind::Label { .. } => "Label",
                        ZorkKind::BoxWidget { .. } => "BoxWidget",
                        ZorkKind::Grid { .. } => "Grid",
                        ZorkKind::Entry { .. } => "Entry",
                        ZorkKind::CheckButton { .. } => "CheckButton",
                        ZorkKind::RadioButton { .. } => "RadioButton",
                        ZorkKind::Dialog { .. } => "Dialog",
                        ZorkKind::Menu => "Menu",
                        ZorkKind::MenuBar => "MenuBar",
                        ZorkKind::SimpleAction => "SimpleAction",
                        ZorkKind::DropDown { .. } => "DropDown",
                        ZorkKind::TextView { .. } => "TextView",
                        ZorkKind::Canvas { .. } => "Canvas",
                        ZorkKind::Overlay => "Overlay",
                        ZorkKind::ScrolledWindow => "ScrolledWindow",
                        ZorkKind::Fixed => "Fixed",
                        ZorkKind::Application => "Application",
                        ZorkKind::Spreadsheet { .. } => "Spreadsheet",
                    }
                    .to_string(),
                    parent: n.parent,
                    children: n.children.clone(),
                    title: None,
                    text: None,
                    label: None,
                    checked: None,
                    selected: None,
                    items: None,
                    props: n.props.clone(),
                    can_focus: n.can_focus,
                    callbacks: n.callbacks.len(),
                    click_hooks: n.pointer.click.len() + n.pointer.click_button.len(),
                    redraws: n.pointer.redraws,
                    cells: n
                        .cells
                        .iter()
                        .map(|((row, col), c)| SheetCellEntry {
                            row: *row,
                            col: *col,
                            text: c.text.clone(),
                            style: c.style,
                            raw: c.raw,
                        })
                        .collect(),
                    dialog_buttons: n.dialog_buttons.clone(),
                    overlays: n.overlays.clone(),
                };
                match &n.kind {
                    ZorkKind::Window { title } | ZorkKind::Dialog { title } => s.title = Some(title.clone()),
                    ZorkKind::Label { text } | ZorkKind::TextView { text } => s.text = Some(text.clone()),
                    ZorkKind::Spreadsheet { border_title } => s.title = Some(border_title.clone()),
                    ZorkKind::Button { label } => s.label = Some(label.clone()),
                    ZorkKind::CheckButton { label, checked } => {
                        s.label = Some(label.clone());
                        s.checked = Some(*checked);
                    }
                    ZorkKind::RadioButton { label, checked, .. } => {
                        s.label = Some(label.clone());
                        s.checked = Some(*checked);
                    }
                    ZorkKind::DropDown { items, selected } => {
                        s.items = Some(items.clone());
                        s.selected = *selected;
                    }
                    _ => {}
                }
                s
            })
            .collect();
        let mut cells = HashMap::new();
        for n in &self.nodes {
            for (k, v) in &n.cells {
                cells.insert(*k, v.clone());
            }
        }
        let mut overlay_pass_through = HashMap::new();
        for n in &self.nodes {
            for (k, v) in &n.overlay_pass_through {
                overlay_pass_through.insert(*k, *v);
            }
        }
        Snapshot {
            nodes,
            menu_items: self.menu_items.clone(),
            current_id: self.current_id,
            running: self.running,
            focused: self.focused,
            overlay_pass_through,
            cells,
        }
    }
}

/// Clamp a scroll value into `[0, upper - page]`, treating a degenerate range
/// as `0` so a caller cannot scroll into a negative offset.
fn clamp_scroll(val: f64, upper: f64, page: f64) -> f64 {
    let max = (upper - page).max(0.0);
    if !val.is_finite() {
        return 0.0;
    }
    val.clamp(0.0, max)
}

// -- Thread-local singleton used by the adapter shim --

thread_local! {
    static ZORK_STATE: RefCell<ZorkState> = RefCell::new(ZorkState::new());
}

pub(crate) fn with_state<F, R>(f: F) -> R
where
    F: FnOnce(&mut ZorkState) -> R,
{
    ZORK_STATE.with(|s| f(&mut s.borrow_mut()))
}

/// Reset the thread-local singleton. Only for tests that need a clean slate
/// between cases.
#[cfg(test)]
pub(crate) fn reset_state() {
    ZORK_STATE.with(|s| *s.borrow_mut() = ZorkState::new());
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn radio_group_is_mutually_exclusive() {
        let mut s = ZorkState::new();
        s.create_window();
        let a = s.create_radiobutton(Some(7), "a");
        let b = s.create_radiobutton(Some(7), "b");
        s.set_radiobutton_checked(a, true);
        assert!(s.get_radiobutton_checked(a));
        s.set_radiobutton_checked(b, true);
        assert!(s.get_radiobutton_checked(b));
        assert!(!s.get_radiobutton_checked(a), "sibling in the same group must clear");
    }

    #[test]
    fn append_child_keeps_existing_children() {
        let mut s = ZorkState::new();
        s.create_window();
        let bx = s.create_box(true, 0);
        let a = s.create_button("a");
        let b = s.create_button("b");
        s.append_child(bx, a);
        s.append_child(bx, b);
        assert_eq!(s.node(bx).unwrap().children, vec![a, b]);
    }

    #[test]
    fn set_child_reparents_without_duplicating() {
        let mut s = ZorkState::new();
        s.create_window();
        let bx1 = s.create_box(true, 0);
        let bx2 = s.create_box(true, 0);
        let a = s.create_button("a");
        s.append_child(bx1, a);
        s.set_child(bx2, a);
        assert_eq!(s.node(bx1).unwrap().children, Vec::<usize>::new());
        assert_eq!(s.node(bx2).unwrap().children, vec![a]);
        assert_eq!(s.node(a).unwrap().parent, Some(bx2));
        // No other node may still list it.
        let listed: Vec<usize> = s.nodes.iter().filter(|n| n.children.contains(&a)).map(|n| n.id).collect();
        assert_eq!(listed, vec![bx2]);
    }

    #[test]
    fn layout_box_positions_children() {
        let mut s = ZorkState::new();
        s.create_window();
        let bx = s.create_box(true, 4);
        let a = s.create_button("a");
        let b = s.create_button("b");
        s.append_child(bx, a);
        s.append_child(bx, b);
        s.set_size_request(a, 10, 6);
        s.set_size_request(b, 20, 6);
        s.layout_box(bx, 0, 0, 200, 100);
        assert_eq!(s.get_offset(a), Some((Some(0), Some(0))));
        assert_eq!(s.get_offset(b), Some((Some(14), Some(0))), "second child sits after the first plus spacing");
    }

    #[test]
    fn layout_box_skips_hidden_children() {
        let mut s = ZorkState::new();
        s.create_window();
        let bx = s.create_box(true, 0);
        let a = s.create_button("a");
        let b = s.create_button("b");
        s.append_child(bx, a);
        s.append_child(bx, b);
        s.set_size_request(a, 10, 6);
        s.set_size_request(b, 10, 6);
        s.set_visible(a, false);
        s.layout_box(bx, 0, 0, 100, 10);
        assert_eq!(s.get_offset(b), Some((Some(0), Some(0))), "hidden first child takes no room");
    }

    #[test]
    fn measure_box_uses_children() {
        let mut s = ZorkState::new();
        s.create_window();
        let bx = s.create_box(false, 2);
        let a = s.create_button("a");
        let b = s.create_button("b");
        s.append_child(bx, a);
        s.append_child(bx, b);
        s.set_size_request(a, 30, 10);
        s.set_size_request(b, 30, 20);
<<<<<<< HEAD
        // Vertical: main axis is height (10+20+2), cross is max width (30).
=======
        // A vertical box packs along y: height = 10 + 20 + 2*spacing,
        // width = the widest child (30).
>>>>>>> 184a0b72 (feat(zork): make the zork backend build and cover the GTK/NWG surface)
        assert_eq!(s.measure_box(bx), Some((30, 32)));
    }

    #[test]
<<<<<<< HEAD
=======
    fn measure_box_horizontal_sums_widths() {
        let mut s = ZorkState::new();
        s.create_window();
        let bx = s.create_box(true, 4);
        let a = s.create_button("a");
        let b = s.create_button("b");
        s.append_child(bx, a);
        s.append_child(bx, b);
        s.set_size_request(a, 10, 6);
        s.set_size_request(b, 20, 9);
        // Horizontal packs along x: width = 10 + 20 + spacing, height = tallest.
        assert_eq!(s.measure_box(bx), Some((34, 9)));
    }

    #[test]
>>>>>>> 184a0b72 (feat(zork): make the zork backend build and cover the GTK/NWG surface)
    fn grid_attach_grows_grid() {
        let mut s = ZorkState::new();
        s.create_window();
        let g = s.create_grid();
        let b = s.create_button("b");
        s.grid_attach(g, b, 2, 3, 1, 1);
        assert_eq!(s.get_offset(b), Some((Some(2), Some(3))));
        match s.node(g).unwrap().kind {
            ZorkKind::Grid { cols, rows } => assert_eq!((cols, rows), (3, 4)),
            _ => panic!("not a grid"),
        }
    }

    #[test]
    fn scroll_is_clamped_to_document() {
        let mut s = ZorkState::new();
        s.create_window();
        let sw = s.create_scrolled_window();
        s.set_scroll(sw, 0.0, 100.0, 20.0, 999.0, 100.0, 20.0);
        assert_eq!(s.get_scroll(sw), (0.0, 80.0));
    }

    #[test]
    fn focus_moves_and_is_queryable() {
        let mut s = ZorkState::new();
        s.create_window();
        let e = s.create_entry();
        assert!(!s.has_focus(e));
        s.set_focus(e);
        assert!(s.has_focus(e));
        assert_eq!(s.get_focus(), Some(e));
    }

    #[test]
    fn focus_refused_for_invisible_node() {
        let mut s = ZorkState::new();
        s.create_window();
        let e = s.create_entry();
        s.set_visible(e, false);
        s.set_focus(e);
        assert_eq!(s.get_focus(), None);
    }

    #[test]
    fn classes_add_remove_and_dedup() {
        let mut s = ZorkState::new();
        s.create_window();
        let b = s.create_button("x");
        s.add_class(b, "primary");
        s.add_class(b, "primary");
        assert_eq!(s.classes(b), vec!["primary".to_string()]);
        s.add_class(b, "flat");
        s.remove_class(b, "primary");
        assert_eq!(s.classes(b), vec!["flat".to_string()]);
        assert!(!s.has_class(b, "primary"));
    }

    #[test]
    fn key_hooks_stop_at_first_consumer() {
        use std::cell::Cell;
        use std::rc::Rc;
        let mut s = ZorkState::new();
        s.create_window();
        let c = s.create_canvas();
        let raw_hits = Rc::new(Cell::new(0));
        s.add_key_hook(c, Box::new(|_| true));
        {
            let r = raw_hits.clone();
            s.add_key_raw_hook(c, Box::new(move |_, _| { r.set(r.get() + 1); true }));
        }
        assert!(s.key(c, b'a' as u32));
        assert_eq!(raw_hits.get(), 0, "key_raw must be skipped when key consumed");
    }

    #[test]
    fn draw_canvas_records_ops() {
        let mut s = ZorkState::new();
        s.create_window();
        let c = s.create_canvas();
        s.set_size_request(c, 100, 50);
        s.set_draw_callback(c, Box::new(|ctx, w, h| {
            assert_eq!((w, h), (100, 50));
            ctx.fill_rect(0.0, 0.0, 10.0, 10.0, 1.0, 0.0, 0.0, 1.0);
        }));
        let ops = s.draw_canvas(c);
        assert_eq!(ops.len(), 1);
        // The callback must still be installed afterwards.
        assert!(s.has_draw_callback(c));
        assert_eq!(s.draw_canvas(c).len(), 1);
    }

    #[test]
    fn menu_select_dispatches_named_action() {
        use std::cell::Cell;
        use std::rc::Rc;
        let mut s = ZorkState::new();
        s.create_window();
        let act = s.create_simple_action("app.open");
        let menu = s.create_menu();
        s.menu_append(menu, "Open", "app.open");
        let fired = Rc::new(Cell::new(false));
        {
            let f = fired.clone();
            s.add_callback(act, Box::new(move || f.set(true)));
        }
        assert_eq!(s.menu_select(menu, 0), Some("Open".to_string()));
        assert!(fired.get());
    }

    #[test]
    fn menu_select_refuses_separator() {
        let mut s = ZorkState::new();
        s.create_window();
        let menu = s.create_menu();
        s.menu_append_separator(menu, "");
        s.menu_append(menu, "Open", "app.open");
        assert_eq!(s.menu_select(menu, 0), None);
        assert_eq!(s.menu_select(menu, 1), Some("Open".to_string()));
    }

    #[test]
    fn menu_check_item_toggles_before_dispatch() {
        use std::cell::Cell;
        use std::rc::Rc;
        let mut s = ZorkState::new();
        s.create_window();
        let act = s.create_simple_action("app.wrap");
        let menu = s.create_menu();
        s.menu_append_check(menu, "Wrap", "app.wrap", false);
        let seen = Rc::new(Cell::new(false));
        {
            let sv = seen.clone();
            let st = s.menu_items[&menu].clone();
            let _ = st;
            s.add_callback(act, Box::new(move || {
                // The action sees the already-flipped state.
                sv.set(true);
            }));
        }
        s.menu_select(menu, 0);
        assert!(seen.get());
        assert!(s.menu_items[&menu][0].checked);
        s.menu_select(menu, 0);
        assert!(!s.menu_items[&menu][0].checked);
    }

    #[test]
    fn radio_menu_items_are_exclusive() {
        let mut s = ZorkState::new();
        s.create_window();
        let menu = s.create_menu();
        let i0 = s.menu_append_radio(menu, "a", "", 1, false);
        let i1 = s.menu_append_radio(menu, "b", "", 1, false);
        s.set_menu_item_checked(menu, i1, true);
        assert!(s.menu_items[&menu][i1].checked);
        assert!(!s.menu_items[&menu][i0].checked);
    }

    #[test]
    fn overlay_layers_and_pass_through() {
        let mut s = ZorkState::new();
        s.create_window();
        let ov = s.create_overlay();
        let base = s.create_label("base");
        let layer = s.create_button("menu");
        s.set_child(ov, base);
        s.overlay_add(ov, layer);
        assert_eq!(s.node(ov).unwrap().children, vec![base]);
        assert_eq!(s.overlay_layers(ov), vec![layer]);
        s.overlay_set_pass_through(ov, layer, true);
        let snap = s.snapshot();
        assert_eq!(snap.overlay_pass_through.get(&layer), Some(&true));
    }

    #[test]
    fn show_all_makes_descendants_visible() {
        let mut s = ZorkState::new();
        s.create_window();
        let bx = s.create_box(true, 0);
        let l = s.create_label("x");
        s.append_child(bx, l);
        s.set_visible(l, false);
        s.show_all(bx);
        assert!(s.get_visible(bx));
        assert!(s.get_visible(l));
    }

    #[test]
    fn dialog_response_requires_a_registered_button() {
        use std::cell::Cell;
        use std::rc::Rc;
        let mut s = ZorkState::new();
        s.create_window();
        let d = s.create_dialog();
        let fired = Rc::new(Cell::new(0));
        {
            let f = fired.clone();
<<<<<<< HEAD
            s.add_callback(d, Box::new(move || f.set(f.get() + 1)));
=======
            s.add_response_callback(d, Box::new(move |_| f.set(f.get() + 1)));
>>>>>>> 184a0b72 (feat(zork): make the zork backend build and cover the GTK/NWG surface)
        }
        s.dialog_respond(d, 1);
        assert_eq!(fired.get(), 0, "unregistered response must not fire");
        s.dialog_add_button(d, "OK", 1);
        s.dialog_respond(d, 1);
        assert_eq!(fired.get(), 1);
    }

    #[test]
    fn dialog_response_handler_receives_the_id() {
        use std::cell::RefCell;
        use std::rc::Rc;
        let mut s = ZorkState::new();
        s.create_window();
        let d = s.create_dialog();
        s.dialog_add_button(d, "OK", 1);
        s.dialog_add_button(d, "Cancel", 2);
        let seen = Rc::new(RefCell::new(Vec::new()));
        {
            let sv = seen.clone();
            s.add_response_callback(d, Box::new(move |r| sv.borrow_mut().push(r)));
        }
        s.dialog_respond(d, 2);
        s.dialog_respond(d, 7);
        assert_eq!(*seen.borrow(), vec![2], "unregistered id must not reach the handler");
    }

    #[test]
    fn scroll_callback_drives_offset() {
        use std::cell::RefCell;
        use std::rc::Rc;
        let mut s = ZorkState::new();
        s.create_window();
        let sw = s.create_scrolled_window();
        let seen = Rc::new(RefCell::new(Vec::new()));
        {
            let sv = seen.clone();
            s.add_scroll_callback(sw, Box::new(move |vert, pos| sv.borrow_mut().push((vert, pos))));
        }
        s.scroll(sw, true, 42.0);
        s.scroll(sw, false, 7.0);
        assert_eq!(*seen.borrow(), vec![(true, 42.0), (false, 7.0)]);
        assert_eq!(s.get_scroll(sw), (7.0, 42.0));
    }

    #[test]
    fn entry_caret_is_tracked() {
        let mut s = ZorkState::new();
        s.create_window();
        let e = s.create_entry();
        s.set_entry_text(e, "hello");
        assert_eq!(s.get_entry_position(e), Some(5));
        s.set_entry_position(e, 99);
        assert_eq!(s.get_entry_position(e), Some(5), "clamped to buffer length");
        s.set_entry_position(e, 2);
        assert_eq!(s.get_entry_position(e), Some(2));
    }

    #[test]
    fn dropdown_selection_is_dropped_when_items_shrink() {
        let mut s = ZorkState::new();
        s.create_window();
        let d = s.create_dropdown(&["a", "b", "c"]);
        s.set_dropdown_selected(d, 2);
        assert_eq!(s.get_dropdown_selected(d), 2);
        s.set_dropdown_items(d, &["x"]);
        assert_eq!(s.get_dropdown_selected(d), -1);
    }

    #[test]
    fn snapshot_includes_properties() {
        let mut s = ZorkState::new();
        s.create_window();
        let b = s.create_button("x");
        s.set_size_request(b, 10, 20);
        s.set_visible(b, false);
        s.add_class(b, "flat");
        s.set_hexpand(b, true);
        let snap = s.snapshot();
        let n = snap.nodes.iter().find(|n| n.id == b).unwrap();
        assert_eq!(n.props.width, Some(10));
        assert!(!n.props.visible);
        assert_eq!(n.props.classes, vec!["flat".to_string()]);
        assert!(n.props.hexpand);
    }

    #[test]
    fn spreadsheet_cells_round_trip() {
        let mut s = ZorkState::new();
        s.create_window();
        let sh = s.create_spreadsheet();
        s.sheet_set_cell(sh, 1, 1, "42", 0);
        s.sheet_set_raw_cell(sh, 1, 2, "=1+1");
        s.sheet_set_border_title(sh, "Sheet1");
        assert_eq!(s.sheet_get_cell(sh, 1, 1), Some("42".to_string()));
        assert_eq!(s.sheet_get_cell(sh, 1, 2), Some("=1+1".to_string()));
        assert_eq!(s.sheet_border_title(sh), Some("Sheet1".to_string()));
        assert!(s.snapshot().cells.contains_key(&(1, 2)));
    }

    #[test]
    fn append_textview_is_incremental() {
        let mut s = ZorkState::new();
        s.create_window();
        let tv = s.create_textview();
        s.append_textview_text(tv, "a");
        s.append_textview_text(tv, "b");
        assert_eq!(s.get_textview_text(tv), Some("ab".to_string()));
    }

    #[test]
    fn fixed_width_sets_size_request() {
        let mut s = ZorkState::new();
        s.create_window();
        let l = s.create_label("x");
        s.set_fixed_width(l, Some(120));
        assert_eq!(s.get_size_request(l), Some((Some(120), None)));
    }

    #[test]
    fn toggle_keeps_radio_group_exclusive() {
        let mut s = ZorkState::new();
        s.create_window();
        let a = s.create_radiobutton(Some(3), "a");
        let b = s.create_radiobutton(Some(3), "b");
        s.toggle(a);
        assert!(s.get_radiobutton_checked(a));
        s.toggle(b);
        assert!(s.get_radiobutton_checked(b));
        assert!(!s.get_radiobutton_checked(a));
    }
}
