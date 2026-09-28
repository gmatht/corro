//! Backwards-compatible free-function facade over [`super::model`].
//!
//! The per-widget adapter ([`crate::backends_zork_adapter`]) was written against
//! the old `backends::zork::*` free functions, which mutably operated on a
//! thread-local model. This module re-exposes those names (each delegating to
//! [`super::model::with_state`]) so the adapter is unchanged. New code should use
//! the typed [`super::harness`] instead.
//!
//! The model grew a property bag ([`super::model::ZorkProps`]) and six new
//! node kinds, so the surface here is much wider than the original
//! label/button/entry set: every property the GUI adapters expose has a
//! free function so the adapter shim stays a thin `id`-passing layer.

use std::os::raw::c_void;

pub use super::model::{MenuItemData, MenuItemKind};
use super::model::with_state;

pub type Callback = Box<dyn FnMut()>;

// -- Factories --

pub fn create_window() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_window()))
}
pub fn create_button(label: &str) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_button(label)))
}
pub fn create_label(text: &str) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_label(text)))
}
pub fn create_box(horizontal: bool, spacing: i32) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_box(horizontal, spacing)))
}
pub fn create_grid() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_grid()))
}
pub fn create_entry() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_entry()))
}
pub fn create_menu() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_menu()))
}
pub fn create_simple_action(name: &str) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_simple_action(name)))
}
pub fn create_menubar(model_id: usize, action_group: *mut c_void) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_menubar(model_id, action_group)))
}
pub fn create_dialog() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_dialog()))
}
pub fn create_dropdown(items: &[&str]) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_dropdown(items)))
}
pub fn create_checkbutton(label: &str) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_checkbutton(label)))
}
pub fn create_radiobutton(group_id: Option<usize>, label: &str) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_radiobutton(group_id, label)))
}
pub fn create_textview() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_textview()))
}
pub fn create_canvas() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_canvas()))
}
pub fn create_overlay() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_overlay()))
}
pub fn create_scrolled_window() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_scrolled_window()))
}
pub fn create_fixed() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_fixed()))
}
pub fn create_application() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_application()))
}
pub fn create_spreadsheet() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
    Ok(with_state(|s| s.create_spreadsheet()))
}

// -- Menu model --

pub fn menu_append(menu_id: usize, label: &str, action: &str) {
    with_state(|s| s.menu_append(menu_id, label, action));
}
pub fn menu_append_submenu(menu_id: usize, label: &str, submenu_id: usize) {
    with_state(|s| s.menu_append_submenu(menu_id, label, submenu_id));
}
pub fn menu_append_with_shortcut(menu_id: usize, label: &str, action: &str, accelerator: &str) {
    with_state(|s| s.menu_append_with_shortcut(menu_id, label, action, accelerator));
}
pub fn menu_append_section(menu_id: usize, label: &str) {
    with_state(|s| s.menu_append_section(menu_id, label));
}
pub fn menu_append_separator(menu_id: usize, label: &str) {
    with_state(|s| s.menu_append_separator(menu_id, label));
}
pub fn menu_append_check(menu_id: usize, label: &str, action: &str, checked: bool) -> usize {
    with_state(|s| s.menu_append_check(menu_id, label, action, checked))
}
pub fn menu_append_radio(menu_id: usize, label: &str, action: &str, group: usize, checked: bool) -> usize {
    with_state(|s| s.menu_append_radio(menu_id, label, action, group, checked))
}
pub fn set_menu_item_checked(menu_id: usize, index: usize, checked: bool) {
    with_state(|s| s.set_menu_item_checked(menu_id, index, checked));
}
pub fn menu_items(menu_id: usize) -> Vec<MenuItemData> {
    with_state(|s| s.menu_items.get(&menu_id).cloned().unwrap_or_default())
}
/// Select a menu item by 0-based index and dispatch the action it names.
/// Returns the item's label, or `None` when the index is invalid or the item
/// is a non-selectable separator/section.
pub fn menu_select(menu_id: usize, index: usize) -> Option<String> {
    with_state(|s| s.menu_select(menu_id, index))
}
pub fn activate_action(action_name: &str) -> bool {
    with_state(|s| s.activate_action(action_name))
}

// -- Text / value setters and getters --

pub fn set_window_title(id: usize, title: &str) {
    with_state(|s| s.set_window_title(id, title));
}
pub fn get_window_title(id: usize) -> Option<String> {
    with_state(|s| s.get_window_title(id))
}
pub fn set_label_text(id: usize, text: &str) {
    with_state(|s| s.set_label_text(id, text));
}
pub fn get_label_text(id: usize) -> Option<String> {
    with_state(|s| s.get_label_text(id))
}
pub fn set_button_label(id: usize, label: &str) {
    with_state(|s| s.set_button_label(id, label));
}
pub fn get_button_label(id: usize) -> Option<String> {
    with_state(|s| s.get_button_label(id))
}
pub fn add_callback(id: usize, cb: Callback) {
    with_state(|s| s.add_callback(id, cb));
}
pub fn set_entry_text(id: usize, text: &str) {
    with_state(|s| s.set_entry_text(id, text));
}
pub fn get_entry_text(id: usize) -> Option<String> {
    with_state(|s| s.get_entry_text(id))
}
pub fn get_entry_position(id: usize) -> Option<usize> {
    with_state(|s| s.get_entry_position(id))
}
pub fn set_entry_position(id: usize, pos: usize) {
    with_state(|s| s.set_entry_position(id, pos));
}
pub fn set_textview_text(id: usize, text: &str) {
    with_state(|s| s.set_textview_text(id, text));
}
pub fn get_textview_text(id: usize) -> Option<String> {
    with_state(|s| s.get_textview_text(id))
}
pub fn append_textview_text(id: usize, text: &str) {
    with_state(|s| s.append_textview_text(id, text));
}
pub fn set_editable(id: usize, editable: bool) {
    with_state(|s| s.set_editable(id, editable));
}
pub fn get_editable(id: usize) -> bool {
    with_state(|s| s.get_editable(id))
}
pub fn set_dropdown_items(id: usize, items: &[&str]) {
    with_state(|s| s.set_dropdown_items(id, items));
}
pub fn set_dropdown_selected(id: usize, idx: i32) {
    with_state(|s| s.set_dropdown_selected(id, idx));
}
pub fn get_dropdown_selected(id: usize) -> i32 {
    with_state(|s| s.get_dropdown_selected(id))
}
pub fn get_checkbutton_checked(id: usize) -> bool {
    with_state(|s| s.get_checkbutton_checked(id))
}
pub fn get_radiobutton_checked(id: usize) -> bool {
    with_state(|s| s.get_radiobutton_checked(id))
}
pub fn set_checkbutton_checked(id: usize, checked: bool) {
    with_state(|s| s.set_checkbutton_checked(id, checked));
}
pub fn set_radiobutton_checked(id: usize, checked: bool) {
    with_state(|s| s.set_radiobutton_checked(id, checked));
}

// -- Tree --

pub fn set_child(parent_id: usize, child_id: usize) {
    with_state(|s| s.set_child(parent_id, child_id));
}
pub fn append_child(parent_id: usize, child_id: usize) {
    with_state(|s| s.append_child(parent_id, child_id));
}
pub fn remove_child(parent_id: usize, child_id: usize) {
    with_state(|s| s.remove_child(parent_id, child_id));
}
pub fn layout_box(id: usize, x: i32, y: i32, w: i32, h: i32) {
    with_state(|s| s.layout_box(id, x, y, w, h));
}
pub fn measure_box(id: usize) -> Option<(i32, i32)> {
    with_state(|s| s.measure_box(id))
}
pub fn grid_attach(grid_id: usize, child_id: usize, left: i32, top: i32, width: i32, height: i32) {
    with_state(|s| s.grid_attach(grid_id, child_id, left, top, width, height));
}

// -- Geometry --

pub fn set_offset(id: usize, x: i32, y: i32) {
    with_state(|s| s.set_offset(id, x, y));
}
pub fn get_offset(id: usize) -> Option<(Option<i32>, Option<i32>)> {
    with_state(|s| s.get_offset(id))
}
pub fn set_size_request(id: usize, w: i32, h: i32) {
    with_state(|s| s.set_size_request(id, w, h));
}
pub fn get_size_request(id: usize) -> Option<(Option<i32>, Option<i32>)> {
    with_state(|s| s.get_size_request(id))
}

// -- Presentation properties --

pub fn set_visible(id: usize, visible: bool) {
    with_state(|s| s.set_visible(id, visible));
}
pub fn get_visible(id: usize) -> bool {
    with_state(|s| s.get_visible(id))
}
pub fn show_all(id: usize) {
    with_state(|s| s.show_all(id));
}
pub fn set_hexpand(id: usize, v: bool) {
    with_state(|s| s.set_hexpand(id, v));
}
pub fn set_vexpand(id: usize, v: bool) {
    with_state(|s| s.set_vexpand(id, v));
}
pub fn set_child_hexpand(parent_id: usize, child_id: usize, v: bool) {
    with_state(|s| s.set_child_hexpand(parent_id, child_id, v));
}
pub fn set_child_vexpand(parent_id: usize, child_id: usize, v: bool) {
    with_state(|s| s.set_child_vexpand(parent_id, child_id, v));
}
pub fn set_margin_start(id: usize, px: i32) {
    with_state(|s| s.set_margin_start(id, px));
}
pub fn set_margin_top(id: usize, px: i32) {
    with_state(|s| s.set_margin_top(id, px));
}
pub fn add_class(id: usize, class_name: &str) {
    with_state(|s| s.add_class(id, class_name));
}
pub fn remove_class(id: usize, class_name: &str) {
    with_state(|s| s.remove_class(id, class_name));
}
pub fn classes(id: usize) -> Vec<String> {
    with_state(|s| s.classes(id))
}
pub fn has_class(id: usize, class_name: &str) -> bool {
    with_state(|s| s.has_class(id, class_name))
}
pub fn set_xalign(id: usize, x: f32) {
    with_state(|s| s.set_xalign(id, x));
}
pub fn set_halign(id: usize, align: i32) {
    with_state(|s| s.set_halign(id, align));
}
pub fn set_valign(id: usize, align: i32) {
    with_state(|s| s.set_valign(id, align));
}
pub fn set_fixed_width(id: usize, w: Option<i32>) {
    with_state(|s| s.set_fixed_width(id, w));
}
/// Record a font family + size, plus the style bits the GTK adapter passes
/// through. See [`super::model::font_style`].
pub fn set_font_style(id: usize, family: &str, size: f64, style_bits: i32) {
    let mut parts: Vec<&str> = Vec::new();
    if style_bits & super::model::font_style::BOLD != 0 {
        parts.push("bold");
    }
    if style_bits & super::model::font_style::ITALIC != 0 {
        parts.push("italic");
    }
    if style_bits & super::model::font_style::UNDERLINE != 0 {
        parts.push("underline");
    }
    if style_bits & super::model::font_style::STRIKETHROUGH != 0 {
        parts.push("strikethrough");
    }
    let family = if parts.is_empty() { family.to_string() } else { format!("{} {}", parts.join(" "), family) };
    with_state(|s| s.set_font_style(id, Some(&family), size));
}

// -- Focus --

pub fn set_focus(id: usize) {
    with_state(|s| s.set_focus(id));
}
pub fn get_focus() -> Option<usize> {
    with_state(|s| s.get_focus())
}
pub fn has_focus(id: usize) -> bool {
    with_state(|s| s.has_focus(id))
}
pub fn set_can_focus(id: usize, can: bool) {
    with_state(|s| s.set_can_focus(id, can));
}

// -- Scroll --

pub fn scroll_to(id: usize, hval: f64, hupper: f64, hpage: f64, vval: f64, vupper: f64, vpage: f64) {
    with_state(|s| s.set_scroll(id, hval, hupper, hpage, vval, vupper, vpage));
}
pub fn get_scroll(id: usize) -> (f64, f64) {
    with_state(|s| s.get_scroll(id))
}

// -- Canvas / drawing --

pub fn set_draw_callback(id: usize, cb: Box<dyn FnMut(&mut dyn crate::core::DrawContext, i32, i32)>) {
    with_state(|s| s.set_draw_callback(id, cb));
}
pub fn clear_draw_callback(id: usize) {
    with_state(|s| s.clear_draw_callback(id));
}
pub fn has_draw_callback(id: usize) -> bool {
    with_state(|s| s.has_draw_callback(id))
}
pub fn queue_redraw(id: usize) {
    with_state(|s| s.queue_redraw(id));
}
pub fn redraw_count(id: usize) -> u32 {
    with_state(|s| s.redraw_count(id))
}
pub fn canvas_size(id: usize) -> (i32, i32) {
    with_state(|s| s.canvas_size(id))
}
/// Run the canvas's draw callback against a recording context and return the
/// recorded ops. This is the model-level `Canvas::force_draw` equivalent.
pub fn draw_canvas(id: usize) -> Vec<crate::backends::headless::DrawOp> {
    with_state(|s| s.draw_canvas(id))
}

// -- Pointer / key --

pub fn add_click_hook(id: usize, cb: Box<dyn FnMut(f64, f64)>) {
    with_state(|s| s.add_click_hook(id, cb));
}
pub fn add_click_button_hook(id: usize, cb: Box<dyn FnMut(f64, f64, u32, u32)>) {
    with_state(|s| s.add_click_button_hook(id, cb));
}
pub fn add_motion_hook(id: usize, cb: Box<dyn FnMut(f64, f64, u32)>) {
    with_state(|s| s.add_motion_hook(id, cb));
}
pub fn add_release_hook(id: usize, cb: Box<dyn FnMut(f64, f64, u32, u32)>) {
    with_state(|s| s.add_release_hook(id, cb));
}
pub fn add_key_hook(id: usize, cb: Box<dyn FnMut(u32) -> bool>) {
    with_state(|s| s.add_key_hook(id, cb));
}
pub fn add_key_raw_hook(id: usize, cb: Box<dyn FnMut(u32, u32) -> bool>) {
    with_state(|s| s.add_key_raw_hook(id, cb));
}
pub fn set_screen_origin(id: usize, origin: Option<(i32, i32)>) {
    with_state(|s| s.set_screen_origin(id, origin));
}
pub fn get_screen_origin(id: usize) -> Option<(i32, i32)> {
    with_state(|s| s.node(id).and_then(|n| n.pointer.screen_origin))
}
pub fn pointer_click(id: usize, x: f64, y: f64) {
    with_state(|s| s.pointer_click(id, x, y));
}
pub fn pointer_click_button(id: usize, x: f64, y: f64, button: u32, state: u32) {
    with_state(|s| s.pointer_click_button(id, x, y, button, state));
}
pub fn pointer_motion(id: usize, x: f64, y: f64, state: u32) {
    with_state(|s| s.pointer_motion(id, x, y, state));
}
pub fn pointer_release(id: usize, x: f64, y: f64, button: u32, state: u32) {
    with_state(|s| s.pointer_release(id, x, y, button, state));
}
pub fn key(id: usize, keyval: u32) -> bool {
    with_state(|s| s.key(id, keyval))
}

// -- Scroll callbacks --

pub fn add_scroll_callback(id: usize, cb: Box<dyn FnMut(bool, f64)>) {
    with_state(|s| s.add_scroll_callback(id, cb));
}
/// Drive a `ScrolledWindow`'s scroll handler as if the user had scrolled.
/// `vertical` selects the axis and `pos` is the new value. A headless viewport
/// never produces a real scroll, so this is how the handler gets exercised.
pub fn scroll(id: usize, vertical: bool, pos: f64) {
    with_state(|s| s.scroll(id, vertical, pos));
}

// -- Dialog response callbacks --

/// Register a dialog response handler. The handler receives the response id.
pub fn add_response_callback(id: usize, cb: Box<dyn FnMut(i32)>) {
    with_state(|s| s.add_response_callback(id, cb));
}

// -- Tree queries --

/// The first child of `id`, if any.
pub fn first_child(id: usize) -> Option<usize> {
    with_state(|s| s.node(id).and_then(|n| n.children.first().copied()))
}
/// The grid's `(cols, rows)` as grown by `grid_attach`.
pub fn grid_dimensions(id: usize) -> (usize, usize) {
    with_state(|s| match s.node(id) {
        Some(n) => match n.kind {
            super::model::ZorkKind::Grid { cols, rows } => (cols, rows),
            _ => (0, 0),
        },
        None => (0, 0),
    })
}
/// Whether a dialog has been closed (`Dialog::close`).
pub fn is_destroyed(id: usize) -> bool {
    with_state(|s| s.node(id).map(|n| n.destroyed).unwrap_or(false))
}
pub fn menu_active(id: usize) -> bool {
    with_state(|s| s.menu_active() || s.node(id).is_some_and(|n| n.overlays.len() > 1))
}

/// Fire a node's callbacks (`Button::emit_clicked`, `Harness::click`).
pub fn fire(id: usize) {
    with_state(|s| s.click(id));
}
/// Flip a `CheckButton`/`RadioButton` and fire its callbacks.
pub fn toggle(id: usize) {
    with_state(|s| s.toggle(id));
}

// -- Overlay --

pub fn overlay_add(overlay_id: usize, child_id: usize) {
    with_state(|s| s.overlay_add(overlay_id, child_id));
}
pub fn overlay_remove(overlay_id: usize, child_id: usize) {
    with_state(|s| s.remove_child(overlay_id, child_id));
}
pub fn overlay_set_pass_through(overlay_id: usize, child_id: usize, pass: bool) {
    with_state(|s| s.overlay_set_pass_through(overlay_id, child_id, pass));
}
pub fn overlay_layers(id: usize) -> Vec<usize> {
    with_state(|s| s.overlay_layers(id))
}

// -- Dialog --

pub fn dialog_add_button(id: usize, label: &str, response_id: i32) {
    with_state(|s| s.dialog_add_button(id, label, response_id));
}
pub fn dialog_set_default_response(id: usize, response_id: i32) {
    with_state(|s| s.dialog_set_default_response(id, response_id));
}
pub fn dialog_buttons(id: usize) -> Vec<(String, i32)> {
    with_state(|s| s.dialog_buttons(id))
}
pub fn dialog_set_transient_for(id: usize, parent: Option<usize>) {
    with_state(|s| s.dialog_set_transient_for(id, parent));
}
pub fn dialog_transient_for(id: usize) -> Option<usize> {
    with_state(|s| s.dialog_transient_for(id))
}
pub fn dialog_mark_destroyed(id: usize) {
    with_state(|s| s.dialog_mark_destroyed(id));
}
/// Deliver a response id to a dialog's `connect_response` handler. Only a
/// response registered with `dialog_add_button` fires.
pub fn dialog_respond(id: usize, response_id: i32) {
    with_state(|s| s.dialog_respond(id, response_id));
}

// -- Spreadsheet --

pub fn sheet_set_cell(id: usize, row: u32, col: u32, text: &str, style: u8) {
    with_state(|s| s.sheet_set_cell(id, row, col, text, style));
}
pub fn sheet_set_raw_cell(id: usize, row: u32, col: u32, text: &str) {
    with_state(|s| s.sheet_set_raw_cell(id, row, col, text));
}
pub fn sheet_get_cell(id: usize, row: u32, col: u32) -> Option<String> {
    with_state(|s| s.sheet_get_cell(id, row, col))
}
pub fn sheet_set_border_title(id: usize, title: &str) {
    with_state(|s| s.sheet_set_border_title(id, title));
}
pub fn sheet_border_title(id: usize) -> Option<String> {
    with_state(|s| s.sheet_border_title(id))
}

// -- Lifecycle --

pub fn quit() {
    with_state(|s| s.quit());
}
pub fn is_running() -> bool {
    with_state(|s| s.running)
}
