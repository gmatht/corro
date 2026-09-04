#[cfg(feature = "pancurses")]
mod pancurses_adapter {
    use std::os::raw::c_void;
    use crate::core::{Error, Widget};

    // -- Window --

    pub struct Window {
        pub(crate) id: usize,
    }

    impl Widget for Window {
        fn raw_handle(&self) -> *mut c_void {
            &self.id as *const usize as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for Window {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    impl Clone for Window {
        fn clone(&self) -> Self { Window { id: self.id } }
    }

    impl Window {
        pub fn set_title(&self, title: &str) {
            crate::backends::pancurses::set_window_title(self.id, title);
        }

        pub fn set_child(&self, child: &impl AsRef<*mut c_void>) {
            let child_ptr = *child.as_ref();
            let child_id = child_ptr as usize;
            crate::backends::pancurses::set_child(self.id, child_id);
        }

        pub fn present(&self) {}

        pub fn insert_action_group(&self, _name: &str, _group_ptr: *mut c_void) {}

        pub fn set_default_size(&self, _width: i32, _height: i32) {}
    }

    // -- Button --

    pub struct Button {
        pub(crate) id: usize,
    }

    impl Widget for Button {
        fn raw_handle(&self) -> *mut c_void {
            &self.id as *const usize as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for Button {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    impl Clone for Button {
        fn clone(&self) -> Self { Button { id: self.id } }
    }

    impl Button {
        pub fn on_click(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            crate::backends::pancurses::add_callback(self.id, Box::new(f));
            Ok(0)
        }

        pub fn emit_clicked(&self) -> Result<u64, Error> {
            // fire synchronously
            Ok(0)
        }
    }

    // -- Label --

    pub struct Label {
        pub(crate) id: usize,
    }

    impl AsRef<*mut c_void> for Label {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    impl Clone for Label {
        fn clone(&self) -> Self { Label { id: self.id } }
    }

    impl Label {
        pub fn set_text(&self, text: &str) {
            crate::backends::pancurses::set_label_text(self.id, text);
        }

        pub fn get_text(&self) -> Option<String> {
            crate::backends::pancurses::get_label_text(self.id)
        }

        pub fn add_class(&self, _class_name: &str) {}
        pub fn remove_class(&self, _class_name: &str) {}
        pub fn set_markup(&self, _markup: &str) {}
        pub fn set_visible(&self, visible: bool) {
            crate::backends::pancurses::set_label_visible(self.id, visible);
        }
        pub fn set_xalign(&self, _x: f32) {}
    }

    // -- BoxWidget --

    pub struct BoxWidget {
        pub(crate) id: usize,
        pub(crate) orientation: Orientation,
        pub(crate) spacing: i32,
    }

    #[derive(Clone, Copy, PartialEq)]
    pub enum Orientation {
        Horizontal,
        Vertical,
    }

    impl AsRef<*mut c_void> for BoxWidget {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    impl BoxWidget {
        pub fn append(&self, child: &impl AsRef<*mut c_void>) {
            let child_ptr = *child.as_ref();
            let child_id = child_ptr as usize;
            crate::backends::pancurses::append_child(self.id, child_id);
        }

        pub fn layout(&self, _x: i32, _y: i32, _w: i32, _h: i32) {
            crate::backends::pancurses::layout_box(self.id);
        }
    }

    // -- Grid --

    pub struct Grid {
        pub(crate) id: usize,
    }

    impl AsRef<*mut c_void> for Grid {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    impl Grid {
        pub fn attach(&self, child: &impl AsRef<*mut c_void>, _left: i32, _top: i32, _width: i32, _height: i32) {
            let child_ptr = *child.as_ref();
            let child_id = child_ptr as usize;
            crate::backends::pancurses::append_child(self.id, child_id);
        }

        pub fn layout(&self) {
            crate::backends::pancurses::layout_grid(self.id);
        }
    }

    // -- Entry --

    pub struct Entry {
        pub(crate) id: usize,
    }

    impl AsRef<*mut c_void> for Entry {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    impl Clone for Entry {
        fn clone(&self) -> Self { Entry { id: self.id } }
    }

    impl Entry {
        pub fn set_text(&self, text: &str) {
            crate::backends::pancurses::entry_set_text(self.id, text);
        }

        pub fn text(&self) -> Option<String> {
            crate::backends::pancurses::entry_text(self.id)
        }

        pub fn get_text(&self) -> Option<String> {
            self.text()
        }

        pub fn connect_changed(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            crate::backends::pancurses::add_callback(self.id, Box::new(f));
            Ok(0)
        }
    }

    // -- Standard dialogs (wxMessageBox / wxFileDialog) --

    pub fn message_box(
        title: &str,
        text: &str,
        kind: crate::MessageBoxKind,
        on_result: Option<Box<dyn FnMut(crate::MessageBoxResult)>>,
    ) {
        crate::backends::pancurses::message_box(title, text, kind, on_result);
    }

    pub fn close_dialog() {
        crate::backends::pancurses::close_dialog();
    }

    pub fn file_open_dialog(on_path: Box<dyn FnMut(Option<String>)>) {
        crate::backends::pancurses::file_open_dialog(on_path);
    }

    pub fn file_save_dialog(default_name: &str, on_path: Box<dyn FnMut(Option<String>)>) {
        crate::backends::pancurses::file_save_dialog(default_name, on_path);
    }

    // -- Menu --

    /// Backend-agnostic menu model: a list of items (actions or submenus).
    /// The same model is used by every rustxwidgets backend, so an application
    /// can build one menu and hand it to any backend's `create_menubar`.
    pub struct Menu {
        pub(crate) id: usize,
        pub(crate) items: std::cell::RefCell<Vec<crate::MenuItem>>,
    }

    impl AsRef<*mut c_void> for Menu {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    impl Menu {
        pub fn append(&self, label: &str, action_name: &str) {
            self.items.borrow_mut().push(crate::MenuItem::Action {
                label: label.to_string(),
                action: action_name.to_string(),
                shortcut: None,
            });
        }
        pub fn append_with_shortcut(&self, label: &str, action_name: &str, shortcut: &str) {
            self.items.borrow_mut().push(crate::MenuItem::Action {
                label: label.to_string(),
                action: action_name.to_string(),
                shortcut: Some(shortcut.to_string()),
            });
        }
        pub fn append_separator(&self) {
            self.items.borrow_mut().push(crate::MenuItem::Separator);
        }
        pub fn append_check(&self, label: &str, action_name: &str, checked: bool) {
            self.items.borrow_mut().push(crate::MenuItem::Check {
                label: label.to_string(),
                action: action_name.to_string(),
                checked,
            });
        }
        pub fn append_radio(&self, label: &str, action_name: &str, group: u32) {
            self.items.borrow_mut().push(crate::MenuItem::Radio {
                label: label.to_string(),
                action: action_name.to_string(),
                group,
            });
        }
        pub fn append_submenu(&self, label: &str, submenu: &Menu) {
            self.append_submenu_with_shortcut(label, "", submenu);
        }
        pub fn append_submenu_with_shortcut(&self, label: &str, shortcut: &str, submenu: &Menu) {
            let sub_items = submenu.items.borrow().clone();
            self.items.borrow_mut().push(crate::MenuItem::Submenu {
                label: label.to_string(),
                items: sub_items,
                shortcut: if shortcut.is_empty() { None } else { Some(shortcut.to_string()) },
            });
        }
        pub fn append_item(&self, _label: &str, _action: &SimpleAction) {}
        pub fn append_section(&self, _label: &str) {}
    }

    /// Extract the menubar model: (root label, items) for each root menu.
    pub(crate) fn collect_menu_items(menu: &Menu) -> Vec<(String, Vec<crate::MenuItem>)> {
        menu.items
            .borrow()
            .iter()
            .filter_map(|item| match item {
                crate::MenuItem::Submenu { label, items, .. } => Some((label.clone(), items.clone())),
                _ => None,
            })
            .collect()
    }

    // -- MenuBar --

    pub struct MenuBar {
        pub(crate) id: usize,
    }

    impl Widget for MenuBar {
        fn raw_handle(&self) -> *mut c_void {
            &self.id as *const usize as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for MenuBar {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    // -- SimpleAction --

    pub struct SimpleAction {
        pub(crate) id: usize,
    }

    impl AsRef<*mut c_void> for SimpleAction {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    impl SimpleAction {
        pub fn on_activate(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            crate::backends::pancurses::add_callback(self.id, Box::new(f));
            Ok(0)
        }
    }

    // -- Dialog --

    pub struct Dialog {
        pub(crate) id: usize,
    }

    impl Widget for Dialog {
        fn raw_handle(&self) -> *mut c_void {
            &self.id as *const usize as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for Dialog {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    impl Dialog {
        pub fn set_title(&self, title: &str) {
            crate::backends::pancurses::set_window_title(self.id, title);
        }
        pub fn set_default_size(&self, _w: i32, _h: i32) {}
        pub fn append_content_area(&self, child: &impl AsRef<*mut c_void>) {
            let child_ptr = *child.as_ref();
            let child_id = child_ptr as usize;
            crate::backends::pancurses::set_child(self.id, child_id);
        }
        pub fn add_button(&self, _label: &str, _response_id: i32) {}
        pub fn connect_response(&self, _f: impl FnMut(i32) + 'static) -> Result<u64, Error> { Ok(0) }
        pub fn present(&self) {}
    }

    // -- DropDown --

    pub struct DropDown {
        pub(crate) id: usize,
    }

    impl Widget for DropDown {
        fn raw_handle(&self) -> *mut c_void {
            &self.id as *const usize as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for DropDown {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    impl Clone for DropDown {
        fn clone(&self) -> Self { DropDown { id: self.id } }
    }

    impl DropDown {
        pub fn set_items(&self, items: &[&str]) {
            crate::backends::pancurses::set_dropdown_items(self.id, items);
        }
        pub fn set_active(&self, idx: i32) {
            crate::backends::pancurses::set_dropdown_selected(self.id, idx);
        }
        pub fn get_active(&self) -> i32 {
            crate::backends::pancurses::get_dropdown_selected(self.id)
        }
        pub fn connect_changed(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            crate::backends::pancurses::add_callback(self.id, Box::new(f));
            Ok(0)
        }
    }

    // -- CheckButton --

    pub struct CheckButton {
        pub(crate) id: usize,
    }

    impl Widget for CheckButton {
        fn raw_handle(&self) -> *mut c_void {
            &self.id as *const usize as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for CheckButton {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    impl Clone for CheckButton {
        fn clone(&self) -> Self { CheckButton { id: self.id } }
    }

    impl CheckButton {
        pub fn set_active(&self, active: bool) {
            crate::backends::pancurses::set_checkbutton_checked(self.id, active);
        }

        pub fn is_active(&self) -> bool {
            crate::backends::pancurses::get_checkbutton_checked(self.id)
        }

        pub fn on_toggle(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            crate::backends::pancurses::add_callback(self.id, Box::new(f));
            Ok(0)
        }

        pub fn connect_toggled(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            self.on_toggle(f)
        }
    }

    // -- RadioButton --

    pub struct RadioButton {
        pub(crate) id: usize,
    }

    impl Widget for RadioButton {
        fn raw_handle(&self) -> *mut c_void {
            &self.id as *const usize as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for RadioButton {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    impl Clone for RadioButton {
        fn clone(&self) -> Self { RadioButton { id: self.id } }
    }

    impl RadioButton {
        pub fn set_active(&self, active: bool) {
            crate::backends::pancurses::set_radiobutton_checked(self.id, active);
        }

        pub fn is_active(&self) -> bool {
            crate::backends::pancurses::get_radiobutton_checked(self.id)
        }

        pub fn on_toggle(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            crate::backends::pancurses::add_callback(self.id, Box::new(f));
            Ok(0)
        }

        pub fn connect_toggled(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            self.on_toggle(f)
        }
    }

    // -- TextView --

    pub struct TextView {
        pub(crate) id: usize,
    }

    impl AsRef<*mut c_void> for TextView {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    impl Widget for TextView {
        fn raw_handle(&self) -> *mut c_void {
            &self.id as *const usize as *mut c_void
        }
    }

    impl TextView {
        pub fn set_text(&self, text: &str) {
            crate::backends::pancurses::set_textview_text(self.id, text);
        }
        pub fn get_text(&self) -> Option<String> {
            crate::backends::pancurses::get_textview_text(self.id)
        }
        pub fn set_wrap_mode(&self, _mode: i32) {}
        pub fn set_size_request(&self, _w: i32, _h: i32) {}
    }

    // -- Spreadsheet --

    pub struct Spreadsheet {
        pub(crate) id: usize,
    }

    impl Clone for Spreadsheet {
        fn clone(&self) -> Self { Spreadsheet { id: self.id } }
    }

    impl AsRef<*mut c_void> for Spreadsheet {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    impl Widget for Spreadsheet {
        fn raw_handle(&self) -> *mut c_void {
            &self.id as *const usize as *mut c_void
        }
    }

    impl Spreadsheet {
        pub fn id(&self) -> usize { self.id }
        pub fn set_cell(&self, row: u32, col: u32, text: &str) {
            crate::backends::pancurses::spreadsheet_set_cell(self.id, row, col, text);
        }
        pub fn get_cell(&self, row: u32, col: u32) -> Option<String> {
            crate::backends::pancurses::spreadsheet_get_cell(self.id, row, col)
        }
        pub fn set_raw_cell(&self, row: u32, col: u32, text: &str) {
            crate::backends::pancurses::spreadsheet_set_raw_cell(self.id, row, col, text);
        }
        pub fn set_cell_style(&self, row: u32, col: u32, style: u8) {
            crate::backends::pancurses::spreadsheet_set_cell_style(self.id, row, col, style);
        }
        pub fn cursor_position(&self) -> Option<(u32, u32)> {
            crate::backends::pancurses::spreadsheet_cursor_position(self.id)
        }
        pub fn set_cursor(&self, row: u32, col: u32) {
            crate::backends::pancurses::spreadsheet_set_cursor(self.id, row, col);
        }
        pub fn set_editing(&self, editing: bool, edit_buf: &str, edit_pos: usize) {
            crate::backends::pancurses::spreadsheet_set_edit_state(self.id, editing, edit_buf, edit_pos);
        }
        pub fn set_grid_config(&self, margin_cols: u32, main_cols: u32) {
            crate::backends::pancurses::spreadsheet_set_grid_config(self.id, margin_cols, main_cols);
        }
        pub fn set_row_counts(&self, header_rows: u32, main_rows: u32) {
            crate::backends::pancurses::spreadsheet_set_row_counts(self.id, header_rows, main_rows);
        }
        pub fn set_column_layout(&self, layout: Vec<(u32, u32, String)>) {
            crate::backends::pancurses::spreadsheet_set_column_layout(self.id, layout);
        }
        pub fn set_row_labels(&self, labels: Vec<(u32, String)>) {
            crate::backends::pancurses::spreadsheet_set_row_labels(self.id, labels);
        }
        pub fn set_menu_text(&self, text: &str) {
            crate::backends::pancurses::spreadsheet_set_menu_text(self.id, text);
        }
        pub fn set_border_title(&self, text: &str) {
            crate::backends::pancurses::spreadsheet_set_border_title(self.id, text);
        }
        pub fn set_status_text(&self, text: &str) {
            crate::backends::pancurses::spreadsheet_set_status_text(self.id, text);
        }
        pub fn set_formula_bar_trailing(&self, text: &str) {
            crate::backends::pancurses::spreadsheet_set_formula_bar_trailing(self.id, text);
        }
        pub fn set_tab_data(&self, titles: &[String], active: usize) {
            crate::backends::pancurses::spreadsheet_set_tab_data(self.id, titles, active);
        }
        pub fn set_formula_bar(&self, address_label: &Label, entry: &Entry) {
            crate::backends::pancurses::spreadsheet_set_formula_bar(
                self.id, address_label.id, entry.id,
            );
        }
        pub fn commit_formula_bar(&self) {
            crate::backends::pancurses::spreadsheet_commit_formula_bar(self.id);
        }
    }

    // -- Factory functions --

    pub fn create_window() -> Result<Window, Error> {
        crate::backends::pancurses::create_window()
            .map(|id| Window { id })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn create_button(label: &str) -> Result<Button, Error> {
        crate::backends::pancurses::create_button(label)
            .map(|id| Button { id })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn create_label(text: &str) -> Result<Label, Error> {
        crate::backends::pancurses::create_label(text)
            .map(|id| Label { id })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn create_box(orientation: Orientation, spacing: i32) -> Result<BoxWidget, Error> {
        let horizontal = orientation == Orientation::Horizontal;
        crate::backends::pancurses::create_box(horizontal, spacing)
            .map(|id| BoxWidget { id, orientation, spacing })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn create_grid() -> Result<Grid, Error> {
        crate::backends::pancurses::create_grid()
            .map(|id| Grid { id })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    // -- DataGrid (a bare grid, mirrors wxGrid; Spreadsheet is built on top) --

    pub struct DataGrid {
        pub(crate) id: usize,
    }

    impl Widget for DataGrid {
        fn raw_handle(&self) -> *mut c_void {
            &self.id as *const usize as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for DataGrid {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    pub fn create_data_grid(rows: u32, cols: u32) -> Result<DataGrid, Error> {
        crate::backends::pancurses::create_data_grid(rows, cols)
            .map(|id| DataGrid { id })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn grid_set_cell(grid: &DataGrid, r: u32, c: u32, text: &str) {
        crate::backends::pancurses::grid_set_cell(grid.id, r, c, text);
    }

    pub fn grid_get_cell(grid: &DataGrid, r: u32, c: u32) -> Option<String> {
        crate::backends::pancurses::grid_get_cell(grid.id, r, c)
    }


    // -- Sizers (wxSizer-like layout) --

    pub struct Sizer {
        pub(crate) id: usize,
    }

    impl AsRef<*mut c_void> for Sizer {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.id as *const usize as *const *mut c_void) }
        }
    }

    pub fn create_box_sizer(horizontal: bool, spacing: i32) -> Result<Sizer, Error> {
        crate::backends::pancurses::create_sizer(crate::Sizer::Box {
            horizontal,
            spacing,
            children: vec![],
        })
        .map(|id| Sizer { id })
        .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn create_grid_sizer(cols: usize, rows: usize) -> Result<Sizer, Error> {
        crate::backends::pancurses::create_sizer(crate::Sizer::Grid {
            cols,
            rows,
            children: vec![],
        })
        .map(|id| Sizer { id })
        .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn create_flex_grid_sizer(cols: usize, rows: usize) -> Result<Sizer, Error> {
        crate::backends::pancurses::create_sizer(crate::Sizer::FlexGrid {
            cols,
            rows,
            children: vec![],
        })
        .map(|id| Sizer { id })
        .map_err(|e| Error::Backend(format!("{}", e)))
    }

    /// Add a child widget to a sizer with weight, border, and flags.
    pub fn sizer_add(
        sizer: &Sizer,
        widget: &impl AsRef<*mut c_void>,
        weight: i32,
        border: i32,
        flags: crate::SizerFlags,
    ) {
        let widget_id = *widget.as_ref() as usize;
        crate::backends::pancurses::sizer_add(sizer.id, widget_id, weight, border, flags);
    }

    pub fn layout_sizer(sizer: &Sizer) {
        crate::backends::pancurses::layout_sizer(sizer.id);
    }

    pub fn set_client_data(widget: &impl AsRef<*mut c_void>, data: &str) {
        let id = *widget.as_ref() as usize;
        crate::backends::pancurses::set_client_data(id, data);
    }

    pub fn get_client_data(widget: &impl AsRef<*mut c_void>) -> Option<String> {
        let id = *widget.as_ref() as usize;
        crate::backends::pancurses::get_client_data(id)
    }

    pub fn set_widget_rect(widget: &impl AsRef<*mut c_void>, x: i32, y: i32, w: i32, h: i32) {
        let id = *widget.as_ref() as usize;
        crate::backends::pancurses::set_widget_rect(id, x, y, w, h);
    }

    pub fn get_widget_rect(widget: &impl AsRef<*mut c_void>) -> Option<(i32, i32, i32, i32)> {
        let id = *widget.as_ref() as usize;
        crate::backends::pancurses::get_widget_rect(id)
    }

    pub fn create_entry() -> Result<Entry, Error> {
        crate::backends::pancurses::create_entry()
            .map(|id| Entry { id })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn create_menu() -> Result<Menu, Error> {
        crate::backends::pancurses::create_menu()
            .map(|id| Menu { id, items: std::cell::RefCell::new(vec![]) })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn create_menubar(model: &Menu, _action_group: *mut c_void) -> Result<MenuBar, Error> {
        let submenu_items = collect_menu_items(model);
        let id = unsafe { crate::backends::pancurses::create_menubar(submenu_items, _action_group) }
            .map_err(|e| Error::Backend(format!("{}", e)))?;
        Ok(MenuBar { id })
    }

    pub fn create_simple_action(name: &str) -> Result<SimpleAction, Error> {
        crate::backends::pancurses::create_simple_action(name)
            .map(|id| SimpleAction { id })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn create_dialog() -> Result<Dialog, Error> {
        crate::backends::pancurses::create_dialog()
            .map(|id| Dialog { id })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn create_dropdown(items: &[&str]) -> Result<DropDown, Error> {
        crate::backends::pancurses::create_dropdown(items)
            .map(|id| DropDown { id })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn create_checkbutton(label: &str) -> Result<CheckButton, Error> {
        crate::backends::pancurses::create_checkbutton(label)
            .map(|id| CheckButton { id })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn create_radiobutton(group: Option<&RadioButton>, label: &str) -> Result<RadioButton, Error> {
        let gid = group.map(|r| r.id);
        crate::backends::pancurses::create_radiobutton(gid, label)
            .map(|id| RadioButton { id })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn create_textview() -> Result<TextView, Error> {
        crate::backends::pancurses::create_textview()
            .map(|id| TextView { id })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn create_spreadsheet(rows: u32, cols: u32) -> Result<Spreadsheet, Error> {
        crate::backends::pancurses::create_spreadsheet(rows, cols)
            .map(|id| Spreadsheet { id })
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    pub fn add_cursor_move_callback<F: FnMut(u32, u32) + 'static>(f: F) {
        crate::backends::pancurses::spreadsheet_add_cursor_move_callback(f);
    }
    pub fn add_commit_edit_callback<F: FnMut(u32, u32, String) + 'static>(f: F) {
        crate::backends::pancurses::spreadsheet_add_commit_edit_callback(f);
    }
    pub fn add_goto_callback<F: FnMut() + 'static>(f: F) {
        crate::backends::pancurses::spreadsheet_add_goto_callback(f);
    }
    pub fn spreadsheet_set_cell(id: usize, r: u32, c: u32, text: &str) {
        crate::backends::pancurses::spreadsheet_set_cell(id, r, c, text);
    }
    pub fn spreadsheet_set_cell_style(id: usize, r: u32, c: u32, style: u8) {
        crate::backends::pancurses::spreadsheet_set_cell_style(id, r, c, style);
    }
    pub fn spreadsheet_set_column_layout(id: usize, layout: Vec<(u32, u32, String)>) {
        crate::backends::pancurses::spreadsheet_set_column_layout(id, layout);
    }
    pub fn spreadsheet_set_border_title(id: usize, text: &str) {
        crate::backends::pancurses::spreadsheet_set_border_title(id, text);
    }
    pub fn spreadsheet_set_row_labels(id: usize, labels: Vec<(u32, String)>) {
        crate::backends::pancurses::spreadsheet_set_row_labels(id, labels);
    }
    pub fn spreadsheet_set_grid_config(id: usize, margin_cols: u32, main_cols: u32) {
        crate::backends::pancurses::spreadsheet_set_grid_config(id, margin_cols, main_cols);
    }
    pub fn spreadsheet_set_edit_state(id: usize, editing: bool, edit_buf: &str, edit_pos: usize) {
        crate::backends::pancurses::spreadsheet_set_edit_state(id, editing, edit_buf, edit_pos);
    }
}

pub use pancurses_adapter::*;

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn menu_model_supports_item_kinds() {
        let menu = create_menu().unwrap();
        menu.append("Open", "open");
        menu.append_with_shortcut("Save", "save", "Ctrl+S");
        menu.append_separator();
        menu.append_check("Bold", "bold", true);
        menu.append_radio("Left", "align_left", 0);
        let sub = create_menu().unwrap();
        sub.append("TSV", "export_tsv");
        menu.append_submenu("Export", &sub);

        let items = menu.items.borrow();
        assert_eq!(items.len(), 6);
        assert!(matches!(&items[0],
            crate::MenuItem::Action { label, action, shortcut }
            if label == "Open" && action == "open" && shortcut.is_none()));
        assert!(matches!(&items[1],
            crate::MenuItem::Action { label, action, shortcut }
            if label == "Save" && action == "save" && shortcut.as_deref() == Some("Ctrl+S")));
        assert!(matches!(&items[2], crate::MenuItem::Separator));
        assert!(matches!(&items[3],
            crate::MenuItem::Check { label, action, checked }
            if label == "Bold" && action == "bold" && *checked));
        assert!(matches!(&items[4],
            crate::MenuItem::Radio { label, action, group }
            if label == "Left" && action == "align_left" && *group == 0));
        assert!(matches!(&items[5],
            crate::MenuItem::Submenu { label, items }
            if label == "Export" && items.len() == 1));
    }

    #[test]
    fn action_registry_tracks_state() {
        crate::backends::pancurses::register_action("bold", true, false);
        crate::backends::pancurses::set_action_checked("bold", true);
        assert_eq!(crate::backends::pancurses::action_state("bold"), Some((true, true)));
        crate::backends::pancurses::set_action_enabled("bold", false);
        assert_eq!(crate::backends::pancurses::action_state("bold"), Some((false, true)));
        assert_eq!(crate::backends::pancurses::action_state("missing"), None);
    }

    #[test]
    fn event_propagation_bubbles_to_parent() {
        let win = create_window().unwrap();
        let boxw = create_box(Orientation::Vertical, 0).unwrap();
        let btn = create_button("Click").unwrap();
        win.set_child(&boxw);
        boxw.append(&btn);

        // Child skips; parent handles.
        let parent_fired = std::rc::Rc::new(std::cell::Cell::new(false));
        let pf = parent_fired.clone();
        crate::backends::pancurses::add_event_callback(boxw.id, Box::new(move |_ev| {
            pf.set(true);
            crate::CallbackResult::Handled
        }));
        crate::backends::pancurses::add_event_callback(btn.id, Box::new(|_ev| crate::CallbackResult::Skip));

        let result = crate::backends::pancurses::fire_event(btn.id, &crate::Event::Activate);
        assert_eq!(result, crate::CallbackResult::Handled);
        assert!(parent_fired.get(), "parent should have received the bubbled event");
    }

    #[test]
    fn event_propagation_stops_when_handled() {
        let win = create_window().unwrap();
        let boxw = create_box(Orientation::Vertical, 0).unwrap();
        let btn = create_button("Click").unwrap();
        win.set_child(&boxw);
        boxw.append(&btn);

        // Child handles; parent must NOT fire.
        let parent_fired = std::rc::Rc::new(std::cell::Cell::new(false));
        let pf = parent_fired.clone();
        crate::backends::pancurses::add_event_callback(boxw.id, Box::new(move |_ev| {
            pf.set(true);
            crate::CallbackResult::Handled
        }));
        crate::backends::pancurses::add_event_callback(btn.id, Box::new(|_ev| crate::CallbackResult::Handled));

        let result = crate::backends::pancurses::fire_event(btn.id, &crate::Event::Activate);
        assert_eq!(result, crate::CallbackResult::Handled);
        assert!(!parent_fired.get(), "parent should NOT fire when child handled the event");
    }

    #[test]
    fn box_sizer_lays_out_children_by_weight() {
        let win = create_window().unwrap();
        let sizer = create_box_sizer(true, 0).unwrap();
        let b1 = create_button("A").unwrap();
        let b2 = create_button("B").unwrap();
        win.set_child(&sizer);
        sizer_add(&sizer, &b1, 1, 0, crate::SizerFlags::default());
        sizer_add(&sizer, &b2, 2, 0, crate::SizerFlags::default());
        set_widget_rect(&sizer, 0, 0, 30, 1);
        layout_sizer(&sizer);

        let (_, _, w1, _) = get_widget_rect(&b1).unwrap();
        let (_, _, w2, _) = get_widget_rect(&b2).unwrap();
        // Weights 1:2 over 30 columns -> 10 and 20.
        assert_eq!(w1, 10, "weight-1 child should get 1/3 of the width");
        assert_eq!(w2, 20, "weight-2 child should get 2/3 of the width");
    }

    #[test]
    fn grid_sizer_arranges_children_in_cells() {
        let win = create_window().unwrap();
        let sizer = create_grid_sizer(2, 2).unwrap();
        let b1 = create_button("A").unwrap();
        let b2 = create_button("B").unwrap();
        let b3 = create_button("C").unwrap();
        let b4 = create_button("D").unwrap();
        win.set_child(&sizer);
        for b in [&b1, &b2, &b3, &b4] {
            sizer_add(&sizer, b, 0, 0, crate::SizerFlags::default());
        }
        set_widget_rect(&sizer, 0, 0, 20, 2);
        layout_sizer(&sizer);

        let (x1, y1, w1, h1) = get_widget_rect(&b1).unwrap();
        let (x2, y2, _, _) = get_widget_rect(&b2).unwrap();
        let (x3, y3, _, _) = get_widget_rect(&b3).unwrap();
        assert_eq!((x1, y1, w1, h1), (0, 0, 10, 1), "cell (0,0)");
        assert_eq!((x2, y2), (10, 0), "cell (1,0)");
        assert_eq!((x3, y3), (0, 1), "cell (0,1)");
    }

    #[test]
    fn message_box_fires_result_on_close() {
        let result = std::rc::Rc::new(std::cell::Cell::new(None));
        let r = result.clone();
        message_box("Test", "Hello world", crate::MessageBoxKind::Info, Some(Box::new(move |res| {
            r.set(Some(res));
        })));
        close_dialog();
        assert_eq!(result.get(), Some(crate::MessageBoxResult::Ok));
    }

    #[test]
    fn client_data_round_trip() {
        let win = create_window().unwrap();
        let btn = create_button("Click").unwrap();
        win.set_child(&btn);
        set_client_data(&btn, "my-data-42");
        assert_eq!(get_client_data(&btn).as_deref(), Some("my-data-42"));
        set_client_data(&btn, "updated");
        assert_eq!(get_client_data(&btn).as_deref(), Some("updated"));
    }

    #[test]
    fn data_grid_holds_cells_independently() {
        let win = create_window().unwrap();
        let grid = create_data_grid(10, 3).unwrap();
        win.set_child(&grid);
        grid_set_cell(&grid, 1, 2, "hello");
        assert_eq!(grid_get_cell(&grid, 1, 2).as_deref(), Some("hello"));
        assert_eq!(grid_get_cell(&grid, 0, 0), None);
        // The DataGrid is a bare grid: it does NOT have spreadsheet chrome, so
        // the spreadsheet setters must not apply to it.
        grid_set_cell(&grid, 5, 5, "out");
        assert_eq!(grid_get_cell(&grid, 5, 5).as_deref(), Some("out"));
    }

    #[test]
    fn menubar_collects_root_submenus() {
        let menubar = create_menu().unwrap();
        for (label, items) in [("File", &["open", "save"][..]), ("Edit", &["cut", "copy"][..])] {
            let sub = create_menu().unwrap();
            for a in items {
                sub.append(a, a);
            }
            menubar.append_submenu(label, &sub);
        }
        let roots = collect_menu_items(&menubar);
        assert_eq!(roots.len(), 2);
        assert_eq!(roots[0].0, "File");
        assert_eq!(roots[0].1.len(), 2);
        assert_eq!(roots[1].0, "Edit");
    }
}
