use std::cell::RefCell;
#[cfg(windows)]
use std::collections::HashMap;
use std::os::raw::c_void;
use std::rc::Rc;

/// Opaque handler id returned when connecting signals
pub type HandlerId = u64;

/// Cross-platform 2D drawing surface.
/// Each backend implements this trait with its own drawing primitives.
pub trait DrawContext {
    fn fill_rect(&mut self, x: f64, y: f64, w: f64, h: f64, r: f64, g: f64, b: f64, a: f64);
    fn stroke_rect(&mut self, x: f64, y: f64, w: f64, h: f64, r: f64, g: f64, b: f64, a: f64, lw: f64);
    /// Draw text with normal (non-bold, non-italic) style.
    fn draw_text(&mut self, x: f64, y: f64, text: &str, font: &str, size: f64, r: f64, g: f64, b: f64, a: f64) {
        self.draw_text_styled(x, y, text, font, size, r, g, b, a, 0, 0)
    }
    /// Draw text with explicit Cairo slant (0=normal, 1=italic, 2=oblique) and
    /// weight (0=normal, 1=bold).
    fn draw_text_styled(&mut self, x: f64, y: f64, text: &str, font: &str, size: f64,
                        r: f64, g: f64, b: f64, a: f64, slant: i32, weight: i32);
    /// Measure text extents with normal (non-bold) weight.
    fn text_extents(&self, text: &str, font: &str, size: f64) -> (f64, f64, f64, f64) {
        self.text_extents_styled(text, font, size, 0, 0)
    }
    /// Measure text extents with explicit Cairo slant and weight.
    fn text_extents_styled(&self, text: &str, font: &str, size: f64, slant: i32, weight: i32) -> (f64, f64, f64, f64);
    fn clear(&mut self, r: f64, g: f64, b: f64, a: f64);
    fn save(&mut self);
    fn restore(&mut self);
    fn clip(&mut self, x: f64, y: f64, w: f64, h: f64);
}

/// Top-level error type
#[derive(thiserror::Error, Debug)]
pub enum Error {
    #[error("backend error: {0}")]
    Backend(String),
}

/// Widget trait: minimal escape hatch for raw handles
pub trait Widget {
    /// Return an opaque raw pointer for backend interop. Use unsafe to deref.
    fn raw_handle(&self) -> *mut c_void;
}

/// Core App wrapper that holds a boxed backend application.
#[derive(Clone)]
pub struct App {
    inner: Rc<RefCell<Option<Box<dyn crate::backends::BackendApp>>>>,
    #[cfg(windows)]
    parent_cell: Rc<RefCell<Option<*mut c_void>>>,
    #[cfg(windows)]
    action_registry: Rc<RefCell<HashMap<String, Box<dyn FnMut()>>>>,
}

impl App {
    /// Initialize the default backend and return an App wrapper.
    /// Uses the priority chain from `backends::init()` (gtk > nwg > wasm > android > pancurses).
    pub fn init() -> Result<Self, Error> {
        let b = match crate::backends::init() {
            Ok(b) => b,
            Err(e) => return Err(Error::Backend(format!("{}", e))),
        };
        #[cfg(not(feature = "pancurses"))]
        {
            #[cfg(windows)]
            {
                // Create a hidden parent window for child controls
                let parent_hwnd = crate::backends::nwg::create_hidden_parent()?;
                return Ok(App {
                    inner: Rc::new(RefCell::new(Some(b))),
                    parent_cell: Rc::new(RefCell::new(Some(parent_hwnd))),
                    action_registry: Rc::new(RefCell::new(HashMap::new())),
                });
            }
            #[cfg(not(windows))]
            return Ok(App {
                inner: Rc::new(RefCell::new(Some(b))),
            });
        }
        #[cfg(feature = "pancurses")]
        return Ok(App { inner: Rc::new(RefCell::new(Some(b))) });
    }

    // -- Linux paths --

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_window(&self) -> Result<crate::backends_gtk_adapter::Window, Error> {
        crate::backends_gtk_adapter::create_window().map_err(|e| e)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_button(&self, label: &str) -> Result<crate::backends_gtk_adapter::Button, Error> {
        crate::backends_gtk_adapter::create_button(label).map_err(|e| e)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_label(&self, text: &str) -> Result<crate::backends_gtk_adapter::Label, Error> {
        crate::backends_gtk_adapter::create_label(text).map_err(|e| e)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_box(&self, orientation: gtk_dynamic_loader::Orientation, spacing: i32) -> Result<crate::backends_gtk_adapter::BoxWidget, Error> {
        crate::backends_gtk_adapter::create_box(orientation, spacing).map_err(|e| e)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_grid(&self) -> Result<crate::backends_gtk_adapter::Grid, Error> {
        crate::backends_gtk_adapter::create_grid().map_err(|e| e)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_entry(&self) -> Result<crate::backends_gtk_adapter::Entry, Error> {
        crate::backends_gtk_adapter::create_entry().map_err(|e| e)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_menu(&self) -> Result<crate::backends_gtk_adapter::Menu, Error> {
        crate::backends_gtk_adapter::create_menu().map_err(|e| e)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    /// # Safety
    /// `action_group` must be a valid GActionGroup pointer or null.
    pub unsafe fn create_menubar(&self, model: &crate::backends_gtk_adapter::Menu, action_group: *mut c_void) -> Result<crate::backends_gtk_adapter::MenuBar, Error> {
        crate::backends_gtk_adapter::create_menubar(model, action_group).map_err(|e| e)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_simple_action(&self, name: &str) -> Result<crate::backends_gtk_adapter::SimpleAction, Error> {
        crate::backends_gtk_adapter::create_simple_action(name).map_err(|e| e)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_dialog(&self) -> Result<crate::backends_gtk_adapter::Dialog, Error> {
        crate::backends_gtk_adapter::create_dialog().map_err(|e| e)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_dropdown(&self, items: &[&str]) -> Result<crate::backends_gtk_adapter::DropDown, Error> {
        crate::backends_gtk_adapter::create_dropdown(items).map_err(|e| e)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_checkbutton(&self, label: &str) -> Result<crate::backends_gtk_adapter::CheckButton, Error> {
        crate::backends_gtk_adapter::create_checkbutton(label).map_err(|e| e)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_radiobutton(&self, label: &str) -> Result<crate::backends_gtk_adapter::RadioButton, Error> {
        crate::backends_gtk_adapter::create_radiobutton(None, label).map_err(|e| e)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_textview(&self) -> Result<crate::backends_gtk_adapter::TextView, Error> {
        crate::backends_gtk_adapter::create_textview().map_err(|e| e)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_canvas(&self) -> Result<crate::backends_gtk_adapter::Canvas, Error> {
        crate::backends_gtk_adapter::create_canvas()
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_scrolled_window(&self) -> Result<crate::backends_gtk_adapter::ScrolledWindow, Error> {
        crate::backends_gtk_adapter::create_scrolled_window()
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_overlay(&self) -> Result<crate::backends_gtk_adapter::Overlay, Error> {
        crate::backends_gtk_adapter::create_overlay()
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn open_file(&self, title: &str) -> Result<Option<String>, Error> {
        crate::backends_gtk_adapter::open_file(title)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn save_file(&self, title: &str) -> Result<Option<String>, Error> {
        crate::backends_gtk_adapter::save_file(title)
    }

    #[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
    pub fn create_spreadsheet(&self, rows: usize, cols: usize) -> Result<crate::backends_gtk_adapter::Spreadsheet, Error> {
        crate::backends_gtk_adapter::create_spreadsheet(rows, cols)
    }

    // -- Windows paths --

    #[cfg(all(windows, not(feature = "zork")))]
    pub fn create_window(&self) -> Result<crate::backends_nwg_adapter::Window, Error> {
        crate::backends_nwg_adapter::create_window(&self.parent_cell)
    }

    #[cfg(all(windows, not(feature = "pancurses"), not(feature = "zork")))]
    pub fn create_button(&self, label: &str) -> Result<crate::backends_nwg_adapter::Button, Error> {
        let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
        crate::backends_nwg_adapter::create_button(parent, label)
    }

    #[cfg(all(windows, not(feature = "zork")))]
    pub fn create_label(&self, text: &str) -> Result<crate::backends_nwg_adapter::Label, Error> {
        let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
        let lbl = crate::backends_nwg_adapter::create_label(parent)?;
        lbl.set_text(text);
        Ok(lbl)
    }

    #[cfg(all(windows, not(feature = "pancurses"), not(feature = "zork")))]
    pub fn create_box(&self, orientation: crate::backends::nwg::Orientation, spacing: i32) -> Result<crate::backends_nwg_adapter::BoxWidget, Error> {
        let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
        crate::backends_nwg_adapter::create_box(orientation, spacing, parent)
    }

    #[cfg(all(windows, not(feature = "zork")))]
    pub fn create_grid(&self) -> Result<crate::backends_nwg_adapter::Grid, Error> {
        crate::backends_nwg_adapter::create_grid()
    }

    #[cfg(all(windows, not(feature = "zork")))]
    pub fn create_entry(&self) -> Result<crate::backends_nwg_adapter::Entry, Error> {
        let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
        crate::backends_nwg_adapter::create_entry(parent)
    }

    #[cfg(all(windows, not(feature = "zork")))]
    pub fn create_menu(&self) -> Result<crate::backends_nwg_adapter::Menu, Error> {
        crate::backends_nwg_adapter::create_menu()
    }

    #[cfg(all(windows, not(feature = "pancurses"), not(feature = "zork")))]
    /// # Safety
    /// `window_hwnd` must be a valid HWND.
    pub unsafe fn create_menubar(&self, model: &crate::backends_nwg_adapter::Menu, window_hwnd: *mut c_void) -> Result<crate::backends_nwg_adapter::MenuBar, Error> {
        crate::backends_nwg_adapter::create_menubar(model, window_hwnd, self.action_registry.clone())
    }

    #[cfg(all(windows, not(feature = "zork")))]
    pub fn create_simple_action(&self, name: &str) -> Result<crate::backends_nwg_adapter::SimpleAction, Error> {
        crate::backends_nwg_adapter::create_simple_action(name, self.action_registry.clone())
    }

    #[cfg(all(windows, not(feature = "pancurses"), not(feature = "zork")))]
    pub fn create_dialog(&self) -> Result<crate::backends_nwg_adapter::Dialog, Error> {
        crate::backends_nwg_adapter::create_dialog(&self.parent_cell)
    }

    #[cfg(all(windows, not(feature = "pancurses"), not(feature = "zork")))]
    pub fn create_dropdown(&self, items: &[&str]) -> Result<crate::backends_nwg_adapter::DropDown, Error> {
        let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
        crate::backends_nwg_adapter::create_dropdown(parent, items)
    }

    #[cfg(all(windows, not(feature = "pancurses"), not(feature = "zork")))]
    pub fn create_checkbutton(&self, label: &str) -> Result<crate::backends_nwg_adapter::CheckButton, Error> {
        let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
        let cb = crate::backends_nwg_adapter::create_checkbutton(parent)?;
        cb.set_label(label);
        Ok(cb)
    }

    #[cfg(all(windows, not(feature = "pancurses"), not(feature = "zork")))]
    pub fn create_radiobutton(&self, label: &str) -> Result<crate::backends_nwg_adapter::RadioButton, Error> {
        let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
        let rb = crate::backends_nwg_adapter::create_radiobutton(parent)?;
        rb.set_label(label);
        Ok(rb)
    }

    #[cfg(all(windows, not(feature = "pancurses"), not(feature = "zork")))]
    pub fn create_textview(&self) -> Result<crate::backends_nwg_adapter::TextView, Error> {
        let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
        crate::backends_nwg_adapter::create_textview(parent)
    }

    #[cfg(all(windows, not(feature = "pancurses"), not(feature = "zork")))]
    pub fn create_canvas(&self) -> Result<crate::backends_nwg_adapter::Canvas, Error> {
        let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
        crate::backends_nwg_adapter::create_canvas(parent)
    }

    #[cfg(all(windows, not(feature = "pancurses"), not(feature = "zork")))]
    pub fn create_overlay(&self) -> Result<crate::backends_nwg_adapter::Overlay, Error> {
        let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
        crate::backends_nwg_adapter::create_overlay(parent)
    }

    #[cfg(all(windows, not(feature = "pancurses"), not(feature = "zork")))]
    pub fn create_scrolled_window(&self) -> Result<crate::backends_nwg_adapter::ScrolledWindow, Error> {
        let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
        crate::backends_nwg_adapter::create_scrolled_window(parent)
    }

    #[cfg(all(windows, not(feature = "pancurses"), not(feature = "zork")))]
    pub fn open_file(&self, title: &str) -> Result<Option<String>, Error> {
        let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
        crate::backends_nwg_adapter::open_file(title, parent)
    }

    #[cfg(all(windows, not(feature = "pancurses"), not(feature = "zork")))]
    pub fn save_file(&self, title: &str) -> Result<Option<String>, Error> {
        let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
        crate::backends_nwg_adapter::save_file(title, parent)
    }

    // -- Pancurses paths --

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_window(&self) -> Result<crate::backends_pancurses_adapter::Window, Error> {
        crate::backends_pancurses_adapter::create_window()
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_button(&self, label: &str) -> Result<crate::backends_pancurses_adapter::Button, Error> {
        crate::backends_pancurses_adapter::create_button(label)
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_label(&self, text: &str) -> Result<crate::backends_pancurses_adapter::Label, Error> {
        crate::backends_pancurses_adapter::create_label(text)
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_box(&self, orientation: crate::backends_pancurses_adapter::Orientation, spacing: i32) -> Result<crate::backends_pancurses_adapter::BoxWidget, Error> {
        crate::backends_pancurses_adapter::create_box(orientation, spacing)
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_grid(&self) -> Result<crate::backends_pancurses_adapter::Grid, Error> {
        crate::backends_pancurses_adapter::create_grid()
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_entry(&self) -> Result<crate::backends_pancurses_adapter::Entry, Error> {
        crate::backends_pancurses_adapter::create_entry()
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_menu(&self) -> Result<crate::backends_pancurses_adapter::Menu, Error> {
        crate::backends_pancurses_adapter::create_menu()
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_menubar(&self, model: &crate::backends_pancurses_adapter::Menu, _action_group: *mut std::os::raw::c_void) -> Result<crate::backends_pancurses_adapter::MenuBar, Error> {
        crate::backends_pancurses_adapter::create_menubar(model, _action_group)
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_simple_action(&self, name: &str) -> Result<crate::backends_pancurses_adapter::SimpleAction, Error> {
        crate::backends_pancurses_adapter::create_simple_action(name)
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_dialog(&self) -> Result<crate::backends_pancurses_adapter::Dialog, Error> {
        crate::backends_pancurses_adapter::create_dialog()
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_dropdown(&self, items: &[&str]) -> Result<crate::backends_pancurses_adapter::DropDown, Error> {
        crate::backends_pancurses_adapter::create_dropdown(items)
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_checkbutton(&self, label: &str) -> Result<crate::backends_pancurses_adapter::CheckButton, Error> {
        crate::backends_pancurses_adapter::create_checkbutton(label)
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_radiobutton(&self, label: &str) -> Result<crate::backends_pancurses_adapter::RadioButton, Error> {
        crate::backends_pancurses_adapter::create_radiobutton(None, label)
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_textview(&self) -> Result<crate::backends_pancurses_adapter::TextView, Error> {
        crate::backends_pancurses_adapter::create_textview()
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_spreadsheet(&self, rows: u32, cols: u32) -> Result<crate::backends_pancurses_adapter::Spreadsheet, Error> {
        crate::backends_pancurses_adapter::create_spreadsheet(rows, cols)
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_canvas(&self) -> Result<crate::backends_pancurses_adapter::Canvas, Error> {
        crate::backends_pancurses_adapter::create_canvas()
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_overlay(&self) -> Result<crate::backends_pancurses_adapter::Overlay, Error> {
        crate::backends_pancurses_adapter::create_overlay()
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn create_scrolled_window(&self) -> Result<crate::backends_pancurses_adapter::ScrolledWindow, Error> {
        crate::backends_pancurses_adapter::create_scrolled_window()
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn open_file(&self, title: &str) -> Result<Option<String>, Error> {
        crate::backends_pancurses_adapter::open_file(title)
    }

    #[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
    pub fn save_file(&self, title: &str) -> Result<Option<String>, Error> {
        crate::backends_pancurses_adapter::save_file(title)
    }

    // -- Zork paths --

    #[cfg(feature = "zork")]
    pub fn create_canvas(&self) -> Result<crate::backends_zork_adapter::Canvas, Error> {
        crate::backends_zork_adapter::create_canvas()
    }

    #[cfg(feature = "zork")]
    pub fn create_overlay(&self) -> Result<crate::backends_zork_adapter::Overlay, Error> {
        crate::backends_zork_adapter::create_overlay()
    }

    #[cfg(feature = "zork")]
    pub fn create_scrolled_window(&self) -> Result<crate::backends_zork_adapter::ScrolledWindow, Error> {
        crate::backends_zork_adapter::create_scrolled_window()
    }

    #[cfg(feature = "zork")]
    pub fn open_file(&self, title: &str) -> Result<Option<String>, Error> {
        crate::backends_zork_adapter::open_file(title)
    }

    #[cfg(feature = "zork")]
    pub fn save_file(&self, title: &str) -> Result<Option<String>, Error> {
        crate::backends_zork_adapter::save_file(title)
    }

    #[cfg(feature = "zork")]
    pub fn create_window(&self) -> Result<crate::backends_zork_adapter::Window, Error> {
        crate::backends_zork_adapter::create_window()
    }

    #[cfg(feature = "zork")]
    pub fn create_button(&self, label: &str) -> Result<crate::backends_zork_adapter::Button, Error> {
        crate::backends_zork_adapter::create_button(label)
    }

    #[cfg(feature = "zork")]
    pub fn create_label(&self, text: &str) -> Result<crate::backends_zork_adapter::Label, Error> {
        crate::backends_zork_adapter::create_label(text)
    }

    #[cfg(feature = "zork")]
    pub fn create_box(&self, orientation: crate::backends_zork_adapter::Orientation, spacing: i32) -> Result<crate::backends_zork_adapter::BoxWidget, Error> {
        crate::backends_zork_adapter::create_box(orientation, spacing)
    }

    #[cfg(feature = "zork")]
    pub fn create_grid(&self) -> Result<crate::backends_zork_adapter::Grid, Error> {
        crate::backends_zork_adapter::create_grid()
    }

    #[cfg(feature = "zork")]
    pub fn create_entry(&self) -> Result<crate::backends_zork_adapter::Entry, Error> {
        crate::backends_zork_adapter::create_entry()
    }

    #[cfg(feature = "zork")]
    pub fn create_menu(&self) -> Result<crate::backends_zork_adapter::Menu, Error> {
        crate::backends_zork_adapter::create_menu()
    }

    #[cfg(feature = "zork")]
    pub fn create_menubar(&self, model: &crate::backends_zork_adapter::Menu, _action_group: *mut std::os::raw::c_void) -> Result<crate::backends_zork_adapter::MenuBar, Error> {
        crate::backends_zork_adapter::create_menubar(model, _action_group)
    }

    #[cfg(feature = "zork")]
    pub fn create_simple_action(&self, name: &str) -> Result<crate::backends_zork_adapter::SimpleAction, Error> {
        crate::backends_zork_adapter::create_simple_action(name)
    }

    #[cfg(feature = "zork")]
    pub fn create_dialog(&self) -> Result<crate::backends_zork_adapter::Dialog, Error> {
        crate::backends_zork_adapter::create_dialog()
    }

    #[cfg(feature = "zork")]
    pub fn create_dropdown(&self, items: &[&str]) -> Result<crate::backends_zork_adapter::DropDown, Error> {
        crate::backends_zork_adapter::create_dropdown(items)
    }

    #[cfg(feature = "zork")]
    pub fn create_checkbutton(&self, label: &str) -> Result<crate::backends_zork_adapter::CheckButton, Error> {
        crate::backends_zork_adapter::create_checkbutton(label)
    }

    #[cfg(feature = "zork")]
    pub fn create_radiobutton(&self, label: &str) -> Result<crate::backends_zork_adapter::RadioButton, Error> {
        crate::backends_zork_adapter::create_radiobutton(None, label)
    }

    #[cfg(feature = "zork")]
    pub fn create_textview(&self) -> Result<crate::backends_zork_adapter::TextView, Error> {
        crate::backends_zork_adapter::create_textview()
    }

    // -- WASM paths --

    #[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
    pub fn create_window(&self) -> Result<crate::backends_wasm_adapter::Window, Error> {
        crate::backends_wasm_adapter::create_window()
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
    pub fn create_button(&self, label: &str) -> Result<crate::backends_wasm_adapter::Button, Error> {
        crate::backends_wasm_adapter::create_button(label)
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
    pub fn create_label(&self, text: &str) -> Result<crate::backends_wasm_adapter::Label, Error> {
        crate::backends_wasm_adapter::create_label(text)
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
    pub fn create_box(&self, orientation: crate::backends_wasm_adapter::Orientation, spacing: i32) -> Result<crate::backends_wasm_adapter::BoxWidget, Error> {
        crate::backends_wasm_adapter::create_box(orientation, spacing)
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
    pub fn create_grid(&self) -> Result<crate::backends_wasm_adapter::Grid, Error> {
        crate::backends_wasm_adapter::create_grid()
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
    pub fn create_entry(&self) -> Result<crate::backends_wasm_adapter::Entry, Error> {
        crate::backends_wasm_adapter::create_entry()
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
    pub fn create_menu(&self) -> Result<crate::backends_wasm_adapter::Menu, Error> {
        crate::backends_wasm_adapter::create_menu()
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
    pub fn create_menubar(&self, model: &crate::backends_wasm_adapter::Menu, action_group: *mut c_void) -> Result<crate::backends_wasm_adapter::MenuBar, Error> {
        crate::backends_wasm_adapter::create_menubar(model, action_group)
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
    pub fn create_simple_action(&self, name: &str) -> Result<crate::backends_wasm_adapter::SimpleAction, Error> {
        crate::backends_wasm_adapter::create_simple_action(name)
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
    pub fn create_dialog(&self) -> Result<crate::backends_wasm_adapter::Dialog, Error> {
        crate::backends_wasm_adapter::create_dialog()
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
    pub fn create_dropdown(&self, items: &[&str]) -> Result<crate::backends_wasm_adapter::DropDown, Error> {
        crate::backends_wasm_adapter::create_dropdown(items)
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
    pub fn create_checkbutton(&self, label: &str) -> Result<crate::backends_wasm_adapter::CheckButton, Error> {
        crate::backends_wasm_adapter::create_checkbutton(label)
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
    pub fn create_radiobutton(&self, label: &str) -> Result<crate::backends_wasm_adapter::RadioButton, Error> {
        crate::backends_wasm_adapter::create_radiobutton(None, label)
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "pancurses"), not(feature = "zork")))]
    pub fn create_textview(&self) -> Result<crate::backends_wasm_adapter::TextView, Error> {
        crate::backends_wasm_adapter::create_textview()
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "pancurses"), not(feature = "zork")))]
    pub fn create_canvas(&self) -> Result<crate::backends_wasm_adapter::Canvas, Error> {
        crate::backends_wasm_adapter::create_canvas()
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "pancurses"), not(feature = "zork")))]
    pub fn create_overlay(&self) -> Result<crate::backends_wasm_adapter::Overlay, Error> {
        crate::backends_wasm_adapter::create_overlay()
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "pancurses"), not(feature = "zork")))]
    pub fn create_scrolled_window(&self) -> Result<crate::backends_wasm_adapter::ScrolledWindow, Error> {
        crate::backends_wasm_adapter::create_scrolled_window()
    }
    #[cfg(all(target_arch = "wasm32", not(feature = "pancurses"), not(feature = "zork")))]
    pub fn open_file(&self, _title: &str) -> Result<Option<String>, Error> {
        Ok(None) // File dialogs not available in WASM
    }

    #[cfg(all(target_arch = "wasm32", not(feature = "pancurses"), not(feature = "zork")))]
    pub fn save_file(&self, _title: &str) -> Result<Option<String>, Error> {
        Ok(None) // File dialogs not available in WASM
    }

// -- Android paths --

#[cfg(all(target_os = "android", not(feature = "zork")))]
pub fn create_window(&self) -> Result<crate::backends_android_adapter::Window, Error> {
    crate::backends_android_adapter::create_window()
}

#[cfg(all(target_os = "android", not(feature = "zork")))]
pub fn create_button(&self, label: &str) -> Result<crate::backends_android_adapter::Button, Error> {
    crate::backends_android_adapter::create_button(label)
}

#[cfg(all(target_os = "android", not(feature = "zork")))]
pub fn create_label(&self, text: &str) -> Result<crate::backends_android_adapter::Label, Error> {
    crate::backends_android_adapter::create_label(text)
}

#[cfg(all(target_os = "android", not(feature = "zork")))]
pub fn create_grid(&self) -> Result<crate::backends_android_adapter::Grid, Error> {
    crate::backends_android_adapter::create_grid()
}

#[cfg(all(target_os = "android", not(feature = "zork")))]
pub fn create_dropdown(&self, items: &[&str]) -> Result<crate::backends_android_adapter::DropDown, Error> {
    crate::backends_android_adapter::create_dropdown(items)
}

#[cfg(all(target_os = "android", not(feature = "zork")))]
pub fn create_checkbutton(&self, label: &str) -> Result<crate::backends_android_adapter::CheckButton, Error> {
    crate::backends_android_adapter::create_checkbutton(label)
}

#[cfg(all(target_os = "android", not(feature = "zork")))]
pub fn create_radiobutton(&self, label: &str) -> Result<crate::backends_android_adapter::RadioButton, Error> {
    crate::backends_android_adapter::create_radiobutton(None, label)
}

#[cfg(all(target_os = "android", not(feature = "zork")))]
pub fn create_dialog(&self) -> Result<crate::backends_android_adapter::Dialog, Error> {
    crate::backends_android_adapter::create_dialog()
}

#[cfg(all(target_os = "android", not(feature = "zork")))]
pub fn create_textview(&self) -> Result<crate::backends_android_adapter::TextView, Error> {
    crate::backends_android_adapter::create_textview()
}

// ---------------------------------------------------------------------------
// High-level wrapper creation methods (return common types)
// ---------------------------------------------------------------------------

    /// Create a new Window and return a platform-independent handle.
    pub fn new_window(&self) -> Result<crate::common::Window, Error> {
        #[cfg(all(feature = "gtk", unix))]
        return crate::backends_gtk_adapter::create_window().map(|w| crate::common::Window { inner: w });
        #[cfg(windows)]
        return crate::backends_nwg_adapter::create_window(&self.parent_cell).map(|w| crate::common::Window { inner: w });
    }

    /// Create a new layout Box.
    pub fn new_box(&self, orientation: crate::common::Orientation, spacing: i32) -> Result<crate::common::WidgetBox, Error> {
        #[cfg(all(feature = "gtk", unix))]
        {
            let gtk_orient = match orientation {
                crate::common::Orientation::Horizontal => gtk_dynamic_loader::Orientation::Horizontal,
                crate::common::Orientation::Vertical => gtk_dynamic_loader::Orientation::Vertical,
            };
            return crate::backends_gtk_adapter::create_box(gtk_orient, spacing).map(|b| crate::common::WidgetBox { inner: b });
        }
        #[cfg(windows)]
        {
            let nwg_orient = match orientation {
                crate::common::Orientation::Horizontal => crate::backends::nwg::Orientation::Horizontal,
                crate::common::Orientation::Vertical => crate::backends::nwg::Orientation::Vertical,
            };
            let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
            return crate::backends_nwg_adapter::create_box(nwg_orient, spacing, parent).map(|b| crate::common::WidgetBox { inner: b });
        }
    }

    /// Create a new Label with the given text.
    pub fn new_label(&self, text: &str) -> Result<crate::common::Label, Error> {
        #[cfg(all(feature = "gtk", unix))]
        {
            let lbl = crate::backends_gtk_adapter::create_label(text)?;
            Ok(crate::common::Label { inner: lbl })
        }
        #[cfg(windows)]
        {
            let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
            let lbl = crate::backends_nwg_adapter::create_label(parent)?;
            lbl.set_text(text);
            Ok(crate::common::Label { inner: lbl })
        }
    }

    /// Create a new text Entry.
    pub fn new_entry(&self) -> Result<crate::common::Entry, Error> {
        #[cfg(all(feature = "gtk", unix))]
        return crate::backends_gtk_adapter::create_entry().map(|e| crate::common::Entry { inner: e });
        #[cfg(windows)]
        {
            let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
            return crate::backends_nwg_adapter::create_entry(parent).map(|e| crate::common::Entry { inner: e });
        }
    }

    /// Create a new Canvas (custom drawing surface).
    pub fn new_canvas(&self) -> Result<crate::common::Canvas, Error> {
        #[cfg(all(feature = "gtk", unix))]
        return crate::backends_gtk_adapter::create_canvas().map(|c| crate::common::Canvas { inner: c });
        #[cfg(windows)]
        {
            let parent = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
            return crate::backends_nwg_adapter::create_canvas(parent).map(|c| crate::common::Canvas { inner: c });
        }
    }

    /// Create a new Menu data model.
    pub fn new_menu(&self) -> Result<crate::common::Menu, Error> {
        #[cfg(all(feature = "gtk", unix))]
        return crate::backends_gtk_adapter::create_menu().map(|m| crate::common::Menu { inner: m });
        #[cfg(windows)]
        return crate::backends_nwg_adapter::create_menu().map(|m| crate::common::Menu { inner: m });
    }

    /// Create a new SimpleAction that will dispatch to the given name.
    /// On Windows the action is registered in the shared action registry.
    pub fn new_simple_action(&self, name: &str) -> Result<crate::common::SimpleAction, Error> {
        #[cfg(all(feature = "gtk", unix))]
        return crate::backends_gtk_adapter::create_simple_action(name).map(|a| crate::common::SimpleAction { inner: a });
        #[cfg(windows)]
        return crate::backends_nwg_adapter::create_simple_action(name, self.action_registry.clone()).map(|a| crate::common::SimpleAction { inner: a });
    }

    /// Create a MenuBar from a Menu model.
    /// `action_group` – on GTK a `*mut c_void` pointer to a `GActionGroup`
    /// (pass null if not available); on Windows it is unused.
    pub fn new_menubar(&self, model: &crate::common::Menu, action_group: *mut c_void) -> Result<crate::common::MenuBar, Error> {
        #[cfg(all(feature = "gtk", unix))]
        return unsafe { crate::backends_gtk_adapter::create_menubar(&model.inner, action_group).map(|m| crate::common::MenuBar { inner: m }) };
        #[cfg(windows)]
        {
            let hwnd = self.parent_cell.borrow().as_ref().copied().unwrap_or(std::ptr::null_mut());
            return crate::backends_nwg_adapter::create_menubar(&model.inner, hwnd, self.action_registry.clone()).map(|m| crate::common::MenuBar { inner: m });
        }
    }

    /// Create a new Dialog.
    pub fn new_dialog(&self) -> Result<crate::common::Dialog, Error> {
        #[cfg(all(feature = "gtk", unix))]
        return crate::backends_gtk_adapter::create_dialog().map(|d| crate::common::Dialog { inner: d });
        #[cfg(windows)]
        return crate::backends_nwg_adapter::create_dialog(&self.parent_cell).map(|d| crate::common::Dialog { inner: d });
    }

/// Run the backend main loop
    pub fn run(self) -> Result<(), Error> {
        let boxed = self.inner.borrow_mut().take().ok_or_else(|| Error::Backend("App::run already called".into()))?;
        boxed.run().map_err(|e| Error::Backend(format!("{}", e)))
    }
}

impl From<Box<dyn crate::backends::BackendApp>> for App {
    fn from(b: Box<dyn crate::backends::BackendApp>) -> Self {
        App {
            inner: Rc::new(RefCell::new(Some(b))),
            #[cfg(windows)]
            parent_cell: Rc::new(RefCell::new(None)),
            #[cfg(windows)]
            action_registry: Rc::new(RefCell::new(HashMap::new())),
        }
    }
}
