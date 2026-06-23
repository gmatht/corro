//! Common type definitions shared across all backends.

// Platform-specific type re-exports using cfg
#[cfg(all(feature = "gtk", target_os = "linux"))]
mod platform {
    pub use crate::backends_gtk_adapter::{Window, BoxWidget, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog, DropDown, CheckButton, RadioButton, TextView, Orientation as GtkOrientation};
}

#[cfg(windows)]
mod platform {
    pub use crate::backends_nwg_adapter::{Window, BoxWidget, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog, DropDown, CheckButton, RadioButton, TextView, Orientation as NwgOrientation};
}

#[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows))))]
mod platform {
    pub use crate::backends_pancurses_adapter::{Window, BoxWidget, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog, DropDown, CheckButton, RadioButton, TextView, Orientation as PancursesOrientation};
}

#[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
mod platform {
    pub use crate::backends_wasm_adapter::{Window, BoxWidget, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog, DropDown, CheckButton, RadioButton, TextView, Orientation as WasmOrientation};
}

#[cfg(all(target_os = "android", not(feature = "zork")))]
mod platform {
    pub use crate::backends_android_adapter::{Window, BoxWidget, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog, DropDown, CheckButton, RadioButton, TextView, Orientation as AndroidOrientation};
}

#[cfg(feature = "zork")]
mod platform {
    pub use crate::backends_zork_adapter::{Window, BoxWidget, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog, DropDown, CheckButton, RadioButton, TextView, Orientation as ZorkOrientation};
}

macro_rules! common_types_mod {
    () => {
        use super::platform::*;
        use crate::core::Widget;

        #[derive(Clone)]
        pub struct Window { pub inner: PlatformWindow }
        #[derive(Clone)]
        pub struct WidgetBox { pub inner: PlatformWidgetBox }
        #[derive(Clone)]
        pub struct Label { pub inner: PlatformLabel }
        #[derive(Clone)]
        pub struct Entry { pub inner: PlatformEntry }
        #[derive(Clone)]
        pub struct Canvas { pub inner: PlatformCanvas }
        #[derive(Clone)]
        pub struct Menu { pub inner: PlatformMenu }
        #[derive(Clone)]
        pub struct SimpleAction { pub inner: PlatformSimpleAction }
        #[derive(Clone)]
        pub struct MenuBar { pub inner: PlatformMenuBar }
        #[derive(Clone)]
        pub struct Dialog { pub inner: PlatformDialog }

        impl Window {
            pub fn set_title(&self, title: &str) { self.inner.set_title(title); }
            pub fn set_default_size(&self, w: i32, h: i32) { self.inner.set_default_size(w, h); }
            pub fn present(&self) { self.inner.present(); }
            pub fn insert_action_group(&self, name: &str, group_ptr: *mut std::os::raw::c_void) { self.inner.insert_action_group(name, group_ptr); }
            pub fn hwnd(&self) -> *mut std::os::raw::c_void { self.inner.hwnd() }
            pub fn set_child_box(&self, bx: &WidgetBox) { self.inner.set_child_box(&bx.inner); }
        }
        impl WidgetBox {
            pub fn append(&mut self, child: &impl AsRef<*mut std::os::raw::c_void>) { self.inner.append(child); }
            pub fn set_child_hexpand(&self, child: &impl AsRef<*mut std::os::raw::c_void>, expand: bool) { self.inner.set_child_hexpand(child, expand); }
            pub fn set_child_vexpand(&self, child: &impl AsRef<*mut std::os::raw::c_void>, expand: bool) { self.inner.set_child_vexpand(child, expand); }
            pub fn set_hexpand(&self, expand: bool) { self.inner.set_hexpand(expand); }
        }
        impl AsRef<*mut std::os::raw::c_void> for WidgetBox {
            fn as_ref(&self) -> &*mut std::os::raw::c_void { self.inner.as_ref() }
        }
        impl Label {
            pub fn set_text(&self, text: &str) { self.inner.set_text(text); }
            pub fn get_text(&self) -> Option<String> { self.inner.get_text() }
            pub fn set_visible(&self, visible: bool) { self.inner.set_visible(visible); }
            pub fn set_markup(&self, markup: &str) { self.inner.set_markup(markup); }
            pub fn raw_handle(&self) -> *mut std::os::raw::c_void { self.inner.raw_handle() }
        }
        impl AsRef<*mut std::os::raw::c_void> for Label {
            fn as_ref(&self) -> &*mut std::os::raw::c_void { self.inner.as_ref() }
        }
        impl Entry {
            pub fn set_text(&self, text: &str) { self.inner.set_text(text); }
            pub fn get_text(&self) -> Option<String> { self.inner.get_text() }
            pub fn grab_focus(&self) { self.inner.grab_focus(); }
            pub fn set_hexpand(&self, expand: bool) { self.inner.set_hexpand(expand); }
            pub fn set_size_request(&self, w: i32, h: i32) { self.inner.set_size_request(w, h); }
            pub fn set_visible(&self, v: bool) { self.inner.set_visible(v); }
            pub fn connect_changed(&self, f: impl FnMut() + 'static) -> Result<u64, crate::Error> { self.inner.connect_changed(f) }
        }
        impl AsRef<*mut std::os::raw::c_void> for Entry {
            fn as_ref(&self) -> &*mut std::os::raw::c_void { self.inner.as_ref() }
        }
        impl Canvas {
            pub fn set_draw_callback(&self, cb: Box<dyn FnMut(&mut dyn crate::core::DrawContext, i32, i32)>) { self.inner.set_draw_callback(cb); }
            pub fn queue_redraw(&self) { self.inner.queue_redraw(); }
            pub fn set_size_request(&self, w: i32, h: i32) { self.inner.set_size_request(w, h); }
            pub fn on_click(&self, cb: Box<dyn FnMut(f64, f64)>) { self.inner.on_click(cb); }
            pub fn set_content_size(&self, w: i32, h: i32) { self.inner.set_content_size(w, h); }
        }
        impl AsRef<*mut std::os::raw::c_void> for Canvas {
            fn as_ref(&self) -> &*mut std::os::raw::c_void { self.inner.as_ref() }
        }
        impl Menu {
            pub fn append(&mut self, label: &str, detailed_action: &str) { self.inner.append(label, detailed_action); }
            pub fn append_submenu(&mut self, label: &str, submenu: &Menu) { self.inner.append_submenu(label, &submenu.inner); }
        }
        impl SimpleAction {
            pub fn connect_activate<F: FnMut(*mut std::os::raw::c_void) + 'static>(&self, f: F) -> Result<u64, crate::Error> { self.inner.connect_activate(f) }
        }
        impl MenuBar {
        }
        impl AsRef<*mut std::os::raw::c_void> for MenuBar {
            fn as_ref(&self) -> &*mut std::os::raw::c_void { self.inner.as_ref() }
        }
        impl Dialog {
            pub fn set_title(&self, title: &str) { self.inner.set_title(title); }
            pub fn set_default_size(&self, w: i32, h: i32) { self.inner.set_default_size(w, h); }
            pub fn append_content_area(&self, child: &impl AsRef<*mut std::os::raw::c_void>) { self.inner.append_content_area(child); }
            pub fn add_button(&self, text: &str, response_id: i32) { self.inner.add_button(text, response_id); }
            pub fn present(&self) { self.inner.present(); }
            pub fn connect_response<F: FnMut(i32) + 'static>(&self, f: F) -> Result<u64, crate::Error> { self.inner.connect_response(f) }
            pub fn close(&self) { self.inner.close(); }
        }
        impl AsRef<*mut std::os::raw::c_void> for Dialog {
            fn as_ref(&self) -> &*mut std::os::raw::c_void { self.inner.as_ref() }
        }
    }
}

// Common wrapper types with `inner` field for the platform-specific types
#[cfg(all(feature = "gtk", target_os = "linux"))]
mod common_types {
    common_types_mod!();
    impl Canvas {
        pub fn on_key(&self, cb: Box<dyn FnMut(u32) -> bool>) {
            self.inner.on_key(Box::new(move |k: u32, _s: u32| -> bool { cb(k) }));
        }
    }
}

#[cfg(windows)]
mod common_types {
    common_types_mod!();
    impl Canvas {
        pub fn on_key(&self, cb: Box<dyn FnMut(u32) -> bool>) { self.inner.on_key(cb); }
    }
    impl Entry {
        pub fn on_key(&self, f: Box<dyn FnMut(u32) -> bool>) { self.inner.on_key(f); }
    }
}

#[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows))))]
mod common_types { common_types_mod!(); }

#[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
mod common_types { common_types_mod!(); }

#[cfg(all(target_os = "android", not(feature = "zork")))]
mod common_types { common_types_mod!(); }

#[cfg(feature = "zork")]
mod common_types { common_types_mod!(); }

#[cfg(all(feature = "gtk", target_os = "linux"))]
pub use common_types::{Window, WidgetBox, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog};

#[cfg(windows)]
pub use common_types::{Window, WidgetBox, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog};

#[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows))))]
pub use common_types::{Window, WidgetBox, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog};

#[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
pub use common_types::{Window, WidgetBox, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog};

#[cfg(all(target_os = "android", not(feature = "zork")))]
pub use common_types::{Window, WidgetBox, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog};

#[cfg(feature = "zork")]
pub use common_types::{Window, WidgetBox, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog};
