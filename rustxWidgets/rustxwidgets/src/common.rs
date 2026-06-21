//! Common type definitions shared across all backends.

/// Orientation for layout containers
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Orientation {
    Horizontal,
    Vertical,
}

// Platform-specific type re-exports using cfg
#[cfg(all(feature = "gtk", target_os = "linux"))]
mod platform {
    pub use crate::backends_gtk_adapter::{Window as PlatformWindow, BoxWidget as PlatformWidgetBox, Label as PlatformLabel, Entry as PlatformEntry, Canvas as PlatformCanvas, Menu as PlatformMenu, SimpleAction as PlatformSimpleAction, MenuBar as PlatformMenuBar, Dialog as PlatformDialog};
}

#[cfg(windows)]
mod platform {
    pub use crate::backends_nwg_adapter::{Window as PlatformWindow, BoxWidget as PlatformWidgetBox, Label as PlatformLabel, Entry as PlatformEntry, Canvas as PlatformCanvas, Menu as PlatformMenu, SimpleAction as PlatformSimpleAction, MenuBar as PlatformMenuBar, Dialog as PlatformDialog};
}

#[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows))))]
mod platform {
    pub use crate::backends_pancurses_adapter::{Window as PlatformWindow, BoxWidget as PlatformWidgetBox, Label as PlatformLabel, Entry as PlatformEntry, Canvas as PlatformCanvas, Menu as PlatformMenu, SimpleAction as PlatformSimpleAction, MenuBar as PlatformMenuBar, Dialog as PlatformDialog};
}

#[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
mod platform {
    pub use crate::backends_wasm_adapter::{Window as PlatformWindow, BoxWidget as PlatformWidgetBox, Label as PlatformLabel, Entry as PlatformEntry, Canvas as PlatformCanvas, Menu as PlatformMenu, SimpleAction as PlatformSimpleAction, MenuBar as PlatformMenuBar, Dialog as PlatformDialog};
}

#[cfg(all(target_os = "android", not(feature = "zork")))]
mod platform {
    pub use crate::backends_android_adapter::{Window as PlatformWindow, BoxWidget as PlatformWidgetBox, Label as PlatformLabel, Entry as PlatformEntry, Canvas as PlatformCanvas, Menu as PlatformMenu, SimpleAction as PlatformSimpleAction, MenuBar as PlatformMenuBar, Dialog as PlatformDialog};
}

#[cfg(feature = "zork")]
mod platform {
    pub use crate::backends_zork_adapter::{Window as PlatformWindow, BoxWidget as PlatformWidgetBox, Label as PlatformLabel, Entry as PlatformEntry, Canvas as PlatformCanvas, Menu as PlatformMenu, SimpleAction as PlatformSimpleAction, MenuBar as PlatformMenuBar, Dialog as PlatformDialog};
}

// Common wrapper types with `inner` field for the platform-specific types
#[cfg(all(feature = "gtk", target_os = "linux"))]
mod common_types {
    use super::platform::*;
    pub struct Window { pub inner: PlatformWindow }
    pub struct WidgetBox { pub inner: PlatformWidgetBox }
    pub struct Label { pub inner: PlatformLabel }
    pub struct Entry { pub inner: PlatformEntry }
    pub struct Canvas { pub inner: PlatformCanvas }
    pub struct Menu { pub inner: PlatformMenu }
    pub struct SimpleAction { pub inner: PlatformSimpleAction }
    pub struct MenuBar { pub inner: PlatformMenuBar }
    pub struct Dialog { pub inner: PlatformDialog }
}

#[cfg(windows)]
mod common_types {
    use super::platform::*;
    pub struct Window { pub inner: PlatformWindow }
    pub struct WidgetBox { pub inner: PlatformWidgetBox }
    pub struct Label { pub inner: PlatformLabel }
    pub struct Entry { pub inner: PlatformEntry }
    pub struct Canvas { pub inner: PlatformCanvas }
    pub struct Menu { pub inner: PlatformMenu }
    pub struct SimpleAction { pub inner: PlatformSimpleAction }
    pub struct MenuBar { pub inner: PlatformMenuBar }
    pub struct Dialog { pub inner: PlatformDialog }
}

#[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows))))]
mod common_types {
    use super::platform::*;
    pub struct Window { pub inner: PlatformWindow }
    pub struct WidgetBox { pub inner: PlatformWidgetBox }
    pub struct Label { pub inner: PlatformLabel }
    pub struct Entry { pub inner: PlatformEntry }
    pub struct Canvas { pub inner: PlatformCanvas }
    pub struct Menu { pub inner: PlatformMenu }
    pub struct SimpleAction { pub inner: PlatformSimpleAction }
    pub struct MenuBar { pub inner: PlatformMenuBar }
    pub struct Dialog { pub inner: PlatformDialog }
}

#[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
mod common_types {
    use super::platform::*;
    pub struct Window { pub inner: PlatformWindow }
    pub struct WidgetBox { pub inner: PlatformWidgetBox }
    pub struct Label { pub inner: PlatformLabel }
    pub struct Entry { pub inner: PlatformEntry }
    pub struct Canvas { pub inner: PlatformCanvas }
    pub struct Menu { pub inner: PlatformMenu }
    pub struct SimpleAction { pub inner: PlatformSimpleAction }
    pub struct MenuBar { pub inner: PlatformMenuBar }
    pub struct Dialog { pub inner: PlatformDialog }
}

#[cfg(all(target_os = "android", not(feature = "zork")))]
mod common_types {
    use super::platform::*;
    pub struct Window { pub inner: PlatformWindow }
    pub struct WidgetBox { pub inner: PlatformWidgetBox }
    pub struct Label { pub inner: PlatformLabel }
    pub struct Entry { pub inner: PlatformEntry }
    pub struct Canvas { pub inner: PlatformCanvas }
    pub struct Menu { pub inner: PlatformMenu }
    pub struct SimpleAction { pub inner: PlatformSimpleAction }
    pub struct MenuBar { pub inner: PlatformMenuBar }
    pub struct Dialog { pub inner: PlatformDialog }
}

#[cfg(feature = "zork")]
mod common_types {
    use super::platform::*;
    pub struct Window { pub inner: PlatformWindow }
    pub struct WidgetBox { pub inner: PlatformWidgetBox }
    pub struct Label { pub inner: PlatformLabel }
    pub struct Entry { pub inner: PlatformEntry }
    pub struct Canvas { pub inner: PlatformCanvas }
    pub struct Menu { pub inner: PlatformMenu }
    pub struct SimpleAction { pub inner: PlatformSimpleAction }
    pub struct MenuBar { pub inner: PlatformMenuBar }
    pub struct Dialog { pub inner: PlatformDialog }
}

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