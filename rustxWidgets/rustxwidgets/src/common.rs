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

// Re-export the platform-specific types as the common types
#[cfg(all(feature = "gtk", target_os = "linux"))]
pub use platform::{Window, BoxWidget as WidgetBox, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog, DropDown, CheckButton, RadioButton, TextView, GtkOrientation as Orientation};

#[cfg(windows)]
pub use platform::{Window, BoxWidget as WidgetBox, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog, DropDown, CheckButton, RadioButton, TextView, NwgOrientation as Orientation};

#[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows))))]
pub use platform::{Window, BoxWidget as WidgetBox, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog, DropDown, CheckButton, RadioButton, TextView, PancursesOrientation as Orientation};

#[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
pub use platform::{Window, BoxWidget as WidgetBox, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog, DropDown, CheckButton, RadioButton, TextView, WasmOrientation as Orientation};

#[cfg(all(target_os = "android", not(feature = "zork")))]
pub use platform::{Window, BoxWidget as WidgetBox, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog, DropDown, CheckButton, RadioButton, TextView, AndroidOrientation as Orientation};

#[cfg(feature = "zork")]
pub use platform::{Window, BoxWidget as WidgetBox, Label, Entry, Canvas, Menu, SimpleAction, MenuBar, Dialog, DropDown, CheckButton, RadioButton, TextView, ZorkOrientation as Orientation};