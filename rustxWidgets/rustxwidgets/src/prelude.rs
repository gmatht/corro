pub use crate::core::{App, DrawContext, Error, HandlerId, Widget};

/// Priority chain for backend type re-exports (same order as backends/mod.rs).
/// Only the highest-priority backend's types are exported to avoid name conflicts.
/// Others are available via their adapter module directly.

#[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
pub use crate::backends_gtk_adapter::{Window, Button, Label, BoxWidget, Grid, Entry, Menu, MenuBar, SimpleAction, Dialog, DropDown, CheckButton, RadioButton, TextView, Orientation, Canvas, Overlay, Spreadsheet, ScrolledWindow};

#[cfg(all(windows, not(feature = "zork")))]
pub use crate::backends_nwg_adapter::{Window, Button, Label, Menu, SimpleAction};

#[cfg(all(target_arch = "wasm32", not(feature = "zork")))]
pub use crate::backends_wasm_adapter::{Window, Button, Label, Menu, SimpleAction, Dialog, DropDown, CheckButton, RadioButton, TextView, Orientation};

#[cfg(all(target_os = "android", not(feature = "zork")))]
pub use crate::backends_android_adapter::{Window, Button, Label, Grid, DropDown, CheckButton, RadioButton, Dialog, TextView};

#[cfg(all(feature = "pancurses", not(any(feature = "gtk", windows, target_arch = "wasm32", target_os = "android"))))]
pub use crate::backends_pancurses_adapter::{Window, Button, Label, BoxWidget, Grid, Entry, Menu, MenuBar, SimpleAction, Dialog, DropDown, CheckButton, RadioButton, TextView, Orientation, Spreadsheet};

#[cfg(all(feature = "zork", not(feature = "pancurses")))]
pub use crate::backends_zork_adapter::{Window, Button, Label, BoxWidget, Grid, Entry, Menu, MenuBar, SimpleAction, Dialog, DropDown, CheckButton, RadioButton, TextView, Orientation};
