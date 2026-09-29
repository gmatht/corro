//! Native dialog widget types, resolved to the right backend once.
//!
//! Every dialog in [`super::dialogs`] builds the same handful of native
//! widgets (`DropDown`, `CheckButton`, `RadioButton`, `Entry`). Under the
//! combined-gui feature the crate root's `prelude` is flipped to the
//! pancurses adapter's types, so a dialog that just says `use
//! rswidgets::prelude::DropDown` would get the wrong (terminal) widget on
//! Linux/Windows — the dialogs must always use the *native* backend's types.
//!
//! That choice is a three-branch `cfg` (gtk on Linux, nwg on Windows, the
//! prelude elsewhere) and used to be repeated verbatim at every dialog
//! construction site. Naming the aliases once here means a new dialog cannot
//! pick the wrong backend's widget, and the platform mapping lives in one
//! place.

// The toolkit's own adapter, where one applies. Deliberately *not* selected by
// `gui-core`: that feature is deliberately toolkit-free, so naming GTK/NWG here
// would drag a toolkit onto a headless build.
#[cfg(all(feature = "gui", target_os = "linux"))]
pub use rswidgets::backends_gtk_adapter::{CheckButton, DropDown, RadioButton};

#[cfg(all(feature = "gui", windows))]
pub use rswidgets::backends_nwg_adapter::{CheckButton, DropDown, RadioButton};

// Everything else — macOS, the mobile backends, and the toolkit-free `gui-core`
// (which is what `zork` selects) — takes the prelude, i.e. whichever backend
// `backends::init` will pick. This must not overlap the two arms above, or the
// names are imported twice.
// Re-exported for `dialogs.rs`, which uses them from any GUI-capable build.
// A `pancurses`-only build compiles that module but constructs none of these
// (it has its own terminal dialog path), so the import is legitimately unused
// there — hence the allow rather than a cfg that would fork the caller too.
#[cfg(not(all(feature = "gui", any(target_os = "linux", windows))))]
#[allow(unused_imports)]
pub use rswidgets::prelude::{CheckButton, DropDown, RadioButton};
