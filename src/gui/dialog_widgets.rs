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
//
// Gated on `gui` / `gui-core`, matching the dialog bodies that are this
// module's only consumer: the six dialog functions in `dialogs.rs` that
// `use super::dialog_widgets::{CheckButton, DropDown, RadioButton}` are each
// `#[cfg(any(feature = "gui", feature = "gui-core"))]`. A `pancurses`, `zork`
// or `wasm` build has no such body -- `pancurses` routes `App::run` to
// `pnc_backend`, and the other two compile their dialog bodies to no-ops -- so
// re-exporting here was a dead-code warning for an import nothing performs.
//
// The two arms above are deliberately *not* gated the same way: they are the
// definition of which backend supplies these widgets, and a build that runs a
// GUI needs them even where no dialog is wired.
// The `not(all(gui, linux|windows))` part is load-bearing, not decoration: the
// two arms above already name these types on a GTK desktop, so a predicate
// that also matched there would import them twice (E0252). The original
// comment said as much and this keeps that half of it.
#[cfg(all(
    any(feature = "gui", feature = "gui-core"),
    not(all(feature = "gui", any(target_os = "linux", windows)))
))]
pub use rswidgets::prelude::{CheckButton, DropDown, RadioButton};
