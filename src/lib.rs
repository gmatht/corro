
#![allow(unexpected_cfgs)] // Win95/rust9x custom target_family gates are intentional

pub mod capture;
pub mod core;
pub mod addr;
pub mod agg;
pub mod balance;
pub mod celladdr;
pub mod export;
pub mod formula;
pub mod extrapolate;
pub mod grid;
pub mod io;
pub mod ods;
pub mod ops;
pub mod debug_log;
pub mod lock;
pub mod ui_core;

// C99 CRT float shims, needed when linking the VC6 static CRT (rust9x msvc
// targets); see src/float_shim.rs. Public so dead-code analysis can't drop
// the #[no_mangle] items before the linker gets a chance to pull them in.
#[cfg(all(target_family = "rust9x", target_env = "msvc"))]
pub mod float_shim;
pub mod editor;
pub use ui_core::format_cell_display;

#[cfg(feature = "ratatui")]
pub mod ui;
/// The rswidgets GUI backends: `gui_backend` holds the shared widget tree
/// (menus, formula bar, sheet canvas, tabs) and is the reference fork every
/// other backend follows.
///
/// Enabled by `gui` (desktop GTK, and the base for pancurses / wasm) and by
/// `pancurses`. On `target_os = "android"` the `gui` feature additionally
/// exposes [`gui::android_backend`], which drives the same tree from the
/// Activity's content view via JNI (see `android/corro`). `examples/
/// android_ui.rs` builds that tree standalone for inspection.
///
/// `gui-macos` enables it for the AppKit backend: a native macOS build cannot
/// use `gui` (that feature resolves `rswidgets/gtk` + `gtk_dynamic_loader`,
/// which do not link on an Apple target), so without this the GUI module
/// would not exist at all on a Mac and `gui::macos_backend` would be
/// unreachable. See `Cargo.toml`.
#[cfg(any(feature = "gui", feature = "gui-core",
    feature = "gui-mobile",
    feature = "gui-macos",
    feature = "pancurses",
    target_arch = "wasm32"
))]
pub mod gui;
#[cfg(feature = "rswidgets-term")]
pub mod rswidgets_term;
