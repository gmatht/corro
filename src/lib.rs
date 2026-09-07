

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
pub mod ui_core;

// C99 CRT float shims, needed when linking the VC6 static CRT (rust9x msvc
// targets); see src/float_shim.rs. Public so dead-code analysis can't drop
// the #[no_mangle] items before the linker gets a chance to pull them in.
#[cfg(all(target_family = "rust9x", target_env = "msvc"))]
pub mod float_shim;
pub use ui_core::format_cell_display;

#[cfg(feature = "ratatui")]
pub mod ui;
#[cfg(any(feature = "gui", feature = "pancurses"))]
pub mod gui;
#[cfg(feature = "rustxwidgets-term")]
pub mod rustxwidgets_term;
