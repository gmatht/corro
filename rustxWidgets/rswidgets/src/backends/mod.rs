use std::error::Error as StdError;

/// Type alias for boxed backend errors
pub type BackendError = Box<dyn StdError + Send + Sync>;

/// Backend application abstraction. Concrete backends provide an implementor boxed via `init()`.
pub trait BackendApp {
    /// Run the backend main loop. Consumes the backend app.
    fn run(self: Box<Self>) -> Result<(), BackendError>;
}

/// Priority chain: each backend module is always compiled when its feature is on,
/// but `init` is re-exported only for the highest-priority backend available.
/// Platform-specific backends (gtk, nwg, wasm, android) naturally exclude each
/// other. Pancurses is a fallback when no platform-native backend applies.
///
/// At runtime, `BACKEND` env var selects between compiled backends:
///   BACKEND=gtk4     uses the new gtk4-rs/dlopen backend
///   BACKEND=gtk3     uses the old gtk_dynamic_loader backend (GTK3/GTK4)
///   (unset)          defaults to gtk4-rs if available, otherwise gtk3

#[cfg(all(feature = "gtk4-rs", target_os = "linux", not(feature = "zork")))]
pub mod gtk4_rs;
// The module itself is always available when its feature is on (combined
// builds need gtk + pancurses side by side); only the `init` re-export
// below is priority-gated so a native backend wins over pancurses —
// pancurses is the fallback (see App::init's docs), reachable explicitly
// via `backends::pancurses::init()` (which is how run_pancurses gets it).
#[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork")))]
pub mod gtk;

#[cfg(all(feature = "gtk4-rs", target_os = "linux", not(feature = "zork"), not(feature = "gtk")))]
pub use self::gtk4_rs::init;

#[cfg(all(feature = "gtk", target_os = "linux", not(feature = "zork"), not(feature = "gtk4-rs")))]
pub use self::gtk::init;

#[cfg(all(feature = "gtk4-rs", feature = "gtk", target_os = "linux", not(feature = "zork")))]
pub fn init() -> Result<Box<dyn BackendApp>, BackendError> {
    let backend = std::env::var("BACKEND").unwrap_or_default();
    if backend == "gtk3" || backend == "gtk" {
        self::gtk::init()
    } else {
        self::gtk4_rs::init()
    }
}

// Same side-by-side story as gtk: the NWG module is always available on
// Windows so combined (gui+pancurses) builds can drive the native GUI
// while the terminal uses `backends::pancurses::init()` directly.
#[cfg(all(windows, not(feature = "zork")))]
pub mod nwg;
#[cfg(all(windows, not(feature = "zork")))]
pub use self::nwg::init;

#[cfg(all(target_arch = "wasm32", not(feature = "pancurses"), not(feature = "zork")))]
pub mod wasm;
#[cfg(all(target_arch = "wasm32", not(feature = "pancurses"), not(feature = "zork")))]
pub use self::wasm::init;

#[cfg(all(target_os = "android", not(feature = "zork")))]
pub mod android;
#[cfg(all(target_os = "android", not(feature = "zork")))]
pub use self::android::init_backend as init;

#[cfg(feature = "pancurses")]
pub mod pancurses;
// Native backends win the default `init` (documented priority); the
// pancurses re-export applies only where no native backend is compiled,
// otherwise it would collide with the native one. Explicit
// `backends::pancurses::init()` stays available everywhere for TUI paths.
#[cfg(all(
    feature = "pancurses",
    not(feature = "zork"),
    not(all(feature = "gtk4-rs", target_os = "linux")),
    not(all(feature = "gtk", target_os = "linux")),
    not(windows)
))]
pub use self::pancurses::init;

#[cfg(feature = "pancurses")]
pub mod pancurses_draw;

#[cfg(feature = "ratatui")]
pub mod ratatui;
#[cfg(all(feature = "ratatui", not(any(feature = "gtk", feature = "gtk4-rs", target_os = "windows", target_arch = "wasm32", target_os = "android", feature = "pancurses", feature = "zork"))))]
pub use self::ratatui::init;

// Always available (not feature-gated): dependency-free pure Rust whose
// documented purpose is testing and behaviour comparison. Unit tests
// (e.g. spreadsheet style coverage) import it unconditionally, so gating
// it broke bare `cargo test`. The `headless` feature name is retained as
// a harmless no-op for existing feature lists.
/// Headless draw-context backend (recording, for tests).
pub mod headless;

#[cfg(feature = "zork")]
pub mod zork;
#[cfg(feature = "zork")]
pub use self::zork::init;
