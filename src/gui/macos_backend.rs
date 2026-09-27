//! macOS backend for corro: the `extern "C"` entry point + host bootstrap.
//!
//! The AppKit sibling of [`super::ios_backend`], and — deliberately — very
//! little code, because almost nothing about corro is platform-specific:
//!
//! * The spreadsheet itself (`super::gui_backend::run_gui`: the `GuiState`,
//!   menus, dialogs, key handling, draw replay) is **the same code** on macOS
//!   as on iOS, Android, GTK, Windows and the TUI. This module does not
//!   re-implement any of it, and adds no macos branches to it.
//! * The widget tree goes through `rswidgets::common::{Window, Canvas, ...}`,
//!   which resolves to the AppKit adapter (`rswidgets::backends_macos_adapter`)
//!   on `target_os = "macos"` — the same seam every other backend uses.
//! * The menu model and the action dispatcher are plain data derived from the
//!   shared menu definition, so they are **literally the same functions** the
//!   iOS backend uses; they are re-exported below rather than copied.
//!
//! What is genuinely macOS-specific here:
//!
//! * [`macos_main`] — root the tree in the host window's content view via
//!   `rswidgets::backends::macos::init_with_root` instead of a desktop
//!   toplevel or the app's root `UIView`.
//! * [`log_macos`] — the system-log line, which is `NSLog` on both Apple
//!   platforms but tagged differently by the host docs.
//!
//! Why the `App` is leaked: the host owns `NSApplication` and calls
//! `-[NSApp run]`, which never returns to us, so there is no stack frame to
//! own the `App`. Every draw/input closure the backend holds a raw pointer to
//! keeps firing after this function returns — exactly the reasoning behind
//! iOS's `ios_main`, Android's `run_android_default` and the wasm `main()`.
//!
//! See `rustxWidgets/docs/MACOS_GUIDELINES.md` for the host-side contract
//! (what the ObjC app must supply) and `rustxWidgets/docs/IOS_GUIDELINES.md`
//! for the iOS twin of this document.
//!
//! The `extern "C"` entry points live in the **`macos/corro`** host crate,
//! not here: they have to be at a crate root to survive the linker, exactly
//! as `ios/corro/src/lib.rs` holds the iOS ones. That crate is the macOS twin
//! of `ios/corro` and carries the AppKit shims alongside them.

// Only the macOS entry points below use this; gated for the same reason as
// the iOS twin (a desktop build compiles this module for the shared model).
#[cfg(target_os = "macos")]
use std::path::PathBuf;

// The menu model and the action dispatcher are *shared*, not per-platform: on
// a desktop they are exercised by `examples/ios_ui.rs`, and both mobile hosts
// (iOS and macOS) read the same `app.*` names out of them. Re-exporting keeps
// one implementation and one set of tests.
pub use super::ios_backend::{menu_model, run_menu_action_by_name};

/// Run the corro GUI on macOS. Mirrors `gui_backend::run_gui` — same
/// spreadsheet, same menus, same key handling — only the host window comes
/// from the window's content view instead of a desktop toplevel.
#[cfg(target_os = "macos")]
pub fn run_macos(corro_app: &mut super::App) -> Result<(), Box<dyn std::error::Error>> {
    super::gui_backend::run_gui(corro_app)
}

/// Boot a fresh corro [`super::App`] with no file and run it.
///
/// Unlike iOS — where documents arrive only through `UIDocumentPicker` — macOS
/// has ordinary paths and a normal filesystem, so a real host is expected to
/// pass the file it was opened with (or one from `NSApplicationDelegate`'s
/// `application:openFile:`). This is the document-less bootstrap, kept for
/// symmetry with [`super::ios_backend::run_ios_default`] and for a host that
/// just wants a blank sheet.
///
/// The App is leaked — see the module docs.
#[cfg(target_os = "macos")]
pub fn run_macos_default() -> Result<(), Box<dyn std::error::Error>> {
    run_macos_with_paths(Vec::<PathBuf>::new())
}

/// As [`run_macos_default`], but opening `paths` (the ordinary macOS launch
/// path: the app is handed the documents to open).
#[cfg(target_os = "macos")]
pub fn run_macos_with_paths(
    paths: Vec<PathBuf>,
) -> Result<(), Box<dyn std::error::Error>> {
    let app = super::App::new_with_paths(paths);
    let app: &'static mut super::App = Box::leak(Box::new(app));
    app.set_backend(super::Backend::Gui);
    app.load_initial()?;
    run_macos(app)
}

/// Publish the menu model to the host. Called from `run_gui` before the tree
/// is presented.
///
/// Unlike iOS — whose host must build a `UIMenu`/`UIBarButtonItem` because a
/// phone has no menubar — macOS has a real `NSMenu`, so the host can build a
/// genuine menubar from this model. The *model* is still what is published,
/// because that is the one shape all three mobile/native backends share; the
/// host chooses how to render it.
///
/// `pub(super)` because `GuiState` is crate-private: this is called from
/// `gui_backend::run_gui`, not from the app.
#[cfg(target_os = "macos")]
pub(super) fn install_menu_model(
    _rxapp: &rswidgets::App,
    _shared: &std::rc::Rc<super::gui_backend::GuiState>,
) -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let menus = menu_model();
    let item_count: usize = menus.iter().map(|(_, items)| items.len()).sum();
    log_macos(&format!(
        "corro: menu model ready ({} menus, {} items); host reads it via corro_macos_menu_model",
        menus.len(),
        item_count
    ));
    Ok(())
}

/// Write a line to the macOS system log (Console.app, or
/// `log stream --predicate 'eventMessage CONTAINS "corro"'`). Best-effort: the
/// backend may not be initialised yet, so failures are swallowed.
pub fn log_macos(msg: &str) {
    #[cfg(target_os = "macos")]
    rswidgets::backends::macos::log_macos(msg);
    #[cfg(not(target_os = "macos"))]
    eprintln!("corro(macos): {msg}");
}

// ---------------------------------------------------------------------------
// The host-facing surface (`macos/corro` calls these)
// ---------------------------------------------------------------------------
//
// The three below are the *only* things a macOS host needs beyond the bootstrap
// and the menu: they wrap the shared GUI helpers that live behind
// `pub(crate)` in `gui_backend` (a module a cdylib cannot reach), so the host
// crate has exactly one place to call instead of a scatter of internal paths.
//
// Unlike their iOS twins these are **not** `#[cfg(target_os = "macos")]`, and
// that is deliberate rather than an oversight. A macOS host has a *mouse* and
// a *scroll wheel*, so the metrics a pointer host needs to turn its own pixel
// deltas into whole cells are part of the contract -- and a contract that can
// only be checked on the one platform it is for is a contract that is never
// checked. Compiling them everywhere is what lets
// `tests/macos_pipeline_preview.rs` assert on Linux CI that the numbers a host
// would use are positive and sane, which is the failure that would otherwise
// divide by zero in the host's pixel-to-cell conversion and only surface on a
// real Mac, on a Retina display, zoomed in.

/// `(row height, column advance)` in points, for a scroll wheel or trackpad
/// that has to turn its own pixel delta into whole cells. Both numbers already
/// carry the Retina factor and the pinch scale, so a host must not compute its
/// own.
pub fn cell_size() -> (f64, f64) {
    super::gui_backend::cell_size()
}

/// Scroll by whole cells -- what a scroll wheel notch or a momentum trackpad
/// pan ends in. Returns the applied `(d_rows, d_cols)`, which can be less
/// than requested at a sheet edge, so the host knows to stop accumulating.
///
/// The move itself needs a published `GuiState`, which only a real window has,
/// so off-target this is the documented "nothing moved" answer rather than a
/// panic -- which is also what a host gets before its first frame.
pub fn scroll_by_cells(d_rows: i32, d_cols: i32) -> (i32, i32) {
    #[cfg(target_os = "macos")]
    {
        super::gui_backend::scroll_by_cells(d_rows, d_cols)
    }
    #[cfg(not(target_os = "macos"))]
    {
        // No live viewport off-target: report no movement, so a host's
        // accumulation loop terminates instead of spinning.
        let _ = (d_rows, d_cols);
        (0, 0)
    }
}

/// Zoom the sheet by `factor` (a trackpad pinch's scale ratio since the last
/// callback), returning the scale actually applied.
pub fn zoom_viewport_by(factor: f64) -> f64 {
    #[cfg(target_os = "macos")]
    {
        super::gui_backend::zoom_viewport_by(factor)
    }
    #[cfg(not(target_os = "macos"))]
    {
        // The pinch scale is a plain thread-local and needs no window, so the
        // *clamping* half is host-independent and genuinely testable
        // off-target -- which is worth keeping rather than stubbing to 1.0.
        super::gui_backend::clamped_view_zoom(factor)
    }
}

/// Reset the pinch scale to 1.0. Returns the applied scale (always 1.0).
pub fn reset_viewport_zoom() -> f64 {
    #[cfg(target_os = "macos")]
    {
        super::gui_backend::reset_viewport_zoom()
    }
    #[cfg(not(target_os = "macos"))]
    {
        super::gui_backend::reset_view_zoom_scale()
    }
}

/// Entry point called from the host once it has a content view.
///
/// Called by a host cdylib's `#[no_mangle] extern "C"`
/// `corro_macos_root_ready(root, window_controller)` (which lives in the host
/// so the linker keeps it). Wires the content view into
/// `rswidgets::backends::macos` and boots corro.
#[cfg(target_os = "macos")]
pub fn macos_main(
    root: *mut std::os::raw::c_void,
    window_controller: *mut std::os::raw::c_void,
) -> Result<(), String> {
    // Step markers: a panic inside an extern "C" frame aborts the process with
    // no unwinding and no symbolised frames, so the only reliable way to locate
    // a startup failure is to say where we got to. Same technique as iOS.
    log_macos("macos_main: init_with_root");
    // SAFETY: the host guarantees both pointers are live Objective-C objects
    // (an NSView and its NSWindowController) that outlive the app — the same
    // contract iOS's `ios_main(UIView*, UIViewController*)` has.
    unsafe {
        rswidgets::backends::macos::init_with_root(root, window_controller)
            .map_err(|e| format!("corro backend init failed: {e}"))?;
    }
    log_macos("macos_main: building the App");
    let app = super::App::new_with_paths(Vec::<PathBuf>::new());
    log_macos("macos_main: load_initial");
    let app: &'static mut super::App = Box::leak(Box::new(app));
    app.set_backend(super::Backend::Gui);
    app.load_initial().map_err(|e| format!("corro load failed: {e}"))?;
    log_macos("macos_main: entering run_gui");
    run_macos(app).map_err(|e| format!("corro run failed: {e}"))?;
    log_macos("macos_main: run_gui returned");
    Ok(())
}
