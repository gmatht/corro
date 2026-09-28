//! iOS backend for corro: the `extern "C"` entry points + host bootstrap.
//!
//! The sheet itself renders through the shared [`super::gui_backend`]
//! pipeline (same `GuiState`, menus, dialogs, key handling as desktop and
//! Android); the only iOS-specific pieces live here:
//!
//! * [`ios_main`] — build the widget tree against the host's root view and
//!   hand the backend its root (the analogue of Android's `native_init`).
//! * [`run_ios_default`] — boot a blank corro [`super::App`] and leak it.
//! * [`install_menu_model`] — flatten the shared menu definition into the
//!   `app.*` action names the host's `UIMenu` dispatches.
//!
//! Why the `App` is leaked: UIKit's `UIApplicationMain` never returns to us,
//! so there is no stack frame to own the `App`. Every draw/input closure the
//! backend holds a raw pointer to keeps firing after this function returns,
//! so the `App` must outlive the call — the same reasoning as Android's
//! `run_android_default` and the wasm `main()`.
//!
//! See `rustxWidgets/docs/IOS_GUIDELINES.md` for the host-side contract (what
//! the Swift/ObjC app must supply) and `ios/corro/README.md` for the build.

// Only the iOS entry points below use this (they pass an empty path list to
// `App::new_with_paths`). Gated so a desktop build -- which compiles this
// module for the shared menu model -- does not warn about it.
#[cfg(target_os = "ios")]
use std::path::PathBuf;

/// Run the corro GUI on iOS. Mirrors `gui_backend::run_gui`: same
/// spreadsheet, same menus, same key handling — only the host window comes
/// from the app's root `UIView` instead of a desktop toplevel.
#[cfg(target_os = "ios")]
pub fn run_ios(corro_app: &mut super::App) -> Result<(), Box<dyn std::error::Error>> {
    super::gui_backend::run_gui(corro_app)
}

/// Boot a fresh corro [`super::App`] with no file (the app starts blank; a
/// document arrives later through the document picker) and run it.
///
/// The App is leaked — see the module docs.
#[cfg(target_os = "ios")]
pub fn run_ios_default() -> Result<(), Box<dyn std::error::Error>> {
    let app = super::App::new_with_paths(Vec::<PathBuf>::new());
    let app: &'static mut super::App = Box::leak(Box::new(app));
    app.set_backend(super::Backend::Gui);
    app.load_initial()?;
    run_ios(app)
}

/// Flatten the shared menu definition into `(top-level label, [(item label,
/// `app.<action>`)])` so the host can build its `UIMenu`.
///
/// A phone cannot show six text menus the way the desktop toolbar does, so
/// the host shows one "⋯" item whose submenu holds every top-level menu,
/// plus a few inline quick actions — the iOS form of the Android menu strip.
/// Submenus are flattened one level (a popup cannot nest arbitrarily on a
/// phone) and mnemonic underscores are dropped (meaningless on touch).
pub fn menu_model() -> Vec<(String, Vec<(String, String)>)> {
    use crate::gui::menu::{action_kind_to_name, MenuAction};

    fn collect(items: &[MenuAction], out: &mut Vec<(String, String)>) {
        for item in items {
            match item.submenu.as_deref() {
                Some(sub) => collect(sub, out),
                None => out.push((
                    item.label.replace('_', ""),
                    format!("app.{}", action_kind_to_name(item.action)),
                )),
            }
        }
    }

    let mut menus = Vec::new();
    for root in super::menu::menu_bar() {
        let mut items = Vec::new();
        collect(root.submenu.as_deref().unwrap_or(&[]), &mut items);
        menus.push((root.label.replace('_', ""), items));
    }
    menus
}

/// Run a menu action by name (called from the host's menu dispatch).
///
/// The name is the same `app.<action>` string the desktop menu registers,
/// so this is a thin lookup rather than a second dispatch table.
pub fn run_menu_action_by_name(name: &str) {
    super::gui_backend::dispatch_mobile_menu_action(name);
}

/// Build the menu model and hand it to the host. Called from `run_gui`
/// before the tree is presented.
///
/// Unlike the Android strip (a Java view the backend builds itself), the iOS
/// menu is a `UIBarButtonItem`, which cannot be constructed from Rust without
/// dragging UIKit view-controller types into this crate. So the *model* is
/// published instead: the host reads it via its own exported
/// `corro_ios_menu_model` callback (see `ios/corro/src/lib.rs`) and builds
/// the menu. A host that reads nothing still gets a working, menu-less UI.
/// Publish the menu model to the host (iOS only; see the body).
///
/// `pub(super)` because `GuiState` is crate-private: this is called from
/// `gui_backend::run_gui`, not from the app (the app reads the model through
/// the exported `corro_ios_menu_*` entry points in `ios/corro/src/lib.rs`).
#[cfg(target_os = "ios")]
pub(super) fn install_menu_model(
    _rxapp: &rswidgets::App,
    _shared: &std::rc::Rc<super::gui_backend::GuiState>,
) -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let menus = menu_model();
    let item_count: usize = menus.iter().map(|(_, items)| items.len()).sum();
    rswidgets::backends::ios::log_ios(&format!(
        "corro: menu model ready ({} menus, {} items); host reads it via corro_ios_menu_model",
        menus.len(),
        item_count
    ));
    Ok(())
}

/// Write a line to the iOS system log (`log stream --predicate 'eventMessage
/// CONTAINS "corro"'` shows these). Best-effort: the backend may not be
/// initialised yet, so failures are swallowed.
pub fn log_ios(msg: &str) {
    #[cfg(target_os = "ios")]
    rswidgets::backends::ios::log_ios(msg);
    #[cfg(not(target_os = "ios"))]
    eprintln!("corro(ios): {msg}");
}

/// A diagnostic line for the mobile backends, on the same footing as
/// [`log_ios`].
///
/// On iOS this goes through rswidgets' logger rather than `eprintln!`, and
/// that is the point. The iOS CI app is launched with `--console-pty`, so its
/// stderr is a PTY with a small fixed buffer; if nothing drains it, a blocking
/// `write(2)` stalls the caller - and the callers here are the edit-script
/// tick and the commit it performs, i.e. the main thread. rswidgets' logger
/// puts the descriptor in non-blocking mode, so a log line can never take the
/// app down. The per-edit commit line is the one diagnostic that says whether
/// the scripted edits ran at all, so it is the last thing that should be able
/// to stop them.
pub fn log_mobile(msg: &str) {
    log_ios(msg);
    // The same second copy rswidgets' logger makes, for the same reason: the
    // app's stderr is a PTY on the iOS CI runner and bursts of diagnostics are
    // dropped when the reader falls behind. A file does not drop them, and the
    // per-edit commit line is the one diagnostic that says whether the
    // scripted edits ran at all.
    append_to_app_log(msg);
}

/// Append one line to the app's own log file. Best-effort: a failure to open
/// is not worth reporting, because the stderr copy has already been made.
fn append_to_app_log(msg: &str) {
    use std::io::Write;
    let mut line = String::with_capacity(msg.len() + 1);
    line.push_str(msg);
    line.push('\n');
    if let Ok(mut f) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open("/tmp/corro-app.log")
    {
        let _ = f.write_all(line.as_bytes());
    }
}

// ---------------------------------------------------------------------------
// Touch gestures (pinch to zoom; drag pans unless a long press armed select)
// ---------------------------------------------------------------------------
//
// The gesture *policy* lives in the shared `gui_backend` (so Android, iOS and
// the widget all behave identically); these are the thin exports the app's
// `SheetView` shim calls. Each one is the iOS spelling of the matching
// `android_backend` entry point.

/// Apply a pinch step: multiply the sheet's view scale by `factor` (the
/// `UIPinchGestureRecognizer`'s scale ratio since the previous callback).
/// Returns the scale actually applied, clamped to the shared range.
#[cfg(target_os = "ios")]
pub fn zoom_viewport(factor: f64) -> f64 {
    super::gui_backend::zoom_viewport_by(factor)
}

/// Reset the pinch scale to 1.0 (a double-tap or a menu item). Returns 1.0.
#[cfg(target_os = "ios")]
pub fn reset_viewport_zoom() -> f64 {
    super::gui_backend::reset_viewport_zoom()
}

/// Begin a pointer gesture: `is_touch` is true for a finger (drag pans;
/// long press arms selection) and false for a mouse/trackpad (drag selects).
#[cfg(target_os = "ios")]
pub fn gesture_down(canvas_id: u64, x: f64, y: f64, is_touch: bool) -> i32 {
    super::gui_backend::mobile_gesture_down(canvas_id, x, y, is_touch)
}

/// The long-press recogniser fired: arm selection for the drag that follows.
#[cfg(target_os = "ios")]
pub fn gesture_long_press(canvas_id: u64, x: f64, y: f64) -> i32 {
    super::gui_backend::mobile_gesture_long_press(canvas_id, x, y)
}

/// A pointer move during a gesture. Returns the shared outcome code
/// (0 ignored, 1 select, 2 scroll, 3 tap, 4 long press).
#[cfg(target_os = "ios")]
pub fn gesture_move(canvas_id: u64, x: f64, y: f64) -> i32 {
    super::gui_backend::mobile_gesture_move(canvas_id, x, y)
}

/// A pointer release. A non-drag release is a tap, which the shared handler
/// has already applied by the time this returns.
#[cfg(target_os = "ios")]
pub fn gesture_up(canvas_id: u64, x: f64, y: f64) -> i32 {
    super::gui_backend::mobile_gesture_up(canvas_id, x, y)
}

/// Cancel an in-flight gesture (a system gesture stole the touch, the app
/// resigned active). Never produces a tap.
#[cfg(target_os = "ios")]
pub fn gesture_cancel(canvas_id: u64) {
    super::gui_backend::mobile_gesture_cancel(canvas_id);
}

/// Pan the sheet by a pixel delta from a scroll drag, converting to whole
/// cells with the current (zoom-aware) metrics. Returns `(d_rows, d_cols)`
/// actually applied, so the host can keep its sub-cell remainder.
#[cfg(target_os = "ios")]
pub fn drag_viewport(canvas_id: u64, dx: f64, dy: f64) -> (i32, i32) {
    super::gui_backend::drag_viewport_by_pixels(canvas_id, dx, dy)
}

/// The grid's row height and default column advance, in points, so the host
/// can convert a drag in points to whole cells with the same metrics the
/// renderer used (and the same ones a pinch scales). Returns `(row_h, col_w)`.
#[cfg(target_os = "ios")]
pub fn touch_cell_size() -> (f64, f64) {
    (super::gui_backend::row_h(), super::gui_backend::char_w())
}

/// Entry point called from the host once it has a root view.
///
/// Called by the `corro_ios` cdylib crate's `#[no_mangle] extern "C"`
/// `corro_ios_root_ready(root, view_controller)` (which lives there so the
/// linker keeps it). Wires the root view into `rswidgets::backends::ios` and
/// boots a blank corro [`super::App`].
#[cfg(target_os = "ios")]
pub fn ios_main(
    root: *mut std::os::raw::c_void,
    view_controller: *mut std::os::raw::c_void,
) -> Result<(), String> {
    // Step markers: a panic inside an extern "C" frame aborts the process with
    // no unwinding and no symbolised frames (the release build strips them), so
    // the only reliable way to locate a failure here is to say where we got to.
    // Cheap, and the alternative is a CI round trip per guess.
    log_ios("ios_main: init_with_root");
    // SAFETY: the host guarantees both pointers are live Objective-C objects
    // (a UIView and its UIViewController) that outlive the app — the same
    // contract Android's `nativeInit(activity, rootLayout)` has.
    unsafe {
        rswidgets::backends::ios::init_with_root(root, view_controller)
            .map_err(|e| format!("corro backend init failed: {e}"))?;
    }
    log_ios("ios_main: building the App");
    let app = super::App::new_with_paths(Vec::<PathBuf>::new());
    log_ios("ios_main: load_initial");
    let app: &'static mut super::App = Box::leak(Box::new(app));
    app.set_backend(super::Backend::Gui);
    app.load_initial().map_err(|e| format!("corro load failed: {e}"))?;
    log_ios("ios_main: entering run_gui");
    run_ios(app).map_err(|e| format!("corro run failed: {e}"))?;
    log_ios("ios_main: run_gui returned");
    Ok(())
}
