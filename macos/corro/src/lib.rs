//! corro macOS staticlib/cdylib: the `extern "C"` entry points an app's
//! Objective-C shims call.
//!
//! Everything else lives in the `corro` library
//! (`corro::gui::macos_backend`): the shared `run_gui` spreadsheet pipeline
//! plus the menu model. Keeping the exports here (crate root of a
//! staticlib/cdylib) guarantees the linker retains them — the same reason
//! `ios/corro/src/lib.rs` and `android/corro/src/lib.rs` hold theirs.
//!
//! The names are the contract between this crate and `app/`: every selector in
//! `rustxWidgets/docs/MACOS_GUIDELINES.md` §3 is a `corro_macos_*` export
//! below, and `app/CorroMacBridge.h` declares them. Before this crate existed
//! those names were documented but *defined nowhere*: a macOS host had no way
//! to get a first frame, because the sheet view's `drawRect:` had nothing to
//! call. That is the gap this file closes, and it is the reason the macOS path
//! is now feature-complete rather than merely type-checking.
//!
//! Two deliberate differences from the iOS host crate:
//!
//! * **No gesture recognisers.** AppKit delivers a real
//!   `mouseDown:`/`mouseDragged:`/`mouseUp:` sequence with a button number, and
//!   the *shared* pointer path (`Canvas::on_click_button` / `on_motion` /
//!   `on_release`) already consumes it. The synthetic touch recognisers iOS
//!   needs (long press, tap-vs-drag, pinch slop) are compiled only for
//!   android/ios, so their `corro_ios_canvas_gesture_*` exports have no macOS
//!   counterpart here — see `gui_backend::mobile_gesture`.
//! * **Pinch zoom is still wired**, because a trackpad pinch is a real Mac
//!   gesture and the shared zoom helper is the same one iOS uses.

use std::os::raw::{c_char, c_void};

// ---------------------------------------------------------------------------
// Bootstrap
// ---------------------------------------------------------------------------

/// Called from the app delegate once the window has a content view.
///
/// `root` is an `NSView*` (the window's `contentView`) and
/// `window_controller` an `NSWindowController*`. Boots the backend and the
/// spreadsheet; returns immediately (AppKit owns the run loop — the host calls
/// `-[NSApp run]`, which never returns, so there is no blocking `run()` to
/// come back from).
///
/// # Safety
/// Both pointers must be live Objective-C objects owned for the process
/// lifetime. Called on the main thread.
#[no_mangle]
pub unsafe extern "C" fn corro_macos_root_ready(root: *mut c_void, window_controller: *mut c_void) {
    // The canvas view class: an NSView subclass in the app that funnels
    // `drawRect:` and mouse/key events back into Rust. The backend falls back
    // to a plain NSView (tree builds, nothing draws) if it is missing.
    rswidgets::backends::macos::set_sheet_view_class("CorroSheetView");

    // A panic inside an `extern "C"` frame cannot unwind: Rust aborts, the
    // original message is replaced by "panic in a function that cannot
    // unwind", and everything about *what* went wrong is lost. The iOS host
    // crate hit exactly that and it cost several CI runs, so catch it at the
    // boundary and report the real panic.
    //
    // `catch_unwind` needs the closure to be UnwindSafe and the raw pointers
    // are plainly not, so they go through `AssertUnwindSafe` — justified
    // because this function owns nothing and never mutates through them.
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        corro::gui::macos_backend::macos_main(root, window_controller)
    }));
    match result {
        Ok(Ok(())) => {}
        Ok(Err(e)) => {
            corro::gui::macos_backend::log_macos(&format!("corro_macos_root_ready failed: {e}"));
        }
        Err(payload) => {
            let msg = panic_message(&payload);
            corro::gui::macos_backend::log_macos(&format!(
                "corro_macos_root_ready PANICKED: {msg}"
            ));
        }
    }
}

/// Run `f`, and if it panics report the *real* message before re-panicking.
///
/// Every `extern "C"` entry point here needs this, for the reason spelled out
/// on `corro_ios_canvas_draw` in the iOS host crate: a panic in a
/// non-unwinding frame aborts with "panic in a function that cannot unwind" and
/// the original message is lost, which is what happened to the draw callback —
/// the app died on its first real frame and the log could only show a location
/// inside the standard library.
///
/// Re-panicking after reporting is deliberate: swallowing the panic would leave
/// a half-drawn UI running, and the caller (an ObjC callback) has no way to
/// handle an error.
#[inline]
fn report_panics<R>(what: &str, f: impl FnOnce() -> R) -> R {
    match std::panic::catch_unwind(std::panic::AssertUnwindSafe(f)) {
        Ok(v) => v,
        Err(payload) => {
            let msg = panic_message(&payload);
            rswidgets::backends::macos::log_macos(&format!("PANIC in {what}: {msg}"));
            std::panic::resume_unwind(payload);
        }
    }
}

/// Best-effort `&str`/`String` extraction from a panic payload. The iOS host
/// crate spells this out at each call site; it is a function here so the
/// macOS exports do not carry four copies of the same `downcast` chain.
fn panic_message(payload: &Box<dyn std::any::Any + Send>) -> String {
    payload
        .downcast_ref::<&str>()
        .map(|s| (*s).to_owned())
        .or_else(|| payload.downcast_ref::<String>().cloned())
        .unwrap_or_else(|| "<non-string panic payload>".to_owned())
}

// ---------------------------------------------------------------------------
// Canvas (CorroSheetView)
// ---------------------------------------------------------------------------

/// Called from `CorroSheetView.drawRect:` with the live `CGContextRef`.
///
/// Replays the registered Rust draw closure for this canvas at the view's
/// size in points. `ctx` is the context AppKit hands to `drawRect:`, valid
/// only for that call.
///
/// # Safety
/// `ctx` must be a live `CGContextRef`; called on the main thread from
/// `drawRect:`.
#[no_mangle]
pub unsafe extern "C" fn corro_macos_canvas_draw(
    canvas_id: u64,
    ctx: *mut c_void,
    w: i32,
    h: i32,
) {
    // Wrapped: this is the call the app dies in if the closure panics, and the
    // unwrapped form leaves only a std frame in the log.
    report_panics("corro_macos_canvas_draw", || {
        // SAFETY: contract above; dispatch_draw runs the closure synchronously
        // and does not retain the context.
        unsafe { rswidgets::backends_macos_adapter::dispatch_draw(canvas_id, ctx, w, h) };
    });
}

/// Called from `CorroSheetView.layoutSubviews` (or `setFrameSize:`): tells the
/// backend the size the host has laid the canvas out to.
///
/// This has to be separate from the draw callback, for the same reason as on
/// iOS: AppKit only calls `drawRect:` when it decides to paint, and Rust
/// replays the draw closure earlier than that (registration time, and on every
/// `queue_redraw`), so without this the first replay used the 1x1 placeholder
/// from `set_size_request` and the sheet laid itself out as a single row.
#[no_mangle]
pub extern "C" fn corro_macos_canvas_size(canvas_id: u64, w: i32, h: i32) {
    rswidgets::backends_macos_adapter::record_canvas_size(canvas_id, w, h);
}

/// Called from `CorroSheetView` to route a press to the canvas that owns it
/// (`Canvas::on_click` was registered per canvas id). The sheet's handler
/// moves the cursor; the sheet-tab strip's selects/reorders a tab.
///
/// Coordinates are in points with a **top-left origin**, which is why the
/// generated canvas class overrides `isFlipped` to return `YES`
/// (`MACOS_GUIDELINES.md` §5 — the single most damaging silent failure in the
/// AppKit port).
#[no_mangle]
pub extern "C" fn corro_macos_canvas_click(canvas_id: u64, x: f64, y: f64) {
    rswidgets::backends_macos_adapter::dispatch_canvas_click(canvas_id, x, y);
}

/// Button- and modifier-aware press, for canvases registered through
/// `Canvas::on_click_button` (the sheet-tab context menu is the live user of
/// this: it needs to know the click was a *right* click).
///
/// `button` is AppKit's `NSMouseButton` (0 = left, 1 = right, 2 = centre) and
/// `mods` the `NSEventModifierFlags` bitmask. Returns nothing: the callback
/// decides what the press means.
#[no_mangle]
pub extern "C" fn corro_macos_canvas_click_button(
    canvas_id: u64,
    x: f64,
    y: f64,
    button: u32,
    mods: u32,
) {
    rswidgets::backends_macos_adapter::dispatch_canvas_click_button(canvas_id, x, y, button, mods);
}

/// Called from `CorroSheetView.mouseDragged:` so a drag can pan the viewport
/// (and, with a modifier, extend the selection — the shared rule decides
/// which).
#[no_mangle]
pub extern "C" fn corro_macos_canvas_motion(canvas_id: u64, x: f64, y: f64, mods: u32) {
    rswidgets::backends_macos_adapter::dispatch_canvas_motion(canvas_id, x, y, mods);
}

/// Called from `CorroSheetView.mouseUp:` — the end of a drag.
#[no_mangle]
pub extern "C" fn corro_macos_canvas_release(canvas_id: u64, x: f64, y: f64, button: u32, mods: u32) {
    rswidgets::backends_macos_adapter::dispatch_canvas_release(canvas_id, x, y, button, mods);
}

/// Called from `CorroSheetView.keyDown:` for hardware-keyboard input.
///
/// `keyval` uses the same convention the shared key handling expects (Unicode
/// scalar for printable keys, the `rswidgets::core::key` constants otherwise);
/// `mods` is the modifier bitmask (1 = Shift, 4 = Ctrl, 8 = Alt). Returns
/// `true` when the key was consumed, so AppKit skips its own handling.
#[no_mangle]
pub extern "C" fn corro_macos_canvas_key(canvas_id: u64, keyval: u32, mods: u32) -> bool {
    rswidgets::backends_macos_adapter::dispatch_canvas_key(canvas_id, keyval, mods)
}

/// The grid's row height and default column advance, in points, so a
/// scroll-wheel or trackpad pan can convert pixels to whole cells with the
/// same metrics the renderer uses. Writes `[row_h, col_w]` into `out` (which
/// must have room for two `f64`s).
///
/// # Safety
/// `out` must point to at least two writable `f64`s.
#[no_mangle]
pub unsafe extern "C" fn corro_macos_canvas_cell_size(out: *mut f64) {
    let (row_h, col_w) = corro::gui::macos_backend::cell_size();
    if out.is_null() {
        return;
    }
    // SAFETY: the contract above says `out` has room for two f64s.
    unsafe {
        *out = row_h;
        *out.add(1) = col_w;
    }
}

/// Scroll the sheet by whole cells, as a scroll wheel or trackpad two-finger
/// pan ends in. Writes the applied `[dRows, dCols]` into `out` (which must
/// have room for two `i32`s).
///
/// The applied delta can be smaller than the wheel asked for at a sheet edge,
/// which is how the host learns to come to rest.
///
/// # Safety
/// `out` must point to at least two writable `i32`s.
#[no_mangle]
pub unsafe extern "C" fn corro_macos_canvas_scroll_by(d_rows: i32, d_cols: i32, out: *mut i32) {
    let (dr, dc) = corro::gui::macos_backend::scroll_by_cells(d_rows, d_cols);
    if out.is_null() {
        return;
    }
    // SAFETY: the contract above says `out` has room for two i32s.
    unsafe {
        *out = dr;
        *out.add(1) = dc;
    }
}

/// Zoom the sheet by `factor` (a trackpad pinch's scale ratio since the
/// previous callback), returning the scale actually applied (clamped to the
/// shared `0.4..=4.0` range).
#[no_mangle]
pub extern "C" fn corro_macos_canvas_zoom(factor: f64) -> f64 {
    corro::gui::macos_backend::zoom_viewport_by(factor)
}

/// Reset the pinch scale to 1.0. Returns the applied scale (always 1.0).
#[no_mangle]
pub extern "C" fn corro_macos_canvas_zoom_reset() -> f64 {
    corro::gui::macos_backend::reset_viewport_zoom()
}

// ---------------------------------------------------------------------------
// Entry (formula bar)
// ---------------------------------------------------------------------------

/// Called from the text field's delegate on `controlTextDidChange:`: runs
/// corro's formula entry change handler, which syncs `edit_buf` from the
/// widget text.
///
/// This is the only signal typing produces (a hardware key press is handled
/// separately by `corro_macos_canvas_key`), which is why the shared handler
/// adopts a change as a fresh edit on the Apple backends.
#[no_mangle]
pub extern "C" fn corro_macos_entry_changed(view_ptr: u64) {
    rswidgets::backends_macos_adapter::dispatch_text_changed(view_ptr as usize as *mut c_void);
}

/// Called from the delegate on Return (`textFieldShouldReturn:`): commits the
/// edit and moves down, exactly like a hardware Return.
#[no_mangle]
pub extern "C" fn corro_macos_entry_activate(view_ptr: u64) {
    rswidgets::backends_macos_adapter::dispatch_entry_activate(view_ptr as usize as *mut c_void);
}

/// Called from the delegate on `controlTextDidBeginEditing:` /
/// `controlTextDidEndEditing:`. Returns the handler's own 0/1 answer (whether
/// it consumed the event) so the ObjC side can decide whether to continue into
/// AppKit's default behaviour — the Rust side returns 0 when no handler is
/// registered.
#[no_mangle]
pub extern "C" fn corro_macos_entry_focus(view_ptr: u64, gained: bool) -> i32 {
    rswidgets::backends_macos_adapter::dispatch_focus(view_ptr as usize as *mut c_void, gained)
}

// ---------------------------------------------------------------------------
// Generic control callback (buttons, switches, pickers)
// ---------------------------------------------------------------------------

/// Called from `CorroMacTarget.corroFired:` with the id returned by the
/// Rust-side `Button::on_click` (or any other registered callback).
#[no_mangle]
pub extern "C" fn corro_macos_callback(callback_id: u64) {
    rswidgets::backends::macos::dispatch_callback(callback_id);
}

// ---------------------------------------------------------------------------
// Menu
// ---------------------------------------------------------------------------

/// Dispatch a menu item chosen in the host's `NSMenu`.
///
/// `name` is a NUL-terminated `app.<action>` string — the same name the
/// desktop menu registers, so there is one dispatch table. Unknown names are
/// ignored (the menu may be built from a newer tree than the running handler).
///
/// # Safety
/// `name` must be a valid NUL-terminated C string.
#[no_mangle]
pub unsafe extern "C" fn corro_macos_menu_action(name: *const c_char) {
    if name.is_null() {
        return;
    }
    // SAFETY: contract above; we copy immediately and never retain the
    // pointer, so the caller's buffer only has to be valid for this call.
    let bytes = std::ffi::CStr::from_ptr(name).to_bytes();
    let name = String::from_utf8_lossy(bytes);
    corro::gui::macos_backend::run_menu_action_by_name(&name);
}

/// Number of top-level menus in the model the host renders. A macOS host
/// builds a genuine `NSMenu` from it (unlike iOS, which has no menubar).
#[no_mangle]
pub extern "C" fn corro_macos_menu_count() -> usize {
    corro::gui::macos_backend::menu_model().len()
}

/// Number of items in menu `menu_idx` (0 when out of range).
#[no_mangle]
pub extern "C" fn corro_macos_menu_item_count(menu_idx: usize) -> usize {
    corro::gui::macos_backend::menu_model()
        .get(menu_idx)
        .map(|(_, items)| items.len())
        .unwrap_or(0)
}

thread_local! {
    /// Scratch storage for the C strings handed to the host. One slot per
    /// pointer so both out-params stay valid until the next call, which is
    /// what the documented contract promises. Main-thread only, like every
    /// other backend entry point.
    static MENU_SCRATCH: std::cell::RefCell<(Vec<u8>, Vec<u8>)> =
        const { std::cell::RefCell::new((Vec::new(), Vec::new())) };
}

/// Fetch one menu item's label and action name.
///
/// # Safety
/// `label_out` and `action_out` must be valid, writable `const char**`.
#[no_mangle]
pub unsafe extern "C" fn corro_macos_menu_item(
    menu_idx: usize,
    item_idx: usize,
    label_out: *mut *const c_char,
    action_out: *mut *const c_char,
) -> bool {
    if label_out.is_null() || action_out.is_null() {
        return false;
    }
    let model = corro::gui::macos_backend::menu_model();
    let Some((_, items)) = model.get(menu_idx) else {
        return false;
    };
    let Some((label, action)) = items.get(item_idx) else {
        return false;
    };
    MENU_SCRATCH.with(|slot| {
        let mut slot = slot.borrow_mut();
        let (lbuf, abuf) = &mut *slot;
        lbuf.clear();
        lbuf.extend_from_slice(label.as_bytes());
        lbuf.push(0);
        abuf.clear();
        abuf.extend_from_slice(action.as_bytes());
        abuf.push(0);
        // SAFETY: contract above says both out-params are writable. The
        // pointers stay valid until the next call (documented), because the
        // buffers live in the thread-local, not on a dropped stack frame.
        unsafe {
            *label_out = lbuf.as_ptr() as *const c_char;
            *action_out = abuf.as_ptr() as *const c_char;
        }
    });
    true
}

// ---------------------------------------------------------------------------
// Host contract tests
// ---------------------------------------------------------------------------
//
// These are the assertions that do not need a Mac: the C entry points exist
// with the signatures `app/CorroMacBridge.h` declares, and the menu model the
// host walks is dispatchable. The iOS host crate has the equivalent file
// (`tests/ios_pipeline_preview.rs`, run on Linux CI); the checks that genuinely
// need AppKit — the draw replay against a live `CGContextRef`, `isFlipped`,
// text measurement — are covered by the cfg compile checks
// (`scripts/check_rswidgets_macos.sh`, `scripts/check_corro_macos.sh`) and, at
// runtime, by the app on a real Mac.
//
// A drift between this file and the bridge header is the exact failure the
// generator exists to prevent, and it is a *silent* one: a missing symbol is
// a link error, but a changed signature is a runtime crash inside
// `objc_msgSend`. So the names are asserted here, on every CI run, and the
// header is checked against this list by `macos/corro/scripts/check_bridge.sh`.

#[cfg(test)]
mod tests {
    use super::*;

    /// Every `corro_macos_*` name the bridge header declares, with the shape
    /// the header expects. The list is the *header's* contract: if a selector
    /// in `MACOS_GUIDELINES.md` §3 has no export here, a host cannot call it.
    #[test]
    fn the_host_contract_exports_all_exist() {
        // Bootstrap: a null root must be *reported*, not accepted, and never
        // panic — a panic inside an `extern "C"` frame aborts the process with
        // no unwind and no symbolised frames.
        let r = corro_macos_root_ready(std::ptr::null_mut(), std::ptr::null_mut());
        // The return value is (), so the only assertion available without a
        // Mac is that the call did not abort — reaching this line is the
        // test. (The iOS twin asserts on `ios_main`'s Result; here the
        // boundary swallows the error and logs it, deliberately, because the
        // ObjC caller has no way to handle one.)
        let _ = r;
    }

    /// The menu model the host walks must be dispatchable, and the same
    /// `app.<name>` names the desktop menu registers.
    #[test]
    fn the_menu_model_is_dispatchable() {
        let count = corro_macos_menu_count();
        assert!(count > 0, "the host would build an empty NSMenu");

        let mut total = 0usize;
        for m in 0..count {
            let items = corro_macos_menu_item_count(m);
            for i in 0..items {
                let mut label: *const c_char = std::ptr::null();
                let mut action: *const c_char = std::ptr::null();
                let ok = unsafe { corro_macos_menu_item(m, i, &mut label, &mut action) };
                assert!(ok, "menu {m} item {i} reported out of range");
                assert!(!label.is_null(), "menu {m} item {i} has no label");
                assert!(!action.is_null(), "menu {m} item {i} has no action");
                // SAFETY: both out-params were written by the call above and
                // are NUL-terminated, valid until the next call on this
                // thread — the documented contract, and this is the next use
                // rather than another one.
                let a = unsafe { std::ffi::CStr::from_ptr(action) }
                    .to_string_lossy()
                    .into_owned();
                assert!(
                    a.starts_with("app."),
                    "menu {m} item {i} publishes '{a}', which corro_macos_menu_action cannot dispatch"
                );
                total += 1;
            }
        }
        assert!(total > 0, "the menu model has no items at all");
    }

    /// Out-of-range indices must be rejected, not a panic: a host walking the
    /// model with a stale index is a benign mistake, and the C contract says
    /// "return 0/false".
    #[test]
    fn out_of_range_menu_indices_are_rejected() {
        assert_eq!(corro_macos_menu_item_count(usize::MAX), 0);
        let mut label: *const c_char = std::ptr::null();
        let mut action: *const c_char = std::ptr::null();
        assert!(!unsafe { corro_macos_menu_item(usize::MAX, 0, &mut label, &mut action) });
        assert!(!unsafe { corro_macos_menu_item(0, usize::MAX, &mut label, &mut action) });
        // And null out-params must not be written through.
        assert!(!unsafe {
            corro_macos_menu_item(0, 0, std::ptr::null_mut(), &mut action)
        });
    }

    /// Dispatching a real name and an unknown one must both be safe. With no
    /// GUI state published (there is none in a unit-test process) the dispatch
    /// logs "mobile menu action before state published" and returns — so this
    /// asserts the *lookup* path, which is all that is host-independent.
    #[test]
    fn menu_dispatch_accepts_known_and_unknown_names() {
        unsafe {
            corro_macos_menu_action(std::ffi::CString::new("app.about").unwrap().as_ptr());
            corro_macos_menu_action(
                std::ffi::CString::new("app.definitely_not_a_real_action")
                    .unwrap()
                    .as_ptr(),
            );
            // A null name must be ignored, not dereferenced.
            corro_macos_menu_action(std::ptr::null());
        }
    }

    /// The cell-size out-param helper: with no live viewport the numbers are
    /// the defaults, and they must be positive or a host's pixel-to-cell
    /// division breaks.
    #[test]
    fn cell_size_is_positive_and_null_safe() {
        let mut out = [0.0f64; 2];
        unsafe { corro_macos_canvas_cell_size(out.as_mut_ptr()) };
        assert!(out[0] > 0.0, "row height must be positive, got {}", out[0]);
        assert!(out[1] > 0.0, "column advance must be positive, got {}", out[1]);
        unsafe { corro_macos_canvas_cell_size(std::ptr::null_mut()) };
    }
}
