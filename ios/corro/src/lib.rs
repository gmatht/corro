//! corro iOS cdylib: the `extern "C"` entry points the app's objc/Swift
//! shims call.
//!
//! Everything else lives in the `corro` library
//! (`corro::gui::ios_backend`): the shared `run_gui` spreadsheet pipeline plus
//! the menu model. Keeping the exports here (crate root of a staticlib/cdylib)
//! guarantees the linker retains them — the same reason
//! `android/corro/src/lib.rs` holds the JNI exports.
//!
//! The names are the contract between this crate and `app/`:
//! `rustxWidgets/docs/IOS_GUIDELINES.md` §3 is the table; each export below
//! repeats the selector it belongs to. Because they are plain C functions,
//! they are visible to any Objective-C file in the app target with a matching
//! declaration — see `app/CorroBridge.h`.

use std::os::raw::{c_char, c_void};

// ---------------------------------------------------------------------------
// Bootstrap
// ---------------------------------------------------------------------------

/// Called from the scene/view-controller once it has a root view.
///
/// `root` is a `UIView*`, `view_controller` a `UIViewController*`. Boots the
/// backend and the spreadsheet; returns immediately (UIKit drives the event
/// loop — there is no blocking `run()` to come back from).
///
/// # Safety
/// Both pointers must be live Objective-C objects owned for the process
/// lifetime. Called on the main thread.
#[no_mangle]
pub unsafe extern "C" fn corro_ios_root_ready(root: *mut c_void, view_controller: *mut c_void) {
    // The canvas view class: a UIView subclass in this app that funnels
    // `drawRect:` and touches back into Rust. The backend falls back to a
    // plain UIView (tree builds, nothing draws) if it is missing.
    rswidgets::backends::ios::set_sheet_view_class("SheetView");

    // A panic inside an `extern "C"` frame cannot unwind: Rust aborts, the
    // original message is replaced by "panic in a function that cannot
    // unwind", and everything about *what* went wrong is lost. That cost
    // several CI runs while bringing the app up, so catch it at the boundary
    // and report the real panic instead. `catch_unwind` needs the closure to
    // be UnwindSafe, and the raw pointers are plainly not, so they are passed
    // through `AssertUnwindSafe` with the justification that this function
    // owns nothing and never mutates through them.
    // SAFETY: this function's contract (both pointers live for the process
    // lifetime, called on the main thread). The whole `catch_unwind` is inside
    // the unsafe block for that reason.
    let result = unsafe {
        std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            corro::gui::ios_backend::ios_main(root, view_controller)
        }))
    };
    match result {
        Ok(Ok(())) => {}
        Ok(Err(e)) => {
            corro::gui::ios_backend::log_ios(&format!("corro_ios_root_ready failed: {e}"));
        }
        Err(payload) => {
            let msg = payload
                .downcast_ref::<&str>()
                .map(|s| (*s).to_owned())
                .or_else(|| payload.downcast_ref::<String>().cloned())
                .unwrap_or_else(|| "<non-string panic payload>".to_owned());
            corro::gui::ios_backend::log_ios(&format!("corro_ios_root_ready PANICKED: {msg}"));
        }
    }
}


/// Run `f`, and if it panics report the *real* message before re-panicking.
///
/// Every `extern "C"` entry point here needs this. A panic inside a
/// non-unwinding frame aborts with "panic in a function that cannot unwind" and
/// the original message is lost - which is exactly what happened to the draw
/// callback: the app died on its first real frame and the only thing the CI log
/// could show was `panicking.rs:225`, a location inside the standard library.
/// The frames that would have named the closure were in the log; the *message*
/// was not. This recovers it.
///
/// Re-panicking after reporting is deliberate: swallowing the panic would leave
/// a half-drawn UI running, and the caller (an ObjC callback) has no way to
/// handle an error.
#[inline]
fn report_panics<R>(what: &str, f: impl FnOnce() -> R) -> R {
    match std::panic::catch_unwind(std::panic::AssertUnwindSafe(f)) {
        Ok(v) => v,
        Err(payload) => {
            let msg = payload
                .downcast_ref::<&str>()
                .map(|s| (*s).to_owned())
                .or_else(|| payload.downcast_ref::<String>().cloned())
                .unwrap_or_else(|| "<non-string panic payload>".to_owned());
            rswidgets::backends::ios::log_ios(&format!("PANIC in {what}: {msg}"));
            std::panic::resume_unwind(payload);
        }
    }
}

// ---------------------------------------------------------------------------
// Canvas (SheetView)
// ---------------------------------------------------------------------------

/// Called from `SheetView.drawRect:` with the live `CGContextRef`.
///
/// Replays the registered Rust draw closure for this canvas at the view's
/// size in points. `ctx` is the context UIKit hands to `drawRect:`, valid only
/// for that call.
///
/// # Safety
/// `ctx` must be a live `CGContextRef`; called on the main thread from
/// `drawRect:`.
#[no_mangle]
pub unsafe extern "C" fn corro_ios_canvas_draw(
    canvas_id: u64,
    ctx: *mut c_void,
    w: i32,
    h: i32,
) {
    // Wrapped: this is where the app died (see `report_panics`), and the
    // unwrapped form left only `panicking.rs:225` in the log.
    report_panics("corro_ios_canvas_draw", || {
        // SAFETY: contract above; dispatch_draw runs the closure synchronously
        // and does not retain the context.
        unsafe { rswidgets::backends_ios_adapter::dispatch_draw(canvas_id, ctx, w, h) };
    });
}

/// Called from `SheetView.layoutSubviews`: tells the backend the size the host
/// has laid the canvas out to.
///
/// This has to be separate from the draw callback. UIKit only calls
/// `drawRect:` when it decides to paint, and Rust replays the draw closure
/// earlier than that (registration time, and on every `queue_redraw`), so
/// without this the first replay used the 1x1 placeholder from
/// `set_size_request` and the sheet laid itself out as a single row — visible
/// in the simulator log as `DRAW_CALLBACK called: w=1 h=1`.
#[no_mangle]
pub extern "C" fn corro_ios_canvas_size(canvas_id: u64, w: i32, h: i32) {
    rswidgets::backends_ios_adapter::record_canvas_size(canvas_id, w, h);
}

/// Called from `SheetView` to route a tap to the canvas that owns it
/// (`Canvas::on_click` was registered per canvas id). The sheet's tap handler
/// moves the cursor; the sheet-tab strip's selects/reorders a tab.
#[no_mangle]
pub extern "C" fn corro_ios_canvas_click(canvas_id: u64, x: f64, y: f64) {
    rswidgets::backends_ios_adapter::dispatch_canvas_click(canvas_id, x, y);
}

// ---------------------------------------------------------------------------
// Pinch to zoom
// ---------------------------------------------------------------------------

/// Called from `SheetView`'s `UIPinchGestureRecognizer`: multiply the sheet's
/// view scale by `factor` — the recogniser's `scale` ratio since the previous
/// callback.
///
/// UIKit has no sheet-level zoom, so the host forwards the ratio and the shared
/// Rust zoom applies it. Everything the grid draws *and* everything a finger
/// hits derives from that one scale, so a pinch moves the grid, the gutter, the
/// headers and the hit targets together.
///
/// Returns the scale actually applied (clamped to the shared
/// `0.4..=4.0` range), so the host can mirror it and log it.
#[no_mangle]
pub extern "C" fn corro_ios_canvas_zoom(factor: f64) -> f64 {
    corro::gui::ios_backend::zoom_viewport(factor)
}

/// Called from `SheetView` on a double-tap: reset the pinch scale to 1.0.
/// Returns the applied scale (always 1.0).
#[no_mangle]
pub extern "C" fn corro_ios_canvas_zoom_reset() -> f64 {
    corro::gui::ios_backend::reset_viewport_zoom()
}

// ---------------------------------------------------------------------------
// Touch gestures (drag pans; long press then drag selects)
// ---------------------------------------------------------------------------
//
// Every call carries the canvas id: the sheet and the sheet-tab strip are both
// canvases, and Rust ignores gestures for any canvas that is not the sheet, so
// the strip keeps its own click/reorder handling.

/// Called from `SheetView` on `touchesBegan:`. `is_touch` is 1 for a finger
/// (drag pans; long press selects) and 0 for a mouse/trackpad (drag selects,
/// as on a desktop).
#[no_mangle]
pub extern "C" fn corro_ios_canvas_gesture_down(canvas_id: u64, x: f64, y: f64, is_touch: i32) {
    let _ = corro::gui::ios_backend::gesture_down(canvas_id, x, y, is_touch != 0);
}

/// Called from `SheetView`'s `UILongPressGestureRecognizer`: arm selection so
/// the drag that follows extends the selection instead of scrolling.
#[no_mangle]
pub extern "C" fn corro_ios_canvas_gesture_long_press(canvas_id: u64, x: f64, y: f64) {
    let _ = corro::gui::ios_backend::gesture_long_press(canvas_id, x, y);
}

/// Called from `SheetView` on `touchesMoved:`. Returns the shared outcome code
/// (0 ignored, 1 select, 2 scroll, 3 tap, 4 long press) so the host knows
/// whether to pan.
#[no_mangle]
pub extern "C" fn corro_ios_canvas_gesture_move(canvas_id: u64, x: f64, y: f64) -> i32 {
    corro::gui::ios_backend::gesture_move(canvas_id, x, y)
}

/// Called from `SheetView` on `touchesEnded:`.
///
/// A release that never dragged is reported as a *tap* (`3`), not applied: the
/// shim then routes it to the canvas's own click handler
/// (`corro_ios_canvas_click`), which is per canvas — so the tab strip's taps do
/// not end up moving the sheet's cursor.
#[no_mangle]
pub extern "C" fn corro_ios_canvas_gesture_up(canvas_id: u64, x: f64, y: f64) -> i32 {
    corro::gui::ios_backend::gesture_up(canvas_id, x, y)
}

/// Called from `SheetView` on `touchesCancelled:`. Never produces a tap.
#[no_mangle]
pub extern "C" fn corro_ios_canvas_gesture_cancel(canvas_id: u64) {
    corro::gui::ios_backend::gesture_cancel(canvas_id);
}

/// Called from `SheetView` for a scroll drag: pan the sheet by the pixel delta
/// since the last event, converting to whole cells with the zoom-aware metrics.
/// Writes the applied `[dRows, dCols]` into `out` (which must have room for
/// two `i32`s), so the host can keep its sub-cell remainder.
///
/// # Safety
/// `out` must point to at least two writable `i32`s.
#[no_mangle]
pub unsafe extern "C" fn corro_ios_canvas_drag_by(
    canvas_id: u64,
    dx: f64,
    dy: f64,
    out: *mut i32,
) {
    let (d_rows, d_cols) = corro::gui::ios_backend::drag_viewport(canvas_id, dx, dy);
    if out.is_null() {
        return;
    }
    // SAFETY: the contract above says `out` has room for two i32s.
    unsafe {
        *out = d_rows;
        *out.add(1) = d_cols;
    }
}

/// The grid's row height and default column advance, in points, so the host's
/// drag accumulation uses the same metrics the renderer does. Writes
/// `[row_h, col_w]` into `out` (which must have room for two `f64`s).
///
/// # Safety
/// `out` must point to at least two writable `f64`s.
#[no_mangle]
pub unsafe extern "C" fn corro_ios_canvas_cell_size(out: *mut f64) {
    let (row_h, col_w) = corro::gui::ios_backend::touch_cell_size();
    if out.is_null() {
        return;
    }
    // SAFETY: the contract above says `out` has room for two f64s.
    unsafe {
        *out = row_h;
        *out.add(1) = col_w;
    }
}

/// Called from `SheetView` for hardware-keyboard input (a physical keyboard
/// on iPad, `simctl` input, the iOS 13+ keyboard accessory bar).
///
/// `keyval` uses the same convention the shared key handling expects (Unicode
/// scalar for printable keys, the `rswidgets::core::key` constants otherwise);
/// `mods` is the modifier bitmask (1 = Shift, 4 = Ctrl, 8 = Alt).
#[no_mangle]
pub extern "C" fn corro_ios_canvas_key(canvas_id: u64, keyval: u32, mods: u32) -> bool {
    rswidgets::backends_ios_adapter::dispatch_canvas_key(canvas_id, keyval, mods)
}

// ---------------------------------------------------------------------------
// Entry (formula bar)
// ---------------------------------------------------------------------------

/// Called from the text field's delegate on `editingChanged`: runs corro's
/// formula entry change handler, which syncs `edit_buf` from the widget text.
///
/// This is the *only* signal soft-keyboard typing produces — there is no key
/// event for it — which is why the shared handler adopts a change as a fresh
/// edit on iOS (see `gui_backend::on_formula_entry_changed`).
#[no_mangle]
pub extern "C" fn corro_ios_entry_changed(view_ptr: u64) {
    rswidgets::backends_ios_adapter::dispatch_text_changed(view_ptr as usize as *mut c_void);
}

/// Called from the delegate on IME Done/Return (`textFieldShouldReturn:`):
/// commits the edit and moves down, exactly like a hardware Return.
#[no_mangle]
pub extern "C" fn corro_ios_entry_activate(view_ptr: u64) {
    rswidgets::backends_ios_adapter::dispatch_entry_activate(view_ptr as usize as *mut c_void);
}

/// Called from the delegate on `textFieldDidBeginEditing:` /
/// `textFieldDidEndEditing:`. Returns the handler's own 0/1 answer (whether
/// it consumed the event) so the ObjC side can decide whether to continue
/// into UIKit's default behaviour — the Rust side returns 0 when no handler
/// is registered.
#[no_mangle]
pub extern "C" fn corro_ios_entry_focus(view_ptr: u64, gained: bool) -> i32 {
    rswidgets::backends_ios_adapter::dispatch_focus(
        view_ptr as usize as *mut c_void,
        gained,
    )
}

// ---------------------------------------------------------------------------
// Generic control callback (buttons, switches, pickers)
// ---------------------------------------------------------------------------

/// Called from `CorroIosTarget.corroFired:` with the id returned by the
/// Rust-side `Button::on_click` (or any other registered callback).
#[no_mangle]
pub extern "C" fn corro_ios_callback(callback_id: u64) {
    rswidgets::backends::ios::dispatch_callback(callback_id);
}

// ---------------------------------------------------------------------------
// Menu
// ---------------------------------------------------------------------------

/// Dispatch a menu item chosen in the host's `UIMenu`/action sheet.
///
/// `name` is a NUL-terminated `app.<action>` string — the same name the
/// desktop menu registers, so there is one dispatch table. Unknown names are
/// ignored (the menu may be built from a newer tree than the running handler).
///
/// # Safety
/// `name` must be a valid NUL-terminated C string.
#[no_mangle]
pub unsafe extern "C" fn corro_ios_menu_action(name: *const c_char) {
    if name.is_null() {
        return;
    }
    // SAFETY: contract above; we copy immediately and never retain the
    // pointer, so the caller's buffer only has to be valid for this call.
    let bytes = unsafe { std::ffi::CStr::from_ptr(name) }.to_bytes();
    let name = String::from_utf8_lossy(bytes);
    corro::gui::ios_backend::run_menu_action_by_name(&name);
}

/// Read the menu model the host renders, one item at a time.
///
/// The host walks the model with this pair rather than receiving a struct:
/// `corro_ios_menu_count()` then, for each index,
/// `corro_ios_menu_item(menu_idx, item_idx, &label, &action)`. Both out
/// parameters are `const char*` valid until the next call (the strings are
/// held in a Rust-side scratch buffer), which is enough for the shim to build
/// `UIAction`s immediately.
///
/// Returns 0 for an out-of-range index, 1 otherwise.
#[no_mangle]
pub extern "C" fn corro_ios_menu_count() -> usize {
    corro::gui::ios_backend::menu_model().len()
}

/// Number of items in menu `menu_idx` (0 when out of range).
#[no_mangle]
pub extern "C" fn corro_ios_menu_item_count(menu_idx: usize) -> usize {
    corro::gui::ios_backend::menu_model()
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
pub unsafe extern "C" fn corro_ios_menu_item(
    menu_idx: usize,
    item_idx: usize,
    label_out: *mut *const c_char,
    action_out: *mut *const c_char,
) -> bool {
    if label_out.is_null() || action_out.is_null() {
        return false;
    }
    let model = corro::gui::ios_backend::menu_model();
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
