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
    // SAFETY: contract above; dispatch_draw runs the closure synchronously
    // and does not retain the context.
    unsafe { rswidgets::backends_ios_adapter::dispatch_draw(canvas_id, ctx, w, h) };
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

/// Called from `SheetView.touchesEnded:` (or its tap recogniser): moves the
/// cursor to the tapped cell and redraws.
#[no_mangle]
pub extern "C" fn corro_ios_canvas_click(canvas_id: u64, x: f64, y: f64) {
    rswidgets::backends_ios_adapter::dispatch_canvas_click(canvas_id, x, y);
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
