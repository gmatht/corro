//! corro Android cdylib: the single JNI export the Activity calls.
//!
//! Everything else lives in the `corro` library
//! (`corro::gui::android_backend`): backend init + the shared `run_gui`
//! spreadsheet pipeline. Keeping the export here (crate root of a cdylib)
//! guarantees the linker retains it.

use jni::objects::{JClass, JObject};
use jni::JNIEnv;

/// Called from `MainActivity.nativeInit(activity, rootLayout)` once the
/// Activity has a content view. Boots the backend and the spreadsheet;
/// returns immediately (Android drives the event loop).
#[no_mangle]
pub extern "system" fn Java_com_corro_MainActivity_nativeInit(
    mut env: JNIEnv,
    _class: JClass,
    activity: JObject,
    root_layout: JObject,
) {
    // Custom sheet view for canvases (grid rendering via onDraw -> Rust).
    rswidgets::backends::android::set_sheet_view_class("com.corro.SheetView");
    if let Err(e) = corro::gui::android_backend::android_main(&mut env, &activity, &root_layout) {
        let _ = env.throw_new("java/lang/RuntimeException", e);
    }
}

/// Called from `SheetView.onDraw`: replays the registered Rust draw closure
/// against the live `android.graphics.Canvas` at the view's pixel size.
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeOnDraw(
    mut env: JNIEnv,
    _class: JClass,
    canvas_id: i64,
    canvas: JObject,
    w: i32,
    h: i32,
) {
    rswidgets::backends_android_adapter::dispatch_draw(canvas_id as u64, canvas, &mut env, w, h);
}

/// Called from `SheetView.onTouchEvent` (ACTION_DOWN): moves the cursor to
/// the tapped cell and redraws.
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeOnTouch(
    _env: JNIEnv,
    _class: JClass,
    canvas_id: i64,
    x: f32,
    y: f32,
) {
    rswidgets::backends_android_adapter::dispatch_canvas_click(canvas_id as u64, x as f64, y as f64);
    // The click handler queues a redraw; make sure onDraw fires.
    // (queue_redraw already invalidates; this is belt-and-braces for the
    // tab strip, whose canvas has no Java view of its own... no-op here.)
}

/// Called from `CorroTextWatcher.afterTextChanged`: runs corro's formula
/// entry change handler, which syncs `edit_buf` from the widget text.
#[no_mangle]
pub extern "system" fn Java_com_corro_CorroTextWatcher_nativeEntryChanged(
    _env: JNIEnv,
    _class: JClass,
    view_ptr: i64,
) {
    rswidgets::backends_android_adapter::dispatch_text_changed(view_ptr as usize as *mut _);
}

/// Called from `CorroEditorAction.onEditorAction` (IME Done/Enter): commits
/// the edit and moves down, exactly like a hardware Return.
#[no_mangle]
pub extern "system" fn Java_com_corro_CorroEditorAction_nativeEntryActivate(
    _env: JNIEnv,
    _class: JClass,
    view_ptr: i64,
) {
    rswidgets::backends_android_adapter::dispatch_entry_activate(view_ptr as usize as *mut _);
}
