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

/// Called from `SheetView.onTouchEvent` (ACTION_MOVE) during a drag: scrolls
/// the sheet by whole rows/columns.
///
/// Android has no native scrolling for the sheet, so the Java side accumulates
/// the drag in pixels and converts it to cell counts using
/// `nativeCellSize()`; this entry point applies the result. `d_rows > 0`
/// scrolls towards later rows (content moves up under the finger), matching
/// the natural drag direction.
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeScrollBy(
    _env: JNIEnv,
    _class: JClass,
    d_rows: i32,
    d_cols: i32,
) {
    corro::gui::android_backend::scroll_viewport(d_rows, d_cols);
}

/// Called from `SheetView` to learn the grid's row height and default column
/// width in device pixels, so a pixel drag converts to whole cells with the
/// same metrics Rust renders with. Returns `[row_h, col_w]`.
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeCellSize(
    env: JNIEnv,
    _class: JClass,
) -> jni::sys::jfloatArray {
    let (row_h, col_w) = corro::gui::android_backend::touch_cell_size();
    match env.new_float_array(2) {
        Ok(arr) => {
            let _ = env.set_float_array_region(&arr, 0, &[row_h as f32, col_w as f32]);
            arr.into_raw()
        }
        Err(_) => std::ptr::null_mut(),
    }
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

/// Called from `CorroKeyListener.onKey` (hardware/adb Enter): same commit
/// path as the IME action above.
#[no_mangle]
pub extern "system" fn Java_com_corro_CorroKeyListener_nativeEntryActivate(
    _env: JNIEnv,
    _class: JClass,
    view_ptr: i64,
) {
    rswidgets::backends_android_adapter::dispatch_entry_activate(view_ptr as usize as *mut _);
}

/// Called from `MenuStrip.nativeMenuAction(action)`: the strip's buttons
/// (quick actions and every overflow item) dispatch by the same `app.*`
/// action name the desktop menu registers, so there is one dispatch table.
#[no_mangle]
pub extern "system" fn Java_com_corro_MenuStrip_nativeMenuAction(
    mut env: JNIEnv,
    _class: JClass,
    action: JObject,
) {
    let jstr: &jni::objects::JString = (&action).into();
    let Ok(action) = env.get_string(jstr) else {
        return;
    };
    let action: String = action.into();
    corro::gui::android_backend::run_menu_action_by_name(&action);
}
