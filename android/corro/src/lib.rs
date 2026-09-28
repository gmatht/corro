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

/// Called from `SheetView` to route a tap to the canvas that owns it
/// (`Canvas::on_click` was registered per canvas id). The sheet's tap handler
/// moves the cursor; the sheet-tab strip's selects/reorders a tab.
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeOnTouch(
    _env: JNIEnv,
    _class: JClass,
    canvas_id: i64,
    x: f32,
    y: f32,
) {
    rswidgets::backends_android_adapter::dispatch_canvas_click(canvas_id as u64, x as f64, y as f64);
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

/// Called from `SheetView.onScale` (a pinch): multiply the sheet's view scale
/// by `factor`, the gesture's span ratio since the previous event.
///
/// Returns the scale actually applied after clamping, so the Java side can
/// refresh its cached cell size (the metrics now include the zoom) and show
/// the value. Pinch works on the sheet, not the Activity: everything the grid
/// draws and everything a tap hits derives from the same scale.
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeZoom(
    _env: JNIEnv,
    _class: JClass,
    factor: f32,
) -> f32 {
    corro::gui::android_backend::zoom_viewport(factor as f64) as f32
}

/// Called from `SheetView` to reset the pinch scale to 1.0 (a double-tap).
/// Returns the applied scale (always 1.0).
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeResetZoom(
    _env: JNIEnv,
    _class: JClass,
) -> f32 {
    corro::gui::android_backend::reset_viewport_zoom() as f32
}

/// Called from `SheetView.onTouchEvent(ACTION_DOWN)`: begins a pointer
/// gesture on the sheet canvas `canvas_id`. `is_touch` is 1 for a finger and
/// 0 for a mouse (the emulator's pointer or a stylus), which is what decides
/// whether a drag selects (mouse) or scrolls (finger).
///
/// The canvas id matters: the sheet and the sheet-tab strip are both
/// `SheetView`s, and Rust ignores gestures for any canvas that is not the
/// sheet, so the strip keeps its own click/reorder handling.
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeGestureDown(
    _env: JNIEnv,
    _class: JClass,
    canvas_id: i64,
    x: f32,
    y: f32,
    is_touch: jni::sys::jboolean,
) {
    let _ = corro::gui::android_backend::gesture_down(canvas_id as u64, x as f64, y as f64, is_touch != 0);
}

/// Called from `SheetView` on `GestureDetector.onLongPress`.
///
/// Returns one of the shared `DragOutcome`-shaped codes: `4` (the long press
/// armed a range selection, and the drag that follows extends it) or `5`
/// (open the cell context menu — a phone's substitute for a right-click).
///
/// The host does not decide which; it asks, and acts on the answer, because
/// the decision needs the gesture state only Rust has (whether the finger
/// moved past the drag slop, whether the press was on the sheet at all).
/// Before this existed the long press could only arm a selection, so the
/// sheet had no route to `open_sheet_context_menu` from touch.
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeGestureLongPress(
    _env: JNIEnv,
    _class: JClass,
    canvas_id: i64,
    x: f32,
    y: f32,
    menu_mode: jni::sys::jboolean,
) -> jni::sys::jint {
    corro::gui::android_backend::gesture_long_press(
        canvas_id as u64,
        x as f64,
        y as f64,
        menu_mode != 0,
    )
}

/// Called from `SheetView.onTouchEvent(ACTION_MOVE)` during an active gesture.
/// Returns the shared outcome code (see `nativeGestureDown`).
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeGestureMove(
    _env: JNIEnv,
    _class: JClass,
    canvas_id: i64,
    x: f32,
    y: f32,
) -> jni::sys::jint {
    corro::gui::android_backend::gesture_move(canvas_id as u64, x as f64, y as f64)
}

/// Called from `SheetView.onTouchEvent(ACTION_UP)`.
///
/// A release that never dragged is reported as a *tap* (`3`), not applied:
/// the Java side then routes it to the canvas's own click handler
/// (`nativeOnTouch`), which is per canvas — so the tab strip's taps do not end
/// up moving the sheet's cursor. Rust applies the tap for the sheet itself.
///
/// Returns the outcome code (0 ignored, 2 scroll, 3 tap, …).
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeGestureUp(
    _env: JNIEnv,
    _class: JClass,
    canvas_id: i64,
    x: f32,
    y: f32,
) -> jni::sys::jint {
    corro::gui::android_backend::gesture_up(canvas_id as u64, x as f64, y as f64)
}

/// Called from `SheetView` when a gesture is cancelled (the system stole the
/// touch, the Activity paused). Never produces a tap.
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeGestureCancel(
    _env: JNIEnv,
    _class: JClass,
    canvas_id: i64,
) {
    corro::gui::android_backend::gesture_cancel(canvas_id as u64);
}

/// Called from `SheetView` for a touch drag that is scrolling: pans the sheet
/// by the pixel delta since the last event, converting to whole cells with the
/// zoom-aware metrics. Returns the counts applied as `[dRows, dCols]` (or
/// null), so the Java side can keep its sub-cell remainder.
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeDragBy(
    env: JNIEnv,
    _class: JClass,
    canvas_id: i64,
    dx: f32,
    dy: f32,
) -> jni::sys::jintArray {
    let (d_rows, d_cols) =
        corro::gui::android_backend::drag_viewport(canvas_id as u64, dx as f64, dy as f64);
    match env.new_int_array(2) {
        Ok(arr) => {
            let _ = env.set_int_array_region(&arr, 0, &[d_rows, d_cols]);
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

/// Called from `CorroTimer.run` (a `Handler.postDelayed` callback): runs the
/// timer Rust registered under `timer_id` and re-posts it if it repeats.
///
/// This is the Android counterpart of GTK's `g_timeout_add` and NWG's message
/// timers, which the same `rswidgets::core::add_periodic_tick` drives on the
/// desktop backends.
#[no_mangle]
pub extern "system" fn Java_com_corro_CorroTimer_nativeFire(
    _env: JNIEnv,
    _class: JClass,
    timer_id: i64,
) {
    rswidgets::backends::android::dispatch_timeout(timer_id as u64);
}

/// Called from `CorroItemSelected.onItemSelected`: a `Spinner` selection
/// reached Rust, so a dropdown's `connect_changed` callback fires.
#[no_mangle]
pub extern "system" fn Java_com_corro_CorroItemSelected_nativeItemSelected(
    _env: JNIEnv,
    _class: JClass,
    callback_id: i64,
    _position: i32,
) {
    rswidgets::backends::android::invoke_local_callback(callback_id as u64);
}

/// Called from `CorroChecked.onCheckedChanged`: a `CheckBox` / `RadioButton`
/// toggle reached Rust, so `connect_toggled` fires. Without this a toggle
/// could be set and read but never *report* a user tap.
#[no_mangle]
pub extern "system" fn Java_com_corro_CorroChecked_nativeChecked(
    _env: JNIEnv,
    _class: JClass,
    callback_id: i64,
    _checked: jni::sys::jboolean,
) {
    rswidgets::backends::android::invoke_local_callback(callback_id as u64);
}

/// Called from `CorroScrolled.onScroll`: the user scrolled a `ScrollView`, so
/// `ScrolledWindow::on_scroll` fires. This is the notification path that was
/// missing — a program scroll (`scroll_to`) and a user scroll now both reach
/// the model, so a host cannot end up with a view and a value that disagree.
#[no_mangle]
pub extern "system" fn Java_com_corro_CorroScrolled_nativeScrolled(
    _env: JNIEnv,
    _class: JClass,
    window_ptr: i64,
    vertical: i32,
    value: i32,
) {
    rswidgets::backends_android_adapter::dispatch_scrolled(
        window_ptr as *mut _,
        vertical != 0,
        value as f64,
    );
}

/// Called from `CorroDialogListener.onClick`: a dialog button was pressed.
///
/// The listener object carries the *role* (`positive` / `negative` /
/// `neutral`), because Android's `OnClickListener` callback receives the
/// dialog and a `which` constant but nothing identifying which button was
/// registered. Rust maps the role to the response id that
/// `Dialog::add_button` recorded, and runs the `connect_response` handler.
#[no_mangle]
pub extern "system" fn Java_com_corro_CorroDialogListener_nativeDialogButton(
    mut env: JNIEnv,
    _class: JClass,
    builder_ptr: i64,
    role: JObject,
) {
    let jstr: &jni::objects::JString = (&role).into();
    let Ok(role) = env.get_string(jstr) else {
        return;
    };
    let role: String = role.into();
    rswidgets::backends_android_adapter::dispatch_dialog_response(
        builder_ptr as *mut _,
        role.as_str(),
    );
}

/// Called from `CorroLayout.onGlobalLayout`: the view's bounds settled, so
/// run the `Window::set_layout_cb` callback. This is Android's counterpart of
/// GTK's size-allocate, and it is the only point at which a real width or
/// height exists — `nativeInit` returns before the first layout pass.
#[no_mangle]
pub extern "system" fn Java_com_corro_CorroLayout_nativeLayout(
    _env: JNIEnv,
    _class: JClass,
    view_ptr: i64,
) {
    rswidgets::backends_android_adapter::dispatch_layout(view_ptr as *mut _);
}

/// Called from `SheetView.onKeyDown` (or `CorroKeyView`), `MainActivity` and
/// `CorroKeyListener`: hand a key to corro's shared `handle_key`.
///
/// This is the bridge that makes the whole of the shared key handler
/// reachable on a phone. Before it existed, `dispatch_canvas_key` had a live
/// registry that nothing ever called, so every desktop binding in
/// `gui_backend::handle_key` — cursor arrows, Shift+arrows to extend a
/// selection, Home/End, PageUp/PageDown, Tab, Escape, Delete/Backspace,
/// F1/F2/F3, and the Ctrl accelerators — was dead on Android. The only key
/// that ever arrived was Enter, through the IME editor-action listener.
///
/// `keyCode` and `metaMask` are Android's own values; the translation to the
/// GDK keysym and the GdkModifierType bitmask happens in Rust, so
/// `handle_key` needs no Android arm. `canvasId` is the canvas that should
/// see the key (0 when none is focused); `viewPtr` is a focused entry, if
/// any, which is routed to the entry's own handler instead.
///
/// Returns true when a handler consumed the key, so the Java side knows
/// whether to swallow it (`onKeyDown` returning true) or let the platform
/// have it.
#[no_mangle]
pub extern "system" fn Java_com_corro_CorroKeyBridge_nativeKey(
    _env: JNIEnv,
    _class: JClass,
    key_code: jni::sys::jint,
    meta_mask: jni::sys::jint,
    canvas_id: jni::sys::jlong,
    view_ptr: jni::sys::jlong,
) -> jni::sys::jboolean {
    // A focused entry takes the key first, exactly as on a desktop where the
    // EditText's controller runs before the window's. Returning early matters
    // because the entry handler is where an in-cell caret key is applied; a
    // key that fell through to the grid handler would move the cursor while
    // the user was editing.
    if view_ptr != 0 && corro::gui::android_backend::dispatch_entry_key(
        view_ptr as u64,
        key_code,
        meta_mask,
    ) {
        return 1;
    }
    let canvas = if canvas_id == 0 {
        None
    } else {
        Some(canvas_id as u64)
    };
    let handled = corro::gui::android_backend::dispatch_key(key_code, meta_mask, canvas);
    jni::sys::jboolean::from(handled)
}

/// Called from `SheetView.onGenericMotionEvent`: a mouse right-click, which on
/// Android is a `BUTTON_SECONDARY` press rather than a touch.
///
/// This is the platform half of `gui_backend::open_sheet_context_menu`, whose
/// `on_click_button` registration previously had no Android source at all. The
/// long-press on a finger reaches the same place through `SheetView`'s gesture
/// detector; this is the mouse equivalent, so the emulator's pointer behaves
/// like a desktop's.
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeClickButton(
    _env: JNIEnv,
    _class: JClass,
    canvas_id: i64,
    x: f32,
    y: f32,
    button: i32,
    mods: i32,
) {
    rswidgets::backends_android_adapter::dispatch_canvas_click_button(
        canvas_id as u64,
        x as f64,
        y as f64,
        button as u32,
        mods as u32,
    );
}

/// Called from `SheetView` after a long press reported that it wants a menu:
/// open the cell context menu for the pressed cell.
///
/// This is the Android half of a desktop right-click. Before the pointer
/// bridge existed, `open_sheet_context_menu` always reported
/// `SHEET_MENU_UNAVAILABLE` — `Canvas::on_click_button` had no Android
/// source and the long press was already spoken for by range selection — so a
/// phone had no way to reach the cell actions at all.
#[no_mangle]
pub extern "system" fn Java_com_corro_SheetView_nativeCellContextMenu(
    _env: JNIEnv,
    _class: JClass,
    x: f32,
    y: f32,
) {
    corro::gui::android_backend::open_cell_context_menu(x as f64, y as f64);
}
