//! Android backend for corro: JNI entry point + run loop.
//!
//! The sheet itself renders through the shared [`super::gui_backend`]
//! pipeline (same `GuiState`, menus, dialogs, key handling as desktop);
//! the only Android-specific pieces live here:
//!
//! * [`run_android`] — build the widget tree against the Activity's root
//!   layout and run the backend event loop (Android drives the loop; this
//!   returns once the backend is up).
//! * [`native_init`] — the `#[no_mangle]` JNI export called from
//!   `MainActivity.nativeInit`, which stashes the JVM/Activity/layout in
//!   [`rswidgets::backends::android`] and boots a corro [`super::App`].

use std::path::PathBuf;

/// Run the corro GUI on Android. Mirrors `gui_backend::run_gui`: same
/// spreadsheet, same menus, same key handling — only the host window comes
/// from the Activity's root layout instead of a desktop toplevel.
pub fn run_android(
    corro_app: &mut super::App,
) -> Result<(), Box<dyn std::error::Error>> {
    super::gui_backend::run_gui(corro_app)
}

/// Boot a fresh corro [`super::App`] with no file (the Android activity
/// starts blank; files arrive later via content URIs) and run it.
///
/// The App is leaked (`Box::leak`): draw/input closures registered with the
/// backend hold a raw pointer to it and keep firing on the UI thread long
/// after this returns (Android drives the event loop — there is no blocking
/// `run()` keeping a stack App alive, unlike desktop). Same reason the wasm
/// `main()` leaks its App.
pub fn run_android_default() -> Result<(), Box<dyn std::error::Error>> {
    let app = super::App::new_with_paths(Vec::<PathBuf>::new());
    let app: &'static mut super::App = Box::leak(Box::new(app));
    app.set_backend(super::Backend::Gui);
    app.load_initial()?;
    run_android(app)
}

/// Build the Android menu strip and hand it to the layout.
///
/// The Android adapter's `create_menubar` has no view behind it (a phone
/// cannot show six text menus the way the GTK build does), so without this
/// the app shows no menus at all. The strip carries the inline quick
/// actions plus an overflow popup holding every top-level menu, and each
/// item dispatches the same `app.*` action name the desktop build uses.
#[cfg(target_os = "android")]
pub fn install_menu_strip() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    use crate::gui::menu;

    let strip_ptr = rswidgets::backends::android::create_menu_strip("corro")?;
    let bar = menu::menu_bar();
    for root in &bar {
        // Alternating label/action pairs, exactly what MenuStrip.addMenu
        // expects. Submenus are flattened one level (a popup submenu cannot
        // nest arbitrarily on a phone); labels lose their mnemonic marker,
        // which is meaningless on touch.
        let mut pairs: Vec<String> = Vec::new();
        collect_menu_pairs(root.submenu.as_deref().unwrap_or(&[]), &mut pairs);
        let refs: Vec<&str> = pairs.iter().map(|s| s.as_str()).collect();
        let label = root.label.replace('_', "");
        logcat(&format!("menu strip: adding {label} with {} entries", refs.len()));
        rswidgets::backends::android::menu_strip_add_menu(
            strip_ptr as *mut std::os::raw::c_void,
            &label,
            &refs,
        );
    }
    // Insert the strip as the FIRST child of the root layout, ahead of the
    // Rust-built vbox that `init_with_layout`/`attach_child` appended. Using
    // addView() alone would put it *after* the vbox, where the vbox's own
    // (opaque, full-height) children paint over it - the strip existed in the
    // hierarchy but never showed a pixel. addView(view, 0) places it at the
    // top, matching the desktop menubar's position.
    //
    // The strip must also be *pinned*: it is a sibling of the sheet inside a
    // vertical LinearLayout, and the sheet carries weight 1, so nothing stops
    // the strip being squeezed or scrolled out of view. WRAP_CONTENT height
    // with zero weight keeps it a fixed band at the top that the sheet
    // expands beneath; the sheet's own scrolling moves grid content only, so
    // the menu stays visible for the whole session.
    let root = rswidgets::backends::android::root_layout()?;
    rswidgets::backends::android::pin_child_at_top(
        root.as_obj().as_raw() as *mut std::os::raw::c_void,
        strip_ptr as *mut std::os::raw::c_void,
    );
    // The layout holds a global ref, so the strip stays alive for the process
    // lifetime; the Rust handle is only needed for the calls above.
    Ok(())
}

/// Flatten a menu subtree into alternating `label, action` strings.
#[cfg(target_os = "android")]
fn collect_menu_pairs(items: &[crate::gui::menu::MenuAction], out: &mut Vec<String>) {
    use crate::gui::menu::action_kind_to_name;
    for item in items {
        match item.submenu.as_deref() {
            Some(sub) => collect_menu_pairs(sub, out),
            None => {
                out.push(item.label.replace('_', ""));
                out.push(format!("app.{}", action_kind_to_name(item.action)));
            }
        }
    }
}

/// Run a menu action by name (from the Android menu strip).
///
/// The name is the same `app.<action>` string the desktop menu registers,
/// so this is a thin lookup rather than a second dispatch table.
#[cfg(target_os = "android")]
pub fn run_menu_action_by_name(name: &str) {
    super::gui_backend::dispatch_android_menu_action(name);
}


/// Report the row/column height in device pixels, so `SheetView` can convert
/// a touch drag in pixels into whole-cell counts (it keeps the sub-cell
/// remainder between events).
///
/// The Java side needs the same numbers Rust renders with — a hardcoded
/// constant there would drift from `row_h()`/`col_w()` as soon as the
/// density or metrics change, and touch scrolling would feel wrong by
/// exactly that factor. Because the metrics include the pinch scale, the Java
/// side re-reads them after a zoom (see `SheetView.refreshCellSize`), or a
/// pinch would leave the drag conversion at the old scale.
#[cfg(target_os = "android")]
pub fn touch_cell_size() -> (f64, f64) {
    (
        super::gui_backend::touch_row_h(),
        super::gui_backend::touch_col_w(),
    )
}

/// Apply a pinch-zoom step: multiply the sheet's view scale by `factor`
/// (`currentSpan / previousSpan` from Android's `ScaleGestureDetector`).
///
/// Called from `SheetView` once per scale-gesture event. Everything the sheet
/// draws *and* everything a finger hits derives from the same scale, so a
/// pinch moves the grid, the gutter, the headers and the hit targets together.
/// Returns the scale actually applied (clamped to
/// `gui_backend::MIN_VIEW_ZOOM`..=`MAX_VIEW_ZOOM`), so `SheetView` can refresh
/// its cached cell size and log the value.
#[cfg(target_os = "android")]
pub fn zoom_viewport(factor: f64) -> f64 {
    super::gui_backend::zoom_viewport_by(factor)
}

/// Reset the pinch scale to 1.0 (a double-tap, or the View menu's reset item).
/// Returns the applied scale (always 1.0).
#[cfg(target_os = "android")]
pub fn reset_viewport_zoom() -> f64 {
    super::gui_backend::reset_viewport_zoom()
}

/// Begin a touch gesture that may become a drag: `kind` is 0 for a mouse (drag
/// selects) and 1 for a finger (drag scrolls until a long press arms
/// selection). Wire this to `ACTION_DOWN`.
///
/// Returns the [`rswidgets::gridview::DragOutcome`] as a small integer so the
/// Java side can react without a second JNI call:
/// `0` ignored, `1` select, `2` scroll (deltas delivered separately by
/// `drag_viewport`), `3` tap, `4` long press.
#[cfg(target_os = "android")]
pub fn gesture_down(canvas_id: u64, x: f64, y: f64, is_touch: bool) -> i32 {
    super::gui_backend::mobile_gesture_down(canvas_id, x, y, is_touch)
}

/// A finger that has been still since `gesture_down` and has now held long
/// enough: arm selection so the following drag extends it instead of
/// scrolling. Wire this to `GestureDetector.onLongPress`.
#[cfg(target_os = "android")]
pub fn gesture_long_press(canvas_id: u64, x: f64, y: f64) -> i32 {
    super::gui_backend::mobile_gesture_long_press(canvas_id, x, y)
}

/// A pointer move while a gesture is active. Returns the
/// [`rswidgets::gridview::DragOutcome`] code (see [`gesture_down`]); for a
/// scroll the Java side applies the pan through [`drag_viewport`].
#[cfg(target_os = "android")]
pub fn gesture_move(canvas_id: u64, x: f64, y: f64) -> i32 {
    super::gui_backend::mobile_gesture_move(canvas_id, x, y)
}

/// A pointer release. Returns the outcome code (see [`gesture_down`]); a
/// `3` (tap) means the host should treat it as a click, which the shared
/// handler has already applied.
#[cfg(target_os = "android")]
pub fn gesture_up(canvas_id: u64, x: f64, y: f64) -> i32 {
    super::gui_backend::mobile_gesture_up(canvas_id, x, y)
}

/// Cancel an in-flight gesture (the system stole the touch, the activity
/// paused). Never produces a tap.
#[cfg(target_os = "android")]
pub fn gesture_cancel(canvas_id: u64) {
    super::gui_backend::mobile_gesture_cancel(canvas_id);
}

/// Pan the sheet by a pixel delta from a touch drag, converting to whole
/// cells with the *current* (zoom-aware) metrics.
///
/// `dx`/`dy` are the movement since the last event in device pixels; positive
/// `dy` means the content should move down the screen (drag downward reveals
/// earlier rows), so the row delta is negated, exactly like the existing
/// `scrollByDrag` path. Returns the row/column counts actually applied so the
/// caller can keep its sub-cell remainder.
#[cfg(target_os = "android")]
pub fn drag_viewport(canvas_id: u64, dx: f64, dy: f64) -> (i32, i32) {
    super::gui_backend::drag_viewport_by_pixels(canvas_id, dx, dy)
}

/// Write a line to logcat (`log -t corro` shows these). Best-effort: the
/// backend may not be initialised yet, so failures are swallowed.
pub fn logcat(msg: &str) {
    let _ = rswidgets::backends::android::with_env_and_activity(|env, _activity| {
        let log_cls = env.find_class("android/util/Log")?;
        let tag = env.new_string("corro")?;
        let jmsg = env.new_string(msg)?;
        env.call_static_method(
            &log_cls,
            "d",
            "(Ljava/lang/String;Ljava/lang/String;)I",
            &[(&tag).into(), (&jmsg).into()],
        )?;
        Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
    });
}

/// JNI entry called from `MainActivity.nativeInit(activity, rootLayout)`.
///
/// Wires the JVM/Activity/layout into [`rswidgets::backends::android`] and
/// boots a blank corro [`super::App`]. Called by the `corro_android` cdylib
/// crate's `#[no_mangle]` export (which lives there so the linker keeps it).
#[cfg(target_os = "android")]
pub fn android_main(
    env: &mut jni::JNIEnv<'_>,
    activity: &jni::objects::JObject<'_>,
    root_layout: &jni::objects::JObject<'_>,
) -> Result<(), String> {
    // NOTE: logcat needs JAVA_VM, which is only set after init_with_layout,
    // so the first line below cannot log yet (it is swallowed silently).
    rswidgets::backends::android::init_with_layout(env, activity, root_layout)
        .map_err(|e| format!("corro backend init failed: {e}"))?;
    run_android_default().map_err(|e| format!("corro run failed: {e}"))?;
    Ok(())
}
