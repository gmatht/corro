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
pub fn run_android_default() -> Result<(), Box<dyn std::error::Error>> {
    let mut app = super::App::new_with_paths(Vec::<PathBuf>::new());
    app.set_backend(super::Backend::Gui);
    app.load_initial()?;
    run_android(&mut app)
}

/// JNI export called from `MainActivity.nativeInit(activity, rootLayout)`.
///
/// Signature: `Java_<package>_MainActivity_nativeInit(JNIEnv, jclass,
/// jobject activity, jobject rootLayout)`. The package prefix is filled in
/// by the APK build (see `android/MainActivity.java`); the body only wires
/// the backend and boots corro.
#[cfg(target_os = "android")]
#[no_mangle]
pub extern "system" fn Java_com_corro_MainActivity_nativeInit(
    mut env: jni::JNIEnv<'_>,
    _class: jni::objects::JClass<'_>,
    activity: jni::objects::JObject<'_>,
    root_layout: jni::objects::JObject<'_>,
) {
    let init = rswidgets::backends::android::init_with_layout(&mut env, &activity, &root_layout);
    if let Err(e) = init {
        let _ = env.throw_new("java/lang/RuntimeException", format!("corro init failed: {e}"));
        return;
    }
    if let Err(e) = run_android_default() {
        let msg = format!("corro run failed: {e}");
        let _ = env.throw_new("java/lang/RuntimeException", msg);
    }
}
