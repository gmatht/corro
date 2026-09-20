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
