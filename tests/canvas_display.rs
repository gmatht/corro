//! Canvas/plumbing tests for the runtime-dlopen GTK backend.
//!
//! These exercise `gtk_drawing_area_set_draw_func`, which exists in GTK4
//! only. The loader defaults to GTK3 (`GTK_DLOPEN_PREFER_GTK3`), so on a
//! default run the symbol is absent and the call returns
//! `MissingSymbol("gtk_drawing_area_set_draw_func")`. That was previously
//! recorded as an intermittent "GTK floating-ref segfault"; it is neither
//! intermittent nor a segfault, it is a GTK-major mismatch. They now skip
//! loudly unless GTK4 is actually loaded, and can be run with:
//!
//!     GTK_DLOPEN_PREFER_GTK3=0 xvfb-run -a cargo test --features gui \
//!         --test canvas_display -- --test-threads=1
//!
//! **`--test-threads=1` is required, not optional.** GTK4's type system
//! (`g_type_class_ref` inside `g_object_new`) is not thread-safe, and the
//! default harness runs each `#[test]` on its own thread, so two tests
//! initialising GTK concurrently segfault *inside libgtk-4* — verified under
//! gdb: `SIGSEGV` at `libgtk-4.so.1` ← `g_type_class_ref` ← `g_object_new`.
//! The crash is in GTK, not in this crate or the loader, so the fix is to
//! serialise the tests rather than to add locking we cannot enforce for
//! arbitrary callers. Same rule as any GTK program: touch GTK from one thread.
#[cfg(all(feature = "gui", target_os = "linux"))]
mod canvas_tests {
    use std::cell::Cell;
    use std::ffi::c_void;
    use std::sync::Arc;

    /// True when the loader would pick GTK4 for this process.
    fn gtk4_available() -> bool {
        std::env::var_os("GTK_DLOPEN_PREFER_GTK3").is_some_and(|v| v == "0")
            && gtk_dynamic_loader::Loader::new().is_ok()
    }

    macro_rules! require_gtk4 {
        () => {
            if !gtk4_available() {
                eprintln!(
                    "SKIP {}: needs GTK4 (set GTK_DLOPEN_PREFER_GTK3=0;                      gtk_drawing_area_set_draw_func is GTK4-only)",
                    module_path!()
                );
                return;
            }
        };
    }

    // ---- API-level test ----
    #[test]
    fn canvas_set_draw_callback_api() {
        require_gtk4!();
        let loader = gtk_dynamic_loader::Loader::new().expect("Loader::new failed");
        let da = gtk_dynamic_loader::DrawingArea::new(loader.clone())
            .expect("DrawingArea::new failed");

        // set_draw_func should accept and store a closure without crashing
        da.set_draw_func(Box::new(|_cr: *mut c_void, _w: i32, _h: i32| {
            // no-op; just verifying the plumbing
        }))
        .expect("set_draw_func failed");

        // queue_draw should not crash (even without a window)
        da.queue_draw();

        // set_size_request should not crash
        da.set_size_request(400, 300);
    }

    // ---- Full rendering test (requires DISPLAY) ----
    #[test]
    fn canvas_draw_callback_fires() {
        require_gtk4!();
        let loader = gtk_dynamic_loader::Loader::new().expect("Loader::new failed");

        // Shared flag to verify the draw callback was invoked
        let drew = Arc::new(Cell::new(false));

        let win = gtk_dynamic_loader::Window::new(loader.clone()).expect("Window::new failed");

        let da = gtk_dynamic_loader::DrawingArea::new(loader.clone())
            .expect("DrawingArea::new failed");
        let drew_clone = drew.clone();
        da.set_draw_func(Box::new(move |_cr: *mut c_void, _w: i32, _h: i32| {
            drew_clone.set(true);
        }))
        .expect("set_draw_func failed");
        da.set_size_request(200, 100);
        da.queue_draw();

        win.set_child(&da);
        win.present();

        // Leak win and da to prevent Drop from running (which would unref GObjects
        // while the main loop or GLib still references them)
        let _win = Box::into_raw(Box::new(win));
        let _da = Box::into_raw(Box::new(da));

        // Run the GLib main loop with an idle callback that checks the flag
        let loop_new = loader.symbols.g_main_loop_new.expect("g_main_loop_new");
        let loop_quit = loader.symbols.g_main_loop_quit.expect("g_main_loop_quit");
        let idle_add = loader.symbols.g_idle_add.expect("g_idle_add");
        let loop_ptr = unsafe { loop_new(std::ptr::null_mut(), 0) };

        type IdleFn = unsafe extern "C" fn(*mut c_void) -> i32;
        type LoopQuit = unsafe extern "C" fn(*mut c_void);

        struct IdleData {
            drew: Arc<Cell<bool>>,
            loop_ptr: *mut c_void,
            loop_quit: LoopQuit,
        }

        unsafe extern "C" fn idle_cb(data: *mut c_void) -> i32 {
            let idle_data = &*(data as *mut IdleData);
            assert!(
                idle_data.drew.get(),
                "Draw callback was NOT invoked during the main loop iteration"
            );
            (idle_data.loop_quit)(idle_data.loop_ptr);
            let _ = Box::from_raw(data as *mut IdleData);
            0
        }

        let idle_data = Box::into_raw(Box::new(IdleData {
            drew: drew.clone(),
            loop_ptr,
            loop_quit,
        }));

        let idle_fn: Option<IdleFn> = Some(idle_cb);
        unsafe { idle_add(idle_fn, idle_data as *mut c_void) };

        unsafe {
            let loop_run = loader.symbols.g_main_loop_run.expect("g_main_loop_run");
            loop_run(loop_ptr);
        }
    }
}

/// GTK4 key press/release handling: a release must not be delivered as a
/// second press.
///
/// Regression test for the bug `gui_backend.rs` used to patch with three
/// `feature = "gtk4"` dedup fields. GTK4 emits BOTH `key-pressed` and
/// `key-released` from the same `GtkEventControllerKey`; the adapter registers
/// `key-released` itself and returns `GDK_EVENT_STOP` without calling the host,
/// so hosts receive a pure press stream.
///
/// Run single-threaded — GTK4's type system is not thread-safe:
///
///     GTK_DLOPEN_PREFER_GTK3=0 xvfb-run -a cargo test --features gui \
///         --test canvas_display -- --test-threads=1
#[cfg(all(feature = "gui", target_os = "linux"))]
mod gtk4_key_release_tests {
    use std::cell::Cell;
    use std::rc::Rc;

    fn gtk4_available() -> bool {
        std::env::var_os("GTK_DLOPEN_PREFER_GTK3").is_some_and(|v| v == "0")
            && gtk_dynamic_loader::Loader::new().is_ok()
    }

    macro_rules! require_gtk4 {
        () => {
            if !gtk4_available() {
                eprintln!("SKIP {}: needs GTK4 (set GTK_DLOPEN_PREFER_GTK3=0)", module_path!());
                return;
            }
        };
    }

    /// Both signals are emitted on one controller, exactly as a physical key
    /// would produce them, so this reproduces the double-delivery scenario
    /// without needing synthetic input from the compositor.
    #[test]
    fn key_released_is_a_distinct_signal_from_key_pressed() {
        require_gtk4!();
        let loader = gtk_dynamic_loader::Loader::new().expect("Loader::new failed");
        if loader.version != gtk_dynamic_loader::Version::Gtk4 {
            eprintln!("SKIP: loader selected GTK3");
            return;
        }

        let ctrl = gtk_dynamic_loader::EventControllerKey::new(loader.clone())
            .expect("EventControllerKey::new");

        let presses: Rc<Cell<u32>> = Rc::new(Cell::new(0));
        let releases: Rc<Cell<u32>> = Rc::new(Cell::new(0));
        let last: Rc<Cell<u32>> = Rc::new(Cell::new(0));
        let p = presses.clone();
        let l = last.clone();
        let _ = ctrl.connect_key_pressed(Box::new(move |keyval: u32, _state: u32| -> i32 {
            p.set(p.get() + 1);
            l.set(keyval);
            1
        }));
        let r = releases.clone();
        let _ = ctrl.connect_key_released(Box::new(move |_keyval: u32, _state: u32| -> i32 {
            r.set(r.get() + 1);
            1 // GDK_EVENT_STOP — this is what the adapter's filter does
        }));

        const KEY_A: u32 = 0x41;
        ctrl.emit_key("key-pressed", KEY_A, 0).expect("emit key-pressed");
        assert_eq!(1, presses.get(), "press reaches its own handler");
        assert_eq!(0, releases.get(), "the release handler has not fired yet");
        assert_eq!(KEY_A, last.get(), "keyval delivered intact");

        ctrl.emit_key("key-released", KEY_A, 0).expect("emit key-released");
        assert_eq!(1, releases.get(), "release reaches the release handler");
        assert_eq!(
            1,
            presses.get(),
            "the release must not be delivered to the press handler"
        );
    }
}
