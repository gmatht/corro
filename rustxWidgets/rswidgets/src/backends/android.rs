#[cfg(target_os = "android")]
mod android_backend {
    use jni::objects::{GlobalRef, JObject};
    use jni::JNIEnv;
    use once_cell::sync::OnceCell;
    use once_cell::sync::Lazy;
    use std::collections::HashMap;
    use std::error::Error as StdError;
    use std::sync::Mutex;

    static JAVA_VM: OnceCell<jni::JavaVM> = OnceCell::new();
    static ACTIVITY: OnceCell<GlobalRef> = OnceCell::new();
    static ROOT_LAYOUT: OnceCell<GlobalRef> = OnceCell::new();
    /// The app's ClassLoader, captured while we are still on the Activity's
    /// thread. `JNIEnv::find_class` on a JNI-attached thread resolves through
    /// the *system* loader and cannot see APK classes (`com.corro.*`), so the
    /// app loader is cached here and used by `load_app_class`.
    static APP_CLASS_LOADER: OnceCell<GlobalRef> = OnceCell::new();

    /// Keeps GlobalRefs alive so raw jobject pointers remain valid.
    static KEEP_ALIVE: Lazy<Mutex<Vec<GlobalRef>>> = Lazy::new(|| Mutex::new(Vec::new()));

    /// Create a GlobalRef from a local JObject, store it, and return the raw pointer.
    pub fn make_global_ref(env: &mut JNIEnv, obj: &JObject<'_>) -> Result<jni::sys::jobject, Box<dyn StdError + Send + Sync>> {
        let gref = env.new_global_ref(obj)?;
        let raw = gref.as_obj().as_raw();
        KEEP_ALIVE.lock().unwrap().push(gref);
        Ok(raw)
    }

    pub fn init(env: &mut JNIEnv, activity: &JObject<'_>) -> Result<(), Box<dyn StdError + Send + Sync>> {
        let vm = env.get_java_vm()?;
        let activity_ref = env.new_global_ref(activity)?;

        let content_id: jni::sys::jint = env.get_static_field(
            "android/R$id",
            "content",
            "I",
        )?.i()?;

        let root_view = env.call_method(
            activity,
            "findViewById",
            "(I)Landroid/view/View;",
            &[content_id.into()],
        )?;
        let root_ref = env.new_global_ref(root_view.l()?)?;

        JAVA_VM.set(vm).map_err(|_| "JAVA_VM already initialized")?;
        ACTIVITY.set(activity_ref).map_err(|_| "ACTIVITY already initialized")?;
        ROOT_LAYOUT.set(root_ref).map_err(|_| "ROOT_LAYOUT already initialized")?;

        Ok(())
    }

    /// Like init() but accepts an explicitly provided root layout ViewGroup.
    /// Avoids the JNI lookup of android.R.id.content.
    pub fn init_with_layout(env: &mut JNIEnv, activity: &JObject<'_>, layout: &JObject<'_>) -> Result<(), Box<dyn StdError + Send + Sync>> {
        let vm = env.get_java_vm()?;
        let activity_ref = env.new_global_ref(activity)?;
        let root_ref = env.new_global_ref(layout)?;

        // Capture the APK's ClassLoader now, while this thread's context
        // loader is the app's (we are inside MainActivity.onCreate).
        if let Ok(loader) = env
            .call_method(activity, "getClassLoader", "()Ljava/lang/ClassLoader;", &[])
            .and_then(|l| l.l())
            .and_then(|obj| env.new_global_ref(&obj))
        {
            let _ = APP_CLASS_LOADER.set(loader);
            logcat_rs("APP_CLASS_LOADER captured");
        } else {
            logcat_rs("APP_CLASS_LOADER capture FAILED");
        }

        JAVA_VM.set(vm).map_err(|_| "JAVA_VM already initialized")?;
        ACTIVITY.set(activity_ref).map_err(|_| "ACTIVITY already initialized")?;
        ROOT_LAYOUT.set(root_ref).map_err(|_| "ROOT_LAYOUT already initialized")?;

        Ok(())
    }

    /// Resolve an app (APK) class by binary name (`com.corro.Foo`) through
    /// the cached app ClassLoader. Falls back to `JNIEnv::find_class` when
    /// no loader was captured (e.g. framework classes still resolve).
    pub fn load_app_class<'a>(
        env: &mut JNIEnv<'a>,
        class_name: &str,
    ) -> Result<jni::objects::JClass<'a>, Box<dyn StdError + Send + Sync>> {
        if let Some(loader) = APP_CLASS_LOADER.get() {
            let jname = env.new_string(class_name.replace('/', "."))?;
            let cls = env.call_method(
                loader.as_obj(),
                "loadClass",
                "(Ljava/lang/String;)Ljava/lang/Class;",
                &[(&jname).into()],
            )?;
            let obj = cls.l()?;
            // Class<...> is a JClass in jni's type model.
            let cls: jni::objects::JClass = jni::objects::JClass::from(obj);
            return Ok(cls);
        }
        Ok(env.find_class(class_name)?)
    }

    pub fn root_layout() -> Result<&'static GlobalRef, Box<dyn StdError + Send + Sync>> {
        ROOT_LAYOUT.get().ok_or_else(|| "ROOT_LAYOUT not initialized (call init first)".into())
    }

    pub fn activity_ref() -> Result<&'static GlobalRef, Box<dyn StdError + Send + Sync>> {
        ACTIVITY.get().ok_or_else(|| "ACTIVITY not initialized (call init first)".into())
    }

    pub fn with_env_and_activity<F, T>(f: F) -> Result<T, Box<dyn StdError + Send + Sync>>
    where
        F: FnOnce(&mut JNIEnv<'_>, &GlobalRef) -> Result<T, Box<dyn StdError + Send + Sync>>,
    {
        let vm = JAVA_VM.get().ok_or_else(|| "JAVA_VM not initialized (call init first)")?;
        let activity = ACTIVITY.get().ok_or_else(|| "ACTIVITY not initialized (call init first)")?;

        let mut env = vm.attach_current_thread()?;
        f(&mut env, activity)
    }

    pub fn is_initialized() -> bool {
        JAVA_VM.get().is_some()
    }

    /// Write a debug line to logcat under the `rswidgets` tag. Best-effort:
    /// silently dropped when the backend is not initialised yet.
    pub fn logcat_rs(msg: &str) {
        let _ = with_env_and_activity(|env, _activity| {
            let log_cls = env.find_class("android/util/Log")?;
            let tag = env.new_string("rswidgets")?;
            let jmsg = env.new_string(msg)?;
            env.call_static_method(
                &log_cls,
                "d",
                "(Ljava/lang/String;Ljava/lang/String;)I",
                &[(&tag).into(), (&jmsg).into()],
            )?;
            Ok::<_, Box<dyn StdError + Send + Sync>>(())
        });
    }

    // ---- Callback registry ----

    static CALLBACKS: Lazy<Mutex<HashMap<u64, Box<dyn FnMut() + Send>>>> = Lazy::new(|| Mutex::new(HashMap::new()));
    static NEXT_CALLBACK_ID: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(1);

    pub fn register_callback(f: Box<dyn FnMut() + Send>) -> u64 {
        let id = NEXT_CALLBACK_ID.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        let mut map = CALLBACKS.lock().unwrap();
        map.insert(id, f);
        id
    }

    pub fn unregister_callback(id: u64) {
        let mut map = CALLBACKS.lock().unwrap();
        map.remove(&id);
    }

    pub fn invoke_callback(id: u64) {
        let mut map = CALLBACKS.lock().unwrap();
        if let Some(f) = map.get_mut(&id) {
            f();
        }
    }

    pub fn dispatch_callback(id: u64) {
        invoke_callback(id);
    }

    /// Create a JNI View.OnClickListener that dispatches to the Rust callback registry.
    /// This looks for a user-provided `RustCallback` class (see example project).
    /// Returns Ok(listener) if the class exists, or Ok(None) if not available.
    pub fn try_create_onclick_listener<'a>(
        env: &mut JNIEnv<'a>,
        callback_id: u64,
    ) -> Result<Option<JObject<'a>>, Box<dyn StdError + Send + Sync>> {
        let cls = match load_app_class(env, "com/example/RustCallback") {
            Ok(c) => c,
            Err(_) => {
                let _ = env.exception_clear();
                return Ok(None);
            }
        };
        let listener = env.new_object(
            &cls,
            "(J)V",
            &[(callback_id as i64).into()],
        )?;
        Ok(Some(listener))
    }

    // ---- AndroidApp ----

    pub struct AndroidApp;

    impl AndroidApp {
        pub fn new() -> Result<Box<dyn crate::backends::BackendApp>, Box<dyn StdError + Send + Sync>> {
            Ok(Box::new(AndroidApp))
        }
    }

    impl crate::backends::BackendApp for AndroidApp {
        fn run(self: Box<Self>) -> Result<(), Box<dyn StdError + Send + Sync>> {
            Ok(())
        }
    }

    pub fn init_backend() -> Result<Box<dyn crate::backends::BackendApp>, Box<dyn StdError + Send + Sync>> {
        AndroidApp::new()
    }

    // ---- Factory functions ----

    pub fn create_window() -> Result<jni::sys::jobject, Box<dyn StdError + Send + Sync>> {
        let activity = activity_ref()?;
        Ok(activity.as_obj().as_raw())
    }

    pub fn create_button(label: &str) -> Result<jni::sys::jobject, Box<dyn StdError + Send + Sync>> {
        with_env_and_activity(|env, activity| {
            let ctx = activity.as_obj();
            let btn = env.new_object(
                "android/widget/Button",
                "(Landroid/content/Context;)V",
                &[(&ctx).into()],
            )?;
            let j_label = env.new_string(label)?;
            env.call_method(
                &btn,
                "setText",
                "(Ljava/lang/CharSequence;)V",
                &[(&j_label).into()],
            )?;
            make_global_ref(env, &btn)
        })
    }

    pub fn create_label(text: &str) -> Result<jni::sys::jobject, Box<dyn StdError + Send + Sync>> {
        with_env_and_activity(|env, activity| {
            let ctx = activity.as_obj();
            let tv = env.new_object(
                "android/widget/TextView",
                "(Landroid/content/Context;)V",
                &[(&ctx).into()],
            )?;
            let j_text = env.new_string(text)?;
            env.call_method(
                &tv,
                "setText",
                "(Ljava/lang/CharSequence;)V",
                &[(&j_text).into()],
            )?;
            make_global_ref(env, &tv)
        })
    }

    pub fn create_box(orientation: i32, _spacing: i32) -> Result<jni::sys::jobject, Box<dyn StdError + Send + Sync>> {
        with_env_and_activity(|env, activity| {
            let ctx = activity.as_obj();
            let layout = env.new_object(
                "android/widget/LinearLayout",
                "(Landroid/content/Context;)V",
                &[(&ctx).into()],
            )?;
            env.call_method(
                &layout,
                "setOrientation",
                "(I)V",
                &[orientation.into()],
            )?;
            make_global_ref(env, &layout)
        })
    }

    pub fn create_entry() -> Result<jni::sys::jobject, Box<dyn StdError + Send + Sync>> {
        with_env_and_activity(|env, activity| {
            let ctx = activity.as_obj();
            let edit = env.new_object(
                "android/widget/EditText",
                "(Landroid/content/Context;)V",
                &[(&ctx).into()],
            )?;
            make_global_ref(env, &edit)
        })
    }

    pub fn create_grid() -> Result<jni::sys::jobject, Box<dyn StdError + Send + Sync>> {
        with_env_and_activity(|env, activity| {
            let ctx = activity.as_obj();
            let grid = env.new_object(
                "android/widget/GridLayout",
                "(Landroid/content/Context;)V",
                &[(&ctx).into()],
            )?;
            make_global_ref(env, &grid)
        })
    }

    pub fn create_dropdown(items: &[&str]) -> Result<jni::sys::jobject, Box<dyn StdError + Send + Sync>> {
        with_env_and_activity(|env, activity| {
            let ctx = activity.as_obj();
            let spinner = env.new_object(
                "android/widget/Spinner",
                "(Landroid/content/Context;)V",
                &[(&ctx).into()],
            )?;
            let array_cls = env.find_class("java/util/ArrayList")?;
            let list = env.new_object(
                &array_cls,
                "()V",
                &[],
            )?;
            for item in items {
                let j_item = env.new_string(item)?;
                env.call_method(
                    &list,
                    "add",
                    "(Ljava/lang/Object;)Z",
                    &[(&j_item).into()],
                )?;
            }
            let layout_id = &env.get_static_field(
                "android/R$layout",
                "simple_spinner_item",
                "I",
            )?.i()?;
            let adapter = env.new_object(
                "android/widget/ArrayAdapter",
                "(Landroid/content/Context;ILjava/util/List;)V",
                &[
                    (&ctx).into(),
                    (*layout_id).into(),
                    (&list).into(),
                ],
            )?;
            env.call_method(
                &spinner,
                "setAdapter",
                "(Landroid/widget/SpinnerAdapter;)V",
                &[(&adapter).into()],
            )?;
            make_global_ref(env, &spinner)
        })
    }

    pub fn create_checkbutton(label: &str) -> Result<jni::sys::jobject, Box<dyn StdError + Send + Sync>> {
        with_env_and_activity(|env, activity| {
            let ctx = activity.as_obj();
            let cb = env.new_object(
                "android/widget/CheckBox",
                "(Landroid/content/Context;)V",
                &[(&ctx).into()],
            )?;
            let j_label = env.new_string(label)?;
            env.call_method(
                &cb,
                "setText",
                "(Ljava/lang/CharSequence;)V",
                &[(&j_label).into()],
            )?;
            make_global_ref(env, &cb)
        })
    }

    pub fn create_radiobutton(group_ptr: *mut std::os::raw::c_void, label: &str) -> Result<jni::sys::jobject, Box<dyn StdError + Send + Sync>> {
        with_env_and_activity(|env, activity| {
            let ctx = activity.as_obj();
            let rb = env.new_object(
                "android/widget/RadioButton",
                "(Landroid/content/Context;)V",
                &[(&ctx).into()],
            )?;
            let j_label = env.new_string(label)?;
            env.call_method(
                &rb,
                "setText",
                "(Ljava/lang/CharSequence;)V",
                &[(&j_label).into()],
            )?;
            if !group_ptr.is_null() {
                let group_obj = unsafe { jni::objects::JObject::from_raw(group_ptr as jni::sys::jobject) };
                env.call_method(
                    &group_obj,
                    "addView",
                    "(Landroid/view/View;)V",
                    &[(&rb).into()],
                )?;
            }
            make_global_ref(env, &rb)
        })
    }

    pub fn create_radiogroup() -> Result<jni::sys::jobject, Box<dyn StdError + Send + Sync>> {
        with_env_and_activity(|env, activity| {
            let ctx = activity.as_obj();
            let rg = env.new_object(
                "android/widget/RadioGroup",
                "(Landroid/content/Context;)V",
                &[(&ctx).into()],
            )?;
            make_global_ref(env, &rg)
        })
    }

    pub fn create_dialog() -> Result<jni::sys::jobject, Box<dyn StdError + Send + Sync>> {
        with_env_and_activity(|env, activity| {
            let ctx = activity.as_obj();
            let builder = env.new_object(
                "android/app/AlertDialog$Builder",
                "(Landroid/content/Context;)V",
                &[(&ctx).into()],
            )?;
            make_global_ref(env, &builder)
        })
    }

    pub fn create_textview() -> Result<jni::sys::jobject, Box<dyn StdError + Send + Sync>> {
        with_env_and_activity(|env, activity| {
            let ctx = activity.as_obj();
            let tv = env.new_object(
                "android/widget/EditText",
                "(Landroid/content/Context;)V",
                &[(&ctx).into()],
            )?;
            let gravity = env.get_static_field(
                "android/view/Gravity",
                "TOP",
                "I",
            )?.i()?;
            env.call_method(
                &tv,
                "setGravity",
                "(I)V",
                &[gravity.into()],
            )?;
            let input_type = env.get_static_field(
                "android/text/InputType",
                "TYPE_TEXT_FLAG_MULTI_LINE",
                "I",
            )?.i()?;
            let class_type = env.get_static_field(
                "android/text/InputType",
                "TYPE_CLASS_TEXT",
                "I",
            )?.i()?;
            env.call_method(
                &tv,
                "setInputType",
                "(I)V",
                &[(class_type | input_type).into()],
            )?;
            env.call_method(
                &tv,
                "setMinLines",
                "(I)V",
                &[3i32.into()],
            )?;
            make_global_ref(env, &tv)
        })
    }

    /// Plain `android.view.View` backing a [`crate::backends_android_adapter::Canvas`].
    /// A custom `SheetView` (with `onDraw` funneling `DrawContext` calls back
    /// into Rust) replaces this once the Java side lands; until then a plain
    /// view keeps the widget tree valid so appends never see bogus handles.
    pub fn create_canvas_view() -> Result<(jni::sys::jobject, u64), Box<dyn StdError + Send + Sync>> {
        let class = SHEET_VIEW_CLASS.get().cloned().unwrap_or_else(|| "android/view/View".to_string());
        with_env_and_activity(|env, activity| {
            let ctx = activity.as_obj();
            // Custom SheetViews take (Context, long canvasId); the platform
            // View takes (Context). Try the two-arg form first, fall back.
            if class != "android/view/View" {
                let canvas_id = NEXT_CANVAS_ID.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
                let resolved = load_app_class(env, &class).ok();
                if let Some(resolved) = resolved {
                    match env.new_object(
                        &resolved,
                        "(Landroid/content/Context;J)V",
                        &[(&ctx).into(), (canvas_id as i64).into()],
                    ) {
                        Ok(v) => {
                            let raw = make_global_ref(env, &v)?;
                            register_canvas_view(raw, canvas_id);
                            return Ok((raw, canvas_id));
                        }
                        Err(_) => {
                            let _ = env.exception_clear();
                        }
                    }
                }
            }
            let canvas_id = NEXT_CANVAS_ID.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
            let view = env.new_object(
                "android/view/View",
                "(Landroid/content/Context;)V",
                &[(&ctx).into()],
            )?;
            let raw = make_global_ref(env, &view)?;
            register_canvas_view(raw, canvas_id);
            Ok((raw, canvas_id))
        })
    }

    static SHEET_VIEW_CLASS: OnceCell<String> = OnceCell::new();
    static NEXT_CANVAS_ID: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(1);

    /// View-pointer -> canvas-id map (the draw registry keys on canvas id;
    /// the Rust `Canvas` handle is the view pointer).
    static CANVAS_IDS: Lazy<Mutex<HashMap<usize, u64>>> = Lazy::new(|| Mutex::new(HashMap::new()));

    /// Views that asked to expand (`set_hexpand`). Android has no expand
    /// flag: the request is recorded here and honoured by
    /// `BoxWidget::append` as `LinearLayout` weight 1, so the formula entry
    /// grows instead of measuring to zero width.
    static EXPANDING: Lazy<Mutex<std::collections::HashSet<usize>>> =
        Lazy::new(|| Mutex::new(std::collections::HashSet::new()));

    /// Minimum content width in characters (GTK `set_width_chars`), kept so
    /// an empty EditText still gets a usable box.
    static MIN_CHARS: Lazy<Mutex<HashMap<usize, i32>>> =
        Lazy::new(|| Mutex::new(HashMap::new()));

    pub fn set_view_expanding(view_ptr: *mut std::os::raw::c_void, expand: bool) {
        let mut set = EXPANDING.lock().unwrap();
        if expand {
            set.insert(view_ptr as usize);
        } else {
            set.remove(&(view_ptr as usize));
        }
    }

    pub fn is_view_expanding(view_ptr: *mut std::os::raw::c_void) -> bool {
        EXPANDING.lock().unwrap().contains(&(view_ptr as usize))
    }

    pub fn set_view_min_chars(view_ptr: *mut std::os::raw::c_void, n: i32) {
        MIN_CHARS.lock().unwrap().insert(view_ptr as usize, n);
    }

    pub fn view_min_chars(view_ptr: *mut std::os::raw::c_void) -> i32 {
        MIN_CHARS.lock().unwrap().get(&(view_ptr as usize)).copied().unwrap_or(0)
    }

    fn register_canvas_view(view_raw: jni::sys::jobject, canvas_id: u64) {
        CANVAS_IDS.lock().unwrap().insert(view_raw as usize, canvas_id);
    }

    /// Canvas id for a view pointer (0 = unknown).
    pub fn canvas_id_for_view(view_ptr: *mut std::os::raw::c_void) -> u64 {
        CANVAS_IDS.lock().unwrap().get(&(view_ptr as usize)).copied().unwrap_or(0)
    }

    /// True when `view_ptr` is a canvas view (SheetView or plain canvas):
    /// used for layout weighting (canvases grow, chrome wraps).
    pub fn is_canvas_view(view_ptr: *mut std::os::raw::c_void) -> bool {
        canvas_id_for_view(view_ptr) != 0
    }

    /// Register the fully-qualified custom View class used for canvases
    /// (e.g. `com.corro.SheetView`). Must have a `(Context, long)` ctor
    /// taking the Rust canvas id. Falls back to a plain View per canvas
    /// when unset or when instantiation fails.
    pub fn set_sheet_view_class(class: &str) {
        let _ = SHEET_VIEW_CLASS.set(class.to_string());
    }

    /// Install the shared `TextWatcher`/`OnEditorActionListener` on an
    /// EditText so text changes and IME "Done" reach Rust. The listeners are
    /// `com.corro.CorroTextWatcher` / `com.corro.CorroEditorAction` when
    /// present; absent classes make this a no-op (backend still works, just
    /// without the typing path).
    pub fn attach_text_watcher(entry_ptr: *mut std::os::raw::c_void) {
        let class_name = match ENTRY_LISTENER_CLASSES.get() {
            Some((watcher, _)) => watcher.clone(),
            None => "com/corro/CorroTextWatcher".to_string(),
        };
        attach_entry_listener(entry_ptr, "addTextChangedListener", &class_name);
    }

    /// See [`attach_text_watcher`]; wires `setOnEditorActionListener`.
    pub fn attach_editor_action(entry_ptr: *mut std::os::raw::c_void) {
        if entry_ptr.is_null() {
            return;
        }
        let (class_name, sig) = match ENTRY_LISTENER_CLASSES.get() {
            Some((_, editor)) => (editor.clone(), "(J)V".to_string()),
            None => ("com/corro/CorroEditorAction".to_string(), "(J)V".to_string()),
        };
        let _ = with_env_and_activity(|env, _activity| {
            let entry = unsafe { jni::objects::JObject::from_raw(entry_ptr as jni::sys::jobject) };
            let cls = match load_app_class(env, &class_name) {
                Ok(c) => c,
                Err(_) => {
                    let _ = env.exception_clear();
                    return Ok::<_, Box<dyn StdError + Send + Sync>>(());
                }
            };
            let listener = env.new_object(&cls, &sig, &[(entry_ptr as i64).into()])?;
            crate::backends::android::logcat_rs("editor-action listener created");
            env.call_method(
                &entry,
                "setOnEditorActionListener",
                "(Landroid/widget/TextView$OnEditorActionListener;)V",
                &[(&listener).into()],
            )?;
            // Hardware/adb-injected Enter arrives as a raw key event, never
            // as an editor action: attach the key fallback too. Same class
            // loader, same native callback.
            let key_cls = load_app_class(env, "com/corro/CorroKeyListener").ok();
            if let Some(key_cls) = key_cls {
                crate::backends::android::logcat_rs("key-listener class resolved");
                if let Ok(key_listener) =
                    env.new_object(&key_cls, "(J)V", &[(entry_ptr as i64).into()])
                {
                    let _ = env.call_method(
                        &entry,
                        "setOnKeyListener",
                        "(Landroid/view/View$OnKeyListener;)V",
                        &[(&key_listener).into()],
                    );
                    crate::backends::android::logcat_rs("key-listener attached");
                } else {
                    let _ = env.exception_clear();
                    crate::backends::android::logcat_rs("key-listener new_object FAILED");
                }
            } else {
                crate::backends::android::logcat_rs("key-listener class NOT FOUND");
            }
            Ok::<_, Box<dyn StdError + Send + Sync>>(())
        });
    }

    fn attach_entry_listener(
        entry_ptr: *mut std::os::raw::c_void,
        method: &str,
        class_name: &str,
    ) {
        if entry_ptr.is_null() {
            return;
        }
        let _ = with_env_and_activity(|env, _activity| {
            let entry = unsafe { jni::objects::JObject::from_raw(entry_ptr as jni::sys::jobject) };
            let cls = match load_app_class(env, class_name) {
                Ok(c) => c,
                Err(_) => {
                    let _ = env.exception_clear();
                    return Ok::<_, Box<dyn StdError + Send + Sync>>(());
                }
            };
            let listener = env.new_object(&cls, "(J)V", &[(entry_ptr as i64).into()])?;
            env.call_method(
                &entry,
                method,
                "(Landroid/text/TextWatcher;)V",
                &[(&listener).into()],
            )?;
            Ok::<_, Box<dyn StdError + Send + Sync>>(())
        });
    }

    /// Override the entry listener class names (tests/alternative apps):
    /// (text-watcher class, editor-action class).
    static ENTRY_LISTENER_CLASSES: OnceCell<(String, String)> = OnceCell::new();

    pub fn set_entry_listener_classes(watcher: &str, editor_action: &str) {
        let _ = ENTRY_LISTENER_CLASSES.set((watcher.to_string(), editor_action.to_string()));
    }

    /// `View.invalidate()` on a canvas handle: schedules `onDraw`, which
    /// funnels back into the registered Rust draw closure. Best-effort.
    pub fn invalidate_view(view_ptr: *mut std::os::raw::c_void) {
        if view_ptr.is_null() {
            return;
        }
        let _ = with_env_and_activity(|env, _activity| {
            let view = unsafe { jni::objects::JObject::from_raw(view_ptr as jni::sys::jobject) };
            env.call_method(&view, "invalidate", "()V", &[])?;
            Ok::<_, Box<dyn StdError + Send + Sync>>(())
        });
    }

    /// Attach `child` to a container view (`FrameLayout` scrolled windows,
    /// overlays): `addView(child)`. Best-effort; ignores null handles.
    pub fn attach_child(container_ptr: *mut std::os::raw::c_void, child_ptr: *mut std::os::raw::c_void) {
        if container_ptr.is_null() || child_ptr.is_null() {
            return;
        }
        let _ = with_env_and_activity(|env, _activity| {
            let container =
                unsafe { jni::objects::JObject::from_raw(container_ptr as jni::sys::jobject) };
            let child =
                unsafe { jni::objects::JObject::from_raw(child_ptr as jni::sys::jobject) };
            env.call_method(
                &container,
                "addView",
                "(Landroid/view/View;)V",
                &[(&child).into()],
            )?;
            Ok::<_, Box<dyn StdError + Send + Sync>>(())
        });
    }

    /// Inert `FrameLayout` container backing a scrolled window. Corro drives
    /// its own viewport, so this never scrolls natively — it only keeps the
    /// shared GUI layout code compiling unchanged with valid handles.
    pub fn create_scrolled_view() -> Result<jni::sys::jobject, Box<dyn StdError + Send + Sync>> {
        with_env_and_activity(|env, activity| {
            let ctx = activity.as_obj();
            let layout = env.new_object(
                "android/widget/FrameLayout",
                "(Landroid/content/Context;)V",
                &[(&ctx).into()],
            )?;
            make_global_ref(env, &layout)
        })
    }
}

#[cfg(target_os = "android")]
pub use android_backend::*;

#[cfg(test)]
#[cfg(target_os = "android")]
mod tests {
    use super::*;

    #[test]
    fn test_register_and_invoke_callback() {
        let called = std::cell::Cell::new(false);
        let id = register_callback(Box::new(move || {
            called.set(true);
        }));

        // Manually invoke (simulates Java dispatchCallback)
        invoke_callback(id);

        // Note: in a real test with a real closure we'd need different approach
        // since the closure is consumed by the Box. This test checks the dispatch
        // mechanism doesn't panic.
    }

    #[test]
    fn test_register_and_unregister() {
        let id = register_callback(Box::new(|| {}));
        unregister_callback(id);
        // After unregister, invoke should be a no-op (not panic)
        invoke_callback(id);
    }

    #[test]
    fn test_dispatch_callback() {
        let id = register_callback(Box::new(|| {}));
        dispatch_callback(id);
        unregister_callback(id);
    }

    #[test]
    fn test_multiple_callbacks() {
        let id1 = register_callback(Box::new(|| {}));
        let id2 = register_callback(Box::new(|| {}));
        assert_ne!(id1, id2);
        invoke_callback(id1);
        invoke_callback(id2);
        unregister_callback(id1);
        unregister_callback(id2);
    }

    #[test]
    fn test_is_initialized() {
        // Before init, should be false
        assert!(!is_initialized());
    }
}
