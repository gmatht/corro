//! Android backend adapter: thin JNI wrappers around Android Views.
//!
//! Every widget is a `#[repr(transparent)]` raw `jobject` handle kept alive
//! by a `GlobalRef` in [`crate::backends::android`] (see `KEEP_ALIVE` there).
//! Methods that need the JVM attach the current thread via
//! `with_env_and_activity` and are best-effort no-ops when the backend was
//! never initialised (host builds, tests, headless runs).
//!
//! The Canvas renders via a custom Java `SheetView` (see
//! `rustxWidgets/examples/android`): the draw closure is registered in the
//! Rust-side callback registry and dispatched by id from `SheetView.onDraw`,
//! which funnels every `DrawContext` call back into Rust through JNI. Until
//! a `SheetView` exists, `queue_redraw` is a no-op and `text_extents_*`
//! return a monospace estimate so layout never divides by zero.

#[cfg(target_os = "android")]
mod android_adapter {
    use crate::core::{DrawContext, Error, Widget};
    use jni::objects::JString;
    use std::cell::RefCell;
    use std::collections::HashMap;
    use std::os::raw::c_void;
    use std::sync::Mutex;

    use once_cell::sync::Lazy;

    /// Monospace advance estimate (px) used by `text_extents_*` until real
    /// measurement exists. Matches the `CHAR_W`-family constants callers use.
    const EST_CHAR_W: f64 = 7.2;
    /// Line-height factor for the estimate.
    const EST_LINE_H: f64 = 1.2;

    fn estimate_extents(text: &str, size: f64, weight: i32) -> (f64, f64, f64, f64) {
        let bold = if weight != 0 { 1.08 } else { 1.0 };
        let w = text.chars().count() as f64 * size * (EST_CHAR_W / 12.0) * bold;
        (0.0, 0.0, w, size * EST_LINE_H)
    }

    // ------------------------------------------------------------------
    // Orientation
    // ------------------------------------------------------------------

    /// Layout direction for [`create_box`]. Discriminants match
    /// `android.widget.LinearLayout` orientation constants
    /// (`HORIZONTAL = 0`, `VERTICAL = 1`).
    #[derive(Clone, Copy, Debug, PartialEq, Eq)]
    pub enum Orientation {
        Horizontal = 0,
        Vertical = 1,
    }

    impl Orientation {
        fn as_jni_int(self) -> i32 {
            self as i32
        }
    }

    /// Alias required by `common.rs` (`Orientation as AndroidOrientation`).
    pub type AndroidOrientation = Orientation;

    // ------------------------------------------------------------------
    // Window
    // ------------------------------------------------------------------

    #[repr(transparent)]
    pub struct Window(pub *mut c_void);

    impl Widget for Window {
        fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl AsRef<*mut c_void> for Window {
        fn as_ref(&self) -> &*mut c_void {
            &self.0
        }
    }

    impl Clone for Window {
        fn clone(&self) -> Self {
            Window(self.0)
        }
    }

    impl Window {
        pub fn set_title(&self, _title: &str) {}

        pub fn set_default_size(&self, _w: i32, _h: i32) {}

        /// # Safety
        /// Kept for API compatibility with the GTK backend; no-op on Android.
        pub unsafe fn insert_action_group(&self, _name: &str, _group_ptr: *mut c_void) {}

        pub fn hwnd(&self) -> *mut c_void {
            std::ptr::null_mut()
        }

        pub fn set_child(&self, child: &impl AsRef<*mut c_void>) {
            let child_ptr = *child.as_ref();
            if child_ptr.is_null() {
                return;
            }
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let root = crate::backends::android::root_layout()?;
                let child_obj = unsafe {
                    // SAFETY: child_ptr was obtained from a JNI-created object. We reconstruct
                    // a JObject from the raw pointer only for the duration of this JNI call.
                    jni::objects::JObject::from_raw(child_ptr as jni::sys::jobject)
                };
                env.call_method(
                    root.as_obj(),
                    "addView",
                    "(Landroid/view/View;)V",
                    &[(&child_obj).into()],
                )?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
        }

        pub fn set_child_box(&self, bx: &BoxWidget) {
            self.set_child(bx);
        }

        pub fn present(&self) {}
        pub fn queue_redraw(&self) {}
        pub fn on_event(&self, _cb: Box<dyn FnMut(*mut c_void) -> i32>) {}
        pub fn on_event_key(&self, _cb: Box<dyn FnMut(u32, u32) -> i32>) {}
        pub fn on_close(&self, _cb: Box<dyn FnMut()>) {}
    }

    // ------------------------------------------------------------------
    // Button
    // ------------------------------------------------------------------

    #[repr(transparent)]
    pub struct Button(pub *mut c_void);

    impl Widget for Button {
        fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl AsRef<*mut c_void> for Button {
        fn as_ref(&self) -> &*mut c_void {
            &self.0
        }
    }

    impl Button {
        pub fn on_click(&self, f: impl FnMut() + Send + 'static) -> Result<u64, Error> {
            let id = crate::backends::android::register_callback(Box::new(f));
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                if let Some(listener) = crate::backends::android::try_create_onclick_listener(env, id)
                    .map_err(|e| format!("{e}"))
                    .unwrap_or(None)
                {
                    let btn = unsafe {
                        // SAFETY: self.0 is a raw jobject from create_button(). We reconstruct
                        // it only for this JNI call while the JVM is attached.
                        jni::objects::JObject::from_raw(self.0 as jni::sys::jobject)
                    };
                    env.call_method(
                        &btn,
                        "setOnClickListener",
                        "(Landroid/view/View$OnClickListener;)V",
                        &[(&listener).into()],
                    )?;
                }
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
            Ok(id)
        }

        pub fn emit_clicked(&self) -> Result<u64, Error> {
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let btn = unsafe {
                    // SAFETY: self.0 is a raw jobject from create_button(). We only use it
                    // synchronously within this JNI call while the JVM is attached.
                    jni::objects::JObject::from_raw(self.0 as jni::sys::jobject)
                };
                env.call_method(&btn, "performClick", "()Z", &[])?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
            Ok(0)
        }
    }

    impl Clone for Button {
        fn clone(&self) -> Self {
            Button(self.0)
        }
    }

    // ------------------------------------------------------------------
    // Label
    // ------------------------------------------------------------------

    #[repr(transparent)]
    pub struct Label(pub *mut c_void);

    impl Widget for Label {
        fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl AsRef<*mut c_void> for Label {
        fn as_ref(&self) -> &*mut c_void {
            &self.0
        }
    }

    impl Label {
        pub fn set_text(&self, text: &str) {
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let tv = unsafe {
                    // SAFETY: self.0 is a raw jobject from create_label(). JVM is attached.
                    jni::objects::JObject::from_raw(self.0 as jni::sys::jobject)
                };
                let j_text = env.new_string(text)?;
                env.call_method(
                    &tv,
                    "setText",
                    "(Ljava/lang/CharSequence;)V",
                    &[(&j_text).into()],
                )?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
        }

        pub fn get_text(&self) -> Option<String> {
            let result = crate::backends::android::with_env_and_activity(|env, _activity| {
                let tv = unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let j_value = env.call_method(&tv, "getText", "()Ljava/lang/CharSequence;", &[])?;
                let j_obj_ref = j_value.l()?;
                let j_obj = unsafe { jni::objects::JObject::from_raw(j_obj_ref.as_raw()) };
                let j_str = JString::from(j_obj);
                let text: String = env.get_string(&j_str)?.into();
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(text)
            });
            result.ok()
        }

        pub fn set_visible(&self, visible: bool) {
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let tv = unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let visibility = if visible { 0i32 } else { 8i32 };
                env.call_method(&tv, "setVisibility", "(I)V", &[visibility.into()])?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
        }

        pub fn set_markup(&self, markup: &str) {
            self.set_text(markup);
        }

        pub fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl Clone for Label {
        fn clone(&self) -> Self {
            Label(self.0)
        }
    }

    // ------------------------------------------------------------------
    // BoxWidget
    // ------------------------------------------------------------------

    #[repr(transparent)]
    pub struct BoxWidget(pub *mut c_void);

    impl Widget for BoxWidget {
        fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl AsRef<*mut c_void> for BoxWidget {
        fn as_ref(&self) -> &*mut c_void {
            &self.0
        }
    }

    impl Clone for BoxWidget {
        fn clone(&self) -> Self {
            BoxWidget(self.0)
        }
    }

    impl BoxWidget {
        /// Attach `child` with parent-appropriate `LayoutParams`: children
        /// of a vertical `LinearLayout` fill the width and share leftover
        /// height by weight (weight=1 only for sheet canvases, which must
        /// grow; chrome keeps wrap-content).
        pub fn append(&self, child: &impl AsRef<*mut c_void>) {
            let child_ptr = *child.as_ref();
            if child_ptr.is_null() {
                return;
            }
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let layout = unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let child_obj = unsafe { jni::objects::JObject::from_raw(child_ptr as jni::sys::jobject) };
                // Axis that grows depends on the parent's orientation, which
                // the child cannot know — query it from the layout.
                let is_canvas = crate::backends::android::is_canvas_view(child_ptr);
                let expands = crate::backends::android::is_view_expanding(child_ptr);
                let vertical_parent = env
                    .call_method(&layout, "getOrientation", "()I", &[])
                    .ok()
                    .and_then(|o| o.i().ok())
                    .map(|o| o == 1)
                    .unwrap_or(true);
                let weight = if is_canvas || expands { 1.0f32 } else { 0.0f32 };
                let (child_w, child_h) = if is_canvas {
                    // The sheet fills whatever the parent gives it.
                    (-1i32, 0i32)
                } else if expands {
                    if vertical_parent { (-1i32, 0i32) } else { (0i32, -1i32) }
                } else {
                    (-2i32, -2i32)
                };
                let params = env.new_object(
                    "android/widget/LinearLayout$LayoutParams",
                    "(IIF)V",
                    &[child_w.into(), child_h.into(), weight.into()],
                )?;
                // An empty EditText measures to zero width; give it a floor
                // in px (density-independent ≈ 9px per character at mdpi).
                let min_chars = {
                    let n = crate::backends::android::view_min_chars(child_ptr);
                    // Expanding children (the formula entry) default to a
                    // usable width even when nothing set one explicitly.
                    if n > 0 { n } else if expands { 12 } else { 0 }
                };
                if min_chars > 0 {
                    let density = env
                        .call_method(&layout, "getResources", "()Landroid/content/res/Resources;", &[])
                        .ok()
                        .and_then(|r| r.l().ok())
                        .and_then(|res| {
                            env.call_method(&res, "getDisplayMetrics", "()Landroid/util/DisplayMetrics;", &[])
                                .ok()
                                .and_then(|m| m.l().ok())
                        })
                        .and_then(|metrics| env.get_field(&metrics, "density", "F").ok())
                        .and_then(|f| f.f().ok())
                        .unwrap_or(1.0);
                    let min_px = (min_chars as f32 * 9.0 * density) as i32;
                    let _ = env.call_method(
                        &child_obj,
                        "setMinimumWidth",
                        "(I)V",
                        &[min_px.into()],
                    );
                }
                env.call_method(
                    &layout,
                    "addView",
                    "(Landroid/view/View;Landroid/view/ViewGroup$LayoutParams;)V",
                    &[(&child_obj).into(), (&params).into()],
                )?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
        }

        pub fn set_child_hexpand(&self, _child: &impl AsRef<*mut c_void>, _expand: bool) {}
        pub fn set_child_vexpand(&self, _child: &impl AsRef<*mut c_void>, _expand: bool) {}
        pub fn set_hexpand(&self, _expand: bool) {}
    }

    // ------------------------------------------------------------------
    // Entry
    // ------------------------------------------------------------------

    #[repr(transparent)]
    pub struct Entry(pub *mut c_void);

    impl Widget for Entry {
        fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl AsRef<*mut c_void> for Entry {
        fn as_ref(&self) -> &*mut c_void {
            &self.0
        }
    }

    impl Entry {
        /// Android laid out the entry at zero width (off-screen) because
        /// the box gave every non-canvas child WRAP_CONTENT and an empty
        /// EditText measures 0. Expand flags are no-ops here, so record
        /// the request and let `BoxWidget::append` weight it.
        pub fn set_hexpand(&self, expand: bool) {
            crate::backends::android::set_view_expanding(self.0, expand);
        }

        pub fn set_vexpand(&self, expand: bool) {
            crate::backends::android::set_view_expanding(self.0, expand);
        }

        pub fn set_width_chars(&self, n: i32) {
            crate::backends::android::set_view_min_chars(self.0, n);
        }
        pub fn set_text(&self, text: &str) {
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let edit = unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let j_text = env.new_string(text)?;
                env.call_method(
                    &edit,
                    "setText",
                    "(Ljava/lang/CharSequence;)V",
                    &[(&j_text).into()],
                )?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
        }

        pub fn get_text(&self) -> Option<String> {
            let result = crate::backends::android::with_env_and_activity(|env, _activity| {
                let edit = unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let j_value = env.call_method(&edit, "getText", "()Ljava/lang/CharSequence;", &[])?;
                let j_obj_ref = j_value.l()?;
                let j_obj = unsafe { jni::objects::JObject::from_raw(j_obj_ref.as_raw()) };
                let j_str = JString::from(j_obj);
                let text: String = env.get_string(&j_str)?.into();
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(text)
            });
            result.ok()
        }

        pub fn set_size_request(&self, _w: i32, _h: i32) {}

        pub fn set_visible(&self, _v: bool) {}
        pub fn add_class(&self, _class_name: &str) {}
        pub fn remove_class(&self, _class_name: &str) {}
        pub fn set_halign(&self, _align: i32) {}
        pub fn set_valign(&self, _align: i32) {}
        pub fn set_margin_start(&self, _px: i32) {}
        pub fn set_margin_top(&self, _px: i32) {}

        pub fn connect_activate<F: FnMut(*mut c_void) + 'static>(
            &self,
            f: F,
        ) -> Result<u64, Error> {
            // RETURN from the soft keyboard: the Java side calls
            // `nativeEntryActivate(entryPtr)` from an OnEditorActionListener.
            let mut map = ENTRY_ACTIVATE.lock().unwrap();
            let mut f = f;
            map.insert(
                self.0 as usize,
                SendEntryActivate(Box::into_raw(Box::new(move |p| f(p)))),
            );
            crate::backends::android::attach_editor_action(self.0);
            Ok(0)
        }

        pub fn connect_focus_in_event<F: FnMut(*mut c_void) -> i32 + 'static>(
            &self,
            _f: F,
        ) -> Result<u64, Error> {
            Ok(0)
        }

        pub fn connect_focus_out_event<F: FnMut(*mut c_void) -> i32 + 'static>(
            &self,
            _f: F,
        ) -> Result<u64, Error> {
            Ok(0)
        }

        pub fn grab_focus(&self) {
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let edit = unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                env.call_method(&edit, "requestFocus", "()Z", &[])?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
        }
        /// Android entries report no caret: callers keep their own.
        pub fn get_position(&self) -> Option<usize> {
            None
        }
        pub fn set_position(&self, _pos: usize) {}
        pub fn on_key_raw(&self, _cb: Box<dyn FnMut(u32, u32) -> bool>) {}

    pub fn connect_changed(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
        // Fire `f` on every text change. The Java side calls
        // `nativeEntryChanged(entryPtr)` from a TextWatcher (installed by
        // `attach_text_watcher`); the callback is keyed by the entry's
        // GlobalRef pointer, exactly like the canvas registries.
        let mut map = TEXT_CHANGED.lock().unwrap();
        let mut f = f;
        map.insert(self.0 as usize, SendTextChanged(Box::into_raw(Box::new(move || f()))));
        crate::backends::android::attach_text_watcher(self.0);
        Ok(0)
    }

        /// No focus query on android entries: report false (unchanged).
        pub fn has_focus(&self) -> bool {
            false
        }

        /// No pointer clicks on android entries: accept and never fire.
        pub fn connect_button_press(&self, _f: impl FnMut() + 'static) -> Result<u64, Error> {
            Ok(0)
        }
    }

    impl Clone for Entry {
        fn clone(&self) -> Self {
            Entry(self.0)
        }
    }

    /// Text-change callbacks keyed by entry view pointer.
    static TEXT_CHANGED: Lazy<Mutex<HashMap<usize, SendTextChanged>>> =
        Lazy::new(|| Mutex::new(HashMap::new()));
    /// RETURN (editor action) callbacks keyed by entry view pointer.
    static ENTRY_ACTIVATE: Lazy<Mutex<HashMap<usize, SendEntryActivate>>> =
        Lazy::new(|| Mutex::new(HashMap::new()));

    struct SendTextChanged(*mut dyn FnMut());
    // SAFETY: guarded by the registry mutex; same pattern as the canvas maps.
    unsafe impl Send for SendTextChanged {}
    struct SendEntryActivate(*mut dyn FnMut(*mut c_void));
    // SAFETY: guarded by the registry mutex; same pattern as the canvas maps.
    unsafe impl Send for SendEntryActivate {}

    /// Called from the Java TextWatcher: run the registered change callback.
    pub fn dispatch_text_changed(entry_ptr: *mut c_void) {
        crate::backends::android::logcat_rs("dispatch_text_changed fired");
        let mut map = TEXT_CHANGED.lock().unwrap();
        if let Some(SendTextChanged(ptr)) = map.get_mut(&(entry_ptr as usize)) {
            let cb: &mut dyn FnMut() = unsafe { &mut **ptr };
            cb();
        }
    }

    /// Called from the Java OnEditorActionListener (IME "Done"/Enter).
    pub fn dispatch_entry_activate(entry_ptr: *mut c_void) {
        crate::backends::android::logcat_rs("dispatch_entry_activate fired");
        let mut map = ENTRY_ACTIVATE.lock().unwrap();
        if let Some(SendEntryActivate(ptr)) = map.get_mut(&(entry_ptr as usize)) {
            let cb: &mut dyn FnMut(*mut c_void) = unsafe { &mut **ptr };
            cb(entry_ptr);
        }
    }

    // ------------------------------------------------------------------
    // Canvas + DrawContext
    // ------------------------------------------------------------------
    //
    // Rendering target for the sheet. The Java side owns a `SheetView`
    // (custom `View`) whose `onDraw(Canvas)` funnels each primitive back
    // into Rust through JNI; the Rust side only stores the draw closure
    // and fires it on `queue_redraw`. Until that view exists (host unit
    // tests, uninitialised backend) the closure still runs against an
    // estimating context so `render_to` logic stays testable.

    /// `DrawContext` implementation used when no Java canvas is attached:
    /// records nothing, measures with the monospace estimate.
    pub struct AndroidDrawContext;

    impl DrawContext for AndroidDrawContext {
        fn fill_rect(&mut self, _x: f64, _y: f64, _w: f64, _h: f64, _r: f64, _g: f64, _b: f64, _a: f64) {
        }
        fn stroke_rect(
            &mut self,
            _x: f64,
            _y: f64,
            _w: f64,
            _h: f64,
            _r: f64,
            _g: f64,
            _b: f64,
            _a: f64,
            _lw: f64,
        ) {
        }
        fn draw_text_styled(
            &mut self,
            _x: f64,
            _y: f64,
            _text: &str,
            _font: &str,
            _size: f64,
            _r: f64,
            _g: f64,
            _b: f64,
            _a: f64,
            _slant: i32,
            _weight: i32,
        ) {
        }
        fn text_extents_styled(
            &self,
            text: &str,
            _font: &str,
            size: f64,
            _slant: i32,
            weight: i32,
        ) -> (f64, f64, f64, f64) {
            estimate_extents(text, size, weight)
        }
        fn clear(&mut self, _r: f64, _g: f64, _b: f64, _a: f64) {}
        fn save(&mut self) {}
        fn restore(&mut self) {}
        fn clip(&mut self, _x: f64, _y: f64, _w: f64, _h: f64) {}
    }

    #[repr(transparent)]
    pub struct Canvas(pub *mut c_void);

    impl Widget for Canvas {
        fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl AsRef<*mut c_void> for Canvas {
        fn as_ref(&self) -> &*mut c_void {
            &self.0
        }
    }

    impl Clone for Canvas {
        fn clone(&self) -> Self {
            Canvas(self.0)
        }
    }

    impl Canvas {
        /// Canvas id for this view (0 = unknown; registry lookups miss).
        fn canvas_id(&self) -> u64 {
            crate::backends::android::canvas_id_for_view(self.0)
        }

        pub fn set_draw_callback(
            &self,
            cb: Box<dyn FnMut(&mut dyn DrawContext, i32, i32)>,
        ) {
            // Stash the closure under our canvas id; the Java SheetView
            // dispatches to it by id from onDraw.
            {
                let mut map = DRAW_CALLBACKS.lock().unwrap();
                map.insert(
                    self.canvas_id(),
                    SendDrawCallback(Box::into_raw(cb)),
                );
            }
            self.queue_redraw();
        }

        pub fn queue_redraw(&self) {
            // Schedule a real onDraw through the Java view (which replays
            // the closure with a JNI-backed context at the live size), and
            // also run it immediately against the estimating context so
            // headless/test callers observe draws without Java.
            crate::backends::android::invalidate_view(self.0);
            let mut map = DRAW_CALLBACKS.lock().unwrap();
            if let Some(SendDrawCallback(ptr)) = map.get_mut(&self.canvas_id()) {
                let cb: &mut dyn FnMut(&mut dyn DrawContext, i32, i32) = unsafe { &mut **ptr };
                let mut dc = AndroidDrawContext;
                let (w, h) = self.replay_size();
                cb(&mut dc, w, h);
            }
        }

        pub fn set_size_request(&self, w: i32, h: i32) {
            CANVAS_SIZE_REQUEST
                .lock()
                .unwrap()
                .insert(self.canvas_id(), (w.max(1), h.max(1)));
        }

        /// Size to replay the draw closure at: the real laid-out size once
        /// `onDraw` has reported one, else the requested size (Android has
        /// not measured the view yet), else a conservative default.
        fn replay_size(&self) -> (i32, i32) {
            let id = self.canvas_id();
            if let Some(&(w, h)) = CANVAS_SIZE.lock().unwrap().get(&id) {
                return (w, h);
            }
            CANVAS_SIZE_REQUEST
                .lock()
                .unwrap()
                .get(&id)
                .copied()
                .unwrap_or((800, 600))
        }

        pub fn set_content_size(&self, w: i32, h: i32) {
            self.set_size_request(w, h);
        }

        pub fn set_visible(&self, _v: bool) {}

        pub fn grab_focus(&self) {}
        pub fn set_can_focus(&self, _can: bool) {}

        pub fn on_click(&self, cb: Box<dyn FnMut(f64, f64)>) {
            let mut map = CLICK_CALLBACKS.lock().unwrap();
            map.insert(self.canvas_id(), SendClickCallback(Box::into_raw(cb)));
        }

        pub fn on_key(&self, cb: Box<dyn FnMut(u32) -> bool>) {
            let mut boxed = cb;
            self.on_key_raw(Box::new(move |k: u32, _s: u32| -> bool { boxed(k) }));
        }

        pub fn on_key_raw(&self, cb: Box<dyn FnMut(u32, u32) -> bool>) {
            let mut map = KEY_CALLBACKS.lock().unwrap();
            map.insert(self.canvas_id(), SendKeyCallback(Box::into_raw(cb)));
        }

        /// Force an immediate draw. Replays the closure like
        /// [`Canvas::queue_redraw`] (plus a view invalidate).
        pub fn force_draw(&self, _window_ptr: *mut c_void, fallback_w: i32, fallback_h: i32) {
            crate::backends::android::invalidate_view(self.0);
            let mut map = DRAW_CALLBACKS.lock().unwrap();
            if let Some(SendDrawCallback(ptr)) = map.get_mut(&self.canvas_id()) {
                let cb: &mut dyn FnMut(&mut dyn DrawContext, i32, i32) = unsafe { &mut **ptr };
                let mut dc = AndroidDrawContext;
                // Prefer the live laid-out size; fall back to the caller's
                // hint only when the view has never been drawn.
                let (w, h) = match CANVAS_SIZE.lock().unwrap().get(&self.canvas_id()).copied() {
                    Some(size) => size,
                    None => (fallback_w.max(1), fallback_h.max(1)),
                };
                cb(&mut dc, w, h);
            }
        }
    }

    static DRAW_CALLBACKS: Lazy<Mutex<HashMap<u64, SendDrawCallback>>> =
        Lazy::new(|| Mutex::new(HashMap::new()));
    static CLICK_CALLBACKS: Lazy<Mutex<HashMap<u64, SendClickCallback>>> =
        Lazy::new(|| Mutex::new(HashMap::new()));

    static KEY_CALLBACKS: Lazy<Mutex<HashMap<u64, SendKeyCallback>>> =
        Lazy::new(|| Mutex::new(HashMap::new()));
    /// The view's *laid-out* size in pixels, learned from `View.onDraw`
    /// (`dispatch_draw`). This is the size the draw closure must be replayed
    /// at: Android does the layout, so Rust cannot know it any earlier.
    static CANVAS_SIZE: Lazy<Mutex<HashMap<u64, (i32, i32)>>> =
        Lazy::new(|| Mutex::new(HashMap::new()));
    /// The size Rust *asked* for via `set_size_request` (often a 1x1
    /// placeholder, since Android measures children itself). Kept separate
    /// from [`CANVAS_SIZE`] so a placeholder request can never overwrite a
    /// real laid-out size and shrink the viewport to a single row.
    static CANVAS_SIZE_REQUEST: Lazy<Mutex<HashMap<u64, (i32, i32)>>> =
        Lazy::new(|| Mutex::new(HashMap::new()));

    struct SendDrawCallback(*mut dyn FnMut(&mut dyn DrawContext, i32, i32));
    // SAFETY: guarded by the registry mutex; same pattern as the wasm adapter.
    unsafe impl Send for SendDrawCallback {}
    struct SendClickCallback(*mut dyn FnMut(f64, f64));
    // SAFETY: guarded by the registry mutex; same pattern as the wasm adapter.
    unsafe impl Send for SendClickCallback {}
    struct SendKeyCallback(*mut dyn FnMut(u32, u32) -> bool);
    // SAFETY: guarded by the registry mutex; same pattern as the wasm adapter.
    unsafe impl Send for SendKeyCallback {}

    /// Dispatch a tap from Java `SheetView` to the registered click closure.
    pub fn dispatch_canvas_click(canvas_id: u64, x: f64, y: f64) {
        let mut map = CLICK_CALLBACKS.lock().unwrap();
        if let Some(SendClickCallback(ptr)) = map.get_mut(&canvas_id) {
            let cb: &mut dyn FnMut(f64, f64) = unsafe { &mut **ptr };
            cb(x, y);
        }
    }

    /// Dispatch a key from Java to the registered key closure.
    pub fn dispatch_canvas_key(canvas_id: u64, keyval: u32, mods: u32) -> bool {
        let mut map = KEY_CALLBACKS.lock().unwrap();
        if let Some(SendKeyCallback(ptr)) = map.get_mut(&canvas_id) {
            let cb: &mut dyn FnMut(u32, u32) -> bool = unsafe { &mut **ptr };
            cb(keyval, mods)
        } else {
            false
        }
    }

    /// Replay the registered draw closure for `canvas_id` against a live
    /// Java `android.graphics.Canvas`. Called from the cdylib's
    /// `SheetView_nativeOnDraw` export (which runs on the UI thread inside
    /// `View.onDraw`). `w`/`h` are the view's pixel size.
    pub fn dispatch_draw<'a, 'b, 'c>(
        canvas_id: u64,
        canvas_obj: jni::objects::JObject<'c>,
        env: &'a mut jni::JNIEnv<'b>,
        w: i32,
        h: i32,
    ) {
        CANVAS_SIZE.lock().unwrap().insert(canvas_id, (w.max(1), h.max(1)));
        // NLL-friendly: resolve the raw callback pointer first, release the
        // registry lock, then build the JNI context and invoke. Holding the
        // lock across JNI calls risks re-entrant deadlock (queue_redraw
        // from inside the closure would block on it).
        let raw = {
            let map = DRAW_CALLBACKS.lock().unwrap();
            map.get(&canvas_id).map(|s| s.0)
        };
        let Some(raw) = raw else {
            return;
        };
        match JniDrawContext::new(env, &canvas_obj) {
            Some(mut ctx) => {
                let cb: &mut dyn FnMut(&mut dyn DrawContext, i32, i32) =
                    unsafe { &mut *raw };
                cb(&mut ctx, w, h);
            }
            None => {
                let mut est = AndroidDrawContext;
                let cb: &mut dyn FnMut(&mut dyn DrawContext, i32, i32) =
                    unsafe { &mut *raw };
                cb(&mut est, w, h);
            }
        }
    }

    /// `DrawContext` backed by a live `android.graphics.Canvas` via JNI.
    /// One instance serves a single `onDraw`: paints are created up front
    /// (fill, stroke, text) and reused across primitives.
    pub struct JniDrawContext<'a, 'b> {
        // Interior mutability: `DrawContext::text_extents_styled` only gets
        // `&self`, but every JNI call needs `&mut JNIEnv`. The context is
        // driven exclusively from the UI thread inside `onDraw`, so the
        // cell never contends; every use goes through `try_borrow_mut` and
        // degrades gracefully instead of panicking.
        env: RefCell<&'a mut jni::JNIEnv<'b>>,
        canvas: jni::objects::GlobalRef,
        fill_paint: jni::objects::GlobalRef,
        stroke_paint: jni::objects::GlobalRef,
        text_paint: jni::objects::GlobalRef,
    }

    impl<'a, 'b> JniDrawContext<'a, 'b> {
        pub fn new(
            env: &'a mut jni::JNIEnv<'b>,
            canvas: &jni::objects::JObject<'_>,
        ) -> Option<Self> {
            let canvas = env.new_global_ref(canvas).ok()?;
            let fill_paint = Self::make_paint(env, 0 /* FILL */)?;
            let stroke_paint = Self::make_paint(env, 1 /* STROKE */)?;
            let text_paint = Self::make_paint(env, 0 /* FILL */)?;
            // Text draws anti-aliased; shapes stay crisp.
            let _ = env.call_method(
                text_paint.as_obj(),
                "setAntiAlias",
                "(Z)V",
                &[true.into()],
            );
            Some(JniDrawContext { env: RefCell::new(env), canvas, fill_paint, stroke_paint, text_paint })
        }

        fn make_paint(
            env: &mut jni::JNIEnv<'_>,
            style: i32,
        ) -> Option<jni::objects::GlobalRef> {
            // 0 = Paint.Style.FILL, 1 = STROKE (ordinal into Style.values()).
            let paint = env
                .new_object("android/graphics/Paint", "()V", &[])
                .ok()?;
            let styles = env
                .call_static_method(
                    "android/graphics/Paint$Style",
                    "values",
                    "()[Landroid/graphics/Paint$Style;",
                    &[],
                )
                .ok()?
                .l()
                .ok()?;
            let styles: jni::objects::JObjectArray =
                jni::objects::JObjectArray::from(styles);
            let style_obj = env.get_object_array_element(&styles, style).ok()?;
            env.call_method(&paint, "setStyle", "(Landroid/graphics/Paint$Style;)V", &[(&style_obj).into()])
                .ok()?;
            env.new_global_ref(&paint).ok()
        }

        fn argb(a: f64, r: f64, g: f64, b: f64) -> i32 {
            let clamp = |v: f64| (v.clamp(0.0, 1.0) * 255.0) as i32;
            (clamp(a) << 24) | (clamp(r) << 16) | (clamp(g) << 8) | clamp(b)
        }

        fn set_paint_color(
            env: &mut jni::JNIEnv<'_>,
            paint: &jni::objects::GlobalRef,
            a: f64,
            r: f64,
            g: f64,
            b: f64,
        ) {
            let _ = env.call_method(
                paint.as_obj(),
                "setColor",
                "(I)V",
                &[Self::argb(a, r, g, b).into()],
            );
        }

        fn set_text_size(
            env: &mut jni::JNIEnv<'_>,
            paint: &jni::objects::GlobalRef,
            size: f64,
            weight: i32,
        ) {
            let _ = env.call_method(
                paint.as_obj(),
                "setTextSize",
                "(F)V",
                &[(size.max(1.0) as f32).into()],
            );
            let _ = env.call_method(
                paint.as_obj(),
                "setFakeBoldText",
                "(Z)V",
                &[(weight != 0).into()],
            );
            let skew = if weight != 0 { 0.0f32 } else { 0.0f32 };
            let _ = env.call_method(
                paint.as_obj(),
                "setTextSkewX",
                "(F)V",
                &[skew.into()],
            );
        }

        /// Real text measurement via `Paint.measureText` (+ ascent/descent
        /// for the height), honouring the same size/weight the draw path
        /// uses. Returns the `(x_bearing, y_bearing, width, height)` tuple
        /// the trait promises: bearings are 0 (left/top origin, matching
        /// the monospace estimate callers already assume).
        fn measure_text(
            env: &mut jni::JNIEnv<'_>,
            paint: &jni::objects::GlobalRef,
            text: &str,
            size: f64,
            weight: i32,
        ) -> Option<(f64, f64, f64, f64)> {
            if text.is_empty() {
                return Some((0.0, 0.0, 0.0, 0.0));
            }
            Self::set_text_size(env, paint, size, weight);
            let jtext = env.new_string(text).ok()?;
            let w = env
                .call_method(
                    paint.as_obj(),
                    "measureText",
                    "(Ljava/lang/String;)F",
                    &[(&jtext).into()],
                )
                .ok()?
                .f()
                .ok()? as f64;
            let ascent = env
                .call_method(paint.as_obj(), "ascent", "()F", &[])
                .ok()?
                .f()
                .ok()? as f64;
            let descent = env
                .call_method(paint.as_obj(), "descent", "()F", &[])
                .ok()?
                .f()
                .ok()? as f64;
            Some((0.0, 0.0, w.max(0.0), (descent - ascent).max(0.0)))
        }
    }

    impl<'a, 'b> DrawContext for JniDrawContext<'a, 'b> {
        fn fill_rect(&mut self, x: f64, y: f64, w: f64, h: f64, r: f64, g: f64, b: f64, a: f64) {
            let Ok(mut env) = self.env.try_borrow_mut() else { return; };
            Self::set_paint_color(&mut env, &self.fill_paint, a, r, g, b);
            let (env, canvas, paint) = (&mut *env, &self.canvas, &self.fill_paint);
            let _ = env.call_method(
                canvas.as_obj(),
                "drawRect",
                "(FFFFLandroid/graphics/Paint;)V",
                &[(x as f32).into(), (y as f32).into(), ((x + w) as f32).into(), ((y + h) as f32).into(), (&paint.as_obj()).into()],
            );
        }

        fn stroke_rect(
            &mut self,
            x: f64,
            y: f64,
            w: f64,
            h: f64,
            r: f64,
            g: f64,
            b: f64,
            a: f64,
            lw: f64,
        ) {
            let Ok(mut env) = self.env.try_borrow_mut() else { return; };
            Self::set_paint_color(&mut env, &self.stroke_paint, a, r, g, b);
            let (env, canvas, paint) = (&mut *env, &self.canvas, &self.stroke_paint);
            let _ = env.call_method(
                paint.as_obj(),
                "setStrokeWidth",
                "(F)V",
                &[(lw.max(0.5) as f32).into()],
            );
            let _ = env.call_method(
                canvas.as_obj(),
                "drawRect",
                "(FFFFLandroid/graphics/Paint;)V",
                &[(x as f32).into(), (y as f32).into(), ((x + w) as f32).into(), ((y + h) as f32).into(), (&paint.as_obj()).into()],
            );
        }

        fn draw_text_styled(
            &mut self,
            x: f64,
            y: f64,
            text: &str,
            _font: &str,
            size: f64,
            r: f64,
            g: f64,
            b: f64,
            a: f64,
            _slant: i32,
            weight: i32,
        ) {
            if text.is_empty() {
                return;
            }
            let Ok(mut env) = self.env.try_borrow_mut() else { return; };
            Self::set_paint_color(&mut env, &self.text_paint, a, r, g, b);
            Self::set_text_size(&mut env, &self.text_paint, size, weight);
            let (env, canvas, paint) = (&mut *env, &self.canvas, &self.text_paint);
            let Ok(jtext) = env.new_string(text) else {
                return;
            };
            // drawText draws with the baseline at y; offset by ascent so
            // callers' top-left convention matches other backends.
            let ascent: f64 = env
                .call_method(paint.as_obj(), "ascent", "()F", &[])
                .map(|v| v.f().unwrap_or(0.0) as f64)
                .unwrap_or(0.0);
            let _ = env.call_method(
                canvas.as_obj(),
                "drawText",
                "(Ljava/lang/String;FFLandroid/graphics/Paint;)V",
                &[(&jtext).into(), (x as f32).into(), ((y - ascent) as f32).into(), (&paint.as_obj()).into()],
            );
        }

        fn text_extents_styled(
            &self,
            text: &str,
            _font: &str,
            size: f64,
            _slant: i32,
            weight: i32,
        ) -> (f64, f64, f64, f64) {
            // Real measurement via Paint.measureText. `&self` is enough:
            // the JNIEnv lives behind a RefCell (UI thread only, so the
            // borrow cannot contend). Any failure degrades to the monospace
            // estimate so layout never divides by zero.
            let Ok(mut env) = self.env.try_borrow_mut() else {
                return estimate_extents(text, size, weight);
            };
            let measured = Self::measure_text(&mut env, &self.text_paint, text, size, weight);
            drop(env);
            measured.unwrap_or_else(|| estimate_extents(text, size, weight))
        }

        fn clear(&mut self, r: f64, g: f64, b: f64, a: f64) {
            let Ok(mut env) = self.env.try_borrow_mut() else { return; };
            let (env, canvas) = (&mut *env, &self.canvas);
            let _ = env.call_method(
                canvas.as_obj(),
                "drawColor",
                "(I)V",
                &[Self::argb(a, r, g, b).into()],
            );
        }

        fn save(&mut self) {
            let Ok(mut env) = self.env.try_borrow_mut() else { return; };
            let (env, canvas) = (&mut *env, &self.canvas);
            let _ = env.call_method(canvas.as_obj(), "save", "()I", &[]);
        }

        fn restore(&mut self) {
            let Ok(mut env) = self.env.try_borrow_mut() else { return; };
            let (env, canvas) = (&mut *env, &self.canvas);
            let _ = env.call_method(canvas.as_obj(), "restore", "()V", &[]);
        }

        fn clip(&mut self, x: f64, y: f64, w: f64, h: f64) {
            let Ok(mut env) = self.env.try_borrow_mut() else { return; };
            let (env, canvas) = (&mut *env, &self.canvas);
            let _ = env.call_method(
                canvas.as_obj(),
                "clipRect",
                "(FFFF)Z",
                &[(x as f32).into(), (y as f32).into(), ((x + w) as f32).into(), ((y + h) as f32).into()],
            );
        }
    }

    // ------------------------------------------------------------------
    // Menu / MenuBar / SimpleAction
    // ------------------------------------------------------------------
    //
    // Android has no desktop menu bar; corro renders its own menu UI on the
    // Canvas (see `MenuBar::handle_menu_key`). These types therefore carry
    // the menu *model* (labels + action names) plus the shared action
    // registry, exactly like the wasm adapter.

    enum MenuItem {
        Item { label: String, action: String },
        Submenu { label: String, items: Vec<MenuItem> },
    }

    impl Clone for MenuItem {
        fn clone(&self) -> Self {
            match self {
                MenuItem::Item { label, action } => MenuItem::Item {
                    label: label.clone(),
                    action: action.clone(),
                },
                MenuItem::Submenu { label, items } => MenuItem::Submenu {
                    label: label.clone(),
                    items: items.clone(),
                },
            }
        }
    }

    #[derive(Default)]
    pub struct Menu {
        items: Vec<MenuItem>,
    }

    impl Clone for Menu {
        fn clone(&self) -> Self {
            Menu { items: self.items.clone() }
        }
    }

    impl Widget for Menu {
        fn raw_handle(&self) -> *mut c_void {
            std::ptr::null_mut()
        }
    }

    impl AsRef<*mut c_void> for Menu {
        fn as_ref(&self) -> &*mut c_void {
            // No stable handle; callers only pass menus back into
            // `append_submenu` / `create_menubar`, which read the model.
            unsafe { &*(&std::ptr::null_mut() as *const *mut c_void) }
        }
    }

    impl Menu {
        pub fn append(&mut self, label: &str, detailed_action: &str) {
            let action = detailed_action.split('(').next().unwrap_or(detailed_action).to_owned();
            self.items.push(MenuItem::Item { label: label.to_owned(), action });
        }

        pub fn append_submenu(&mut self, label: &str, submenu: &Menu) {
            self.items.push(MenuItem::Submenu {
                label: label.to_owned(),
                items: submenu.items.clone(),
            });
        }
    }

    pub struct MenuBar {
        items: Vec<MenuItem>,
    }

    impl Clone for MenuBar {
        fn clone(&self) -> Self {
            MenuBar { items: self.items.clone() }
        }
    }

    impl Widget for MenuBar {
        fn raw_handle(&self) -> *mut c_void {
            std::ptr::null_mut()
        }
    }

    impl AsRef<*mut c_void> for MenuBar {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&std::ptr::null_mut() as *const *mut c_void) }
        }
    }

    impl MenuBar {
        pub fn activate_submenu_by_mnemonic(&self, _keyval: u32) -> bool {
            false
        }
        pub fn activate_submenu_item_by_mnemonic(&self, _keyval: u32) -> bool {
            false
        }
        /// # Safety
        /// Kept for API compatibility; no-op on Android.
        pub unsafe fn insert_action_group(&self, _name: &str, _group_ptr: *mut c_void) {}
        pub fn handle_mnemonic_key(&self, _keyval: u32) -> bool {
            false
        }
        pub fn handle_menu_key(&self, _keyval: u32, _modifiers: u32) -> bool {
            false
        }
        pub fn menu_active(&self) -> bool {
            false
        }
        pub fn menu_close(&self) {}
    }

    #[derive(Clone)]
    pub struct SimpleAction {
        name: String,
    }

    impl Widget for SimpleAction {
        fn raw_handle(&self) -> *mut c_void {
            std::ptr::null_mut()
        }
    }

    impl AsRef<*mut c_void> for SimpleAction {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&std::ptr::null_mut() as *const *mut c_void) }
        }
    }

    impl SimpleAction {
        pub fn connect_activate<F: FnMut(*mut c_void) + 'static>(&self, f: F) -> Result<u64, Error> {
            register_action(&self.name, Box::new(f));
            Ok(0)
        }
    }

    struct SendFnPtr(*mut dyn FnMut(*mut c_void));
    // SAFETY: guarded by the registry mutex; same pattern as the wasm adapter.
    unsafe impl Send for SendFnPtr {}

    static ACTION_REGISTRY: Lazy<Mutex<HashMap<String, SendFnPtr>>> =
        Lazy::new(|| Mutex::new(HashMap::new()));

    fn register_action(name: &str, f: Box<dyn FnMut(*mut c_void)>) {
        let mut map = ACTION_REGISTRY.lock().unwrap();
        let ptr = Box::into_raw(Box::new(f) as Box<dyn FnMut(*mut c_void)>);
        map.insert(name.to_owned(), SendFnPtr(ptr));
    }

    /// Dispatch a menu action by name (called from `handle_menu_action`
    /// plumbing and future Java menu shims).
    pub fn invoke_action(name: &str, param: *mut c_void) {
        if let Ok(mut map) = ACTION_REGISTRY.lock() {
            if let Some(SendFnPtr(ptr)) = map.get_mut(name) {
                let cb: &mut dyn FnMut(*mut c_void) = unsafe { &mut **ptr };
                cb(param);
            }
        }
    }

    // ------------------------------------------------------------------
    // Grid
    // ------------------------------------------------------------------

    #[repr(transparent)]
    pub struct Grid(pub *mut c_void);

    impl Widget for Grid {
        fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl AsRef<*mut c_void> for Grid {
        fn as_ref(&self) -> &*mut c_void {
            &self.0
        }
    }

    impl Clone for Grid {
        fn clone(&self) -> Self {
            Grid(self.0)
        }
    }

    impl Grid {
        pub fn attach(
            &self,
            child: &impl AsRef<*mut c_void>,
            _left: i32,
            _top: i32,
            _width: i32,
            _height: i32,
        ) {
            let child_ptr = *child.as_ref();
            if child_ptr.is_null() {
                return;
            }
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let grid = unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let child_obj =
                    unsafe { jni::objects::JObject::from_raw(child_ptr as jni::sys::jobject) };
                env.call_method(
                    &grid,
                    "addView",
                    "(Landroid/view/View;)V",
                    &[(&child_obj).into()],
                )?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
        }
    }

    // ------------------------------------------------------------------
    // DropDown
    // ------------------------------------------------------------------

    #[repr(transparent)]
    pub struct DropDown(pub *mut c_void);

    impl Widget for DropDown {
        fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl AsRef<*mut c_void> for DropDown {
        fn as_ref(&self) -> &*mut c_void {
            &self.0
        }
    }

    impl DropDown {
        pub fn set_active(&self, index: Option<u32>) {
            if let Some(idx) = index {
                let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                    let spinner =
                        unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                    env.call_method(&spinner, "setSelection", "(I)V", &[(idx as i32).into()])?;
                    Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
                });
            }
        }

        pub fn get_active(&self) -> i32 {
            let r = crate::backends::android::with_env_and_activity(|env, _activity| {
                let spinner =
                    unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let pos =
                    env.call_method(&spinner, "getSelectedItemPosition", "()I", &[])?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(pos.i()?)
            });
            r.unwrap_or(-1)
        }

        pub fn connect_changed(&self, _f: impl FnMut() + 'static) -> Result<u64, Error> {
            Ok(0) // TODO: OnItemSelectedListener via JNI trampoline
        }

        pub fn set_hexpand(&self, _expand: bool) {}
        pub fn set_vexpand(&self, _expand: bool) {}
        pub fn grab_focus(&self) {}
    }

    impl Clone for DropDown {
        fn clone(&self) -> Self {
            DropDown(self.0)
        }
    }

    // ------------------------------------------------------------------
    // CheckButton
    // ------------------------------------------------------------------

    #[repr(transparent)]
    pub struct CheckButton(pub *mut c_void);

    impl Widget for CheckButton {
        fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl AsRef<*mut c_void> for CheckButton {
        fn as_ref(&self) -> &*mut c_void {
            &self.0
        }
    }

    impl CheckButton {
        pub fn is_active(&self) -> bool {
            let r = crate::backends::android::with_env_and_activity(|env, _activity| {
                let cb = unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let checked = env.call_method(&cb, "isChecked", "()Z", &[])?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(checked.z()?)
            });
            r.unwrap_or(false)
        }

        pub fn set_active(&self, active: bool) {
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let cb = unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                env.call_method(&cb, "setChecked", "(Z)V", &[active.into()])?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
        }

        pub fn connect_toggled(&self, _f: impl FnMut() + 'static) -> Result<u64, Error> {
            Ok(0) // TODO: OnCheckedChangeListener via JNI trampoline
        }
    }

    impl Clone for CheckButton {
        fn clone(&self) -> Self {
            CheckButton(self.0)
        }
    }

    // ------------------------------------------------------------------
    // RadioButton
    // ------------------------------------------------------------------

    #[repr(transparent)]
    pub struct RadioButton(pub *mut c_void);

    impl Widget for RadioButton {
        fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl AsRef<*mut c_void> for RadioButton {
        fn as_ref(&self) -> &*mut c_void {
            &self.0
        }
    }

    impl RadioButton {
        pub fn is_active(&self) -> bool {
            let r = crate::backends::android::with_env_and_activity(|env, _activity| {
                let rb = unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let checked = env.call_method(&rb, "isChecked", "()Z", &[])?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(checked.z()?)
            });
            r.unwrap_or(false)
        }

        pub fn set_active(&self, active: bool) {
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let rb = unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                env.call_method(&rb, "setChecked", "(Z)V", &[active.into()])?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
        }

        pub fn connect_toggled(&self, _f: impl FnMut() + 'static) -> Result<u64, Error> {
            Ok(0) // TODO: OnCheckedChangeListener via JNI trampoline
        }
    }

    impl Clone for RadioButton {
        fn clone(&self) -> Self {
            RadioButton(self.0)
        }
    }

    // ------------------------------------------------------------------
    // Dialog
    // ------------------------------------------------------------------

    #[repr(transparent)]
    pub struct Dialog(pub *mut c_void);

    impl Widget for Dialog {
        fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl AsRef<*mut c_void> for Dialog {
        fn as_ref(&self) -> &*mut c_void {
            &self.0
        }
    }

    impl Clone for Dialog {
        fn clone(&self) -> Self {
            Dialog(self.0)
        }
    }

    impl Dialog {
        pub fn set_title(&self, title: &str) {
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let builder =
                    unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let j_title = env.new_string(title)?;
                env.call_method(
                    &builder,
                    "setTitle",
                    "(Ljava/lang/CharSequence;)Landroid/app/AlertDialog$Builder;",
                    &[(&j_title).into()],
                )?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
        }

        pub fn set_default_size(&self, _w: i32, _h: i32) {}

        /// Parent this dialog to `parent`. No-op on Android (the activity
        /// context already parents dialogs).
        pub fn set_transient_for(&self, _parent: *mut c_void) {}

        pub fn append_content_area(&self, _child: &impl AsRef<*mut c_void>) {}

        pub fn add_button(&self, text: &str, _response_id: i32) {
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let builder =
                    unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let j_text = env.new_string(text)?;
                env.call_method(
                    &builder,
                    "setPositiveButton",
                    "(Ljava/lang/CharSequence;Landroid/content/DialogInterface$OnClickListener;)Landroid/app/AlertDialog$Builder;",
                    &[(&j_text).into(), (&jni::objects::JObject::null()).into()],
                )?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
        }

        pub fn present(&self) {
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let builder =
                    unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let dialog =
                    env.call_method(&builder, "create", "()Landroid/app/AlertDialog;", &[])?;
                env.call_method(dialog.l()?, "show", "()V", &[])?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
        }

        pub fn connect_response(&self, _f: impl FnMut(i32) + 'static) -> Result<u64, Error> {
            Ok(0) // TODO: DialogInterface.OnClickListener via JNI trampoline
        }

        /// Default focused button. No-op until response trampolines land.
        pub fn set_default_response(&self, _response_id: i32) {}

        pub fn close(&self) {}
    }

    // ------------------------------------------------------------------
    // TextView
    // ------------------------------------------------------------------

    #[repr(transparent)]
    pub struct TextView(pub *mut c_void);

    impl Widget for TextView {
        fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl AsRef<*mut c_void> for TextView {
        fn as_ref(&self) -> &*mut c_void {
            &self.0
        }
    }

    impl TextView {
        pub fn set_text(&self, text: &str) {
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let tv = unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let j_text = env.new_string(text)?;
                env.call_method(
                    &tv,
                    "setText",
                    "(Ljava/lang/CharSequence;)V",
                    &[(&j_text).into()],
                )?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
        }

        pub fn get_text(&self) -> Option<String> {
            let r = crate::backends::android::with_env_and_activity(|env, _activity| {
                let tv = unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let j_value =
                    env.call_method(&tv, "getText", "()Ljava/lang/CharSequence;", &[])?;
                let j_obj_ref = j_value.l()?;
                let j_obj = unsafe { jni::objects::JObject::from_raw(j_obj_ref.as_raw()) };
                let j_str = jni::objects::JString::from(j_obj);
                let text: String = env.get_string(&j_str)?.into();
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(text)
            });
            r.ok()
        }

        pub fn set_wrap_mode(&self, _wrap_mode: i32) {}

        pub fn set_size_request(&self, _w: i32, _h: i32) {}

        pub fn set_hexpand(&self, _expand: bool) {}

        pub fn set_vexpand(&self, _expand: bool) {}
    }

    impl Clone for TextView {
        fn clone(&self) -> Self {
            TextView(self.0)
        }
    }

    // ------------------------------------------------------------------
    // ScrolledWindow + Overlay
    // ------------------------------------------------------------------
    //
    // Android scrolls natively (ScrollView); corro drives the viewport
    // itself, so these are inert containers that keep shared GUI code
    // compiling unchanged.

    #[repr(transparent)]
    #[derive(Clone, Copy)]
    pub struct ScrolledWindow(pub *mut c_void);

    impl Widget for ScrolledWindow {
        fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl ScrolledWindow {
        /// Attach the sheet canvas to this container (the one real
        /// parent-child edge Android needs: scrolled window -> canvas).
        pub fn attach_canvas(&self, canvas: &Canvas) {
            crate::backends::android::attach_child(self.0, canvas.0);
        }
    }

    impl AsRef<*mut c_void> for ScrolledWindow {
        fn as_ref(&self) -> &*mut c_void {
            &self.0
        }
    }

    impl Clone for Overlay {
        fn clone(&self) -> Self {
            Overlay(self.0)
        }
    }

    #[repr(transparent)]
    pub struct Overlay(pub *mut c_void);

    impl Widget for Overlay {
        fn raw_handle(&self) -> *mut c_void {
            self.0
        }
    }

    impl AsRef<*mut c_void> for Overlay {
        fn as_ref(&self) -> &*mut c_void {
            &self.0
        }
    }

    impl Overlay {
        pub fn set_child(&self, child: &impl AsRef<*mut c_void>) {
            let child_ptr = *child.as_ref();
            if child_ptr.is_null() || self.0.is_null() {
                return;
            }
            let parent = self.0;
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let layout =
                    unsafe { jni::objects::JObject::from_raw(parent as jni::sys::jobject) };
                let child_obj =
                    unsafe { jni::objects::JObject::from_raw(child_ptr as jni::sys::jobject) };
                env.call_method(
                    &layout,
                    "addView",
                    "(Landroid/view/View;)V",
                    &[(&child_obj).into()],
                )?;
                Ok::<_, Box<dyn std::error::Error + Send + Sync>>(())
            });
        }

        pub fn add_overlay(&self, child: &impl AsRef<*mut c_void>) {
            self.set_child(child);
        }

        pub fn set_overlay_pass_through(&self, _child: &impl AsRef<*mut c_void>, _pass: bool) {
        }

        pub fn remove(&self, _child: &impl AsRef<*mut c_void>) {}

        pub fn show_all(&self) {}

        pub fn set_size_request(&self, _w: i32, _h: i32) {}

        pub fn set_vexpand(&self, _expand: bool) {}

        pub fn set_hexpand(&self, _expand: bool) {}
    }

    // ------------------------------------------------------------------
    // Factories
    // ------------------------------------------------------------------

    pub fn create_window() -> Result<Window, Error> {
        let ptr = crate::backends::android::create_window()
            .map_err(|e| Error::Backend(format!("{e}")))?;
        Ok(Window(ptr as *mut c_void))
    }

    pub fn create_button(label: &str) -> Result<Button, Error> {
        let ptr = crate::backends::android::create_button(label)
            .map_err(|e| Error::Backend(format!("{e}")))?;
        Ok(Button(ptr as *mut c_void))
    }

    pub fn create_label(text: &str) -> Result<Label, Error> {
        let ptr = crate::backends::android::create_label(text)
            .map_err(|e| Error::Backend(format!("{e}")))?;
        Ok(Label(ptr as *mut c_void))
    }

    pub fn create_box(orientation: Orientation, spacing: i32) -> Result<BoxWidget, Error> {
        let ptr = crate::backends::android::create_box(orientation.as_jni_int(), spacing)
            .map_err(|e| Error::Backend(format!("{e}")))?;
        Ok(BoxWidget(ptr as *mut c_void))
    }

    pub fn create_entry() -> Result<Entry, Error> {
        let ptr = crate::backends::android::create_entry()
            .map_err(|e| Error::Backend(format!("{e}")))?;
        Ok(Entry(ptr as *mut c_void))
    }

    pub fn create_grid() -> Result<Grid, Error> {
        let ptr = crate::backends::android::create_grid()
            .map_err(|e| Error::Backend(format!("{e}")))?;
        Ok(Grid(ptr as *mut c_void))
    }

    pub fn create_menu() -> Result<Menu, Error> {
        Ok(Menu { items: Vec::new() })
    }

    pub fn create_menubar(model: &Menu, _action_group: *mut c_void) -> Result<MenuBar, Error> {
        Ok(MenuBar { items: model.items.clone() })
    }

    pub fn create_simple_action(name: &str) -> Result<SimpleAction, Error> {
        Ok(SimpleAction { name: name.to_owned() })
    }

    pub fn create_canvas() -> Result<Canvas, Error> {
        // Real Java view (custom SheetView once registered, else a plain
        // View) so the widget tree never holds a bogus handle. Callback
        // registries key on the canvas id (see backend CANVAS_IDS map).
        let (ptr, _canvas_id) = crate::backends::android::create_canvas_view()
            .map_err(|e| Error::Backend(format!("{e}")))?;
        Ok(Canvas(ptr as *mut c_void))
    }

    pub fn create_overlay() -> Result<Overlay, Error> {
        let ptr = crate::backends::android::create_scrolled_view()
            .map_err(|e| Error::Backend(format!("{e}")))?;
        Ok(Overlay(ptr as *mut c_void))
    }

    pub fn create_scrolled_window() -> Result<ScrolledWindow, Error> {
        let ptr = crate::backends::android::create_scrolled_view()
            .map_err(|e| Error::Backend(format!("{e}")))?;
        Ok(ScrolledWindow(ptr as *mut c_void))
    }

    pub fn create_dropdown(items: &[&str]) -> Result<DropDown, Error> {
        let ptr = crate::backends::android::create_dropdown(items)
            .map_err(|e| Error::Backend(format!("{e}")))?;
        Ok(DropDown(ptr as *mut c_void))
    }

    pub fn create_checkbutton(label: &str) -> Result<CheckButton, Error> {
        let ptr = crate::backends::android::create_checkbutton(label)
            .map_err(|e| Error::Backend(format!("{e}")))?;
        Ok(CheckButton(ptr as *mut c_void))
    }

    pub fn create_radiobutton(
        group: Option<&RadioButton>,
        label: &str,
    ) -> Result<RadioButton, Error> {
        let group_ptr = group.map(|g| g.0).unwrap_or(std::ptr::null_mut());
        let ptr = crate::backends::android::create_radiobutton(group_ptr, label)
            .map_err(|e| Error::Backend(format!("{e}")))?;
        Ok(RadioButton(ptr as *mut c_void))
    }

    pub fn create_dialog() -> Result<Dialog, Error> {
        let ptr = crate::backends::android::create_dialog()
            .map_err(|e| Error::Backend(format!("{e}")))?;
        Ok(Dialog(ptr as *mut c_void))
    }

    pub fn create_textview() -> Result<TextView, Error> {
        let ptr = crate::backends::android::create_textview()
            .map_err(|e| Error::Backend(format!("{e}")))?;
        Ok(TextView(ptr as *mut c_void))
    }

    /// The display's density factor (`DisplayMetrics.density`): 1.0 at mdpi,
    /// 2.625 on a 420dpi phone, and so on.
    ///
    /// Hosts need this to size things in *pixels* while thinking in
    /// density-independent units: a hardcoded 12px font is 12dp on a desktop
    /// monitor but only ~4.6dp on a 420dpi phone, i.e. about a third of
    /// Android's 14sp body-text floor. Returns `None` before the backend is
    /// initialised (or if the JNI call fails), so callers can fall back to
    /// 1.0 and still render.
    pub fn display_density() -> Option<f64> {
        crate::backends::android::with_env_and_activity(|env, activity| {
            let res = env
                .call_method(activity.as_obj(), "getResources", "()Landroid/content/res/Resources;", &[])?
                .l()?;
            let metrics = env
                .call_method(&res, "getDisplayMetrics", "()Landroid/util/DisplayMetrics;", &[])?
                .l()?;
            let density = env.get_field(&metrics, "density", "F")?.f()?;
            Ok::<f64, Box<dyn std::error::Error + Send + Sync>>(density as f64)
        })
        .ok()
        .filter(|d| d.is_finite() && *d > 0.0)
    }
}

#[cfg(target_os = "android")]
pub use android_adapter::*;

#[cfg(test)]
#[cfg(target_os = "android")]
mod tests {
    use crate::core::Widget;

    #[test]
    fn test_null_window_handle() {
        let w = super::Window(std::ptr::null_mut());
        assert!(w.raw_handle().is_null());
    }

    #[test]
    fn test_null_button_handle() {
        let b = super::Button(std::ptr::null_mut());
        assert!(b.as_ref().is_null());
    }

    #[test]
    fn test_null_label_handle() {
        let l = super::Label(std::ptr::null_mut());
        assert!(l.as_ref().is_null());
    }

    #[test]
    fn test_null_box_handle() {
        let b = super::BoxWidget(std::ptr::null_mut());
        assert!(b.as_ref().is_null());
    }

    #[test]
    fn test_null_entry_handle() {
        let e = super::Entry(std::ptr::null_mut());
        assert!(e.as_ref().is_null());
    }

    #[test]
    fn test_clone_button() {
        let b1 = super::Button(0x1234 as *mut _);
        let b2 = b1.clone();
        assert_eq!(b1.as_ref(), b2.as_ref());
    }

    #[test]
    fn test_clone_label() {
        let l1 = super::Label(0x5678 as *mut _);
        let l2 = l1.clone();
        assert_eq!(l1.as_ref(), l2.as_ref());
    }

    #[test]
    fn test_clone_entry() {
        let e1 = super::Entry(0x9abc as *mut _);
        let e2 = e1.clone();
        assert_eq!(e1.as_ref(), e2.as_ref());
    }

    #[test]
    fn test_orientation_values_match_linear_layout() {
        // Discriminants mirror android.widget.LinearLayout constants.
        assert_eq!(super::Orientation::Horizontal as i32, 0);
        assert_eq!(super::Orientation::Vertical as i32, 1);
    }

    #[test]
    fn test_menu_submenu_clone_keeps_items() {
        let mut sub = super::Menu { items: Vec::new() };
        sub.append("Open", "app.open");
        let mut root = super::Menu { items: Vec::new() };
        root.append_submenu("File", &sub);
        let bar = super::create_menubar(&root, std::ptr::null_mut()).unwrap();
        let _ = bar.menu_active();
    }

    #[test]
    fn test_canvas_estimate_extents_nonzero() {
        let dc = super::AndroidDrawContext;
        let (x, y, w, h) = crate::core::DrawContext::text_extents_styled(&dc, "hello", "monospace", 12.0, 0, 0);
        let _ = (x, y);
        assert!(w > 0.0 && h > 0.0);
    }
}
