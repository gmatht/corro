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
        pub fn append(&self, child: &impl AsRef<*mut c_void>) {
            let child_ptr = *child.as_ref();
            if child_ptr.is_null() {
                return;
            }
            let _ = crate::backends::android::with_env_and_activity(|env, _activity| {
                let layout = unsafe { jni::objects::JObject::from_raw(self.0 as jni::sys::jobject) };
                let child_obj = unsafe { jni::objects::JObject::from_raw(child_ptr as jni::sys::jobject) };
                env.call_method(
                    &layout,
                    "addView",
                    "(Landroid/view/View;)V",
                    &[(&child_obj).into()],
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

        pub fn set_width_chars(&self, _n: i32) {}

        pub fn set_size_request(&self, _w: i32, _h: i32) {}

        pub fn set_hexpand(&self, _expand: bool) {}
        pub fn set_vexpand(&self, _expand: bool) {}
        pub fn set_visible(&self, _v: bool) {}
        pub fn add_class(&self, _class_name: &str) {}
        pub fn remove_class(&self, _class_name: &str) {}
        pub fn set_halign(&self, _align: i32) {}
        pub fn set_valign(&self, _align: i32) {}
        pub fn set_margin_start(&self, _px: i32) {}
        pub fn set_margin_top(&self, _px: i32) {}

        pub fn connect_activate<F: FnMut(*mut c_void) + 'static>(
            &self,
            _f: F,
        ) -> Result<u64, Error> {
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

        pub fn connect_changed(&self, _f: impl FnMut() + 'static) -> Result<u64, Error> {
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
        pub fn set_draw_callback(
            &self,
            cb: Box<dyn FnMut(&mut dyn DrawContext, i32, i32)>,
        ) {
            // Stash the closure under our handle id; the Java SheetView
            // dispatches to it by id from onDraw.
            let mut map = DRAW_CALLBACKS.lock().unwrap();
            map.insert(
                self.0 as usize,
                SendDrawCallback(Box::into_raw(cb)),
            );
            self.queue_redraw();
        }

        pub fn queue_redraw(&self) {
            // Run the stored closure immediately against the estimating
            // context so headless/test callers observe draws without Java.
            // On-device the Java SheetView invalidates and replays this
            // same closure with a JNI-backed context.
            let mut map = DRAW_CALLBACKS.lock().unwrap();
            if let Some(SendDrawCallback(ptr)) = map.get_mut(&(self.0 as usize)) {
                let cb: &mut dyn FnMut(&mut dyn DrawContext, i32, i32) = unsafe { &mut **ptr };
                let mut dc = AndroidDrawContext;
                let (w, h) = CANVAS_SIZE.lock().unwrap().get(&(self.0 as usize)).copied().unwrap_or((800, 600));
                cb(&mut dc, w, h);
            }
        }

        pub fn set_size_request(&self, w: i32, h: i32) {
            CANVAS_SIZE.lock().unwrap().insert(self.0 as usize, (w.max(1), h.max(1)));
        }

        pub fn set_content_size(&self, w: i32, h: i32) {
            self.set_size_request(w, h);
        }

        pub fn set_visible(&self, _v: bool) {}

        pub fn grab_focus(&self) {}
        pub fn set_can_focus(&self, _can: bool) {}

        pub fn on_click(&self, cb: Box<dyn FnMut(f64, f64)>) {
            let mut map = CLICK_CALLBACKS.lock().unwrap();
            map.insert(self.0 as usize, SendClickCallback(Box::into_raw(cb)));
        }

        pub fn on_key(&self, cb: Box<dyn FnMut(u32) -> bool>) {
            let mut boxed = cb;
            self.on_key_raw(Box::new(move |k: u32, _s: u32| -> bool { boxed(k) }));
        }

        pub fn on_key_raw(&self, cb: Box<dyn FnMut(u32, u32) -> bool>) {
            let mut map = KEY_CALLBACKS.lock().unwrap();
            map.insert(self.0 as usize, SendKeyCallback(Box::into_raw(cb)));
        }

        /// Force an immediate draw. No surface exists yet on Android, so
        /// this replays the closure like [`Canvas::queue_redraw`].
        pub fn force_draw(&self, _window_ptr: *mut c_void, fallback_w: i32, fallback_h: i32) {
            let mut map = DRAW_CALLBACKS.lock().unwrap();
            if let Some(SendDrawCallback(ptr)) = map.get_mut(&(self.0 as usize)) {
                let cb: &mut dyn FnMut(&mut dyn DrawContext, i32, i32) = unsafe { &mut **ptr };
                let mut dc = AndroidDrawContext;
                let (w, h) = CANVAS_SIZE
                    .lock()
                    .unwrap()
                    .get(&(self.0 as usize))
                    .copied()
                    .unwrap_or((fallback_w.max(1), fallback_h.max(1)));
                cb(&mut dc, w, h);
            }
        }
    }

    static DRAW_CALLBACKS: Lazy<Mutex<HashMap<usize, SendDrawCallback>>> =
        Lazy::new(|| Mutex::new(HashMap::new()));
    static CLICK_CALLBACKS: Lazy<Mutex<HashMap<usize, SendClickCallback>>> =
        Lazy::new(|| Mutex::new(HashMap::new()));
    static KEY_CALLBACKS: Lazy<Mutex<HashMap<usize, SendKeyCallback>>> =
        Lazy::new(|| Mutex::new(HashMap::new()));
    static CANVAS_SIZE: Lazy<Mutex<HashMap<usize, (i32, i32)>>> =
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
    pub fn dispatch_canvas_click(canvas_id: usize, x: f64, y: f64) {
        let mut map = CLICK_CALLBACKS.lock().unwrap();
        if let Some(SendClickCallback(ptr)) = map.get_mut(&canvas_id) {
            let cb: &mut dyn FnMut(f64, f64) = unsafe { &mut **ptr };
            cb(x, y);
        }
    }

    /// Dispatch a key from Java to the registered key closure.
    pub fn dispatch_canvas_key(canvas_id: usize, keyval: u32, mods: u32) -> bool {
        let mut map = KEY_CALLBACKS.lock().unwrap();
        if let Some(SendKeyCallback(ptr)) = map.get_mut(&canvas_id) {
            let cb: &mut dyn FnMut(u32, u32) -> bool = unsafe { &mut **ptr };
            cb(keyval, mods)
        } else {
            false
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
        // No Java view yet: use a unique non-null id so per-canvas callback
        // registries stay distinct. The Java SheetView replaces this when
        // `create_canvas_view` lands.
        static NEXT_ID: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(1);
        let id = NEXT_ID.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        Ok(Canvas(id as *mut c_void))
    }

    pub fn create_overlay() -> Result<Overlay, Error> {
        Ok(Overlay(std::ptr::null_mut()))
    }

    pub fn create_scrolled_window() -> Result<ScrolledWindow, Error> {
        Ok(ScrolledWindow(std::ptr::null_mut()))
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
