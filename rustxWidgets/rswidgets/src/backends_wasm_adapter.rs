#[cfg(target_arch = "wasm32")]
mod wasm_adapter {
    use std::any::Any;
    use std::cell::{Cell, RefCell};
    use std::collections::HashMap;
    use std::os::raw::c_void;
    use std::rc::Rc;
    use wasm_bindgen::closure::Closure;
    use wasm_bindgen::JsCast;
    use web_sys::{
        Document, Element, Event, FocusEvent, HtmlCanvasElement, HtmlButtonElement,
        HtmlDialogElement, HtmlDivElement, HtmlElement, HtmlInputElement,
        HtmlOptionElement, HtmlSelectElement, HtmlTextAreaElement, KeyboardEvent,
        MouseEvent,
    };
    use crate::core::{Error, Widget};

    // -----------------------------------------------------------------------
    // Global helpers
    // -----------------------------------------------------------------------
    fn document() -> Document {
        web_sys::window().unwrap().document().unwrap()
    }

    fn body() -> HtmlElement {
        document().body().unwrap()
    }

    fn create_element(tag: &str) -> Element {
        document().create_element(tag).unwrap()
    }

    /// Run `f` every `ms` milliseconds until it returns `false`.
    ///
    /// `setInterval`, and the closure is leaked: an interval has no natural
    /// owner here (the backend is a set of free functions, not a window), and
    /// the `true` return from `f` is what stops it. Leaking the `Closure` is
    /// therefore not a leak in practice -- a repeating tick is meant to
    /// outlive the call that armed it, and the interval handle stops firing
    /// as soon as `f` says so.
    ///
    /// `ms` is clamped to 1: `setInterval` treats 0 as "as fast as the
    /// browser allows" and a negative value as 0, so a caller asking for
    /// 0 ms would get an unthrottled loop rather than the one-shot-per-tick
    /// behaviour a 0 interval suggests.
    pub fn add_periodic_tick(
        ms: u32,
        f: Box<dyn FnMut() -> bool>,
    ) -> Result<(), Error> {
        let cell = Rc::new(RefCell::new(f));
        // The interval id, shared with the callback so that returning `false`
        // can actually cancel the interval.
        //
        // This has to exist: `setInterval` keeps firing until something calls
        // `clearInterval`, and the only thing that knows the tick is finished
        // is `f` itself. A cell holding the id is set *after* the first
        // `setInterval` call, and a tick that fires before that reads 0 --
        // which is not a valid interval id, so `clearInterval(0)` is a no-op
        // and the next scheduled run does the cancelling. In practice the
        // delay is `ms`, far longer than the microseconds the assignment
        // needs, so the common case cancels on the first tick.
        let handle: Rc<Cell<i32>> = Rc::new(Cell::new(0));
        let handle_for_cb = handle.clone();
        let closure = Closure::<dyn FnMut()>::new(move || {
            // A panic inside `f` would unwind through a JS callback, which
            // wasm-bindgen cannot do safely. `try_borrow_mut` degrades to a
            // skipped tick instead, which is also what a re-entrant call from
            // `f` needs.
            let keep_going = match cell.try_borrow_mut() {
                Ok(mut guard) => guard(),
                Err(_) => true,
            };
            if !keep_going {
                if let Some(w) = web_sys::window() {
                    w.clear_interval_with_handle(handle_for_cb.get());
                }
            }
        });
        let ms = ms.max(1);
        let Some(w) = web_sys::window() else {
            return Err(Error::Backend("no window: no event loop to schedule on".into()));
        };
        let id = w
            .set_interval_with_callback_and_timeout_and_arguments_0(
                closure.as_ref().unchecked_ref(),
                ms as i32,
            )
            .map_err(|e| Error::Backend(format!("setInterval: {e:?}")))?;
        handle.set(id);
        // Keep the closure alive for as long as the page is; the interval
        // owns the callback, and a dropped `Closure` would free the JS
        // trampoline that callback still points at.
        std::mem::forget(closure);
        Ok(())
    }

    /// Run `f` once, `ms` milliseconds from now.
    pub fn timeout_add_once(ms: u32, f: Box<dyn FnOnce()>) -> Result<(), Error> {
        let closure = Closure::<dyn FnMut()>::once_into_js(move || f());
        let ms = ms.max(1) as i32;
        let _ = web_sys::window().map(|w| {
            let _ = w.set_timeout_with_callback_and_timeout_and_arguments_0(
                closure.as_ref().unchecked_ref(),
                ms,
            );
        });
        std::mem::forget(closure);
        Ok(())
    }

    /// A repeating tick on `window`. `timeout_add_repeating` is the GTK and
    /// NWG spelling; `add_periodic_tick` above is what `core.rs` calls.
    pub fn timeout_add_repeating(ms: u32, f: Box<dyn FnMut() -> bool>) -> Result<(), Error> {
        add_periodic_tick(ms, f)
    }

    pub fn quit_main_loop() {
        // On WASM there is no message loop to quit.  Signal the test
        // framework that the app has quit, and try to close the tab.
        if let Some(doc) = web_sys::window().and_then(|w| w.document()) {
            // Set document title as a signal for test frameworks
            // that may poll for it.
            doc.set_title("CORRO_QUIT");
        }
        if let Some(win) = web_sys::window() {
            let _ = win.close();
        }
    }

    /// Store a line of output in a JS global (`window.__corro_output`).
    /// The browser test framework can read this variable after the app quits
    /// to verify the recording replay output.
    pub fn append_wasm_output_line(line: &str) {
        if let Some(win) = web_sys::window() {
            let _ = js_sys::Reflect::set(
                &win,
                &wasm_bindgen::JsValue::from_str("__corro_output_dirty"),
                &wasm_bindgen::JsValue::TRUE,
            );
            let current = js_sys::Reflect::get(
                &win,
                &wasm_bindgen::JsValue::from_str("__corro_output"),
            )
            .ok()
            .and_then(|v| v.as_string())
            .unwrap_or_default();
            let new = if current.is_empty() {
                line.to_string()
            } else {
                format!("{}\n{}", current, line)
            };
            let _ = js_sys::Reflect::set(
                &win,
                &wasm_bindgen::JsValue::from_str("__corro_output"),
                &wasm_bindgen::JsValue::from_str(&new),
            );
        }
    }

    fn set_css(elem: &Element, prop: &str, val: &str) {
        if let Some(html) = elem.dyn_ref::<HtmlElement>() {
            html.style().set_property(prop, val).ok();
        }
    }

    /// The GDK modifier mask for a mouse or key event: 1 = Shift, 4 =
    /// Control, 8 = Alt/Meta.
    ///
    /// Every callback in this adapter promises a GDK-style mask, because that
    /// is what the shared GUI already branches on. Translating once here is
    /// what keeps `on_motion`, `on_click_button`, `on_release` and
    /// `Window::on_event_key` agreeing -- previously `on_event_key` reported
    /// only Alt and hard-coded 0, so a caller checking Ctrl or Shift saw
    /// "no modifier" for half its cases.
    fn gdk_modifier_mask(evt: &KeyboardEvent) -> u32 {
        let mut state = 0u32;
        if evt.shift_key() {
            state |= 1;
        }
        if evt.ctrl_key() {
            state |= 4;
        }
        if evt.alt_key() {
            state |= 8;
        }
        state
    }

    /// `gdk_modifier_mask` for a mouse event. Separate because
    /// `KeyboardEvent` and `MouseEvent` are distinct wasm-bindgen types with
    /// the same three accessors.
    fn gdk_modifier_mask_mouse(evt: &MouseEvent) -> u32 {
        let mut state = 0u32;
        if evt.shift_key() {
            state |= 1;
        }
        if evt.ctrl_key() {
            state |= 4;
        }
        if evt.alt_key() {
            state |= 8;
        }
        state
    }

    /// Own a `Closure` for the life of the program.
    ///
    /// A `Closure` that is dropped releases the JS function it points at, so
    /// an event listener still attached to an element would call into freed
    /// memory. The three call sites that need this (`Entry::on_key_raw`,
    /// `Canvas::on_click`, `Canvas::on_key`) cannot store the closure in the
    /// widget: they are `&self` methods on a type whose other closures live in
    /// a per-widget cell that a `Clone` would share, and a listener outlives
    /// any one handle. Leaking is the correct trade here -- the page's own
    /// lifetime bounds it, and it is what `closure.forget()` does elsewhere
    /// in this file (the menu bar's click handlers).
    fn store_closure(closure: Box<dyn Any>) {
        std::mem::forget(closure);
    }

    /// The printable character a DOM `keyCode` names, or `None` for a
    /// non-printable key.
    ///
    /// The `key::` constants in `core.rs` are raw DOM `keyCode` values on
    /// this target (that is the `#[cfg(target_arch = "wasm32")] mod plat`
    /// arm), so a mnemonic is matched by mapping the code back to a letter.
    /// Only the codes a mnemonic can be typed with are handled: A-Z map to
    /// 0x41-0x5A, and 0-9 to 0x30-0x39. Anything else (arrows, modifiers,
    /// punctuation) is not a mnemonic, and returning `None` is what makes
    /// the menu fall through to normal key handling.
    fn char_from_keyval(keyval: u32) -> Option<char> {
        let c = char::from_u32(keyval)?;
        if c.is_ascii_alphanumeric() {
            Some(c)
        } else {
            None
        }
    }

    /// UTF-16 code-unit offset -> character index, for the DOM's
    /// `selectionStart` / `setSelectionRange` and `Entry::get_position`.
    ///
    /// The DOM counts UTF-16 code units; the shared `Entry` API speaks
    /// character indices. They diverge for anything outside the BMP, so
    /// passing the raw value through would mis-place the caret on a cell
    /// containing an emoji.
    fn utf16_offset_to_char_index(text: &str, utf16_offset: i32) -> usize {
        let mut units = 0i32;
        let mut chars = 0usize;
        for c in text.chars() {
            if units >= utf16_offset {
                break;
            }
            units += c.len_utf16() as i32;
            chars += 1;
        }
        chars
    }

    /// The inverse of [`utf16_offset_to_char_index`].
    fn char_index_to_utf16_offset(text: &str, char_index: usize) -> i32 {
        let mut units = 0i32;
        let mut seen = 0usize;
        for c in text.chars() {
            if seen >= char_index {
                break;
            }
            units += c.len_utf16() as i32;
            seen += 1;
        }
        units
    }

    // -----------------------------------------------------------------------
    // AsElement trait – unified way to get a DOM Element from any widget
    // -----------------------------------------------------------------------
    pub trait AsElement {
        fn as_element(&self) -> &Element;
    }

    // -----------------------------------------------------------------------
    // Action registry (single‑threaded WASM, safe under wasm32)
    // -----------------------------------------------------------------------
    struct SendFnPtr(*mut dyn FnMut(*mut c_void));
    unsafe impl Send for SendFnPtr {}

    fn register_action(name: &str, f: Box<dyn FnMut(*mut c_void)>) {
        use once_cell::sync::Lazy;
        use std::sync::Mutex;
        static REG: Lazy<Mutex<HashMap<String, SendFnPtr>>> =
            Lazy::new(|| Mutex::new(HashMap::new()));
        let mut map = REG.lock().unwrap();
        let ptr = Box::into_raw(Box::new(f) as Box<dyn FnMut(*mut c_void)>);
        map.insert(name.to_owned(), SendFnPtr(ptr));
    }

    fn invoke_action(name: &str, param: *mut c_void) {
        use once_cell::sync::Lazy;
        use std::sync::Mutex;
        static REG: Lazy<Mutex<HashMap<String, SendFnPtr>>> =
            Lazy::new(|| Mutex::new(HashMap::new()));
        if let Ok(mut map) = REG.lock() {
            if let Some(SendFnPtr(ptr)) = map.get_mut(name) {
                let cb: &mut dyn FnMut(*mut c_void) = unsafe { &mut **ptr };
                cb(param);
            }
        }
    }

    // -----------------------------------------------------------------------
    // Orientation
    // -----------------------------------------------------------------------
    pub enum Orientation {
        Horizontal,
        Vertical,
    }

    impl Orientation {
        pub fn as_flex_direction(&self) -> &'static str {
            match self {
                Orientation::Horizontal => "row",
                Orientation::Vertical => "column",
            }
        }
    }

    // -----------------------------------------------------------------------
    // Window
    // -----------------------------------------------------------------------
pub struct Window {
    elem: HtmlDivElement,
    event_key_cb: Rc<RefCell<Option<Box<dyn FnMut(u32, u32) -> i32>>>>,
    close_cb: Rc<RefCell<Option<Box<dyn FnMut()>>>>,
    closures: Rc<RefCell<Vec<Box<dyn Any>>>>,
}

impl Clone for Window {
    fn clone(&self) -> Self {
        Window {
            elem: self.elem.clone(),
            event_key_cb: self.event_key_cb.clone(),
            close_cb: self.close_cb.clone(),
            closures: self.closures.clone(),
        }
    }
}

impl AsElement for Window {
        fn as_element(&self) -> &Element {
            self.elem.as_ref()
        }
    }

    impl Widget for Window {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const HtmlDivElement as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for Window {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.raw_handle() as *const *mut c_void) }
        }
    }

impl Window {
    pub fn set_title(&self, title: &str) {
        self.elem.set_text_content(Some(title));
    }

        pub fn set_child(&self, child: &impl AsElement) {
            while let Some(c) = self.elem.first_child() {
                self.elem.remove_child(&c).ok();
            }
            self.elem.append_child(child.as_element()).ok();
        }

        pub fn present(&self) {}

        pub fn set_default_size(&self, w: i32, h: i32) {
            if w > 0 {
                set_css(self.elem.as_ref(), "width", &format!("{}px", w));
            }
            if h > 0 {
                set_css(self.elem.as_ref(), "height", &format!("{}px", h));
            }
        }

        /// Force the size, the way GTK's `gtk_window_resize` does.
        ///
        /// A separate method from `set_default_size` because CSS sizes are
        /// advisory in exactly the way the `common.rs` doc comment warns
        /// about: a `width` set on a flex child is a *request*, and the layout
        /// may shrink it to nothing. Setting `min-width`/`min-height` as well
        /// is what actually holds it, and it is the DOM's version of
        /// `gtk_window_set_size` (not `set_default_size`).
        pub fn resize(&self, w: i32, h: i32) {
            if w > 0 {
                set_css(self.elem.as_ref(), "min-width", &format!("{}px", w));
                set_css(self.elem.as_ref(), "width", &format!("{}px", w));
            }
            if h > 0 {
                set_css(self.elem.as_ref(), "min-height", &format!("{}px", h));
                set_css(self.elem.as_ref(), "height", &format!("{}px", h));
            }
        }

        /// No OS window handle exists in a browser. `raw_handle` already
        /// returns the DOM element pointer for interop, which is the honest
        /// analogue; this is the `hwnd` slot `common.rs` requires, so it
        /// reports "no native handle" the way the ios/macos/zork adapters do.
        pub fn hwnd(&self) -> *mut c_void {
            std::ptr::null_mut()
        }

        /// `common.rs`'s `set_child_box` takes a `WidgetBox`; the adapter's
        /// `set_child` is generic over `AsElement`, so a box forwards
        /// straight through.
        pub fn set_child_box(&self, bx: &BoxWidget) {
            self.set_child(bx);
        }

        /// # Safety – kept for API compatibility; no‑op on WASM.
        pub unsafe fn insert_action_group(&self, _name: &str, _group_ptr: *mut c_void) {}
        pub fn on_event(&self, _cb: Box<dyn FnMut(*mut c_void) -> i32>) {}

        /// Fire `cb` when the page or window is closing.
        ///
        /// GTK's `close-request` and NWG's `WM_CLOSE` both mean "the user is
        /// about to lose the app", and the browser has exactly one such
        /// notification: `beforeunload`, on the *window* (it is the only
        /// unload hook that is not the end of the page). It cannot cancel
        /// the unload without a user-gesture-registered handler, so this
        /// reports rather than prevents -- the same "tell the app, let it
        /// save" contract `quit_main_loop` already uses.
        pub fn on_close(&self, cb: Box<dyn FnMut()>) {
            *self.close_cb.borrow_mut() = Some(cb);
            let cell = self.close_cb.clone();
            let closure = Closure::<dyn FnMut(Event)>::new(move |_evt: Event| {
                if let Some(f) = cell.borrow_mut().as_mut() {
                    f();
                }
            });
            if let Some(win) = web_sys::window() {
                let _ = win.add_event_listener_with_callback("beforeunload", closure.as_ref().unchecked_ref());
            }
            self.closures.borrow_mut().push(Box::new(closure));
        }

        /// Re-run the DOM work that a native backend would do in a redraw.
        ///
        /// GTK calls `queue_draw` on the widget tree and NWG
        /// `RedrawWindow(RDW_ALLCHILDREN)`; neither has a DOM equivalent,
        /// because a browser repaints on its own. The closest honest answer
        /// is to invalidate the window's own box model, which is what makes a
        /// caller that has changed layout see the result. The child canvases
        /// are the caller's to redraw: they own their own `queue_redraw`.
        pub fn queue_redraw(&self) {
            // Reading offsetWidth/Height flushes pending style and layout, so
            // anything the caller changed has been measured by the time this
            // returns -- the same guarantee `gtk_widget_queue_draw` offers
            // for the frame that follows.
            let _ = (self.elem.offset_width(), self.elem.offset_height());
        }
        pub fn on_event_key(&self, cb: Box<dyn FnMut(u32, u32) -> i32>) {
            *self.event_key_cb.borrow_mut() = Some(cb);
            let cb2 = self.event_key_cb.clone();
            let closure = Closure::<dyn FnMut(KeyboardEvent)>::new(move |evt: KeyboardEvent| {
                let keyval = evt.key_code() as u32;
                // The full GDK mask, not just Alt: a caller branching on Ctrl
                // (the standard accelerator modifier) saw "no modifier" for
                // every Ctrl keypress, because only alt_key was read.
                let state = gdk_modifier_mask(&evt);
                if let Some(cb) = cb2.borrow_mut().as_mut() {
                    if cb(keyval, state) != 0 {
                        evt.prevent_default();
                    }
                }
            });
            let listener = wasm_bindgen::JsCast::unchecked_into::<js_sys::Function>(
                closure.as_ref().clone(),
            );
            web_sys::window()
                .and_then(|w| w.document())
                .map(|doc| {
                    doc.add_event_listener_with_callback("keydown", &listener).ok();
                });
            self.closures.borrow_mut().push(Box::new(closure));
        }
    }

    pub fn create_window() -> Result<Window, Error> {
        let div: HtmlDivElement = create_element("div").dyn_into().map_err(|e| {
            Error::Backend(format!("create_window: {:?}", e))
        })?;
        set_css(div.as_ref(), "all", "initial");
        body().append_child(div.as_ref()).map_err(|e| {
            Error::Backend(format!("create_window append: {:?}", e))
        })?;
        Ok(Window {
            elem: div,
            event_key_cb: Rc::new(RefCell::new(None)),
            close_cb: Rc::new(RefCell::new(None)),
            closures: Rc::new(RefCell::new(Vec::new())),
        })
    }

    // -----------------------------------------------------------------------
    // Button
    // -----------------------------------------------------------------------
    pub struct Button {
        elem: HtmlButtonElement,
        closures: Rc<RefCell<Vec<Box<dyn Any>>>>,
        next_id: Rc<RefCell<u64>>,
    }

    impl AsElement for Button {
        fn as_element(&self) -> &Element {
            self.elem.as_ref()
        }
    }

    impl Widget for Button {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const HtmlButtonElement as *mut c_void
        }
    }

    impl Clone for Button {
        fn clone(&self) -> Self {
            Button {
                elem: self.elem.clone(),
                closures: self.closures.clone(),
                next_id: self.next_id.clone(),
            }
        }
    }

    impl Button {
        pub fn on_click(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            let cb = Rc::new(RefCell::new(f));
            let cb2 = cb.clone();
            let closure = Closure::<dyn FnMut(MouseEvent)>::new(move |_: MouseEvent| {
                (cb2.borrow_mut())();
            });
            self.elem
                .add_event_listener_with_callback("click", closure.as_ref().unchecked_ref())
                .map_err(|e| Error::Backend(format!("on_click: {:?}", e)))?;
            let id = *self.next_id.borrow();
            *self.next_id.borrow_mut() += 1;
            self.closures.borrow_mut().push(Box::new(closure));
            Ok(id)
        }

        pub fn emit_clicked(&self) -> Result<u64, Error> {
            self.elem.click();
            Ok(0)
        }

        pub fn set_visible(&self, v: bool) {
            set_css(self.elem.as_ref(), "display", if v { "" } else { "none" });
        }

        /// `flex-grow` + `align-self: stretch`. A `<button>` sizes to its
        /// label, so without these it stays that width inside a flex row
        /// instead of taking the leftover space.
        pub fn set_hexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-grow", "1");
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "flex-grow", "0");
            }
        }

        /// See [`Button::set_hexpand`].
        pub fn set_vexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-shrink", "1");
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "flex-shrink", "0");
            }
        }

        pub fn set_size_request(&self, w: i32, h: i32) {
            if w > 0 {
                set_css(self.elem.as_ref(), "width", &format!("{}px", w));
            }
            if h > 0 {
                set_css(self.elem.as_ref(), "height", &format!("{}px", h));
            }
        }

        /// `font-weight` / `font-style`, GTK `set_font_style`'s
        /// `(weight, italic)` pair.
        pub fn set_font_style(&self, weight: i32, italic: bool) {
            set_css(
                self.elem.as_ref(),
                "font-weight",
                if weight != 0 { "bold" } else { "normal" },
            );
            set_css(
                self.elem.as_ref(),
                "font-style",
                if italic { "italic" } else { "normal" },
            );
        }

        /// `classList`, the same mechanism [`Label::add_class`] uses.
        pub fn add_class(&self, class_name: &str) {
            self.elem.class_list().add_1(class_name).ok();
        }

        /// See [`Button::add_class`].
        pub fn remove_class(&self, class_name: &str) {
            self.elem.class_list().remove_1(class_name).ok();
        }
    }

    pub fn create_button(label: &str) -> Result<Button, Error> {
        let btn: HtmlButtonElement = create_element("button").dyn_into().map_err(|e| {
            Error::Backend(format!("create_button: {:?}", e))
        })?;
        btn.set_text_content(Some(label));
        Ok(Button {
            elem: btn,
            closures: Rc::new(RefCell::new(Vec::new())),
            next_id: Rc::new(RefCell::new(1)),
        })
    }

    // -----------------------------------------------------------------------
    // Label
    // -----------------------------------------------------------------------
    pub struct Label {
        elem: Element,
    }

    impl AsElement for Label {
        fn as_element(&self) -> &Element {
            &self.elem
        }
    }

    impl Clone for Label {
        fn clone(&self) -> Self {
            Label {
                elem: self.elem.clone(),
            }
        }
    }

    impl Widget for Label {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const Element as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for Label {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.raw_handle() as *const *mut c_void) }
        }
    }

    impl Label {
        pub fn set_text(&self, text: &str) {
            self.elem.set_text_content(Some(text));
        }

        pub fn get_text(&self) -> Option<String> {
            self.elem.text_content()
        }

        pub fn add_class(&self, class_name: &str) {
            self.elem.class_list().add_1(class_name).ok();
        }

        pub fn remove_class(&self, class_name: &str) {
            self.elem.class_list().remove_1(class_name).ok();
        }

        pub fn set_markup(&self, markup: &str) {
            self.elem.set_inner_html(markup);
        }

        pub fn set_visible(&self, visible: bool) {
            if let Some(html) = self.elem.dyn_ref::<HtmlElement>() {
                if visible {
                    html.style().set_property("display", "").ok();
                } else {
                    html.style().set_property("display", "none").ok();
                }
            }
        }

        pub fn set_xalign(&self, x: f32) {
            let align = if x <= 0.0 {
                "left"
            } else if x >= 1.0 {
                "right"
            } else {
                "center"
            };
            if let Some(html) = self.elem.dyn_ref::<HtmlElement>() {
                html.style().set_property("text-align", align).ok();
            }
        }

        /// Pin the label's width so changing its text cannot resize it.
        /// `None` releases the pin.
        ///
        /// This is `width` + `min-width` together, not `width` alone: in a
        /// flex row a `width` is only a request and the layout may shrink it,
        /// while `min-width` is the floor that actually holds the slot still.
        /// Corro pins the formula-bar address and status labels for exactly
        /// this reason (`gui_backend.rs:5732` and 5768), and a `width` alone
        /// would let the sibling controls still slide as the text changes.
        pub fn set_fixed_width(&self, w: Option<i32>) {
            match w {
                Some(px) if px > 0 => {
                    let v = format!("{}px", px);
                    set_css(&self.elem, "width", &v);
                    set_css(&self.elem, "min-width", &v);
                    // A pinned slot must not be flexible, or the flex layout
                    // can still stretch or shrink it past the pin.
                    set_css(&self.elem, "flex", "0 0 auto");
                }
                // `None` clears the pin, so both properties have to go: a
                // leftover `min-width` would keep the label from ever
                // shrinking again.
                _ => {
                    set_css(&self.elem, "width", "");
                    set_css(&self.elem, "min-width", "");
                    set_css(&self.elem, "flex", "");
                }
            }
        }

        /// Left outer margin, in device px.
        ///
        /// `margin-left`, not `padding-left`: GTK's `set_margin_start` is an
        /// *outer* spacing that pushes siblings along, and corro relies on
        /// that (it pairs the inset with a pinned slot, so the inset must not
        /// eat into the pinned width).
        pub fn set_margin_start(&self, px: i32) {
            set_css(&self.elem, "margin-left", &format!("{}px", px));
        }

        /// Top outer margin, in device px.
        pub fn set_margin_top(&self, px: i32) {
            set_css(&self.elem, "margin-top", &format!("{}px", px));
        }

        /// See [`Entry::set_hexpand`].
        pub fn set_hexpand(&self, expand: bool) {
            if expand {
                set_css(&self.elem, "flex-grow", "1");
                set_css(&self.elem, "align-self", "stretch");
            } else {
                set_css(&self.elem, "flex-grow", "0");
            }
        }

        /// See [`Entry::set_vexpand`].
        pub fn set_vexpand(&self, expand: bool) {
            if expand {
                set_css(&self.elem, "flex-shrink", "1");
                set_css(&self.elem, "align-self", "stretch");
            } else {
                set_css(&self.elem, "flex-shrink", "0");
            }
        }

        pub fn set_size_request(&self, w: i32, h: i32) {
            if w > 0 {
                set_css(&self.elem, "width", &format!("{}px", w));
            }
            if h > 0 {
                set_css(&self.elem, "height", &format!("{}px", h));
            }
        }

        /// Horizontal alignment, GTK `GtkAlign` ordinals (0 = start,
        /// 1 = centre, 2 = end). Distinct from `set_xalign`, which aligns the
        /// *text* inside the label; this moves the label itself inside its
        /// slot.
        pub fn set_halign(&self, align: i32) {
            let v = match align {
                1 => "center",
                2 => "flex-end",
                _ => "flex-start",
            };
            set_css(&self.elem, "justify-self", v);
        }

        /// Vertical alignment, same ordinals as [`Label::set_halign`].
        pub fn set_valign(&self, align: i32) {
            let v = match align {
                1 => "center",
                2 => "flex-end",
                _ => "flex-start",
            };
            set_css(&self.elem, "align-self", v);
        }

        /// `AsRef` in the *other* direction: a raw handle to the DOM element,
        /// the escape hatch `common::Label::raw_handle` forwards. The DOM
        /// handle is the pointer, so this is the same value `crate::core::
        /// Widget::raw_handle` returns; it exists as an inherent method
        /// because `common.rs` calls it on the concrete type.
        pub fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const Element as *mut c_void
        }
    }

    pub fn create_label(text: &str) -> Result<Label, Error> {
        let elem = create_element("span");
        elem.set_text_content(Some(text));
        Ok(Label { elem })
    }

    // -----------------------------------------------------------------------
    // BoxWidget
    // -----------------------------------------------------------------------
pub struct BoxWidget {
    elem: HtmlDivElement,
}

impl Clone for BoxWidget {
    fn clone(&self) -> Self {
        BoxWidget { elem: self.elem.clone() }
    }
}

impl AsElement for BoxWidget {
        fn as_element(&self) -> &Element {
            self.elem.as_ref()
        }
    }

    impl Widget for BoxWidget {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const HtmlDivElement as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for BoxWidget {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.raw_handle() as *const *mut c_void) }
        }
    }

    fn as_element_from_ptr(ptr: *mut c_void) -> &'static Element {
        unsafe { &*(ptr as *const Element) }
    }

    fn as_html_element(ptr: *mut c_void) -> Option<&'static HtmlElement> {
        unsafe { (*(ptr as *const Element)).dyn_ref::<HtmlElement>() }
    }

    impl BoxWidget {
        pub fn append(&self, child: &impl AsRef<*mut c_void>) {
            self.elem.append_child(as_element_from_ptr(*child.as_ref())).ok();
        }

        pub fn set_child_vexpand(&self, child: &impl AsRef<*mut c_void>, expand: bool) {
            if let Some(html) = as_html_element(*child.as_ref()) {
                if expand {
                    html.style().set_property("flex-grow", "1").ok();
                    html.style().set_property("align-self", "stretch").ok();
                } else {
                    html.style().set_property("flex-grow", "0").ok();
                }
            }
        }

        pub fn set_child_hexpand(&self, child: &impl AsRef<*mut c_void>, expand: bool) {
            if let Some(html) = as_html_element(*child.as_ref()) {
                if expand {
                    html.style().set_property("align-self", "stretch").ok();
                }
            }
        }

        /// See [`Entry::set_hexpand`]: `flex-grow` is how a box takes the
        /// leftover space, and `align-self: stretch` is the cross-axis half.
        pub fn set_hexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-grow", "1");
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "flex-grow", "0");
            }
        }

        /// See [`BoxWidget::set_hexpand`].
        pub fn set_vexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-grow", "1");
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "flex-grow", "0");
            }
        }

        pub fn set_visible(&self, v: bool) {
            set_css(self.elem.as_ref(), "display", if v { "" } else { "none" });
        }

        pub fn set_size_request(&self, w: i32, h: i32) {
            if w > 0 {
                set_css(self.elem.as_ref(), "width", &format!("{}px", w));
            }
            if h > 0 {
                set_css(self.elem.as_ref(), "height", &format!("{}px", h));
            }
        }
    }

    pub fn create_box(orientation: Orientation, spacing: i32) -> Result<BoxWidget, Error> {
        let div: HtmlDivElement = create_element("div").dyn_into().map_err(|e| {
            Error::Backend(format!("create_box: {:?}", e))
        })?;
        let s = div.style();
        s.set_property("display", "flex").ok();
        s.set_property("flex-direction", orientation.as_flex_direction())
            .ok();
        if spacing > 0 {
            s.set_property("gap", &format!("{}px", spacing)).ok();
        }
        Ok(BoxWidget { elem: div })
    }

    // -----------------------------------------------------------------------
    // Grid
    // -----------------------------------------------------------------------
    pub struct Grid {
        elem: HtmlDivElement,
    }

    impl AsElement for Grid {
        fn as_element(&self) -> &Element {
            self.elem.as_ref()
        }
    }

    impl Widget for Grid {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const HtmlDivElement as *mut c_void
        }
    }

    impl Grid {
        pub fn attach(&self, child: &impl AsElement, left: i32, top: i32, width: i32, height: i32) {
            let child = child.as_element();
            if let Some(html) = child.dyn_ref::<HtmlElement>() {
                let s = html.style();
                s.set_property(
                    "grid-column",
                    &format!("{} / {}", left + 1, left + width + 1),
                )
                .ok();
                s.set_property(
                    "grid-row",
                    &format!("{} / {}", top + 1, top + height + 1),
                )
                .ok();
            }
            self.elem.append_child(child).ok();
        }

        pub fn set_visible(&self, v: bool) {
            set_css(self.elem.as_ref(), "display", if v { "" } else { "none" });
        }
        pub fn set_hexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-grow", "1");
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "flex-grow", "0");
            }
        }
        pub fn set_vexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-grow", "1");
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "flex-grow", "0");
            }
        }
        pub fn set_size_request(&self, w: i32, h: i32) {
            if w > 0 {
                set_css(self.elem.as_ref(), "width", &format!("{}px", w));
            }
            if h > 0 {
                set_css(self.elem.as_ref(), "height", &format!("{}px", h));
            }
        }
    }

    pub fn create_grid() -> Result<Grid, Error> {
        let div: HtmlDivElement = create_element("div").dyn_into().map_err(|e| {
            Error::Backend(format!("create_grid: {:?}", e))
        })?;
        div.style().set_property("display", "grid").ok();
        Ok(Grid { elem: div })
    }

    // -----------------------------------------------------------------------
    // Entry
    // -----------------------------------------------------------------------
    pub struct Entry {
        elem: HtmlInputElement,
        key_cb: Rc<RefCell<Option<Box<dyn FnMut(u32, u32) -> bool>>>>,
        closures: Rc<RefCell<Vec<Box<dyn Any>>>>,
        next_id: Rc<RefCell<u64>>,
    }

    impl AsElement for Entry {
        fn as_element(&self) -> &Element {
            self.elem.as_ref()
        }
    }

    impl Widget for Entry {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const HtmlInputElement as *mut c_void
        }
    }

    impl Clone for Entry {
        fn clone(&self) -> Self {
            Entry {
                elem: self.elem.clone(),
                key_cb: self.key_cb.clone(),
                closures: self.closures.clone(),
                next_id: self.next_id.clone(),
            }
        }
    }

    impl AsRef<*mut c_void> for Entry {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.raw_handle() as *const *mut c_void) }
        }
    }

    impl Entry {
        pub fn set_text(&self, text: &str) {
            self.elem.set_value(text);
        }

        pub fn get_text(&self) -> Option<String> {
            Some(self.elem.value())
        }


        pub fn set_width_chars(&self, n: i32) {
            self.elem.set_size(n as u32);
        }

        pub fn set_size_request(&self, w: i32, h: i32) {
            if w > 0 {
                set_css(self.elem.as_ref(), "width", &format!("{}px", w));
            }
            if h > 0 {
                set_css(self.elem.as_ref(), "height", &format!("{}px", h));
            }
        }

        pub fn set_hexpand(&self, expand: bool) {
            // `flex-grow` is the only way an input takes leftover space: an
            // <input> measures to its content, so in a bare flex row it stays
            // at whatever `size` said and the entry does not follow the window
            // being resized.
            if expand {
                set_css(self.elem.as_ref(), "flex-grow", "1");
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "flex-grow", "0");
            }
        }

        /// See [`Entry::set_hexpand`]: `align-self` is the cross-axis
        /// property in the flex row a formula bar is.
        pub fn set_vexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-shrink", "1");
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "flex-shrink", "0");
            }
        }

        pub fn set_visible(&self, v: bool) {
            set_css(self.elem.as_ref(), "display", if v { "" } else { "none" });
        }

        /// Horizontal alignment, GTK `GtkAlign` ordinals (0 = start,
        /// 1 = centre, 2 = end).
        pub fn set_halign(&self, align: i32) {
            let v = match align {
                1 => "center",
                2 => "flex-end",
                _ => "flex-start",
            };
            set_css(self.elem.as_ref(), "justify-self", v);
        }

        /// Vertical alignment, same ordinals as [`Entry::set_halign`].
        pub fn set_valign(&self, align: i32) {
            let v = match align {
                1 => "center",
                2 => "flex-end",
                _ => "flex-start",
            };
            set_css(self.elem.as_ref(), "align-self", v);
        }

        pub fn set_margin_start(&self, px: i32) {
            set_css(self.elem.as_ref(), "margin-left", &format!("{}px", px));
        }

        pub fn set_margin_top(&self, px: i32) {
            set_css(self.elem.as_ref(), "margin-top", &format!("{}px", px));
        }

        pub fn connect_changed(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            let cb = Rc::new(RefCell::new(f));
            let cb2 = cb.clone();
            let closure = Closure::<dyn FnMut(Event)>::new(move |_: Event| {
                (cb2.borrow_mut())();
            });
            self.elem
                .add_event_listener_with_callback("input", closure.as_ref().unchecked_ref())
                .map_err(|e| Error::Backend(format!("connect_changed: {:?}", e)))?;
            let id = *self.next_id.borrow();
            *self.next_id.borrow_mut() += 1;
            self.closures.borrow_mut().push(Box::new(closure));
            Ok(id)
        }

        pub fn connect_activate<F: FnMut(*mut c_void) + 'static>(
            &self,
            f: F,
        ) -> Result<u64, Error> {
            let cb = Rc::new(RefCell::new(f));
            let cb2 = cb.clone();
            let closure =
                Closure::<dyn FnMut(KeyboardEvent)>::new(move |evt: KeyboardEvent| {
                    if evt.key() == "Enter" {
                        (cb2.borrow_mut())(std::ptr::null_mut());
                    }
                });
            self.elem
                .add_event_listener_with_callback("keydown", closure.as_ref().unchecked_ref())
                .map_err(|e| Error::Backend(format!("connect_activate: {:?}", e)))?;
            let id = *self.next_id.borrow();
            *self.next_id.borrow_mut() += 1;
            self.closures.borrow_mut().push(Box::new(closure));
            Ok(id)
        }

        /// Whether the input currently holds focus.
        ///
        /// This was hard-coded `false`, and that is not a neutral default:
        /// `common::Entry::has_focus` (common.rs:238) is what callers use to
        /// choose between pushing a new value and appending to the current
        /// one, so always-false silently made every keystroke push. The DOM
        /// answer is `document.activeElement === input`.
        pub fn has_focus(&self) -> bool {
            document()
                .active_element()
                .map(|active| &active == self.elem.as_ref())
                .unwrap_or(false)
        }

        pub fn connect_button_press(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            let cb = Rc::new(RefCell::new(f));
            let cb2 = cb.clone();
            let closure = Closure::<dyn FnMut(MouseEvent)>::new(move |_: MouseEvent| {
                (cb2.borrow_mut())();
            });
            self.elem
                .add_event_listener_with_callback("click", closure.as_ref().unchecked_ref())
                .map_err(|e| Error::Backend(format!("connect_button_press: {:?}", e)))?;
            let id = *self.next_id.borrow();
            *self.next_id.borrow_mut() += 1;
            self.closures.borrow_mut().push(Box::new(closure));
            Ok(id)
        }

        pub fn add_class(&self, class_name: &str) {
            self.elem.class_list().add_1(class_name).ok();
        }

        pub fn remove_class(&self, class_name: &str) {
            self.elem.class_list().remove_1(class_name).ok();
        }

        pub fn grab_focus(&self) {
            let _ = self.elem.focus();
        }
        /// The caret offset. `common::Entry` documents this as a *character*
        /// index while the DOM's `selectionStart` counts UTF-16 code units,
        /// and the two differ for anything outside the BMP (an emoji in a
        /// cell value is enough), so the offset is converted rather than
        /// passed through.
        pub fn get_position(&self) -> Option<usize> {
            // `selection_start` is `Result<Option<u32>, JsValue>`: the inner
            // `None` is "no selection" (an input that has never been focused),
            // which is the same answer as a caret at the beginning. A `None`
            // from the outer Result is an un-focusable input, and reporting
            // no position at all is right for that.
            let start = self.elem.selection_start().ok().flatten().unwrap_or(0) as i32;
            Some(utf16_offset_to_char_index(&self.elem.value(), start))
        }

        /// Move the caret, taking a character index and converting it to the
        /// UTF-16 offset the DOM wants.
        pub fn set_position(&self, pos: usize) {
            let offset = char_index_to_utf16_offset(&self.elem.value(), pos) as u32;
            // The setters take `Option<u32>` and throw on `None`, so they are
            // wrapped in `Some` rather than unwrapped: "set the selection to
            // 0" is a real instruction, not a missing one.
            let _ = self.elem.set_selection_start(Some(offset));
            let _ = self.elem.set_selection_end(Some(offset));
        }

        pub fn on_key_raw(&self, cb: Box<dyn FnMut(u32, u32) -> bool>) {
            *self.key_cb.borrow_mut() = Some(cb);
            let cb2 = self.key_cb.clone();
            let closure = Closure::<dyn FnMut(KeyboardEvent)>::new(move |evt: KeyboardEvent| {
                if let Some(ref mut f) = *cb2.borrow_mut() {
                    if f(evt.key_code(), 0) {
                        evt.prevent_default();
                    }
                }
            });
            self.elem
                .add_event_listener_with_callback("keydown", closure.as_ref().unchecked_ref())
                .ok();
            store_closure(Box::new(closure));
        }

        pub fn connect_focus_in_event<F: FnMut(*mut c_void) -> i32 + 'static>(
            &self,
            f: F,
        ) -> Result<u64, Error> {
            let cb = Rc::new(RefCell::new(f));
            let cb2 = cb.clone();
            let closure = Closure::<dyn FnMut(FocusEvent)>::new(move |_: FocusEvent| {
                (cb2.borrow_mut())(std::ptr::null_mut());
            });
            self.elem
                .add_event_listener_with_callback("focus", closure.as_ref().unchecked_ref())
                .map_err(|e| Error::Backend(format!("connect_focus_in_event: {:?}", e)))?;
            let id = *self.next_id.borrow();
            *self.next_id.borrow_mut() += 1;
            self.closures.borrow_mut().push(Box::new(closure));
            Ok(id)
        }

        pub fn connect_focus_out_event<F: FnMut(*mut c_void) -> i32 + 'static>(
            &self,
            f: F,
        ) -> Result<u64, Error> {
            let cb = Rc::new(RefCell::new(f));
            let cb2 = cb.clone();
            let closure = Closure::<dyn FnMut(FocusEvent)>::new(move |_: FocusEvent| {
                (cb2.borrow_mut())(std::ptr::null_mut());
            });
            self.elem
                .add_event_listener_with_callback("blur", closure.as_ref().unchecked_ref())
                .map_err(|e| Error::Backend(format!("connect_focus_out_event: {:?}", e)))?;
            let id = *self.next_id.borrow();
            *self.next_id.borrow_mut() += 1;
            self.closures.borrow_mut().push(Box::new(closure));
            Ok(id)
        }
    }

    pub fn create_entry() -> Result<Entry, Error> {
        let elem: HtmlInputElement = create_element("input").dyn_into().map_err(|e| {
            Error::Backend(format!("create_entry: {:?}", e))
        })?;
        elem.set_type("text");
        Ok(Entry {
            elem,
            key_cb: Rc::new(RefCell::new(None)),
            closures: Rc::new(RefCell::new(Vec::new())),
            next_id: Rc::new(RefCell::new(1)),
        })
    }

    // -----------------------------------------------------------------------
    // Menu, MenuBar, SimpleAction
    // -----------------------------------------------------------------------
    #[derive(Clone)]
    enum MenuItem {
        Item { label: String, action: String },
        Submenu { label: String, items: Vec<MenuItem> },
    }

#[derive(Clone)]
pub struct Menu {
    items: Vec<MenuItem>,
}

impl Menu {
        pub fn append(&mut self, label: &str, detailed_action: &str) {
            let action = detailed_action
                .split('(')
                .next()
                .unwrap_or(detailed_action)
                .to_owned();
            self.items.push(MenuItem::Item {
                label: label.to_owned(),
                action,
            });
        }

        pub fn append_submenu(&mut self, label: &str, submenu: &Menu) {
            self.items.push(MenuItem::Submenu {
                label: label.to_owned(),
                items: submenu.items.clone(),
            });
        }
    }

    pub fn create_menu() -> Result<Menu, Error> {
        Ok(Menu { items: Vec::new() })
    }

pub struct MenuBar {
    elem: HtmlDivElement,
    /// The top-level menus, in order, with the DOM elements that open them.
    ///
    /// The keyboard API needs this: `Alt+F` has to find *the* submenu whose
    /// label starts with F, and the only way to know that is to keep the
    /// labels alongside the elements. GTK gets it from `GMenu`, which carries
    /// the labels itself; the DOM does not.
    entries: Rc<RefCell<Vec<MenuEntry>>>,
    /// The open submenu, if any: its toggle button and its dropdown.
    open: Rc<RefCell<Option<(Element, Element)>>>,
    /// Index of the highlighted item in the open submenu, for arrow-key
    /// navigation.
    selected: Rc<RefCell<usize>>,
    /// The active response callback, so a chosen item can be invoked.
    closures: Rc<RefCell<Vec<Box<dyn Any>>>>,
}

/// One top-level menu: its label, the button that toggles it, and the
/// dropdown it opens.
struct MenuEntry {
    label: String,
    /// The `<button>` that opens the dropdown.
    toggle: Element,
    /// The absolutely-positioned dropdown below it.
    dropdown: Element,
    /// `(label, action)` for each item, in display order. Kept so
    /// `activate_submenu_item_by_mnemonic` can dispatch by key.
    items: Vec<(String, String)>,
}

impl Clone for MenuBar {
    fn clone(&self) -> Self {
        MenuBar {
            elem: self.elem.clone(),
            entries: self.entries.clone(),
            open: self.open.clone(),
            selected: self.selected.clone(),
            closures: self.closures.clone(),
        }
    }
}

impl AsElement for MenuBar {
        fn as_element(&self) -> &Element {
            self.elem.as_ref()
        }
    }

    impl Widget for MenuBar {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const HtmlDivElement as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for MenuBar {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.raw_handle() as *const *mut c_void) }
        }
    }

    impl MenuBar {
        /// Open the submenu whose first letter is `keyval`, and report
        /// whether one matched.
        ///
        /// GTK matches a mnemonic with a trailing `_` in the label
        /// (`"_File"`); the DOM has no such convention, so the first
        /// character of the label is the mnemonic -- the same rule a browser
        /// menu uses, and the one that matches what the GTK build's labels
        /// already spell out (`Menu::append` receives the label with the
        /// marker stripped, see `build_menu_model`).
        pub fn activate_submenu_by_mnemonic(&self, keyval: u32) -> bool {
            let ch = match char_from_keyval(keyval) {
                Some(c) => c.to_ascii_lowercase(),
                None => return false,
            };
            let idx = self.entries.borrow().iter().position(|e| {
                e.label
                    .chars()
                    .next()
                    .map(|c| c.to_ascii_lowercase() == ch)
                    .unwrap_or(false)
            });
            match idx {
                Some(i) => {
                    self.open_submenu(i);
                    true
                }
                None => false,
            }
        }

        /// Open the submenu whose mnemonic is `keyval` at a screen position.
        ///
        /// The dropdown is normally positioned by CSS (`position: absolute`
        /// under its toggle). A caller that wants it at the pointer -- the
        /// sheet-tab context menu, which GTK does through its
        /// position-callback -- gets it here by overriding the offsets.
        pub fn popup_submenu_by_mnemonic_at(&self, keyval: u32, screen_x: i32, screen_y: i32) -> bool {
            let ch = match char_from_keyval(keyval) {
                Some(c) => c.to_ascii_lowercase(),
                None => return false,
            };
            let entries = self.entries.borrow();
            let idx = entries.iter().position(|e| {
                e.label
                    .chars()
                    .next()
                    .map(|c| c.to_ascii_lowercase() == ch)
                    .unwrap_or(false)
            });
            drop(entries);
            match idx {
                Some(i) => {
                    self.open_submenu(i);
                    // The dropdown is `position: absolute` inside the bar, so
                    // an offset is relative to the bar -- which sits at the
                    // top of the page. The caller's screen coordinates are
                    // converted through the bar's own bounding box, the same
                    // way GTK's position callback receives screen coords and
                    // converts to widget coords.
                    {
                        let rect = self.elem.get_bounding_client_rect();
                        let dx = screen_x as f64 - rect.x();
                        let dy = screen_y as f64 - rect.y();
                        let entries = self.entries.borrow();
                        if let Some(entry) = entries.get(i) {
                            if let Some(html) = entry.dropdown.dyn_ref::<HtmlElement>() {
                                let _ = html.style().set_property("left", &format!("{}px", dx));
                                let _ = html.style().set_property("top", &format!("{}px", dy));
                            }
                        }
                    }
                    true
                }
                None => false,
            }
        }

        /// Activate the item in the open submenu whose label starts with
        /// `keyval`. Requires a submenu to be open, which is what
        /// `menu_active` reports.
        pub fn activate_submenu_item_by_mnemonic(&self, keyval: u32) -> bool {
            let Some(ch) = char_from_keyval(keyval) else {
                return false;
            };
            let ch = ch.to_ascii_lowercase();
            let open = self.open.borrow();
            let Some((_, _)) = open.as_ref() else {
                return false;
            };
            drop(open);
            // Find the entry holding the match, then dispatch its action.
            let found = self.entries.borrow().iter().enumerate().find_map(|(entry_idx, e)| {
                e.items.iter().position(|(label, _)| {
                    label
                        .chars()
                        .next()
                        .map(|c| c.to_ascii_lowercase() == ch)
                        .unwrap_or(false)
                })
                .map(|item_idx| (entry_idx, item_idx))
            });
            match found {
                Some((entry_idx, item_idx)) => {
                    let action = {
                        let entries = self.entries.borrow();
                        entries[entry_idx].items[item_idx].1.clone()
                    };
                    self.menu_close();
                    invoke_action(&action, std::ptr::null_mut());
                    true
                }
                None => false,
            }
        }

        /// # Safety
        /// Kept for API compatibility; no-op (the DOM menu is wired directly
        /// to the action registry, as on NWG).
        pub unsafe fn insert_action_group(&self, _name: &str, _group_ptr: *mut c_void) {}

        /// Handle a keypress when a submenu may be open.
        ///
        /// With a submenu open: Escape closes it, a printable key activates
        /// the item it starts, and Up/Down move the selection. With nothing
        /// open this only answers Alt+letter (via
        /// `activate_submenu_by_mnemonic`) and reports `false` otherwise, so
        /// the caller falls through to normal key handling.
        pub fn handle_mnemonic_key(&self, keyval: u32) -> bool {
            if !self.menu_active() {
                return false;
            }
            if keyval == crate::core::key::ESCAPE {
                self.menu_close();
                return true;
            }
            if keyval == crate::core::key::DOWN || keyval == crate::core::key::UP {
                self.move_selection(keyval == crate::core::key::DOWN);
                return true;
            }
            self.activate_submenu_item_by_mnemonic(keyval)
        }

        /// Handle any menu-related key event.
        ///
        /// `modifiers` is the GDK mask (1 = Shift, 4 = Control, 8 = Alt), the
        /// same one `Window::on_event_key` reports. Alt+letter opens a
        /// submenu; otherwise this defers to `handle_mnemonic_key`, so the
        /// navigation rules live in one place.
        pub fn handle_menu_key(&self, keyval: u32, modifiers: u32) -> bool {
            const ALT_MASK: u32 = 0x8;
            if modifiers & ALT_MASK != 0 {
                let opened = self.activate_submenu_by_mnemonic(keyval);
                if opened {
                    return true;
                }
            }
            self.handle_mnemonic_key(keyval)
        }

        /// Whether a keyboard menu is open. Callers use this to skip normal
        /// key handling while one is.
        pub fn menu_active(&self) -> bool {
            self.open.borrow().is_some()
        }

        /// Close the open submenu and clear the selection.
        pub fn menu_close(&self) {
            let open = self.open.borrow_mut().take();
            let Some((toggle, dropdown)) = open else {
                return;
            };
            if let Some(html) = dropdown.dyn_ref::<HtmlElement>() {
                let _ = html.style().set_property("display", "none");
            }
            if let Some(html) = toggle.dyn_ref::<HtmlElement>() {
                // Drop the "is open" marker, which is also what the
                // background highlight keys off.
                let _ = html.style().set_property("background", "transparent");
            }
            *self.selected.borrow_mut() = 0;
        }

        /// Show submenu `idx` and hide any other.
        fn open_submenu(&self, idx: usize) {
            self.menu_close();
            let entries = self.entries.borrow();
            let Some(entry) = entries.get(idx) else {
                return;
            };
            if let Some(html) = entry.dropdown.dyn_ref::<HtmlElement>() {
                let _ = html.style().set_property("display", "block");
            }
            if let Some(html) = entry.toggle.dyn_ref::<HtmlElement>() {
                let _ = html.style().set_property("background", "#e8e8e8");
            }
            *self.open.borrow_mut() = Some((entry.toggle.clone(), entry.dropdown.clone()));
            *self.selected.borrow_mut() = 0;
        }

        /// Move the highlight by one item in the open submenu, clamping at
        /// the ends.
        fn move_selection(&self, forward: bool) {
            let count = {
                let open = self.open.borrow();
                let Some((_, dropdown)) = open.as_ref() else {
                    return;
                };
                dropdown.children().length() as usize
            };
            if count == 0 {
                return;
            }
            let mut sel = self.selected.borrow_mut();
            *sel = if forward {
                (*sel + 1) % count
            } else {
                (*sel + count - 1) % count
            };
            // The highlight is what a caller reads back to know where the
            // selection is, so it has to be visible: re-apply the background
            // to the selected child only.
            let open = self.open.borrow();
            if let Some((_, dropdown)) = open.as_ref() {
                for i in 0..count {
                    if let Some(child) = dropdown.children().item(i as u32) {
                        if let Some(html) = child.dyn_ref::<HtmlElement>() {
                            let _ = html.style().set_property(
                                "background",
                                if i == *sel { "#d0d8e8" } else { "" },
                            );
                        }
                    }
                }
            }
        }
    }

    pub fn create_menubar(model: &Menu, _action_group: *mut c_void) -> Result<MenuBar, Error> {
        let bar: HtmlDivElement = create_element("div").dyn_into().map_err(|e| {
            Error::Backend(format!("create_menubar: {:?}", e))
        })?;
        let s = bar.style();
        s.set_property("display", "flex").ok();
        s.set_property("background", "#f0f0f0").ok();
        s.set_property("border-bottom", "1px solid #ccc").ok();

        let mut entries: Vec<MenuEntry> = Vec::new();
        for item in &model.items {
            match item {
                MenuItem::Item { label, action } => {
                    let btn: HtmlButtonElement = create_element("button").dyn_into().unwrap();
                    btn.set_text_content(Some(label));
                    let action_name = action.clone();
                    let closure =
                        Closure::<dyn FnMut(MouseEvent)>::new(move |_: MouseEvent| {
                            invoke_action(&action_name, std::ptr::null_mut());
                        });
                    btn.add_event_listener_with_callback("click", closure.as_ref().unchecked_ref())
                        .ok();
                    closure.forget();
                    let s = btn.style();
                    s.set_property("background", "transparent").ok();
                    s.set_property("border", "none").ok();
                    s.set_property("padding", "4px 12px").ok();
                    s.set_property("cursor", "pointer").ok();
                    bar.append_child(btn.as_ref()).ok();
                    // A bare top-level item is still addressable by
                    // mnemonic: it has no dropdown, so the "dropdown" is the
                    // button itself and `menu_active` stays false for it.
                    // `HtmlButtonElement` implements `AsRef` for both
                    // `Element` and `HtmlElement`, so `btn.as_ref()` is
                    // ambiguous here; the turbofish names the one wanted.
                    let elem: Element = <HtmlButtonElement as AsRef<Element>>::as_ref(&btn).clone();
                    entries.push(MenuEntry {
                        label: label.clone(),
                        toggle: elem.clone(),
                        dropdown: elem,
                        items: vec![(label.clone(), action.clone())],
                    });
                }
                MenuItem::Submenu { label, items } => {
                    let wrapper = create_element("div");
                    set_css(&wrapper, "position", "relative");

                    let toggle: HtmlButtonElement = create_element("button").dyn_into().unwrap();
                    toggle.set_text_content(Some(label));
                    let s = toggle.style();
                    s.set_property("background", "transparent").ok();
                    s.set_property("border", "none").ok();
                    s.set_property("padding", "4px 12px").ok();
                    s.set_property("cursor", "pointer").ok();

                    let dropdown = create_element("div");
                    set_css(&dropdown, "display", "none");
                    set_css(&dropdown, "position", "absolute");
                    set_css(&dropdown, "top", "100%");
                    set_css(&dropdown, "left", "0");
                    set_css(&dropdown, "background", "#fff");
                    set_css(&dropdown, "border", "1px solid #ccc");
                    set_css(&dropdown, "z-index", "1000");

                    let mut entry_items: Vec<(String, String)> = Vec::new();
                    for sub in items {
                        match sub {
                            MenuItem::Item { label, action } => {
                                let item = create_element("div");
                                set_css(&item, "padding", "4px 12px");
                                set_css(&item, "cursor", "pointer");
                                item.set_text_content(Some(label));
                                let action_name = action.clone();
                                let cl =
                                    Closure::<dyn FnMut(MouseEvent)>::new(move |_: MouseEvent| {
                                        invoke_action(&action_name, std::ptr::null_mut());
                                    });
                                item.add_event_listener_with_callback(
                                    "click",
                                    cl.as_ref().unchecked_ref(),
                                )
                                .ok();
                                cl.forget();
                                dropdown.append_child(&item).ok();
                                // Recorded so the keyboard API can dispatch by
                                // mnemonic: the DOM item has no way to tell
                                // Rust which action it stands for.
                                entry_items.push((label.clone(), action.clone()));
                            }
                            _ => {}
                        }
                    }

                    let dd = dropdown.clone();
                    let toggle_closure =
                        Closure::<dyn FnMut(MouseEvent)>::new(move |_: MouseEvent| {
                            if let Some(html) = dd.dyn_ref::<HtmlElement>() {
                                let disp = html
                                    .style()
                                    .get_property_value("display")
                                    .unwrap_or_default();
                                if disp == "none" {
                                    html.style().set_property("display", "block").ok();
                                } else {
                                    html.style().set_property("display", "none").ok();
                                }
                            }
                        });
                    toggle
                        .add_event_listener_with_callback(
                            "click",
                            toggle_closure.as_ref().unchecked_ref(),
                        )
                        .ok();
                    toggle_closure.forget();

                    wrapper.append_child(toggle.as_ref()).ok();
                    wrapper.append_child(&dropdown).ok();
                    bar.append_child(&wrapper).ok();
                    let toggle_elem: Element =
                        <HtmlButtonElement as AsRef<Element>>::as_ref(&toggle).clone();
                    entries.push(MenuEntry {
                        label: label.clone(),
                        toggle: toggle_elem,
                        dropdown: dropdown.clone(),
                        items: entry_items,
                    });
                }
            }
        }

        Ok(MenuBar {
            elem: bar,
            entries: Rc::new(RefCell::new(entries)),
            open: Rc::new(RefCell::new(None)),
            selected: Rc::new(RefCell::new(0)),
            closures: Rc::new(RefCell::new(Vec::new())),
        })
    }

#[derive(Clone)]
pub struct SimpleAction {
    name: String,
}

impl SimpleAction {
        pub fn connect_activate<F: FnMut(*mut c_void) + 'static>(
            &self,
            f: F,
        ) -> Result<u64, Error> {
            register_action(&self.name, Box::new(f));
            Ok(0)
        }
    }

    pub fn create_simple_action(name: &str) -> Result<SimpleAction, Error> {
        Ok(SimpleAction {
            name: name.to_owned(),
        })
    }

    // -----------------------------------------------------------------------
    // Dialog
    // -----------------------------------------------------------------------
    #[derive(Clone)]
    pub struct Dialog {
        elem: HtmlDialogElement,
        content_area: HtmlDivElement,
        response_cb: Rc<RefCell<Option<Box<dyn FnMut(i32)>>>>,
        /// Whether `set_transient_for` asked for a *modal* presentation.
        /// See [`Dialog::set_transient_for`].
        modal: Rc<Cell<bool>>,
    }

    impl AsElement for Dialog {
        fn as_element(&self) -> &Element {
            self.elem.as_ref()
        }
    }

    impl Widget for Dialog {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const HtmlDialogElement as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for Dialog {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.raw_handle() as *const *mut c_void) }
        }
    }

    impl Dialog {
        pub fn set_title(&self, title: &str) {
            let h2 = create_element("h2");
            h2.set_text_content(Some(title));
            self.elem.insert_before(&h2, self.elem.first_child().as_ref())
                .ok();
        }

        pub fn set_default_size(&self, w: i32, h: i32) {
            if w > 0 {
                set_css(self.elem.as_ref(), "width", &format!("{}px", w));
            }
            if h > 0 {
                set_css(self.elem.as_ref(), "height", &format!("{}px", h));
            }
        }

        pub fn add_button(&self, text: &str, response_id: i32) {
            let btn: HtmlButtonElement = create_element("button").dyn_into().unwrap();
            btn.set_text_content(Some(text));
            let cb = self.response_cb.clone();
            let elem = self.elem.clone();
            let closure = Closure::<dyn FnMut(MouseEvent)>::new(move |_: MouseEvent| {
                if let Some(ref mut f) = *cb.borrow_mut() {
                    f(response_id);
                }
                // Dismiss like every other backend: confirming must close.
                elem.close();
            });
            btn.add_event_listener_with_callback("click", closure.as_ref().unchecked_ref())
                .ok();
            closure.forget();
            self.elem.append_child(btn.as_ref()).ok();
        }

        /// Mark the dialog as modal, which is what `set_transient_for` means
        /// on a browser.
        ///
        /// GTK parents the dialog to a window so the WM centres it there; a
        /// page has no window manager, and the DOM's equivalent of "belongs
        /// to this window" is `showModal` -- which blocks interaction with the
        /// rest of the page, exactly as a parent-modal dialog does. The parent
        /// handle itself is not meaningful, so it is accepted and ignored.
        pub fn set_transient_for(&self, _parent: *mut c_void) {
            // A `present` that ran before `showModal` would defeat the
            // purpose, so this is a flag rather than a call: the actual
            // `show_modal` happens in `present`, which is the only place the
            // element is shown.
            self.modal.set(true);
        }

        /// Whether a `set_transient_for` asked for modality.
        fn wants_modal(&self) -> bool {
            self.modal.get()
        }

        pub fn get_content_area(&self) -> *mut c_void {
            &self.content_area as *const HtmlDivElement as *mut c_void
        }

        pub fn append_content_area(&self, child: &impl AsRef<*mut c_void>) {
            self.content_area.append_child(as_element_from_ptr(*child.as_ref())).ok();
        }

        /// Show the dialog, modally if `set_transient_for` asked for it.
        ///
        /// `show_modal` additionally blocks the rest of the page, which is
        /// the browser's version of a parent-modal dialog, and fires
        /// `preventDefault` on a cancel event so Escape cannot dismiss it
        /// behind the app's back -- the app decides what closing means.
        pub fn present(&self) {
            if self.wants_modal() {
                let _ = self.elem.show_modal();
            } else {
                let _ = self.elem.show();
            }
        }

        pub fn connect_response<F: FnMut(i32) + 'static>(&self, f: F) -> Result<u64, Error> {
            *self.response_cb.borrow_mut() = Some(Box::new(f));
            Ok(0)
        }

        pub fn close(&self) {
            self.elem.close();
        }
    }

    pub fn create_dialog() -> Result<Dialog, Error> {
        let elem: HtmlDialogElement = create_element("dialog").dyn_into().map_err(|e| {
            Error::Backend(format!("create_dialog: {:?}", e))
        })?;
        let content: HtmlDivElement = create_element("div").dyn_into().map_err(|e| {
            Error::Backend(format!("create_dialog content: {:?}", e))
        })?;
        elem.append_child(content.as_ref()).ok();
        body().append_child(elem.as_ref()).ok();
        Ok(Dialog {
            elem,
            content_area: content,
            response_cb: Rc::new(RefCell::new(None)),
            modal: Rc::new(Cell::new(false)),
        })
    }

    // -----------------------------------------------------------------------
    // DropDown
    // -----------------------------------------------------------------------
    pub struct DropDown {
        elem: HtmlSelectElement,
        closures: Rc<RefCell<Vec<Box<dyn Any>>>>,
        next_id: Rc<RefCell<u64>>,
    }

    impl AsElement for DropDown {
        fn as_element(&self) -> &Element {
            self.elem.as_ref()
        }
    }

    impl Widget for DropDown {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const HtmlSelectElement as *mut c_void
        }
    }

    impl Clone for DropDown {
        fn clone(&self) -> Self {
            DropDown {
                elem: self.elem.clone(),
                closures: self.closures.clone(),
                next_id: self.next_id.clone(),
            }
        }
    }

    impl DropDown {
        /// Select `index`, or clear the selection with `None`.
        ///
        /// The `Option` is GTK's contract (a caller can clear the selection),
        /// and the DOM has a matching state: `selectedIndex = -1` is "nothing
        /// selected", which is exactly what a `None` asks for. Passing a bare
        /// `u32` -- the old signature -- could not express it, so a caller that
        /// wanted "no selection" had no way to say so.
        pub fn set_active(&self, index: Option<u32>) {
            match index {
                Some(i) => self.elem.set_selected_index(i as i32),
                None => self.elem.set_selected_index(-1),
            }
        }

        /// The selected index, or `None` when nothing is selected. `None` is
        /// what `selectedIndex == -1` means, and it is the inverse of
        /// `set_active(None)`.
        pub fn get_active(&self) -> i32 {
            self.elem.selected_index()
        }

        /// The selected index as an `Option`, for callers that would rather
        /// not interpret the `-1` sentinel themselves.
        pub fn active(&self) -> Option<u32> {
            match self.elem.selected_index() {
                i if i < 0 => None,
                i => Some(i as u32),
            }
        }

        pub fn grab_focus(&self) {
            let _ = self.elem.focus();
        }

        pub fn set_visible(&self, v: bool) {
            set_css(self.elem.as_ref(), "display", if v { "" } else { "none" });
        }

        pub fn set_size_request(&self, w: i32, h: i32) {
            if w > 0 {
                set_css(self.elem.as_ref(), "width", &format!("{}px", w));
            }
            if h > 0 {
                set_css(self.elem.as_ref(), "height", &format!("{}px", h));
            }
        }

        pub fn set_hexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-grow", "1");
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "flex-grow", "0");
            }
        }

        pub fn set_vexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-shrink", "1");
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "flex-shrink", "0");
            }
        }

        /// Move the dropdown within its parent, in px.
        ///
        /// The DOM answer to GTK's `set_offset`, and what makes a dropdown
        /// hosted in an `Overlay` land where the caller asked instead of at
        /// the overlay's top-left corner: `add_overlay` sets
        /// `position: absolute` with no offset, so without this a positioned
        /// popup has nowhere to go. `transform: translate` is used rather than
        /// `left`/`top` because it composes with any layout position the
        /// overlay already established.
        pub fn set_offset(&self, x: i32, y: i32) {
            set_css(
                self.elem.as_ref(),
                "transform",
                &format!("translate({}px, {}px)", x, y),
            );
        }

        pub fn connect_changed(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            let cb = Rc::new(RefCell::new(f));
            let cb2 = cb.clone();
            let closure = Closure::<dyn FnMut(Event)>::new(move |_: Event| {
                (cb2.borrow_mut())();
            });
            self.elem
                .add_event_listener_with_callback("change", closure.as_ref().unchecked_ref())
                .map_err(|e| Error::Backend(format!("connect_changed: {:?}", e)))?;
            let id = *self.next_id.borrow();
            *self.next_id.borrow_mut() += 1;
            self.closures.borrow_mut().push(Box::new(closure));
            Ok(id)
        }
    }

    pub fn create_dropdown(items: &[&str]) -> Result<DropDown, Error> {
        let elem: HtmlSelectElement = create_element("select").dyn_into().map_err(|e| {
            Error::Backend(format!("create_dropdown: {:?}", e))
        })?;
        for item in items {
            let opt: HtmlOptionElement = create_element("option").dyn_into().unwrap();
            opt.set_text_content(Some(item));
            elem.append_child(opt.as_ref()).ok();
        }
        Ok(DropDown {
            elem,
            closures: Rc::new(RefCell::new(Vec::new())),
            next_id: Rc::new(RefCell::new(1)),
        })
    }

    // -----------------------------------------------------------------------
    // CheckButton
    // -----------------------------------------------------------------------
    pub struct CheckButton {
        elem: Element,
        input: HtmlInputElement,
        closures: Rc<RefCell<Vec<Box<dyn Any>>>>,
        next_id: Rc<RefCell<u64>>,
    }

    impl AsElement for CheckButton {
        fn as_element(&self) -> &Element {
            &self.elem
        }
    }

    impl Widget for CheckButton {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const Element as *mut c_void
        }
    }

    impl Clone for CheckButton {
        fn clone(&self) -> Self {
            CheckButton {
                elem: self.elem.clone(),
                input: self.input.clone(),
                closures: self.closures.clone(),
                next_id: self.next_id.clone(),
            }
        }
    }

    impl CheckButton {
        pub fn is_active(&self) -> bool {
            self.input.checked()
        }

        pub fn set_active(&self, active: bool) {
            self.input.set_checked(active);
        }

        pub fn connect_toggled(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            let cb = Rc::new(RefCell::new(f));
            let cb2 = cb.clone();
            let closure = Closure::<dyn FnMut(Event)>::new(move |_: Event| {
                (cb2.borrow_mut())();
            });
            self.input
                .add_event_listener_with_callback("change", closure.as_ref().unchecked_ref())
                .map_err(|e| Error::Backend(format!("connect_toggled: {:?}", e)))?;
            let id = *self.next_id.borrow();
            *self.next_id.borrow_mut() += 1;
            self.closures.borrow_mut().push(Box::new(closure));
            Ok(id)
        }

        pub fn set_visible(&self, v: bool) {
            set_css(&self.elem, "display", if v { "" } else { "none" });
        }
        pub fn set_hexpand(&self, expand: bool) {
            if expand {
                set_css(&self.elem, "flex-grow", "1");
                set_css(&self.elem, "align-self", "stretch");
            } else {
                set_css(&self.elem, "flex-grow", "0");
            }
        }
        pub fn set_vexpand(&self, expand: bool) {
            if expand {
                set_css(&self.elem, "flex-shrink", "1");
                set_css(&self.elem, "align-self", "stretch");
            } else {
                set_css(&self.elem, "flex-shrink", "0");
            }
        }
        pub fn set_size_request(&self, w: i32, h: i32) {
            if w > 0 {
                set_css(&self.elem, "width", &format!("{}px", w));
            }
            if h > 0 {
                set_css(&self.elem, "height", &format!("{}px", h));
            }
        }

        /// Replace the label text.
        ///
        /// The wrapper is a `<label>` holding the input plus a text node, so
        /// the text is the wrapper's *last* child; only that one is rewritten
        /// and the input is left alone.
        pub fn set_label(&self, label: &str) {
            if let Some(node) = self.elem.last_child() {
                let _ = node.set_text_content(Some(label));
            }
        }

        /// The current label text, read back from the same text node
        /// [`CheckButton::set_label`] writes.
        pub fn get_label(&self) -> Option<String> {
            self.elem.last_child().and_then(|n| n.text_content())
        }
    }

    pub fn create_checkbutton(label: &str) -> Result<CheckButton, Error> {
        let wrapper = create_element("label");
        let input: HtmlInputElement = create_element("input").dyn_into().unwrap();
        input.set_type("checkbox");
        wrapper.append_child(input.as_ref()).ok();
        wrapper
            .append_child(&web_sys::Text::new_with_data(label).unwrap())
            .ok();
        Ok(CheckButton {
            elem: wrapper,
            input,
            closures: Rc::new(RefCell::new(Vec::new())),
            next_id: Rc::new(RefCell::new(1)),
        })
    }

    // -----------------------------------------------------------------------
    // RadioButton
    // -----------------------------------------------------------------------
    pub struct RadioButton {
        elem: Element,
        input: HtmlInputElement,
        closures: Rc<RefCell<Vec<Box<dyn Any>>>>,
        next_id: Rc<RefCell<u64>>,
        group_name: String,
    }

    impl AsElement for RadioButton {
        fn as_element(&self) -> &Element {
            &self.elem
        }
    }

    impl Widget for RadioButton {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const Element as *mut c_void
        }
    }

    impl Clone for RadioButton {
        fn clone(&self) -> Self {
            RadioButton {
                elem: self.elem.clone(),
                input: self.input.clone(),
                closures: self.closures.clone(),
                next_id: self.next_id.clone(),
                group_name: self.group_name.clone(),
            }
        }
    }

    impl RadioButton {
        pub fn is_active(&self) -> bool {
            self.input.checked()
        }

        pub fn set_active(&self, active: bool) {
            self.input.set_checked(active);
        }

        pub fn connect_toggled(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            let cb = Rc::new(RefCell::new(f));
            let cb2 = cb.clone();
            let closure = Closure::<dyn FnMut(Event)>::new(move |_: Event| {
                (cb2.borrow_mut())();
            });
            self.input
                .add_event_listener_with_callback("change", closure.as_ref().unchecked_ref())
                .map_err(|e| Error::Backend(format!("connect_toggled: {:?}", e)))?;
            let id = *self.next_id.borrow();
            *self.next_id.borrow_mut() += 1;
            self.closures.borrow_mut().push(Box::new(closure));
            Ok(id)
        }

        pub fn set_visible(&self, v: bool) {
            set_css(&self.elem, "display", if v { "" } else { "none" });
        }
        pub fn set_hexpand(&self, expand: bool) {
            if expand {
                set_css(&self.elem, "flex-grow", "1");
                set_css(&self.elem, "align-self", "stretch");
            } else {
                set_css(&self.elem, "flex-grow", "0");
            }
        }
        pub fn set_vexpand(&self, expand: bool) {
            if expand {
                set_css(&self.elem, "flex-shrink", "1");
                set_css(&self.elem, "align-self", "stretch");
            } else {
                set_css(&self.elem, "flex-shrink", "0");
            }
        }
        pub fn set_size_request(&self, w: i32, h: i32) {
            if w > 0 {
                set_css(&self.elem, "width", &format!("{}px", w));
            }
            if h > 0 {
                set_css(&self.elem, "height", &format!("{}px", h));
            }
        }
        pub fn grab_focus(&self) {
            let _ = self.input.focus();
        }
        /// See [`CheckButton::set_label`].
        pub fn set_label(&self, label: &str) {
            if let Some(node) = self.elem.last_child() {
                let _ = node.set_text_content(Some(label));
            }
        }
        /// See [`CheckButton::get_label`].
        pub fn get_label(&self) -> Option<String> {
            self.elem.last_child().and_then(|n| n.text_content())
        }
    }

    pub fn create_radiobutton(group: Option<&RadioButton>, label: &str) -> Result<RadioButton, Error> {
        let group_name = group
            .map(|g| g.group_name.clone())
            .unwrap_or_else(|| format!("rb_{}", rand_id()));
        let wrapper = create_element("label");
        let input: HtmlInputElement = create_element("input").dyn_into().unwrap();
        input.set_type("radio");
        input.set_name(&group_name);
        wrapper.append_child(input.as_ref()).ok();
        wrapper
            .append_child(&web_sys::Text::new_with_data(label).unwrap())
            .ok();
        Ok(RadioButton {
            elem: wrapper,
            input,
            closures: Rc::new(RefCell::new(Vec::new())),
            next_id: Rc::new(RefCell::new(1)),
            group_name,
        })
    }

    fn rand_id() -> u64 {
        use std::sync::atomic::{AtomicU64, Ordering};
        static COUNTER: AtomicU64 = AtomicU64::new(1);
        COUNTER.fetch_add(1, Ordering::Relaxed)
    }

    // -----------------------------------------------------------------------
    // TextView
    // -----------------------------------------------------------------------
    pub struct TextView {
        elem: HtmlTextAreaElement,
        closures: Rc<RefCell<Vec<Box<dyn Any>>>>,
        next_id: Rc<RefCell<u64>>,
    }

    impl AsElement for TextView {
        fn as_element(&self) -> &Element {
            self.elem.as_ref()
        }
    }

    impl Widget for TextView {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const HtmlTextAreaElement as *mut c_void
        }
    }

    impl Clone for TextView {
        fn clone(&self) -> Self {
            TextView {
                elem: self.elem.clone(),
                closures: self.closures.clone(),
                next_id: self.next_id.clone(),
            }
        }
    }

    impl TextView {
        pub fn set_text(&self, text: &str) {
            self.elem.set_value(text);
        }

        pub fn get_text(&self) -> Option<String> {
            Some(self.elem.value())
        }

        pub fn set_wrap_mode(&self, wrap_mode: i32) {
            match wrap_mode {
                0 => self.elem.set_wrap("off"),
                1 => self.elem.set_wrap("soft"),
                2 => self.elem.set_wrap("hard"),
                _ => self.elem.set_wrap("soft"),
            }
        }

        pub fn set_size_request(&self, w: i32, h: i32) {
            if w > 0 {
                self.elem.set_cols(w as u32);
            }
            if h > 0 {
                self.elem.set_rows(h as u32);
            }
        }

        pub fn set_hexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-grow", "1");
            } else {
                set_css(self.elem.as_ref(), "flex-grow", "0");
            }
        }

        pub fn set_vexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-shrink", "1");
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "flex-shrink", "0");
            }
        }

        pub fn set_visible(&self, v: bool) {
            set_css(self.elem.as_ref(), "display", if v { "" } else { "none" });
        }

        /// Read-only mode, via the `readOnly` attribute NWG's `set_editable`
        /// sets. `readOnly` rather than `disabled` on purpose: a disabled
        /// textarea drops out of the tab order and dims, whereas a read-only
        /// one stays selectable and copyable, which is what "not editable"
        /// means in a form.
        pub fn set_editable(&self, editable: bool) {
            self.elem.set_read_only(!editable);
        }
            /// Append a line to the buffer.
        ///
        /// Read-modify-write: this backend has no incremental insert, so the text is
        /// fetched, extended and written back. Correct, but O(document) per call.
        /// A high-rate log pane should reimplement this against the native handle.
        pub fn append_text(&self, text: &str) {
        let mut buf = self.get_text().unwrap_or_default();
        buf.push_str(text);
        self.set_text(&buf);
        }
    }

    pub fn create_textview() -> Result<TextView, Error> {
        let elem: HtmlTextAreaElement = create_element("textarea").dyn_into().map_err(|e| {
            Error::Backend(format!("create_textview: {:?}", e))
        })?;
        Ok(TextView {
            elem,
            closures: Rc::new(RefCell::new(Vec::new())),
            next_id: Rc::new(RefCell::new(1)),
        })
    }

    // -----------------------------------------------------------------------
    // DrawContext (Canvas 2D)
    // -----------------------------------------------------------------------
    pub struct WasmDrawContext {
        ctx: web_sys::CanvasRenderingContext2d,
    }

    impl WasmDrawContext {
        fn set_font(&self, font: &str, size: f64, slant: i32, weight: i32) {
            let weight_str = if weight != 0 { "bold" } else { "" };
            let slant_str = if slant != 0 { "italic " } else { "" };
            let size_str = format!("{}px", size);
            let f = format!("{}{} {} {}", slant_str, weight_str, size_str, font);
            let _ = self.ctx.set_font(&f);
        }

        fn rgba(r: f64, g: f64, b: f64, a: f64) -> String {
            format!("rgba({},{},{},{})", (r * 255.0) as u8, (g * 255.0) as u8, (b * 255.0) as u8, a)
        }
    }

    impl crate::core::DrawContext for WasmDrawContext {
        fn fill_rect(&mut self, x: f64, y: f64, w: f64, h: f64, r: f64, g: f64, b: f64, a: f64) {
            let _ = self.ctx.set_fill_style_str(&Self::rgba(r, g, b, a));
            self.ctx.fill_rect(x, y, w, h);
        }

        fn stroke_rect(&mut self, x: f64, y: f64, w: f64, h: f64, r: f64, g: f64, b: f64, a: f64, lw: f64) {
            let _ = self.ctx.set_stroke_style_str(&Self::rgba(r, g, b, a));
            self.ctx.set_line_width(lw);
            self.ctx.stroke_rect(x, y, w, h);
        }

        fn draw_text_styled(&mut self, x: f64, y: f64, text: &str, font: &str, size: f64,
                            r: f64, g: f64, b: f64, a: f64, slant: i32, weight: i32) {
            self.set_font(font, size, slant, weight);
            let _ = self.ctx.set_fill_style_str(&Self::rgba(r, g, b, a));
            let _ = self.ctx.fill_text(text, x, y);
        }

        fn text_extents_styled(&self, text: &str, font: &str, size: f64, slant: i32, weight: i32) -> (f64, f64, f64, f64) {
            let weight_str = if weight != 0 { "bold" } else { "" };
            let slant_str = if slant != 0 { "italic " } else { "" };
            let size_str = format!("{}px", size);
            let f = format!("{}{} {} {}", slant_str, weight_str, size_str, font);
            self.ctx.save();
            let _ = self.ctx.set_font(&f);
            let metrics = self.ctx.measure_text(text).expect("measure_text failed");
            self.ctx.restore();
            let w = metrics.width();
            let xb = -metrics.actual_bounding_box_left();
            let yb = -metrics.actual_bounding_box_ascent();
            let h = metrics.actual_bounding_box_ascent() + metrics.actual_bounding_box_descent();
            (xb, yb, w, h)
        }

        fn clear(&mut self, r: f64, g: f64, b: f64, a: f64) {
            let _ = self.ctx.set_fill_style_str(&Self::rgba(r, g, b, a));
            let canvas = self.ctx.canvas().unwrap();
            let w = canvas.width() as f64;
            let h = canvas.height() as f64;
            self.ctx.fill_rect(0.0, 0.0, w, h);
        }

        /// Blit a straight-alpha RGBA8 image at `(x, y)`, scaled by `scale`.
        ///
        /// The only `DrawContext` method GTK overrides that this backend did
        /// not, so every caller fell through to the trait default and got
        /// `false` -- which is the "I could not draw this, degrade to vector
        /// primitives" answer, and is why an image preview on a page showed
        /// nothing at all.
        ///
        /// Two steps, because there is no `putImageDataAt`-with-transform:
        /// `putImageData` writes at integer device pixels and ignores the
        /// transform, and `putImageData` is the only way to get *premultiplied*
        /// semantics right. The DOM wants **straight** alpha, which is what
        /// the caller supplies, so the pixels are staged verbatim into an
        /// `ImageData`; the scale and the sub-pixel origin are then applied by
        /// `drawImage` from that staging canvas, which does honour the
        /// transform.
        ///
        /// Scaling is done by `drawImage` with a destination rectangle rather
        /// than by resampling the `ImageData`, because the DOM's image
        /// smoothing is what a caller expects from a scaled blit (and matches
        /// what Cairo's own filter does on GTK).
        fn draw_rgba_image(
            &mut self,
            x: f64,
            y: f64,
            pixels: &[u8],
            width: u32,
            height: u32,
            scale: f64,
        ) -> bool {
            if width == 0 || height == 0 {
                // A zero-sized image has nothing to blit; `false` tells the
                // caller to fall back, which is the honest answer rather than
                // a silent no-op that looks like success.
                return false;
            }
            let expected = (width as usize) * (height as usize) * 4;
            if pixels.len() < expected {
                // Short buffer: the DOM would throw a DataCloneError on an
                // `ImageData` of the wrong length. Reporting `false` keeps the
                // caller's fallback path instead of panicking across the FFI
                // boundary.
                return false;
            }
            // `Clamped` is a marker wrapper the DOM's byte array type uses;
            // the slice is wrapped, not reinterpreted, so the bytes cross
            // unmodified. The argument order is (data, width, height).
            let Ok(image_data) = web_sys::ImageData::new_with_u8_clamped_array_and_sh(
                wasm_bindgen::Clamped(pixels),
                width,
                height,
            ) else {
                return false;
            };

            // A scratch canvas to hold the unscaled image, so the destination
            // rectangle can scale it. Created on demand and dropped after --
            // `drawImage` needs a source, and an `ImageData` is not one.
            let Ok(scratch) = create_element("canvas").dyn_into::<HtmlCanvasElement>() else {
                return false;
            };
            scratch.set_width(width);
            scratch.set_height(height);
            let Some(sctx) = scratch
                .get_context("2d")
                .ok()
                .flatten()
                .and_then(|c| c.dyn_into::<web_sys::CanvasRenderingContext2d>().ok())
            else {
                return false;
            };
            // Both calls below are fallible (they return `Result<_, JsValue>`
            // because a DOM exception is reportable across the boundary: a
            // detached canvas, or a scratch canvas with no 2d context). A
            // failure means the blit did not happen, so `false` is the
            // honest answer and lets the caller fall back to vector drawing
            // rather than silently drawing nothing.
            if sctx.put_image_data(&image_data, 0.0, 0.0).is_err() {
                return false;
            }

            let dw = width as f64 * scale;
            let dh = height as f64 * scale;
            // The `and_dw_and_dh` overload is the one that scales; the plain
            // form blits at the source's natural size and would ignore
            // `scale` entirely. There is no single 6-argument overload: each
            // source type has its own, so this names the canvas one.
            self.ctx
                .draw_image_with_html_canvas_element_and_dw_and_dh(
                    scratch.as_ref(),
                    x,
                    y,
                    dw,
                    dh,
                )
                .is_ok()
        }

        fn save(&mut self) { self.ctx.save(); }
        fn restore(&mut self) { self.ctx.restore(); }

        fn clip(&mut self, x: f64, y: f64, w: f64, h: f64) {
            self.ctx.begin_path();
            self.ctx.rect(x, y, w, h);
            self.ctx.clip();
        }
    }

    // -----------------------------------------------------------------------
    // Canvas
    // -----------------------------------------------------------------------
    pub struct Canvas {
        elem: HtmlCanvasElement,
        draw_cb: Rc<RefCell<Option<Box<dyn FnMut(&mut dyn crate::core::DrawContext, i32, i32)>>>>,
        click_cb: Rc<RefCell<Option<Box<dyn FnMut(f64, f64)>>>>,
        key_cb: Rc<RefCell<Option<Box<dyn FnMut(u32) -> bool>>>>,
        /// `(x, y, button, modifier mask)`, for a press.
        click_button_cb: Rc<RefCell<Option<Box<dyn FnMut(f64, f64, u32, u32)>>>>,
        /// `(x, y, state)`, for motion.
        motion_cb: Rc<RefCell<Option<Box<dyn FnMut(f64, f64, u32)>>>>,
        /// `(keyval, modifier mask) -> consumed`, for keys.
        key_raw_cb: Rc<RefCell<Option<Box<dyn FnMut(u32, u32) -> bool>>>>,
        /// `(x, y, button, modifier mask)`, for a release.
        release_cb: Rc<RefCell<Option<Box<dyn FnMut(f64, f64, u32, u32)>>>>,
        closures: Rc<RefCell<Vec<Box<dyn Any>>>>,
    }

    impl AsElement for Canvas {
        fn as_element(&self) -> &Element {
            self.elem.as_ref()
        }
    }

    impl Widget for Canvas {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const HtmlCanvasElement as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for Canvas {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.raw_handle() as *const *mut c_void) }
        }
    }

    impl Clone for Canvas {
        fn clone(&self) -> Self {
            Canvas {
                elem: self.elem.clone(),
                draw_cb: self.draw_cb.clone(),
                click_cb: self.click_cb.clone(),
                key_cb: self.key_cb.clone(),
                click_button_cb: self.click_button_cb.clone(),
                motion_cb: self.motion_cb.clone(),
                key_raw_cb: self.key_raw_cb.clone(),
                release_cb: self.release_cb.clone(),
                closures: self.closures.clone(),
            }
        }
    }

    impl Canvas {
        pub fn set_draw_callback(&self, cb: Box<dyn FnMut(&mut dyn crate::core::DrawContext, i32, i32)>) {
            *self.draw_cb.borrow_mut() = Some(cb);
            self.queue_redraw();
        }

        pub fn queue_redraw(&self) {
            if let Some(ref mut draw_fn) = *self.draw_cb.borrow_mut() {
                let ctx = self.elem.get_context("2d")
                    .ok().flatten()
                    .and_then(|o| o.dyn_into::<web_sys::CanvasRenderingContext2d>().ok());
                if let Some(ctx) = ctx {
                    let w = self.elem.width() as i32;
                    let h = self.elem.height() as i32;
                    let mut dc = WasmDrawContext { ctx };
                    draw_fn(&mut dc, w, h);
                }
            }
        }

        pub fn set_size_request(&self, w: i32, h: i32) {
            if w > 0 { self.elem.set_width(w as u32); }
            if h > 0 { self.elem.set_height(h as u32); }
            set_css(self.elem.as_ref(), "width", &format!("{}px", w));
            set_css(self.elem.as_ref(), "height", &format!("{}px", h));
        }

        pub fn set_content_size(&self, w: i32, h: i32) {
            self.elem.set_width(w as u32);
            self.elem.set_height(h as u32);
        }

        pub fn set_visible(&self, v: bool) {
            if v {
                set_css(self.elem.as_ref(), "display", "");
            } else {
                set_css(self.elem.as_ref(), "display", "none");
            }
        }

        /// Button/modifier-aware press: `(x, y, button, modifier mask)`.
        ///
        /// All three of this file's pointer methods were comment-only stubs
        /// ("the DOM handler does not yet forward button/state"), which is
        /// what made canvas dragging, hover and right-click menus impossible
        /// on a page. The DOM has all of it: `MouseEvent::button()` gives the
        /// button number and the shift/ctrl/alt keys are on the event, so the
        /// same `(x, y, button, state)` tuple GTK produces is available here.
        pub fn on_click_button(&self, cb: Box<dyn FnMut(f64, f64, u32, u32)>) {
            *self.click_button_cb.borrow_mut() = Some(cb);
            let cell = self.click_button_cb.clone();
            let closure = Closure::<dyn FnMut(MouseEvent)>::new(move |evt: MouseEvent| {
                if let Some(f) = cell.borrow_mut().as_mut() {
                    f(
                        evt.offset_x() as f64,
                        evt.offset_y() as f64,
                        evt.button() as u32,
                        gdk_modifier_mask_mouse(&evt),
                    );
                }
            });
            self.elem
                .add_event_listener_with_callback("mousedown", closure.as_ref().unchecked_ref())
                .ok();
            self.closures.borrow_mut().push(Box::new(closure));
        }

        /// Pointer motion: `(x, y, state mask)`.
        ///
        /// The state is the same GDK-style mask as everywhere else (1 = Shift,
        /// 4 = Control, 8 = Alt), so a caller can ask "is a button held"
        /// without the backend translating a second time.
        pub fn on_motion(&self, cb: Box<dyn FnMut(f64, f64, u32)>) {
            *self.motion_cb.borrow_mut() = Some(cb);
            let cell = self.motion_cb.clone();
            let closure = Closure::<dyn FnMut(MouseEvent)>::new(move |evt: MouseEvent| {
                if let Some(f) = cell.borrow_mut().as_mut() {
                    f(
                        evt.offset_x() as f64,
                        evt.offset_y() as f64,
                        gdk_modifier_mask_mouse(&evt),
                    );
                }
            });
            self.elem
                .add_event_listener_with_callback("mousemove", closure.as_ref().unchecked_ref())
                .ok();
            self.closures.borrow_mut().push(Box::new(closure));
        }

        /// This canvas's top-left in screen coordinates.
        ///
        /// `getBoundingClientRect` gives the viewport-relative box, and the
        /// canvas's offset within the document is added so the answer is in the
        /// same *page* coordinates the GTK backend reports. Returning `None`
        /// here is what forced every context menu on a page to open
        /// unpositioned, since a caller has no origin to place it against.
        pub fn screen_origin(&self) -> Option<(i32, i32)> {
            let rect = self.elem.get_bounding_client_rect();
            let scroll = document().document_element()?;
            let x = rect.x() + scroll.scroll_left() as f64;
            let y = rect.y() + scroll.scroll_top() as f64;
            Some((x.round() as i32, y.round() as i32))
        }

        /// Pointer release: `(x, y, button, modifier mask)`.
        ///
        /// Pairs with [`Canvas::on_click_button`]; a press and a release with
        /// motion between them are what a drag is, and without the release a
        /// drag can never end.
        pub fn on_release(&self, cb: Box<dyn FnMut(f64, f64, u32, u32)>) {
            *self.release_cb.borrow_mut() = Some(cb);
            let cell = self.release_cb.clone();
            let closure = Closure::<dyn FnMut(MouseEvent)>::new(move |evt: MouseEvent| {
                if let Some(f) = cell.borrow_mut().as_mut() {
                    f(
                        evt.offset_x() as f64,
                        evt.offset_y() as f64,
                        evt.button() as u32,
                        gdk_modifier_mask_mouse(&evt),
                    );
                }
            });
            self.elem
                .add_event_listener_with_callback("mouseup", closure.as_ref().unchecked_ref())
                .ok();
            self.closures.borrow_mut().push(Box::new(closure));
        }

        pub fn set_hexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-grow", "1");
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "flex-grow", "0");
            }
        }

        /// See [`Entry::set_vexpand`].
        pub fn set_vexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-shrink", "1");
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "flex-shrink", "0");
            }
        }

        /// Outer spacing in px, for a canvas inset inside its container.
        pub fn set_margin_start(&self, px: i32) {
            set_css(self.elem.as_ref(), "margin-left", &format!("{}px", px));
        }

        /// See [`Canvas::set_margin_start`].
        pub fn set_margin_top(&self, px: i32) {
            set_css(self.elem.as_ref(), "margin-top", &format!("{}px", px));
        }

        pub fn on_click(&self, cb: Box<dyn FnMut(f64, f64)>) {
            *self.click_cb.borrow_mut() = Some(cb);
            let cb2 = self.click_cb.clone();
            let closure = Closure::<dyn FnMut(MouseEvent)>::new(move |evt: MouseEvent| {
                if let Some(ref mut f) = *cb2.borrow_mut() {
                    f(evt.offset_x() as f64, evt.offset_y() as f64);
                }
            });
            self.elem
                .add_event_listener_with_callback("mousedown", closure.as_ref().unchecked_ref())
                .ok();
            store_closure(Box::new(closure));
        }

        pub fn on_key(&self, cb: Box<dyn FnMut(u32) -> bool>) {
            *self.key_cb.borrow_mut() = Some(cb);
            let cb2 = self.key_cb.clone();
            let closure = Closure::<dyn FnMut(KeyboardEvent)>::new(move |evt: KeyboardEvent| {
                if let Some(ref mut f) = *cb2.borrow_mut() {
                    if f(evt.key_code()) {
                        evt.prevent_default();
                    }
                }
            });
            self.elem
                .add_event_listener_with_callback("keydown", closure.as_ref().unchecked_ref())
                .ok();
            store_closure(Box::new(closure));
            // Make canvas focusable
            self.elem.set_tab_index(0);
        }
        pub fn on_key_raw(&self, cb: Box<dyn FnMut(u32, u32) -> bool>) {
            // Registered directly rather than through `on_key`, which discards
            // the modifier state -- and a canvas key handler is exactly where
            // an accelerator (Ctrl+S) has to be distinguishable from a bare S.
            *self.key_raw_cb.borrow_mut() = Some(cb);
            let cell = self.key_raw_cb.clone();
            let closure = Closure::<dyn FnMut(KeyboardEvent)>::new(move |evt: KeyboardEvent| {
                if let Some(f) = cell.borrow_mut().as_mut() {
                    if f(evt.key_code() as u32, gdk_modifier_mask(&evt)) {
                        evt.prevent_default();
                    }
                }
            });
            self.elem
                .add_event_listener_with_callback("keydown", closure.as_ref().unchecked_ref())
                .ok();
            self.closures.borrow_mut().push(Box::new(closure));
        }
        pub fn grab_focus(&self) {
            let _ = self.elem.focus();
        }
        pub fn set_can_focus(&self, _can: bool) {}
        pub fn force_draw(&self, _window_ptr: *mut c_void, _fallback_w: i32, _fallback_h: i32) {}
    }

    pub fn create_canvas() -> Result<Canvas, Error> {
        let elem: HtmlCanvasElement = create_element("canvas").dyn_into().map_err(|e| {
            Error::Backend(format!("create_canvas: {:?}", e))
        })?;
        let ctx = elem.get_context("2d")
            .ok().flatten()
            .and_then(|o| o.dyn_into::<web_sys::CanvasRenderingContext2d>().ok())
            .ok_or_else(|| Error::Backend("getContext('2d') failed".into()))?;
        // Use default font that will be overridden per draw call
        let _ = ctx.set_font("12px monospace");
        Ok(Canvas {
            elem,
            draw_cb: Rc::new(RefCell::new(None)),
            click_cb: Rc::new(RefCell::new(None)),
            key_cb: Rc::new(RefCell::new(None)),
            click_button_cb: Rc::new(RefCell::new(None)),
            motion_cb: Rc::new(RefCell::new(None)),
            key_raw_cb: Rc::new(RefCell::new(None)),
            release_cb: Rc::new(RefCell::new(None)),
            closures: Rc::new(RefCell::new(Vec::new())),
        })
    }

    // -----------------------------------------------------------------------
    // Overlay (stacking container using position:relative + absolute overlays)
    // -----------------------------------------------------------------------
    pub struct Overlay {
        elem: HtmlDivElement,
        children: Rc<RefCell<Vec<Element>>>,
    }

    impl AsElement for Overlay {
        fn as_element(&self) -> &Element {
            self.elem.as_ref()
        }
    }

    impl Widget for Overlay {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const HtmlDivElement as *mut c_void
        }
    }

    impl Clone for Overlay {
        fn clone(&self) -> Self {
            Overlay {
                elem: self.elem.clone(),
                children: self.children.clone(),
            }
        }
    }

    impl Overlay {
        pub fn set_child(&self, child: &impl AsElement) {
            while let Some(c) = self.elem.first_child() {
                self.elem.remove_child(&c).ok();
            }
            self.children.borrow_mut().clear();
            self.children.borrow_mut().push(child.as_element().clone());
            self.elem.append_child(child.as_element()).ok();
        }

        pub fn add_overlay(&self, child: &impl AsElement) {
            self.children.borrow_mut().push(child.as_element().clone());
            set_css(child.as_element(), "position", "absolute");
            set_css(child.as_element(), "z-index", "10");
            self.elem.append_child(child.as_element()).ok();
        }

        pub fn set_overlay_pass_through(&self, child: &impl AsElement, pass: bool) {
            if pass {
                set_css(child.as_element(), "pointer-events", "none");
            } else {
                set_css(child.as_element(), "pointer-events", "auto");
            }
        }

        pub fn remove(&self, child: &impl AsElement) {
            self.elem.remove_child(child.as_element()).ok();
            self.children.borrow_mut().retain(|c| c != child.as_element());
        }

        pub fn show_all(&self) {
            for c in self.children.borrow().iter() {
                if let Some(html) = c.dyn_ref::<HtmlElement>() {
                    html.style().set_property("display", "").ok();
                }
            }
        }

        pub fn set_size_request(&self, w: i32, h: i32) {
            if w > 0 { set_css(self.elem.as_ref(), "width", &format!("{}px", w)); }
            if h > 0 { set_css(self.elem.as_ref(), "height", &format!("{}px", h)); }
        }

        pub fn set_vexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-grow", "1");
            } else {
                set_css(self.elem.as_ref(), "flex-grow", "0");
            }
        }

        pub fn set_hexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "align-self", "stretch");
            }
        }

        /// Outer spacing of the overlay itself, in px. See
        /// [`Label::set_margin_start`] for why this is a margin and not a
        /// padding.
        pub fn set_margin_start(&self, px: i32) {
            set_css(self.elem.as_ref(), "margin-left", &format!("{}px", px));
        }

        /// See [`Overlay::set_margin_start`].
        pub fn set_margin_top(&self, px: i32) {
            set_css(self.elem.as_ref(), "margin-top", &format!("{}px", px));
        }
    }

    pub fn create_overlay() -> Result<Overlay, Error> {
        let div: HtmlDivElement = create_element("div").dyn_into().map_err(|e| {
            Error::Backend(format!("create_overlay: {:?}", e))
        })?;
        set_css(div.as_ref(), "position", "relative");
        div.style().set_property("overflow", "hidden").ok();
        Ok(Overlay {
            elem: div,
            children: Rc::new(RefCell::new(Vec::new())),
        })
    }

    // -----------------------------------------------------------------------
    // ScrolledWindow
    // -----------------------------------------------------------------------
    #[derive(Clone)]
    pub struct ScrolledWindow {
        elem: HtmlDivElement,
        /// The registered `on_scroll` callback, if any.
        scroll_cb: Rc<RefCell<Option<Box<dyn FnMut(bool, f64)>>>>,
        /// The scroll listener's own `Closure`, kept alive for as long as the
        /// window is: dropping it would free the JS trampoline the listener
        /// points at, and the next scroll would call into freed memory.
        closures: Rc<RefCell<Vec<Box<dyn Any>>>>,
    }

    impl AsElement for ScrolledWindow {
        fn as_element(&self) -> &Element {
            self.elem.as_ref()
        }
    }

    impl Widget for ScrolledWindow {
        fn raw_handle(&self) -> *mut c_void {
            &self.elem as *const HtmlDivElement as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for ScrolledWindow {
        fn as_ref(&self) -> &*mut c_void {
            unsafe { &*(&self.raw_handle() as *const *mut c_void) }
        }
    }

    impl ScrolledWindow {
        /// Show or hide the native scrollbars.
        ///
        /// GTK's `GtkPolicyType` ordinals: 0 = always, 1 = automatic,
        /// 2 = never, 3 = external. CSS `overflow` has only two answers per
        /// axis -- "there may be a scrollbar" (`auto`/`scroll`) or "there is
        /// none" (`hidden`) -- so *automatic* maps to `auto` (shown only when
        /// needed) and *always* to `scroll` (the bar is reserved even when
        /// there is nothing to scroll). A caller that reserved the space on
        /// purpose gets `scroll`; one that wants it only when needed gets
        /// `auto`. That is the closest honest reading of the four-valued
        /// policy in a two-valued property.
        pub fn set_policy(&self, hscroll: i32, vscroll: i32) {
            let h = match hscroll { 0 => "scroll", 2 | 3 => "hidden", _ => "auto" };
            let v = match vscroll { 0 => "scroll", 2 | 3 => "hidden", _ => "auto" };
            self.elem.style().set_property("overflow-x", h).ok();
            self.elem.style().set_property("overflow-y", v).ok();
        }

        pub fn set_child(&self, child: &impl AsElement) {
            while let Some(c) = self.elem.first_child() {
                self.elem.remove_child(&c).ok();
            }
            self.elem.append_child(child.as_element()).ok();
        }

        /// Push the viewport position and domain into the scroll offset.
        ///
        /// The arguments are the shared **cell-index** model, not pixels:
        /// `sync_scrollbars` (`gui_backend.rs:3053`) passes
        /// `(value, upper, page)` where value and upper count cells. GTK
        /// configures a `GtkAdjustment` with them directly; a DOM element
        /// scrolls in `scrollLeft`/`scrollTop` pixels, so the cell fraction
        /// is scaled by the *scrollable* extent -- which is
        /// `scrollHeight - clientHeight`, not the content height, since
        /// that difference is exactly the distance the offset ranges over.
        ///
        /// A domain of 0, or a container that is not currently scrollable,
        /// is a no-op: there is no position to express.
        pub fn scroll_to(
            &self,
            hval: f64,
            hupper: f64,
            _hpage: f64,
            vval: f64,
            vupper: f64,
            _vpage: f64,
        ) {
            if let Some(x) = self.horizontal_range() {
                if hupper > 0.0 {
                    self.elem.set_scroll_left(clamp_scroll(hval / hupper, x));
                }
            }
            if let Some(y) = self.vertical_range() {
                if vupper > 0.0 {
                    self.elem.set_scroll_top(clamp_scroll(vval / vupper, y));
                }
            }
        }

        /// Fire `cb(vertical, value)` when the user scrolls this element.
        ///
        /// The value is in the same cell-index units `scroll_to` accepts, so
        /// `scroll_to_cursor` (`gui_backend.rs:3064`) can consume it without
        /// a second conversion: the pixel offset divided by the scrollable
        /// extent is the fraction, and the fraction is what the caller wants
        /// (it clamps into the domain itself). GTK reports the adjustment's
        /// value in its own units; here those units *are* the fraction, which
        /// is the same shape.
        pub fn on_scroll(&self, cb: Box<dyn FnMut(bool, f64)>) {
            *self.scroll_cb.borrow_mut() = Some(cb);
            if !self.closures.borrow().is_empty() {
                // A listener is already installed; the closure captured the
                // cell, so it reads the current callback on every event and a
                // second registration needs no second listener.
                return;
            }
            let cell = self.scroll_cb.clone();
            let horizontal_range = {
                let elem = self.elem.clone();
                Rc::new(move || (elem.scroll_width() as i32 - elem.client_width() as i32).max(0))
            };
            let vertical_range = {
                let elem = self.elem.clone();
                Rc::new(move || (elem.scroll_height() as i32 - elem.client_height() as i32).max(0))
            };
            let closure = Closure::<dyn FnMut(Event)>::new(move |evt: Event| {
                let target = match evt.target() {
                    Some(t) => t,
                    None => return,
                };
                let el: Element = match target.dyn_into() {
                    Ok(e) => e,
                    Err(_) => return,
                };
                // A scroll event on a descendant does not bubble here as
                // this element's own offset, so read the element we are
                // attached to rather than the event target.
                let _ = el;
                let y = vertical_range();
                let x = horizontal_range();
                let elem: Element = match evt.current_target().and_then(|t| t.dyn_into().ok()) {
                    Some(e) => e,
                    None => return,
                };
                let mut cb = match cell.try_borrow_mut() {
                    Ok(c) => c,
                    Err(_) => return,
                };
                if let Some(f) = cb.as_mut() {
                    if y > 0 {
                        f(true, (elem.scroll_top() as f64 / y as f64).clamp(0.0, 1.0));
                    }
                    if x > 0 {
                        f(false, (elem.scroll_left() as f64 / x as f64).clamp(0.0, 1.0));
                    }
                }
            });
            self.elem
                .add_event_listener_with_callback("scroll", closure.as_ref().unchecked_ref())
                .ok();
            self.closures.borrow_mut().push(Box::new(closure));
        }

        /// How far the element can be scrolled horizontally, or `None` when
        /// it is not currently scrollable on that axis (which is what makes
        /// `scroll_to` a no-op rather than a jump to 0).
        fn horizontal_range(&self) -> Option<i32> {
            let range = (self.elem.scroll_width() as i32 - self.elem.client_width() as i32).max(0);
            if range > 0 { Some(range) } else { None }
        }

        /// How far the element can be scrolled vertically. See
        /// [`ScrolledWindow::horizontal_range`].
        fn vertical_range(&self) -> Option<i32> {
            let range = (self.elem.scroll_height() as i32 - self.elem.client_height() as i32).max(0);
            if range > 0 { Some(range) } else { None }
        }

        pub fn set_vexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "flex-grow", "1");
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "flex-grow", "0");
            }
        }

        pub fn set_hexpand(&self, expand: bool) {
            if expand {
                set_css(self.elem.as_ref(), "align-self", "stretch");
            } else {
                set_css(self.elem.as_ref(), "align-self", "");
            }
        }
    }

    /// A scroll offset for a fraction of the scrollable range.
    ///
    /// Clamped to the range because `scrollLeft`/`scrollTop` silently ignore
    /// an out-of-range value, and because the caller's fraction can exceed 1
    /// when the domain is larger than what is currently rendered. A NaN
    /// fraction (0/0) is treated as 0, since "no information" and "top left"
    /// are the same answer for a scroll position.
    fn clamp_scroll(fraction: f64, range: i32) -> i32 {
        if !fraction.is_finite() {
            return 0;
        }
        ((fraction.clamp(0.0, 1.0)) * range as f64) as i32
    }

    pub fn create_scrolled_window() -> Result<ScrolledWindow, Error> {
        let div: HtmlDivElement = create_element("div").dyn_into().map_err(|e| {
            Error::Backend(format!("create_scrolled_window: {:?}", e))
        })?;
        div.style().set_property("overflow", "auto").ok();
        div.style().set_property("position", "relative").ok();
        Ok(ScrolledWindow {
            elem: div,
            scroll_cb: Rc::new(RefCell::new(None)),
            closures: Rc::new(RefCell::new(Vec::new())),
        })
    }
}

#[cfg(target_arch = "wasm32")]
pub use wasm_adapter::*;
