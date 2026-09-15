#[cfg(windows)]
mod nwg_adapter {
    use native_windows_gui as nwg;
    use crate::core::{Error, Widget, DrawContext};
    use std::cell::RefCell;
    use std::collections::HashMap;
    use std::os::raw::c_void;
    use std::rc::Rc;
    use std::sync::atomic::{AtomicUsize, Ordering};

    fn set_window_pos(hwnd: *mut c_void, x: i32, y: i32, w: i32, h: i32) {
        unsafe {
            winapi::um::winuser::SetWindowPos(
                hwnd as winapi::shared::windef::HWND,
                std::ptr::null_mut(), x, y, w, h,
                winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_SHOWWINDOW,
            );
        }
    }

    // TEMPORARY Win95 diagnosis: raw file marker (std::fs broken on 9x).
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    fn mark95a(s: &[u8]) {
        unsafe {
            extern "system" {
                fn CreateFileA(name: *const u8, access: u32, share: u32,
                    sa: *mut c_void, disp: u32, flags: u32,
                    tmpl: *mut c_void) -> *mut c_void;
                fn SetFilePointer(h: *mut c_void, lo: i32, hi: *mut i32, how: u32) -> u32;
                fn WriteFile(h: *mut c_void, buf: *const u8, len: u32, w: *mut u32,
                    ov: *mut c_void) -> i32;
                fn CloseHandle(h: *mut c_void) -> i32;
            }
            let h = CreateFileA(b"c:\\gcorro.log\0".as_ptr(), 0x4000_0000, 1,
                std::ptr::null_mut(), 4, 0x80, std::ptr::null_mut());
            if !h.is_null() && h as isize != -1 {
                SetFilePointer(h, 0, std::ptr::null_mut(), 2);
                let mut w = 0u32;
                WriteFile(h, s.as_ptr(), s.len() as u32, &mut w, std::ptr::null_mut());
                CloseHandle(h);
            }
        }
    }

    // TEMPORARY Win95 diagnosis: log a label + two i32s as hex.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    fn mark95xy(tag: &[u8; 5], a: i32, b: i32) {
        let hx = b"0123456789abcdef";
        let mut s = [0u8; 32];
        let mut p = 0;
        for i in 0..5 { s[p] = tag[i]; p += 1; }
        s[p] = b' '; p += 1;
        for v in [a as u32, b as u32] {
            for sh in [28u32, 24, 20, 16, 12, 8, 4, 0] {
                s[p] = hx[((v >> sh) & 0xf) as usize]; p += 1;
            }
            s[p] = b' '; p += 1;
        }
        s[p] = b'\n'; p += 1;
        mark95a(&s[..p]);
    }

    /// Translate a WM_KEYDOWN wParam (raw Win32 virtual-key code) for the
    /// app key callback, honouring Shift/CapsLock and the active keyboard
    /// layout (Shift+9 is '(' on US layouts, not '9'; Shift+letter gives
    /// capitals). Returns:
    /// - `Some(raw VK)` when Ctrl/Alt is held (accelerator/menu paths must
    ///   keep raw codes) or when no printable character results (special
    ///   keys keep their VK so app dispatch by constant keeps working);
    /// - `Some(translated char)` when the char cannot collide with an app
    ///   dispatched VK;
    /// - `None` when the translated char numerically collides with an app
    ///   dispatched VK (`!"#$%&'()*` are VK_PRIOR..VK_DOWN, `.` is
    ///   VK_DELETE). Callers must decline (return None) so the native
    ///   control inserts the character and the change event resyncs app
    ///   state. Passing the char on would misfire menu/cursor actions —
    ///   e.g. '(' (0x28) would commit the edit and move down as VK_DOWN.
    /// Known limitation: dead-key compositions fall back to the raw VK.
    fn translate_vk(wparam: usize) -> Option<u32> {
        let vk = (wparam & 0xFF) as u32;
        unsafe {
            const VK_CONTROL: i32 = 0x11;
            const VK_MENU: i32 = 0x12;
            if winapi::um::winuser::GetKeyState(VK_CONTROL) as u16 & 0x8000 != 0 {
                return Some(vk);
            }
            if winapi::um::winuser::GetKeyState(VK_MENU) as u16 & 0x8000 != 0 {
                return Some(vk);
            }
            let mut state: [u8; 256] = [0; 256];
            if winapi::um::winuser::GetKeyboardState(state.as_mut_ptr()) == 0 {
                return Some(vk);
            }
            let scan =
                winapi::um::winuser::MapVirtualKeyW(vk, winapi::um::winuser::MAPVK_VK_TO_VSC);
            // Win95: ToUnicodeEx is a Win2000+ USER32 export (and absent from
            // Win95's export table entirely — a hard loader failure, not just
            // a W-stub). ToAsciiEx is the Win95-era ANSI equivalent: same
            // VK/scan/keyboard-state inputs, same "1 char, 0 none, -1 dead
            // key" contract, but it writes ANSI bytes into a WORD array.
            // The app's key path only ever deals with ASCII key names and
            // control codes, so the ANSI result is what it wants.
            let mut buf: [u16; 4] = [0; 4];
            let n = winapi::um::winuser::ToAsciiEx(
                vk,
                scan,
                state.as_ptr(),
                buf.as_mut_ptr(),
                0,
                winapi::um::winuser::GetKeyboardLayout(0),
            );
            // Exactly one printable char, no pending dead-key composition —
            // unless it collides with an app-dispatched VK (see doc).
            if n == 1 && buf[0] >= 32 {
                let c = buf[0] as u32;
                if matches!(c, 0x21..=0x28 | 0x2E) {
                    return None;
                }
                return Some(c);
            }
            Some(vk)
        }
    }

    /// GDK-compatible modifier mask for the current physical key state:
    /// bit 0 = Shift, bit 2 = Ctrl, bit 3 = Alt. Lets widget callbacks see
    /// the same modifiers GTK delivers, so Ctrl/Alt guards in app code
    /// (copy/paste decline, menu propagation, quit) work on Windows too.
    /// Previously only Shift was reported, so Ctrl+C typed 'C' instead of
    /// copying and Ctrl+Q could never quit.
    fn modifier_state() -> u32 {
        unsafe {
            const VK_SHIFT: i32 = 0x10;
            const VK_CONTROL: i32 = 0x11;
            const VK_MENU: i32 = 0x12;
            let mut mods: u32 = 0;
            if winapi::um::winuser::GetKeyState(VK_SHIFT) as u16 & 0x8000 != 0 {
                mods |= 1;
            }
            if winapi::um::winuser::GetKeyState(VK_CONTROL) as u16 & 0x8000 != 0 {
                mods |= 4;
            }
            if winapi::um::winuser::GetKeyState(VK_MENU) as u16 & 0x8000 != 0 {
                mods |= 8;
            }
            mods
        }
    }

    /// True when pure Ctrl (no Alt) is physically held: accelerators and
    /// native edit shortcuts own the key, so a canvas must decline rather
    /// than type it. Ctrl+Alt (AltGr) returns false so international text
    /// input still reaches the widget.
    fn pure_ctrl_held() -> bool {
        modifier_state() & 0xC == 0x4
    }

    /// True when pure Alt (no Ctrl) is physically held: the key belongs to
    /// OS menu/accelerator handling, not widget editing. Ctrl+Alt (AltGr)
    /// returns false so international text input still reaches the widget.
    fn pure_alt_held() -> bool {
        unsafe {
            const VK_MENU: i32 = 0x12;
            const VK_CONTROL: i32 = 0x11;
            let alt = winapi::um::winuser::GetKeyState(VK_MENU) as u16 & 0x8000 != 0;
            let ctrl = winapi::um::winuser::GetKeyState(VK_CONTROL) as u16 & 0x8000 != 0;
            crate::win32_portable::should_yield_to_menu(alt, ctrl)
        }
    }

    /// Current outer size of a control, if measurable.
    fn control_size(hwnd: winapi::shared::windef::HWND) -> Option<(i32, i32)> {
        unsafe {
            let mut rect: winapi::shared::windef::RECT = std::mem::zeroed();
            if winapi::um::winuser::GetWindowRect(hwnd, &mut rect) == 0 {
                return None;
            }
            Some((rect.right - rect.left, rect.bottom - rect.top))
        }
    }

    /// Dialog content+button layout shared by `Dialog::layout_dialog` and
    /// the dialog WM_SIZE binding (which fires before any Dialog exists).
    fn layout_nwg_dialog_parts(
        dlg_hwnd: winapi::shared::windef::HWND,
        content: &RefCell<Vec<*mut c_void>>,
        buttons: &RefCell<Vec<(nwg::Button, nwg::EventHandler)>>,
    ) {
        let (cw, ch) = unsafe {
            let mut rect: winapi::shared::windef::RECT = std::mem::zeroed();
            if winapi::um::winuser::GetClientRect(dlg_hwnd, &mut rect) == 0 {
                return;
            }
            (rect.right - rect.left, rect.bottom - rect.top)
        };
        let mut specs: Vec<(i32, i32, bool)> = Vec::new();
        for &ptr in content.borrow().iter() {
            let (w, h) = control_size(ptr as _).unwrap_or((0, 0));
            if w <= 8 && h <= 8 {
                specs.push((0, 0, true));
            } else {
                specs.push((w, if h > 10 { h } else { 26 }, false));
            }
        }
        let mut btn_sizes: Vec<(i32, i32)> = Vec::new();
        for (btn, _) in buttons.borrow().iter() {
            let (w, h) = btn
                .handle
                .hwnd()
                .and_then(control_size)
                .unwrap_or((0, 0));
            btn_sizes.push((
                if (40..=220).contains(&w) { w } else { 96 },
                if (16..=48).contains(&h) { h } else { 28 },
            ));
        }
        let (content_rects, btn_rects) =
            crate::win32_portable::dialog_layout_geometry(cw, ch, &specs, &btn_sizes);
        let content = content.borrow();
        for (i, r) in content_rects.iter().enumerate() {
            if i >= content.len() {
                break;
            }
            set_window_pos(content[i], r.x, r.y, r.w, r.h);
            unsafe {
                let l = ((r.h & 0xFFFF) << 16) | (r.w & 0xFFFF);
                winapi::um::winuser::SendMessageW(
                    content[i] as _,
                    winapi::um::winuser::WM_SIZE,
                    0,
                    l as _,
                );
            }
        }
        let buttons = buttons.borrow();
        for (i, r) in btn_rects.iter().enumerate() {
            if i >= buttons.len() {
                break;
            }
            if let Some(h) = buttons[i].0.handle.hwnd() {
                set_window_pos(h as *mut c_void, r.x, r.y, r.w, r.h);
            }
        }
    }

    // -- Window --

    pub struct Window {
        pub(crate) inner: Rc<nwg::Window>,
        pub(crate) _handler: Rc<nwg::EventHandler>,
        // Stable hwnd snapshot for AsRef. The previous body referenced a
        // temporary (`&handle.hwnd().unwrap_or(..) as ...`) — a dangling
        // reference (UB), garbage hwnds in release builds.
        hwnd: *mut c_void,
        root_child: Rc<RefCell<Option<*mut c_void>>>,
        layout_cb: Rc<RefCell<Option<Box<dyn FnMut(i32, i32)>>>>,
        event_key_cb: Rc<RefCell<Option<Box<dyn FnMut(u32, u32) -> i32>>>>,
        close_cb: Rc<RefCell<Option<Box<dyn FnMut()>>>>,
    }

    impl Clone for Window {
        fn clone(&self) -> Self {
            Window {
                inner: self.inner.clone(),
                _handler: self._handler.clone(),
                hwnd: self.hwnd,
                root_child: self.root_child.clone(),
                layout_cb: self.layout_cb.clone(),
                event_key_cb: self.event_key_cb.clone(),
                close_cb: self.close_cb.clone(),
            }
        }
    }

    impl Widget for Window {
        fn raw_handle(&self) -> *mut c_void {
            self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for Window {
        fn as_ref(&self) -> &*mut c_void {
            &self.hwnd
        }
    }

    impl Window {
        pub fn set_title(&self, title: &str) { self.inner.set_text(title); }
        pub fn set_child(&self, child: &impl AsRef<*mut c_void>) {
            let ptr = *child.as_ref();
            if !ptr.is_null() {
                unsafe {
                    winapi::um::winuser::SetParent(ptr as _, self.hwnd() as _);
                }
                *self.root_child.borrow_mut() = Some(ptr);
                let hwnd = ptr;
                *self.layout_cb.borrow_mut() = Some(Box::new(move |w, h| {
                    set_window_pos(hwnd, 0, 0, w, h);
                }));
            }
        }
        pub fn set_child_box(&self, bx: &BoxWidget) {
            if !bx.hwnd.is_null() {
                unsafe {
                    winapi::um::winuser::SetParent(bx.hwnd as _, self.hwnd() as _);
                }
            }
            let hwnd = bx.hwnd;
            let bx = bx.clone();
            *self.layout_cb.borrow_mut() = Some(Box::new(move |w, h| {
                set_window_pos(hwnd, 0, 0, w, h);
                bx.layout(0, 0, w, h);
            }));
        }
        pub fn set_layout_cb<F: FnMut(i32, i32) + 'static>(&self, f: F) {
            *self.layout_cb.borrow_mut() = Some(Box::new(f));
        }
        pub fn set_default_size(&self, w: i32, h: i32) {
            let hwnd = self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut());
            if hwnd != std::ptr::null_mut() {
                unsafe {
                    winapi::um::winuser::SetWindowPos(
                        hwnd as _,
                        std::ptr::null_mut(), 0, 0, w, h,
                        winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_NOMOVE,
                    );
                }
            }
        }
        pub fn present(&self) {
            let hwnd = self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut());
            if hwnd != std::ptr::null_mut() {
                unsafe {
                    // Bring window to foreground so child controls can
                    // receive keyboard focus (required by SetFocus).
                    winapi::um::winuser::SetForegroundWindow(hwnd);
                    winapi::um::winuser::BringWindowToTop(hwnd);
                    let mut rect: winapi::shared::windef::RECT = std::mem::zeroed();
                    winapi::um::winuser::GetClientRect(hwnd, &mut rect);
                    let w = rect.right - rect.left;
                    let h = rect.bottom - rect.top;
                    // TEMPORARY Win95 diagnosis.
                    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
                    mark95xy(b"pres ", w, h);
                    if let Some(ref mut cb) = *self.layout_cb.borrow_mut() {
                        cb(w, h);
                    }
                }
            }
        }
        pub fn insert_action_group(&self, _name: &str, _group_ptr: *mut c_void) {}
        pub fn hwnd(&self) -> *mut c_void {
            self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void
        }
        pub fn queue_redraw(&self) {}
        pub fn on_event(&self, _cb: Box<dyn FnMut(*mut c_void) -> i32>) {}
        pub fn on_event_key(&self, cb: Box<dyn FnMut(u32, u32) -> i32>) {
            *self.event_key_cb.borrow_mut() = Some(cb);
        }
        pub fn on_close(&self, cb: Box<dyn FnMut()>) {
            *self.close_cb.borrow_mut() = Some(cb);
        }
    }

    pub fn create_window(parent_cell: &Rc<RefCell<Option<*mut c_void>>>) -> Result<Window, Error> {
        let (inner, handler) = crate::backends::nwg::create_window(parent_cell)
            .map_err(|e| Error::Backend(format!("{}", e)))?;
        let root_child: Rc<RefCell<Option<*mut c_void>>> = Rc::new(RefCell::new(None));
        let layout_cb: Rc<RefCell<Option<Box<dyn FnMut(i32, i32)>>>> = Rc::new(RefCell::new(None));
        let event_key_cb: Rc<RefCell<Option<Box<dyn FnMut(u32, u32) -> i32>>>> = Rc::new(RefCell::new(None));
        let close_cb: Rc<RefCell<Option<Box<dyn FnMut()>>>> = Rc::new(RefCell::new(None));

        let hwnd = inner.handle.hwnd().unwrap_or(std::ptr::null_mut());
        if hwnd != std::ptr::null_mut() {
            // Bind raw WM_SIZE handler
            let cb = layout_cb.clone();
            static RAW_HANDLER_ID: AtomicUsize = AtomicUsize::new(0x10000000);
            let handler_id = RAW_HANDLER_ID.fetch_add(1, Ordering::SeqCst);
            nwg::bind_raw_event_handler(
                &nwg::ControlHandle::Hwnd(hwnd),
                handler_id,
                move |_h, msg, _w, l| {
                    if msg == winapi::um::winuser::WM_SIZE {
                        let w = (l & 0xFFFF) as i32;
                        let h = ((l >> 16) & 0xFFFF) as i32;
                        // Drop zero sizes (stale setup-storm leftovers; a
                        // zero-size toplevel has nothing to lay out and the
                        // next real size repairs). See the box handler below.
                        if w <= 0 || h <= 0 { return None; }
                        if let Some(ref mut cb) = *cb.borrow_mut() {
                            cb(w, h);
                        }
                    }
                    None
                },
            ).map_err(|e| Error::Backend(format!("{}", e)))?;

            // Bind raw WM_KEYDOWN/WM_SYSKEYDOWN handler for on_event_key.
            // The state parameter is a GDK-compatible modifier mask:
            //   bit 0 = Shift, bit 2 = Ctrl, bit 3 = Alt (MOD1_MASK)
            // Alt is detected from WM_SYSKEYDOWN; Shift/Ctrl via GetKeyState.
            // If the callback returns 0 (not consumed), the message is forwarded
            // to the focused child window via PostMessage so the canvas or entry
            // raw handlers can process it.  After forwarding, we consume the
            // message (return Some(0)) so DefWindowProc does NOT process it.
            // This prevents WM_SYSKEYDOWN(Alt) from activating the menu bar,
            // which would steal focus from the window-level quit handler.
            let kcb = event_key_cb.clone();
            static KEY_HANDLER_ID: AtomicUsize = AtomicUsize::new(0x30000000);
            let key_id = KEY_HANDLER_ID.fetch_add(1, Ordering::SeqCst);
            let parent_hwnd = hwnd;
            nwg::bind_raw_event_handler(
                &nwg::ControlHandle::Hwnd(hwnd),
                key_id,
                move |_h, msg, w, l| {
                    if msg != winapi::um::winuser::WM_KEYDOWN && msg != winapi::um::winuser::WM_SYSKEYDOWN {
                        return None;
                    }
                    if let Some(ref mut f) = *kcb.borrow_mut() {
                        let mut state: u32 = 0;
                        if msg == winapi::um::winuser::WM_SYSKEYDOWN {
                            state |= 8; // GDK_MOD1_MASK (Alt)
                        }
                        // Win32 key messages carry no modifier state; query
                        // it directly so Shift+arrows (selection extend) and
                        // Ctrl accelerators (Ctrl+Q quit) work. Alt keeps its
                        // message-derived bit above.
                        state |= modifier_state() & 0x5;
                        // Translate Shift pairs/capitals; specials keep raw
                        // VKs. The forwarded PostMessage below stays raw so
                        // the child translates for itself. A colliding
                        // translation (None) skips the callback; flow continues
                        // to forwarding below like any unhandled key.
                        if let Some(k) = translate_vk(w) {
                            if f(k, state) != 0 {
                                return Some(0); // consumed, do not forward
                            }
                        }
                    }
                    // Alt+letter (WM_SYSKEYDOWN): let DefWindowProc activate the
                    // native menu bar so the user can keyboard-navigate the menu
                    // (Alt+letter to open, arrows to move, Enter to activate,
                    // Esc to close). The native menu is modal, so it handles the
                    // navigation keys itself; item activation fires WM_MENUCOMMAND
                    // which dispatches the action. Previously we consumed every
                    // key to stop the menu from stealing focus, which made the
                    // menu mouse-only.
                    if msg == winapi::um::winuser::WM_SYSKEYDOWN {
                        return None;
                    }
                    // Forward other keyboard messages to the correct child.
                    // The window-level raw handler consumes non-Alt key events
                    // (returning Some(0) below) so they are not processed twice;
                    // we manually post the message to the focused child (or a
                    // suitable descendant if no child has focus).
                    unsafe {
                        let focused = winapi::um::winuser::GetFocus();
                        if focused != std::ptr::null_mut() && focused != parent_hwnd {
                            // Forward to the focused child.
                            winapi::um::winuser::PostMessageW(focused, msg, w, l);
                        } else {
                            // No child has focus.  Post to all descendant
                            // windows recursively.  This ensures the formula
                            // entry (great-great-grandchild of the main window
                            // in the VBox -> formula bar -> entry hierarchy)
                            // receives keyboard messages even when no child
                            // has keyboard focus.
                            unsafe fn post_to_descendants(hwnd: winapi::shared::windef::HWND, msg: u32, w: winapi::shared::minwindef::WPARAM, l: winapi::shared::minwindef::LPARAM) {
                                let mut child = winapi::um::winuser::GetWindow(hwnd, winapi::um::winuser::GW_CHILD);
                                while child != std::ptr::null_mut() {
                                    winapi::um::winuser::PostMessageW(child, msg, w, l);
                                    post_to_descendants(child, msg, w, l);
                                    child = winapi::um::winuser::GetWindow(child, winapi::um::winuser::GW_HWNDNEXT);
                                }
                            }
                            post_to_descendants(parent_hwnd, msg, w, l);
                        }
                    }
                    Some(0) // consumed — prevent DefWindowProc from activating menu
                },
            ).map_err(|e| Error::Backend(format!("{}", e)))?;

            // Bind raw WM_CLOSE handler: replayer tests can post WM_CLOSE
            // directly to the main window as a reliable quit mechanism that
            // does not depend on the foreground-window focus state.
            // Calls the registered close callback (save_before_quit) before
            // quitting so that pending edits are committed to the output file.
            {
                let cb = close_cb.clone();
                static CLOSE_ID: AtomicUsize = AtomicUsize::new(0x40000000);
                let cid = CLOSE_ID.fetch_add(1, Ordering::SeqCst);
                nwg::bind_raw_event_handler(
                    &nwg::ControlHandle::Hwnd(hwnd), cid,
                    move |_h, msg, _w, _l| {
                        if msg == winapi::um::winuser::WM_CLOSE {
                            // TEMPORARY Win95 diagnosis.
                            #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
                            unsafe {
                                extern "system" {
                                    fn CreateFileA(name: *const u8, access: u32, share: u32,
                                        sa: *mut std::os::raw::c_void, disp: u32, flags: u32,
                                        tmpl: *mut std::os::raw::c_void) -> *mut std::os::raw::c_void;
                                    fn SetFilePointer(h: *mut std::os::raw::c_void, lo: i32,
                                        hi: *mut i32, how: u32) -> u32;
                                    fn WriteFile(h: *mut std::os::raw::c_void, buf: *const u8,
                                        len: u32, w: *mut u32, ov: *mut std::os::raw::c_void) -> i32;
                                    fn CloseHandle(h: *mut std::os::raw::c_void) -> i32;
                                }
                                let h = CreateFileA(b"c:\\gcorro.log\0".as_ptr(), 0x4000_0000, 1,
                                    std::ptr::null_mut(), 4, 0x80, std::ptr::null_mut());
                                if !h.is_null() && h as isize != -1 {
                                    SetFilePointer(h, 0, std::ptr::null_mut(), 2);
                                    let mut w = 0u32;
                                    let s = b"got-close\n";
                                    WriteFile(h, s.as_ptr(), s.len() as u32, &mut w, std::ptr::null_mut());
                                    CloseHandle(h);
                                }
                            }
                            if let Some(ref mut f) = *cb.borrow_mut() {
                                f();
                            }
                            crate::backends::nwg::quit_main_loop();
                            Some(0)
                        } else {
                            None
                        }
                    },
                ).map_err(|e| Error::Backend(format!("{}", e)))?;
            }
        }

        Ok(Window { hwnd: inner.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void, inner: Rc::new(inner), _handler: Rc::new(handler), root_child, layout_cb, event_key_cb, close_cb })
    }

    // -- Button --

    #[derive(Clone)]
    pub struct Button {
        pub(crate) inner: Rc<nwg::Button>,
        pub(crate) _handler: Rc<nwg::EventHandler>,
        // Stable hwnd snapshot for AsRef (see Window). *mut c_void is
        // Copy, so #[derive(Clone)] keeps working.
        hwnd: *mut c_void,
        pub(crate) click_cb: Rc<RefCell<Option<Box<dyn FnMut()>>>>,
        pub(crate) _raw_click_handler: Rc<RefCell<Vec<nwg::RawEventHandler>>>,
    }

    impl Button {
        pub fn on_click(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            *self.click_cb.borrow_mut() = Some(Box::new(f));
            Ok(0)
        }
        pub fn emit_clicked(&self) -> Result<u64, Error> {
            if let Some(ref mut cb) = *self.click_cb.borrow_mut() { cb(); }
            Ok(0)
        }
        pub fn set_size_request(&self, w: i32, h: i32) {
            if let Some(hwnd) = self.inner.handle.hwnd() {
                unsafe {
                    winapi::um::winuser::SetWindowPos(
                        hwnd as _,
                        std::ptr::null_mut(), 0, 0, w, h,
                        winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_NOMOVE | winapi::um::winuser::SWP_SHOWWINDOW,
                    );
                }
            }
        }
        pub fn set_font_style(&self, weight: i32, italic: bool) {
            if let Some(hwnd) = self.inner.handle.hwnd() {
                unsafe {
                    let mut lf: winapi::um::wingdi::LOGFONTW = std::mem::zeroed();
                    lf.lfHeight = -13;
                    lf.lfWeight = weight;
                    lf.lfItalic = italic as u8;
                    lf.lfCharSet = winapi::um::wingdi::ANSI_CHARSET as u8;
                    lf.lfOutPrecision = winapi::um::wingdi::OUT_DEFAULT_PRECIS as u8;
                    lf.lfClipPrecision = winapi::um::wingdi::CLIP_DEFAULT_PRECIS as u8;
                    lf.lfQuality = winapi::um::wingdi::PROOF_QUALITY as u8;
                    lf.lfPitchAndFamily = winapi::um::wingdi::DEFAULT_PITCH as u8;
                    let face = "Segoe UI\0".encode_utf16().collect::<Vec<_>>();
                    let mut i = 0;
                    while i < face.len().min(32) {
                        lf.lfFaceName[i] = face[i];
                        i += 1;
                    }
                    let hfont = winapi::um::wingdi::CreateFontIndirectW(&lf);
                    if !hfont.is_null() {
                        winapi::um::winuser::SendMessageW(
                            hwnd as _,
                            winapi::um::winuser::WM_SETFONT,
                            hfont as usize,
                            1,
                        );
                    }
                }
            }
        }
        pub fn add_class(&self, _class: &str) {}
        pub fn remove_class(&self, _class: &str) {}
    }

    impl AsRef<*mut c_void> for Button {
        fn as_ref(&self) -> &*mut c_void {
            &self.hwnd
        }
    }
    pub fn create_button(parent: *mut c_void, text: &str) -> Result<Button, Error> {
        let (inner, click_cb, handler) = crate::backends::nwg::create_button(parent, text)
            .map_err(|e| Error::Backend(format!("{}", e)))?;
        let hwnd = inner.handle.hwnd().unwrap_or(std::ptr::null_mut());
        let raw_handlers: Rc<RefCell<Vec<nwg::RawEventHandler>>> = Rc::new(RefCell::new(Vec::new()));
        if hwnd != std::ptr::null_mut() {
            static RAW_BTN_CLICK_ID: AtomicUsize = AtomicUsize::new(0x60000000);
            let cb = click_cb.clone();
            let rid = RAW_BTN_CLICK_ID.fetch_add(1, Ordering::SeqCst);
            if let Ok(raw) = nwg::bind_raw_event_handler(
                &nwg::ControlHandle::Hwnd(hwnd), rid,
                move |_h, msg, w, l| {
                    match msg {
                        winapi::um::winuser::WM_LBUTTONUP => {
                            let mut rect = std::mem::MaybeUninit::zeroed();
                            let rc = unsafe {
                                winapi::um::winuser::GetClientRect(hwnd as _, rect.as_mut_ptr());
                                rect.assume_init()
                            };
                            let x = (l & 0xFFFF) as i16;
                            let y = ((l >> 16) & 0xFFFF) as i16;
                            if i32::from(x) <= rc.right && i32::from(y) <= rc.bottom {
                                if let Some(ref mut f) = *cb.borrow_mut() { f(); }
                            }
                            None
                        }
                        winapi::um::winuser::WM_KEYUP => {
                            if w == winapi::um::winuser::VK_SPACE as usize {
                                if let Some(ref mut f) = *cb.borrow_mut() { f(); }
                            }
                            None
                        }
                        _ => None
                    }
                },
            ) {
                raw_handlers.borrow_mut().push(raw);
            }
        }
        Ok(Button { hwnd: hwnd as *mut c_void, inner: Rc::new(inner), _handler: Rc::new(handler), click_cb, _raw_click_handler: raw_handlers })
    }

    // -- Label --

    #[derive(Clone)]
    pub struct Label {
        pub(crate) inner: Rc<nwg::Label>,
        // Stable hwnd snapshot for AsRef (see Window).
        hwnd: *mut c_void,
    }

    impl Label {
        pub fn set_text(&self, text: &str) {
            self.inner.set_text(text);
            // Nudge the parent to re-run its layout (if it has a WM_SIZE
            // layout handler, i.e. a BoxWidget): label width is measured
            // from text at layout time, so a text change must re-layout
            // to keep the label fitted (parity with GTK auto-sizing).
            // Harmless when the parent has no such handler.
            if let Some(hwnd) = self.inner.handle.hwnd() {
                unsafe {
                    let parent = winapi::um::winuser::GetParent(hwnd as _);
                    if !parent.is_null() {
                        let mut rect: winapi::shared::windef::RECT = std::mem::zeroed();
                        winapi::um::winuser::GetClientRect(parent, &mut rect);
                        let l = (((rect.bottom & 0xFFFF) << 16) | (rect.right & 0xFFFF)) as isize;
                        winapi::um::winuser::PostMessageW(parent,
                            winapi::um::winuser::WM_SIZE, 0, l as _);
                    }
                }
            }
        }
        pub fn get_text(&self) -> Option<String> { Some(self.inner.text()) }
        pub fn set_visible(&self, visible: bool) { self.inner.set_visible(visible); }
        pub fn set_markup(&self, markup: &str) { self.inner.set_text(markup); }
        pub fn set_margin_start(&self, _px: i32) {}
        pub fn set_margin_top(&self, _px: i32) {}
        pub fn set_halign(&self, _align: i32) {}
        pub fn set_valign(&self, _align: i32) {}
        pub fn raw_handle(&self) -> *mut c_void {
            self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void
        }
    }

    impl AsRef<*mut c_void> for Label {
        fn as_ref(&self) -> &*mut c_void {
            &self.hwnd
        }
    }

    impl Widget for Label {
        fn raw_handle(&self) -> *mut c_void {
            self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void
        }
    }

    pub fn create_label(parent: *mut c_void) -> Result<Label, Error> {
        crate::backends::nwg::create_label(parent).map(|l| {
            let hwnd = l.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void;
            Label { inner: Rc::new(l), hwnd }
        }).map_err(|e| Error::Backend(format!("{}", e)))
    }

    // -- BoxWidget --

    pub struct BoxWidget {
        pub(crate) frame: Option<Rc<nwg::Frame>>,
        pub(crate) hwnd: *mut c_void,
        pub(crate) children: Rc<RefCell<Vec<*mut c_void>>>,
        pub(crate) child_vexpand: Rc<RefCell<Vec<bool>>>,
        pub(crate) child_hexpand: Rc<RefCell<Vec<bool>>>,
        pub(crate) orientation: crate::backends::nwg::Orientation,
        pub(crate) spacing: i32,
    }

    impl Clone for BoxWidget {
        fn clone(&self) -> Self {
            BoxWidget {
                frame: self.frame.clone(),
                hwnd: self.hwnd,
                children: self.children.clone(),
                child_vexpand: self.child_vexpand.clone(),
                child_hexpand: self.child_hexpand.clone(),
                orientation: self.orientation,
                spacing: self.spacing,
            }
        }
    }

    impl AsRef<*mut c_void> for BoxWidget {
        fn as_ref(&self) -> &*mut c_void {
            &self.hwnd
        }
    }

    impl Widget for BoxWidget {
        fn raw_handle(&self) -> *mut c_void {
            self.hwnd
        }
    }

    pub trait Appendable {
        fn collect_hwnds(&self) -> Vec<*mut c_void>;
    }

    impl<T: AsRef<*mut c_void>> Appendable for T {
        fn collect_hwnds(&self) -> Vec<*mut c_void> {
            let ptr = *self.as_ref();
            if ptr.is_null() { vec![] } else { vec![ptr] }
        }
    }

    impl BoxWidget {
        pub fn append(&self, child: &impl Appendable) {
            let hwnds = child.collect_hwnds();
            // A null HWND can never participate in layout (it is filtered
            // below) — silently dropping it crams the remaining widgets
            // with no diagnostic. Fail loudly in debug builds instead.
            debug_assert!(
                hwnds.iter().all(|&p| !p.is_null()),
                "BoxWidget::append: child with null HWND cannot be laid out"
            );
            for &ptr in &hwnds {
                if !ptr.is_null() && !self.hwnd.is_null() {
                    unsafe {
                        winapi::um::winuser::SetParent(ptr as _, self.hwnd as _);
                    }
                }
            }
            let mut children = self.children.borrow_mut();
            let mut vex = self.child_vexpand.borrow_mut();
            let mut hex = self.child_hexpand.borrow_mut();
            children.extend(hwnds.into_iter().filter(|&c| !c.is_null()));
            vex.resize(children.len(), false);
            hex.resize(children.len(), false);
            drop(children);
            drop(vex);
            drop(hex);
            // Never depend on a future WM_SIZE: lay out now with the
            // current size (re-runs on every later WM_SIZE anyway).
            self.request_layout();
        }
        pub fn set_child_vexpand(&self, child: &impl AsRef<*mut c_void>, expand: bool) {
            let ptr = *child.as_ref();
            let children = self.children.borrow();
            let mut vex = self.child_vexpand.borrow_mut();
            // Silently missing the child leaves it fixed-size (fx-bar cram
            // shape): fail loudly in debug builds instead.
            debug_assert!(
                children.iter().any(|&c| c == ptr),
                "BoxWidget::set_child_vexpand: child not found in box"
            );
            if let Some(idx) = children.iter().position(|&c| c == ptr) {
                vex[idx] = expand;
            }
            drop(children);
            drop(vex);
            self.request_layout();
        }
        pub fn set_child_hexpand(&self, child: &impl AsRef<*mut c_void>, expand: bool) {
            let ptr = *child.as_ref();
            let children = self.children.borrow();
            let mut hex = self.child_hexpand.borrow_mut();
            // Silently missing the child leaves it fixed-size (fx-bar cram
            // shape): fail loudly in debug builds instead.
            debug_assert!(
                children.iter().any(|&c| c == ptr),
                "BoxWidget::set_child_hexpand: child not found in box"
            );
            if let Some(idx) = children.iter().position(|&c| c == ptr) {
                hex[idx] = expand;
            }
            drop(children);
            drop(hex);
            self.request_layout();
        }
        pub fn layout(&self, _x: i32, _y: i32, w: i32, h: i32) {
            // TEMPORARY Win95 diagnosis.
            #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
            mark95xy(b"layot", w, h);
            // Negative/zero sizes arise transiently (a box laid out before
            // its parent is sized, or a wrapped synthetic WM_SIZE). A
            // negative size is never valid: SetWindowPos clamps it to 0
            // (hiding the child) while the synthetic WM_SIZE below would
            // wrap it to ~65526 and fling children off-screen. Clamp here
            // so garbage layouts are harmless no-ops at the right place.
            let w = w.max(0);
            let h = h.max(0);
            let children = self.children.borrow();
            let vex = self.child_vexpand.borrow();
            let hex = self.child_hexpand.borrow();
            let n = children.len();
            if n == 0 { return; }
            let (fixed_w, fixed_h) = match self.orientation {
                crate::backends::nwg::Orientation::Horizontal => {
                    (0, h - 10)
                }
                crate::backends::nwg::Orientation::Vertical => {
                    (w - 10, 0)
                }
            };
            let mut desired_sizes: Vec<i32> = Vec::with_capacity(n);
            for i in 0..n {
                // Hidden children take no space (e.g. the sheet tab strip
                // with a single sheet): hiding alone would otherwise leave
                // a blank gap in the layout.
                let hidden = unsafe {
                    winapi::um::winuser::IsWindowVisible(children[i] as _) == 0
                };
                if hidden {
                    desired_sizes.push(0);
                    continue;
                }
                let is_expand = match self.orientation {
                    crate::backends::nwg::Orientation::Horizontal => hex[i],
                    crate::backends::nwg::Orientation::Vertical => vex[i],
                };
                if !is_expand {
                    let hardcoded = match self.orientation {
                        crate::backends::nwg::Orientation::Horizontal => 60,
                        crate::backends::nwg::Orientation::Vertical => 28,
                    };
                    // Labels (STATIC controls) are fitted to their text
                    // (parity with GTK auto-sizing); NWG gives them a wide
                    // default that would otherwise leave gaps in the row.
                    // Other controls keep their current window size.
                    let fitted = match self.orientation {
                        crate::backends::nwg::Orientation::Horizontal =>
                            static_text_width(children[i] as _),
                        crate::backends::nwg::Orientation::Vertical => None,
                    };
                    if let Some(w) = fitted {
                        desired_sizes.push(w);
                        continue;
                    }
                    unsafe {
                        let mut rect: winapi::shared::windef::RECT = std::mem::zeroed();
                        if winapi::um::winuser::GetWindowRect(children[i] as _, &mut rect) != 0 {
                            let sz = match self.orientation {
                                crate::backends::nwg::Orientation::Horizontal => rect.right - rect.left,
                                crate::backends::nwg::Orientation::Vertical => rect.bottom - rect.top,
                            };
                            desired_sizes.push(if sz > 10 { sz } else { hardcoded });
                        } else {
                            desired_sizes.push(hardcoded);
                        }
                    }
                } else {
                    desired_sizes.push(0);
                }
            }
            // Shared span math (unit-tested in win32_portable): fixed sizes
            // from `desired_sizes`, expanders splitting the remainder with
            // GTK fill parity (no lost remainder pixel).
            let flags: Vec<bool> = (0..n)
                .map(|i| match self.orientation {
                    crate::backends::nwg::Orientation::Horizontal => hex[i],
                    crate::backends::nwg::Orientation::Vertical => vex[i],
                })
                .collect();
            let avail = match self.orientation {
                crate::backends::nwg::Orientation::Horizontal => w - 10,
                crate::backends::nwg::Orientation::Vertical => h - 10,
            };
            let spans = crate::win32_portable::distribute_spans(
                5,
                avail,
                self.spacing,
                &desired_sizes,
                &flags,
            );

            for i in 0..n {
                let child = children[i];
                let (pos, span) = spans[i];
                let (cw, ch) = match self.orientation {
                    crate::backends::nwg::Orientation::Horizontal => (span, fixed_h),
                    crate::backends::nwg::Orientation::Vertical => (fixed_w, span),
                };
                let (cx, cy) = match self.orientation {
                    crate::backends::nwg::Orientation::Horizontal => (pos, 5),
                    crate::backends::nwg::Orientation::Vertical => (5, pos),
                };
                // Clamp the span: distribute_spans can return negatives
                // when the box itself was laid out at/near zero size
                // (avail = w - 10 < fixed total). A negative size must
                // never reach SetWindowPos (Wine clamps to 0, hiding the
                // child) nor the synthetic WM_SIZE below (it would wrap
                // to ~65526 and fling nested children off-screen).
                let (cw, ch) = (cw.max(0), ch.max(0));
                set_window_pos(child, cx, cy, cw, ch);
                // Airtight cascade: SetWindowPos only delivers WM_SIZE when
                // the size actually changed, so a nested box that keeps its
                // size would never re-lay-out its own children (the cram
                // failure). Synthesize WM_SIZE unconditionally — leaf
                // controls ignore it, nested boxes re-run their layout.
                // Guard: a zero-size child has nothing to lay out, and
                // packing a non-positive size would wrap (see above).
                if cw > 0 && ch > 0 {
                    unsafe {
                        let l = ((ch & 0xFFFF) << 16) | (cw & 0xFFFF);
                        winapi::um::winuser::SendMessageW(
                            child as _,
                            winapi::um::winuser::WM_SIZE,
                            0,
                            l as _,
                        );
                    }
                }
            }
        }
        /// Re-run layout with the box's current client size. Called after
        /// every structural mutation (append, expand-flag change) so layout
        /// never depends on a future WM_SIZE that may never arrive.
        pub fn request_layout(&self) {
            if self.hwnd.is_null() {
                return;
            }
            unsafe {
                let mut rect: winapi::shared::windef::RECT = std::mem::zeroed();
                if winapi::um::winuser::GetClientRect(self.hwnd as _, &mut rect) != 0 {
                    self.layout(0, 0, rect.right - rect.left, rect.bottom - rect.top);
                }
            }
        }
        pub fn set_vexpand(&self, _expand: bool) {}
        pub fn set_hexpand(&self, _expand: bool) {}
    }

    /// Measure a STATIC (label) control's text width in pixels, for
    /// shrink-to-fit layout (parity with GTK label auto-sizing).
    /// Returns None for non-label controls or on any measurement failure
    /// (callers fall back to the window rect / hardcoded size).
    fn static_text_width(hwnd: winapi::shared::windef::HWND) -> Option<i32> {
        unsafe {
            let mut cls: [u16; 256] = [0; 256];
            let n = winapi::um::winuser::GetClassNameW(hwnd, cls.as_mut_ptr(), 256);
            if n <= 0 { return None; }
            if !String::from_utf16_lossy(&cls[..n as usize]).eq_ignore_ascii_case("Static") {
                return None;
            }
            let tlen = winapi::um::winuser::GetWindowTextLengthW(hwnd);
            let mut buf: Vec<u16> = vec![0; (tlen + 1) as usize];
            winapi::um::winuser::GetWindowTextW(hwnd, buf.as_mut_ptr(), tlen + 1);
            let hdc = winapi::um::winuser::GetDC(hwnd);
            if hdc.is_null() { return None; }
            let hfont = winapi::um::winuser::SendMessageW(hwnd, winapi::um::winuser::WM_GETFONT, 0, 0);
            let old = if hfont != 0 {
                winapi::um::wingdi::SelectObject(hdc, hfont as _)
            } else {
                std::ptr::null_mut()
            };
            let mut sz: winapi::shared::windef::SIZE = std::mem::zeroed();
            let ok = winapi::um::wingdi::GetTextExtentPoint32W(hdc, buf.as_ptr(), tlen, &mut sz);
            if !old.is_null() { winapi::um::wingdi::SelectObject(hdc, old); }
            winapi::um::winuser::ReleaseDC(hwnd, hdc);
            if ok == 0 { return None; }
            Some(sz.cx + 8) // small horizontal padding
        }
    }

    pub fn create_box(orientation: crate::backends::nwg::Orientation, spacing: i32, parent: *mut c_void) -> Result<BoxWidget, Error> {
        let mut frame = nwg::Frame::default();
        if !parent.is_null() {
            nwg::Frame::builder()
                .flags(nwg::FrameFlags::VISIBLE)
                .size((0, 0))
                .position((0, 0))
                .parent(&nwg::ControlHandle::Hwnd(parent as _))
                .build(&mut frame)
                .map_err(|e| Error::Backend(format!("{}", e)))?;
        }
        let hwnd = frame.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void;
        let children: Rc<RefCell<Vec<*mut c_void>>> = Rc::new(RefCell::new(Vec::new()));
        let child_vexpand: Rc<RefCell<Vec<bool>>> = Rc::new(RefCell::new(Vec::new()));
        let child_hexpand: Rc<RefCell<Vec<bool>>> = Rc::new(RefCell::new(Vec::new()));
        let bw = BoxWidget {
            frame: Some(Rc::new(frame)), hwnd,
            children: children.clone(),
            child_vexpand: child_vexpand.clone(),
            child_hexpand: child_hexpand.clone(),
            orientation, spacing,
        };
        // Auto-layout on WM_SIZE — now shares children via Rc<RefCell>
        if hwnd != std::ptr::null_mut() {
            let bw2 = bw.clone();
            static BOX_SIZE_ID: AtomicUsize = AtomicUsize::new(0xB0000000);
            let id = BOX_SIZE_ID.fetch_add(1, Ordering::SeqCst);
            let _ = nwg::bind_raw_event_handler(
                &nwg::ControlHandle::Hwnd(hwnd as _), id,
                move |_h, msg, _w, l| {
                    if msg == winapi::um::winuser::WM_SIZE {
                        let w = (l & 0xFFFF) as i32;
                        let h = ((l >> 16) & 0xFFFF) as i32;
                        // A zero-size box has nothing to lay out, and these
                        // arrive as stale queue leftovers from the un-pumped
                        // setup storm (pump_events is a no-op on Windows):
                        // honoring them crushes children to 0 and the white
                        // screen never repairs (nothing re-invalidates).
                        // The next real (non-zero) size re-runs layout.
                        if w <= 0 || h <= 0 { return None; }
                        bw2.layout(0, 0, w, h);
                    }
                    None
                },
            );
        }
        Ok(bw)
    }

    // -- Grid --

    pub struct Grid;

    impl Grid {
        pub fn attach(&self, child: &impl AsRef<*mut c_void>, left: i32, top: i32, width: i32, height: i32) {
            set_window_pos(*child.as_ref(), left, top, width, height);
        }
    }

    pub fn create_grid() -> Result<Grid, Error> { Ok(Grid) }

    // -- Entry --

    pub struct Entry {
        pub(crate) inner: Rc<nwg::TextInput>,
        pub(crate) _handler: Rc<nwg::EventHandler>,
        // Stable hwnd snapshot for AsRef (see Window).
        hwnd: *mut c_void,
        pub(crate) changed_cb: Rc<RefCell<Option<Box<dyn FnMut()>>>>,
        focus_in_cb: Rc<RefCell<Option<Box<dyn FnMut(*mut c_void) -> i32>>>>,
        focus_out_cb: Rc<RefCell<Option<Box<dyn FnMut(*mut c_void) -> i32>>>>,
        _focus_in_handler: Option<nwg::RawEventHandler>,
        _focus_out_handler: Option<nwg::RawEventHandler>,
        key_cb: Rc<RefCell<Option<Box<dyn FnMut(u32, u32) -> bool>>>>,
        _key_handler: Option<nwg::RawEventHandler>,
        activate_cb: Rc<RefCell<Option<Box<dyn FnMut(*mut c_void)>>>>,
        pub(crate) pos_x: std::cell::Cell<i32>,
        pub(crate) pos_y: std::cell::Cell<i32>,
    }

    impl Clone for Entry {
        fn clone(&self) -> Self {
            Entry {
                inner: self.inner.clone(),
                _handler: self._handler.clone(),
                changed_cb: self.changed_cb.clone(),
                focus_in_cb: self.focus_in_cb.clone(),
                focus_out_cb: self.focus_out_cb.clone(),
                _focus_in_handler: None,
                _focus_out_handler: None,
                key_cb: self.key_cb.clone(),
                _key_handler: None,
                activate_cb: self.activate_cb.clone(),
                hwnd: self.hwnd,
                pos_x: std::cell::Cell::new(self.pos_x.get()),
                pos_y: std::cell::Cell::new(self.pos_y.get()),
            }
        }
    }

    impl Entry {
        pub fn set_text(&self, text: &str) {
            self.inner.set_text(text);
            // Caret to end: SetWindowText leaves the caret at position 0,
            // so a natively-inserted char (a WM_CHAR the key handlers
            // declined, e.g. '(') would land at the front ("(=" for
            // "=("), fail the resync guard, and be clobbered by the next
            // sync. Caret-at-end keeps native insertions appending, which
            // is also what every other backend does after a programmatic
            // set. Harmless when unfocused (caret invisible).
            if let Some(hwnd) = self.inner.handle.hwnd() {
                unsafe {
                    let len = winapi::um::winuser::GetWindowTextLengthW(hwnd as _);
                    winapi::um::winuser::SendMessageW(
                        hwnd as _,
                        winapi::um::winuser::EM_SETSEL as u32,
                        len as usize,
                        len as isize,
                    );
                }
            }
        }
        pub fn get_text(&self) -> Option<String> { Some(self.inner.text()) }
        pub fn connect_changed(&self, f: impl FnMut() + 'static) -> Result<u64, Error> {
            *self.changed_cb.borrow_mut() = Some(Box::new(f));
            Ok(0)
        }
        pub fn set_width_chars(&self, n: i32) {
            if let Some(hwnd) = self.inner.handle.hwnd() {
                unsafe {
                    let mut rect: winapi::shared::windef::RECT = std::mem::zeroed();
                    if winapi::um::winuser::GetWindowRect(hwnd as _, &mut rect) != 0 {
                        let h = rect.bottom - rect.top;
                        winapi::um::winuser::SetWindowPos(
                            hwnd as _, std::ptr::null_mut(), 0, 0, n * 8, h,
                            winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_NOMOVE | winapi::um::winuser::SWP_SHOWWINDOW,
                        );
                    }
                }
            }
        }
        pub fn set_size_request(&self, w: i32, h: i32) {
            if let Some(hwnd) = self.inner.handle.hwnd() {
                unsafe {
                    winapi::um::winuser::SetWindowPos(
                        hwnd as _, std::ptr::null_mut(), 0, 0, w, h,
                        winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_NOMOVE | winapi::um::winuser::SWP_SHOWWINDOW,
                    );
                }
            }
        }
        pub fn set_visible(&self, v: bool) { self.inner.set_visible(v); }
        pub fn grab_focus(&self) {
            let _ = self.inner.set_focus();
            if let Some(hwnd) = self.inner.handle.hwnd() {
                unsafe {
                    winapi::um::winuser::ShowWindow(hwnd as _, winapi::um::winuser::SW_SHOW);
                    winapi::um::winuser::SetWindowPos(
                        hwnd as _, winapi::um::winuser::HWND_TOP,
                        0, 0, 0, 0,
                        winapi::um::winuser::SWP_NOMOVE | winapi::um::winuser::SWP_NOSIZE | winapi::um::winuser::SWP_SHOWWINDOW,
                    );
                    winapi::um::winuser::RedrawWindow(
                        hwnd as _,
                        std::ptr::null_mut(),
                        std::ptr::null_mut(),
                        winapi::um::winuser::RDW_INVALIDATE | winapi::um::winuser::RDW_UPDATENOW | winapi::um::winuser::RDW_ERASE | winapi::um::winuser::RDW_FRAME,
                    );
                    let parent = winapi::um::winuser::GetParent(hwnd as _);
                    if !parent.is_null() {
                        winapi::um::winuser::RedrawWindow(
                            parent,
                            std::ptr::null_mut(),
                            std::ptr::null_mut(),
                            winapi::um::winuser::RDW_INVALIDATE | winapi::um::winuser::RDW_UPDATENOW | winapi::um::winuser::RDW_ALLCHILDREN | winapi::um::winuser::RDW_FRAME,
                        );
                    }
                }
            }
        }
        pub fn on_key(&self, mut f: Box<dyn FnMut(u32) -> bool>) {
            *self.key_cb.borrow_mut() = Some(Box::new(move |k: u32, _s: u32| -> bool { f(k) }));
        }
        pub fn set_hexpand(&self, _expand: bool) {}
        pub fn set_vexpand(&self, _expand: bool) {}
        pub fn add_class(&self, _class: &str) {}
        pub fn remove_class(&self, _class: &str) {}
        pub fn set_margin_start(&self, px: i32) {
            self.pos_x.set(px);
            if let Some(hwnd) = self.inner.handle.hwnd() {
                unsafe {
                    let _ = winapi::um::winuser::SetWindowPos(
                        hwnd as _, std::ptr::null_mut(), px, self.pos_y.get(), 0, 0,
                        winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_NOSIZE | winapi::um::winuser::SWP_SHOWWINDOW,
                    );
                }
            } else {
            }
        }
        pub fn set_margin_top(&self, px: i32) {
            self.pos_y.set(px);
            if let Some(hwnd) = self.inner.handle.hwnd() {
                unsafe {
                    let _ = winapi::um::winuser::SetWindowPos(
                        hwnd as _, std::ptr::null_mut(), self.pos_x.get(), px, 0, 0,
                        winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_NOSIZE | winapi::um::winuser::SWP_SHOWWINDOW,
                    );
                }
            } else {
            }
        }
        pub fn set_halign(&self, _align: i32) {}
        pub fn set_valign(&self, _align: i32) {}
        pub fn on_key_raw(&self, cb: Box<dyn FnMut(u32, u32) -> bool>) {
            *self.key_cb.borrow_mut() = Some(cb);
        }
        pub fn connect_activate(&self, f: impl FnMut(*mut c_void) + 'static) -> Result<u64, Error> {
            *self.activate_cb.borrow_mut() = Some(Box::new(f));
            Ok(0)
        }
        pub fn connect_focus_in_event(&self, f: impl FnMut(*mut c_void) -> i32 + 'static) -> Result<u64, Error> {
            *self.focus_in_cb.borrow_mut() = Some(Box::new(f));
            Ok(0)
        }
        pub fn connect_focus_out_event(&self, f: impl FnMut(*mut c_void) -> i32 + 'static) -> Result<u64, Error> {
            *self.focus_out_cb.borrow_mut() = Some(Box::new(f));
            Ok(0)
        }
    }

    impl AsRef<*mut c_void> for Entry {
        fn as_ref(&self) -> &*mut c_void {
            &self.hwnd
        }
    }

    impl Widget for Entry {
        fn raw_handle(&self) -> *mut c_void {
            self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void
        }
    }

    pub fn create_entry(parent: *mut c_void) -> Result<Entry, Error> {
        let (inner, changed_cb, handler) = crate::backends::nwg::create_entry(parent)
            .map_err(|e| Error::Backend(format!("{}", e)))?;
        #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
        mark95a(b"a-e0\n");
        let focus_in_cb: Rc<RefCell<Option<Box<dyn FnMut(*mut c_void) -> i32>>>> = Rc::new(RefCell::new(None));
        let focus_out_cb: Rc<RefCell<Option<Box<dyn FnMut(*mut c_void) -> i32>>>> = Rc::new(RefCell::new(None));
        let key_cb: Rc<RefCell<Option<Box<dyn FnMut(u32, u32) -> bool>>>> = Rc::new(RefCell::new(None));
        let activate_cb: Rc<RefCell<Option<Box<dyn FnMut(*mut c_void)>>>> =
            Rc::new(RefCell::new(None));
        let hwnd = inner.handle.hwnd().unwrap_or(std::ptr::null_mut());
        #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
        mark95xy(b"a-hwn", hwnd as i32, 0);
        if hwnd != std::ptr::null_mut() {
            unsafe {
                let ex = winapi::um::winuser::GetWindowLongW(
                    hwnd as _, winapi::um::winuser::GWL_EXSTYLE);
                #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
                mark95xy(b"a-exs", ex, 0);
                winapi::um::winuser::SetWindowLongW(
                    hwnd as _, winapi::um::winuser::GWL_EXSTYLE,
                    ex | winapi::um::winuser::WS_EX_CLIENTEDGE as i32);
                #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
                mark95a(b"a-slw\n");
                winapi::um::winuser::SetWindowPos(
                    hwnd as _, std::ptr::null_mut(), 0, 0, 0, 0,
                    winapi::um::winuser::SWP_NOMOVE | winapi::um::winuser::SWP_NOSIZE
                    | winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_FRAMECHANGED);
            }
        }

        #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
        mark95a(b"a-e1\n");
        let _focus_in_handler = if hwnd != std::ptr::null_mut() {
            let cb = focus_in_cb.clone();
            static FOCUS_IN_ID: AtomicUsize = AtomicUsize::new(0x40000000);
            let id = FOCUS_IN_ID.fetch_add(1, Ordering::SeqCst);
            nwg::bind_raw_event_handler(
                &nwg::ControlHandle::Hwnd(hwnd), id,
                move |_h, msg, _w, _l| {
                    if msg == winapi::um::winuser::WM_SETFOCUS {
                        if let Some(ref mut f) = *cb.borrow_mut() { f(std::ptr::null_mut()); }
                    }
                    None
                },
            ).ok()
        } else { None };

        #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
        mark95a(b"a-e2\n");
        let _focus_out_handler = if hwnd != std::ptr::null_mut() {
            let cb = focus_out_cb.clone();
            static FOCUS_OUT_ID: AtomicUsize = AtomicUsize::new(0x50000000);
            let id = FOCUS_OUT_ID.fetch_add(1, Ordering::SeqCst);
            nwg::bind_raw_event_handler(
                &nwg::ControlHandle::Hwnd(hwnd), id,
                move |_h, msg, _w, _l| {
                    if msg == winapi::um::winuser::WM_KILLFOCUS {
                        if let Some(ref mut f) = *cb.borrow_mut() { f(std::ptr::null_mut()); }
                    }
                    None
                },
            ).ok()
        } else { None };

        #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
        mark95a(b"a-e3\n");
        let _key_handler = if hwnd != std::ptr::null_mut() {
            let kc = key_cb.clone();
            let act = activate_cb.clone();
            let act_hwnd = hwnd as *mut c_void;
            // When a WM_KEYDOWN is consumed by key_cb (the app handled the
            // key itself and synced the widget text), the message loop's
            // TranslateMessage still posts a WM_CHAR for it, and the edit
            // control's default handler would insert the char a second time
            // ("AA" for a single keypress). Remember the consumed key and
            // swallow its WM_CHAR (same idea as the Enter/Escape suppression
            // below). Only printable VKs set the flag — they reliably produce
            // exactly one WM_CHAR; arrows etc. produce none.
            let suppress_char: std::rc::Rc<std::cell::Cell<bool>> =
                std::rc::Rc::new(std::cell::Cell::new(false));
            let suppress_set = suppress_char.clone();
            let suppress_get = suppress_char.clone();
            static ENTRY_KEY_ID: AtomicUsize = AtomicUsize::new(0x60000000);
            let id = ENTRY_KEY_ID.fetch_add(1, Ordering::SeqCst);
            nwg::bind_raw_event_handler(
                &nwg::ControlHandle::Hwnd(hwnd), id,
                move |_h, msg, w, _l| {
                    // Consume WM_CHAR for Enter/Escape to prevent beep
                    if msg == winapi::um::winuser::WM_CHAR {
                        let c = (w & 0xFF) as u8;
                        if c == 0x0D || c == 0x1B {
                            return Some(0);
                        }
                        if suppress_get.get() {
                            suppress_get.set(false);
                            return Some(0);
                        }
                        return None;
                    }
                    if msg == winapi::um::winuser::WM_KEYDOWN || msg == winapi::um::winuser::WM_SYSKEYDOWN {
                        // Pure Alt+key belongs to the native menu (DefWindowProc
                        // opens the popup); decline so it is neither typed nor
                        // swallowed. Without this, Alt+F in a focused entry
                        // was consumed as printable text and menus were
                        // mouse-only. Ctrl+Alt (AltGr) still passes through.
                        if pure_alt_held() {
                            return None;
                        }
                        // GTK4 parity: an entry-level activate callback owns
                        // Return exclusively (the CAPTURE handler consumes it
                        // before on_key_raw fires there), so exactly one
                        // submit path runs per press on every backend.
                        if crate::win32_portable::entry_return_fires_activate(
                            (w & 0xFF) as u32,
                            act.borrow().is_some(),
                        ) {
                            if let Some(ref mut a) = *act.borrow_mut() {
                                a(act_hwnd);
                            }
                            return Some(0);
                        }
                        if let Some(ref mut f) = *kc.borrow_mut() {
                            // Full modifier mask (Shift/Ctrl/Alt): the app
                            // needs Shift for arrows, Ctrl/Alt to decline
                            // accelerators (copy/paste stay native) instead
                            // of typing them.
                            let mods = modifier_state();
                            // Translate Shift pairs/capitals (Shift+9 is '(',
                            // not '9'); specials keep raw VKs for dispatch. A
                            // colliding translation (None) declines so native
                            // insertion + resync handle the character.
                            if let Some(k) = translate_vk(w) {
                                if f(k, mods) {
                                    // Suppress the follow-on WM_CHAR exactly
                                    // when this VK produces one — otherwise
                                    // the native control inserts a second
                                    // copy ("==" for '='). Always *set*
                                    // (never just arm): a stale flag from a
                                    // non-producing key would swallow the
                                    // *next* genuine character.
                                    let vk = w & 0xFF;
                                    let mut suppress =
                                        crate::win32_portable::vk_produces_wm_char(vk as u32);
                                    if suppress && (0x60..=0x6F).contains(&vk) {
                                        // Numpad without NumLock is navigation
                                        // (no WM_CHAR follows).
                                        unsafe {
                                            const VK_NUMLOCK: i32 = 0x90;
                                            if winapi::um::winuser::GetKeyState(VK_NUMLOCK) as u16
                                                & 1
                                                == 0
                                            {
                                                suppress = false;
                                            }
                                        }
                                    }
                                    suppress_set.set(suppress);
                                    return Some(0);
                                }
                            } else {
                                return None;
                            }
                        }
                    }
                    None
                },
            ).ok()
        } else { None };

        #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
        mark95a(b"a-e4\n");
        Ok(Entry { hwnd: hwnd as *mut c_void, inner: Rc::new(inner), _handler: Rc::new(handler), changed_cb, focus_in_cb, focus_out_cb, _focus_in_handler, _focus_out_handler, key_cb, _key_handler, activate_cb, pos_x: std::cell::Cell::new(0), pos_y: std::cell::Cell::new(0) })
    }

    // ========== DropDown ==========

    #[derive(Clone)]
    pub struct DropDown {
        pub(crate) inner: Rc<nwg::ComboBox<String>>,
        pub(crate) _handler: Rc<nwg::EventHandler>,
        // Stable hwnd snapshot for AsRef (see Window).
        hwnd: *mut c_void,
    }

    impl DropDown {
        pub fn set_active(&self, index: Option<u32>) {
            self.inner.set_selection(index.map(|i| i as usize));
        }
        pub fn active(&self) -> Option<u32> {
            self.inner.selection().map(|i| i as u32)
        }
        pub fn get_active(&self) -> u32 {
            self.inner.selection().unwrap_or(0) as u32
        }
        pub fn connect_changed(&self, _f: impl FnMut() + 'static) -> Result<u64, Error> {
            Ok(0)
        }
        pub fn set_hexpand(&self, _expand: bool) {}
        pub fn set_vexpand(&self, _expand: bool) {}
    }

    impl AsRef<*mut c_void> for DropDown {
        fn as_ref(&self) -> &*mut c_void {
            &self.hwnd
        }
    }

    pub fn create_dropdown(parent: *mut c_void, items: &[&str]) -> Result<DropDown, Error> {
        let (inner, handler) = crate::backends::nwg::create_dropdown(parent, items)
            .map_err(|e| Error::Backend(format!("{}", e)))?;
        Ok(DropDown { hwnd: inner.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void, inner: Rc::new(inner), _handler: Rc::new(handler) })
    }

    // ========== CheckButton ==========

    #[derive(Clone)]
    pub struct CheckButton {
        pub(crate) inner: Rc<nwg::CheckBox>,
        pub(crate) _handler: Rc<nwg::EventHandler>,
        // Stable hwnd snapshot for AsRef (see Window).
        hwnd: *mut c_void,
    }

    impl CheckButton {
        pub fn set_active(&self, active: bool) {
            self.inner.set_check_state(if active { nwg::CheckBoxState::Checked } else { nwg::CheckBoxState::Unchecked });
        }
        pub fn is_active(&self) -> bool {
            matches!(self.inner.check_state(), nwg::CheckBoxState::Checked)
        }
        pub fn set_label(&self, label: &str) { self.inner.set_text(label); }
        pub fn get_label(&self) -> Option<String> { Some(self.inner.text()) }
        pub fn connect_toggled(&self, _f: impl FnMut() + 'static) -> Result<u64, Error> {
            Ok(0)
        }
    }

    impl AsRef<*mut c_void> for CheckButton {
        fn as_ref(&self) -> &*mut c_void {
            &self.hwnd
        }
    }

    pub fn create_checkbutton(parent: *mut c_void) -> Result<CheckButton, Error> {
        let (inner, handler) = crate::backends::nwg::create_checkbox(parent, "Check")
            .map_err(|e| Error::Backend(format!("{}", e)))?;
        Ok(CheckButton { hwnd: inner.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void, inner: Rc::new(inner), _handler: Rc::new(handler) })
    }

    // ========== RadioButton ==========

    #[derive(Clone)]
    pub struct RadioButton {
        pub(crate) inner: Rc<nwg::RadioButton>,
        pub(crate) _handler: Rc<nwg::EventHandler>,
        // Stable hwnd snapshot for AsRef (see Window).
        hwnd: *mut c_void,
    }

    impl RadioButton {
        pub fn grab_focus(&self) {
            unsafe {
                winapi::um::winuser::SetFocus(self.hwnd as _);
            }
        }
        pub fn set_active(&self, active: bool) {
            self.inner.set_check_state(if active { nwg::RadioButtonState::Checked } else { nwg::RadioButtonState::Unchecked });
        }
        pub fn is_active(&self) -> bool {
            self.inner.check_state() == nwg::RadioButtonState::Checked
        }
        pub fn set_label(&self, label: &str) { self.inner.set_text(label); }
        pub fn get_label(&self) -> Option<String> { Some(self.inner.text()) }
        pub fn connect_toggled(&self, _f: impl FnMut() + 'static) -> Result<u64, Error> {
            Ok(0)
        }
        pub fn set_hexpand(&self, _expand: bool) {}
        pub fn set_vexpand(&self, _expand: bool) {}
    }

    impl AsRef<*mut c_void> for RadioButton {
        fn as_ref(&self) -> &*mut c_void {
            &self.hwnd
        }
    }

    /// `group_start` marks the first radio of a mutually-exclusive group
    /// (WS_GROUP): arrows then move selection natively within the group.
    /// Without it every radio is independent and arrows do nothing.
    pub fn create_radiobutton(parent: *mut c_void, group_start: bool) -> Result<RadioButton, Error> {
        let (inner, handler) = crate::backends::nwg::create_radiobutton(parent, "Radio", group_start)
            .map_err(|e| Error::Backend(format!("{}", e)))?;
        Ok(RadioButton { hwnd: inner.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void, inner: Rc::new(inner), _handler: Rc::new(handler) })
    }

    // ========== TextView ==========

    #[derive(Clone)]
    pub struct TextView {
        pub(crate) inner: Rc<nwg::TextBox>,
        pub(crate) changed_cb: Rc<RefCell<Option<Box<dyn FnMut()>>>>,
        pub(crate) _handler: Rc<nwg::EventHandler>,
        // Stable hwnd snapshot for AsRef (see Window).
        hwnd: *mut c_void,
    }

    impl TextView {
        pub fn get_buffer(&self) -> &RefCell<Option<Box<dyn FnMut()>>> { &self.changed_cb }
        pub fn set_text(&self, text: &str) { self.inner.set_text(text); }
        pub fn get_text(&self) -> Option<String> { Some(self.inner.text()) }
        pub fn set_editable(&self, editable: bool) { self.inner.set_readonly(!editable); }
        pub fn set_size_request(&self, _w: i32, _h: i32) {}
        pub fn set_wrap_mode(&self, _mode: i32) {}
    }

    impl AsRef<*mut c_void> for TextView {
        fn as_ref(&self) -> &*mut c_void {
            &self.hwnd
        }
    }

    pub fn create_textview(parent: *mut c_void) -> Result<TextView, Error> {
        let (inner, changed_cb, handler) = crate::backends::nwg::create_textview(parent)
            .map_err(|e| Error::Backend(format!("{}", e)))?;
        Ok(TextView { hwnd: inner.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void, inner: Rc::new(inner), changed_cb, _handler: Rc::new(handler) })
    }

    // ========== Dialog ==========

    pub struct Dialog {
        pub(crate) inner: Rc<nwg::Window>,
        // Stable hwnd snapshot for AsRef (see Window).
        hwnd: *mut c_void,
        pub(crate) buttons: Rc<RefCell<Vec<(nwg::Button, nwg::EventHandler)>>>,
        pub(crate) response_cb: Rc<RefCell<Option<Box<dyn FnMut(i32)>>>>,
        pub(crate) _handler: Rc<nwg::EventHandler>,
        layout_cb: Rc<RefCell<Vec<Box<dyn FnMut(i32, i32)>>>>,
        dlg_key_handlers: Rc<RefCell<Vec<nwg::RawEventHandler>>>,
        esc_bound: Rc<RefCell<bool>>,
        content: Rc<RefCell<Vec<*mut c_void>>>,
    }

    impl Clone for Dialog {
        fn clone(&self) -> Self {
            Dialog {
                inner: self.inner.clone(),
                buttons: self.buttons.clone(),
                response_cb: self.response_cb.clone(),
                _handler: self._handler.clone(),
                hwnd: self.hwnd,
                layout_cb: self.layout_cb.clone(),
                dlg_key_handlers: self.dlg_key_handlers.clone(),
                esc_bound: self.esc_bound.clone(),
                content: self.content.clone(),
            }
        }
    }

    impl Dialog {
        pub fn run(&self) -> i32 { 0 }
        pub fn set_title(&self, title: &str) { self.inner.set_text(title); }
        pub fn set_size_request(&self, _w: i32, _h: i32) {}
        pub fn set_default_size(&self, w: i32, h: i32) {
            let hwnd = self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut());
            if hwnd != std::ptr::null_mut() {
                unsafe {
                    winapi::um::winuser::SetWindowPos(
                        hwnd as _, std::ptr::null_mut(), 0, 0, w, h,
                        winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_NOMOVE,
                    );
                }
            }
        }
        pub fn present(&self) {
            self.inner.set_visible(true);
            let hwnd = self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut());
            if hwnd != std::ptr::null_mut() {
                // Parity with GtkDialog (which natively emits a cancel response
                // on Escape): dismiss the dialog on Escape and fire the response
                // callback with 0 (the Cancel convention used by corro's prompt
                // dialogs). Keyboard focus is usually on a child (button/entry),
                // so bind on the dialog and every descendant. Bound once; the
                // RawEventHandlers are kept alive in dlg_key_handlers. All call sites
                // append content before present(), so the subtree is complete.
                // Own once-flag (not vec-emptiness): other dialog key bindings
                // (e.g. set_default_response) share the vec.
                if !*self.esc_bound.borrow() {
                    *self.esc_bound.borrow_mut() = true;
                    self.bind_esc_dismiss(hwnd);
                }
                unsafe {
                    winapi::um::winuser::SetForegroundWindow(hwnd as _);
                    let mut rect: winapi::shared::windef::RECT = std::mem::zeroed();
                    winapi::um::winuser::GetClientRect(hwnd, &mut rect);
                    let w = rect.right - rect.left;
                    let h = rect.bottom - rect.top;
                    for cb in self.layout_cb.borrow_mut().iter_mut() {
                        cb(w, h);
                    }
                    // Content over buttons (GtkDialog parity); also runs on
                    // dialog WM_SIZE and after add_button().
                    self.layout_dialog();
                    winapi::um::winuser::RedrawWindow(
                        hwnd as _,
                        std::ptr::null_mut(),
                        std::ptr::null_mut(),
                        winapi::um::winuser::RDW_INVALIDATE | winapi::um::winuser::RDW_UPDATENOW | winapi::um::winuser::RDW_ALLCHILDREN | winapi::um::winuser::RDW_ERASE | winapi::um::winuser::RDW_FRAME,
                    );
                }
            }
        }
        pub fn append_content_area(&self, child: &impl Appendable) {
            for &ptr in &child.collect_hwnds() {
                if !ptr.is_null() {
                    unsafe {
                        let dlg = self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut());
                        winapi::um::winuser::SetParent(ptr as _, dlg as _);
                        winapi::um::winuser::ShowWindow(ptr as _, winapi::um::winuser::SW_SHOW);
                    }
                    // Recorded for layout_dialog(): content stacks vertically
                    // over the button row (GtkDialog anatomy). Positioning
                    // happens there — never a full-area stretch per child,
                    // which overlapped every child at (0,0).
                    self.content.borrow_mut().push(ptr);
                }
            }
        }
        /// Lay out content (stacked, full-width, above the buttons) and the
        /// button row (bottom-right), GtkDialog parity. Runs on present(),
        /// on dialog WM_SIZE, and after add_button().
        pub fn layout_dialog(&self) {
            let hwnd = self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut());
            if hwnd == std::ptr::null_mut() {
                return;
            }
            layout_nwg_dialog_parts(hwnd as _, &self.content, &self.buttons);
        }
        pub fn add_button(&self, text: &str, response_id: i32) {
            let hwnd = self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut());
            let cb = self.response_cb.clone();
            let result = crate::backends::nwg::create_dialog_button(hwnd as *mut c_void, text, response_id, cb);
            if let Ok(btn) = result {
                self.buttons.borrow_mut().push(btn);
                // Buttons were never positioned (all piled at 0,0); lay out
                // now in case the dialog is already visible.
                self.layout_dialog();
            }
        }
        pub fn connect_response<F: FnMut(i32) + 'static>(&self, f: F) -> Result<u64, Error> {
            *self.response_cb.borrow_mut() = Some(Box::new(f));
            Ok(0)
        }
        /// Make `response_id` the dialog's default response: Return anywhere
        /// unhandled (e.g. on a focused radio button, which consumes nothing)
        /// fires it and dismisses, mirroring GtkDialog default-response
        /// semantics. Handlers bind on the dialog and all current
        /// descendants (like bind_esc_dismiss) because keys go to the focused
        /// control. Children that consume Return themselves (entries via
        /// activate) swallow it first, so no double-confirm.
        ///

        pub fn set_default_response(&self, response_id: i32) {
            let dlg_hwnd = self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut());
            if dlg_hwnd.is_null() {
                return;
            }
            fn collect(hwnd: winapi::shared::windef::HWND, out: &mut Vec<winapi::shared::windef::HWND>) {
                out.push(hwnd);
                unsafe {
                    let mut child = winapi::um::winuser::GetWindow(hwnd, winapi::um::winuser::GW_CHILD);
                    while child != std::ptr::null_mut() {
                        collect(child, out);
                        child = winapi::um::winuser::GetWindow(child, winapi::um::winuser::GW_HWNDNEXT);
                    }
                }
            }
            let mut hwnds = Vec::new();
            collect(dlg_hwnd as _, &mut hwnds);
            static DEFRESP_ID: AtomicUsize = AtomicUsize::new(0xD1000000);
            for child in hwnds {
                let cb = self.response_cb.clone();
                let id = DEFRESP_ID.fetch_add(1, Ordering::SeqCst);
                if let Ok(h) = nwg::bind_raw_event_handler(
                    &nwg::ControlHandle::Hwnd(child), id,
                    move |_h, msg, w, _l| {
                        if msg == winapi::um::winuser::WM_KEYDOWN && (w & 0xFF) == 0x0D {
                            if let Some(ref mut f) = *cb.borrow_mut() { f(response_id); }
                            // Dismiss like the button path: confirming must close.
                            unsafe {
                                winapi::um::winuser::ShowWindow(dlg_hwnd as _, winapi::um::winuser::SW_HIDE);
                            }
                            return Some(0);
                        }
                        None
                    },
                ) {
                    self.dlg_key_handlers.borrow_mut().push(h);
                }
            }
        }
        /// Bind an Escape-to-dismiss raw handler on the dialog window and all
        /// of its current descendants (see present()).
        fn bind_esc_dismiss(&self, dlg_hwnd: winapi::shared::windef::HWND) {
            // Collect dialog + recursive descendants.
            fn collect(hwnd: winapi::shared::windef::HWND, out: &mut Vec<winapi::shared::windef::HWND>) {
                out.push(hwnd);
                unsafe {
                    let mut child = winapi::um::winuser::GetWindow(hwnd, winapi::um::winuser::GW_CHILD);
                    while child != std::ptr::null_mut() {
                        collect(child, out);
                        child = winapi::um::winuser::GetWindow(child, winapi::um::winuser::GW_HWNDNEXT);
                    }
                }
            }
            let mut hwnds = Vec::new();
            collect(dlg_hwnd as _, &mut hwnds);
            static ESC_ID: AtomicUsize = AtomicUsize::new(0xD0000000);
            for child in hwnds {
                let cb = self.response_cb.clone();
                let id = ESC_ID.fetch_add(1, Ordering::SeqCst);
                if let Ok(h) = nwg::bind_raw_event_handler(
                    &nwg::ControlHandle::Hwnd(child), id,
                    move |_h, msg, w, _l| {
                        if (msg == winapi::um::winuser::WM_KEYDOWN
                            || msg == winapi::um::winuser::WM_SYSKEYDOWN)
                            && w == winapi::um::winuser::VK_ESCAPE as usize
                        {
                            if let Some(ref mut f) = *cb.borrow_mut() { f(0); }
                            unsafe {
                                winapi::um::winuser::ShowWindow(dlg_hwnd as _,
                                    winapi::um::winuser::SW_HIDE);
                            }
                            return Some(0);
                        }
                        None
                    },
                ) {
                    self.dlg_key_handlers.borrow_mut().push(h);
                }
            }
        }
        pub fn close(&self) {
            if let Some(hwnd) = self.inner.handle.hwnd() {
                unsafe {
                    winapi::um::winuser::ShowWindow(hwnd as _, winapi::um::winuser::SW_HIDE);
                }
            }
        }
    }

    impl AsRef<*mut c_void> for Dialog {
        fn as_ref(&self) -> &*mut c_void {
            &self.hwnd
        }
    }

    impl Widget for Dialog {
        fn raw_handle(&self) -> *mut c_void {
            self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void
        }
    }

    impl Dialog {
        pub fn set_visible(&self, v: bool) {
            unsafe {
                winapi::um::winuser::ShowWindow(self.inner.handle.hwnd().unwrap_or(std::ptr::null_mut()) as _, if v { winapi::um::winuser::SW_SHOW } else { winapi::um::winuser::SW_HIDE });
            }
        }
    }

    pub fn create_dialog(
        parent_cell: &Rc<RefCell<Option<*mut c_void>>>,
    ) -> Result<Dialog, Error> {
        let response_cb: Rc<RefCell<Option<Box<dyn FnMut(i32)>>>> = Rc::new(RefCell::new(None));
        let btn_cb = response_cb.clone();
        let (inner, _, handler) = crate::backends::nwg::create_dialog(parent_cell, btn_cb)
            .map_err(|e| Error::Backend(format!("{}", e)))?;

        let layout_cb: Rc<RefCell<Vec<Box<dyn FnMut(i32, i32)>>>> = Rc::new(RefCell::new(Vec::new()));

        let content: Rc<RefCell<Vec<*mut c_void>>> = Rc::new(RefCell::new(Vec::new()));
        let buttons: Rc<RefCell<Vec<(nwg::Button, nwg::EventHandler)>>> =
            Rc::new(RefCell::new(Vec::new()));
        // Bind raw WM_SIZE handler for dialog
        let dlg_hwnd = inner.handle.hwnd().unwrap_or(std::ptr::null_mut());
        if dlg_hwnd != std::ptr::null_mut() {
            let cb = layout_cb.clone();
            let content_wm = content.clone();
            let buttons_wm = buttons.clone();
            static DIALOG_RAW_HANDLER_ID: AtomicUsize = AtomicUsize::new(0x20000000);
            let handler_id = DIALOG_RAW_HANDLER_ID.fetch_add(1, Ordering::SeqCst);
            nwg::bind_raw_event_handler(
                &nwg::ControlHandle::Hwnd(dlg_hwnd),
                handler_id,
                move |_h, msg, _w, l| {
                    if msg == winapi::um::winuser::WM_SIZE {
                        let w = (l & 0xFFFF) as i32;
                        let h = ((l >> 16) & 0xFFFF) as i32;
                        for cb_item in cb.borrow_mut().iter_mut() {
                            cb_item(w, h);
                        }
                        layout_nwg_dialog_parts(dlg_hwnd, &content_wm, &buttons_wm);
                    }
                    None
                },
            ).map_err(|e| Error::Backend(format!("{}", e)))?;
        }

        Ok(Dialog { hwnd: inner.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void, inner: Rc::new(inner), buttons, response_cb, _handler: Rc::new(handler), layout_cb, dlg_key_handlers: Rc::new(RefCell::new(Vec::new())), esc_bound: Rc::new(RefCell::new(false)), content })
    }

    pub fn create_dialog_button(
        parent: *mut c_void,
        text: &str,
        response_id: i32,
        cb: Rc<RefCell<Option<Box<dyn FnMut(i32)>>>>,
    ) -> Result<(nwg::Button, nwg::EventHandler), Error> {
        crate::backends::nwg::create_dialog_button(parent, text, response_id, cb)
            .map_err(|e| Error::Backend(format!("{}", e)))
    }

    // ========== GDI DrawContext ==========

    pub struct NwgDrawContext {
        hdc: winapi::shared::windef::HDC,
        w: i32,
        h: i32,
    }

    impl NwgDrawContext {
        fn make_font(name: &str, size: f64, weight: i32, italic: bool) -> winapi::shared::windef::HFONT {
            unsafe {
                let is_mono = name.eq_ignore_ascii_case("monospace");
                let face = if is_mono { "Courier New" } else { name };
                let wide_name: Vec<u16> = face.encode_utf16().chain(std::iter::once(0)).collect();
                let pitch = if is_mono { winapi::um::wingdi::FF_MODERN } else { 0 };
                winapi::um::wingdi::CreateFontW(
                    -(size.abs() as i32), 0, 0, 0, weight as i32, italic as u32, 0, 0,
                    winapi::um::wingdi::ANSI_CHARSET,
                    winapi::um::wingdi::OUT_DEFAULT_PRECIS,
                    winapi::um::wingdi::CLIP_DEFAULT_PRECIS,
                    winapi::um::wingdi::PROOF_QUALITY,
                    winapi::um::wingdi::DEFAULT_PITCH | pitch,
                    wide_name.as_ptr(),
                )
            }
        }
    }

    impl DrawContext for NwgDrawContext {
        fn fill_rect(&mut self, x: f64, y: f64, w: f64, h: f64, r: f64, g: f64, b: f64, _a: f64) {
            unsafe {
                let color: u32 = winapi::um::wingdi::RGB(
                    (r * 255.0) as u8, (g * 255.0) as u8, (b * 255.0) as u8);
                let brush = winapi::um::wingdi::CreateSolidBrush(color);
                if !brush.is_null() {
                    let mut rect = winapi::shared::windef::RECT {
                        left: x as i32, top: y as i32,
                        right: (x + w) as i32, bottom: (y + h) as i32,
                    };
                    winapi::um::winuser::FillRect(self.hdc, &mut rect, brush);
                    winapi::um::wingdi::DeleteObject(brush as _);
                }
            }
        }
        fn stroke_rect(&mut self, x: f64, y: f64, w: f64, h: f64, r: f64, g: f64, b: f64, _a: f64, lw: f64) {
            unsafe {
                let color: u32 = winapi::um::wingdi::RGB(
                    (r * 255.0) as u8, (g * 255.0) as u8, (b * 255.0) as u8);
                let pen = winapi::um::wingdi::CreatePen(winapi::um::wingdi::PS_SOLID as i32, lw as i32, color);
                if !pen.is_null() {
                    let old_pen = winapi::um::wingdi::SelectObject(self.hdc, pen as _);
                    if w == 0.0 && h != 0.0 {
                        winapi::um::wingdi::MoveToEx(self.hdc, x as i32, y as i32, std::ptr::null_mut());
                        winapi::um::wingdi::LineTo(self.hdc, x as i32, (y + h) as i32);
                    } else if h == 0.0 && w != 0.0 {
                        winapi::um::wingdi::MoveToEx(self.hdc, x as i32, y as i32, std::ptr::null_mut());
                        winapi::um::wingdi::LineTo(self.hdc, (x + w) as i32, y as i32);
                    } else {
                        let null_brush = winapi::um::wingdi::GetStockObject(winapi::um::wingdi::NULL_BRUSH as i32);
                        let old_brush = winapi::um::wingdi::SelectObject(self.hdc, null_brush);
                        winapi::um::wingdi::Rectangle(self.hdc, x as i32, y as i32, (x + w) as i32, (y + h) as i32);
                        winapi::um::wingdi::SelectObject(self.hdc, old_brush);
                    }
                    winapi::um::wingdi::SelectObject(self.hdc, old_pen);
                    winapi::um::wingdi::DeleteObject(pen as _);
                }
            }
        }
        fn draw_text_styled(&mut self, x: f64, y: f64, text: &str, font: &str, size: f64,
                            r: f64, g: f64, b: f64, _a: f64, _slant: i32, weight: i32) {
            unsafe {
                let wide: Vec<u16> = text.encode_utf16().collect();
                let color: u32 = winapi::um::wingdi::RGB(
                    (r * 255.0) as u8, (g * 255.0) as u8, (b * 255.0) as u8);
                winapi::um::wingdi::SetTextColor(self.hdc, color);
                let old_bkmode = winapi::um::wingdi::SetBkMode(self.hdc, winapi::um::wingdi::TRANSPARENT as i32);
                let hfont = Self::make_font(font, size, if weight != 0 { winapi::um::wingdi::FW_BOLD as i32 } else { winapi::um::wingdi::FW_NORMAL as i32 }, _slant != 0);
                if hfont.is_null() {
                    let sys = winapi::um::wingdi::GetStockObject(winapi::um::wingdi::SYSTEM_FONT as i32) as winapi::shared::windef::HFONT;
                    let old = winapi::um::wingdi::SelectObject(self.hdc, sys as _);
                    winapi::um::wingdi::TextOutW(self.hdc, x as i32, y as i32, wide.as_ptr() as _, wide.len() as i32);
                    winapi::um::wingdi::SelectObject(self.hdc, old);
                } else {
                    let old_font = winapi::um::wingdi::SelectObject(self.hdc, hfont as _);
                    winapi::um::wingdi::TextOutW(self.hdc, x as i32, y as i32, wide.as_ptr() as _, wide.len() as i32);
                    winapi::um::wingdi::SelectObject(self.hdc, old_font);
                    winapi::um::wingdi::DeleteObject(hfont as _);
                }
                winapi::um::wingdi::SetBkMode(self.hdc, old_bkmode);
            }
        }
        fn text_extents_styled(&self, text: &str, font: &str, size: f64, _slant: i32, weight: i32) -> (f64, f64, f64, f64) {
            unsafe {
                let wide: Vec<u16> = text.encode_utf16().collect();
                let hfont = Self::make_font(font, size, if weight != 0 { winapi::um::wingdi::FW_BOLD as i32 } else { winapi::um::wingdi::FW_NORMAL as i32 }, _slant != 0);
                if hfont.is_null() { return (0.0, 0.0, 0.0, 0.0); }
                let old_font = winapi::um::wingdi::SelectObject(self.hdc, hfont as _);
                let mut size_tag: winapi::shared::windef::SIZE = std::mem::zeroed();
                winapi::um::wingdi::GetTextExtentPoint32W(self.hdc, wide.as_ptr() as _, wide.len() as i32, &mut size_tag);
                winapi::um::wingdi::SelectObject(self.hdc, old_font);
                winapi::um::wingdi::DeleteObject(hfont as _);
                (0.0, 0.0, size_tag.cx as f64, size_tag.cy as f64)
            }
        }
        fn clear(&mut self, r: f64, g: f64, b: f64, _a: f64) {
            self.fill_rect(0.0, 0.0, self.w as f64, self.h as f64, r, g, b, 1.0);
        }
        fn save(&mut self) {
            unsafe { winapi::um::wingdi::SaveDC(self.hdc); }
        }
        fn restore(&mut self) {
            unsafe { winapi::um::wingdi::RestoreDC(self.hdc, -1); }
        }
        fn clip(&mut self, x: f64, y: f64, w: f64, h: f64) {
            unsafe { winapi::um::wingdi::IntersectClipRect(self.hdc, x as i32, y as i32, (x + w) as i32, (y + h) as i32); }
        }
    }

    // ========== Canvas ==========

    pub struct Canvas {
        frame: Option<Rc<nwg::Frame>>,
        hwnd: *mut c_void,
        draw_cb: Rc<RefCell<Option<Box<dyn FnMut(&mut dyn crate::core::DrawContext, i32, i32)>>>>,
        click_cb: Rc<RefCell<Option<Box<dyn FnMut(f64, f64)>>>>,
        key_cb: Rc<RefCell<Option<Box<dyn FnMut(u32) -> bool>>>>,
        _raw_handlers: Rc<Vec<nwg::RawEventHandler>>,
        painting: Rc<RefCell<bool>>,
    }

    impl Canvas {
        pub fn set_draw_callback(&self, cb: Box<dyn FnMut(&mut dyn DrawContext, i32, i32)>) {
            *self.draw_cb.borrow_mut() = Some(cb);
            self.queue_redraw();
        }
        pub fn queue_redraw(&self) {
            if !self.hwnd.is_null() {
                unsafe {
                    winapi::um::winuser::InvalidateRect(self.hwnd as _, std::ptr::null_mut(), 0);
                }
            }
        }
        pub fn set_size_request(&self, w: i32, h: i32) {
            if !self.hwnd.is_null() {
                unsafe {
                    winapi::um::winuser::SetWindowPos(
                        self.hwnd as _,
                        std::ptr::null_mut(), 0, 0, w, h,
                        winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_NOMOVE | winapi::um::winuser::SWP_SHOWWINDOW,
                    );
                }
            }
        }

        pub fn set_visible(&self, visible: bool) {
            if !self.hwnd.is_null() {
                unsafe {
                    winapi::um::winuser::ShowWindow(
                        self.hwnd as _,
                        if visible { winapi::um::winuser::SW_SHOW } else { winapi::um::winuser::SW_HIDE },
                    );
                }
            }
        }
        pub fn set_content_size(&self, _w: i32, _h: i32) {}
        pub fn on_click(&self, cb: Box<dyn FnMut(f64, f64)>) {
            *self.click_cb.borrow_mut() = Some(cb);
        }
        pub fn on_key(&self, cb: Box<dyn FnMut(u32) -> bool>) {
            *self.key_cb.borrow_mut() = Some(cb);
        }
        pub fn on_key_raw(&self, cb: Box<dyn FnMut(u32, u32) -> bool>) {
            let mut cb = cb;
            *self.key_cb.borrow_mut() = Some(Box::new(move |k: u32| -> bool { cb(k, 0) }));
        }
        pub fn grab_focus(&self) {
            unsafe { winapi::um::winuser::SetFocus(self.hwnd as _); }
        }
        pub fn set_can_focus(&self, _can: bool) {}
        pub fn force_draw(&self, _window_ptr: *mut c_void, _fallback_w: i32, _fallback_h: i32) {}
    }

    impl Clone for Canvas {
        fn clone(&self) -> Self {
            Canvas {
                frame: self.frame.clone(),
                hwnd: self.hwnd,
                draw_cb: self.draw_cb.clone(),
                click_cb: self.click_cb.clone(),
                key_cb: self.key_cb.clone(),
                _raw_handlers: Rc::new(Vec::new()),
                painting: self.painting.clone(),
            }
        }
    }

    impl AsRef<*mut c_void> for Canvas {
        fn as_ref(&self) -> &*mut c_void {
            &self.hwnd
        }
    }

    impl Widget for Canvas {
        fn raw_handle(&self) -> *mut c_void { self.hwnd }
    }

    pub fn create_canvas(parent: *mut c_void) -> Result<Canvas, Error> {
        let mut frame = nwg::Frame::default();
        if !parent.is_null() {
            nwg::Frame::builder()
                .flags(nwg::FrameFlags::VISIBLE)
                .size((0, 0))
                .position((0, 0))
                .parent(&nwg::ControlHandle::Hwnd(parent as _))
                .build(&mut frame)
                .map_err(|e| Error::Backend(format!("{}", e)))?;
        }
        let hwnd = frame.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void;
        let draw_cb: Rc<RefCell<Option<Box<dyn FnMut(&mut dyn DrawContext, i32, i32)>>>> = Rc::new(RefCell::new(None));
        let painting: Rc<RefCell<bool>> = Rc::new(RefCell::new(false));
        let click_cb: Rc<RefCell<Option<Box<dyn FnMut(f64, f64)>>>> = Rc::new(RefCell::new(None));
        let key_cb: Rc<RefCell<Option<Box<dyn FnMut(u32) -> bool>>>> = Rc::new(RefCell::new(None));

        let mut handlers: Vec<nwg::RawEventHandler> = Vec::new();

        if hwnd != std::ptr::null_mut() {
            let raw_hwnd: winapi::shared::windef::HWND = hwnd as _;

            // WM_KEYDOWN/WM_SYSKEYDOWN handler for keyboard input.
            // Without this, the Canvas's key_cb is never called because
            // no raw handler is registered to process keystroke messages.
            {
                let kc = key_cb.clone();
                static KEYBOARD_ID: AtomicUsize = AtomicUsize::new(0x70000000);
                let kid = KEYBOARD_ID.fetch_add(1, Ordering::SeqCst);
                if let Some(h) = nwg::bind_raw_event_handler(
                    &nwg::ControlHandle::Hwnd(raw_hwnd), kid,
                    move |_h, msg, w, _l| {
                        if msg == winapi::um::winuser::WM_KEYDOWN || msg == winapi::um::winuser::WM_SYSKEYDOWN {
                            // Same pure-Alt yield as the entry handler: Alt+letter
                            // must reach DefWindowProc (native menu), never the
                            // grid edit path. Ctrl+Alt (AltGr) passes through.
                            if pure_alt_held() {
                                return None;
                            }
                            // Pure Ctrl is accelerators, not grid input: decline
                            // instead of typing the letter (Ctrl+C must never
                            // insert 'c'). Ctrl+Alt (AltGr) passes through.
                            if pure_ctrl_held() {
                                return None;
                            }
                            if let Some(ref mut f) = *kc.borrow_mut() {
                                // Translate Shift pairs/capitals like the
                                // entry path; specials keep raw VKs. A
                                // colliding translation (None) declines.
                                if let Some(k) = translate_vk(w) {
                                    if f(k) { return Some(0); }
                                } else {
                                    return None;
                                }
                            }
                        }
                        None
                    },
                ).ok() { handlers.push(h); }
            }

            // Suppress WM_ERASEBKGND (prevent flash from class background brush)
            {
                static ERASE_ID: AtomicUsize = AtomicUsize::new(0x60000000);
                let eid = ERASE_ID.fetch_add(1, Ordering::SeqCst);
                if let Some(h) = nwg::bind_raw_event_handler(
                    &nwg::ControlHandle::Hwnd(raw_hwnd), eid,
                    move |_h, msg, _w, _l| {
                        if msg == winapi::um::winuser::WM_ERASEBKGND { Some(1) } else { None }
                    },
                ).ok() { handlers.push(h); }
            }

            // WM_PAINT handler
            {
                let cb = draw_cb.clone();
                let paint_flag = painting.clone();
                static CANVAS_PAINT_ID: AtomicUsize = AtomicUsize::new(0x30000000);
                let pid = CANVAS_PAINT_ID.fetch_add(1, Ordering::SeqCst);
                if let Some(h) = nwg::bind_raw_event_handler(
                    &nwg::ControlHandle::Hwnd(raw_hwnd), pid,
                    move |_h, msg, _w, _l| {
                        if msg != winapi::um::winuser::WM_PAINT { return None; }
                        if *paint_flag.borrow() { return Some(0); }
                        *paint_flag.borrow_mut() = true;
                        unsafe {
                            let mut ps: winapi::um::winuser::PAINTSTRUCT = std::mem::zeroed();
                            let hdc = winapi::um::winuser::BeginPaint(hwnd as _, &mut ps);
                            let mut rect: winapi::shared::windef::RECT = std::mem::zeroed();
                            winapi::um::winuser::GetClientRect(hwnd as _, &mut rect);
                            let w = rect.right;
                            let h = rect.bottom;
                            if w > 0 && h > 0 {
                                let mem_dc = winapi::um::wingdi::CreateCompatibleDC(hdc);
                                if !mem_dc.is_null() {
                                    let bmp = winapi::um::wingdi::CreateCompatibleBitmap(hdc, w, h);
                                    if !bmp.is_null() {
                                        let old = winapi::um::wingdi::SelectObject(mem_dc, bmp as _);
                                        if let Some(ref mut draw_fn) = *cb.borrow_mut() {
                                            let mut ctx = NwgDrawContext { hdc: mem_dc, w, h };
                                            draw_fn(&mut ctx, w, h);
                                        }
                                        winapi::um::wingdi::BitBlt(hdc, 0, 0, w, h, mem_dc, 0, 0, winapi::um::wingdi::SRCCOPY);
                                        winapi::um::wingdi::SelectObject(mem_dc, old);
                                        winapi::um::wingdi::DeleteObject(bmp as _);
                                    }
                                    winapi::um::wingdi::DeleteDC(mem_dc);
                                }
                            }
                            winapi::um::winuser::EndPaint(hwnd as _, &mut ps);
                        }
                        *paint_flag.borrow_mut() = false;
                        Some(0)
                    },
                ).ok() { handlers.push(h); }
            }

            // WM_LBUTTONDOWN handler
            {
                let cc = click_cb.clone();
                static CLICK_ID: AtomicUsize = AtomicUsize::new(0x40000000);
                let cid = CLICK_ID.fetch_add(1, Ordering::SeqCst);
                if let Some(h) = nwg::bind_raw_event_handler(
                    &nwg::ControlHandle::Hwnd(raw_hwnd), cid,
                    move |_h, msg, _w, l| {
                        if msg != winapi::um::winuser::WM_LBUTTONDOWN { return None; }
                        {
                            let x = (l & 0xFFFF) as i16 as f64;
                            let y = ((l >> 16) & 0xFFFF) as i16 as f64;
                            if let Some(ref mut f) = *cc.borrow_mut() {
                                f(x, y);
                            }
                        }
                        Some(0)
                    },
                ).ok() { handlers.push(h); }
            }

            // WM_KEYDOWN handler — need focus first; forward WM_SETFOCUS to force keyboard input
            {
                let kc = key_cb.clone();
                static KEY_ID: AtomicUsize = AtomicUsize::new(0x50000000);
                let kid = KEY_ID.fetch_add(1, Ordering::SeqCst);
                if let Some(h) = nwg::bind_raw_event_handler(
                    &nwg::ControlHandle::Hwnd(raw_hwnd), kid,
                    move |_h, msg, w, _l| {
                        if msg != winapi::um::winuser::WM_KEYDOWN && msg != winapi::um::winuser::WM_SYSKEYDOWN { return None; }
                        // Pure Alt/Ctrl yield (menus/accelerators own those
                        // keys); Ctrl+Alt (AltGr) passes through.
                        if pure_alt_held() || pure_ctrl_held() {
                            return None;
                        }
                        if let Some(ref mut f) = *kc.borrow_mut() {
                            // Translate Shift pairs/capitals like the entry
                            // path; specials keep raw VKs. A colliding
                            // translation (None) declines.
                            if let Some(k) = translate_vk(w) {
                                if f(k) { return Some(0); }
                            } else {
                                return None;
                            }
                        }
                        None
                    },
                ).ok() { handlers.push(h); }
            }
        }

        Ok(Canvas {
            frame: Some(Rc::new(frame)),
            hwnd,
            draw_cb,
            click_cb,
            key_cb,
            _raw_handlers: Rc::new(handlers),
            painting,
        })
    }

    // ========== Overlay ==========

    pub struct Overlay {
        frame: Option<Rc<nwg::Frame>>,
        hwnd: *mut c_void,
    }

    impl Overlay {
        pub fn set_child(&self, child: &impl AsRef<*mut c_void>) {
            let ptr = *child.as_ref();
            if !ptr.is_null() && !self.hwnd.is_null() {
                unsafe {
                    winapi::um::winuser::SetParent(ptr as _, self.hwnd as _);
                }
            }
        }
        pub fn add_overlay(&self, child: &impl AsRef<*mut c_void>) {
            let ptr = *child.as_ref();
            if !ptr.is_null() && !self.hwnd.is_null() {
                unsafe {
                    winapi::um::winuser::SetParent(ptr as _, self.hwnd as _);
                    winapi::um::winuser::SetWindowPos(
                        ptr as _, winapi::um::winuser::HWND_TOP,
                        0, 0, 0, 0,
                        winapi::um::winuser::SWP_NOSIZE | winapi::um::winuser::SWP_SHOWWINDOW,
                    );
                }
            }
        }
        pub fn set_overlay_pass_through(&self, _child: &impl AsRef<*mut c_void>, _pass: bool) {}
        pub fn remove(&self, _child: &impl AsRef<*mut c_void>) {}
        pub fn show_all(&self) {
            if !self.hwnd.is_null() {
                unsafe {
                    winapi::um::winuser::ShowWindow(self.hwnd as _, winapi::um::winuser::SW_SHOW);
                }
            }
        }
        pub fn set_vexpand(&self, _expand: bool) {}
        pub fn set_hexpand(&self, _expand: bool) {}
        pub fn set_size_request(&self, w: i32, h: i32) {
            if !self.hwnd.is_null() {
                unsafe {
                    winapi::um::winuser::SetWindowPos(
                        self.hwnd as _,
                        std::ptr::null_mut(), 0, 0, w, h,
                        winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_NOMOVE | winapi::um::winuser::SWP_SHOWWINDOW,
                    );
                }
            }
        }
    }

    impl Clone for Overlay {
        fn clone(&self) -> Self {
            Overlay { frame: self.frame.clone(), hwnd: self.hwnd }
        }
    }

    impl AsRef<*mut c_void> for Overlay {
        fn as_ref(&self) -> &*mut c_void {
            &self.hwnd
        }
    }

    impl Widget for Overlay {
        fn raw_handle(&self) -> *mut c_void { self.hwnd }
    }

    pub fn create_overlay(parent: *mut c_void) -> Result<Overlay, Error> {
        let mut frame = nwg::Frame::default();
        if !parent.is_null() {
            nwg::Frame::builder()
                .flags(nwg::FrameFlags::VISIBLE)
                .size((0, 0))
                .position((0, 0))
                .parent(&nwg::ControlHandle::Hwnd(parent as _))
                .build(&mut frame)
                .map_err(|e| Error::Backend(format!("{}", e)))?;
        }
        let hwnd = frame.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void;
        Ok(Overlay { frame: Some(Rc::new(frame)), hwnd })
    }

    // ========== ScrolledWindow ==========

    pub struct ScrolledWindow {
        frame: Option<Rc<nwg::Frame>>,
        hwnd: *mut c_void,
        child: Rc<RefCell<Option<*mut c_void>>>,
        vscroll: Rc<RefCell<nwg::ScrollBar>>,
        hscroll: Rc<RefCell<nwg::ScrollBar>>,
        _handlers: Rc<Vec<nwg::RawEventHandler>>,
        on_scroll_cb: Rc<RefCell<Option<Box<dyn FnMut(bool, f64)>>>>,
    }

    impl Clone for ScrolledWindow {
        fn clone(&self) -> Self {
            ScrolledWindow {
                frame: self.frame.clone(),
                hwnd: self.hwnd,
                child: self.child.clone(),
                vscroll: self.vscroll.clone(),
                hscroll: self.hscroll.clone(),
                _handlers: self._handlers.clone(),
                on_scroll_cb: self.on_scroll_cb.clone(),
            }
        }
    }

    impl ScrolledWindow {
        pub fn set_child(&self, child: &impl AsRef<*mut c_void>) {
            let ptr = *child.as_ref();
            if ptr.is_null() || self.hwnd.is_null() { return; }
            unsafe {
                winapi::um::winuser::SetParent(ptr as _, self.hwnd as _);
                winapi::um::winuser::SetWindowPos(
                    ptr as _, std::ptr::null_mut(),
                    0, 0, 0, 0,
                    winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_NOSIZE | winapi::um::winuser::SWP_SHOWWINDOW,
                );
            }
            *self.child.borrow_mut() = Some(ptr);
            // Size the child to the frame's client area on the next WM_SIZE.
            // (Scroll ranges are driven by the host via scroll_to; the child
            // always fills the viewport and nothing pans.)
            unsafe {
                let mut rect: winapi::shared::windef::RECT = std::mem::zeroed();
                winapi::um::winuser::GetClientRect(self.hwnd as _, &mut rect);
                let fw = rect.right - rect.left;
                let fh = rect.bottom - rect.top;
                if fw > 0 && fh > 0 {
                    winapi::um::winuser::SetWindowPos(
                        ptr as _, std::ptr::null_mut(),
                        0, 0, (fw - 20).max(0), (fh - 20).max(0),
                        winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_SHOWWINDOW,
                    );
                }
            }
        }

        pub fn set_policy(&self, _h: u32, _v: u32) {
            // NWG: always show both scrollbars
        }

        /// Drive both scrollbars from item indices (shared cursor-centric
        /// model with the GTK backend): value = cursor index, upper = domain
        /// size, page = visible count (page sizes the thumb where supported).
        pub fn scroll_to(&self, hval: f64, hupper: f64, _hpage: f64, vval: f64, vupper: f64, _vpage: f64) {
            if let Ok(sb) = self.hscroll.try_borrow() {
                sb.set_range(0..(hupper.max(1.0) as usize));
                sb.set_pos((hval.max(0.0) as usize).min(hupper.max(1.0) as usize));
            }
            if let Ok(sb) = self.vscroll.try_borrow() {
                sb.set_range(0..(vupper.max(1.0) as usize));
                sb.set_pos((vval.max(0.0) as usize).min(vupper.max(1.0) as usize));
            }
        }

        /// Notify on user scrollbar interaction: `cb(vertical, pos)`.
        /// Fires for thumb drags and trough clicks. The host moves its cursor
        /// (the canvas keeps filling the viewport; nothing pans).
        pub fn on_scroll(&self, cb: Box<dyn FnMut(bool, f64)>) {
            *self.on_scroll_cb.borrow_mut() = Some(cb);
        }

        pub fn set_vexpand(&self, _v: bool) {}
        pub fn set_hexpand(&self, _h: bool) {}

    }

    impl AsRef<*mut c_void> for ScrolledWindow {
        fn as_ref(&self) -> &*mut c_void {
            &self.hwnd
        }
    }

    pub fn create_scrolled_window(parent: *mut c_void) -> Result<ScrolledWindow, Error> {
        let mut frame = nwg::Frame::default();
        if !parent.is_null() {
            nwg::Frame::builder()
                .flags(nwg::FrameFlags::VISIBLE)
                .size((0, 0))
                .position((0, 0))
                .parent(&nwg::ControlHandle::Hwnd(parent as _))
                .build(&mut frame)
                .map_err(|e| Error::Backend(format!("{}", e)))?;
        }
        let hwnd = frame.handle.hwnd().unwrap_or(std::ptr::null_mut()) as *mut c_void;

        let mut vscroll = nwg::ScrollBar::default();
        nwg::ScrollBar::builder()
            .flags(nwg::ScrollBarFlags::VISIBLE | nwg::ScrollBarFlags::VERTICAL)
            .parent(&nwg::ControlHandle::Hwnd(hwnd as _))
            .range(Some(0..1))
            .pos(Some(0))
            .build(&mut vscroll)
            .map_err(|e| Error::Backend(format!("{}", e)))?;

        let mut hscroll = nwg::ScrollBar::default();
        nwg::ScrollBar::builder()
            .flags(nwg::ScrollBarFlags::VISIBLE | nwg::ScrollBarFlags::HORIZONTAL)
            .parent(&nwg::ControlHandle::Hwnd(hwnd as _))
            .range(Some(0..1))
            .pos(Some(0))
            .build(&mut hscroll)
            .map_err(|e| Error::Backend(format!("{}", e)))?;

        let vscroll = Rc::new(RefCell::new(vscroll));
        let hscroll = Rc::new(RefCell::new(hscroll));
        let child: Rc<RefCell<Option<*mut c_void>>> = Rc::new(RefCell::new(None));
        let on_scroll_cb: Rc<RefCell<Option<Box<dyn FnMut(bool, f64)>>>> =
            Rc::new(RefCell::new(None));

        let mut handlers: Vec<nwg::RawEventHandler> = Vec::new();

        // WM_SIZE on frame to reposition scrollbars and update range
        if hwnd != std::ptr::null_mut() {
            let vscroll_sz = vscroll.clone();
            let hscroll_sz = hscroll.clone();
            let child_sz_child = child.clone();
            let c_hwnd = hwnd;
            static SIZE_ID: AtomicUsize = AtomicUsize::new(0x80000000);
            let sid = SIZE_ID.fetch_add(1, Ordering::SeqCst);
            if let Some(h) = nwg::bind_raw_event_handler(
                &nwg::ControlHandle::Hwnd(hwnd as _), sid,
                move |_h, msg, _w, _l| {
                    if msg != winapi::um::winuser::WM_SIZE { return None; }
                    unsafe {
                        let mut rect: winapi::shared::windef::RECT = std::mem::zeroed();
                        winapi::um::winuser::GetClientRect(c_hwnd as _, &mut rect);
                        let w = rect.right;
                        let h = rect.bottom;
                        // Drop zero sizes (stale setup-storm leftovers):
                        // resizing the canvas to 0x0 would silence its
                        // WM_PAINT forever (empty update region, nothing
                        // re-invalidates). The next real size repairs.
                        if w <= 0 || h <= 0 { return None; }
                        let scroll_w = 20i32;
                        let scroll_h = 20i32;
                        if let Ok(sb) = vscroll_sz.try_borrow() {
                            if let Some(vh) = sb.handle.hwnd() {
                                winapi::um::winuser::SetWindowPos(
                                    vh as _, std::ptr::null_mut(),
                                    w - scroll_w, 0, scroll_w, h - scroll_h,
                                    winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_SHOWWINDOW,
                                );
                            }
                        }
                        if let Ok(sb) = hscroll_sz.try_borrow() {
                            if let Some(hh) = sb.handle.hwnd() {
                                winapi::um::winuser::SetWindowPos(
                                    hh as _, std::ptr::null_mut(),
                                    0, h - scroll_h, w - scroll_w, scroll_h,
                                    winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_SHOWWINDOW,
                                );
                            }
                        }
                        // The child always fills the viewport (ranges are driven
                        // by the host via scroll_to; nothing pans).
                        if let Some(child_ptr) = *child_sz_child.borrow() {
                            winapi::um::winuser::SetWindowPos(
                                child_ptr as _, std::ptr::null_mut(),
                                0, 0, (w - scroll_w).max(0), (h - scroll_h).max(0),
                                winapi::um::winuser::SWP_NOZORDER | winapi::um::winuser::SWP_SHOWWINDOW,
                            );
                        }
                    }
                    None
                },
            ).ok() { handlers.push(h); }

            // Scroll messages from the scrollbar children arrive at the frame
            // (their parent), not at the scrollbars themselves. Compute the new
            // thumb position from the request and report it; the host moves its
            // cursor (the child keeps filling the viewport; nothing pans).
            // SB_LINEUP/DOWN = 0/1, SB_PAGEUP/DOWN = 2/3, SB_THUMBPOSITION = 4,
            // SB_THUMBTRACK = 5, SB_TOP/BOTTOM = 6/7, SB_ENDSCROLL = 8.
            let vscroll_msg = vscroll.clone();
            let hscroll_msg = hscroll.clone();
            let scroll_cb = on_scroll_cb.clone();
            static SCROLL_MSG_ID: AtomicUsize = AtomicUsize::new(0x90000000);
            let smid = SCROLL_MSG_ID.fetch_add(1, Ordering::SeqCst);
            if let Some(h) = nwg::bind_raw_event_handler(
                &nwg::ControlHandle::Hwnd(hwnd as _), smid,
                move |_h, msg, w, _l| {
                    let vertical = if msg == winapi::um::winuser::WM_VSCROLL {
                        true
                    } else if msg == winapi::um::winuser::WM_HSCROLL {
                        false
                    } else {
                        return None;
                    };
                    let sb = if vertical { &vscroll_msg } else { &hscroll_msg };
                    let Ok(sb) = sb.try_borrow() else { return Some(0); };
                    let req = (w & 0xFFFF) as u32;
                    if req == 8 {
                        return Some(0); // SB_ENDSCROLL: nothing to do
                    }
                    let cur = sb.pos();
                    let end = sb.range().end.max(1);
                    let track = ((w >> 16) & 0xFFFF) as usize;
                    let new_pos = match req {
                        4 | 5 => track.min(end), // thumb position/track
                        0 => cur.saturating_sub(1), // line up/left
                        1 => (cur + 1).min(end), // line down/right
                        2 => cur.saturating_sub(10), // page up/left
                        3 => (cur + 10).min(end), // page down/right
                        6 => 0,  // top/left end
                        7 => end, // bottom/right end
                        _ => cur,
                    };
                    sb.set_pos(new_pos);
                    if let Some(ref mut f) = *scroll_cb.borrow_mut() {
                        f(vertical, new_pos as f64);
                    }
                    Some(0)
                },
            ).ok() { handlers.push(h); }

        }

        Ok(ScrolledWindow {
            frame: Some(Rc::new(frame)),
            hwnd,
            child,
            vscroll,
            hscroll,
            _handlers: Rc::new(handlers),
            on_scroll_cb,
        })
    }

    // ========== Menu / MenuBar / SimpleAction ==========

    /// Convert adapter MenuItemData → nwg builder format
    fn as_nwg_data(items: &[MenuItem]) -> Vec<crate::backends::nwg::MenuItemData> {
        items.iter().map(|i| crate::backends::nwg::MenuItemData {
            label: i.label.clone(),
            detailed_action: i.action.clone(),
            submenu: i.submenu.as_ref().map(|s| as_nwg_data(&s.items)),
        }).collect()
    }

    // -- Menu (data-only, builds nothing until consumed by MenuBar) --

    struct MenuItem {
        label: String,
        action: String,
        submenu: Option<Menu>,
    }

    pub struct Menu {
        items: Vec<MenuItem>,
    }

    impl Menu {
        pub fn append(&mut self, label: &str, detailed_action: &str) {
            self.items.push(MenuItem {
                // The model speaks GTK (`_` mnemonics); translate once to
                // the Win32 `&` marker here so no call site invents its own.
                label: crate::win32_portable::gtk_mnemonic_to_win32(label),
                action: detailed_action.to_string(),
                submenu: None,
            });
        }
        pub fn append_submenu(&mut self, label: &str, submenu: &Menu) {
            self.items.push(MenuItem {
                label: crate::win32_portable::gtk_mnemonic_to_win32(label),
                action: String::new(),
                submenu: Some(submenu.clone()),
            });
        }
    }

    impl Clone for Menu { fn clone(&self) -> Self { Menu { items: self.items.iter().map(|i| MenuItem {
        label: i.label.clone(), action: i.action.clone(), submenu: i.submenu.clone(),
    }).collect() } } }

    pub fn create_menu() -> Result<Menu, Error> { Ok(Menu { items: Vec::new() }) }

    // -- MenuBar: builds NWG menus, wires actions --

    /// Table mapping (hmenu_ptr, item_index) → stripped action name
    type MenuIndex = HashMap<(*mut std::ffi::c_void, u32), String>;

    pub struct MenuBar {
        pub(crate) _menus: Rc<Vec<nwg::Menu>>,
        pub(crate) _items: Rc<Vec<nwg::MenuItem>>,
        pub(crate) _raw_handler: Rc<nwg::RawEventHandler>,
        pub(crate) action_registry: Rc<RefCell<HashMap<String, Box<dyn FnMut()>>>>,
    }

    impl Clone for MenuBar {
        fn clone(&self) -> Self {
            MenuBar {
                _menus: self._menus.clone(),
                _items: self._items.clone(),
                _raw_handler: self._raw_handler.clone(),
                action_registry: self.action_registry.clone(),
            }
        }
    }

    impl AsRef<*mut c_void> for MenuBar {
        fn as_ref(&self) -> &*mut c_void {
            static NULL: usize = 0;
            unsafe { &*(&NULL as *const usize as *const *mut c_void) }
        }
    }

    impl Widget for MenuBar {
        fn raw_handle(&self) -> *mut c_void {
            std::ptr::null_mut()
        }
    }

    // Keyboard-menu shims required by common.rs: native Win32 menus handle
    // Alt+letter mnemonics in the OS (see the `&` prefixes in
    // `create_menubar`), so the manual popover-navigation contract is a
    // no-op here — matching the documented "On NWG ... this is a no-op".
    impl MenuBar {
        pub fn activate_submenu_by_mnemonic(&self, _keyval: u32) -> bool {
            false
        }
        pub fn activate_submenu_item_by_mnemonic(&self, _keyval: u32) -> bool {
            false
        }
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

    /// Recursively build NWG menu items, recording the (hmenu, index)→action mapping.
    /// Collectors `menus` and `items` keep the NWG objects alive (else Drop destroys them).
    fn build_and_index(
        parent_handle: &nwg::ControlHandle,
        items: &[crate::backends::nwg::MenuItemData],
        index: &mut MenuIndex,
        menus: &mut Vec<nwg::Menu>,
        items_collector: &mut Vec<nwg::MenuItem>,
    ) -> Result<(), nwg::NwgError> {
        for (i, item) in items.iter().enumerate() {
            if let Some(ref children) = item.submenu {
                let mut sub = nwg::Menu::default();
                nwg::Menu::builder()
                    .text(&item.label)
                    .popup(false)
                    .parent(parent_handle.clone())
                    .build(&mut sub)?;
                let sub_hmenu = sub.handle.hmenu().map(|(_, h)| h).unwrap_or(std::ptr::null_mut());
                let sub_handle = match parent_handle {
                    nwg::ControlHandle::Hwnd(h) => nwg::ControlHandle::PopMenu(*h, sub_hmenu),
                    nwg::ControlHandle::PopMenu(h, _) => nwg::ControlHandle::PopMenu(*h, sub_hmenu),
                    _ => unreachable!(),
                };
                menus.push(sub);
                build_and_index(&sub_handle, children, index, menus, items_collector)?;
            } else {
                let mut mi = nwg::MenuItem::default();
                nwg::MenuItem::builder()
                    .text(&item.label)
                    .parent(parent_handle.clone())
                    .build(&mut mi)?;
                let parent_hmenu: *mut c_void = match parent_handle {
                    nwg::ControlHandle::Hwnd(h) => *h as *mut c_void,
                    nwg::ControlHandle::PopMenu(_h, m) => *m as *mut c_void,
                    _ => unreachable!(),
                };
                if !item.detailed_action.is_empty() {
                    index.insert((parent_hmenu as *mut c_void, i as u32), item.detailed_action.clone());
                }
                items_collector.push(mi);
            }
        }
        Ok(())
    }

    /// Count leaf items in the menu tree (sequential numbering).
#[allow(dead_code)]
    fn count_leaves(items: &[crate::backends::nwg::MenuItemData]) -> u32 {
        let mut n = 0;
        for i in items {
            if let Some(ref children) = i.submenu {
                n += count_leaves(children);
            } else {
                n += 1;
            }
        }
        n
    }

    pub fn create_menubar(
        model: &Menu,
        window_hwnd: *mut c_void,
        action_registry: Rc<RefCell<HashMap<String, Box<dyn FnMut()>>>>,
    ) -> Result<MenuBar, Error> {
        let window_handle = nwg::ControlHandle::Hwnd(window_hwnd as _);
        let mut index: MenuIndex = HashMap::new();
        let data = as_nwg_data(&model.items);
        let mut menus: Vec<nwg::Menu> = Vec::new();
        let mut items: Vec<nwg::MenuItem> = Vec::new();

        // Build top-level menubar entries (File, Edit, Help, ...)
        for item in &data {
            if let Some(ref children) = item.submenu {
                let mut menu = nwg::Menu::default();
                nwg::Menu::builder()
                    .text(&item.label)
                    .popup(false)
                    .parent(&window_handle)
                    .build(&mut menu)
                    .map_err(|e| Error::Backend(format!("{}", e)))?;
                let menu_hmenu = menu.handle.hmenu().map(|(_, h)| h).unwrap_or(std::ptr::null_mut());
                let menu_handle = nwg::ControlHandle::PopMenu(
                    window_hwnd as *mut std::ffi::c_void as _,
                    menu_hmenu,
                );
                menus.push(menu);
                build_and_index(&menu_handle, children, &mut index, &mut menus, &mut items)
                    .map_err(|e| Error::Backend(format!("{}", e)))?;
            } else {
                let mut mi = nwg::MenuItem::default();
                nwg::MenuItem::builder()
                    .text(&item.label)
                    .parent(&window_handle)
                    .build(&mut mi)
                    .map_err(|e| Error::Backend(format!("{}", e)))?;
                items.push(mi);
            }
        }

        // Bind raw event handler for WM_MENUCOMMAND
        // handler_id must be > 0xFFFF (NWG reserves lower IDs)
        const RAW_MENU_ID: usize = 0x10001;
        let idx = index.clone();
        let reg = action_registry.clone();
        let raw_handler = nwg::bind_raw_event_handler(
            &nwg::ControlHandle::Hwnd(window_hwnd as _),
            RAW_MENU_ID,
            move |_hwnd, msg, wparam, lparam| {
                if msg != winapi::um::winuser::WM_MENUCOMMAND { return None; }
                let item_index = (wparam & 0xFFFF) as u32;
                let hmenu = lparam as *mut c_void;
                let key = (hmenu, item_index);
                if let Some(action_name) = idx.get(&key) {
                    let stripped = action_name.rsplit('.').next().unwrap_or(action_name);
                    if let Some(cb) = reg.borrow_mut().get_mut(stripped) {
                        cb();
                    }
                }
                Some(0)
            },
        ).map_err(|e| Error::Backend(format!("{}", e)))?;

        Ok(MenuBar { _menus: Rc::new(menus), _items: Rc::new(items), _raw_handler: Rc::new(raw_handler), action_registry })
    }

    // -- SimpleAction: stores callback by action name --

    pub struct SimpleAction {
        pub(crate) name: String,
        pub(crate) registry: Rc<RefCell<HashMap<String, Box<dyn FnMut()>>>>,
    }

    impl Clone for SimpleAction {
        fn clone(&self) -> Self {
            SimpleAction {
                name: self.name.clone(),
                registry: self.registry.clone(),
            }
        }
    }

    impl SimpleAction {
        pub fn connect_activate<F: FnMut(*mut c_void) + 'static>(&self, mut f: F) -> Result<u64, Error> {
            let mut map = self.registry.borrow_mut();
            map.insert(self.name.clone(), Box::new(move || f(std::ptr::null_mut())));
            Ok(0)
        }
    }

    pub fn create_simple_action(
        name: &str,
        registry: Rc<RefCell<HashMap<String, Box<dyn FnMut()>>>>,
    ) -> Result<SimpleAction, Error> {
        Ok(SimpleAction {
            name: name.to_string(),
            registry,
        })
    }
    // ---- File dialogs ----

    /// nwg filter-spec for `(name, patterns)` pairs:
    /// `Name (*.a;*.b)|*.a;*.b|…`. Empty list means no filter.
    /// nwg filter spec for `(name, patterns)` pairs; the format
    /// (`Name (*.a;*.b)|*.a;*.b|…`) is built portably and unit-tested in
    /// `win32_portable::join_dialog_filters`.
    #[cfg(windows)]
    fn join_filters(filters: &[(&str, &[&str])]) -> Option<String> {
        crate::win32_portable::join_dialog_filters(filters)
    }

    pub fn open_file(title: &str, parent: *mut c_void) -> Result<Option<String>, Error> {
        open_file_filtered(title, parent, &[])
    }

    pub fn open_file_filtered(title: &str, parent: *mut c_void, filters: &[(&str, &[&str])]) -> Result<Option<String>, Error> {
        let mut dialog = nwg::FileDialog::default();
        let mut builder = nwg::FileDialog::builder();
        builder = builder.title(title).action(nwg::FileDialogAction::Open);
        if let Some(spec) = join_filters(filters) {
            builder = builder.filters(&spec);
        }
        builder
            .build(&mut dialog)
            .map_err(|e| Error::Backend(format!("{}", e)))?;
        let parent_handle = nwg::ControlHandle::Hwnd(parent as _);
        if dialog.run(Some(&parent_handle)) {
            match dialog.get_selected_item() {
                Ok(path) => Ok(Some(path.to_string_lossy().into_owned())),
                Err(_) => Ok(None),
            }
        } else {
            Ok(None)
        }
    }

    pub fn save_file(title: &str, parent: *mut c_void) -> Result<Option<String>, Error> {
        save_file_filtered(title, parent, &[], "")
    }

    /// Save dialog with file-type filters and an optional suggested
    /// filename. Filters follow nwg's `Name (*.a;*.b)|*.a;*.b` spec; the
    /// native dialog offers the selected type's extension, prompts on
    /// overwrite, and appends the extension itself.
    pub fn save_file_filtered(title: &str, parent: *mut c_void, filters: &[(&str, &[&str])], _current_name: &str) -> Result<Option<String>, Error> {
        let mut dialog = nwg::FileDialog::default();
        let mut builder = nwg::FileDialog::builder();
        builder = builder.title(title).action(nwg::FileDialogAction::Save);
        if let Some(spec) = join_filters(filters) {
            builder = builder.filters(&spec);
        }
        // NOTE: nwg exposes no initial-filename setter; the suggested name
        // is GTK-only (callers still append the extension themselves, so
        // both backends land on the same bytes).
        builder
            .build(&mut dialog)
            .map_err(|e| Error::Backend(format!("{}", e)))?;
        let parent_handle = nwg::ControlHandle::Hwnd(parent as _);
        if dialog.run(Some(&parent_handle)) {
            match dialog.get_selected_item() {
                Ok(path) => Ok(Some(path.to_string_lossy().into_owned())),
                Err(_) => Ok(None),
            }
        } else {
            Ok(None)
        }
    }

    pub fn quit_main_loop() {
        crate::backends::nwg::quit_main_loop();
    }

    /// Diagnostic: dump the native window tree under `root` (hwnd, class,
    /// rect, visibility, leading text) to `path`. Env-gated by callers;
    /// no behavior change when unused. Exists so layout bugs can be
    /// diagnosed from real Windows screenshots' ground truth.
    pub fn debug_dump_native_tree(root: *mut c_void, path: &str) {
        struct Ctx {
            lines: Vec<String>,
        }
        unsafe extern "system" fn enum_cb(
            hwnd: winapi::shared::windef::HWND,
            lparam: winapi::shared::minwindef::LPARAM,
        ) -> i32 {
            let ctx = &mut *(lparam as *mut Ctx);
            unsafe {
                let mut cls: [u16; 64] = [0; 64];
                let n = winapi::um::winuser::GetClassNameW(hwnd, cls.as_mut_ptr(), 64);
                let class = String::from_utf16_lossy(&cls[..n.max(0) as usize]);
                let mut rect: winapi::shared::windef::RECT = std::mem::zeroed();
                winapi::um::winuser::GetWindowRect(hwnd, &mut rect);
                let vis = winapi::um::winuser::IsWindowVisible(hwnd) != 0;
                let tlen = winapi::um::winuser::GetWindowTextLengthW(hwnd);
                let mut text = String::new();
                if tlen > 0 && tlen < 80 {
                    let mut buf: Vec<u16> = vec![0; (tlen + 1) as usize];
                    winapi::um::winuser::GetWindowTextW(hwnd, buf.as_mut_ptr(), tlen + 1);
                    text = String::from_utf16_lossy(&buf[..tlen as usize]);
                }
                ctx.lines.push(format!(
                    "hwnd={:p} class={} rect=({},{})-({},{}) visible={} text={:?}",
                    hwnd, class, rect.left, rect.top, rect.right, rect.bottom, vis, text
                ));
            }
            1
        }
        unsafe {
            let mut ctx = Ctx { lines: Vec::new() };
            // Root itself first.
            enum_cb(root as _, &mut ctx as *mut Ctx as _);
            winapi::um::winuser::EnumChildWindows(root as _, Some(enum_cb), &mut ctx as *mut Ctx as _);
            let _ = std::fs::write(path, ctx.lines.join("\n") + "\n");
        }
    }
}

#[cfg(windows)]
pub use nwg_adapter::*;
#[cfg(windows)]
pub use crate::backends::nwg::Orientation;
