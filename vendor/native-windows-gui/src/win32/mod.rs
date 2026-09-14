pub(crate) mod base_helper;
pub(crate) mod window_helper;
pub(crate) mod resources_helper;
pub(crate) mod window;
pub(crate) mod message_box;
pub(crate) mod high_dpi;
pub(crate) mod monitor;

#[cfg(feature = "menu")]
pub(crate) mod menu;

#[cfg(feature = "cursor")]
pub(crate) mod cursor;

#[cfg(feature = "clipboard")]
pub(crate) mod clipboard;

#[cfg(feature = "tabs")]
pub(crate) mod tabs;

#[cfg(feature = "extern-canvas")]
pub(crate) mod extern_canvas;

#[cfg(feature = "image-decoder")]
pub(crate) mod image_decoder;

#[cfg(feature = "rich-textbox")]
pub(crate) mod richedit;

#[cfg(feature = "plotting")]
pub(crate) mod plotters_d2d;

use std::{fs, mem, ptr};
use crate::errors::NwgError;


use winapi::um::winuser::{IsDialogMessageA, GetParent, TranslateMessage, DispatchMessageA};
use winapi::shared::windef::HWND;

/**
    Win95-compatible replacement for `GetAncestor(hwnd, GA_ROOT)`.
    `GetAncestor` requires Windows 98 or later and statically importing it makes
    the whole binary unloadable on Windows 95, so walk the parent chain with
    `GetParent` (available since Windows 3.1) instead.
*/
unsafe fn get_root_window(mut hwnd: HWND) -> HWND {
    loop {
        let parent = GetParent(hwnd);
        if parent.is_null() {
            return hwnd;
        }
        hwnd = parent;
    }
}

/**
    Dispatch system events in the current thread. This method will pause the thread until there are events to process.
*/
pub fn dispatch_thread_events() {
    use winapi::um::winuser::MSG;
    // Win95 patch: ANSI message pump (GetMessageW/DispatchMessageW are stubs on Win95)
    use winapi::um::winuser::GetMessageA;

    unsafe {
        let mut msg: MSG = mem::zeroed();
        while GetMessageA(&mut msg, ptr::null_mut(), 0, 0) != 0 {
            if IsDialogMessageA(get_root_window(msg.hwnd), &mut msg) == 0 {
                TranslateMessage(&msg); 
                DispatchMessageA(&msg); 
            }
        }
    }
}


/**
    Dispatch system events in the current thread AND execute a callback after each peeking attempt.
    Unlike `dispath_thread_events`, this method will not pause the thread while waiting for events.
*/
pub fn dispatch_thread_events_with_callback<F>(mut cb: F) 
    where F: FnMut() -> () + 'static
{
    use winapi::um::winuser::MSG;
    use winapi::um::winuser::{PeekMessageA, PM_REMOVE, WM_QUIT};

    unsafe {
        let mut msg: MSG = mem::zeroed();
        while msg.message != WM_QUIT {
            let has_message = PeekMessageA(&mut msg, ptr::null_mut(), 0, 0, PM_REMOVE) != 0;
            if has_message {
                if IsDialogMessageA(get_root_window(msg.hwnd), &mut msg) == 0 {
                    TranslateMessage(&msg); 
                    DispatchMessageA(&msg); 
                }
            }

            cb();
        }
    }
}

/**
    Break the events loop running on the current thread
*/
pub fn stop_thread_dispatch() {
  use winapi::um::winuser::PostMessageA;
  use winapi::um::winuser::WM_QUIT;

  unsafe { PostMessageA(ptr::null_mut(), WM_QUIT, 0, 0) };
}


/**
  Enable the Windows visual style in the application without having to use a manifest

  Win95 note (vendored patch): visual styles do not exist on Windows 95, and the
  original implementation statically imported `CreateActCtxW`/`ActivateActCtx`
  (Windows XP and later), which made the binary refuse to load on Windows 95.
  This is now a no-op so the executable stays loadable on Win9x.
*/
pub fn enable_visual_styles() {
}

/**
    Ensure that the dll containing the winapi controls is loaded.
    Also register the custom classes used by NWG
*/
pub fn init_common_controls() -> Result<(), NwgError> {
    use winapi::um::objbase::CoInitialize;
    use winapi::um::libloaderapi::LoadLibraryW;
    use winapi::um::commctrl::{InitCommonControlsEx, INITCOMMONCONTROLSEX};
    use winapi::um::commctrl::{ICC_BAR_CLASSES, ICC_STANDARD_CLASSES, ICC_DATE_CLASSES, ICC_PROGRESS_CLASS,
     ICC_TAB_CLASSES, ICC_TREEVIEW_CLASSES, ICC_LISTVIEW_CLASSES};
    use winapi::shared::winerror::{S_OK, S_FALSE};

    unsafe {
        let mut classes = ICC_BAR_CLASSES | ICC_STANDARD_CLASSES;

        if cfg!(feature = "datetime-picker") {
            classes |= ICC_DATE_CLASSES;
        }

        if cfg!(feature = "progress-bar") {
            classes |= ICC_PROGRESS_CLASS;
        }

        if cfg!(feature = "tabs") {
            classes |= ICC_TAB_CLASSES;
        }

        if cfg!(feature = "tree-view") {
            classes |= ICC_TREEVIEW_CLASSES;
        }

        if cfg!(feature = "list-view") {
            classes |= ICC_LISTVIEW_CLASSES;
        }

        if cfg!(feature = "rich-textbox") {
            let lib = base_helper::to_utf16("Msftedit.dll");
            LoadLibraryW(lib.as_ptr());
        }

        let data = INITCOMMONCONTROLSEX {
            dwSize: mem::size_of::<INITCOMMONCONTROLSEX>() as u32,
            dwICC: classes
        };

        InitCommonControlsEx(&data);
    }

    window::init_window_class()?;
    tabs_init()?;
    extern_canvas_init()?;
    frame_init()?;
    
    match unsafe { CoInitialize(ptr::null_mut()) } {
        S_OK | S_FALSE => Ok(()),
        // (Patch-Win95) On a pristine Win95 registry OLE32.CoInitialize can
        // fail with E_FAIL (0x80004005). Nothing in the ANSI build uses COM
        // (no OLE drag & drop, no shell COM), so keep going without it —
        // aborting init made every app die right after window creation.
        _ => Ok(()),
    }
}

#[cfg(feature = "tabs")]
fn tabs_init() -> Result<(), NwgError> { tabs::create_tab_classes() }

#[cfg(not(feature = "tabs"))]
fn tabs_init() -> Result<(), NwgError> { Ok(()) }

#[cfg(feature = "extern-canvas")]
fn extern_canvas_init() -> Result<(), NwgError> { extern_canvas::create_extern_canvas_classes() }

#[cfg(not(feature = "extern-canvas"))]
fn extern_canvas_init() -> Result<(), NwgError> { Ok(()) }

#[cfg(feature = "frame")]
fn frame_init() -> Result<(), NwgError> { window::create_frame_classes() }

#[cfg(not(feature = "frame"))]
fn frame_init() -> Result<(), NwgError> { Ok(()) }

