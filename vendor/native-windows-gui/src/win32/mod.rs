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
// TEMPORARY Win95 diagnosis: raw file marker (std::fs broken on 9x).
#[cfg(all(target_family = "rust9x", target_env = "msvc"))]
unsafe fn mark95w(s: &[u8]) {
    use winapi::um::fileapi::{CreateFileA, SetFilePointer, WriteFile, OPEN_ALWAYS};
    use winapi::um::handleapi::CloseHandle;
    use winapi::um::winnt::{GENERIC_WRITE, FILE_SHARE_READ, FILE_ATTRIBUTE_NORMAL, HANDLE};
    use winapi::shared::minwindef::{DWORD, LPCVOID, LPDWORD};
    use winapi::um::winnt::LPCSTR;
    use std::{ptr, mem};
    let _ = mem::size_of::<u8>();
    let h: HANDLE = CreateFileA(b"c:\\gcorro.log\0".as_ptr() as LPCSTR,
        GENERIC_WRITE, FILE_SHARE_READ, ptr::null_mut(), OPEN_ALWAYS,
        FILE_ATTRIBUTE_NORMAL, ptr::null_mut());
    if h.is_null() || h == winapi::um::handleapi::INVALID_HANDLE_VALUE {
        return;
    }
    SetFilePointer(h, 0, ptr::null_mut(), 2);
    let mut w: DWORD = 0;
    WriteFile(h, s.as_ptr() as LPCVOID, s.len() as DWORD, &mut w as LPDWORD, ptr::null_mut());
    CloseHandle(h);
}
pub fn init_common_controls() -> Result<(), NwgError> {
    use winapi::um::objbase::CoInitialize;
    use winapi::um::libloaderapi::LoadLibraryW;
    use winapi::um::commctrl::{InitCommonControlsEx, INITCOMMONCONTROLSEX};
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    use winapi::um::commctrl::InitCommonControls;
    use winapi::um::commctrl::{ICC_BAR_CLASSES, ICC_STANDARD_CLASSES, ICC_DATE_CLASSES, ICC_PROGRESS_CLASS,
     ICC_TAB_CLASSES, ICC_TREEVIEW_CLASSES, ICC_LISTVIEW_CLASSES};
    use winapi::shared::winerror::{S_OK, S_FALSE};

    unsafe {
        // TEMPORARY Win95 diagnosis.
        #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
        mark95w(b"coinit\n");
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
        // TEMPORARY Win95 diagnosis.
        #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
        mark95w(b"preicex\n");

        // Win95 note (vendored patch): this box's comctl32 faults inside
        // InitCommonControlsEx (~0x1C7 into the function, EIP 0x148dd4) for
        // the ICC set nwg requests. The Win95 feature set needs no comctl32
        // classes (combobox/textbox/scroll-bar are USER32 built-ins; the rest
        // are nwg custom classes), so the legacy InitCommonControls suffices.
        // TEMPORARY diagnosis: skip the call entirely.
        //#[cfg(all(target_family = "rust9x", target_env = "msvc"))]
        //InitCommonControls();
        //#[cfg(not(all(target_family = "rust9x", target_env = "msvc")))]
        //InitCommonControlsEx(&data);
        #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
        mark95w(b"skipcc\n");
        // TEMPORARY Win95 diagnosis.
        #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
        mark95w(b"icex\n");
    }

    // TEMPORARY Win95 diagnosis.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95w(b"prewinclass\n"); }
    window::init_window_class()?;
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95w(b"wclass\n"); }
    tabs_init()?;
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95w(b"tabs\n"); }
    extern_canvas_init()?;
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe { mark95w(b"canvas\n"); }
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

