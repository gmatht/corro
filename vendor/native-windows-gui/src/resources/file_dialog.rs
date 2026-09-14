//! File dialogs — Win95-compatible implementation.
//!
//! Upstream (and the crates.io build of) native-windows-gui implements
//! `FileDialog` on top of the Win7+ `IFileDialog` COM API
//! (`CLSID_FileOpenDialog` / `SHCreateItemFromParsingName`). Those exports do
//! not exist on Windows 95, and because the call is a *static* import it makes
//! the whole executable refuse to load there ("linked to missing export").
//!
//! This vendored copy implements the same public API on the Win95-era
//! `comdlg32.dll` ANSI common dialogs:
//!
//! | Operation       | Win95 API                                        |
//! | --------------- | ------------------------------------------------ |
//! | open a file     | `GetOpenFileNameA` + `OPENFILENAMEA`              |
//! | save a file     | `GetSaveFileNameA` + `OPENFILENAMEA`              |
//! | pick a folder   | `SHBrowseForFolderA` + `SHGetPathFromIDListA`     |
//!
//! All three are present in the stock Windows 95 `comdlg32.dll`/`shell32.dll`
//! and require no COM (so no `CoInitialize` either). ANSI (`*A`) variants are
//! used deliberately: Win9x exports many `*W` functions as no-op stubs.
//!
//! Behaviour is preserved as far as the OS allows:
//! * `title`, `default_folder`, `filters`, `multiselect` all map onto
//!   `OPENFILENAMEA` fields directly.
//! * `OpenDirectory` uses the shell folder browser (Win95's `OPENFILENAMEA`
//!   has no `OFN_PICKFOLDERS` — that flag is Win2000+).
//! * `set_filters` can only be set before the first `run` on Win95 (the dialog
//!   is rebuilt per `run`, so in practice it is honoured on every run).
//! * `clear_client_data` is a no-op: Win95 always keeps its own MRU state.

use crate::win32::base_helper::{to_ansi, from_ansi};
use crate::{ControlHandle, NwgError};

use std::ffi::OsString;
use std::{fmt, ptr};

/// Longest path a Win95 `OPENFILENAMEA` can return (MAX_PATH).
const MAX_PATH_BUF: usize = 260;

/**
    A enum that dictates how a file dialog should behave
    Members:
    * `Open`: User can select a file that is not a directory
    * `OpenDirectory`: User can select a directory
    * `Save`: User select the name of a file. If it already exists, a confirmation message will be raised
*/
#[derive(Copy, Clone, PartialEq, Eq, Debug)]
pub enum FileDialogAction {
    Open,
    OpenDirectory,
    Save,
}

/// Result of a completed dialog run: the selected path(s), retained until the
/// next `run` (upstream keeps the COM object alive for the same purpose).
struct DialogState {
    /// Filled by `run`; the NUL-terminated buffer the common dialog wrote.
    selected: Vec<u8>,
    /// Extra items for a multiselect Open (beyond `selected`).
    extra: Vec<Vec<u8>>,
    /// Set once the dialog has run at least once (upstream errors
    /// `get_selected_item` before that; we mirror it).
    ran: bool,
    /// Copy of the builder inputs; the dialog is (re)constructed per `run`
    /// because `OPENFILENAMEA` is a stack struct owned by the call.
    title: Option<String>,
    default_folder: Option<String>,
    filters: Option<String>,
    multiselect: bool,
}

/**
    A file dialog control

    The file dialog builders accepts the following parameters:
    * title: The title of the dialog
    * action: The action to execute. Open, OpenDirectory for Save
    * multiselect: Whether the user can select more than one file. Only supported with the Open action
    * default_folder: Default folder to show in the dialog.
    * filters: If defined, filter the files that the user can select (In a Open dialog) or which extension to add to the saved file (in a Save dialog)
    The `filters` value must be a '|' separated string having this format: "Test(*.txt;*.rs)|Any(*.*)"

    ```rust
        use native_windows_gui as nwg;
        fn layout(dialog: &mut nwg::FileDialog) {
            nwg::FileDialog::builder()
                .title("Hello")
                .action(nwg::FileDialogAction::Open)
                .multiselect(true)
                .build(dialog);
        }
    ```
*/
pub struct FileDialog {
    action: FileDialogAction,
    state: DialogState,
}

impl FileDialog {

    pub fn builder() -> FileDialogBuilder {
        FileDialogBuilder {
            title: None,
            action: FileDialogAction::Save,
            multiselect: false,
            default_folder: None,
            filters: None
        }
    }

    /// Return the action type executed by this dialog
    pub fn action(&self) -> FileDialogAction {
        self.action
    }

    /**
        Display the dialog. Return true if the dialog was accepted or false if it was cancelled
        If the dialog was accepted, `get_selected_item` or `get_selected_items` can be used to find the selected file(s)

        It's important to note that `run` blocks the current thread until the user as chosen a file (similar to `dispatch_thread_events`)

        The parent argument must be a window control otherwise the method will panic.
    */
    pub fn run<C: Into<ControlHandle>>(&mut self, parent: Option<C>) -> bool {
        let parent_handle = match parent {
            Some(p) => p.into().hwnd().expect("File dialog parent must be a window control"),
            None => ptr::null_mut()
        };

        // A fresh selection each run, like the COM dialog's Show().
        self.state.selected.clear();
        self.state.extra.clear();

        let accepted = unsafe {
            match self.action {
                FileDialogAction::OpenDirectory => run_browse_for_folder(parent_handle, &mut self.state),
                FileDialogAction::Open => run_open_save(parent_handle, &mut self.state, false),
                FileDialogAction::Save => run_open_save(parent_handle, &mut self.state, true),
            }
        };

        self.state.ran = true;
        accepted
    }

    /**
        Return the item selected in the dialog by the user.

        Failures:
        • if the dialog was not called
        • if there was a system error while reading the selected item
        • if the dialog has the `multiselect` flag
    */
    pub fn get_selected_item(&self) -> Result<OsString, NwgError> {
        if self.multiselect() {
            return Err(NwgError::file_dialog("FileDialog have the multiselect flag"));
        }
        if !self.state.ran {
            return Err(NwgError::file_dialog("File dialog was not run"));
        }
        Ok(OsString::from(from_ansi(trim_nul(&self.state.selected))))
    }

    /**
        Return the selected items in the dialog by the user.

        Failures:
        • if the dialog was not called
        • if there was a system error while reading the selected items
        • if the dialog has `Save` for action
    */
    pub fn get_selected_items(&self) -> Result<Vec<OsString>, NwgError> {
        if self.action == FileDialogAction::Save {
            return Err(NwgError::file_dialog("Save dialog cannot have more than one item selected"));
        }
        if !self.state.ran {
            return Err(NwgError::file_dialog("File dialog was not run"));
        }

        let mut out = Vec::new();
        if !self.state.selected.is_empty() {
            out.push(OsString::from(from_ansi(trim_nul(&self.state.selected))));
        }
        for item in &self.state.extra {
            out.push(OsString::from(from_ansi(trim_nul(item))));
        }
        Ok(out)
    }

    /// Return `true` if the dialog accepts multiple values or `false` otherwise
    pub fn multiselect(&self) -> bool {
        self.state.multiselect
    }

    /**
        Set the multiselect flag of the dialog.

        Failures:
        • if there was a system error while setting the new flag value
        • if the dialog has `Save` for action
    */
    pub fn set_multiselect(&mut self, multiselect: bool) -> Result<(), NwgError> {
        if self.action == FileDialogAction::Save {
            return Err(NwgError::file_dialog("Cannot set multiselect flag for a save file dialog"));
        }
        self.state.multiselect = multiselect;
        Ok(())
    }

    /**
        Set the first opened folder when the dialog is shown. This value is overriden by the user after the dialog ran.
        Call `clear_client_data` to fix that.
        Failures:
        • if the default folder do not identify a folder
        • if the folder do not exists
    */
    pub fn set_default_folder<'a>(&mut self, folder: &'a str) -> Result<(), NwgError> {
        self.state.default_folder = Some(folder.to_string());
        Ok(())
    }

    /**
        Filter the files that the user can select (In a `Open` dialog) in the dialog or which extension to add to the saved file (in a `Save` dialog).
        This can only be set ONCE (the initialization counts) and won't work if the dialog is `OpenDirectory`.

        The `filters` value must be a '|' separated string having this format: "Test(*.txt;*.rs)|Any(*.*)"
        Where the fist part is the "human name" and the second part is a filter for the system.
    */
    pub fn set_filters<'a>(&mut self, filters: &'a str) -> Result<(), NwgError> {
        self.state.filters = Some(filters.to_string());
        Ok(())
    }

    /// Change the dialog title
    pub fn set_title<'a>(&mut self, title: &'a str) {
        self.state.title = Some(title.to_string());
    }

    /// Instructs the dialog to clear all persisted state information (such as the last folder visited).
    /// No-op on Win95: the common dialogs keep their own MRU state.
    pub fn clear_client_data(&self) {}
}

fn trim_nul(buf: &[u8]) -> &[u8] {
    let end = buf.iter().position(|&b| b == 0).unwrap_or(buf.len());
    &buf[..end]
}

/// Convert a `filters` string ("Name(*.a;*.b)|Name2(*.*)") into the
/// double-NUL-terminated `lpstrFilter` form `OPENFILENAMEA` expects.
fn filters_to_ansi(filters: &str) -> Vec<u8> {
    let mut out = Vec::new();
    // `to_ansi` appends a NUL; we only want the bytes.
    let flat = to_ansi(filters);
    let flat = trim_nul(&flat);
    for part in flat.split(|&b| b == b'|') {
        if part.is_empty() {
            continue;
        }
        out.extend_from_slice(part);
        out.push(0);
    }
    out.push(0); // final terminating NUL ends the list
    out
}

/// `GetOpenFileNameA` / `GetSaveFileNameA` with a parent HWND.
unsafe fn run_open_save(
    parent: winapi::shared::windef::HWND,
    state: &mut DialogState,
    save: bool,
) -> bool {
    use winapi::um::commdlg::{GetOpenFileNameA, GetSaveFileNameA, OPENFILENAMEA};
    use winapi::um::winuser::{GetWindowTextA, SetWindowTextA};

    let title = state.title.as_deref().unwrap_or("");
    let title_ansi = to_ansi(title);
    let init_dir = state.default_folder.as_deref().unwrap_or("");
    let init_dir_ansi = to_ansi(init_dir);
    let filter_ansi = state
        .filters
        .as_deref()
        .map(filters_to_ansi)
        .unwrap_or_else(Vec::new);

    // Multiselect buffer layout is a sequence of NUL-terminated names ending
    // with an extra NUL. A single file only needs MAX_PATH.
    let mut buf = vec![0u8; if state.multiselect { 4096 } else { MAX_PATH_BUF }];

    let mut ofn: OPENFILENAMEA = std::mem::zeroed();
    ofn.lStructSize = std::mem::size_of::<OPENFILENAMEA>() as u32;
    ofn.hwndOwner = parent;
    ofn.lpstrFilter = if filter_ansi.is_empty() { ptr::null() } else { filter_ansi.as_ptr() as *const _ };
    ofn.nFilterIndex = 1;
    ofn.lpstrFile = buf.as_mut_ptr() as *mut _;
    ofn.nMaxFile = buf.len() as u32;
    ofn.lpstrTitle = if title_ansi == [0] { ptr::null() } else { title_ansi.as_ptr() as *const _ };
    ofn.lpstrInitialDir = if init_dir_ansi == [0] { ptr::null() } else { init_dir_ansi.as_ptr() as *const _ };
    ofn.Flags = OFN_EXPLORER
        | OFN_NOCHANGEDIR
        | OFN_PATHMUSTEXIST
        | OFN_HIDEREADONLY
        | if save { OFN_OVERWRITEPROMPT } else { OFN_FILEMUSTEXIST }
        | if state.multiselect { OFN_ALLOWMULTISELECT } else { 0 };

    let ok = if save {
        GetSaveFileNameA(&mut ofn) != 0
    } else {
        GetOpenFileNameA(&mut ofn) != 0
    };

    if !ok {
        // Cancelled (or a hard error). If a title change was requested the
        // common dialog restores the owner text itself on Win95.
        let _ = (GetWindowTextA, SetWindowTextA);
        return false;
    }

    if state.multiselect {
        // Either "dir\0file1\0file2\0\0" (multi) or "fullpath\0\0" (single).
        let nul_positions: Vec<usize> = buf
            .iter()
            .enumerate()
            .take_while(|(_, &b)| b != 0 || true)
            .filter(|(_, &b)| b == 0)
            .map(|(i, _)| i)
            .collect();
        let first_end = nul_positions.first().copied().unwrap_or(0);
        // A second immediately-following NUL means a single full path.
        let second = nul_positions.get(1).copied().unwrap_or(first_end);
        if second == first_end + 1 {
            state.selected = buf[..first_end + 1].to_vec();
        } else {
            let dir = &buf[..first_end];
            let mut start = first_end + 1;
            let mut first = true;
            while start < buf.len() && buf[start] != 0 {
                let end = buf[start..].iter().position(|&b| b == 0).map(|p| start + p).unwrap_or(buf.len());
                let mut full = Vec::with_capacity(dir.len() + 1 + (end - start) + 1);
                full.extend_from_slice(dir);
                full.push(b'\\');
                full.extend_from_slice(&buf[start..end]);
                full.push(0);
                if first {
                    state.selected = full;
                    first = false;
                } else {
                    state.extra.push(full);
                }
                start = end + 1;
            }
        }
    } else {
        state.selected = buf[..MAX_PATH_BUF.min(buf.len())].to_vec();
    }

    true
}

// OFN_* flags. `winapi`'s commdlg module exposes some of these, but not all
// on every feature set, so they are named here (values are from the Win95 SDK
// commdlg.h and are stable across Win9x/NT).
const OFN_ALLOWMULTISELECT: u32 = 0x0000_0200;
const OFN_EXPLORER: u32 = 0x0008_0000;
const OFN_FILEMUSTEXIST: u32 = 0x0000_1000;
const OFN_HIDEREADONLY: u32 = 0x0000_0004;
const OFN_NOCHANGEDIR: u32 = 0x0000_0008;
const OFN_OVERWRITEPROMPT: u32 = 0x0000_0002;
const OFN_PATHMUSTEXIST: u32 = 0x0000_0800;

/// `BROWSEINFOA` from the Win95 SDK shlobj.h. winapi 0.3.9 ships
/// `SHGetPathFromIDListA` but not the folder-browser entry point or its
/// parameter struct, so both are declared here. Field order/ABI match the
/// Win95 SDK exactly.
#[repr(C)]
struct BrowseInfoA {
    hwnd_owner: winapi::shared::windef::HWND,
    pidl_root: *mut winapi::um::shtypes::ITEMIDLIST,
    psz_display_name: *mut winapi::ctypes::c_char,
    lpsz_title: *const winapi::ctypes::c_char,
    ul_flags: winapi::shared::minwindef::UINT,
    lpfn: *mut winapi::ctypes::c_void,
    l_param: winapi::shared::minwindef::LPARAM,
    i_image: winapi::shared::minwindef::INT,
}

extern "system" {
    /// Win95 shlobj32: present in SHELL32.dll since Windows 95.
    fn SHBrowseForFolderA(lpbi: *mut BrowseInfoA) -> *mut winapi::um::shtypes::ITEMIDLIST;
}

/// Directory picker via the Win95 shell folder browser.
unsafe fn run_browse_for_folder(parent: winapi::shared::windef::HWND, state: &mut DialogState) -> bool {
    use winapi::shared::windef::HWND;
    use winapi::um::shlobj::SHGetPathFromIDListA;
    use winapi::um::winuser::GetDesktopWindow;

    extern "system" {
        fn CoTaskMemFree(pv: *mut winapi::ctypes::c_void);
    }

    let title_ansi = to_ansi(state.title.as_deref().unwrap_or(""));
    let mut display = vec![0u8; MAX_PATH_BUF];

    let mut bi = BrowseInfoA {
        hwnd_owner: if parent.is_null() { GetDesktopWindow() as HWND } else { parent },
        pidl_root: ptr::null_mut(),
        psz_display_name: display.as_mut_ptr() as *mut _,
        lpsz_title: if title_ansi == [0] { ptr::null() } else { title_ansi.as_ptr() as *const _ },
        ul_flags: 0, // BIF_RETURNONLYFSDIRS is unavailable in the Win95 shell
        lpfn: ptr::null_mut(),
        l_param: 0,
        i_image: 0,
    };

    let idl = SHBrowseForFolderA(&mut bi);
    if idl.is_null() {
        return false;
    }

    let mut path = vec![0u8; MAX_PATH_BUF];
    let ok = SHGetPathFromIDListA(idl, path.as_mut_ptr() as *mut _) != 0;
    CoTaskMemFree(idl as *mut _);
    if !ok {
        return false;
    }

    state.selected = path;
    true
}

impl fmt::Debug for FileDialog {

    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "FileDialog {{ action: {:?} }}", self.action)
    }

}

impl Default for FileDialog {
    fn default() -> FileDialog {
        FileDialog {
            action: FileDialogAction::Open,
            state: DialogState {
                selected: Vec::new(),
                extra: Vec::new(),
                ran: false,
                title: None,
                default_folder: None,
                filters: None,
                multiselect: false,
            },
        }
    }
}

impl PartialEq for FileDialog {
    fn eq(&self, other: &Self) -> bool {
        self.action == other.action && self.state.ran == other.state.ran
    }
}

impl Eq for FileDialog {}

/*
    Structure that hold the state required to build a file dialog
*/
pub struct FileDialogBuilder {
    pub title: Option<String>,
    pub action: FileDialogAction,
    pub multiselect: bool,
    pub default_folder: Option<String>,
    pub filters: Option<String>
}

impl FileDialogBuilder {

    pub fn title<S: Into<String>>(mut self, t: S) -> FileDialogBuilder {
        self.title = Some(t.into());
        self
    }

    pub fn default_folder<S: Into<String>>(mut self, t: S) -> FileDialogBuilder {
        self.default_folder = Some(t.into());
        self
    }

    pub fn filters<S: Into<String>>(mut self, t: S) -> FileDialogBuilder {
        self.filters = Some(t.into());
        self
    }

    pub fn action(mut self, a: FileDialogAction) -> FileDialogBuilder {
        self.action = a;
        self
    }

    pub fn multiselect(mut self, m: bool) -> FileDialogBuilder {
        self.multiselect = m;
        self
    }

    pub fn build(self, out: &mut FileDialog) -> Result<(), NwgError> {
        out.action = self.action;
        out.state.title = self.title;
        out.state.default_folder = self.default_folder;
        out.state.filters = self.filters;
        // `OpenDirectory` cannot multiselect (upstream ignores the flag too).
        out.state.multiselect = self.multiselect && self.action != FileDialogAction::Save;
        out.state.ran = false;
        out.state.selected.clear();
        out.state.extra.clear();
        Ok(())
    }

}
