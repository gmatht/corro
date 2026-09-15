//! corro — append-only collaborative spreadsheet TUI.
#![allow(unexpected_cfgs)] // Win95/rust9x custom target_family gates are intentional

// Win95 (rust9x targets): std::rt initialization hangs inside KERNEL32 on
// Windows 95 (DBCS conversion loop). With `no_main`, the VC6 CRT startup
// calls our `main` directly, skipping std::rt entirely. Normal targets keep
// the standard Rust entry (and `cargo test` keeps working).
#![cfg_attr(all(target_family = "rust9x", target_env = "msvc"), no_main)]

#[cfg(feature = "ratatui")]
use corro::ui::App as TuiApp;
#[cfg(any(feature = "gui", feature = "pancurses"))]
use corro::gui::App as GuiApp;
use std::path::PathBuf;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum UiKind {
    Ratatui,
    #[allow(dead_code)]
    Gui,
    #[allow(dead_code)]
    Pancurses,
}

struct Args {
    revision: Option<RevisionMode>,
    files: Vec<PathBuf>,
    export: Option<PathBuf>,
    movie: bool,
    movie_typing_cps: f64,
    movie_confirm_ms: u64,
    movie_menu_hold_ms: u64,
    show_help: bool,
    show_version: bool,
    debug_no_number: bool,
    ui: UiKind,
    capture_html: Option<PathBuf>,
    convert_ansi: Option<PathBuf>,
}

enum RevisionMode {
    Browse,
    Limit(usize),
}

#[allow(dead_code)]
fn cli_option_suggestion(arg: &str) -> Option<&'static str> {
    match arg {
        "--movie-cps" => Some("--movie-typing-cps"),
        _ => None,
    }
}

#[cfg(target_arch = "wasm32")]
fn determine_default_ui() -> UiKind { UiKind::Gui }

// Terminal-first: a compiled-in terminal UI is always the default; the GUI
// is opt-in via an explicit --gui flag (see --help text). Without this,
// `cargo run --features gui` would pop a GUI window even though the user
// never asked for one.
#[cfg(all(not(target_arch = "wasm32"), feature = "ratatui"))]
fn determine_default_ui() -> UiKind { UiKind::Ratatui }

#[cfg(all(
    not(target_arch = "wasm32"),
    not(feature = "ratatui"),
    feature = "pancurses"
))]
fn determine_default_ui() -> UiKind { UiKind::Pancurses }

#[cfg(all(
    not(target_arch = "wasm32"),
    not(feature = "ratatui"),
    not(feature = "pancurses")
))]
fn determine_default_ui() -> UiKind { UiKind::Gui }

// Win95-safe command line: `GetCommandLineW` is a no-op stub on Windows 95
// (returns NULL), so `std::env::args()` cannot be used there. Read the ANSI
// command line instead. DBCS bytes decode lossily, which is fine for the
// ASCII-only options corro takes.
#[cfg(all(target_family = "rust9x", target_env = "msvc"))]
fn cli_args() -> impl Iterator<Item = String> {
    win95_args().into_iter().skip(1)
}

/// Default UI implied by the program name (busybox-style argv[0] dispatch):
/// invoked as `pcorro*` prefers the pancurses UI, as `gcorro*` the GUI.
/// Returns None when the name requests nothing or the requested backend is
/// not compiled in (callers fall back to [`determine_default_ui`]); explicit
/// `--gui`/`--ratatui`/`--pancurses` flags always win over this default.
fn argv0_ui(program: &str) -> Option<UiKind> {
    let base = program.rsplit(|c| c == '/' || c == '\\').next().unwrap_or(program);
    let base = base
        .strip_suffix(".exe")
        .or_else(|| base.strip_suffix(".EXE"))
        .unwrap_or(base);
    let lower = base.to_ascii_lowercase();
    #[cfg(feature = "pancurses")]
    if lower.starts_with("pcorro") {
        return Some(UiKind::Pancurses);
    }
    #[cfg(feature = "gui")]
    if lower.starts_with("gcorro") {
        return Some(UiKind::Gui);
    }
    None
}

/// Windows console attach for GUI-subsystem launches (MSIX/Store builds).
///
/// A GUI-subsystem exe is born console-less even when started from a
/// terminal. When the user wants a TUI anyway, AttachConsole claims the
/// parent's console; success is also the "launched in a console" signal
/// (a GetConsoleWindow-style check alone always reports NULL here, since a
/// GUI-subsystem process starts detached). AttachConsole/AllocConsole are
/// NT+ only and must stay out of the Win95 (rust9x) builds, hence the gate
/// (raw declarations, no new dependencies).
#[cfg(all(target_os = "windows", not(target_family = "rust9x")))]
mod wincon {
    #[link(name = "kernel32")]
    extern "system" {
        fn AttachConsole(dwProcessId: u32) -> i32;
        fn FreeConsole() -> i32;
        fn AllocConsole() -> i32;
        fn GetStdHandle(nStdHandle: i32) -> *mut std::ffi::c_void;
        fn GetConsoleMode(
            hConsoleHandle: *mut std::ffi::c_void,
            lpMode: *mut u32,
        ) -> i32;
    }
    pub const ATTACH_PARENT_PROCESS: u32 = 0xFFFF_FFFF;
    const STD_INPUT_HANDLE: i32 = -10;
    const STD_OUTPUT_HANDLE: i32 = -11;

    /// Claim the parent's console. True = launched from a console
    /// (Explorer/tile/Run launches have no parent console: false).
    pub fn attach_parent() -> bool {
        unsafe { AttachConsole(ATTACH_PARENT_PROCESS) != 0 }
    }
    pub fn detach() {
        unsafe {
            FreeConsole();
        }
    }
    /// Pop a fresh console window (tile/Run launches with explicit TUI).
    pub fn alloc() -> bool {
        unsafe { AllocConsole() != 0 }
    }
    fn handle_is_console(std: i32) -> bool {
        unsafe {
            let h = GetStdHandle(std);
            if h.is_null() {
                return false;
            }
            let mut mode = 0u32;
            GetConsoleMode(h, &mut mode) != 0
        }
    }
    /// stdio are live console handles (not redirected to file/pipe/nul).
    /// Guards scripting: `gcorro --help > out.txt` from a console must not
    /// flip into the TUI just because a parent console exists.
    pub fn stdio_is_console() -> bool {
        handle_is_console(STD_INPUT_HANDLE) && handle_is_console(STD_OUTPUT_HANDLE)
    }
}

/// Headless-launch fallback for Unix desktops (pure; inputs are parameters
/// so the table is unit-testable). A terminal UI with stdio on /dev/null
/// (double-clicked in a file manager, launched without a terminal) can only
/// die, so with no explicit --flag, no argv[0] request, and no terminal on
/// stdio, a GUI-capable build opens the GUI instead. Explicit choices
/// (flags and argv[0] names) are honored untouched.
fn resolve_headless_default(
    ui: UiKind,
    explicit: bool,
    argv0_matched: bool,
    has_tty: bool,
    gui_fallback: Option<UiKind>,
) -> UiKind {
    if explicit || argv0_matched || has_tty {
        return ui;
    }
    gui_fallback.unwrap_or(ui)
}

/// Console-vs-GUI default resolution (pure: FFI inputs are parameters so the
/// table is unit-testable on any host). Returns (ui, keep_attached).
///
/// - An explicit --gui/--ratatui/--pancurses flag always wins.
/// - Otherwise, launched-in-a-console with live console stdio defaults to
///   the terminal UI (`tui_default`: caller's first compiled-in TUI).
/// - Otherwise the argv[0]/compiled default stands; an unneeded attachment
///   is released (second element false = caller should FreeConsole).
fn resolve_console_default(
    ui: UiKind,
    explicit: bool,
    attached: bool,
    stdio_console: bool,
    tui_default: Option<UiKind>,
) -> (UiKind, bool) {
    if explicit {
        // Caller ensures a console exists for an explicit TUI request.
        return (ui, attached);
    }
    if attached && stdio_console {
        if let Some(tui) = tui_default {
            return (tui, true);
        }
    }
    (ui, false)
}

#[cfg(not(all(target_family = "rust9x", target_env = "msvc")))]
fn cli_args() -> impl Iterator<Item = String> {
    std::env::args().skip(1)
}

#[cfg(all(target_family = "rust9x", target_env = "msvc"))]
fn win95_args() -> Vec<String> {
    #[link(name = "kernel32")]
    unsafe extern "system" {
        fn GetCommandLineA() -> *const u8;
    }
    unsafe {
        let p = GetCommandLineA();
        if p.is_null() {
            return Vec::new();
        }
        let mut raw = Vec::new();
        let mut i = 0usize;
        while *p.add(i) != 0 {
            raw.push(*p.add(i));
            i += 1;
        }
        let mut out: Vec<String> = Vec::new();
        let mut cur = String::new();
        let mut in_quotes = false;
        for &b in &raw {
            match b {
                b'"' => in_quotes = !in_quotes,
                b' ' | b'\t' if !in_quotes => {
                    if !cur.is_empty() {
                        out.push(std::mem::take(&mut cur));
                    }
                }
                _ => cur.push(b as char),
            }
        }
        if !cur.is_empty() {
            out.push(cur);
        }
        out
    }
}

fn parse_args() -> Result<Args, String> {
    let mut revision = None;
    let mut export = None;
    let mut movie = false;
    let movie_typing_cps = 22.0f64;
    let movie_confirm_ms = 120u64;
    let movie_menu_hold_ms = 1200u64;
    let mut show_help = false;
    let mut show_version = false;
    let debug_no_number = false;
    // Whether --gui/--ratatui/--pancurses explicitly chose the UI (always
    // wins over argv[0] and over the Windows console-launch default below).
    // Modern-Windows-only input: other targets never read it.
    #[cfg(any(all(target_os = "windows", not(target_family = "rust9x")), all(target_family = "unix", not(target_arch = "wasm32"))))]
    let mut ui_explicit = false;
    // (argv0_matched records whether the name itself requested a UI; the
    // Unix headless fallback honors a name match like an explicit choice.)
    // argv[0] dispatch comes first so explicit flags below can override
    // it; a non-matching or uncompiled name falls back to the default.
    // argv0 outlives this block: the Unix headless fallback below re-reads
    // it to honor a name match like an explicit choice.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    let argv0 = win95_args().into_iter().next();
    #[cfg(not(all(target_family = "rust9x", target_env = "msvc")))]
    let argv0 = std::env::args().next();
    let mut ui = {
        argv0
            .as_deref()
            .and_then(argv0_ui)
            .unwrap_or_else(determine_default_ui)
    };
    let mut capture_html = None;
    let mut convert_ansi = None;
    let mut positional = Vec::new();
    let mut it = cli_args().peekable();

    while let Some(arg) = it.next() {
        match arg.as_str() {
            "-h" | "-?" | "--help" => {
                show_help = true;
            }
            "-v" | "--version" => {
                show_version = true;
            }
            "-r" | "--revision" => {
                if let Some(next) = it.peek() {
                    if let Ok(value) = next.parse::<usize>() {
                        let _ = it.next();
                        revision = Some(RevisionMode::Limit(value));
                        continue;
                    }
                }
                revision = Some(RevisionMode::Browse);
            }
            "-e" | "--export" => {
                let Some(path) = it.next() else {
                    return Err("--export requires a file path".into());
                };
                export = Some(PathBuf::from(path));
            }
            "--movie" => {
                movie = true;
            }
            "--ratatui" => {
                #[cfg(any(all(target_os = "windows", not(target_family = "rust9x")), all(target_family = "unix", not(target_arch = "wasm32"))))]
                { ui_explicit = true; }
                #[cfg(feature = "ratatui")]
                { ui = UiKind::Ratatui; }
                #[cfg(not(feature = "ratatui"))]
                { return Err("ratatui UI not compiled in; rebuild with --features ratatui".into()); }
            }
            "--gui" => {
                #[cfg(any(all(target_os = "windows", not(target_family = "rust9x")), all(target_family = "unix", not(target_arch = "wasm32"))))]
                { ui_explicit = true; }
                #[cfg(feature = "gui")]
                { ui = UiKind::Gui; }
                #[cfg(not(feature = "gui"))]
                { return Err("GTK GUI not compiled in; rebuild with --features gui".into()); }
            }
            "--pancurses" => {
                #[cfg(any(all(target_os = "windows", not(target_family = "rust9x")), all(target_family = "unix", not(target_arch = "wasm32"))))]
                { ui_explicit = true; }
                #[cfg(feature = "pancurses")]
                { ui = UiKind::Pancurses; }
                #[cfg(not(feature = "pancurses"))]
                { return Err("pancurses UI not compiled in; rebuild with --features pancurses".into()); }
            }
            "--capture-html" => {
                let Some(path) = it.next() else {
                    return Err("--capture-html requires a file path".into());
                };
                capture_html = Some(PathBuf::from(path));
            }
            "--convert-ansi" => {
                let Some(path) = it.next() else {
                    return Err("--convert-ansi requires an input file path".into());
                };
                convert_ansi = Some(PathBuf::from(path));
            }
            _ => positional.push(arg),
        }
    }

    let files = positional.into_iter().map(PathBuf::from).collect();

    // Unix headless-launch fallback (no attach step exists here: stdio is
    // always inherited, so a missing terminal means stdio is /dev/null).
    // Double-clicked from a file manager with no UI choice, a GUI-capable
    // build opens the GUI instead of dying on dead stdio. Explicit flags
    // and argv[0] names are honored untouched.
    #[cfg(all(target_family = "unix", not(target_arch = "wasm32")))]
    {
        #[cfg(feature = "gui")]
        let gui_fallback = Some(UiKind::Gui);
        #[cfg(not(feature = "gui"))]
        let gui_fallback: Option<UiKind> = None;
        // SAFETY: isatty takes a raw fd, no retained state; -1 impossible
        // here (constants), return is a plain 0/1 boolean.
        let has_tty = unsafe {
            libc::isatty(libc::STDIN_FILENO) != 0 && libc::isatty(libc::STDOUT_FILENO) != 0
        };
        let argv0_matched = argv0.as_deref().and_then(argv0_ui).is_some();
        ui = resolve_headless_default(ui, ui_explicit, argv0_matched, has_tty, gui_fallback);
    }

    // Windows console handling (modern Windows only; Win95 builds are
    // separate console/GUI exes and never reach this). A GUI-subsystem
    // launch (e.g. the MSIX build) starts detached: without this block a
    // TUI request fails silently on dead console handles.
    #[cfg(all(target_os = "windows", not(target_family = "rust9x")))]
    {
        #[cfg(feature = "ratatui")]
        let tui_default = Some(UiKind::Ratatui);
        #[cfg(all(not(feature = "ratatui"), feature = "pancurses"))]
        let tui_default = Some(UiKind::Pancurses);
        #[cfg(all(not(feature = "ratatui"), not(feature = "pancurses")))]
        let tui_default: Option<UiKind> = None;
        if ui_explicit && matches!(ui, UiKind::Ratatui | UiKind::Pancurses) {
            // Explicit TUI: guarantee a console. From a terminal the
            // attach makes console APIs work; with no parent console
            // (tile/Run dialog) pop a fresh window instead of dying
            // silently on dead handles.
            if !(wincon::attach_parent() || wincon::alloc()) {
                return Err("no console available for the terminal UI".into());
            }
        } else if !ui_explicit {
            // No explicit choice: a claimed parent console with live
            // console stdio means "launched in a console" -> default to
            // the TUI instead of popping a GUI window. Explorer/tile
            // launches (no parent console) and redirected stdio
            // (scripting) keep the argv[0]/compiled default.
            let attached = wincon::attach_parent();
            let (new_ui, keep) = resolve_console_default(
                ui,
                false,
                attached,
                if attached {
                    wincon::stdio_is_console()
                } else {
                    false
                },
                tui_default,
            );
            ui = new_ui;
            if attached && !keep {
                wincon::detach();
            }
        }
        // Explicit --gui (or anything else explicit): untouched, detached.
    }

    Ok(Args {
        revision,
        files,
        export,
        movie,
        movie_typing_cps,
        movie_confirm_ms,
        movie_menu_hold_ms,
        show_help,
        show_version,
        debug_no_number,
        ui,
        capture_html,
        convert_ansi,
    })
}

// Win95 entry point (rust9x msvc targets only; see the no_main note above).
// The VC6 CRT startup calls this directly. Errors still exit non-zero via
// std::process::exit inside corro_main.
// NOTE: `not(feature = "gui-subsystem")` matters. The VC6 CRT ships two
// startup objects: crt0.obj defines `_main`+`_mainCRTStartup` (console) and
// wincrt0.obj defines `_WinMain@16`+`_WinMainCRTStartup` (GUI). Defining BOTH
// `main` and `WinMain` makes the linker pull BOTH objects, which then collide
// on the shared CRT globals (__amsg_exit, __aexit_rtn, ...) — "duplicate
// symbol". Each build therefore exports exactly one entry point.
#[cfg(all(
    target_family = "rust9x",
    target_env = "msvc",
    not(feature = "gui-subsystem")
))]
#[no_mangle]
pub extern "C" fn main() -> i32 {
    win9x_redirect_console_output();
    corro_main();
    0
}

// GUI-subsystem entry point for the nwg build (rust9x msvc, `gui` feature).
//
// A PE linked with /SUBSYSTEM:WINDOWS is entered through the CRT's
// `WinMainCRTStartup`, which calls `WinMain` — *not* `main`. Without this the
// nwg build would link against the GUI subsystem and then die before running
// a single line. The signature is the classic one; the parameters are unused
// because corro takes its input from the command line (fetched through the
// same `win95_args` shim `main` uses) and from argv[0] dispatch.
//
// This is compiled only for the GUI link, so the console TUI build keeps its
// `main` entry and console subsystem (see build95.sh / build95-gui.sh).
#[cfg(all(
    target_family = "rust9x",
    target_env = "msvc",
    feature = "gui",
    feature = "gui-subsystem"
))]
#[no_mangle]
pub extern "system" fn WinMain(
    _instance: *mut std::os::raw::c_void,
    _prev_instance: *mut std::os::raw::c_void,
    _cmd_line: *mut u8,
    _show_cmd: i32,
) -> i32 {
    // TEMPORARY Win95 diagnosis: startup progression markers.
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    unsafe {
        mark95(b"winmain\n");
    }
    // TEMPORARY Win95 diagnosis: route panic messages to the log file so the
    // aborting unwrap/expect identifies itself (GUI has no console).
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    win9x_redirect_console_output();
    // TEMPORARY Win95 diagnosis: raw panic hook (default hook dies in TLS).
    #[cfg(all(target_family = "rust9x", target_env = "msvc"))]
    install_raw_panic_hook95();
    corro_main();
    0
}

/// TEMPORARY Win95 diagnosis: panic hook that records the panic site via raw
/// CreateFileA (no std::fs, no TLS, no heap-alloc in the write path — the
/// default hook dies in thread-local storage on 9x, masking the real site).
#[cfg(all(target_family = "rust9x", target_env = "msvc"))]
fn install_raw_panic_hook95() {
    use std::os::raw::c_void;
    struct SliceW<'a> { b: &'a mut [u8], p: usize }
    impl<'a> core::fmt::Write for SliceW<'a> {
        fn write_str(&mut self, s: &str) -> core::fmt::Result {
            let n = s.len().min(self.b.len().saturating_sub(self.p));
            self.b[self.p..self.p + n].copy_from_slice(&s.as_bytes()[..n]);
            self.p += n;
            Ok(())
        }
    }
    std::panic::set_hook(Box::new(|info| {
        unsafe extern "system" {
            fn CreateFileA(name: *const u8, access: u32, share: u32, sa: *mut c_void,
                disp: u32, flags: u32, tmpl: *mut c_void) -> *mut c_void;
            fn SetFilePointer(h: *mut c_void, lo: i32, hi: *mut i32, how: u32) -> u32;
            fn WriteFile(h: *mut c_void, buf: *const u8, len: u32, w: *mut u32, ov: *mut c_void) -> i32;
            fn CloseHandle(h: *mut c_void) -> i32;
        }
        let mut buf = [0u8; 768];
        let mut w = SliceW { b: &mut buf, p: 0 };
        {
            use core::fmt::Write as _;
            let _ = write!(w, "PANIC ");
            if let Some(loc) = info.location() {
                let _ = write!(w, "{}:{}:{} ", loc.file(), loc.line(), loc.column());
            }
            // (rust9x std predates PanicHookInfo::message; use payload.)
            let pl = info.payload();
            if let Some(s) = pl.downcast_ref::<&str>() {
                let _ = write!(w, "{}", s);
            } else if let Some(s) = pl.downcast_ref::<String>() {
                let _ = write!(w, "{}", s);
            } else {
                let _ = write!(w, "<non-string payload>");
            }
            let _ = write!(w, "\n");
        }
        let n = w.p;
        unsafe {
            let h = CreateFileA(b"c:\\panic95.log\0".as_ptr(), 0x4000_0000, 1,
                std::ptr::null_mut(), 4, 0x80, std::ptr::null_mut());
            if !h.is_null() && h as isize != -1 {
                SetFilePointer(h, 0, std::ptr::null_mut(), 2);
                let mut wr = 0u32;
                WriteFile(h, buf.as_ptr(), n as u32, &mut wr, std::ptr::null_mut());
                CloseHandle(h);
            }
        }
    }));
}

/// TEMPORARY Win95 diagnosis: append bytes to c:\gcorro.log via raw
/// CreateFileA (std::fs is broken on 9x: CreateFileW stub, error 120).
#[cfg(all(target_family = "rust9x", target_env = "msvc"))]
unsafe fn mark95(s: &[u8]) {
    use std::os::raw::c_void;
    unsafe extern "system" {
        fn CreateFileA(name: *const u8, access: u32, share: u32, sa: *mut c_void,
            disp: u32, flags: u32, tmpl: *mut c_void) -> *mut c_void;
        fn SetFilePointer(h: *mut c_void, lo: i32, hi: *mut i32, how: u32) -> u32;
        fn WriteFile(h: *mut c_void, buf: *const u8, len: u32, w: *mut u32, ov: *mut c_void) -> i32;
        fn CloseHandle(h: *mut c_void) -> i32;
    }
    let h = CreateFileA(b"c:\\gcorro.log\0".as_ptr(), 0x4000_0000, 1,
        std::ptr::null_mut(), 4, 0x80, std::ptr::null_mut());
    if h.is_null() || h as isize == -1 {
        return;
    }
    SetFilePointer(h, 0, std::ptr::null_mut(), 2);
    let mut w = 0u32;
    WriteFile(h, s.as_ptr(), s.len() as u32, &mut w, std::ptr::null_mut());
    CloseHandle(h);
}

// Windows 9x only: std's console layer writes through WriteConsoleW, a no-op
// stub on Windows 95/98/ME — every eprintln!/println! panics ("failed
// printing to stderr") and aborts under panic="abort". Redirect
// STDOUT/ERROR to a log file with SetStdHandle so all std console macros
// write via WriteFile instead. The NT line (Win2000+) has real W console
// APIs and keeps its console handles (the crossterm TUI needs them there).
#[cfg(all(target_family = "rust9x", target_env = "msvc"))]
fn win9x_redirect_console_output() {
    use std::os::raw::c_void;
    #[repr(C)]
    struct OsVersionInfoA {
        size: u32,
        major: u32,
        minor: u32,
        build: u32,
        platform: u32,
        csd: [u8; 128],
    }
    unsafe extern "system" {
        fn GetVersionExA(info: *mut c_void) -> i32;
        fn CreateFileA(
            name: *const u8,
            access: u32,
            share: u32,
            sa: *mut c_void,
            disp: u32,
            flags: u32,
            tmpl: *mut c_void,
        ) -> *mut c_void;
        fn SetStdHandle(which: u32, handle: *mut c_void) -> i32;
    }
    const VER_PLATFORM_WIN32_WINDOWS: u32 = 1; // 95/98/ME
    const GENERIC_WRITE: u32 = 0x4000_0000;
    const FILE_SHARE_READ: u32 = 1;
    const OPEN_ALWAYS: u32 = 4;
    const FILE_ATTRIBUTE_NORMAL: u32 = 0x80;
    const STD_OUTPUT_HANDLE: u32 = -11i32 as u32;
    const STD_ERROR_HANDLE: u32 = -12i32 as u32;
    unsafe {
        let mut vi = OsVersionInfoA {
            size: std::mem::size_of::<OsVersionInfoA>() as u32,
            major: 0,
            minor: 0,
            build: 0,
            platform: 0,
            csd: [0; 128],
        };
        if GetVersionExA(&mut vi as *mut _ as *mut c_void) == 0
            || vi.platform != VER_PLATFORM_WIN32_WINDOWS
        {
            return; // NT line: keep the real console handles
        }
        let h = CreateFileA(
            b"c:\\corro-win95.log\0".as_ptr(),
            GENERIC_WRITE,
            FILE_SHARE_READ,
            std::ptr::null_mut(),
            OPEN_ALWAYS,
            FILE_ATTRIBUTE_NORMAL,
            std::ptr::null_mut(),
        );
        if h.is_null() || h as isize == -1 {
            return;
        }
        SetStdHandle(STD_OUTPUT_HANDLE, h);
        SetStdHandle(STD_ERROR_HANDLE, h);
    }
}

#[cfg(target_arch = "wasm32")]
fn main() {
    let mut app = if let Some(path) = std::env::args().nth(1) {
        let mut a = corro::gui::App::new_with_paths(vec![std::path::PathBuf::from(path)]);
        if let Err(e) = a.load_initial() {
            eprintln!("corro: load error: {e}");
        }
        a
    } else {
        let mut a = corro::gui::App::new_with_paths(vec![]);
        let _ = a.load_initial();
        a
    };
    // Leak the App so its memory stays valid after main() returns.
    // The GuiState holds a raw pointer to this App; DOM event listeners
    // registered by the WASM adapter continue to fire after main() exits
    // because their closures are stored in a global static (CLOSURES).
    // If we let the App drop, the raw pointer becomes dangling and keyboard
    // callbacks would access freed memory (undefined behavior).
    let app: &'static mut corro::gui::App = Box::leak(Box::new(app));
    if let Err(e) = app.run() {
        eprintln!("corro: run error: {e}");
    }
}

#[cfg(all(
    not(target_arch = "wasm32"),
    not(all(target_family = "rust9x", target_env = "msvc"))
))]
fn main() {
    corro_main();
}

fn corro_main() {
    // Log every panic to a file as well as stderr: release profiles use
    // panic="abort" (silent death, no message), and GUI launches often
    // detach stderr, so an undebuggable abort is otherwise all you get.
    // /tmp/corro-panic.log then names the panic site (e.g. which key
    // sequence tripped it) instead of leaving only "Aborted (core dumped)".
    let default_hook = std::panic::take_hook();
    std::panic::set_hook(Box::new(move |info| {
        let msg = format!("corro panic: {info}\n{}", std::backtrace::Backtrace::force_capture());
        let _ = std::fs::write("/tmp/corro-panic.log", &msg);
        default_hook(info);
    }));
    let (res, exit_message) = try_main();
    if let Some(msg) = exit_message {
        // Print to both stderr and stdout and flush so the message is
        // visible after the TUI restores the terminal. Also write a
        // fallback file under XDG_STATE_HOME/corro/last-exit-hint or
        // ~/.corro/last-exit-hint so the message can be discovered when
        // terminal output is unreliable.
        use std::io::Write as _;
        let _ = writeln!(std::io::stderr(), "{}", msg);
        let _ = writeln!(std::io::stdout(), "{}", msg);
        let _ = std::io::stderr().flush();
        let _ = std::io::stdout().flush();

        // Fallback file write.
        if let Ok(xdg) = std::env::var("XDG_STATE_HOME") {
            let mut dir = std::path::PathBuf::from(xdg);
            dir.push("corro");
                if std::fs::create_dir_all(&dir).is_ok() {
                let path = dir.join("last-exit-hint");
                let _ = std::fs::write(path, &msg);
            }
        } else if let Ok(home) = std::env::var("HOME") {
            let mut dir = std::path::PathBuf::from(home);
            dir.push(".corro");
            if std::fs::create_dir_all(&dir).is_ok() {
                let path = dir.join("last-exit-hint");
                let _ = std::fs::write(path, &msg);
            }
        }

        // Also write an exit hint to the debug log if possible. Prefer
        // CORRO_DEBUG_LOG, otherwise XDG_STATE_HOME/corro/debug.log or
        // ~/.corro/debug.log. Ignore errors; this is best-effort only.
        if let Some(path) = std::env::var("CORRO_DEBUG_LOG").ok().or_else(|| {
            std::env::var("XDG_STATE_HOME").ok().map(|xdg| format!("{}/corro/debug.log", xdg))
        }).or_else(|| std::env::var("HOME").ok().map(|h| format!("{}/.corro/debug.log", h))) {
            let p = std::path::PathBuf::from(path);
            if let Some(dir) = p.parent() {
                let _ = std::fs::create_dir_all(dir);
            }
            let _ = std::fs::OpenOptions::new().create(true).append(true).open(&p).and_then(|mut f| {
                use std::io::Write as _;
                writeln!(f, "{}", msg)
            });
        }
    }
    if let Err(e) = res {
        eprintln!("{e}");
        std::process::exit(1);
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn try_main() -> (Result<(), Box<dyn std::error::Error>>, Option<String>) {
    // Parse args; return early with no exit message on CLI errors/help/version.
    let args = match parse_args() {
        Ok(a) => a,
        Err(s) => return (Err(s.into()), None),
    };

    // Redirect stderr to a per-user debug log so debug traces do not
    // interleave with the TUI. Prefer CORRO_DEBUG_LOG if set; otherwise
    // use XDG_STATE_HOME/corro/debug.log or ~/.corro/debug.log. Attempt to
    // open/create the log and duplicate it onto STDERR so existing
    // eprintln! calls go to the file on Unix platforms. Ignore errors.
    if let Some(path) = std::env::var("CORRO_DEBUG_LOG").ok().or_else(|| {
        std::env::var("XDG_STATE_HOME").ok().map(|xdg| format!("{}/corro/debug.log", xdg))
    }).or_else(|| std::env::var("HOME").ok().map(|h| format!("{}/.corro/debug.log", h))) {
        let p = std::path::PathBuf::from(path);
        if let Some(dir) = p.parent() {
            let _ = std::fs::create_dir_all(dir);
        }
        if let Ok(f) = std::fs::OpenOptions::new().create(true).append(true).open(&p) {
            #[cfg(unix)]
            {
                use std::os::unix::io::AsRawFd;
                // Duplicate the debug file onto STDERR_FILENO so existing
                // eprintln! calls write to the file.
                unsafe {
                    libc::dup2(f.as_raw_fd(), libc::STDERR_FILENO);
                }
                // Prevent the File's Drop from closing the fd we just
                // duplicated; leak it intentionally until process exit.
                let _ = Box::leak(Box::new(f));
            }
            #[cfg(not(unix))]
            {
                // On non-Unix platforms, just keep the file open but do not
                // attempt to replace STDERR; callers may still see eprintln
                // output on the terminal.
                let _ = f;
            }
        }
    }
    // If requested via CLI, expose the debug-no-number flag as an env var so
    // debug instrumentation in other modules can observe it without changing
    // many function signatures.
    if args.debug_no_number {
        let _ = std::env::set_var("CORRO_DEBUG_NO_NUMBER", "1");
    }
    if args.show_help {
        println!("{}", cli_help_text());
        return (Ok(()), None);
    }
    if args.show_version {
        println!("corro {}", env!("CARGO_PKG_VERSION"));
        return (Ok(()), None);
    }
    if let Some(ref ansi_path) = args.convert_ansi {
        let out_path = args.files.first().cloned().unwrap_or_else(|| PathBuf::from("output.html"));
        match corro::capture::convert_ansi_file(ansi_path, &out_path) {
            Ok(()) => println!("Converted {} to {}", ansi_path.display(), out_path.display()),
            Err(e) => eprintln!("Conversion failed: {e}"),
        }
        return (Ok(()), None);
    }
    if let Some(export_path) = args.export {
            let input_path = match args.files.first() {
                Some(p) => p.clone(),
                None => {
                    return (
                        Err("--export requires an input file argument".into()),
                        None,
                    )
                }
            };
            if args.files.len() > 1 {
                return (
                    Err("--export accepts exactly one input file".into()),
                    None,
                );
            }
            let workbook = match load_workbook_for_export(&input_path) {
                Ok(w) => w,
                Err(e) => return (Err(e.into()), None),
            };
            if let Err(e) = export_workbook_to_path(&workbook, &export_path) {
                return (Err(e.into()), None);
            }
            return (Ok(()), None);
        }
        if args.movie && args.revision.is_some() {
            return (
                Err("--movie cannot be combined with --revision".into()),
                None,
            );
        }
        if args.revision.is_some() && args.files.len() > 1 {
            return (
                Err("--revision accepts exactly one input file".into()),
                None,
            );
        }
        if args.movie && args.files.len() > 1 {
            return (
                Err("--movie accepts exactly one input file".into()),
                None,
            );
        }
    let (res, exit_msg) = match args.ui {
        #[cfg(feature = "ratatui")]
        UiKind::Ratatui => {
            let capture = match args.capture_html.as_ref() {
                Some(p) => match corro::capture::HtmlCapture::new(p) {
                    Ok(c) => Some(c),
                    Err(e) => return (Err(e.into()), None),
                },
                None => None,
            };
            let mut app = match args.revision {
                None => TuiApp::new_with_paths(args.files),
                Some(RevisionMode::Browse) => TuiApp::new_with_revision_browser(args.files.first().cloned()),
                Some(RevisionMode::Limit(revision)) => {
                    TuiApp::new_with_revision_limit(args.files.first().cloned(), Some(revision))
                }
            };
            app.set_capturer(capture);
            let res = if args.movie {
                app.run_movie(corro::ui::MovieReplayOptions {
                    typing_cps: args.movie_typing_cps,
                    confirm_delay_ms: args.movie_confirm_ms,
                    menu_hold_ms: args.movie_menu_hold_ms,
                })
            } else {
                match app.load_initial() {
                    Ok(()) => app.run(),
                    Err(e) => Err(e.into()),
                }
            };
            let exit_msg = app.take_final_exit_hint();
            (res.map_err(|e| e.into()), exit_msg)
        }
        #[cfg(feature = "gui")]
        UiKind::Gui => {
            let mut app = match args.revision {
                None => GuiApp::new_with_paths(args.files),
                Some(RevisionMode::Browse) => GuiApp::new_with_revision_browser(args.files.first().cloned()),
                Some(RevisionMode::Limit(revision)) => {
                    GuiApp::new_with_revision_limit(args.files.first().cloned(), Some(revision))
                }
            };
            app.set_backend(corro::gui::Backend::Gui);
            let res = match app.load_initial() {
                Ok(()) => app.run(),
                Err(e) => Err(e),
            };
            let exit_msg = app.take_final_exit_hint();
            (res, exit_msg)
        }
        #[cfg(feature = "pancurses")]
        UiKind::Pancurses => {
            let mut app = match args.revision {
                None => GuiApp::new_with_paths(args.files),
                Some(RevisionMode::Browse) => GuiApp::new_with_revision_browser(args.files.first().cloned()),
                Some(RevisionMode::Limit(revision)) => {
                    GuiApp::new_with_revision_limit(args.files.first().cloned(), Some(revision))
                }
            };
            #[cfg(feature = "pancurses")]
            app.set_backend(corro::gui::Backend::Pancurses);
            let res = match app.load_initial() {
                Ok(()) => app.run(),
                Err(e) => Err(e),
            };
            let exit_msg = app.take_final_exit_hint();
            (res, exit_msg)
        }
        #[allow(unreachable_patterns)]
        _ => {
            let msg = format!("{:?} UI backend not compiled in", args.ui);
            (Err(msg.into()), None)
        }
    };
    (res, exit_msg)
}

fn cli_help_text() -> String {
    let mut ui_opts = String::new();
    #[cfg(feature = "ratatui")]
    { ui_opts.push_str("  --ratatui                Use ratatui terminal UI (default)\n"); }
    #[cfg(feature = "gui")]
    { ui_opts.push_str("  --gui                    Use GTK native GUI\n"); }
    #[cfg(feature = "pancurses")]
    { ui_opts.push_str("  --pancurses              Use pancurses terminal UI\n"); }
    #[cfg(feature = "pancurses")]
    { ui_opts.push_str("  (invoked as pcorro* defaults to pancurses)\n"); }
    #[cfg(feature = "gui")]
    { ui_opts.push_str("  (invoked as gcorro* defaults to the GUI)\n"); }
    #[cfg(all(target_os = "windows", not(target_family = "rust9x")))]
    { ui_opts.push_str("  (a console launch with no UI flag defaults to the terminal UI)\n"); }
    #[cfg(all(target_family = "unix", not(target_arch = "wasm32")))]
    { ui_opts.push_str("  (a terminal-less launch with no UI flag defaults to the GUI)\n"); }
    format!(
        "corro {}\n\
\n\
USAGE:\n\
  corro [OPTIONS] [FILE ...]\n\
\n\
OPTIONS:\n\
  -h, -?, --help            Show help\n\
  -v, --version             Show version\n\
  -r, --revision [N]        Browse revisions (or limit to N)\n\
  -e, --export <PATH>       Export input FILE to PATH (.tsv, .csv, .txt/.ascii, .ods)\n\
  --movie                   Replay a .corro file line-by-line, then quit\n\
  --movie-typing-cps <N>    Movie typing speed in chars/sec (default: 22)\n\
  --movie-confirm-ms <N>    Delay before Enter/confirm per line (default: 120)\n\
  --movie-menu-hold-ms <N>  Hold menu/dialog moments in movie mode (default: 1200)\n\
{}",
        env!("CARGO_PKG_VERSION"),
        ui_opts,
    )
}

fn load_workbook_for_export(path: &std::path::Path) -> Result<corro::ops::WorkbookState, String> {
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_ascii_lowercase();
    match ext.as_str() {
        "corro" => {
            let mut workbook = corro::ops::WorkbookState::new();
            let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
            let _ = corro::io::load_workbook_revisions(path, usize::MAX, &mut workbook, &mut active_sheet)
                .map_err(|e| format!("failed to read .corro workbook: {e}"))?;
            if let Some(i) = workbook.sheets.iter().position(|s| s.id == active_sheet) {
                workbook.active_sheet = i;
            }
            Ok(workbook)
        }
        "ods" => corro::ods::import_ods_workbook(path)
            .map_err(|e| format!("failed to import ODS: {e}")),
        "tsv" => {
            let data = std::fs::read_to_string(path)
                .map_err(|e| format!("failed to read TSV: {e}"))?;
            let mut workbook = corro::ops::WorkbookState::new();
            let state = workbook.active_sheet_mut();
            corro::io::import_tsv(&data, state);
            Ok(workbook)
        }
        "csv" => {
            let data = std::fs::read_to_string(path)
                .map_err(|e| format!("failed to read CSV: {e}"))?;
            let mut workbook = corro::ops::WorkbookState::new();
            let state = workbook.active_sheet_mut();
            corro::io::import_csv(&data, state);
            Ok(workbook)
        }
        _ => Err(format!(
            "unsupported input extension: {} (expected .corro, .ods, .tsv, .csv)",
            if ext.is_empty() { "<none>" } else { ext.as_str() }
        )),
    }
}

fn export_workbook_to_path(
    workbook: &corro::ops::WorkbookState,
    path: &std::path::Path,
) -> Result<(), String> {
    let ext = path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_ascii_lowercase();
    match ext.as_str() {
        "tsv" => {
            let mut buf = Vec::new();
            corro::export::export_tsv_with_options(
                &workbook.active_sheet().grid,
                &mut buf,
                &corro::export::DelimitedExportOptions::default(),
            );
            std::fs::write(path, buf).map_err(|e| format!("failed to write TSV: {e}"))
        }
        "csv" => {
            let mut buf = Vec::new();
            corro::export::export_csv_with_options(
                &workbook.active_sheet().grid,
                &mut buf,
                &corro::export::DelimitedExportOptions::default(),
            );
            std::fs::write(path, buf).map_err(|e| format!("failed to write CSV: {e}"))
        }
        "txt" | "ascii" => {
            let mut buf = Vec::new();
            corro::export::export_ascii_table_with_options(
                &workbook.active_sheet().grid,
                &mut buf,
                &corro::export::AsciiTableOptions::default(),
            );
            std::fs::write(path, buf).map_err(|e| format!("failed to write ASCII text: {e}"))
        }
        "ods" => {
            let ods_options = corro::export::DelimitedExportOptions {
                content: corro::export::ExportContent::Generic,
                ..Default::default()
            };
            let bytes = corro::ods::export_ods_bytes_workbook_with_options(
                workbook,
                &ods_options,
            )
            .map_err(|e| format!("failed to export ODS: {e}"))?;
            std::fs::write(path, bytes).map_err(|e| format!("failed to write ODS: {e}"))
        }
        _ => Err(format!(
            "unsupported export extension: {} (expected .tsv, .csv, .txt, .ascii, .ods)",
            if ext.is_empty() { "<none>" } else { ext.as_str() }
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::{resolve_console_default, resolve_headless_default, Args, RevisionMode, UiKind};
    use std::path::PathBuf;

    // resolve_console_default decision table (pure; the AttachConsole /
    // GetConsoleMode inputs are parameters so this runs on any host).
    #[test]
    fn console_launch_defaults_to_tui() {
        // gcorro-named, no flag, parent console + live stdio -> TUI.
        let (ui, keep) =
            resolve_console_default(UiKind::Gui, false, true, true, Some(UiKind::Ratatui));
        assert_eq!(ui, UiKind::Ratatui);
        assert!(keep);
    }

    #[test]
    fn explicit_flag_always_wins() {
        // --gui from a console stays GUI (but keeps the attachment).
        let (ui, keep) =
            resolve_console_default(UiKind::Gui, true, true, true, Some(UiKind::Ratatui));
        assert_eq!(ui, UiKind::Gui);
        assert!(keep);
        // Explicit TUI from a tile keeps the request; caller allocs.
        let (ui, _) =
            resolve_console_default(UiKind::Ratatui, true, false, false, Some(UiKind::Ratatui));
        assert_eq!(ui, UiKind::Ratatui);
    }

    #[test]
    fn tile_and_redirected_launches_stay_gui() {
        // No parent console (tile/Explorer): GUI, nothing to release.
        let (ui, keep) =
            resolve_console_default(UiKind::Gui, false, false, false, Some(UiKind::Ratatui));
        assert_eq!(ui, UiKind::Gui);
        assert!(!keep);
        // Parent console exists but stdio redirected (scripting): GUI and
        // release the unneeded attachment.
        let (ui, keep) =
            resolve_console_default(UiKind::Gui, false, true, false, Some(UiKind::Ratatui));
        assert_eq!(ui, UiKind::Gui);
        assert!(!keep);
    }

    #[test]
    fn no_tui_compiled_keeps_gui() {
        // Pure-GUI build in a console: nothing to switch to, detach.
        let (ui, keep) = resolve_console_default(UiKind::Gui, false, true, true, None);
        assert_eq!(ui, UiKind::Gui);
        assert!(!keep);
    }

    #[test]
    fn pancurses_fallback_when_ratatui_absent() {
        let (ui, keep) =
            resolve_console_default(UiKind::Gui, false, true, true, Some(UiKind::Pancurses));
        assert_eq!(ui, UiKind::Pancurses);
        assert!(keep);
    }

    // resolve_headless_default table (Unix desktops; also pure).
    #[test]
    fn headless_launch_defaults_to_gui() {
        // No flag, neutral name, no tty, GUI compiled -> GUI.
        assert_eq!(
            resolve_headless_default(UiKind::Ratatui, false, false, false, Some(UiKind::Gui)),
            UiKind::Gui
        );
    }

    #[test]
    fn headless_honors_explicit_choices() {
        // Explicit --ratatui with piped stdio stays a TUI attempt.
        assert_eq!(
            resolve_headless_default(UiKind::Ratatui, true, false, false, Some(UiKind::Gui)),
            UiKind::Ratatui
        );
        // pcorro-named binary double-clicked: the name wins, GUI ignored.
        assert_eq!(
            resolve_headless_default(UiKind::Pancurses, false, true, false, Some(UiKind::Gui)),
            UiKind::Pancurses
        );
    }

    #[test]
    fn tty_or_gui_less_stays_put() {
        // Terminal present: default stands.
        assert_eq!(
            resolve_headless_default(UiKind::Ratatui, false, false, true, Some(UiKind::Gui)),
            UiKind::Ratatui
        );
        // Headless with no GUI compiled: nothing to fall back to.
        assert_eq!(
            resolve_headless_default(UiKind::Ratatui, false, false, false, None),
            UiKind::Ratatui
        );
    }

    #[test]
    fn parses_revision_limit() {
        let args = parse_args_from(["corro", "--revision", "2", "docs/test/main.corro"]);
        assert!(matches!(args.revision, Some(RevisionMode::Limit(2))));
        assert_eq!(args.files, vec![PathBuf::from("docs/test/main.corro")]);
    }

    #[test]
    fn parses_browse_mode() {
        let args = parse_args_from(["corro", "-r", "docs/test/main.corro"]);
        assert!(matches!(args.revision, Some(RevisionMode::Browse)));
        assert_eq!(args.files, vec![PathBuf::from("docs/test/main.corro")]);
    }

    fn parse_args_from<I, S>(iter: I) -> Args
    where
        I: IntoIterator<Item = S>,
        S: Into<std::ffi::OsString> + Clone,
    {
        let mut it = iter.into_iter();
        let program = it.next();
        let mut ui = program
            .clone()
            .map(|p| p.into().to_string_lossy().into_owned())
            .as_deref()
            .and_then(super::argv0_ui)
            .unwrap_or_else(super::determine_default_ui);
        let mut revision = None;
        let mut export = None;
        let mut movie = false;
        let mut movie_typing_cps = 22.0f64;
        let mut movie_confirm_ms = 120u64;
        let mut movie_menu_hold_ms = 1200u64;
        let mut show_help = false;
        let mut show_version = false;
        let mut positional = Vec::new();
        let mut rest = it.peekable();

        while let Some(arg) = rest.next() {
            let arg = arg.into();
            let arg = arg.to_string_lossy().into_owned();
            match arg.as_str() {
                "-h" | "-?" | "--help" => {
                    show_help = true;
                }
                "-v" | "--version" => {
                    show_version = true;
                }
                "-r" | "--revision" => {
                    if let Some(next) = rest.peek() {
                        let next = next.clone().into().to_string_lossy().into_owned();
                        if let Ok(value) = next.parse::<usize>() {
                            let _ = rest.next();
                            revision = Some(RevisionMode::Limit(value));
                            continue;
                        }
                    }
                    revision = Some(RevisionMode::Browse);
                }
                "-e" | "--export" => {
                    let next = rest.next().expect("export path");
                    let next = next.into().to_string_lossy().into_owned();
                    export = Some(PathBuf::from(next));
                }
                "--movie" => {
                    movie = true;
                }
                "--ratatui" => {
                    #[cfg(feature = "ratatui")]
                    { ui = super::UiKind::Ratatui; }
                }
                "--gui" => {
                    #[cfg(feature = "gui")]
                    { ui = super::UiKind::Gui; }
                }
                "--pancurses" => {
                    #[cfg(feature = "pancurses")]
                    { ui = super::UiKind::Pancurses; }
                }
                "--movie-typing-cps" => {
                    let next = rest.next().expect("movie typing cps");
                    let next = next.into().to_string_lossy().into_owned();
                    movie_typing_cps = next.parse::<f64>().expect("valid movie typing cps");
                }
                "--movie-confirm-ms" => {
                    let next = rest.next().expect("movie confirm delay");
                    let next = next.into().to_string_lossy().into_owned();
                    movie_confirm_ms = next.parse::<u64>().expect("valid movie confirm delay");
                }
                "--movie-menu-hold-ms" => {
                    let next = rest.next().expect("movie menu hold delay");
                    let next = next.into().to_string_lossy().into_owned();
                    movie_menu_hold_ms = next.parse::<u64>().expect("valid movie menu hold delay");
                }
                _ if arg.starts_with('-') => panic!("unexpected option"),
                _ => positional.push(arg),
            }
        }

        let files = positional.into_iter().map(PathBuf::from).collect();

        Args {
            revision,
            files,
            export,
            movie,
            movie_typing_cps,
            movie_confirm_ms,
            movie_menu_hold_ms,
            show_help,
            show_version,
            debug_no_number: false,
            ui,
            capture_html: None,
            convert_ansi: None,
        }
    }

    #[test]
    fn default_ui_is_terminal_unless_requested() {
        // No UI flag: the default must be a terminal UI whenever one is
        // compiled in — the GUI is opt-in via --gui (matches --help text).
        // Regression: with the `gui` feature enabled the default used to be
        // Gui, so `cargo run --features gui` popped a window unasked.
        let args = parse_args_from(["corro"]);
        #[cfg(all(not(target_arch = "wasm32"), feature = "ratatui"))]
        assert!(
            matches!(args.ui, super::UiKind::Ratatui),
            "default UI should be ratatui when compiled in"
        );
        #[cfg(all(
            not(target_arch = "wasm32"),
            not(feature = "ratatui"),
            feature = "pancurses"
        ))]
        assert!(
            matches!(args.ui, super::UiKind::Pancurses),
            "default UI should be pancurses when it is the only terminal UI"
        );
        #[cfg(target_arch = "wasm32")]
        assert!(
            matches!(args.ui, super::UiKind::Gui),
            "wasm default UI should be gui"
        );
    }

    #[test]
    fn explicit_ui_flags_select_backends() {
        // Explicit flags always win over the default, per compiled backend.
        #[cfg(feature = "ratatui")]
        assert!(matches!(
            parse_args_from(["corro", "--ratatui"]).ui,
            super::UiKind::Ratatui
        ));
        #[cfg(feature = "pancurses")]
        assert!(matches!(
            parse_args_from(["corro", "--pancurses"]).ui,
            super::UiKind::Pancurses
        ));
        #[cfg(feature = "gui")]
        assert!(matches!(
            parse_args_from(["corro", "--gui"]).ui,
            super::UiKind::Gui
        ));
    }

    #[test]
    fn argv0_selects_default_backend() {
        // Busybox-style dispatch: the program name selects the default UI.
        // Explicit flags still win (covered by the next test).
        #[cfg(feature = "pancurses")]
        assert!(matches!(
            parse_args_from(["pcorro"]).ui,
            super::UiKind::Pancurses
        ));
        #[cfg(feature = "gui")]
        assert!(matches!(
            parse_args_from(["gcorro"]).ui,
            super::UiKind::Gui
        ));
        // Prefix, path, .exe suffix, and case variants all dispatch.
        #[cfg(feature = "pancurses")]
        assert!(matches!(
            parse_args_from(["/usr/local/bin/pcorro-debug"]).ui,
            super::UiKind::Pancurses
        ));
        #[cfg(feature = "gui")]
        assert!(matches!(
            parse_args_from(["C:\\tools\\GCORRO.EXE"]).ui,
            super::UiKind::Gui
        ));
        // Unrelated names fall back to the normal default.
        #[cfg(all(not(target_arch = "wasm32"), feature = "ratatui"))]
        assert!(matches!(
            parse_args_from(["corro"]).ui,
            super::UiKind::Ratatui
        ));
        // An argv[0]-requested backend that is not compiled in falls back
        // to the normal default instead of erroring.
        #[cfg(all(
            not(target_arch = "wasm32"),
            feature = "ratatui",
            not(feature = "pancurses")
        ))]
        assert!(matches!(
            parse_args_from(["pcorro"]).ui,
            super::UiKind::Ratatui
        ));
    }

    #[test]
    fn explicit_ui_flags_override_argv0() {
        // An explicit flag always beats the program-name default.
        #[cfg(all(feature = "gui", feature = "ratatui"))]
        assert!(matches!(
            parse_args_from(["pcorro", "--gui"]).ui,
            super::UiKind::Gui
        ));
        #[cfg(all(feature = "pancurses", feature = "ratatui"))]
        assert!(matches!(
            parse_args_from(["gcorro", "--pancurses"]).ui,
            super::UiKind::Pancurses
        ));
        #[cfg(all(feature = "ratatui", feature = "gui"))]
        assert!(matches!(
            parse_args_from(["gcorro", "--ratatui"]).ui,
            super::UiKind::Ratatui
        ));
    }

    #[test]
    fn parses_export_path() {
        let args = parse_args_from(["corro", "--export", "out.ods", "docs/test/main.corro"]);
        assert_eq!(
            args.export.as_deref(),
            Some(std::path::Path::new("out.ods"))
        );
        assert_eq!(args.files, vec![PathBuf::from("docs/test/main.corro")]);
    }

    #[test]
    fn parses_multiple_tabular_inputs() {
        let args = parse_args_from(["corro", "a.csv", "b.tsv"]);
        assert_eq!(args.files, vec![PathBuf::from("a.csv"), PathBuf::from("b.tsv")]);
    }

    #[test]
    fn parses_movie_options() {
        let args = parse_args_from([
            "corro",
            "--movie",
            "--movie-typing-cps",
            "30",
            "--movie-confirm-ms",
            "500",
            "--movie-menu-hold-ms",
            "1600",
            "docs/test/main.corro",
        ]);
        assert!(args.movie);
        assert!((args.movie_typing_cps - 30.0).abs() < f64::EPSILON);
        assert_eq!(args.movie_confirm_ms, 500);
        assert_eq!(args.movie_menu_hold_ms, 1600);
    }

    #[test]
    fn parses_help_variants() {
        assert!(parse_args_from(["corro", "--help"]).show_help);
        assert!(parse_args_from(["corro", "-h"]).show_help);
        assert!(parse_args_from(["corro", "-?"]).show_help);
    }

    #[test]
    fn parses_version_variants() {
        assert!(parse_args_from(["corro", "--version"]).show_version);
        assert!(parse_args_from(["corro", "-v"]).show_version);
    }

    #[test]
    fn movie_cps_option_has_typing_cps_suggestion() {
        assert_eq!(
            super::cli_option_suggestion("--movie-cps"),
            Some("--movie-typing-cps")
        );
        assert_eq!(super::cli_option_suggestion("--unknown"), None);
    }
}
