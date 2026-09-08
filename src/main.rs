//! corro — append-only collaborative spreadsheet TUI.

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

#[cfg(any(target_arch = "wasm32", feature = "gui"))]
fn determine_default_ui() -> UiKind { UiKind::Gui }

#[cfg(all(
    not(any(target_arch = "wasm32", feature = "gui")),
    feature = "pancurses"
))]
fn determine_default_ui() -> UiKind { UiKind::Pancurses }

#[cfg(all(
    not(any(target_arch = "wasm32", feature = "gui")),
    not(feature = "pancurses")
))]
fn determine_default_ui() -> UiKind { UiKind::Ratatui }

// Win95-safe command line: `GetCommandLineW` is a no-op stub on Windows 95
// (returns NULL), so `std::env::args()` cannot be used there. Read the ANSI
// command line instead. DBCS bytes decode lossily, which is fine for the
// ASCII-only options corro takes.
#[cfg(all(target_family = "rust9x", target_env = "msvc"))]
fn cli_args() -> impl Iterator<Item = String> {
    win95_args().into_iter().skip(1)
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
    let mut ui = determine_default_ui();
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
                #[cfg(feature = "ratatui")]
                { ui = UiKind::Ratatui; }
                #[cfg(not(feature = "ratatui"))]
                { return Err("ratatui UI not compiled in; rebuild with --features ratatui".into()); }
            }
            "--gui" => {
                #[cfg(feature = "gui")]
                { ui = UiKind::Gui; }
                #[cfg(not(feature = "gui"))]
                { return Err("GTK GUI not compiled in; rebuild with --features gui".into()); }
            }
            "--pancurses" => {
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
#[cfg(all(target_family = "rust9x", target_env = "msvc"))]
#[no_mangle]
pub extern "C" fn main() -> i32 {
    win9x_redirect_console_output();
    corro_main();
    0
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
    use super::{Args, RevisionMode};
    use std::path::PathBuf;

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
        let _program = it.next();
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
            ui: super::determine_default_ui(),
            capture_html: None,
            convert_ansi: None,
        }
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
