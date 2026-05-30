//! corro — append-only collaborative spreadsheet TUI.

use corro::ui::App;
use std::path::PathBuf;

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
}

enum RevisionMode {
    Browse,
    Limit(usize),
}

fn cli_option_suggestion(arg: &str) -> Option<&'static str> {
    match arg {
        "--movie-cps" => Some("--movie-typing-cps"),
        _ => None,
    }
}

fn parse_args() -> Result<Args, String> {
    let mut revision = None;
    let mut export = None;
    let mut movie = false;
    let mut movie_typing_cps = 22.0f64;
    let mut movie_confirm_ms = 120u64;
    let mut movie_menu_hold_ms = 1200u64;
    let mut show_help = false;
    let mut show_version = false;
    let mut debug_no_number = false;
    let mut positional = Vec::new();
    let mut it = std::env::args().skip(1).peekable();

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
            "--movie-typing-cps" => {
                let Some(raw) = it.next() else {
                    return Err("--movie-typing-cps requires a positive number".into());
                };
                let Ok(value) = raw.parse::<f64>() else {
                    return Err(format!("invalid --movie-typing-cps value: {raw}"));
                };
                if value <= 0.0 {
                    return Err("--movie-typing-cps must be > 0".into());
                }
                movie_typing_cps = value;
            }
            "--movie-confirm-ms" => {
                let Some(raw) = it.next() else {
                    return Err("--movie-confirm-ms requires a non-negative integer".into());
                };
                let Ok(value) = raw.parse::<u64>() else {
                    return Err(format!("invalid --movie-confirm-ms value: {raw}"));
                };
                movie_confirm_ms = value;
            }
            "--movie-menu-hold-ms" => {
                let Some(raw) = it.next() else {
                    return Err("--movie-menu-hold-ms requires a non-negative integer".into());
                };
                let Ok(value) = raw.parse::<u64>() else {
                    return Err(format!("invalid --movie-menu-hold-ms value: {raw}"));
                };
                movie_menu_hold_ms = value;
            }
            "--debug-no-number" => {
                // When set, emit a backtrace and panic if a SET line with a
                // purely numeric value is about to be appended. The runtime
                // detection reads CORRO_DEBUG_NO_NUMBER which we set in
                // try_main after parsing (so the rest of the program can
                // observe it as an env var).
                debug_no_number = true;
            }
            _ if arg.starts_with('-') => {
                if let Some(suggested) = cli_option_suggestion(&arg) {
                    return Err(format!(
                        "unrecognized option: {arg} (did you mean {suggested}?)"
                    ));
                }
                return Err(format!("unrecognized option: {arg}"));
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
    })
}

fn main() {
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

fn try_main() -> (Result<(), corro::ui::RunError>, Option<String>) {
    // Parse args; return early with no exit message on CLI errors/help/version.
    let args = match parse_args() {
        Ok(a) => a,
        Err(s) => return (Err(std::io::Error::other(s).into()), None),
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
    if let Some(export_path) = args.export {
        let input_path = match args.files.first() {
            Some(p) => p.clone(),
            None => {
                return (
                    Err(std::io::Error::other("--export requires an input file argument").into()),
                    None,
                )
            }
        };
        if args.files.len() > 1 {
            return (
                Err(std::io::Error::other("--export accepts exactly one input file").into()),
                None,
            );
        }
        let workbook = match load_workbook_for_export(&input_path) {
            Ok(w) => w,
            Err(e) => return (Err(std::io::Error::other(e).into()), None),
        };
        if let Err(e) = export_workbook_to_path(&workbook, &export_path) {
            return (Err(std::io::Error::other(e).into()), None);
        }
        return (Ok(()), None);
    }
    if args.movie && args.revision.is_some() {
        return (
            Err(std::io::Error::other("--movie cannot be combined with --revision").into()),
            None,
        );
    }
    if args.revision.is_some() && args.files.len() > 1 {
        return (
            Err(std::io::Error::other("--revision accepts exactly one input file").into()),
            None,
        );
    }
    if args.movie && args.files.len() > 1 {
        return (
            Err(std::io::Error::other("--movie accepts exactly one input file").into()),
            None,
        );
    }
    let mut app = match args.revision {
        None => App::new_with_paths(args.files),
        Some(RevisionMode::Browse) => App::new_with_revision_browser(args.files.first().cloned()),
        Some(RevisionMode::Limit(revision)) => {
            App::new_with_revision_limit(args.files.first().cloned(), Some(revision))
        }
    };
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
    (res, exit_msg)
}

fn cli_help_text() -> String {
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
\n\
ARGS:\n\
  FILE                      Input file(s) (.corro, .ods, .tsv, .csv)\n",
        env!("CARGO_PKG_VERSION")
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
