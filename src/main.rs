//! corro — append-only collaborative spreadsheet TUI.

use corro::ui::App;
use std::path::PathBuf;

struct Args {
    revision: Option<RevisionMode>,
    file: Option<PathBuf>,
    export: Option<PathBuf>,
    movie: bool,
    movie_typing_cps: f64,
    movie_confirm_ms: u64,
    movie_menu_hold_ms: u64,
    show_help: bool,
    show_version: bool,
}

enum RevisionMode {
    Browse,
    Limit(usize),
}

fn parse_args() -> Result<Args, String> {
    let mut revision = None;
    let mut file = None;
    let mut export = None;
    let mut movie = false;
    let mut movie_typing_cps = 22.0f64;
    let mut movie_confirm_ms = 120u64;
    let mut movie_menu_hold_ms = 1200u64;
    let mut show_help = false;
    let mut show_version = false;
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
            _ if arg.starts_with('-') => {
                return Err(format!("unrecognized option: {arg}"));
            }
            _ => positional.push(arg),
        }
    }

    if positional.len() > 1 {
        return Err("too many positional arguments".into());
    }
    if let Some(path) = positional.pop() {
        file = Some(PathBuf::from(path));
    }

    Ok(Args {
        revision,
        file,
        export,
        movie,
        movie_typing_cps,
        movie_confirm_ms,
        movie_menu_hold_ms,
        show_help,
        show_version,
    })
}

fn main() {
    if let Err(e) = try_main() {
        eprintln!("{e}");
        std::process::exit(1);
    }
}

fn try_main() -> Result<(), corro::ui::RunError> {
    let args = parse_args().map_err(std::io::Error::other)?;
    if args.show_help {
        println!("{}", cli_help_text());
        return Ok(());
    }
    if args.show_version {
        println!("corro {}", env!("CARGO_PKG_VERSION"));
        return Ok(());
    }
    if let Some(export_path) = args.export {
        let Some(input_path) = args.file else {
            return Err(std::io::Error::other(
                "--export requires an input file argument",
            )
            .into());
        };
        let workbook = load_workbook_for_export(&input_path).map_err(std::io::Error::other)?;
        export_workbook_to_path(&workbook, &export_path).map_err(std::io::Error::other)?;
        return Ok(());
    }
    if args.movie && args.revision.is_some() {
        return Err(std::io::Error::other("--movie cannot be combined with --revision").into());
    }
    let mut app = match args.revision {
        None => App::new(args.file),
        Some(RevisionMode::Browse) => App::new_with_revision_browser(args.file),
        Some(RevisionMode::Limit(revision)) => {
            App::new_with_revision_limit(args.file, Some(revision))
        }
    };
    if args.movie {
        app.run_movie(corro::ui::MovieReplayOptions {
            typing_cps: args.movie_typing_cps,
            confirm_delay_ms: args.movie_confirm_ms,
            menu_hold_ms: args.movie_menu_hold_ms,
        })?;
    } else {
        app.load_initial()?;
        app.run()?;
    }
    Ok(())
}

fn cli_help_text() -> String {
    format!(
        "corro {}\n\
\n\
USAGE:\n\
  corro [OPTIONS] [FILE]\n\
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
  FILE                      Input file (.corro, .ods, .tsv, .csv)\n",
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
        assert_eq!(
            args.file.as_deref(),
            Some(std::path::Path::new("docs/test/main.corro"))
        );
    }

    #[test]
    fn parses_browse_mode() {
        let args = parse_args_from(["corro", "-r", "docs/test/main.corro"]);
        assert!(matches!(args.revision, Some(RevisionMode::Browse)));
        assert_eq!(
            args.file.as_deref(),
            Some(std::path::Path::new("docs/test/main.corro"))
        );
    }

    fn parse_args_from<I, S>(iter: I) -> Args
    where
        I: IntoIterator<Item = S>,
        S: Into<std::ffi::OsString> + Clone,
    {
        let mut it = iter.into_iter();
        let _program = it.next();
        let mut revision = None;
        let mut file = None;
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

        if let Some(path) = positional.pop() {
            file = Some(PathBuf::from(path));
        }

        Args {
            revision,
            file,
            export,
            movie,
            movie_typing_cps,
            movie_confirm_ms,
            movie_menu_hold_ms,
            show_help,
            show_version,
        }
    }

    #[test]
    fn parses_export_path() {
        let args = parse_args_from(["corro", "--export", "out.ods", "docs/test/main.corro"]);
        assert_eq!(
            args.export.as_deref(),
            Some(std::path::Path::new("out.ods"))
        );
        assert_eq!(
            args.file.as_deref(),
            Some(std::path::Path::new("docs/test/main.corro"))
        );
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
}
