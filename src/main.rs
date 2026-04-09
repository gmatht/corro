//! corro — append-only collaborative spreadsheet TUI.

use corro::ui::App;
use std::path::PathBuf;

struct Args {
    revision: Option<RevisionMode>,
    file: Option<PathBuf>,
}

enum RevisionMode {
    Browse,
    Limit(usize),
}

fn parse_args() -> Result<Args, String> {
    let mut revision = None;
    let mut file = None;
    let mut positional = Vec::new();
    let mut it = std::env::args().skip(1).peekable();

    while let Some(arg) = it.next() {
        match arg.as_str() {
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

    Ok(Args { revision, file })
}

fn main() {
    if let Err(e) = try_main() {
        eprintln!("{e}");
        std::process::exit(1);
    }
}

fn try_main() -> Result<(), corro::ui::RunError> {
    let args = parse_args().map_err(std::io::Error::other)?;
    let mut app = match args.revision {
        None => App::new(args.file),
        Some(RevisionMode::Browse) => App::new_with_revision_browser(args.file),
        Some(RevisionMode::Limit(revision)) => {
            App::new_with_revision_limit(args.file, Some(revision))
        }
    };
    app.load_initial()?;
    app.run()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{Args, RevisionMode};
    use std::path::PathBuf;

    #[test]
    fn parses_revision_limit() {
        let args = parse_args_from(["corro", "--revision", "2", "test5.corro"]);
        assert!(matches!(args.revision, Some(RevisionMode::Limit(2))));
        assert_eq!(
            args.file.as_deref(),
            Some(std::path::Path::new("test5.corro"))
        );
    }

    #[test]
    fn parses_browse_mode() {
        let args = parse_args_from(["corro", "-r", "test5.corro"]);
        assert!(matches!(args.revision, Some(RevisionMode::Browse)));
        assert_eq!(
            args.file.as_deref(),
            Some(std::path::Path::new("test5.corro"))
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
        let mut positional = Vec::new();
        let mut rest = it.peekable();

        while let Some(arg) = rest.next() {
            let arg = arg.into();
            let arg = arg.to_string_lossy().into_owned();
            match arg.as_str() {
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
                _ if arg.starts_with('-') => panic!("unexpected option"),
                _ => positional.push(arg),
            }
        }

        if let Some(path) = positional.pop() {
            file = Some(PathBuf::from(path));
        }

        Args { revision, file }
    }
}
