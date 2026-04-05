//! corro — append-only collaborative spreadsheet TUI.

use clap::Parser;
use corro::ui::App;
use std::path::PathBuf;

#[derive(Parser, Debug)]
#[command(
    name = "corro",
    about = "Terminal spreadsheet with append-only text log"
)]
struct Args {
    /// Load at most this many revisions from a log file.
    #[arg(short = 'r', long = "revision", value_name = "N")]
    revision: Option<usize>,

    /// Path to the `.corro` log file. Created on first write.
    #[arg(value_name = "FILE")]
    file: Option<PathBuf>,
}

fn main() {
    if let Err(e) = try_main() {
        eprintln!("{e}");
        std::process::exit(1);
    }
}

fn try_main() -> Result<(), corro::ui::RunError> {
    let args = Args::parse();
    let mut app = App::new_with_revision_limit(args.file, args.revision);
    app.load_initial()?;
    app.run()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::Args;
    use clap::Parser;

    #[test]
    fn parses_revision_limit() {
        let args = Args::parse_from(["corro", "--revision", "2", "test5.corro"]);
        assert_eq!(args.revision, Some(2));
        assert_eq!(
            args.file.as_deref(),
            Some(std::path::Path::new("test5.corro"))
        );
    }

    #[test]
    fn parses_short_revision_limit() {
        let args = Args::parse_from(["corro", "-r", "3", "test5.corro"]);
        assert_eq!(args.revision, Some(3));
    }
}
