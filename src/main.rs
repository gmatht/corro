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
    let mut app = App::new(args.file);
    app.load_initial()?;
    app.run()?;
    Ok(())
}
