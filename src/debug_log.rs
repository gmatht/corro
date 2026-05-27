use std::fs::OpenOptions;
use std::io::Write;
use std::path::PathBuf;

fn debug_log_path() -> Option<PathBuf> {
    std::env::var("CORRO_DEBUG_LOG").ok().map(PathBuf::from).or_else(|| {
        std::env::var("XDG_STATE_HOME").ok().map(|xdg| PathBuf::from(format!("{}/corro/debug.log", xdg)))
    })
    .or_else(|| std::env::var("HOME").ok().map(|h| PathBuf::from(format!("{}/.corro/debug.log", h))))
}

/// Write a debug line to the configured debug log (best-effort).
pub fn log(msg: &str) {
    if let Some(p) = debug_log_path() {
        if let Some(dir) = p.parent() {
            let _ = std::fs::create_dir_all(dir);
        }
        if let Ok(mut f) = OpenOptions::new().create(true).append(true).open(&p) {
            let _ = writeln!(f, "{}", msg);
        }
    }
}
