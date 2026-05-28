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

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;

    #[test]
    fn debug_log_writes_when_env_set() {
        let path = "/tmp/corro-debug.log";
        // Clean up any existing file; ignore errors
        let _ = fs::remove_file(path);
        // Ensure the env var is set for this test process
        std::env::set_var("CORRO_DEBUG_LOG", path);
        log("TEST debug log entry");
        let content = fs::read_to_string(path).expect("debug log file was not created");
        assert!(content.contains("TEST debug log entry"));
        // Cleanup
        let _ = fs::remove_file(path);
    }
}
