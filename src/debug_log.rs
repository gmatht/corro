use std::fs::OpenOptions;
use std::io::Write;
use std::path::PathBuf;

/// Where debug traces go, or None when no location can be determined.
/// Checked in order: `CORRO_DEBUG_LOG`, `XDG_STATE_HOME`, `HOME` (Unix),
/// then `LOCALAPPDATA` (Windows; mirrors the unsaved-file directory).
/// Shared with `try_main`, which redirects the process's stderr here so
/// debug `eprintln!` traces never corrupt the TUI.
pub fn debug_log_path() -> Option<PathBuf> {
    fn nonempty(v: Option<String>) -> Option<String> {
        v.filter(|s| !s.trim().is_empty())
    }
    if let Some(p) = nonempty(std::env::var("CORRO_DEBUG_LOG").ok()) {
        return Some(PathBuf::from(p));
    }
    if let Some(xdg) = nonempty(std::env::var("XDG_STATE_HOME").ok()) {
        return Some(PathBuf::from(xdg).join("corro").join("debug.log"));
    }
    if let Some(home) = nonempty(std::env::var("HOME").ok()) {
        return Some(PathBuf::from(home).join(".corro").join("debug.log"));
    }
    if let Some(local) = nonempty(std::env::var("LOCALAPPDATA").ok()) {
        return Some(PathBuf::from(local).join("corro").join("debug.log"));
    }
    None
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
