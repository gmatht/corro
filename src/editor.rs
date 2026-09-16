//! Open the user's text editor on a scratch file and read back the result.
//!
//! Used by every UI for "Edit in Text Editor" (long cell contents are
//! painful in a one-line formula bar) and, via
//! [`edit_workbook_externally`], for "Edit ▸ Workbook (External)" (the
//! append-only log itself). The caller owns everything around the call:
//! terminal suspend/resume on TUI backends, committing the returned text,
//! status messages, reloading after a workbook edit. This module only
//! resolves the editor, runs it, and normalizes the result.
//!
//! Conventions (documented, tested below):
//! - `$VISUAL`, then `$EDITOR`, then a platform fallback (`notepad` on
//!   Windows, `vi` elsewhere). Words after the first are passed as arguments
//!   (`EDITOR="code --wait"` works); no shell is involved.
//! - An explicitly set but unlaunchable editor is an error (no silent
//!   fallback to something the user didn't ask for).
//! - Non-zero editor exit means *abort* (vim `:cq`): `Ok(None)`.
//! - Unchanged content (ignoring a single trailing line break the editor may
//!   have added) is `Ok(None)`.
//! - Otherwise `Ok(Some(text))` with one trailing line break stripped.

use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};

/// Process-unique counter for scratch directory names (no extra deps).
static SCRATCH_SEQ: AtomicU64 = AtomicU64::new(0);

/// A temp dir removed on drop (std-only; `tempfile` is dev-dependencies).
struct ScratchDir {
    path: PathBuf,
}

impl ScratchDir {
    fn new() -> Result<Self, String> {
        for _ in 0..100 {
            let seq = SCRATCH_SEQ.fetch_add(1, Ordering::Relaxed);
            let path =
                std::env::temp_dir().join(format!("corro-edit-{}-{}", std::process::id(), seq));
            match std::fs::create_dir(&path) {
                Ok(()) => return Ok(Self { path }),
                Err(e) if e.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(e) => return Err(format!("scratch dir: {e}")),
            }
        }
        Err("scratch dir: too many collisions".into())
    }
}

impl Drop for ScratchDir {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.path);
    }
}

pub fn edit_text_externally(initial: &str) -> Result<Option<String>, String> {
    #[cfg(test)]
    if let Some(behavior) = TEST_EDITOR.with(|cell| cell.borrow().clone()) {
        return match behavior {
            TestEditor::Output(text) => {
                let back = strip_added_line_break(&text);
                if back == initial {
                    Ok(None)
                } else {
                    Ok(Some(back.to_string()))
                }
            }
            TestEditor::Unchanged => Ok(None),
            TestEditor::Error(e) => Err(e),
        };
    }
    let (prog, args) = resolve_editor()?;
    let scratch = ScratchDir::new()?;
    let path = scratch.path.join("cell.txt");
    std::fs::write(&path, initial).map_err(|e| format!("write scratch: {e}"))?;
    let status = std::process::Command::new(&prog)
        .args(&args)
        .arg(&path)
        .status()
        .map_err(|e| format!("launch {prog}: {e}"))?;
    if !status.success() {
        // Non-zero exit = abort (e.g. vim :cq): leave the cell alone.
        return Ok(None);
    }
    let back = std::fs::read_to_string(&path).map_err(|e| format!("read back: {e}"))?;
    let back = strip_added_line_break(&back);
    if back == initial {
        Ok(None)
    } else {
        Ok(Some(back.to_string()))
    }
    // `scratch` drops here, removing the scratch dir.
}

/// Open the user's editor on the real workbook file at `path` (Edit ▸
/// Workbook (External)). Unlike [`edit_text_externally`] this does **not**
/// copy to a scratch file: the user is editing the live append-only log, so
/// their edits are the file. Returns `Ok(true)` when the file changed
/// (mtime or size), `Ok(false)` when the editor saved nothing, and `Err` when
/// the editor could not be launched. A non-zero editor exit is still an
/// error here (unlike the cell roundtrip, where it means "abort") because the
/// file may have been half-written; callers report and reload regardless.
pub fn edit_workbook_externally(path: &std::path::Path) -> Result<bool, String> {
    #[cfg(test)]
    if let Some(behavior) = TEST_EDITOR.with(|cell| cell.borrow().clone()) {
        return match behavior {
            TestEditor::Output(text) => {
                std::fs::write(path, text).map_err(|e| format!("test editor write: {e}"))?;
                Ok(true)
            }
            TestEditor::Unchanged => Ok(false),
            TestEditor::Error(e) => Err(e),
        };
    }
    let before = std::fs::read(path).ok();
    let (prog, args) = resolve_editor()?;
    let status = std::process::Command::new(&prog)
        .args(&args)
        .arg(path)
        .status()
        .map_err(|e| format!("launch {prog}: {e}"))?;
    if !status.success() {
        return Err(format!("{prog} exited with {status}"));
    }
    // Compare content, not size+mtime: editors that save within the same
    // timestamp tick (and same length, e.g. "old" -> "new") would otherwise
    // look unchanged. The log is small, so reading it back is cheap.
    let after = std::fs::read(path).ok();
    Ok(before != after)
}

fn resolve_editor() -> Result<(String, Vec<String>), String> {
    for var in ["VISUAL", "EDITOR"] {
        if let Ok(val) = std::env::var(var) {
            let val = val.trim().to_string();
            if val.is_empty() {
                continue;
            }
            let mut words = val.split_whitespace().map(str::to_string);
            let prog = words.next().expect("non-empty after trim");
            return Ok((prog, words.collect()));
        }
    }
    #[cfg(target_os = "windows")]
    return Ok(("notepad".to_string(), Vec::new()));
    #[cfg(not(target_os = "windows"))]
    return Ok(("vi".to_string(), Vec::new()));
}

/// Strip at most one trailing line break (what editors add on save).
fn strip_added_line_break(s: &str) -> &str {
    s.strip_suffix("\r\n")
        .or_else(|| s.strip_suffix('\n'))
        .unwrap_or(s)
}

#[cfg(test)]
thread_local! {
    static TEST_EDITOR: std::cell::RefCell<Option<TestEditor>> =
        const { std::cell::RefCell::new(None) };
}

/// Test hook (mirrors the `TEST_CLIPBOARD` pattern): bypasses process
/// spawning so tests stay hermetic. Thread-local, so parallel tests are safe.
#[cfg(test)]
#[derive(Clone, Debug)]
pub enum TestEditor {
    /// Pretend the editor left this text in the file (normalization applies).
    Output(String),
    /// Pretend the editor saved without changes.
    Unchanged,
    /// Pretend spawning/editing failed.
    Error(String),
}

#[cfg(test)]
pub fn set_test_editor(behavior: Option<TestEditor>) {
    TEST_EDITOR.with(|cell| *cell.borrow_mut() = behavior);
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::sync::{Mutex, MutexGuard};

    // $VISUAL/$EDITOR are process-global: serialize these tests against each
    // other and restore the environment afterwards (the 40-site env race in
    // ui tests shows what happens otherwise). The lock lives in the guard so
    // it is held for the whole test body.
    static ENV_LOCK: Mutex<()> = Mutex::new(());
    struct EnvGuard {
        _lock: MutexGuard<'static, ()>,
        visual: Option<String>,
        editor: Option<String>,
    }
    impl EnvGuard {
        fn take() -> Self {
            let lock = ENV_LOCK.lock().unwrap();
            let visual = std::env::var("VISUAL").ok();
            let editor = std::env::var("EDITOR").ok();
            std::env::remove_var("VISUAL");
            std::env::remove_var("EDITOR");
            Self {
                _lock: lock,
                visual,
                editor,
            }
        }
    }
    impl Drop for EnvGuard {
        fn drop(&mut self) {
            match &self.visual {
                Some(v) => std::env::set_var("VISUAL", v),
                None => std::env::remove_var("VISUAL"),
            }
            match &self.editor {
                Some(e) => std::env::set_var("EDITOR", e),
                None => std::env::remove_var("EDITOR"),
            }
            // `_lock` drops last, releasing the mutex after restore.
        }
    }

    #[cfg(unix)]
    fn script(dir: &std::path::Path, name: &str, body: &str) -> String {
        let p = dir.join(name);
        std::fs::write(&p, body).unwrap();
        use std::os::unix::fs::PermissionsExt;
        std::fs::set_permissions(&p, std::fs::Permissions::from_mode(0o755)).unwrap();
        p.to_string_lossy().into_owned()
    }

    #[cfg(unix)]
    #[test]
    fn unchanged_file_means_no_change() {
        let _g = EnvGuard::take();
        // `true` ignores its argument: file untouched.
        std::env::set_var("EDITOR", "true");
        assert_eq!(edit_text_externally("hello").unwrap(), None);
    }

    #[cfg(unix)]
    #[test]
    fn changed_file_returns_new_text() {
        let _g = EnvGuard::take();
        let dir = tempfile::tempdir().unwrap();
        let ed = script(dir.path(), "ed.sh", "#!/bin/sh\nprintf 'NEW' > \"$1\"\n");
        std::env::set_var("EDITOR", &ed);
        assert_eq!(
            edit_text_externally("old").unwrap(),
            Some("NEW".to_string())
        );
    }

    #[cfg(unix)]
    #[test]
    fn added_trailing_newline_is_ignored() {
        let _g = EnvGuard::take();
        let dir = tempfile::tempdir().unwrap();
        let ed = script(dir.path(), "ed.sh", "#!/bin/sh\nprintf 'same\\n' > \"$1\"\n");
        std::env::set_var("EDITOR", &ed);
        assert_eq!(edit_text_externally("same").unwrap(), None);
    }

    #[cfg(unix)]
    #[test]
    fn nonzero_exit_means_abort() {
        let _g = EnvGuard::take();
        std::env::set_var("EDITOR", "false");
        assert_eq!(edit_text_externally("old").unwrap(), None);
    }

    #[test]
    fn missing_editor_is_an_error() {
        let _g = EnvGuard::take();
        std::env::set_var("EDITOR", "/nonexistent-editor-xyz");
        assert!(edit_text_externally("old").is_err());
    }

    #[cfg(unix)]
    #[test]
    fn editor_args_are_supported() {
        let _g = EnvGuard::take();
        let dir = tempfile::tempdir().unwrap();
        // Probe script records its two args into the target file; the helper
        // invokes "<script> --flag <path>", so the read-back must show both.
        let ed = script(
            dir.path(),
            "ed.sh",
            "#!/bin/sh\nprintf 'VIA=%s\\n' \"$1\" > \"$2\"\n",
        );
        std::env::set_var("EDITOR", format!("{ed} --flag"));
        let out = edit_text_externally("old").unwrap().unwrap();
        assert!(out.contains("--flag"), "{out:?}");
    }

    #[test]
    fn test_hook_short_circuits_spawn() {
        let _g = EnvGuard::take();
        std::env::set_var("EDITOR", "/nonexistent-editor-xyz");
        set_test_editor(Some(TestEditor::Output("hooked\n".into())));
        assert_eq!(edit_text_externally("hooked").unwrap(), None);
        set_test_editor(Some(TestEditor::Output("changed".into())));
        assert_eq!(
            edit_text_externally("old").unwrap(),
            Some("changed".to_string())
        );
        set_test_editor(Some(TestEditor::Unchanged));
        assert_eq!(edit_text_externally("old").unwrap(), None);
        set_test_editor(Some(TestEditor::Error("boom".into())));
        assert_eq!(edit_text_externally("old").unwrap_err(), "boom");
        set_test_editor(None);
    }

    /// The workbook variant edits the file in place (not a scratch copy), so
    /// a changed file is reported as `true` — the caller reloads from it.
    #[cfg(unix)]
    #[test]
    fn workbook_edit_reports_change_and_leaves_the_file_edited() {
        let _g = EnvGuard::take();
        let dir = tempfile::tempdir().unwrap();
        let wb = dir.path().join("book.corro");
        std::fs::write(&wb, "CORRO_LOG 1\nSET A1 old\n").unwrap();
        let ed = script(
            dir.path(),
            "ed.sh",
            "#!/bin/sh\nprintf 'CORRO_LOG 1\\nSET A1 new\\n' > \"$1\"\n",
        );
        std::env::set_var("EDITOR", &ed);
        assert!(edit_workbook_externally(&wb).unwrap());
        assert_eq!(
            std::fs::read_to_string(&wb).unwrap(),
            "CORRO_LOG 1\nSET A1 new\n"
        );
    }

    #[cfg(unix)]
    #[test]
    fn workbook_edit_without_change_reports_false() {
        let _g = EnvGuard::take();
        let dir = tempfile::tempdir().unwrap();
        let wb = dir.path().join("book.corro");
        std::fs::write(&wb, "CORRO_LOG 1\n").unwrap();
        std::env::set_var("EDITOR", "true"); // ignores its argument
        assert!(!edit_workbook_externally(&wb).unwrap());
    }

    #[cfg(unix)]
    #[test]
    fn workbook_edit_nonzero_exit_is_an_error() {
        let _g = EnvGuard::take();
        let dir = tempfile::tempdir().unwrap();
        let wb = dir.path().join("book.corro");
        std::fs::write(&wb, "CORRO_LOG 1\n").unwrap();
        std::env::set_var("EDITOR", "false");
        assert!(edit_workbook_externally(&wb).is_err());
    }

    #[cfg(unix)]
    #[test]
    fn workbook_edit_missing_editor_is_an_error() {
        let _g = EnvGuard::take();
        let dir = tempfile::tempdir().unwrap();
        let wb = dir.path().join("book.corro");
        std::fs::write(&wb, "CORRO_LOG 1\n").unwrap();
        std::env::set_var("EDITOR", "/nonexistent-editor-xyz");
        assert!(edit_workbook_externally(&wb).is_err());
    }

    #[test]
    fn workbook_edit_test_hook_writes_and_reports_change() {
        let _g = EnvGuard::take();
        std::env::set_var("EDITOR", "/nonexistent-editor-xyz");
        let dir = tempfile::tempdir().unwrap();
        let wb = dir.path().join("book.corro");
        std::fs::write(&wb, "CORRO_LOG 1\n").unwrap();

        set_test_editor(Some(TestEditor::Output("CORRO_LOG 1\nSET A1 hooked\n".into())));
        assert!(edit_workbook_externally(&wb).unwrap());
        assert_eq!(
            std::fs::read_to_string(&wb).unwrap(),
            "CORRO_LOG 1\nSET A1 hooked\n"
        );

        set_test_editor(Some(TestEditor::Unchanged));
        assert!(!edit_workbook_externally(&wb).unwrap());

        set_test_editor(Some(TestEditor::Error("boom".into())));
        assert_eq!(edit_workbook_externally(&wb).unwrap_err(), "boom");
        set_test_editor(None);
    }
}
