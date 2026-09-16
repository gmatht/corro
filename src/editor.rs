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
//! - From a native GUI backend (no terminal) there is no `$EDITOR` roundtrip:
//!   [`open_gui_editor_detached`] launches a GUI editor instead
//!   (`$CORRO_GUI_EDITOR`, else `xdg-open` on Linux, `open` on macOS,
//!   `notepad` on Windows) and returns immediately. The caller reports
//!   "opened" and the log tail (polled by every backend) picks up the
//!   user's save — so a blocking wait would only race the editor.
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

/// Known Linux GUI text editors (binary names), lightest first: a `.corro`
/// log wants a fast window, not an IDE. GNOME Text Editor, Xed (Mint),
/// Mousepad (Xfce, ultra-light), Geany, Kate (KDE), Bluefish, legacy gedit,
/// Pluma (MATE); the heavyweights (Sublime, VS Code) come last so they are
/// only picked when nothing lighter is installed. All of them open `<file>`
/// detached with no extra flags.
#[cfg(not(any(target_os = "windows", target_os = "macos")))]
const KNOWN_GUI_EDITORS: &[&str] = &[
    "gnome-text-editor",
    "xed",
    "mousepad",
    "geany",
    "kate",
    "bluefish",
    "gedit",
    "pluma",
    "subl",
    "code",
];

/// Front-runner editors for the current desktop, from `$XDG_CURRENT_DESKTOP`
/// (colon-separated, e.g. `"ubuntu:GNOME"`): the native editor beats the
/// generic lightest-first list because it matches the user's theme and file
/// associations.
#[cfg(not(any(target_os = "windows", target_os = "macos")))]
fn desktop_preferred_editors() -> Vec<&'static str> {
    let desktops = std::env::var("XDG_CURRENT_DESKTOP").unwrap_or_default();
    let mut out = Vec::new();
    for d in desktops.split([':', ';']) {
        match d.trim().to_ascii_uppercase().as_str() {
            "GNOME" | "UBUNTU" | "PANTHEON" => out.extend(["gnome-text-editor", "gedit"]),
            "KDE" | "PLASMA" => out.extend(["kate"]),
            "XFCE" => out.extend(["mousepad"]),
            // Cinnamon reports "X-Cinnamon".
            "CINNAMON" | "X-CINNAMON" => out.extend(["xed"]),
            "MATE" => out.extend(["pluma"]),
            _ => {}
        }
    }
    out
}

/// Is `prog` an executable file somewhere on `$PATH`? (Resolution only —
/// nothing is spawned, so tests can point `$PATH` at a scratch dir.)
#[cfg(not(any(target_os = "windows", target_os = "macos")))]
fn find_in_path(prog: &str) -> bool {
    let Ok(path_var) = std::env::var("PATH") else {
        return false;
    };
    for dir in std::env::split_paths(&path_var) {
        let candidate = dir.join(prog);
        let Ok(meta) = std::fs::metadata(&candidate) else {
            continue;
        };
        if !meta.is_file() {
            continue;
        }
        #[cfg(unix)]
        {
            use std::os::unix::fs::PermissionsExt;
            if meta.permissions().mode() & 0o111 == 0 {
                continue;
            }
        }
        return true;
    }
    false
}

/// Resolve the GUI editor for `path` (Edit ▸ Workbook (External) from a
/// native GUI backend). Mirrors [`crate::ui_core::opener_command`]: an
/// explicit `$CORRO_GUI_EDITOR` spec wins (words after the first are
/// arguments, `{}` is replaced with the path, otherwise the path is
/// appended; no shell is involved), else the platform GUI opener (`open`
/// on macOS, `notepad` on Windows). On Linux the desktop's native editor
/// is probed first (`$XDG_CURRENT_DESKTOP`: GNOME Text Editor on GNOME,
/// Kate on KDE, Xed on Cinnamon, Mousepad on Xfce), then the known-editor
/// list lightest-first, with `xdg-open` as the last resort (it often has
/// nothing registered for `.corro`, but may surprise on configured
/// systems).
///
/// The terminal `$VISUAL`/`$EDITOR` are deliberately ignored here: they
/// usually name terminal editors, which cannot run without a terminal.
pub fn gui_editor_command(path: &std::path::Path) -> (String, Vec<String>) {
    if let Ok(spec) = std::env::var("CORRO_GUI_EDITOR") {
        let spec = spec.trim();
        if !spec.is_empty() {
            let mut parts: Vec<String> =
                spec.split_whitespace().map(|s| s.to_string()).collect();
            let prog = parts.remove(0);
            let path_str = path.to_string_lossy().into_owned();
            if parts.iter().any(|a| a.contains("{}")) {
                for a in parts.iter_mut() {
                    if a.contains("{}") {
                        *a = a.replace("{}", &path_str);
                    }
                }
            } else {
                parts.push(path_str);
            }
            return (prog, parts);
        }
    }
    #[cfg(target_os = "windows")]
    {
        ("notepad".to_string(), vec![path.to_string_lossy().into_owned()])
    }
    #[cfg(target_os = "macos")]
    {
        ("open".to_string(), vec![path.to_string_lossy().into_owned()])
    }
    #[cfg(not(any(target_os = "windows", target_os = "macos")))]
    {
        let path_str = path.to_string_lossy().into_owned();
        // Native editor for this desktop first, then lightest-first, so a
        // GNOME box opens GNOME Text Editor and a Mint box opens Xed —
        // while a bare container with only VS Code still opens *something*.
        let mut ordered: Vec<&str> = desktop_preferred_editors();
        ordered.extend(KNOWN_GUI_EDITORS.iter().copied());
        if let Some(found) = ordered.into_iter().find(|e| find_in_path(e)) {
            return (found.to_string(), vec![path_str]);
        }
        ("xdg-open".to_string(), vec![path_str])
    }
}

/// Launch the GUI editor on `path` (native GUI backends only). Detached
/// like [`crate::ui_core::open_url_detached`]: stdin/stdout nulled so the
/// child can never scribble on the UI. Returns the program name for status
/// text.
///
/// Spawn success only means fork+exec worked — a missing display, an
/// opener with no handler (`xdg-open` on an unregistered type), or bad
/// args all die *after* spawn. So the child gets a short grace window
/// (~300 ms): an early nonzero exit is a real error (reported with the
/// child's last stderr line), an early zero exit is a handoff launcher
/// (`xdg-open`/`code`) done delegating, and a still-running child is the
/// editor window itself.
///
/// There is no "changed?" answer here — the editor hasn't even opened
/// yet. Callers report "opened" and let the log-tail poll pick up the
/// save, exactly like another window's Save landing mid-session.
pub fn open_gui_editor_detached(path: &std::path::Path) -> Result<String, String> {
    #[cfg(test)]
    if let Some(behavior) = TEST_EDITOR.with(|cell| cell.borrow().clone()) {
        return match behavior {
            TestEditor::Output(text) => {
                std::fs::write(path, text).map_err(|e| format!("test editor write: {e}"))?;
                Ok("test-editor".to_string())
            }
            TestEditor::Unchanged => Ok("test-editor".to_string()),
            TestEditor::Error(e) => Err(e),
        };
    }
    let (prog, args) = gui_editor_command(path);
    let mut child = std::process::Command::new(&prog)
        .args(&args)
        .stdin(std::process::Stdio::null())
        .stdout(std::process::Stdio::null())
        .stderr(std::process::Stdio::piped())
        .spawn()
        .map_err(|e| format!("launch {prog}: {e}"))?;
    let mut early: Option<std::process::ExitStatus> = None;
    for _ in 0..12 {
        match child.try_wait().map_err(|e| format!("wait {prog}: {e}"))? {
            Some(status) => {
                early = Some(status);
                break;
            }
            None => std::thread::sleep(std::time::Duration::from_millis(25)),
        }
    }
    match early {
        Some(status) if !status.success() => {
            let mut err = String::new();
            if let Some(mut stderr) = child.stderr.take() {
                use std::io::Read;
                let _ = stderr.read_to_string(&mut err);
            }
            let detail = err.lines().last().unwrap_or("").trim();
            if detail.is_empty() {
                Err(format!("{prog} exited with {status}"))
            } else {
                Err(format!("{prog} exited with {status}: {detail}"))
            }
        }
        _ => Ok(prog),
    }
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
        gui_editor: Option<String>,
        path: Option<String>,
        xdg_desktop: Option<String>,
    }
    impl EnvGuard {
        fn take() -> Self {
            let lock = ENV_LOCK.lock().unwrap();
            let visual = std::env::var("VISUAL").ok();
            let editor = std::env::var("EDITOR").ok();
            let gui_editor = std::env::var("CORRO_GUI_EDITOR").ok();
            let path = std::env::var("PATH").ok();
            let xdg_desktop = std::env::var("XDG_CURRENT_DESKTOP").ok();
            std::env::remove_var("VISUAL");
            std::env::remove_var("EDITOR");
            std::env::remove_var("CORRO_GUI_EDITOR");
            std::env::remove_var("XDG_CURRENT_DESKTOP");
            Self {
                _lock: lock,
                visual,
                editor,
                gui_editor,
                path,
                xdg_desktop,
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
            match &self.gui_editor {
                Some(e) => std::env::set_var("CORRO_GUI_EDITOR", e),
                None => std::env::remove_var("CORRO_GUI_EDITOR"),
            }
            match &self.path {
                Some(p) => std::env::set_var("PATH", p),
                None => std::env::remove_var("PATH"),
            }
            match &self.xdg_desktop {
                Some(d) => std::env::set_var("XDG_CURRENT_DESKTOP", d),
                None => std::env::remove_var("XDG_CURRENT_DESKTOP"),
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

    /// The GUI opener ignores the terminal $VISUAL/$EDITOR (they usually
    /// name terminal editors) and uses $CORRO_GUI_EDITOR, else the platform
    /// GUI opener. No shell is involved; `{}` marks the path position.
    #[test]
    fn gui_editor_command_prefers_override_then_platform_default() {
        let _g = EnvGuard::take();
        let path = std::path::Path::new("/tmp/book.corro");
        std::env::set_var("VISUAL", "vim");
        std::env::set_var("EDITOR", "vim");
        std::env::set_var("CORRO_GUI_EDITOR", "myedit --new-window");
        assert_eq!(
            gui_editor_command(path),
            (
                "myedit".to_string(),
                vec!["--new-window".to_string(), "/tmp/book.corro".to_string()]
            )
        );
        std::env::set_var("CORRO_GUI_EDITOR", "myedit --file {} --done");
        assert_eq!(
            gui_editor_command(path),
            (
                "myedit".to_string(),
                vec!["--file".to_string(), "/tmp/book.corro".to_string(), "--done".to_string()]
            )
        );
        std::env::remove_var("CORRO_GUI_EDITOR");
        #[cfg(target_os = "windows")]
        assert_eq!(
            gui_editor_command(path),
            ("notepad".to_string(), vec!["/tmp/book.corro".to_string()])
        );
        #[cfg(target_os = "macos")]
        assert_eq!(
            gui_editor_command(path),
            ("open".to_string(), vec!["/tmp/book.corro".to_string()])
        );
        // On Linux the probe runs: with an empty $PATH nothing is found, so
        // the last-resort xdg-open applies (whatever the desktop claims).
        #[cfg(not(any(target_os = "windows", target_os = "macos")))]
        {
            let empty = tempfile::tempdir().unwrap();
            std::env::set_var("PATH", empty.path());
            std::env::set_var("XDG_CURRENT_DESKTOP", "KDE");
            assert_eq!(
                gui_editor_command(path),
                ("xdg-open".to_string(), vec!["/tmp/book.corro".to_string()])
            );
        }
    }

    /// Linux probe: the desktop-native installed editor wins (Kate on KDE,
    /// Xed on Cinnamon), otherwise the lightest installed editor, otherwise
    /// xdg-open. A scratch $PATH keeps the test hermetic — resolution never
    /// spawns, it only stats.
    #[cfg(not(any(target_os = "windows", target_os = "macos")))]
    #[test]
    fn gui_editor_command_probes_desktop_then_lightest_first() {
        let _g = EnvGuard::take();
        let path = std::path::Path::new("/tmp/book.corro");
        let dir = tempfile::tempdir().unwrap();
        // Install kate, xed and mousepad — but not the GNOME default — so a
        // generic lightest-first pick would be mousepad (xed outranks it)
        // while KDE must still prefer its native kate.
        for name in ["kate", "xed", "mousepad"] {
            script(dir.path(), name, "#!/bin/sh\nexit 0\n");
        }
        std::env::set_var("PATH", dir.path());
        let cmd = |desktop: &str| {
            std::env::set_var("XDG_CURRENT_DESKTOP", desktop);
            gui_editor_command(path)
        };
        assert_eq!(
            cmd("KDE"),
            ("kate".to_string(), vec!["/tmp/book.corro".to_string()]),
            "KDE prefers its native Kate"
        );
        assert_eq!(
            cmd("X-Cinnamon"),
            ("xed".to_string(), vec!["/tmp/book.corro".to_string()]),
            "Cinnamon prefers its native Xed"
        );
        assert_eq!(
            cmd("XFCE"),
            ("mousepad".to_string(), vec!["/tmp/book.corro".to_string()]),
            "Xfce prefers its native Mousepad"
        );
        // Unknown desktop: generic lightest-first order (xed outranks
        // mousepad), even though kate is installed too.
        assert_eq!(
            cmd("BUDGIE"),
            ("xed".to_string(), vec!["/tmp/book.corro".to_string()]),
            "unknown desktop falls back to lightest-first"
        );
    }

    /// A heavyweight-only box still opens *something*: VS Code is picked
    /// ahead of the xdg-open last resort.
    #[cfg(not(any(target_os = "windows", target_os = "macos")))]
    #[test]
    fn gui_editor_command_prefers_heavyweight_over_xdg_open() {
        let _g = EnvGuard::take();
        let path = std::path::Path::new("/tmp/book.corro");
        let dir = tempfile::tempdir().unwrap();
        script(dir.path(), "code", "#!/bin/sh\nexit 0\n");
        std::env::set_var("PATH", dir.path());
        std::env::remove_var("XDG_CURRENT_DESKTOP");
        assert_eq!(
            gui_editor_command(path),
            ("code".to_string(), vec!["/tmp/book.corro".to_string()])
        );
    }

    /// The detached GUI launch returns the program name without waiting:
    /// a `true` opener reports success immediately and leaves the file
    /// alone (the save arrives later via the log-tail poll).
    #[cfg(unix)]
    #[test]
    fn gui_editor_launch_is_detached_and_leaves_the_file_for_the_tail() {
        let _g = EnvGuard::take();
        let dir = tempfile::tempdir().unwrap();
        let wb = dir.path().join("book.corro");
        std::fs::write(&wb, "CORRO_LOG 1\n").unwrap();
        std::env::set_var("CORRO_GUI_EDITOR", "true");
        assert_eq!(open_gui_editor_detached(&wb).unwrap(), "true");
        assert_eq!(std::fs::read_to_string(&wb).unwrap(), "CORRO_LOG 1\n");
    }

    #[cfg(unix)]
    #[test]
    fn gui_editor_launch_missing_program_is_an_error() {
        let _g = EnvGuard::take();
        let dir = tempfile::tempdir().unwrap();
        let wb = dir.path().join("book.corro");
        std::fs::write(&wb, "CORRO_LOG 1\n").unwrap();
        std::env::set_var("CORRO_GUI_EDITOR", "/nonexistent-gui-editor-xyz");
        assert!(open_gui_editor_detached(&wb).is_err());
    }

    /// Spawn success is not "opened": an opener that dies instantly (bad
    /// display, `xdg-open` with no handler) must surface as an error, with
    /// the child's own last stderr line for diagnosability.
    #[cfg(unix)]
    #[test]
    fn gui_editor_launch_instant_death_is_an_error_with_stderr() {
        let _g = EnvGuard::take();
        let dir = tempfile::tempdir().unwrap();
        let wb = dir.path().join("book.corro");
        std::fs::write(&wb, "CORRO_LOG 1\n").unwrap();
        let ed = script(
            dir.path(),
            "ed.sh",
            "#!/bin/sh\necho 'cannot open display:' >&2\nexit 1\n",
        );
        std::env::set_var("CORRO_GUI_EDITOR", &ed);
        let err = open_gui_editor_detached(&wb).unwrap_err();
        assert!(err.contains("cannot open display"), "got {err:?}");
    }

    /// A child that stays alive past the grace window is the editor window
    /// itself: reported opened without waiting for it to exit.
    #[cfg(unix)]
    #[test]
    fn gui_editor_launch_running_child_reports_opened() {
        let _g = EnvGuard::take();
        let dir = tempfile::tempdir().unwrap();
        let wb = dir.path().join("book.corro");
        std::fs::write(&wb, "CORRO_LOG 1\n").unwrap();
        let ed = script(dir.path(), "ed.sh", "#!/bin/sh\nsleep 2\n");
        std::env::set_var("CORRO_GUI_EDITOR", &ed);
        assert_eq!(open_gui_editor_detached(&wb).unwrap(), ed);
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
