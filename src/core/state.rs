use crate::grid::{CellAddr, MainRange, SelectionKind, SheetCursor, SortSpec};
use crate::io::{tail_apply_workbook, IoError, LogWatcher};
use crate::ops::{apply_log_line_to_workbook, Op, SheetState, WorkbookState};
use std::collections::HashMap;
use std::path::PathBuf;
use std::time::SystemTime;

/// Bottom-strip hints for the normal state, shared by every UI: the
/// ratatui bottom row always shows these (never status — that lives in
/// the formula bar trailing), and the GUI bottom strip shows exactly the
/// same text. Pure over four booleans so terminal (`ui::App`) and GUI
/// (`CoreApp`) share one source of truth instead of two hint tables.
/// Callers pass live state (undo/redo availability, save target); the
/// text must stay byte-identical everywhere it renders.
/// Bottom-strip hint for revision-browse mode, shared by every UI so the
/// text cannot drift. The ratatui reference renders this exact string in
/// `Mode::RevisionBrowse`, and both GUIs show it while `rev_browse` is set.
pub const REVISION_BROWSE_HINTS: &str =
    "  left/right·step revisions   Enter·close   Esc·close";

/// Status text when a Sheet context menu cannot be opened (no menu bar on the
/// backend, or its menus cannot be opened programmatically). Shared so the
/// GUI and pancurses backends report the same thing instead of drifting.
pub const SHEET_MENU_UNAVAILABLE: &str = "Sheet menu unavailable";

/// Status text when the aggregate picker is invoked on a cell that no margin
/// key governs. Shared across backends.
pub const NO_AGG_KEY_FOR_CELL: &str =
    "Aggregate: no margin TOTAL/MAX/… key for this cell";

pub fn normal_hints(anchored: bool, can_undo: bool, can_redo: bool, has_path: bool) -> String {
    if anchored {
        return "  r·move-rows   c·move-cols   v·deselect   Esc·cancel".into();
    }
    let mut hints = vec!["type/F2·edit", "Ctrl+C·copy", "Ctrl+X·cut", "Ctrl+V·paste"];
    if can_undo {
        hints.push("Ctrl+Z·undo");
    }
    if can_redo {
        hints.push("Ctrl+Y·redo");
    }
    hints.push("Ctrl+;·date");
    hints.push("Ctrl+:·time");
    hints.push(if has_path {
        "Ctrl+S·save"
    } else {
        "Ctrl+S·save as"
    });
    hints.push("F1·help");
    format!("  {}", hints.join("; "))
}

pub struct CoreApp {
    pub path: Option<PathBuf>,
    pub import_source: Option<PathBuf>,
    pub source_path: Option<PathBuf>,
    pub revision_limit: Option<usize>,
    pub revision_browse: bool,
    pub revision_browse_limit: usize,
    pub offset: u64,
    pub state: SheetState,
    pub workbook: WorkbookState,
    pub cursor: SheetCursor,
    pub anchor: Option<SheetCursor>,
    /// Which axes the anchor↔cursor selection covers. `Cells` (the default)
    /// is the classic rectangle; `Rows`/`Cols` are set by a gutter-header
    /// click (or the row/column selection commands) so the whole main-body
    /// row/column highlights and only the matching gutter headers glow.
    /// Shared by every non-ratatui backend, so the widget-tree GUI and
    /// pancurses agree on what a selection means.
    pub selection_kind: SelectionKind,
    pub watcher: Option<LogWatcher>,
    pub status: String,
    pub ops_applied: usize,
    pub op_history: Vec<Op>,
    pub redo_history: Vec<Op>,
    pub view_sheet_id: u32,
    pub persisted_view_sort_cols: HashMap<u32, Vec<SortSpec>>,
    pub linked_source_mtimes: HashMap<PathBuf, SystemTime>,
    pub unsaved_file: Option<PathBuf>,
    pub unsaved_auto_create: bool,
    pub exit_message: Option<String>,
    pub clipboard_snapshot: Option<(MainRange, String)>,
    pub edit_target_addr: Option<CellAddr>,
    pub edit_range_addrs: Option<Vec<CellAddr>>,
    pub pending_lost_edit: Option<(CellAddr, String)>,
    pub pending_fit_to_content_on_commit: bool,
    /// Frozen rows/columns (the padlock feature). Session-only view state,
    /// never written to the log, and cleared when the active sheet changes
    /// (see [`crate::lock::LockState::ensure_sheet`]).
    pub locks: crate::lock::LockState,
}

/// Serialises tests that install the unsaved-file environment variables
/// (`CORRO_UNSAVED_TEST_DIR`, `CORRO_AUTO_UNSAVED_TEST`, `XDG_STATE_HOME`).
///
/// `std::env` is process-global, so two tests that each point the unsaved
/// directory at their own tempdir race: one reads the other's directory and
/// the failure looks like a product bug. Every test in the crate that touches
/// those variables holds this lock for its whole body (the `ui` module's
/// `UnsavedEnv` guard reuses it too).
///
/// Declared unconditionally (not `#[cfg(test)]`) so the `ui` module's test
/// guard can alias it regardless of how the crate is compiled; outside tests
/// nothing references it.
#[allow(dead_code)]
pub(crate) static UNSAVED_ENV_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

/// Log-tail polling for backends that drive [`CoreApp`] directly (the GUI).
///
/// Twin of the log-tail half of `ui::App::sync_external` (which the TUI
/// loop calls every iteration): when another window appends revisions to
/// the bound file, tail-apply them here so both windows show the same
/// cells. Commits append to the log immediately, so mirroring needs no
/// save step on either side. Keep the two in sync: same watcher-then-size
/// check order, same tail-apply-then-extent-restore sequence, same status
/// texts, same full-reload fallback when the file was rewritten (Save)
/// instead of appended to.
/// Outcome of one revision step, so callers can report it (status line,
/// on-canvas indicator) without re-deriving the arithmetic.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RevisionStep {
    /// The limit moved and the log was re-replayed at the new revision.
    Moved,
    /// Already at revision 1 and asked to go back — the reference clamps
    /// rather than wrapping, so this is a no-op, not an error.
    AtOldest,
    /// Asked to go forward past the last revision. The re-replay is still
    /// performed (the file may have grown since), but the limit only ever
    /// lands on a real revision; if it did not move, nothing changed.
    AtNewest,
    /// No `source_path` to re-replay (e.g. the file was renamed away).
    NoSource,
}

/// A replay's status line under an arbitrary verb, shared by every UI so the
/// text cannot drift between the ratatui reference and the GUIs.
///
/// `PartialReplay` reports a truncated log (`failed_line` + `error`) rather
/// than an `Err`, so a partially-played file still says *where* the play
/// stopped; the plain case names the revision reached. The reference uses
/// "Replayed" for a full File▸Replay and "Loaded"/"Browsing" elsewhere, so
/// the prefix is the caller's.
pub fn replay_status(prefix: &str, path: &std::path::Path, replay: &crate::io::PartialReplay) -> String {
    match (replay.failed_line, replay.error.as_deref()) {
        (Some(line), Some(err)) => format!(
            "{prefix} {} @ revision {} stopped at line {line}: {err}",
            path.display(),
            replay.op_count
        ),
        _ => format!("{prefix} {} @ revision {}", path.display(), replay.op_count),
    }
}

/// The status line for a browsed revision (see [`replay_status`]).
pub fn revision_browse_status(
    path: &std::path::Path,
    replay: &crate::io::PartialReplay,
) -> String {
    replay_status("Browsing", path, replay)
}

/// Outcome of one handled [`RevisionStep`], so a UI can log what the
/// keystroke did without re-deriving the arithmetic from the step itself.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RevisionStepAction {
    /// The limit moved and the log was re-replayed at the new revision.
    Moved,
    /// Already at the oldest (Left) / newest (Right) revision: a clamped
    /// no-op, not an error.
    AtBoundary,
    /// No `source_path` to re-replay (e.g. the file was renamed away).
    NoSource,
    /// The re-replay failed; the message is in `CoreApp::status`.
    Errored,
}

/// The name [`RevisionStepAction`] is logged under, so every backend's log
/// line uses the same vocabulary.
impl RevisionStepAction {
    pub fn name(self) -> &'static str {
        match self {
            RevisionStepAction::Moved => "moved",
            RevisionStepAction::AtBoundary => "at_boundary",
            RevisionStepAction::NoSource => "no_source",
            RevisionStepAction::Errored => "error",
        }
    }
}

/// The shared body of a revision-browse reload, over the state every
/// backend keeps (ratatui's `App` holds these in separate fields rather
/// than a nested `CoreApp`).
///
/// The reference resets more than the workbook: a stale active-sheet
/// cache, a cursor past the new extent, or a column width fitted to a
/// revision that no longer exists all show up as visible corruption, so
/// every one of those is reset here too.
///
/// `revision_browse_limit` is read as the *requested* revision and, on
/// return, holds the revision actually reached (the re-replay clamps it
/// to the log's real op count, which is how Right past the end becomes a
/// no-op rather than an empty sheet).
#[allow(clippy::too_many_arguments)]
pub fn reload_revision_browse_fields(
    source_path: &mut Option<PathBuf>,
    revision_browse_limit: &mut usize,
    workbook: &mut WorkbookState,
    state: &mut SheetState,
    cursor: &mut SheetCursor,
    anchor: &mut Option<SheetCursor>,
    selection_kind: &mut SelectionKind,
    path: &mut Option<PathBuf>,
    watcher: &mut Option<LogWatcher>,
    view_sheet_id: &mut u32,
    offset: &mut u64,
    ops_applied: &mut usize,
    status: &mut String,
) -> Result<(), IoError> {
    let Some(src) = source_path.clone() else {
        return Ok(());
    };
    *workbook = WorkbookState::new();
    *state = SheetState::new(1, 1);
    *cursor = SheetCursor {
        row: crate::grid::HEADER_ROWS,
        col: crate::grid::MARGIN_COLS,
    };
    *anchor = None;
    // A collapsed selection also drops its Rows/Cols coverage, or a reloaded
    // revision would inherit the previous row/column highlight.
    *selection_kind = SelectionKind::Cells;
    // A browsed revision set is detached from the on-disk log: no commit
    // path and no watcher, exactly as the reference does, so stepping
    // revisions can never append to the file being read.
    *path = None;
    *watcher = None;

    let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
    let requested_limit = *revision_browse_limit;
    let (off, replay) = crate::io::load_workbook_revisions_partial(
        &src,
        requested_limit,
        workbook,
        &mut active_sheet,
    )?;

    // Re-resolve the active sheet against the *new* workbook: the
    // reloaded revision set may not contain the sheet the cache pointed
    // at.
    *view_sheet_id = active_sheet;
    workbook.ensure_active_sheet();
    if let Some(idx) = workbook.sheet_index_by_id(active_sheet) {
        workbook.active_sheet = idx;
        *state = workbook.sheets[idx].state.clone();
    }

    // Re-fit columns to the values at this revision. Without this the
    // grid keeps the widths fitted to the previous revision.
    let main_cols = state.grid.main_cols();
    for c in 0..main_cols {
        let global_col = crate::grid::MARGIN_COLS + c;
        if let Some(rw) = crate::ui_core::rendered_width_for_column(&state.grid, global_col) {
            let capped = rw.min(state.grid.max_col_width());
            state.grid.set_col_width(global_col, Some(capped));
        }
    }

    *offset = off;
    *ops_applied = replay.op_count;
    *revision_browse_limit = replay.op_count;
    *status = revision_browse_status(&src, &replay);
    cursor.clamp(&state.grid);
    Ok(())
}

/// One revision step's clamped-boundary arithmetic over the raw fields, so
/// a backend that keeps browse state outside a [`CoreApp`] (the ratatui
/// reference's `App`) runs the same Left/Right semantics.
///
/// Sets `revision_browse_limit` to the *requested* revision and returns:
/// * [`RevisionStep::NoSource`] / [`RevisionStep::AtOldest`] when the step
///   is refused outright (nothing to re-replay, or Left at revision 1);
/// * [`RevisionStep::Moved`] when the caller should now re-replay
///   (see [`CoreApp::step_revision_back`]). The re-replay clamps the
///   limit to the revision actually reached, so the caller classifies the
///   result afterwards with [`revision_step_outcome`].
pub fn revision_step_request(
    back: bool,
    source_path: &Option<PathBuf>,
    revision_browse_limit: &mut usize,
) -> RevisionStep {
    if source_path.is_none() {
        return RevisionStep::NoSource;
    }
    // Left at revision 1 is the reference's `limit > 1` guard: a no-op,
    // never a wrap to the newest revision.
    if back && *revision_browse_limit <= 1 {
        return RevisionStep::AtOldest;
    }
    if back {
        *revision_browse_limit -= 1;
    } else {
        // Right past the newest revision saturates; the clamp proper lives
        // in the re-replay, so this only re-reads the same revision (and
        // picks up appends).
        *revision_browse_limit = revision_browse_limit.saturating_add(1);
    }
    RevisionStep::Moved
}

/// Classify a step once its re-replay has run: `Moved` iff the limit came
/// out different from `before` (the re-replay clamps it).
pub fn revision_step_outcome(back: bool, before: usize, after: usize) -> RevisionStep {
    if after != before {
        RevisionStep::Moved
    } else if back {
        RevisionStep::AtOldest
    } else {
        RevisionStep::AtNewest
    }
}

impl CoreApp {
    /// Re-replay `source_path` at `revision_browse_limit`.
    ///
    /// The shared bookkeeping lives in [`reload_revision_browse_fields`]
    /// so every backend (this CoreApp, plus the ratatui reference, which
    /// keeps the same state in separate fields) runs one implementation.
    pub fn reload_revision_browse(&mut self) -> Result<(), IoError> {
        reload_revision_browse_fields(
            &mut self.source_path,
            &mut self.revision_browse_limit,
            &mut self.workbook,
            &mut self.state,
            &mut self.cursor,
            &mut self.anchor,
            &mut self.selection_kind,
            &mut self.path,
            &mut self.watcher,
            &mut self.view_sheet_id,
            &mut self.offset,
            &mut self.ops_applied,
            &mut self.status,
        )
    }

    /// Step backwards one revision (Left arrow in revision-browse mode).
    pub fn step_revision_back(&mut self) -> Result<RevisionStep, IoError> {
        let before = self.revision_browse_limit;
        let requested = revision_step_request(true, &self.source_path, &mut self.revision_browse_limit);
        if requested != RevisionStep::Moved {
            return Ok(requested);
        }
        self.reload_revision_browse()?;
        Ok(revision_step_outcome(true, before, self.revision_browse_limit))
    }

    /// Step forwards one revision (Right arrow in revision-browse mode).
    pub fn step_revision_forward(&mut self) -> Result<RevisionStep, IoError> {
        let before = self.revision_browse_limit;
        let requested = revision_step_request(false, &self.source_path, &mut self.revision_browse_limit);
        if requested != RevisionStep::Moved {
            return Ok(requested);
        }
        self.reload_revision_browse()?;
        Ok(revision_step_outcome(false, before, self.revision_browse_limit))
    }

    /// One arrow-key revision step, with the user-visible reporting folded
    /// in: the status line a backend would set is set here, and the result
    /// is reduced to what a keystroke handler needs to log.
    ///
    /// Both GUIs drive this so the left/right semantics, the clamped
    /// boundaries and the "Browsing: no source log to step" /
    /// "Browsing error: …" texts are identical everywhere; a handler that
    /// only needs the new state does nothing else.
    pub fn step_revision(&mut self, back: bool) -> RevisionStepAction {
        let result = if back {
            self.step_revision_back()
        } else {
            self.step_revision_forward()
        };
        match result {
            Ok(RevisionStep::Moved) => RevisionStepAction::Moved,
            Ok(RevisionStep::AtOldest | RevisionStep::AtNewest) => RevisionStepAction::AtBoundary,
            Ok(RevisionStep::NoSource) => {
                self.status = "Browsing: no source log to step".into();
                RevisionStepAction::NoSource
            }
            Err(e) => {
                self.status = format!("Browsing error: {e}");
                RevisionStepAction::Errored
            }
        }
    }

    pub fn poll_log_tail(&mut self) -> Result<bool, IoError> {
        let Some(ref p) = self.path.clone() else {
            return Ok(false);
        };
        // notify cannot watch a path that does not exist yet; (re)try until
        // it does. Absent watcher => the size check below still fires.
        if self.watcher.is_none() {
            if let Ok(w) = LogWatcher::new(p.clone()) {
                self.watcher = Some(w);
            }
        }
        let mut should_tail = false;
        if let Some(w) = &self.watcher {
            if w.poll_dirty() {
                should_tail = true;
            }
        }
        if !should_tail {
            if let Ok(meta) = std::fs::metadata(p) {
                if meta.len() > self.offset {
                    should_tail = true;
                }
            }
        }
        if !should_tail {
            return Ok(false);
        }
        // Save the in-memory extent so we can restore it after the reload.
        let saved_rows = self.state.grid.main_rows();
        let saved_cols = self.state.grid.main_cols();
        match tail_apply_workbook(p, self.offset, &mut self.workbook, &mut self.view_sheet_id) {
            Ok(new_off) => {
                if new_off > self.offset {
                    self.offset = new_off;
                    self.sync_sheet_cache();
                    let cur_rows = self.state.grid.main_rows();
                    let cur_cols = self.state.grid.main_cols();
                    if saved_rows > cur_rows || saved_cols > cur_cols {
                        self.state
                            .grid
                            .set_main_size(saved_rows.max(cur_rows), saved_cols.max(cur_cols));
                    }
                    self.status = "External change applied".into();
                    Ok(true)
                } else {
                    self.offset = new_off;
                    Ok(false)
                }
            }
            Err(_) => {
                // Counterpart rewrote the file (Save) instead of appending:
                // fall back to a full reload.
                let data = std::fs::read_to_string(p)?;
                let mut workbook = WorkbookState::new();
                let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
                for line in data.lines() {
                    let t = line.trim();
                    if t.is_empty() {
                        continue;
                    }
                    apply_log_line_to_workbook(t, &mut workbook, &mut active_sheet)?;
                }
                self.workbook = workbook;
                self.view_sheet_id = active_sheet;
                self.sync_sheet_cache();
                let cur_rows = self.state.grid.main_rows();
                let cur_cols = self.state.grid.main_cols();
                if saved_rows > cur_rows || saved_cols > cur_cols {
                    self.state
                        .grid
                        .set_main_size(saved_rows.max(cur_rows), saved_cols.max(cur_cols));
                }
                self.offset = data.len() as u64;
                self.ops_applied = data.lines().filter(|line| !line.trim().is_empty()).count();
                self.status = "File reset; full reload".into();
                Ok(true)
            }
        }
    }

    fn sync_sheet_cache(&mut self) {
        self.workbook.ensure_active_sheet();
        if let Some(idx) = self.workbook.sheet_index_by_id(self.view_sheet_id) {
            self.workbook.active_sheet = idx;
            self.state = self.workbook.sheets[idx].state.clone();
        } else {
            self.view_sheet_id = self.workbook.sheet_id(self.workbook.active_sheet);
            self.state = self.workbook.active_sheet().clone();
        }
    }

    // ── Shared operations ──────────────────────────────────────────────
    //
    // These live on `CoreApp` so every backend (TUI, GUI, pancurses, …)
    // uses the same commit/save/unsaved-file logic instead of each
    // re-implementing it with subtle drift.

    /// Directory for unsaved per-user `.corro` files.
    ///
    /// `CORRO_UNSAVED_TEST_DIR` overrides for tests.  Platform-specific
    /// defaults: `$XDG_STATE_HOME/corro/unsaved` on Linux,
    /// `~/Library/Application Support/corro/unsaved` on macOS,
    /// `%LOCALAPPDATA%\corro\unsaved` on Windows.
    pub fn default_unsaved_dir() -> PathBuf {        if let Ok(test_dir) = std::env::var("CORRO_UNSAVED_TEST_DIR") {
            return PathBuf::from(test_dir);
        }
        if cfg!(target_os = "linux") {
            if let Ok(x) = std::env::var("XDG_STATE_HOME") {
                return PathBuf::from(x).join("corro/unsaved");
            }
            if let Ok(home) = std::env::var("HOME") {
                return PathBuf::from(home).join(".corro/unsaved");
            }
        }
        if cfg!(target_os = "macos") {
            if let Ok(home) = std::env::var("HOME") {
                return PathBuf::from(home).join("Library/Application Support/corro/unsaved");
            }
        }
        if cfg!(target_os = "windows") {
            if let Ok(local) = std::env::var("LOCALAPPDATA") {
                return PathBuf::from(local).join("corro\\unsaved");
            }
            if let Ok(appdata) = std::env::var("APPDATA") {
                return PathBuf::from(appdata).join("corro\\unsaved");
            }
        }
        if let Ok(home) = std::env::var("HOME") {
            return PathBuf::from(home).join(".corro/unsaved");
        }
        std::env::temp_dir().join("corro/unsaved")
    }

    /// Ensure `path` has a `.corro` extension (TUI parity).
    pub fn to_corro_path(path: &std::path::Path) -> PathBuf {
        if path.extension().and_then(|e| e.to_str()) == Some("corro") {
            path.to_path_buf()
        } else {
            let mut p = path.to_path_buf();
            p.set_extension("corro");
            p
        }
    }

    /// Best-guess source path for naming unsaved files.
    fn preferred_import_source_path(&self) -> Option<&std::path::Path> {
        self.path
            .as_deref()
            .or(self.import_source.as_deref())
            .or_else(|| {
                self.workbook
                    .sheets
                    .iter()
                    .find_map(|sheet| sheet.linked_source.as_ref().map(|source| source.path.as_path()))
            })
    }

    /// Full initial log text for a freshly created unsaved file.
    fn initial_unsaved_log_text(&self) -> String {
        crate::io::serialize_workbook_log(&self.workbook, &self.persisted_view_sort_cols)
    }

    /// Create an on-disk unsaved `.corro` file and bind it to `self.path`.
    ///
    /// If `self.path` is already set, returns it immediately.  Otherwise
    /// creates a new file in [`Self::default_unsaved_dir`], writes the full
    /// workbook serialization, and binds `self.path` and `self.unsaved_file`
    /// to it.  All subsequent commits append to this file, and
    /// [`Self::save_to_path`] can rename it to the user's chosen destination.
    pub fn ensure_unsaved_file(&mut self) -> Result<PathBuf, String> {
        if let Some(ref p) = self.path.clone() {
            return Ok(p.clone());
        }

        let candidate_dir = std::env::var("CORRO_UNSAVED_TEST_DIR")
            .ok()
            .map(PathBuf::from)
            .unwrap_or_else(Self::default_unsaved_dir);

        let base = self
            .preferred_import_source_path()
            .and_then(|p| p.file_stem())
            .and_then(|os| os.to_str())
            .filter(|s| !s.is_empty())
            .map(|s| s.to_string())
            .unwrap_or_else(|| "untitled".to_string());

        // Try base-first filenames in candidate_dir.
        if std::fs::create_dir_all(&candidate_dir).is_ok() {
            for i in 0..1000 {
                let name = if i == 0 {
                    format!("{}.corro", base)
                } else {
                    format!("{}_{}.corro", base, i)
                };
                let cand = candidate_dir.join(&name);
                match std::fs::OpenOptions::new()
                    .create_new(true)
                    .write(true)
                    .open(&cand)
                {
                    Ok(_) => {
                        let initial = self.initial_unsaved_log_text();
                        std::fs::write(&cand, initial.as_bytes())
                            .map_err(|e| format!("failed to write unsaved file: {e}"))?;
                        self.unsaved_file = Some(cand.clone());
                        self.path = Some(cand.clone());
                        self.exit_message = None;
                        self.status = format!("Created unsaved file: {}", cand.display());
                        let meta = std::fs::metadata(&cand)
                            .map_err(|e| format!("stat unsaved file: {e}"))?;
                        self.offset = meta.len();
                        return Ok(cand);
                    }
                    Err(e) if e.kind() == std::io::ErrorKind::AlreadyExists => continue,
                    Err(_) => break,
                }
            }
        }

        // Fallback: timestamped filename in the unsaved dir.
        let dir = Self::default_unsaved_dir();
        std::fs::create_dir_all(&dir)
            .map_err(|e| format!("create unsaved dir: {e}"))?;
        let source_basename = self
            .preferred_import_source_path()
            .and_then(|p| p.file_name())
            .and_then(|os| os.to_str())
            .map(|s| {
                let mut out = String::with_capacity(s.len());
                for ch in s.chars() {
                    if ch.is_ascii_alphanumeric() || ch == '-' || ch == '_' || ch == '.' {
                        out.push(ch);
                    } else {
                        out.push('_');
                    }
                }
                if out.is_empty() { s.to_string() } else { out }
            });

        for _ in 0..10 {
            let pid = std::process::id();
            let now = std::time::SystemTime::now()
                .duration_since(std::time::SystemTime::UNIX_EPOCH)
                .map(|d| d.as_nanos())
                .unwrap_or(0);
            let name = if let Some(ref src) = source_basename {
                format!("unsaved-{}-{}-{}.corro", pid, src, now)
            } else {
                format!("unsaved-{}-{}.corro", pid, now)
            };
            let cand = dir.join(&name);
            match std::fs::OpenOptions::new()
                .create_new(true)
                .write(true)
                .open(&cand)
            {
                Ok(_) => {
                    let initial = self.initial_unsaved_log_text();
                    std::fs::write(&cand, initial.as_bytes())
                        .map_err(|e| format!("failed to write unsaved file: {e}"))?;
                    self.unsaved_file = Some(cand.clone());
                    self.path = Some(cand.clone());
                    self.exit_message = None;
                    self.status = format!("Created unsaved file: {}", cand.display());
                    let meta = std::fs::metadata(&cand)
                        .map_err(|e| format!("stat unsaved file: {e}"))?;
                    self.offset = meta.len();
                    return Ok(cand);
                }
                Err(e) if e.kind() == std::io::ErrorKind::AlreadyExists => continue,
                Err(e) => return Err(format!("create unsaved file: {e}")),
            }
        }
        Err("could not create unsaved file (too many attempts)".into())
    }

    /// Commit a sheet-level op to disk (or apply in-memory if no path).
    ///
    /// If `self.path` is `None` and `self.unsaved_auto_create` is `true`,
    /// an unsaved file is created first so the op is persisted immediately.
    /// Returns `Ok(true)` if persisted to disk, `Ok(false)` if applied
    /// in-memory only.
    pub fn commit_sheet_op(&mut self, op: Op) -> Result<bool, String> {
        if self.path.is_none() && self.unsaved_auto_create {
            let _ = self.ensure_unsaved_file()?;
        }

        if let Some(ref p) = self.path.clone() {
            // Resolve the sheet to commit against through the workbook, not
            // a possibly-stale `view_sheet_id`: a view id that no longer
            // exists (a sheet deleted, or a cache never seeded — a fresh GUI
            // app used to start at id 0, which is not a sheet) would make
            // `commit_workbook_op` fail with "unknown sheet id" and silently
            // drop the edit. `view_sheet_id` is the intent; this is the
            // guarantee.
            let sheet_id = if self.workbook.sheet_index_by_id(self.view_sheet_id).is_some() {
                self.view_sheet_id
            } else {
                self.workbook.sheet_id(self.workbook.active_sheet)
            };
            let wbo = crate::ops::WorkbookOp::SheetOp {
                sheet_id,
                op,
            };
            let mut active_sheet = sheet_id;
            crate::io::commit_workbook_op(
                p,
                &mut self.offset,
                &mut self.workbook,
                &mut active_sheet,
                &wbo,
            )
            .map_err(|e| format!("commit_workbook_op: {e}"))?;
            self.ops_applied = self.ops_applied.saturating_add(1);
            self.sync_sheet_cache();
            Ok(true)
        } else {
            // In-memory only.
            op.apply(&mut self.state);
            Ok(false)
        }
    }

    /// Save the workbook to `path`, moving an unsaved file if possible.
    ///
    /// If the current file lives under the unsaved directory it is
    /// atomically renamed (or copied+renamed across devices).  Otherwise
    /// the in-memory workbook is serialized to `path`.
    pub fn save_to_path(&mut self, path: &std::path::Path) -> Result<(), String> {
        let path = Self::to_corro_path(path);

        // Fast-path: move an unsaved file to the destination.
        let unsaved_dir = Self::default_unsaved_dir();
        let cur = self
            .path
            .clone()
            .or_else(|| self.unsaved_file.clone())
            .filter(|p| p.exists() && p.ancestors().any(|a| a == unsaved_dir.as_path()));
        if let Some(cur) = cur {
            if let Some(parent) = path.parent() {
                std::fs::create_dir_all(parent)
                    .map_err(|e| format!("create dir: {e}"))?;
            }
            match std::fs::rename(&cur, &path) {
                Ok(()) => {
                    self.path = Some(path.clone());
                    self.import_source = None;
                    self.source_path = None;
                    self.revision_limit = None;
                    self.unsaved_file = None;
                    self.status = format!("Saved {}", path.display());
                    if let Ok(w) = LogWatcher::new(path.clone()) {
                        self.watcher = Some(w);
                    }
                    let meta = std::fs::metadata(&path)
                        .map_err(|e| format!("stat saved file: {e}"))?;
                    self.offset = meta.len();
                    return Ok(());
                }
                Err(e) => {
                    // Cross-device: copy + rename.
                    if e.raw_os_error() == Some(libc::EXDEV) {
                        let tmp = crate::io::temp_sibling_path(path.parent(), &path);
                        std::fs::copy(&cur, &tmp)
                            .map_err(|e| format!("copy unsaved: {e}"))?;
                        if path.exists() {
                            std::fs::remove_file(&path)
                                .map_err(|e| format!("remove old: {e}"))?;
                        }
                        std::fs::rename(&tmp, &path)
                            .map_err(|e| format!("rename tmp: {e}"))?;
                        let _ = std::fs::remove_file(&cur);
                        self.path = Some(path.clone());
                        self.import_source = None;
                        self.source_path = None;
                        self.revision_limit = None;
                        self.unsaved_file = None;
                        self.status = format!("Saved {}", path.display());
                        if let Ok(w) = LogWatcher::new(path.clone()) {
                            self.watcher = Some(w);
                        }
                        let meta = std::fs::metadata(&path)
                            .map_err(|e| format!("stat saved file: {e}"))?;
                        self.offset = meta.len();
                        return Ok(());
                    }
                    // Other errors fall through to serialize.
                }
            }
        }

        // Fallback: serialize the workbook to the destination.
        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent)
                .map_err(|e| format!("create dir: {e}"))?;
        }
        let tmp = crate::io::temp_sibling_path(path.parent(), &path);
        {
            use std::io::Write;
            let mut f = std::fs::File::create(&tmp)
                .map_err(|e| format!("create temp: {e}"))?;
            let data = self.initial_unsaved_log_text();
            f.write_all(data.as_bytes())
                .map_err(|e| format!("write temp: {e}"))?;
            f.flush().map_err(|e| format!("flush temp: {e}"))?;
        }
        if path.exists() {
            std::fs::remove_file(&path)
                .map_err(|e| format!("remove old: {e}"))?;
        }
        std::fs::rename(&tmp, &path)
            .map_err(|e| format!("rename to dest: {e}"))?;

        self.path = Some(path.clone());
        self.import_source = None;
        self.source_path = None;
        self.revision_limit = None;
        self.unsaved_file = None;
        self.status = format!("Saved {}", path.display());
        if let Ok(w) = LogWatcher::new(path.clone()) {
            self.watcher = Some(w);
        }
        let meta = std::fs::metadata(&path)
            .map_err(|e| format!("stat saved file: {e}"))?;
        self.offset = meta.len();
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::normal_hints;
    use super::CoreApp;
    use super::RevisionStep;
    use crate::grid::{CellAddr, SelectionKind};
    use crate::io::{commit_workbook_op, load_workbook_revisions_partial};
    use crate::ops::{apply_workbook_op, Op, WorkbookOp, WorkbookState};
    use std::path::PathBuf;

    fn test_core_app(path: Option<std::path::PathBuf>, workbook: WorkbookState, offset: u64) -> CoreApp {
        super::CoreApp {
            path,
            import_source: None,
            source_path: None,
            revision_limit: None,
            revision_browse: false,
            revision_browse_limit: 0,
            offset,
            state: Default::default(),
            workbook,
            cursor: crate::grid::SheetCursor { row: 0, col: 0 },
            anchor: None,
            selection_kind: SelectionKind::Cells,
            watcher: None,
            ops_applied: 0,
            op_history: Vec::new(),
            redo_history: Vec::new(),
            view_sheet_id: 0,
            persisted_view_sort_cols: Default::default(),
            linked_source_mtimes: Default::default(),
            unsaved_file: None,
            unsaved_auto_create: false,
            status: String::new(),
            exit_message: None,
            clipboard_snapshot: None,
            edit_target_addr: None,
            edit_range_addrs: None,
            pending_lost_edit: None,
            pending_fit_to_content_on_commit: false,
            locks: Default::default(),
        }
    }

    /// CoreApp twin of the two-window mirror: B.poll_log_tail() picks up
    /// A's committed cell. This is the exact path the GUI keystroke hook
    /// drives (gui_backend polls app.core on every key).
    #[test]
    fn core_second_window_tails_first_windows_commit() {
        let dir = tempfile::tempdir().unwrap();
        let path = dir.path().join("shared.corro");
        let cell = CellAddr::Main { row: 0, col: 0 };
        // Window A opens a fresh sheet bound to the file.
        let mut a = test_core_app(Some(path.clone()), WorkbookState::new_seeded(), 0);
        let mut a_active = a.workbook.sheets[0].id;
        let seed = WorkbookOp::SheetOp {
            sheet_id: a_active,
            op: Op::SetCell {
                addr: cell.clone(),
                value: "seed".into(),
            },
        };
        apply_workbook_op(&mut a.workbook, &mut a_active, seed.clone()).unwrap();
        commit_workbook_op(&path, &mut a.offset, &mut a.workbook, &mut a_active, &seed).unwrap();
        // Window B opens the file (full load, offset at end).
        let mut b_wb = WorkbookState::new();
        let mut b_active = 0u32;
        let (off, _) =
            load_workbook_revisions_partial(&path, usize::MAX, &mut b_wb, &mut b_active).unwrap();
        let mut b = test_core_app(Some(path.clone()), b_wb, off);
        b.view_sheet_id = b_active;
        // Window A commits a new cell (append to log + apply in memory).
        let op = WorkbookOp::SheetOp {
            sheet_id: a_active,
            op: Op::SetCell {
                addr: cell.clone(),
                value: "MIRROR".into(),
            },
        };
        apply_workbook_op(&mut a.workbook, &mut a_active, op.clone()).unwrap();
        commit_workbook_op(&path, &mut a.offset, &mut a.workbook, &mut a_active, &op).unwrap();
        // B is stale until it polls ...
        assert_ne!(
            b.workbook.sheets[0].state.grid.get(&cell).as_deref(),
            Some("MIRROR")
        );
        // ... then the tail applies A's commit with no save step.
        assert!(b.poll_log_tail().unwrap(), "B should observe A's commit");
        assert_eq!(
            b.workbook.sheets[0].state.grid.get(&cell).as_deref(),
            Some("MIRROR")
        );
        assert_eq!(b.offset, a.offset);
        assert_eq!(b.status, "External change applied");
        // A second poll is quiet.
        assert!(!b.poll_log_tail().unwrap());
    }

    #[test]
    fn normal_hints_covers_anchor_undo_redo_and_save_states() {
        let fresh = normal_hints(false, false, false, false);
        assert!(fresh.contains("type/F2·edit"), "{fresh:?}");
        assert!(fresh.contains("Ctrl+S·save as"), "{fresh:?}");
        assert!(!fresh.contains("undo"), "{fresh:?}");
        let undo = normal_hints(false, true, false, true);
        assert!(undo.contains("Ctrl+Z·undo"), "{undo:?}");
        assert!(undo.contains("Ctrl+S·save"), "{undo:?}");
        assert!(!undo.contains("save as"), "{undo:?}");
        let redo = normal_hints(false, false, true, false);
        assert!(redo.contains("Ctrl+Y·redo"), "{redo:?}");
        let anchored = normal_hints(true, true, true, true);
        assert!(anchored.contains("move-rows"), "{anchored:?}");
        assert!(!anchored.contains("Ctrl+C"), "{anchored:?}");
    }

    /// A three-revision log, browsable from `CoreApp` exactly as
    /// File▸Replay leaves it: detached source, limit at the op count.
    fn browsable_core_app(dir: &tempfile::TempDir) -> (CoreApp, PathBuf) {
        let path = dir.path().join("browse.corro");
        // Three real revisions: A1, then B1, then A2 overwritten.
        std::fs::write(
            &path,
            "SET $1:A1 one\nSET $1:B1 two\nSET $1:A1 three\n",
        )
        .unwrap();
        let mut app = test_core_app(None, WorkbookState::new(), 0);
        app.source_path = Some(path.clone());
        app.revision_browse = true;
        app.revision_browse_limit = 3;
        app.reload_revision_browse().unwrap();
        (app, path)
    }

    fn main_cell(app: &CoreApp, row: u32, col: u32) -> Option<String> {
        app.state
            .grid
            .get(&CellAddr::Main { row, col })
            .map(|s| s.to_string())
    }

    /// The whole point of the feature: Left walks the log backwards and
    /// Right walks it forwards, showing the *cells at that revision*.
    #[test]
    fn revision_step_changes_the_cells_through_the_log() {
        let dir = tempfile::tempdir().unwrap();
        let (mut app, _) = browsable_core_app(&dir);

        // Fully replayed: the last write wins.
        assert_eq!(app.revision_browse_limit, 3);
        assert_eq!(main_cell(&app, 0, 0).as_deref(), Some("three"));
        assert_eq!(main_cell(&app, 0, 1).as_deref(), Some("two"));

        // Left -> revision 2: A1 is "one" again, B1 still set.
        assert_eq!(app.step_revision_back().unwrap(), RevisionStep::Moved);
        assert_eq!(app.revision_browse_limit, 2);
        assert_eq!(main_cell(&app, 0, 0).as_deref(), Some("one"));
        assert_eq!(main_cell(&app, 0, 1).as_deref(), Some("two"));

        // Left -> revision 1: B1 has not happened yet.
        assert_eq!(app.step_revision_back().unwrap(), RevisionStep::Moved);
        assert_eq!(app.revision_browse_limit, 1);
        assert_eq!(main_cell(&app, 0, 0).as_deref(), Some("one"));
        assert_eq!(main_cell(&app, 0, 1), None);

        // Left again at revision 1 is the reference's no-op guard.
        assert_eq!(app.step_revision_back().unwrap(), RevisionStep::AtOldest);
        assert_eq!(app.revision_browse_limit, 1);

        // Right walks back up through the same revisions.
        assert_eq!(app.step_revision_forward().unwrap(), RevisionStep::Moved);
        assert_eq!(app.revision_browse_limit, 2);
        assert_eq!(main_cell(&app, 0, 0).as_deref(), Some("one"));
        assert_eq!(main_cell(&app, 0, 1).as_deref(), Some("two"));
        assert_eq!(app.step_revision_forward().unwrap(), RevisionStep::Moved);
        assert_eq!(app.revision_browse_limit, 3);
        assert_eq!(main_cell(&app, 0, 0).as_deref(), Some("three"));

        // Right past the newest revision re-replays but does not invent a
        // revision: the limit is clamped, not incremented.
        assert_eq!(app.step_revision_forward().unwrap(), RevisionStep::AtNewest);
        assert_eq!(app.revision_browse_limit, 3);
        assert_eq!(main_cell(&app, 0, 0).as_deref(), Some("three"));
    }

    /// Browsing must not arm a commit path: a stepped revision is a
    /// read-only view of the log, so the file is untouched and no watcher
    /// is attached (a watcher would re-apply appends onto a past revision).
    #[test]
    fn revision_step_detaches_from_the_file_and_preserves_it() {
        let dir = tempfile::tempdir().unwrap();
        let (mut app, path) = browsable_core_app(&dir);
        let before = std::fs::read_to_string(&path).unwrap();

        app.step_revision_back().unwrap();
        app.step_revision_back().unwrap();

        assert!(app.path.is_none(), "browsing must keep path detached");
        assert!(app.watcher.is_none(), "browsing must not attach a watcher");
        assert_eq!(std::fs::read_to_string(&path).unwrap(), before);
    }

    /// The status names the revision actually reached, which is what the
    /// status bar and the on-canvas indicator render.
    #[test]
    fn revision_step_status_names_the_revision() {
        let dir = tempfile::tempdir().unwrap();
        let (mut app, path) = browsable_core_app(&dir);
        assert!(app.status.contains("@ revision 3"), "{}", app.status);
        app.step_revision_back().unwrap();
        assert!(app.status.contains("@ revision 2"), "{}", app.status);
        assert!(app.status.contains(&path.display().to_string()), "{}", app.status);
    }

    /// No source to re-replay is reported, not panicked on.
    #[test]
    fn revision_step_without_source_reports_no_source() {
        let mut app = test_core_app(None, WorkbookState::new(), 0);
        app.revision_browse = true;
        assert_eq!(app.step_revision_back().unwrap(), RevisionStep::NoSource);
        assert_eq!(app.step_revision_forward().unwrap(), RevisionStep::NoSource);
    }

    /// A freshly seeded workbook must carry the built-in margin TOTAL
    /// seeds.  These are the `[A_1` footer-under-margin and `]A~1`
    /// header-beside-margin cells that every new document shows.
    #[test]
    fn new_seeded_workbook_has_total_prefills() {
        let wb = WorkbookState::new_seeded();
        let sheet = wb.active_sheet();
        let footer_addr = crate::grid::CellAddr::Footer {
            row: 0,
            col: crate::grid::ColumnAddr::Left(crate::grid::MARGIN_COLS - 1),
        };
        let header_addr = crate::grid::CellAddr::Header {
            row: (crate::grid::HEADER_ROWS - 1) as u32,
            col: crate::grid::ColumnAddr::Right(0),
        };
        assert_eq!(
            sheet.grid.get(&footer_addr).as_deref(),
            Some("TOTAL"),
            "footer margin cell [A_1 must contain TOTAL"
        );
        assert_eq!(
            sheet.grid.get(&header_addr).as_deref(),
            Some("TOTAL"),
            "header margin cell ]A~1 must contain TOTAL"
        );
    }

    /// TOTAL seeds must survive serialization → reload round-trip
    /// (the ensure_unsaved_file path uses the same serialization).
    #[test]
    fn total_prefills_surround_unsaved_file_round_trip() {
        // Process-global env: serialise with the other unsaved-file tests.
        let _env_lock = super::UNSAVED_ENV_LOCK
            .lock()
            .unwrap_or_else(|e| e.into_inner());
        let dir = tempfile::tempdir().unwrap();
        let unsaved_dir = dir.path().join("unsaved");
        std::env::set_var("CORRO_UNSAVED_TEST_DIR", &unsaved_dir);

        let mut app = CoreApp {
            path: None,
            import_source: None,
            source_path: None,
            revision_limit: None,
            revision_browse: false,
            revision_browse_limit: 0,
            offset: 0,
            state: crate::ops::SheetState::new(1, 1),
            workbook: WorkbookState::new_seeded(),
            cursor: crate::grid::SheetCursor { row: 0, col: 0 },
            anchor: None,
            selection_kind: SelectionKind::Cells,
            watcher: None,
            ops_applied: 0,
            op_history: Vec::new(),
            redo_history: Vec::new(),
            view_sheet_id: 0,
            persisted_view_sort_cols: Default::default(),
            linked_source_mtimes: Default::default(),
            unsaved_file: None,
            unsaved_auto_create: true,
            status: String::new(),
            exit_message: None,
            clipboard_snapshot: None,
            edit_target_addr: None,
            edit_range_addrs: None,
            pending_lost_edit: None,
            pending_fit_to_content_on_commit: false,
            locks: Default::default(),
        };

        // ensure_unsaved_file creates the file and binds path.
        let path = app.ensure_unsaved_file().expect("ensure_unsaved_file");
        assert!(path.exists(), "unsaved file must exist on disk");

        // The serialized log must contain the TOTAL seeds.
        let log = std::fs::read_to_string(&path).unwrap();
        assert!(
            log.contains("TOTAL"),
            "unsaved log must contain TOTAL seeds; log:\n{log}"
        );

        // Reload the file into a fresh workbook and verify the seeds survived.
        let mut wb2 = WorkbookState::new();
        let mut active = wb2.sheet_id(wb2.active_sheet);
        let (_, replay) = crate::io::load_workbook_revisions_partial(
            &path, usize::MAX, &mut wb2, &mut active,
        )
        .expect("reload unsaved file");
        assert!(replay.op_count > 0, "replayed at least one op");

        let sheet2 = wb2.active_sheet();
        let footer_addr = crate::grid::CellAddr::Footer {
            row: 0,
            col: crate::grid::ColumnAddr::Left(crate::grid::MARGIN_COLS - 1),
        };
        let header_addr = crate::grid::CellAddr::Header {
            row: (crate::grid::HEADER_ROWS - 1) as u32,
            col: crate::grid::ColumnAddr::Right(0),
        };
        assert_eq!(
            sheet2.grid.get(&footer_addr).as_deref(),
            Some("TOTAL"),
            "reloaded footer margin must contain TOTAL"
        );
        assert_eq!(
            sheet2.grid.get(&header_addr).as_deref(),
            Some("TOTAL"),
            "reloaded header margin must contain TOTAL"
        );

        // Cleanup.
        std::env::remove_var("CORRO_UNSAVED_TEST_DIR");
        let _ = dir.close();
    }

    /// commit_sheet_op must persist the op to the unsaved file when
    /// unsaved_auto_create is true and path was initially None.
    #[test]
    fn commit_sheet_op_creates_unsaved_file_and_persists() {
        // Process-global env: serialise with the other unsaved-file tests.
        let _env_lock = super::UNSAVED_ENV_LOCK
            .lock()
            .unwrap_or_else(|e| e.into_inner());
        let dir = tempfile::tempdir().unwrap();
        let unsaved_dir = dir.path().join("unsaved");
        std::env::set_var("CORRO_UNSAVED_TEST_DIR", &unsaved_dir);

        let mut app = CoreApp {
            path: None,
            import_source: None,
            source_path: None,
            revision_limit: None,
            revision_browse: false,
            revision_browse_limit: 0,
            offset: 0,
            state: crate::ops::SheetState::new_seeded(),
            workbook: WorkbookState::new_seeded(),
            cursor: crate::grid::SheetCursor { row: 0, col: 0 },
            anchor: None,
            selection_kind: SelectionKind::Cells,
            watcher: None,
            ops_applied: 0,
            op_history: Vec::new(),
            redo_history: Vec::new(),
            view_sheet_id: 1,
            persisted_view_sort_cols: Default::default(),
            linked_source_mtimes: Default::default(),
            unsaved_file: None,
            unsaved_auto_create: true,
            status: String::new(),
            exit_message: None,
            clipboard_snapshot: None,
            edit_target_addr: None,
            edit_range_addrs: None,
            pending_lost_edit: None,
            pending_fit_to_content_on_commit: false,
            locks: Default::default(),
        };

        let op = Op::SetCell {
            addr: crate::grid::CellAddr::Main { row: 0, col: 0 },
            value: "42".into(),
        };
        let persisted = app.commit_sheet_op(op).expect("commit_sheet_op");
        assert!(persisted, "op must be persisted to unsaved file");
        assert!(app.path.is_some(), "path must be set after commit");
        assert!(app.ops_applied >= 1, "ops_applied must advance");

        // Verify the value is in the file.
        let log = std::fs::read_to_string(app.path.as_ref().unwrap()).unwrap();
        assert!(log.contains("42"), "committed value must be in log; log:\n{log}");

        // Cleanup.
        std::env::remove_var("CORRO_UNSAVED_TEST_DIR");
        let _ = dir.close();
    }

    /// A stale `view_sheet_id` (one that no longer names a sheet) must not
    /// silently drop the edit: `commit_sheet_op` falls back to the workbook's
    /// active sheet, so the op still lands.
    ///
    /// Regression: a fresh GUI app initialised `view_sheet_id` to 0 — not a
    /// valid sheet id (ids start at 1) — and every commit failed with
    /// "unknown sheet id", so the first cell edit was lost on a new document.
    #[test]
    fn commit_sheet_op_falls_back_when_view_sheet_id_is_stale() {
        // Process-global env: serialise with the other unsaved-file tests.
        let _env_lock = super::UNSAVED_ENV_LOCK
            .lock()
            .unwrap_or_else(|e| e.into_inner());
        let dir = tempfile::tempdir().unwrap();
        std::env::set_var("CORRO_UNSAVED_TEST_DIR", dir.path().join("unsaved"));

        let mut app = test_core_app(None, WorkbookState::new_seeded(), 0);
        app.unsaved_auto_create = true;
        // A view id that matches no sheet, exactly like a 0-seeded cache.
        let real_id = app.workbook.sheet_id(app.workbook.active_sheet);
        app.view_sheet_id = 0;
        assert!(app.workbook.sheet_index_by_id(0).is_none(), "0 is not a sheet id");

        let op = Op::SetCell {
            addr: crate::grid::CellAddr::Main { row: 0, col: 0 },
            value: "42".into(),
        };
        let persisted = app
            .commit_sheet_op(op)
            .expect("a stale view id must not fail the commit");
        assert!(persisted, "op must be persisted despite the stale view id");
        // The commit went to the real sheet, and the cache was repaired.
        assert_eq!(app.view_sheet_id, real_id, "cache must be re-synced");
        let log = std::fs::read_to_string(app.path.as_ref().unwrap()).unwrap();
        assert!(log.contains("42"), "value must reach the log; log:\n{log}");

        std::env::remove_var("CORRO_UNSAVED_TEST_DIR");
        let _ = dir.close();
    }
}
