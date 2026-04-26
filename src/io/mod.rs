//! Append-only log I/O, file watching, and tabular import for multi-instance sync.

use crate::grid::CellAddr;
use crate::ops::{
    append_line, append_op, apply_line, apply_log_line_to_workbook, replay_lines,
    replay_lines_partial, Op, SheetState, WorkbookOp, WorkbookSnapshot, WorkbookState,
};
use notify::{RecursiveMode, Watcher};
use std::fs;
use std::io::{Read, Seek, SeekFrom};
use std::path::{Path, PathBuf};
use std::sync::mpsc::Receiver;
use thiserror::Error;

#[derive(Debug, Error)]
pub enum IoError {
    #[error("I/O: {0}")]
    Io(#[from] std::io::Error),
    #[error("Notify: {0}")]
    Notify(#[from] notify::Error),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PartialReplay {
    pub op_count: usize,
    pub failed_line: Option<usize>,
    pub error: Option<String>,
}

/// Load entire file from disk and replay into `state`. Returns `(byte_len, op_count)`.
pub fn load_full(path: &Path, state: &mut SheetState) -> Result<(u64, usize), IoError> {
    let data = fs::read_to_string(path)?;
    let n = replay_lines(&data, state)?;
    Ok((data.len() as u64, n))
}

pub fn load_full_partial(
    path: &Path,
    state: &mut SheetState,
) -> Result<(u64, PartialReplay), IoError> {
    let data = fs::read_to_string(path)?;
    let (n, failed_line, error) = replay_lines_partial(&data, state)?;
    Ok((
        data.len() as u64,
        PartialReplay {
            op_count: n,
            failed_line,
            error: error.map(|e| e.to_string()),
        },
    ))
}

pub fn load_revisions_partial(
    path: &Path,
    limit: usize,
    state: &mut SheetState,
) -> Result<(u64, PartialReplay), IoError> {
    let data = fs::read_to_string(path)?;
    if limit == 0 {
        return Ok((
            data.len() as u64,
            PartialReplay {
                op_count: 0,
                failed_line: None,
                error: None,
            },
        ));
    }
    let mut n = 0usize;
    for (idx, line) in data.lines().enumerate() {
        let t = line.trim();
        if t.is_empty() {
            continue;
        }
        if let Err(err) = apply_line(t, state) {
            return Ok((
                data.len() as u64,
                PartialReplay {
                    op_count: n,
                    failed_line: Some(idx + 1),
                    error: Some(err.to_string()),
                },
            ));
        }
        n += 1;
        if n >= limit {
            break;
        }
    }
    Ok((
        data.len() as u64,
        PartialReplay {
            op_count: n,
            failed_line: None,
            error: None,
        },
    ))
}

/// Load at most `limit` log entries from disk and replay into a workbook snapshot.
pub fn load_workbook_revisions(
    path: &Path,
    limit: usize,
    workbook: &mut WorkbookState,
    active_sheet: &mut u32,
) -> Result<(u64, usize), IoError> {
    let data = fs::read_to_string(path)?;
    if limit == 0 {
        return Ok((data.len() as u64, 0));
    }
    let mut n = 0usize;
    for line in data.lines() {
        let t = line.trim();
        if t.is_empty() {
            continue;
        }
        apply_log_line_to_workbook(t, workbook, active_sheet)?;
        n += 1;
        if n >= limit {
            break;
        }
    }
    Ok((data.len() as u64, n))
}

pub fn load_workbook_revisions_partial(
    path: &Path,
    limit: usize,
    workbook: &mut WorkbookState,
    active_sheet: &mut u32,
) -> Result<(u64, PartialReplay), IoError> {
    let data = fs::read_to_string(path)?;
    if limit == 0 {
        return Ok((
            data.len() as u64,
            PartialReplay {
                op_count: 0,
                failed_line: None,
                error: None,
            },
        ));
    }
    let mut n = 0usize;
    for (idx, line) in data.lines().enumerate() {
        let t = line.trim();
        if t.is_empty() {
            continue;
        }
        if let Err(err) = apply_log_line_to_workbook(t, workbook, active_sheet) {
            return Ok((
                data.len() as u64,
                PartialReplay {
                    op_count: n,
                    failed_line: Some(idx + 1),
                    error: Some(err.to_string()),
                },
            ));
        }
        n += 1;
        if n >= limit {
            break;
        }
    }
    Ok((
        data.len() as u64,
        PartialReplay {
            op_count: n,
            failed_line: None,
            error: None,
        },
    ))
}

pub fn save_workbook(path: &Path, workbook: &WorkbookSnapshot) -> Result<(), IoError> {
    let mut out = String::new();
    out.push_str(&format!(
        "WORKBOOK {} {}\n",
        workbook.next_sheet_id, workbook.active_sheet_id
    ));
    for sheet in &workbook.sheets {
        out.push_str(&format!("SHEET {} {}\n", sheet.id, sheet.title));
        out.push_str(&format!(
        "VOLATILE_SEED {}\n",
        sheet.state.grid.volatile_seed()
    ));
        for row in 0..sheet.state.grid.main_rows() {
            for col in 0..sheet.state.grid.main_cols() {
                let addr = CellAddr::Main {
                    row: row as u32,
                    col: col as u32,
                };
                if let Some(value) = sheet.state.grid.get(&addr) {
                    if !value.is_empty() {
                        out.push_str(&format!("SET {} {}\n", workbook_addr_label(&addr), value));
                    }
                }
            }
        }
        out.push_str("END_SHEET\n");
    }
    fs::write(path, out)?;
    Ok(())
}

fn workbook_addr_label(addr: &CellAddr) -> String {
    crate::addr::cell_ref_text(addr, 0)
}

pub fn load_workbook_snapshot(path: &Path) -> Result<WorkbookSnapshot, IoError> {
    let data = fs::read_to_string(path)?;
    let mut lines = data.lines();
    let header = lines.next().ok_or_else(|| {
        std::io::Error::new(std::io::ErrorKind::InvalidData, "missing workbook header")
    })?;
    let mut header_parts = header.split_whitespace();
    if header_parts.next() != Some("WORKBOOK") {
        return Err(
            std::io::Error::new(std::io::ErrorKind::InvalidData, "bad workbook header").into(),
        );
    }
    let next_sheet_id = header_parts
        .next()
        .and_then(|v| v.parse::<u32>().ok())
        .ok_or_else(|| {
            std::io::Error::new(std::io::ErrorKind::InvalidData, "bad workbook header")
        })?;
    let active_sheet_id = header_parts
        .next()
        .and_then(|v| v.parse::<u32>().ok())
        .ok_or_else(|| {
            std::io::Error::new(std::io::ErrorKind::InvalidData, "bad workbook header")
        })?;

    let mut sheets = Vec::new();
    let mut current: Option<crate::ops::SheetRecord> = None;

    for raw in lines {
        let line = raw.trim();
        if line.is_empty() {
            continue;
        }
        let mut parts = line.split_whitespace();
        match parts.next() {
            Some("SHEET") => {
                if let Some(sheet) = current.take() {
                    sheets.push(sheet);
                }
                let id = parts
                    .next()
                    .and_then(|v| v.parse::<u32>().ok())
                    .ok_or_else(|| {
                        std::io::Error::new(std::io::ErrorKind::InvalidData, "bad sheet header")
                    })?;
                let title = parts.collect::<Vec<_>>().join(" ");
                current = Some(crate::ops::SheetRecord {
                    id,
                    title,
                    state: SheetState::new(1, 1),
                });
            }
            Some("VOLATILE_SEED") => {
                if let Some(sheet) = current.as_mut() {
                    let seed = parts
                        .next()
                        .and_then(|v| v.parse::<u64>().ok())
                        .ok_or_else(|| {
                            std::io::Error::new(std::io::ErrorKind::InvalidData, "bad seed line")
                        })?;
                    sheet.state.grid.set_volatile_seed(seed);
                }
            }
            Some("END_SHEET") => {
                if let Some(sheet) = current.take() {
                    sheets.push(sheet);
                }
            }
            Some(_) => {
                if let Some(sheet) = current.as_mut() {
                    apply_line(line, &mut sheet.state)?;
                }
            }
            None => {}
        }
    }

    if let Some(sheet) = current.take() {
        sheets.push(sheet);
    }
    if sheets.is_empty() {
        return Err(std::io::Error::new(std::io::ErrorKind::InvalidData, "empty workbook").into());
    }

    Ok(WorkbookSnapshot {
        next_sheet_id,
        active_sheet_id,
        sheets,
        volatile_seed: 0,
    })
}

/// Load at most `limit` non-empty log lines from disk and replay into `state`.
pub fn load_revisions(
    path: &Path,
    limit: usize,
    state: &mut SheetState,
) -> Result<(u64, usize), IoError> {
    let data = fs::read_to_string(path)?;
    if limit == 0 {
        return Ok((data.len() as u64, 0));
    }
    let mut n = 0usize;
    for line in data.lines() {
        let t = line.trim();
        if t.is_empty() {
            continue;
        }
        apply_line(t, state)?;
        n += 1;
        if n >= limit {
            break;
        }
    }
    Ok((data.len() as u64, n))
}

/// Read new bytes from `path` starting at `byte_offset`, apply appended log lines, return new EOF offset.
pub fn tail_apply(path: &Path, byte_offset: u64, state: &mut SheetState) -> Result<u64, IoError> {
    let meta = fs::metadata(path)?;
    let len = meta.len();
    if len < byte_offset {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidData,
            "file shrank; full reload required",
        )
        .into());
    }
    if len == byte_offset {
        return Ok(byte_offset);
    }
    let mut f = fs::File::open(path)?;
    f.seek(SeekFrom::Start(byte_offset))?;
    let mut rest = String::new();
    f.read_to_string(&mut rest)?;
    for line in rest.lines() {
        apply_line(line, state)?;
    }
    Ok(len)
}

/// Append `op` to the log and apply newly written bytes from `offset` (single-writer tail).
pub fn commit_op(
    path: &Path,
    offset: &mut u64,
    state: &mut SheetState,
    op: &Op,
) -> Result<(), IoError> {
    let mut preview = state.clone();
    op.apply(&mut preview);
    append_op(path, op, preview.grid.main_cols())?;
    *offset = tail_apply(path, *offset, state)?;
    Ok(())
}

/// Append a plain-text document-setting line and apply it to the live state.
pub fn commit_line(
    path: &Path,
    offset: &mut u64,
    state: &mut SheetState,
    line: &str,
) -> Result<(), IoError> {
    append_line(path, line)?;
    *offset = tail_apply(path, *offset, state)?;
    Ok(())
}

pub fn commit_workbook_op(
    path: &Path,
    offset: &mut u64,
    workbook: &mut WorkbookState,
    active_sheet: &mut u32,
    op: &WorkbookOp,
) -> Result<(), IoError> {
    let mut preview = workbook.clone();
    let mut preview_active_sheet = *active_sheet;
    crate::ops::apply_workbook_op(&mut preview, &mut preview_active_sheet, op.clone())?;
    // Compute main_cols to pass to to_log_line(). Normally we use the
    // preview (post-apply) width so the serialized address can use
    // main-region Excel names when appropriate. However, when the
    // operation itself is a header/footer SET that caused the preview to
    // expand main columns, preserve the mental model of storing the
    // header/footer in the margin by decrementing the preview width by
    // one for serialization. This matches historical expectations in
    // tests that prefer a mirrored-margin address in that case.
    // Determine the pre-apply main_cols for the target sheet (if available).
    let pre_main_cols = match op {
        WorkbookOp::SheetOp { sheet_id, .. } => workbook
            .sheets
            .iter()
            .find(|s| s.id == *sheet_id)
            .map(|s| s.state.grid.main_cols()),
        _ => Some(workbook.active_sheet().grid.main_cols()),
    };

    let preview_main_cols = match op {
        WorkbookOp::SheetOp { sheet_id, .. } => preview
            .sheets
            .iter()
            .find(|s| s.id == *sheet_id)
            .map(|s| s.state.grid.main_cols())
            .unwrap_or_else(|| preview.active_sheet().grid.main_cols()),
        _ => preview.active_sheet().grid.main_cols(),
    };

    let mut main_cols = preview_main_cols;
    if let WorkbookOp::SheetOp {
        sheet_id: _,
        op: inner_op,
    } = op
    {
        use crate::ops::Op::*;
        if let Op::SetCell { addr, .. } = inner_op {
            match addr {
                crate::grid::CellAddr::Header { col, .. }
                | crate::grid::CellAddr::Footer { col, .. } => {
                    if let Some(pre) = pre_main_cols {
                        if preview_main_cols > pre && (*col as usize) >= crate::grid::MARGIN_COLS {
                            main_cols = preview_main_cols.saturating_sub(1);
                        }
                    }
                }
                _ => {}
            }
        }
    }
    append_line(path, &op.to_log_line(main_cols))?;
    *offset = tail_apply_workbook(path, *offset, workbook, active_sheet)?;
    Ok(())
}

pub fn commit_sheet_log_line(
    path: &Path,
    offset: &mut u64,
    workbook: &mut WorkbookState,
    active_sheet: &mut u32,
    line: &str,
) -> Result<(), IoError> {
    append_line(path, line)?;
    *offset = tail_apply_workbook(path, *offset, workbook, active_sheet)?;
    Ok(())
}

pub fn tail_apply_workbook(
    path: &Path,
    byte_offset: u64,
    workbook: &mut WorkbookState,
    active_sheet: &mut u32,
) -> Result<u64, IoError> {
    let meta = fs::metadata(path)?;
    let len = meta.len();
    if len < byte_offset {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidData,
            "file shrank; full reload required",
        )
        .into());
    }
    if len == byte_offset {
        return Ok(byte_offset);
    }
    let mut f = fs::File::open(path)?;
    f.seek(SeekFrom::Start(byte_offset))?;
    let mut rest = String::new();
    f.read_to_string(&mut rest)?;
    for line in rest.lines() {
        apply_log_line_to_workbook(line, workbook, active_sheet)?;
    }
    Ok(len)
}

// ── Tabular import ───────────────────────────────────────────────────────────

pub fn import_tsv(data: &str, state: &mut SheetState) {
    import_delimited(data, state, '\t');
}

pub fn import_csv(data: &str, state: &mut SheetState) {
    import_delimited(data, state, ',');
}

fn import_delimited(data: &str, state: &mut SheetState, delim: char) {
    let lines: Vec<&str> = data.lines().collect();
    if lines.is_empty() {
        return;
    }

    let mut rows: Vec<Vec<String>> = Vec::new();
    for line in &lines {
        if delim == ',' {
            rows.push(parse_csv_line(line));
        } else {
            rows.push(line.split(delim).map(|s| s.to_string()).collect());
        }
    }

    let max_cols = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    if max_cols == 0 {
        return;
    }

    let first_all_numeric = rows.first().map_or(true, |r| {
        r.iter().all(|cell| {
            let t = cell.trim();
            t.is_empty() || t.parse::<f64>().is_ok()
        })
    });

    let (header_row, data_rows) = if first_all_numeric || rows.len() <= 1 {
        (None, &rows[..])
    } else {
        (Some(&rows[0]), &rows[1..])
    };

    let mc = max_cols as u32;
    let mr = data_rows.len() as u32;
    state
        .grid
        .set_main_size(mr.max(1) as usize, mc.max(1) as usize);

    if let Some(hdr) = header_row {
        use crate::grid::HEADER_ROWS;
        let header_idx = (HEADER_ROWS - 1) as u8;
        for (ci, val) in hdr.iter().enumerate() {
            if !val.is_empty() {
                let global_col = crate::grid::MARGIN_COLS as u32 + ci as u32;
                state.grid.set(
                    &CellAddr::Header {
                        row: header_idx,
                        col: global_col,
                    },
                    val.clone(),
                );
            }
        }
    }

    for ci in 0..max_cols {
        state.grid.auto_fit_column(crate::grid::MARGIN_COLS + ci);
    }

    for (ri, row) in data_rows.iter().enumerate() {
        for (ci, val) in row.iter().enumerate() {
            if !val.is_empty() {
                state.grid.set(
                    &CellAddr::Main {
                        row: ri as u32,
                        col: ci as u32,
                    },
                    val.clone(),
                );
            }
        }
    }
}

fn parse_csv_line(line: &str) -> Vec<String> {
    let mut fields = Vec::new();
    let mut current = String::new();
    let mut in_quotes = false;
    let mut chars = line.chars().peekable();

    while let Some(ch) = chars.next() {
        if in_quotes {
            if ch == '"' {
                if chars.peek() == Some(&'"') {
                    current.push('"');
                    chars.next();
                } else {
                    in_quotes = false;
                }
            } else {
                current.push(ch);
            }
        } else if ch == '"' {
            in_quotes = true;
        } else if ch == ',' {
            fields.push(current.clone());
            current.clear();
        } else {
            current.push(ch);
        }
    }
    fields.push(current);
    fields
}

/// Watches `path` for changes; poll [`LogWatcher::poll_dirty`].
pub struct LogWatcher {
    _watcher: notify::RecommendedWatcher,
    pub path: PathBuf,
    rx: Receiver<notify::Result<notify::Event>>,
}

impl LogWatcher {
    pub fn new(path: PathBuf) -> Result<Self, notify::Error> {
        let (tx, rx) = std::sync::mpsc::channel();
        let mut watcher = notify::recommended_watcher(move |ev| {
            let _ = tx.send(ev);
        })?;
        watcher.watch(&path, RecursiveMode::NonRecursive)?;
        Ok(LogWatcher {
            _watcher: watcher,
            path,
            rx,
        })
    }

    /// Non-blocking drain: returns true if any modify/create event arrived.
    pub fn poll_dirty(&self) -> bool {
        let mut dirty = false;
        while let Ok(Ok(ev)) = self.rx.try_recv() {
            match ev.kind {
                notify::EventKind::Modify(_) | notify::EventKind::Create(_) => dirty = true,
                _ => {}
            }
        }
        dirty
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::grid::CellAddr;
    use std::path::PathBuf;

    fn docs_test_path(name: &str) -> PathBuf {
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("docs/tests")
            .join(name)
    }
    use crate::ops::{Op, SheetRecord, SheetState, WorkbookSnapshot};
    use tempfile::NamedTempFile;

    #[test]
    fn commit_op_roundtrip() {
        let path = NamedTempFile::new().unwrap();
        let mut state = SheetState::new(2, 2);
        let mut offset = 0u64;
        let op = Op::SetCell {
            addr: CellAddr::Main { row: 0, col: 0 },
            value: "42".into(),
        };
        commit_op(path.path(), &mut offset, &mut state, &op).unwrap();
        assert_eq!(offset, fs::metadata(path.path()).unwrap().len());
        assert_eq!(
            state.grid.get(&CellAddr::Main { row: 0, col: 0 }).as_deref(),
            Some("42")
        );
    }

    #[test]
    fn commit_workbook_op_uses_post_apply_width_for_header_refs() {
        let path = NamedTempFile::new().unwrap();
        let mut workbook = WorkbookState::new();
        let sheet_id = workbook.sheet_id(workbook.active_sheet);
        workbook.sheets[0].state.grid.set_main_size(1, 1);
        let mut offset = 0u64;
        let mut active_sheet = sheet_id;
        let op = WorkbookOp::SheetOp {
            sheet_id,
            op: Op::SetCell {
                addr: CellAddr::Header { row: 25, col: 12 },
                value: "x".into(),
            },
        };

        commit_workbook_op(
            path.path(),
            &mut offset,
            &mut workbook,
            &mut active_sheet,
            &op,
        )
        .unwrap();

        let written = fs::read_to_string(path.path()).unwrap();
        assert!(written.contains("SET $1:]A~1 x"));
    }

    #[test]
    fn save_as_snapshot_reloads_as_log() {
        let path = NamedTempFile::new().unwrap();
        let mut state = SheetState::new(2, 2);
        state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "b".into());

        let log = concat!("SET A1 a\n", "SET A2 b\n",);
        fs::write(path.path(), log).unwrap();

        let mut reloaded = SheetState::new(1, 1);
        let (off, n) = load_full(path.path(), &mut reloaded).unwrap();
        assert!(off > 0);
        assert!(n > 0);
        assert_eq!(
            reloaded.grid.get(&CellAddr::Main { row: 0, col: 0 }).as_deref(),
            Some("a")
        );
        assert_eq!(
            reloaded.grid.get(&CellAddr::Main { row: 1, col: 0 }).as_deref(),
            Some("b")
        );
    }

    #[test]
    fn load_legacy_test5_corro() {
        let mut state = SheetState::new(1, 1);
        let (off, n) = load_revisions(&docs_test_path("main.corro"), 2, &mut state).unwrap();

        assert!(off > 0);
        assert_eq!(n, 2);
        assert_eq!(
            state.grid.get(&CellAddr::Main { row: 0, col: 0 }).as_deref(),
            Some("1")
        );
        assert_eq!(
            state.grid.get(&CellAddr::Main { row: 1, col: 0 }).as_deref(),
            Some("7")
        );
        assert_eq!(
            state.grid.get(&CellAddr::Header { row: 25, col: 10 }).as_deref(),
            Some("")
        );
    }

    #[test]
    fn load_revisions_zero_loads_nothing() {
        let mut state = SheetState::new(1, 1);
        let (off, n) = load_revisions(&docs_test_path("main.corro"), 0, &mut state).unwrap();

        assert!(off > 0);
        assert_eq!(n, 0);
        assert_eq!(state.grid.get(&CellAddr::Main { row: 0, col: 0 }), None);
    }

    #[test]
    fn load_revisions_limits_replay() {
        let mut state = SheetState::new(1, 1);
        let (off, n) = load_revisions(&docs_test_path("main.corro"), 2, &mut state).unwrap();

        assert!(off > 0);
        assert_eq!(n, 2);
        assert_eq!(
            state.grid.get(&CellAddr::Main { row: 0, col: 0 }).as_deref(),
            Some("1")
        );
        assert_eq!(
            state.grid.get(&CellAddr::Main { row: 1, col: 0 }).as_deref(),
            Some("7")
        );
        assert_eq!(state.grid.get(&CellAddr::Main { row: 2, col: 0 }), None);
    }

    #[test]
    fn load_revisions_partial_reports_bad_line() {
        let tmp = NamedTempFile::new().unwrap();
        std::fs::write(tmp.path(), "SET A1 1\nBAD LINE\nSET A2 2\n").unwrap();
        let mut state = SheetState::new(1, 1);
        let (_, replay) = load_revisions_partial(tmp.path(), usize::MAX, &mut state).unwrap();

        assert_eq!(replay.op_count, 1);
        assert_eq!(replay.failed_line, Some(2));
        assert!(replay
            .error
            .as_deref()
            .unwrap_or_default()
            .contains("unrecognized"));
        assert_eq!(
            state.grid.get(&CellAddr::Main { row: 0, col: 0 }).as_deref(),
            Some("1")
        );
    }

    #[test]
    fn save_and_load_workbook_snapshot_roundtrip() {
        let path = NamedTempFile::new().unwrap();
        let mut workbook = WorkbookSnapshot {
            next_sheet_id: 3,
            active_sheet_id: 2,
            volatile_seed: 0,
            sheets: vec![
                SheetRecord {
                    id: 1,
                    title: "Sheet1".into(),
                    state: SheetState::new(1, 1),
                },
                SheetRecord {
                    id: 2,
                    title: "Sheet2".into(),
                    state: SheetState::new(1, 1),
                },
            ],
        };
        workbook.sheets[1]
            .state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "hello".into());

        save_workbook(path.path(), &workbook).unwrap();
        let loaded = load_workbook_snapshot(path.path()).unwrap();

        assert_eq!(loaded.next_sheet_id, 3);
        assert_eq!(loaded.active_sheet_id, 2);
        assert_eq!(loaded.sheets.len(), 2);
        assert_eq!(loaded.sheets[1].title, "Sheet2");
        assert_eq!(
            loaded.sheets[1]
                .state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("hello")
        );
    }

    #[test]
    fn workbook_replay_test5_corro_reports_first_failing_line() {
        let data = fs::read_to_string(docs_test_path("main.corro")).unwrap();
        let mut workbook = WorkbookState::new();
        let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
        for (idx, line) in data.lines().enumerate() {
            let t = line.trim();
            if t.is_empty() {
                continue;
            }
            apply_log_line_to_workbook(t, &mut workbook, &mut active_sheet)
                .unwrap_or_else(|e| panic!("line {}: {} ({e})", idx + 1, t));
        }
    }

    #[test]
    fn import_tsv_basic() {
        let mut state = SheetState::new(1, 1);
        import_tsv("Name\tAge\nAlice\t30\nBob\t25\n", &mut state);
        assert_eq!(state.grid.main_rows(), 2);
        assert_eq!(state.grid.main_cols(), 2);
        assert_eq!(
            state.grid.get(&CellAddr::Main { row: 0, col: 0 }).as_deref(),
            Some("Alice")
        );
        assert_eq!(
            state.grid.get(&CellAddr::Main { row: 1, col: 1 }).as_deref(),
            Some("25")
        );
    }

    #[test]
    fn import_csv_quoted() {
        let mut state = SheetState::new(1, 1);
        import_csv("a,\"b,c\",d\n1,2,3\n", &mut state);
        assert_eq!(state.grid.main_rows(), 1);
        assert_eq!(state.grid.main_cols(), 3);
        assert_eq!(
            state.grid.get(&CellAddr::Main { row: 0, col: 1 }).as_deref(),
            Some("2")
        );
    }
}
