//! Append-only log I/O, file watching, and tabular import for multi-instance sync.

use crate::grid::CellAddr;
use crate::ops::{append_line, append_op, apply_line, replay_lines, Op, SheetState};
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
    #[error("JSON: {0}")]
    Json(#[from] serde_json::Error),
    #[error("Notify: {0}")]
    Notify(#[from] notify::Error),
}

/// Load entire file from disk and replay into `state`. Returns `(byte_len, op_count)`.
pub fn load_full(path: &Path, state: &mut SheetState) -> Result<(u64, usize), IoError> {
    let data = fs::read_to_string(path)?;
    let n = replay_lines(&data, state)?;
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
    append_op(path, op)?;
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
    use crate::ops::{Op, SheetState};
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
            state.grid.get(&CellAddr::Main { row: 0, col: 0 }),
            Some("42")
        );
    }

    #[test]
    fn import_tsv_basic() {
        let mut state = SheetState::new(1, 1);
        import_tsv("Name\tAge\nAlice\t30\nBob\t25\n", &mut state);
        assert_eq!(state.grid.main_rows(), 2);
        assert_eq!(state.grid.main_cols(), 2);
        assert_eq!(
            state.grid.get(&CellAddr::Main { row: 0, col: 0 }),
            Some("Alice")
        );
        assert_eq!(
            state.grid.get(&CellAddr::Main { row: 1, col: 1 }),
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
            state.grid.get(&CellAddr::Main { row: 0, col: 1 }),
            Some("2")
        );
    }
}
