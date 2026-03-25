//! Append-only log I/O and file watching for multi-instance sync.

use crate::ops::{apply_line, append_op, replay_lines, Op, SheetState};
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

/// Read new bytes from `path` starting at `byte_offset`, apply as JSONL lines, return new EOF offset.
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
pub fn commit_op(path: &Path, offset: &mut u64, state: &mut SheetState, op: &Op) -> Result<(), IoError> {
    append_op(path, op)?;
    *offset = tail_apply(path, *offset, state)?;
    Ok(())
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
}
