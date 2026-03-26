//! Append-only JSONL operations and replay onto [`SheetState`].

use crate::grid::{CellAddr, Grid, MainRange};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs::OpenOptions;
use std::io::Write;
use std::path::Path;

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum AggFunc {
    Sum,
    Mean,
    Median,
    Min,
    Max,
    Count,
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
pub struct AggregateDef {
    pub func: AggFunc,
    pub source: MainRange,
}

#[derive(Clone, Debug, Default)]
pub struct SheetState {
    pub grid: Grid,
    pub aggregates: HashMap<CellAddr, AggregateDef>,
}

impl SheetState {
    pub fn new(main_rows: usize, main_cols: usize) -> Self {
        SheetState {
            grid: Grid::new(main_rows as u32, main_cols as u32),
            aggregates: HashMap::new(),
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(tag = "op", rename_all = "snake_case")]
pub enum Op {
    SetCell {
        addr: CellAddr,
        value: String,
    },
    SetAggregate {
        addr: CellAddr,
        def: AggregateDef,
    },
    SetMainSize {
        main_rows: u32,
        main_cols: u32,
    },
    MoveRowRange {
        from: u32,
        count: u32,
        to: u32,
    },
    MoveColRange {
        from: u32,
        count: u32,
        to: u32,
    },
}

impl Op {
    pub fn apply(&self, state: &mut SheetState) {
        match self {
            Op::SetCell { addr, value } => {
                state.grid.set(addr, value.clone());
            }
            Op::SetAggregate { addr, def } => {
                state.aggregates.insert(addr.clone(), def.clone());
            }
            Op::SetMainSize {
                main_rows,
                main_cols,
            } => {
                state
                    .grid
                    .set_main_size(*main_rows as usize, *main_cols as usize);
            }
            Op::MoveRowRange { from, count, to } => {
                state.grid.move_main_rows(
                    *from as usize,
                    *count as usize,
                    *to as usize,
                );
            }
            Op::MoveColRange { from, count, to } => {
                state.grid.move_main_cols(
                    *from as usize,
                    *count as usize,
                    *to as usize,
                );
            }
        }
    }
}

/// Replay JSONL from a string (full load).
pub fn replay_lines(data: &str, state: &mut SheetState) -> Result<usize, serde_json::Error> {
    let mut n = 0usize;
    for line in data.lines() {
        let t = line.trim();
        if t.is_empty() {
            continue;
        }
        let op: Op = serde_json::from_str(t)?;
        op.apply(state);
        n += 1;
    }
    Ok(n)
}

/// Parse a single line and apply; used when tailing.
pub fn apply_line(line: &str, state: &mut SheetState) -> Result<(), serde_json::Error> {
    let t = line.trim();
    if t.is_empty() {
        return Ok(());
    }
    let op: Op = serde_json::from_str(t)?;
    op.apply(state);
    Ok(())
}

/// Append one op as JSONL to `path` (creates file if missing).
pub fn append_op(path: &Path, op: &Op) -> std::io::Result<()> {
    let mut f = OpenOptions::new()
        .create(true)
        .append(true)
        .open(path)?;
    let line = serde_json::to_string(op)?;
    writeln!(f, "{line}")?;
    f.sync_all()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::grid::CellAddr;

    #[test]
    fn replay_set_cell() {
        let mut s = SheetState::new(2, 2);
        let json = r#"{"op":"set_cell","addr":{"kind":"main","row":0,"col":0},"value":"x"}"#;
        apply_line(json, &mut s).unwrap();
        assert_eq!(s.grid.get(&CellAddr::Main { row: 0, col: 0 }), Some("x"));
    }

    #[test]
    fn op_json_roundtrip() {
        let op = Op::MoveRowRange {
            from: 0,
            count: 2,
            to: 3,
        };
        let j = serde_json::to_string(&op).unwrap();
        let back: Op = serde_json::from_str(&j).unwrap();
        assert_eq!(op, back);
    }
}
