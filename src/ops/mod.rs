//! Append-only log operations and replay onto [`SheetState`].

use crate::addr::{
    parse_cell_ref_at, parse_excel_column, parse_sheet_id_prefix_at,
    parse_sheet_qualified_cell_ref_at,
};
use crate::grid::{CellAddr, Grid, MainRange, SortSpec, MARGIN_COLS};
use std::fs::OpenOptions;
use std::io::Write;
use std::path::Path;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum AggFunc {
    Sum,
    Mean,
    Median,
    Min,
    Max,
    Count,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct AggregateDef {
    pub func: AggFunc,
    pub source: MainRange,
}

#[derive(Clone, Debug, Default)]
pub struct SheetState {
    pub grid: Grid,
}

impl SheetState {
    pub fn new(main_rows: usize, main_cols: usize) -> Self {
        SheetState {
            grid: Grid::new(main_rows as u32, main_cols as u32),
        }
    }
}

#[derive(Clone, Debug, Default)]
pub struct WorkbookState {
    pub sheets: Vec<SheetRecord>,
    pub active_sheet: usize,
    pub next_sheet_id: u32,
}

#[derive(Clone, Debug)]
pub struct SheetRecord {
    pub id: u32,
    pub title: String,
    pub state: SheetState,
}

impl WorkbookState {
    pub fn new() -> Self {
        Self {
            sheets: vec![SheetRecord {
                id: 1,
                title: "Sheet1".into(),
                state: SheetState::new(1, 1),
            }],
            active_sheet: 0,
            next_sheet_id: 2,
        }
    }

    pub fn active_sheet(&self) -> &SheetState {
        &self.sheets[self.active_sheet].state
    }

    pub fn active_sheet_mut(&mut self) -> &mut SheetState {
        &mut self.sheets[self.active_sheet].state
    }

    pub fn ensure_active_sheet(&mut self) {
        if self.sheets.is_empty() {
            self.sheets.push(SheetRecord {
                id: 1,
                title: "Sheet1".into(),
                state: SheetState::new(1, 1),
            });
            self.active_sheet = 0;
            self.next_sheet_id = 2;
        } else if self.active_sheet >= self.sheets.len() {
            self.active_sheet = 0;
        }
    }

    pub fn sheet_count(&self) -> usize {
        self.sheets.len()
    }

    pub fn sheet_title(&self, index: usize) -> &str {
        self.sheets
            .get(index)
            .map(|s| s.title.as_str())
            .unwrap_or("")
    }

    pub fn sheet_id(&self, index: usize) -> u32 {
        self.sheets.get(index).map(|s| s.id).unwrap_or(0)
    }

    pub fn add_sheet(&mut self, title: String, state: SheetState) -> usize {
        let id = self.next_sheet_id;
        self.next_sheet_id += 1;
        self.sheets.push(SheetRecord { id, title, state });
        self.sheets.len() - 1
    }

    pub fn add_sheet_record(&mut self, record: SheetRecord) -> usize {
        self.next_sheet_id = self.next_sheet_id.max(record.id.saturating_add(1));
        self.sheets.push(record);
        self.sheets.len() - 1
    }

    pub fn sheet_index_by_id(&self, id: u32) -> Option<usize> {
        self.sheets.iter().position(|s| s.id == id)
    }

    pub fn sheet_mut_by_index(&mut self, index: usize) -> Option<&mut SheetState> {
        self.sheets.get_mut(index).map(|sheet| &mut sheet.state)
    }

    pub fn sheet_mut_by_id(&mut self, id: u32) -> Option<&mut SheetState> {
        let index = self.sheet_index_by_id(id)?;
        self.sheets.get_mut(index).map(|sheet| &mut sheet.state)
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum Op {
    SetCell { addr: CellAddr, value: String },
    SetMainSize { main_rows: u32, main_cols: u32 },
    MoveRowRange { from: u32, count: u32, to: u32 },
    MoveColRange { from: u32, count: u32, to: u32 },
    SetMaxColWidth { width: usize },
    SetColWidth { col: usize, width: Option<usize> },
    SetViewSortCols { cols: Vec<SortSpec> },
    Undo { target: String },
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum WorkbookOp {
    NewSheet { id: u32, title: String },
    ActivateSheet { id: u32 },
    RenameSheet { id: u32, title: String },
    SheetOp { sheet_id: u32, op: Op },
}

fn sheet_prefix(sheet_id: u32) -> String {
    format!("${sheet_id}:")
}

#[derive(Clone, Debug)]
pub struct WorkbookSnapshot {
    pub next_sheet_id: u32,
    pub active_sheet_id: u32,
    pub sheets: Vec<SheetRecord>,
    pub volatile_seed: u64,
}

impl WorkbookSnapshot {
    pub fn from_workbook(workbook: &WorkbookState) -> Self {
        Self {
            next_sheet_id: workbook.next_sheet_id,
            active_sheet_id: workbook.sheet_id(workbook.active_sheet),
            sheets: workbook.sheets.clone(),
            volatile_seed: workbook.active_sheet().grid.volatile_seed,
        }
    }
}

impl Op {
    pub fn apply(&self, state: &mut SheetState) {
        match self {
            Op::SetCell { addr, value } => {
                state.grid.set(addr, value.clone());
                state.grid.bump_volatile_seed();
            }
            Op::SetMainSize {
                main_rows,
                main_cols,
            } => {
                state
                    .grid
                    .set_main_size(*main_rows as usize, *main_cols as usize);
                state.grid.bump_volatile_seed();
            }
            Op::MoveRowRange { from, count, to } => {
                state
                    .grid
                    .move_main_rows(*from as usize, *count as usize, *to as usize);
                state.grid.bump_volatile_seed();
            }
            Op::MoveColRange { from, count, to } => {
                state
                    .grid
                    .move_main_cols(*from as usize, *count as usize, *to as usize);
                state.grid.bump_volatile_seed();
            }
            Op::SetMaxColWidth { width } => {
                state.grid.set_max_col_width(*width);
            }
            Op::SetColWidth { col, width } => {
                state.grid.set_col_width(*col, *width);
            }
            Op::SetViewSortCols { cols } => {
                state.grid.set_view_sort_cols(cols.clone());
            }
            Op::Undo { .. } => {}
        }
    }
}

fn addr_text(addr: &CellAddr) -> String {
    match addr {
        CellAddr::Header { row, col } => format!(
            "~{}{}",
            crate::grid::HEADER_ROWS - *row as usize,
            crate::addr::excel_column_name(*col as usize)
        ),
        CellAddr::Footer { row, col } => format!(
            "_{}{}",
            *row as usize + 1,
            crate::addr::excel_column_name(*col as usize)
        ),
        CellAddr::Main { row, col } => format!(
            "{}{}",
            crate::addr::excel_column_name(*col as usize),
            row + 1
        ),
        CellAddr::Left { col, row } => format!(
            "[{}{}",
            crate::addr::mirror_margin_column_name(*col as usize, true),
            row + 1
        ),
        CellAddr::Right { col, row } => format!(
            "]{}{}",
            crate::addr::mirror_margin_column_name(*col as usize, false),
            row + 1
        ),
    }
}

fn parse_op_text(line: &str) -> Option<Op> {
    let mut parts = line.split_whitespace();
    let cmd = parts.next()?.to_ascii_uppercase();
    match cmd.as_str() {
        "SET" => {
            let addr = parts.next()?;
            let value = parts.collect::<Vec<_>>().join(" ");
            let (addr, _) = crate::addr::parse_sheet_qualified_cell_ref_at(addr)
                .map(|(_, addr, len)| (addr, len))
                .or_else(|| crate::addr::parse_cell_ref_at(addr))?;
            Some(Op::SetCell { addr, value })
        }
        "MOVE" => {
            let kind = parts.next()?.to_ascii_uppercase();
            let from = parts.next()?.parse::<u32>().ok()?;
            let count = parts.next()?.parse::<u32>().ok()?;
            let to = parts.next()?.parse::<u32>().ok()?;
            match kind.as_str() {
                "ROW" => Some(Op::MoveRowRange { from, count, to }),
                "COL" => Some(Op::MoveColRange { from, count, to }),
                _ => None,
            }
        }
        "MAX_COL_WIDTH" => parts
            .next()?
            .parse::<usize>()
            .ok()
            .map(|width| Op::SetMaxColWidth { width }),
        "COL_WIDTH" => {
            let col = parts.next()?;
            let col = parse_excel_column(col).map(|c| crate::grid::MARGIN_COLS + c as usize)?;
            let width = parts.next().and_then(|w| w.parse::<usize>().ok());
            Some(Op::SetColWidth { col, width })
        }
        "SORT" => {
            let cols = parts
                .map(|s| parse_excel_column(s).map(|c| crate::grid::MARGIN_COLS + c as usize))
                .collect::<Option<Vec<_>>>()?
                .into_iter()
                .map(|col| SortSpec { col, desc: false })
                .collect::<Vec<_>>();
            Some(Op::SetViewSortCols { cols })
        }
        _ => None,
    }
}

pub fn parse_op_line(line: &str) -> Option<Op> {
    parse_op_text(line)
}

impl Op {
    pub fn to_log_line(&self) -> String {
        match self {
            Op::SetCell { addr, value } => format!("SET {} {}", addr_text(addr), value),
            Op::MoveRowRange { from, count, to } => format!("MOVE ROW {from} {count} {to}"),
            Op::MoveColRange { from, count, to } => format!("MOVE COL {from} {count} {to}"),
            Op::SetMainSize {
                main_rows,
                main_cols,
            } => format!("SIZE {main_rows} {main_cols}"),
            Op::SetMaxColWidth { width } => format!("MAX_COL_WIDTH {width}"),
            Op::SetColWidth { col, width } => {
                let name =
                    crate::addr::excel_column_name(col.saturating_sub(crate::grid::MARGIN_COLS));
                match width {
                    Some(w) => format!("COL_WIDTH {name} {w}"),
                    None => format!("COL_WIDTH {name}"),
                }
            }
            Op::SetViewSortCols { cols } => format!(
                "SORT {}",
                cols.iter()
                    .map(|spec| {
                        let name =
                            crate::addr::excel_column_name(spec.col.saturating_sub(MARGIN_COLS));
                        if spec.desc {
                            format!("!{name}")
                        } else {
                            name
                        }
                    })
                    .collect::<Vec<_>>()
                    .join(" ")
            ),
            Op::Undo { target } => format!("UNDO {target}"),
        }
    }
}

impl WorkbookOp {
    pub fn to_log_line(&self) -> String {
        match self {
            WorkbookOp::NewSheet { id, title } => format!("${id}:NEW_SHEET {title}"),
            WorkbookOp::ActivateSheet { id } => format!("${id}:ACTIVATE_SHEET"),
            WorkbookOp::RenameSheet { id, title } => format!("${id}:RENAME_SHEET {title}"),
            WorkbookOp::SheetOp { sheet_id, op } => match op {
                Op::SetCell { addr, value } => {
                    format!("SET ${sheet_id}:{} {value}", addr_text(addr))
                }
                _ => format!("{}{}", sheet_prefix(*sheet_id), op.to_log_line()),
            },
        }
    }
}

fn parse_sheet_set_addr(addr: &str) -> Option<(u32, CellAddr, usize)> {
    if let Some(parsed) = parse_sheet_qualified_cell_ref_at(addr) {
        return Some(parsed);
    }

    let (sheet_id, prefix_len) = parse_sheet_id_prefix_at(addr)?;
    let rest = addr.get(prefix_len..)?;
    let cell_ref = rest.strip_prefix("::")?;
    let (cell_addr, cell_len) = parse_cell_ref_at(cell_ref)?;
    Some((sheet_id, cell_addr, prefix_len + 2 + cell_len))
}

pub fn parse_workbook_line(line: &str) -> Result<WorkbookOp, std::io::Error> {
    let t = line.trim();
    if let Some(rest) = t.strip_prefix("SET ") {
        let mut parts = rest.split_whitespace();
        let addr = parts
            .next()
            .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::InvalidData, "bad SET line"))?;
        if let Some((sheet_id, addr, len)) = parse_sheet_set_addr(addr) {
            let value = rest.get(len..).unwrap_or("").trim_start().to_string();
            return Ok(WorkbookOp::SheetOp {
                sheet_id,
                op: Op::SetCell { addr, value },
            });
        }
    }
    let (sheet_id, prefix_len) = parse_sheet_id_prefix_at(t)
        .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::InvalidData, "missing sheet id"))?;
    let rest = t
        .get(prefix_len..)
        .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::InvalidData, "bad sheet prefix"))?;
    let rest = rest
        .strip_prefix(':')
        .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::InvalidData, "bad sheet prefix"))?;
    let mut parts = rest.split_whitespace();
    let cmd = parts
        .next()
        .map(|s| s.to_ascii_uppercase())
        .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::InvalidData, "empty line"))?;
    let bad = |msg: &'static str| std::io::Error::new(std::io::ErrorKind::InvalidData, msg);

    match cmd.as_str() {
        "NEW_SHEET" => {
            let title = parts.collect::<Vec<_>>().join(" ");
            Ok(WorkbookOp::NewSheet {
                id: sheet_id,
                title,
            })
        }
        "ACTIVATE_SHEET" => Ok(WorkbookOp::ActivateSheet { id: sheet_id }),
        "RENAME_SHEET" => {
            let title = parts.collect::<Vec<_>>().join(" ");
            Ok(WorkbookOp::RenameSheet {
                id: sheet_id,
                title,
            })
        }
        _ => {
            let op = parse_op_line(rest).ok_or_else(|| bad("bad sheet op line"))?;
            Ok(WorkbookOp::SheetOp { sheet_id, op })
        }
    }
}

pub fn apply_workbook_op(
    workbook: &mut WorkbookState,
    active_sheet: &mut u32,
    op: WorkbookOp,
) -> Result<(), std::io::Error> {
    let bad = |msg: &'static str| std::io::Error::new(std::io::ErrorKind::InvalidData, msg);
    match op {
        WorkbookOp::NewSheet { id, title } => {
            if workbook.sheet_index_by_id(id).is_none() {
                workbook.add_sheet_record(SheetRecord {
                    id,
                    title,
                    state: SheetState::new(1, 1),
                });
            }
            Ok(())
        }
        WorkbookOp::ActivateSheet { id } => {
            let idx = workbook
                .sheet_index_by_id(id)
                .ok_or_else(|| bad("unknown sheet id"))?;
            workbook.active_sheet = idx;
            *active_sheet = id;
            Ok(())
        }
        WorkbookOp::RenameSheet { id, title } => {
            let sheet = workbook
                .sheets
                .iter_mut()
                .find(|s| s.id == id)
                .ok_or_else(|| bad("unknown sheet id"))?;
            sheet.title = title;
            Ok(())
        }
        WorkbookOp::SheetOp { sheet_id, op } => {
            let sheet = workbook
                .sheet_mut_by_id(sheet_id)
                .ok_or_else(|| bad("unknown sheet id"))?;
            op.apply(sheet);
            sheet.grid.bump_volatile_seed();
            Ok(())
        }
    }
}

impl SheetState {
    pub fn reverse_op(&self, op: &Op) -> Option<Op> {
        match op {
            Op::SetCell { addr, .. } => {
                let prev_value = self.grid.get(addr).unwrap_or("").to_string();
                Some(Op::SetCell {
                    addr: addr.clone(),
                    value: prev_value,
                })
            }
            Op::MoveRowRange { from, count, to } => {
                let insert_at = if *to > *from { *from + *count } else { *to };
                Some(Op::MoveRowRange {
                    from: insert_at,
                    count: *count,
                    to: *from,
                })
            }
            Op::MoveColRange { from, count, to } => {
                let insert_at = if *to > *from { *from + *count } else { *to };
                Some(Op::MoveColRange {
                    from: insert_at,
                    count: *count,
                    to: *from,
                })
            }
            Op::SetMainSize { .. } => Some(Op::SetMainSize {
                main_rows: self.grid.main_rows() as u32,
                main_cols: self.grid.main_cols() as u32,
            }),
            Op::SetMaxColWidth { .. } => Some(Op::SetMaxColWidth {
                width: self.grid.max_col_width,
            }),
            Op::SetColWidth { col, .. } => Some(Op::SetColWidth {
                col: *col,
                width: self.grid.col_width_overrides.get(col).copied(),
            }),
            Op::SetViewSortCols { .. } => None,
            Op::Undo { .. } => None,
        }
    }
}

/// Replay text log lines from a string (full load).
pub fn replay_lines(data: &str, state: &mut SheetState) -> Result<usize, std::io::Error> {
    let mut n = 0usize;
    for line in data.lines() {
        let t = line.trim();
        if t.is_empty() {
            continue;
        }
        apply_any_line(t, state)?;
        n += 1;
    }
    Ok(n)
}

/// Parse a single line and apply; used when tailing.
pub fn apply_line(line: &str, state: &mut SheetState) -> Result<(), std::io::Error> {
    let t = line.trim();
    if t.is_empty() {
        return Ok(());
    }
    apply_any_line(t, state)
}

pub fn apply_log_line_to_workbook(
    line: &str,
    workbook: &mut WorkbookState,
    active_sheet: &mut u32,
) -> Result<(), std::io::Error> {
    let t = line.trim();
    if t.is_empty() {
        return Ok(());
    }
    if let Ok(op) = parse_workbook_line(t) {
        return apply_workbook_op(workbook, active_sheet, op);
    }
    if let Some(op) = parse_op_line(t) {
        return apply_workbook_op(
            workbook,
            active_sheet,
            WorkbookOp::SheetOp {
                sheet_id: *active_sheet,
                op,
            },
        );
    }
    let sheet = workbook.sheet_mut_by_id(*active_sheet).ok_or_else(|| {
        std::io::Error::new(std::io::ErrorKind::InvalidData, "unknown active sheet")
    })?;
    apply_any_line(t, sheet)
}

fn apply_any_line(line: &str, state: &mut SheetState) -> Result<(), std::io::Error> {
    if line.starts_with("<<<<<<<") || line.starts_with("=======") || line.starts_with(">>>>>>>") {
        return Ok(());
    }
    match line.trim().to_ascii_uppercase().as_str() {
        "SUM" | "TOTAL" | "MEAN" | "AVERAGE" | "AVG" | "MEDIAN" | "MIN" | "MINIMUM" | "MAX"
        | "MAXIMUM" | "COUNT" => return Ok(()),
        _ => {}
    }
    if let Some(op) = parse_op_text(line) {
        op.apply(state);
        return Ok(());
    }
    let mut parts = line.split_whitespace();
    let cmd = match parts.next() {
        Some(cmd) => cmd.to_ascii_uppercase(),
        None => return Ok(()),
    };

    let bad = |msg: &'static str| std::io::Error::new(std::io::ErrorKind::InvalidData, msg);

    match cmd.as_str() {
        "MAX_COL_WIDTH" => {
            let width = parts
                .next()
                .and_then(|w| w.parse::<usize>().ok())
                .ok_or_else(|| bad("bad MAX_COL_WIDTH line"))?;
            if parts.next().is_some() {
                return Err(bad("bad MAX_COL_WIDTH line"));
            }
            state.grid.set_max_col_width(width);
            Ok(())
        }
        "COL_WIDTH" => {
            let col_name = parts.next().ok_or_else(|| bad("bad COL_WIDTH line"))?;
            let col = parse_excel_column(col_name)
                .map(|c| crate::grid::MARGIN_COLS + c as usize)
                .ok_or_else(|| bad("bad COL_WIDTH line"))?;
            let width = match parts.next() {
                Some(w) => Some(w.parse::<usize>().map_err(|_| bad("bad COL_WIDTH line"))?),
                None => None,
            };
            if parts.next().is_some() {
                return Err(bad("bad COL_WIDTH line"));
            }
            state.grid.set_col_width(col, width);
            Ok(())
        }
        "SORT" => {
            let cols = parts
                .map(|s| {
                    let (desc, raw) = if let Some(rest) = s.strip_prefix('!') {
                        (true, rest)
                    } else {
                        (false, s)
                    };
                    parse_excel_column(raw)
                        .map(|c| SortSpec {
                            col: MARGIN_COLS + c as usize,
                            desc,
                        })
                        .ok_or_else(|| bad("bad SORT line"))
                })
                .collect::<Result<Vec<_>, _>>()?;
            state.grid.set_view_sort_cols(cols);
            Ok(())
        }
        "SIZE" => {
            let rows = parts
                .next()
                .and_then(|v| v.parse::<usize>().ok())
                .ok_or_else(|| bad("bad SIZE line"))?;
            let cols = parts
                .next()
                .and_then(|v| v.parse::<usize>().ok())
                .ok_or_else(|| bad("bad SIZE line"))?;
            if parts.next().is_some() {
                return Err(bad("bad SIZE line"));
            }
            state.grid.set_main_size(rows, cols);
            Ok(())
        }
        _ => Err(bad("unrecognized log line")),
    }
}

/// Append one op as text to `path` (creates file if missing).
pub fn append_op(path: &Path, op: &Op) -> std::io::Result<()> {
    let mut f = OpenOptions::new().create(true).append(true).open(path)?;
    let line = match op {
        Op::SetCell { addr, value } => format!("SET {} {}", addr_text(addr), value),
        Op::MoveRowRange { from, count, to } => format!("MOVE ROW {from} {count} {to}"),
        Op::MoveColRange { from, count, to } => format!("MOVE COL {from} {count} {to}"),
        Op::SetMainSize {
            main_rows,
            main_cols,
        } => {
            format!("SIZE {main_rows} {main_cols}")
        }
        Op::SetMaxColWidth { width } => format!("MAX_COL_WIDTH {width}"),
        Op::SetColWidth { col, width } => {
            let name = crate::addr::excel_column_name(col.saturating_sub(crate::grid::MARGIN_COLS));
            match width {
                Some(w) => format!("COL_WIDTH {name} {w}"),
                None => format!("COL_WIDTH {name}"),
            }
        }
        Op::SetViewSortCols { cols } => format!(
            "SORT {}",
            cols.iter()
                .map(|spec| {
                    let name = crate::addr::excel_column_name(spec.col.saturating_sub(MARGIN_COLS));
                    if spec.desc {
                        format!("!{name}")
                    } else {
                        name
                    }
                })
                .collect::<Vec<_>>()
                .join(" ")
        ),
        Op::Undo { target } => format!("UNDO {target}"),
    };
    writeln!(f, "{line}")?;
    f.sync_all()?;
    Ok(())
}

/// Append a plain-text log line.
pub fn append_line(path: &Path, line: &str) -> std::io::Result<()> {
    let mut f = OpenOptions::new().create(true).append(true).open(path)?;
    writeln!(f, "{line}")?;
    f.sync_all()?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::grid::CellAddr;

    #[test]
    fn replay_doc_settings_lines() {
        let mut s = SheetState::new(1, 3);
        apply_line("MAX_COL_WIDTH 17", &mut s).unwrap();
        apply_line("COL_WIDTH B 9", &mut s).unwrap();
        assert_eq!(s.grid.max_col_width, 17);
        assert_eq!(s.grid.col_width(crate::grid::MARGIN_COLS + 1), 9);
    }

    #[test]
    fn replay_size_line() {
        let mut s = SheetState::new(1, 1);
        apply_line("SIZE 7 1", &mut s).unwrap();
        assert_eq!(s.grid.main_rows(), 7);
        assert_eq!(s.grid.main_cols(), 1);
    }

    #[test]
    fn replay_ignores_git_conflict_markers() {
        let mut s = SheetState::new(1, 1);
        let log = concat!(
            "<<<<<<< HEAD\n",
            "SET A1 left\n",
            "=======\n",
            "SET A1 right\n",
            ">>>>>>> other\n"
        );
        replay_lines(log, &mut s).unwrap();
        assert_eq!(
            s.grid.get(&CellAddr::Main { row: 0, col: 0 }),
            Some("right")
        );
    }

    #[test]
    fn replay_ignores_bare_aggregate_labels() {
        let mut s = SheetState::new(1, 1);
        apply_line("TOTAL", &mut s).unwrap();
        apply_line("SUM", &mut s).unwrap();
    }

    #[test]
    fn workbook_sheet_set_log_line_uses_single_colon() {
        let op = WorkbookOp::SheetOp {
            sheet_id: 2,
            op: Op::SetCell {
                addr: CellAddr::Main { row: 1, col: 0 },
                value: "is A2".into(),
            },
        };
        assert_eq!(op.to_log_line(), "SET $2:A2 is A2");
    }

    #[test]
    fn workbook_sheet_set_parser_accepts_legacy_double_colon() {
        let op = parse_workbook_line("SET $2::A2 is A2").unwrap();
        assert_eq!(
            op,
            WorkbookOp::SheetOp {
                sheet_id: 2,
                op: Op::SetCell {
                    addr: CellAddr::Main { row: 1, col: 0 },
                    value: "is A2".into(),
                },
            }
        );
    }
}
