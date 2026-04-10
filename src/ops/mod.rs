//! Append-only log operations and replay onto [`SheetState`].

use crate::addr::{
    parse_cell_ref_at, parse_excel_column, parse_main_range_at, parse_sheet_id_prefix_at,
    parse_ui_column_fragment,
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

    pub fn from_snapshot(snapshot: &WorkbookSnapshot) -> Self {
        let mut workbook = Self {
            sheets: snapshot.sheets.clone(),
            active_sheet: 0,
            next_sheet_id: snapshot.next_sheet_id,
        };
        workbook.active_sheet = workbook
            .sheet_index_by_id(snapshot.active_sheet_id)
            .unwrap_or(0);
        workbook
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
    SetCell {
        addr: CellAddr,
        value: String,
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
    FillRange {
        cells: Vec<(CellAddr, String)>,
    },
    CopyFromTo {
        source: MainRange,
        target: MainRange,
    },
    SetMaxColWidth {
        width: usize,
    },
    SetColWidth {
        col: usize,
        width: Option<usize>,
    },
    SetViewSortCols {
        cols: Vec<SortSpec>,
    },
    Undo {
        target: String,
    },
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub enum WorkbookOp {
    NewSheet {
        id: u32,
        title: String,
    },
    CopySheet {
        source_id: u32,
        id: u32,
        title: String,
    },
    ActivateSheet {
        id: u32,
    },
    RenameSheet {
        id: u32,
        title: String,
    },
    MoveSheet {
        id: u32,
    },
    BalanceReport {
        id: u32,
        title: String,
        source_sheet_id: u32,
        amount_col: usize,
        direction: crate::balance::BalanceDirection,
        row_order: Vec<usize>,
        preserve_formulas: bool,
    },
    SheetOp {
        sheet_id: u32,
        op: Op,
    },
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
            Op::FillRange { cells } => {
                for (addr, value) in cells {
                    state.grid.set(addr, value.clone());
                }
                state.grid.bump_volatile_seed();
            }
            Op::CopyFromTo { source, target } => {
                let rows = source.row_end.saturating_sub(source.row_start);
                let cols = source.col_end.saturating_sub(source.col_start);
                let target_rows = target.row_end.saturating_sub(target.row_start);
                let target_cols = target.col_end.saturating_sub(target.col_start);
                let rows = rows.min(target_rows);
                let cols = cols.min(target_cols);

                let mut cells = Vec::with_capacity(rows.saturating_mul(cols) as usize);
                for r in 0..rows {
                    for c in 0..cols {
                        let src = CellAddr::Main {
                            row: source.row_start + r,
                            col: source.col_start + c,
                        };
                        let dst = CellAddr::Main {
                            row: target.row_start + r,
                            col: target.col_start + c,
                        };
                        cells.push((dst, state.grid.get(&src).unwrap_or("").to_string()));
                    }
                }
                for (addr, value) in cells {
                    state.grid.set(&addr, value);
                }
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
            crate::addr::ui_column_fragment(*col as usize, 0)
        ),
        CellAddr::Footer { row, col } => {
            format!(
                "_{}{}",
                *row as usize + 1,
                crate::addr::ui_column_fragment(*col as usize, 0)
            )
        }
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

fn main_range_text(range: &MainRange) -> String {
    let start = CellAddr::Main {
        row: range.row_start,
        col: range.col_start,
    };
    let end = CellAddr::Main {
        row: range.row_end.saturating_sub(1),
        col: range.col_end.saturating_sub(1),
    };
    format!("{}:{}", addr_text(&start), addr_text(&end))
}

fn encode_log_value(value: &str) -> String {
    let mut out = String::new();
    for b in value.bytes() {
        match b {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'_' | b'.' | b'~' => {
                out.push(b as char)
            }
            _ => out.push_str(&format!("%{b:02X}")),
        }
    }
    out
}

fn decode_log_value(value: &str) -> Option<String> {
    let bytes = value.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0usize;
    while i < bytes.len() {
        if bytes[i] == b'%' {
            if i + 2 >= bytes.len() {
                return None;
            }
            let hi = (bytes[i + 1] as char).to_digit(16)? as u8;
            let lo = (bytes[i + 2] as char).to_digit(16)? as u8;
            out.push((hi << 4) | lo);
            i += 3;
        } else {
            out.push(bytes[i]);
            i += 1;
        }
    }
    String::from_utf8(out).ok()
}

fn parse_op_text(line: &str) -> Option<Op> {
    let mut parts = line.split_whitespace();
    let cmd = parts.next()?.to_ascii_uppercase();
    match cmd.as_str() {
        "SET" => {
            let addr = parts.next()?;
            let value = parts.collect::<Vec<_>>().join(" ");
            let (addr, _) = parse_log_addr(addr, 0)?;
            Some(Op::SetCell { addr, value })
        }
        "FILL" => {
            let mut cells = Vec::new();
            for token in parts {
                let (addr, value) = token.split_once('=')?;
                let (addr, _) = parse_log_addr(addr, 0)?;
                cells.push((addr, decode_log_value(value)?));
            }
            Some(Op::FillRange { cells })
        }
        "COPY_FROM_TO" => {
            let source_text = parts.next()?;
            let target_text = parts.next()?;
            if parts.next().is_some() {
                return None;
            }
            let (source, _) = parse_main_range_at(source_text)?;
            let (target, _) = parse_main_range_at(target_text)?;
            Some(Op::CopyFromTo { source, target })
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
        "SIZE" => {
            let rows = parts.next()?.parse::<u32>().ok()?;
            let cols = parts.next()?.parse::<u32>().ok()?;
            Some(Op::SetMainSize {
                main_rows: rows,
                main_cols: cols,
            })
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
            Op::FillRange { cells } => format!(
                "FILL {}",
                cells
                    .iter()
                    .map(|(addr, value)| format!("{}={}", addr_text(addr), encode_log_value(value)))
                    .collect::<Vec<_>>()
                    .join(" ")
            ),
            Op::CopyFromTo { source, target } => {
                format!(
                    "COPY_FROM_TO {} {}",
                    main_range_text(source),
                    main_range_text(target)
                )
            }
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
    pub fn to_log_line(&self, main_cols: usize) -> String {
        match self {
            WorkbookOp::NewSheet { id, title } => format!("${id}:NEW_SHEET {title}"),
            WorkbookOp::CopySheet {
                source_id,
                id,
                title,
            } => format!("${id}:COPY_SHEET {source_id} {title}"),
            WorkbookOp::ActivateSheet { id } => format!("${id}:ACTIVATE_SHEET"),
            WorkbookOp::RenameSheet { id, title } => format!("${id}:RENAME_SHEET {title}"),
            WorkbookOp::MoveSheet { id } => format!("${id}:MOVE_SHEET"),
            WorkbookOp::BalanceReport {
                id,
                title,
                source_sheet_id,
                amount_col,
                direction,
                row_order,
                preserve_formulas,
            } => format!(
                "${id}:BALANCE_REPORT {title} {source_sheet_id} {amount_col} {:?} {} {}",
                direction,
                if *preserve_formulas { 1 } else { 0 },
                row_order
                    .iter()
                    .map(|n| n.to_string())
                    .collect::<Vec<_>>()
                    .join(",")
            ),
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
    let (sheet_id, prefix_len) = parse_sheet_id_prefix_at(addr)?;
    let rest = addr.get(prefix_len..)?;
    let cell_ref = rest.strip_prefix(':')?;
    let (cell_addr, cell_len) = parse_log_addr(cell_ref, 0)?;
    Some((sheet_id, cell_addr, prefix_len + 1 + cell_len))
}

fn parse_log_addr(addr: &str, main_cols: usize) -> Option<(CellAddr, usize)> {
    let bytes = addr.as_bytes();
    if bytes.is_empty() {
        return None;
    }
    match bytes[0] {
        b'~' => {
            let rest = &addr[1..];
            let row_digits = rest.chars().take_while(|c| c.is_ascii_digit()).count();
            if row_digits == 0 {
                return None;
            }
            let row_num: usize = rest[..row_digits].parse().ok()?;
            let row = if row_num == 0 || row_num > crate::grid::HEADER_ROWS {
                return None;
            } else {
                (crate::grid::HEADER_ROWS - row_num) as u8
            };
            let after = &rest[row_digits..];
            let col_len = after.chars().take_while(|c| c.is_ascii_uppercase()).count();
            if col_len == 0 {
                return None;
            }
            let col = crate::addr::parse_excel_column(&after[..col_len])?;
            Some((CellAddr::Header { row, col }, 1 + row_digits + col_len))
        }
        b'_' => {
            let rest = &addr[1..];
            let row_digits = rest.chars().take_while(|c| c.is_ascii_digit()).count();
            if row_digits == 0 {
                return None;
            }
            let row_num: usize = rest[..row_digits].parse().ok()?;
            let row = if row_num == 0 || row_num > crate::grid::FOOTER_ROWS {
                return None;
            } else {
                (row_num - 1) as u8
            };
            let after = &rest[row_digits..];
            let col_len = after.chars().take_while(|c| c.is_ascii_uppercase()).count();
            if col_len == 0 {
                return None;
            }
            let col = crate::addr::parse_excel_column(&after[..col_len])?;
            Some((CellAddr::Footer { row, col }, 1 + row_digits + col_len))
        }
        b'[' | b']' => {
            let left_side = bytes[0] == b'[';
            let rest = &addr[1..];
            let col_len = rest.chars().take_while(|c| c.is_ascii_uppercase()).count();
            if col_len != 1 {
                return None;
            }
            let col = crate::addr::parse_mirror_margin_column_name(&rest[..col_len], left_side)?;
            let after = &rest[col_len..];
            let row_digits = after.chars().take_while(|c| c.is_ascii_digit()).count();
            if row_digits == 0 {
                return None;
            }
            let row: u32 = after[..row_digits].parse().ok()?;
            if row == 0 {
                return None;
            }
            let consumed = 1 + col_len + row_digits;
            Some((
                if left_side {
                    CellAddr::Left { col, row: row - 1 }
                } else {
                    CellAddr::Right { col, row: row - 1 }
                },
                consumed,
            ))
        }
        _ => parse_cell_ref_at(addr),
    }
}

pub fn parse_workbook_line(line: &str) -> Result<WorkbookOp, std::io::Error> {
    let t = line.trim();
    if let Some(rest) = t.strip_prefix("SET ") {
        let mut parts = rest.split_whitespace();
        let addr = parts
            .next()
            .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::InvalidData, "bad SET line"))?;
        if let Some((sheet_id, cell_addr, len)) = parse_sheet_set_addr(addr) {
            let value = rest.get(len..).unwrap_or("").trim_start().to_string();
            return Ok(WorkbookOp::SheetOp {
                sheet_id,
                op: Op::SetCell {
                    addr: cell_addr,
                    value,
                },
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
        "COPY_SHEET" => {
            let source_id = parts
                .next()
                .and_then(|v| v.parse::<u32>().ok())
                .ok_or_else(|| bad("bad sheet copy line"))?;
            let title = parts.collect::<Vec<_>>().join(" ");
            Ok(WorkbookOp::CopySheet {
                source_id,
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
        "MOVE_SHEET" => Ok(WorkbookOp::MoveSheet { id: sheet_id }),
        "BALANCE_REPORT" => {
            let title = parts
                .next()
                .ok_or_else(|| bad("bad balance line"))?
                .to_string();
            let source_sheet_id = parts
                .next()
                .and_then(|v| v.parse::<u32>().ok())
                .ok_or_else(|| bad("bad balance line"))?;
            let amount_col = parts
                .next()
                .and_then(|v| v.parse::<usize>().ok())
                .ok_or_else(|| bad("bad balance line"))?;
            let direction = match parts.next() {
                Some("PosToNeg") => crate::balance::BalanceDirection::PosToNeg,
                Some("NegToPos") => crate::balance::BalanceDirection::NegToPos,
                _ => return Err(bad("bad balance line")),
            };
            let preserve_formulas = match parts.next() {
                Some("1") => true,
                Some("0") => false,
                _ => return Err(bad("bad balance line")),
            };
            let row_order = parts
                .next()
                .unwrap_or("")
                .split(',')
                .filter(|s| !s.is_empty())
                .map(|s| s.parse::<usize>().map_err(|_| bad("bad balance line")))
                .collect::<Result<Vec<_>, _>>()?;
            Ok(WorkbookOp::BalanceReport {
                id: sheet_id,
                title,
                source_sheet_id,
                amount_col,
                direction,
                row_order,
                preserve_formulas,
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
        WorkbookOp::CopySheet {
            source_id,
            id,
            title,
        } => {
            let source = workbook
                .sheets
                .iter()
                .find(|s| s.id == source_id)
                .ok_or_else(|| bad("unknown sheet id"))?
                .clone();
            if let Some(idx) = workbook.sheet_index_by_id(id) {
                workbook.sheets[idx].title = title;
                workbook.sheets[idx].state = source.state.clone();
            } else {
                workbook.add_sheet_record(SheetRecord {
                    id,
                    title,
                    state: source.state,
                });
            }
            workbook.active_sheet = workbook
                .sheet_index_by_id(id)
                .unwrap_or(workbook.active_sheet);
            *active_sheet = id;
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
        WorkbookOp::MoveSheet { id } => {
            let idx = workbook
                .sheet_index_by_id(id)
                .ok_or_else(|| bad("unknown sheet id"))?;
            let sheet = workbook.sheets.remove(idx);
            workbook.sheets.push(sheet);
            workbook.active_sheet = workbook
                .sheet_index_by_id(id)
                .unwrap_or(workbook.active_sheet);
            *active_sheet = id;
            Ok(())
        }
        WorkbookOp::BalanceReport {
            id,
            title,
            source_sheet_id,
            amount_col,
            direction,
            row_order,
            preserve_formulas,
        } => {
            let source = workbook
                .sheets
                .iter()
                .find(|s| s.id == source_sheet_id)
                .ok_or_else(|| bad("unknown sheet id"))?
                .clone();
            let report = crate::balance::BalanceReport {
                direction,
                amount_col,
                groups: Vec::new(),
                leftovers: row_order,
            };
            let plan = crate::balance::balance_copy_plan(
                source_sheet_id,
                source.title.clone(),
                id,
                title,
                amount_col,
                source.state.grid.main_rows(),
                &report,
                preserve_formulas,
            );
            let mut target_state =
                SheetState::new(source.state.grid.main_rows(), source.state.grid.main_cols());
            crate::balance::apply_balance_copy(&source.state, &mut target_state, &plan);
            if workbook.sheet_index_by_id(id).is_none() {
                workbook.add_sheet_record(SheetRecord {
                    id,
                    title: plan.target_title.clone(),
                    state: target_state.clone(),
                });
            }
            let sheet = workbook
                .sheets
                .iter_mut()
                .find(|s| s.id == id)
                .ok_or_else(|| bad("unknown sheet id"))?;
            sheet.title = plan.target_title;
            sheet.state = target_state;
            workbook.active_sheet = workbook
                .sheet_index_by_id(id)
                .unwrap_or(workbook.active_sheet);
            *active_sheet = id;
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
            Op::FillRange { cells } => Some(Op::FillRange {
                cells: cells
                    .iter()
                    .map(|(addr, _)| {
                        let prev_value = self.grid.get(addr).unwrap_or("").to_string();
                        (addr.clone(), prev_value)
                    })
                    .collect(),
            }),
            Op::CopyFromTo { target, .. } => {
                let mut cells = Vec::new();
                for r in target.row_start..target.row_end {
                    for c in target.col_start..target.col_end {
                        let addr = CellAddr::Main { row: r, col: c };
                        let prev_value = self.grid.get(&addr).unwrap_or("").to_string();
                        cells.push((addr, prev_value));
                    }
                }
                Some(Op::FillRange { cells })
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

/// Replay text log lines until the first malformed entry.
pub fn replay_lines_partial(
    data: &str,
    state: &mut SheetState,
) -> Result<(usize, Option<usize>, Option<std::io::Error>), std::io::Error> {
    let mut n = 0usize;
    for (idx, line) in data.lines().enumerate() {
        let t = line.trim();
        if t.is_empty() {
            continue;
        }
        if let Err(err) = apply_any_line(t, state) {
            return Ok((n, Some(idx + 1), Some(err)));
        }
        n += 1;
    }
    Ok((n, None, None))
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
        Op::FillRange { cells } => format!(
            "FILL {}",
            cells
                .iter()
                .map(|(addr, value)| format!("{}={}", addr_text(addr), encode_log_value(value)))
                .collect::<Vec<_>>()
                .join(" ")
        ),
        Op::CopyFromTo { source, target } => {
            format!(
                "COPY_FROM_TO {} {}",
                main_range_text(source),
                main_range_text(target)
            )
        }
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
        assert_eq!(op.to_log_line(1), "SET $2:A2 is A2");
    }

    #[test]
    fn workbook_sheet_set_parser_accepts_ui_notation() {
        let op = parse_workbook_line("SET $2:A2 is A2").unwrap();
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

    #[test]
    fn workbook_log_parser_keeps_header_footer_columns_absolute() {
        let header = parse_workbook_line("SET $1:~1K x").unwrap();
        let footer = parse_workbook_line("SET $1:_1K y").unwrap();
        assert!(matches!(
            header,
            WorkbookOp::SheetOp {
                op: Op::SetCell {
                    addr: CellAddr::Header { col: 10, .. },
                    ..
                },
                ..
            }
        ));
        assert!(matches!(
            footer,
            WorkbookOp::SheetOp {
                op: Op::SetCell {
                    addr: CellAddr::Footer { col: 10, .. },
                    ..
                },
                ..
            }
        ));
    }

    #[test]
    fn fill_range_round_trips_through_log_line() {
        let op = Op::FillRange {
            cells: vec![
                (CellAddr::Main { row: 0, col: 0 }, "1".into()),
                (CellAddr::Main { row: 0, col: 1 }, "2".into()),
            ],
        };
        let line = op.to_log_line();
        assert_eq!(line, "FILL A1=1 B1=2");
        assert_eq!(parse_op_line(&line), Some(op));
    }

    #[test]
    fn copy_from_to_round_trips_through_log_line() {
        let op = Op::CopyFromTo {
            source: MainRange {
                row_start: 0,
                row_end: 2,
                col_start: 0,
                col_end: 2,
            },
            target: MainRange {
                row_start: 2,
                row_end: 4,
                col_start: 1,
                col_end: 3,
            },
        };
        let line = op.to_log_line();
        assert_eq!(line, "COPY_FROM_TO A1:B2 B3:C4");
        assert_eq!(parse_op_line(&line), Some(op));
    }

    #[test]
    fn balance_report_replays_as_copied_sheet() {
        let mut workbook = WorkbookState::new();
        workbook.add_sheet("Src".into(), SheetState::new(2, 2));
        let src_idx = workbook.sheet_index_by_id(1).unwrap();
        workbook.sheets[src_idx]
            .state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "10".into());
        workbook.sheets[src_idx]
            .state
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "=A1".into());
        workbook.sheets[src_idx]
            .state
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "-10".into());
        workbook.sheets[src_idx]
            .state
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "=A2".into());

        let op = WorkbookOp::BalanceReport {
            id: 2,
            title: "Dst".into(),
            source_sheet_id: 1,
            amount_col: 0,
            direction: crate::balance::BalanceDirection::PosToNeg,
            row_order: vec![1, 0],
            preserve_formulas: true,
        };

        let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
        apply_workbook_op(&mut workbook, &mut active_sheet, op).unwrap();

        let dst = workbook.sheet_index_by_id(2).unwrap();
        assert_eq!(
            workbook.sheets[dst]
                .state
                .grid
                .get(&CellAddr::Main { row: 0, col: 1 }),
            Some("=A1")
        );
        assert_eq!(
            workbook.sheets[dst]
                .state
                .grid
                .get(&CellAddr::Main { row: 1, col: 1 }),
            Some("=A2")
        );
    }

    #[test]
    fn copy_sheet_replays_as_one_log_op() {
        let mut workbook = WorkbookState::new();
        workbook.sheets[0]
            .state
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "src".into());
        let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
        apply_workbook_op(
            &mut workbook,
            &mut active_sheet,
            WorkbookOp::CopySheet {
                source_id: 1,
                id: 2,
                title: "Copy".into(),
            },
        )
        .unwrap();

        assert_eq!(workbook.sheet_count(), 2);
        assert_eq!(workbook.sheets[1].id, 2);
        assert_eq!(workbook.sheets[1].title, "Copy");
        assert_eq!(
            workbook.sheets[1]
                .state
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 }),
            Some("src")
        );
    }

    #[test]
    fn move_sheet_preserves_ids_while_reordering() {
        let mut workbook = WorkbookState::new();
        workbook.add_sheet("Two".into(), SheetState::new(1, 1));
        workbook.add_sheet("Three".into(), SheetState::new(1, 1));
        let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
        apply_workbook_op(
            &mut workbook,
            &mut active_sheet,
            WorkbookOp::MoveSheet { id: 1 },
        )
        .unwrap();

        assert_eq!(
            workbook.sheets.iter().map(|s| s.id).collect::<Vec<_>>(),
            vec![2, 3, 1]
        );
        assert_eq!(active_sheet, 1);
    }
}
