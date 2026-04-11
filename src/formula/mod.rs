//! `=...` cell formulas: parse, evaluate, display.

use crate::addr::{excel_column_name, parse_cell_ref_at, parse_main_range_at};
use crate::grid::{CellAddr, Grid, MainRange, HEADER_ROWS, MARGIN_COLS};
use crate::ops::WorkbookState;
use std::cell::RefCell;

mod functions;

thread_local! {
    static EVAL_WORKBOOK: RefCell<Option<WorkbookState>> = const { RefCell::new(None) };
}

#[derive(Clone, Debug)]
pub struct FormulaCopyContext {
    pub source_sheet_id: u32,
    pub target_sheet_id: u32,
    pub row_map: Vec<u32>,
    pub main_cols: usize,
}

pub struct EvalContextGuard;

impl Drop for EvalContextGuard {
    fn drop(&mut self) {
        EVAL_WORKBOOK.with(|wb| *wb.borrow_mut() = None);
    }
}

pub fn set_eval_context(workbook: &WorkbookState) -> EvalContextGuard {
    EVAL_WORKBOOK.with(|wb| *wb.borrow_mut() = Some(workbook.clone()));
    EvalContextGuard
}

fn workbook_lookup(sheet_id: u32) -> Option<Grid> {
    EVAL_WORKBOOK.with(|wb| {
        wb.borrow()
            .as_ref()
            .and_then(|w| w.sheets.iter().find(|s| s.id == sheet_id))
            .map(|s| s.state.grid.clone())
    })
}

fn workbook_lookup_sheet_ref(sheet: &str) -> Option<(u32, Grid)> {
    EVAL_WORKBOOK.with(|wb| {
        let wb = wb.borrow();
        let wb = wb.as_ref()?;
        if let Some(rest) = sheet.strip_prefix("Sheet") {
            if let Ok(id) = rest.parse::<u32>() {
                let rec = wb.sheets.iter().find(|s| s.id == id)?;
                return Some((rec.id, rec.state.grid.clone()));
            }
        }
        let rec = wb.sheets.iter().find(|s| s.title == sheet)?;
        Some((rec.id, rec.state.grid.clone()))
    })
}
const DEFAULT_BUDGET: usize = 10_000;

const SPEED_OF_LIGHT_MPS: f64 = 299_792_458.0;

/// Evaluation step budget for one aggregate range scan (many cells).
pub const EVAL_BUDGET_AGG: usize = 1_000_000;

/// Result of evaluating a cell (formula or plain).
#[derive(Clone, Debug, PartialEq)]
pub enum EvalResult {
    Number(f64),
    Text(String),
    Array(Vec<Vec<EvalResult>>),
    /// Display as `#msg` in the UI.
    Error(&'static str),
}

impl EvalResult {
    pub(crate) fn scalar_coerce(self) -> EvalResult {
        match self {
            EvalResult::Array(rows) => rows
                .into_iter()
                .next()
                .and_then(|row| row.into_iter().next())
                .unwrap_or(EvalResult::Error("CALC"))
                .scalar_coerce(),
            other => other,
        }
    }

    fn top_left(&self) -> Option<&EvalResult> {
        match self {
            EvalResult::Array(rows) => rows.first().and_then(|row| row.first()),
            _ => Some(self),
        }
    }

    fn as_text(&self) -> Option<&str> {
        match self {
            EvalResult::Text(s) => Some(s.as_str()),
            _ => None,
        }
    }
}

fn parse_number_literal(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    t.parse::<f64>().ok()
}

fn resolve_name(name: &str, bindings: &[(String, EvalResult)]) -> Option<EvalResult> {
    if let Some((_, value)) = bindings.iter().rev().find(|(n, _)| n == name) {
        return Some(value.clone());
    }

    match name {
        "π" => Some(EvalResult::Number(std::f64::consts::PI)),
        "e" => Some(EvalResult::Number(std::f64::consts::E)),
        "c" => Some(EvalResult::Number(SPEED_OF_LIGHT_MPS)),
        _ => None,
    }
}

fn split_labeled_formula(raw: &str) -> Option<(&str, &str)> {
    let t = raw.trim();
    let expr = t.strip_prefix('=')?;
    let (expr, label) = expr.rsplit_once(" -- ")?;
    let expr = expr.trim();
    let label = label.trim();
    if expr.is_empty() || label.is_empty() {
        return None;
    }
    Some((expr, label))
}

fn render_addr(addr: &CellAddr) -> String {
    crate::addr::cell_ref_text(addr, 0)
}

fn render_ast(ast: &Ast) -> String {
    match ast {
        Ast::Number(n) => format!("{n}"),
        Ast::Text(s) => format!("\"{}\"", s.replace('"', "\"\"")),
        Ast::Name(name) => name.clone(),
        Ast::Ref(addr) => render_addr(addr),
        Ast::SheetRef { sheet_id, addr } => format!("#{sheet_id}!{}", render_addr(addr)),
        Ast::Range(range) => format!(
            "{}:{}",
            render_addr(&CellAddr::Main {
                row: range.row_start,
                col: range.col_start,
            }),
            render_addr(&CellAddr::Main {
                row: range.row_end.saturating_sub(1),
                col: range.col_end.saturating_sub(1),
            })
        ),
        Ast::Neg(a) => format!("(-{})", render_ast(a)),
        Ast::Add(a, b) => format!("({}+{})", render_ast(a), render_ast(b)),
        Ast::Sub(a, b) => format!("({}-{})", render_ast(a), render_ast(b)),
        Ast::Mul(a, b) => format!("({}*{})", render_ast(a), render_ast(b)),
        Ast::Div(a, b) => format!("({}/{})", render_ast(a), render_ast(b)),
        Ast::Pow(a, b) => format!("({}^{})", render_ast(a), render_ast(b)),
        Ast::Call { name, args } => format!(
            "{}({})",
            name,
            args.iter().map(render_ast).collect::<Vec<_>>().join(",")
        ),
    }
}

fn translate_addr(addr: &CellAddr, ctx: &FormulaCopyContext) -> Option<CellAddr> {
    match addr {
        CellAddr::Header { .. } | CellAddr::Footer { .. } => Some(addr.clone()),
        CellAddr::Main { row, col } => Some(CellAddr::Main {
            row: *ctx.row_map.get(*row as usize)?,
            col: *col,
        }),
        CellAddr::Left { col, row } => Some(CellAddr::Left {
            col: *col,
            row: *ctx.row_map.get(*row as usize)?,
        }),
        CellAddr::Right { col, row } => Some(CellAddr::Right {
            col: *col,
            row: *ctx.row_map.get(*row as usize)?,
        }),
    }
}

fn translate_range(range: &MainRange, ctx: &FormulaCopyContext) -> Option<MainRange> {
    let mut mapped_rows = Vec::new();
    for row in range.row_start..range.row_end {
        mapped_rows.push(*ctx.row_map.get(row as usize)?);
    }
    if mapped_rows.is_empty() {
        return Some(range.clone());
    }
    if !mapped_rows.windows(2).all(|w| w[1] == w[0] + 1) {
        return None;
    }
    Some(MainRange {
        row_start: *mapped_rows.first()?,
        row_end: mapped_rows.last().copied()?.saturating_add(1),
        col_start: range.col_start,
        col_end: range.col_end,
    })
}

pub fn translate_formula_text_by_offset(
    raw: &str,
    row_delta: i32,
    col_delta: i32,
) -> Option<String> {
    if row_delta == 0 && col_delta == 0 {
        return Some(raw.trim().to_string());
    }
    let t = raw.trim();
    let (expr, label) =
        split_labeled_formula(t).map_or((t, None), |(expr, label)| (expr, Some(label)));
    if !expr.starts_with('=') {
        return None;
    }
    let mut parser = Parser {
        s: &expr[1..],
        i: 0,
        main_cols: 0,
    };
    let ast = parser.parse_expr().ok()?;
    parser.skip_ws();
    if parser.i != parser.s.len() {
        return None;
    }
    let translated = translate_ast_by_offset(&ast, row_delta, col_delta)?;
    let mut out = format!("={}", render_ast(&translated));
    if let Some(label) = label {
        out.push_str(" -- ");
        out.push_str(label);
    }
    Some(out)
}

fn translate_cell_addr_by_offset(
    addr: &CellAddr,
    row_delta: i32,
    col_delta: i32,
) -> Option<CellAddr> {
    let shift_u32 = |v: u32, delta: i32| -> Option<u32> {
        if delta >= 0 {
            v.checked_add(delta as u32)
        } else {
            v.checked_sub(delta.unsigned_abs())
        }
    };
    match addr {
        CellAddr::Header { row, col } => Some(CellAddr::Header {
            row: shift_u32(*row as u32, row_delta)? as u8,
            col: shift_u32(*col, col_delta)?,
        }),
        CellAddr::Footer { row, col } => Some(CellAddr::Footer {
            row: shift_u32(*row as u32, row_delta)? as u8,
            col: shift_u32(*col, col_delta)?,
        }),
        CellAddr::Main { row, col } => Some(CellAddr::Main {
            row: shift_u32(*row, row_delta)?,
            col: shift_u32(*col, col_delta)?,
        }),
        CellAddr::Left { col, row } => Some(CellAddr::Left {
            col: shift_u32(*col as u32, col_delta)? as u8,
            row: shift_u32(*row, row_delta)?,
        }),
        CellAddr::Right { col, row } => Some(CellAddr::Right {
            col: shift_u32(*col as u32, col_delta)? as u8,
            row: shift_u32(*row, row_delta)?,
        }),
    }
}

fn translate_range_by_offset(
    range: &MainRange,
    row_delta: i32,
    _col_delta: i32,
) -> Option<MainRange> {
    let shift_u32 = |v: u32, delta: i32| -> Option<u32> {
        if delta >= 0 {
            v.checked_add(delta as u32)
        } else {
            v.checked_sub(delta.unsigned_abs())
        }
    };
    Some(MainRange {
        row_start: shift_u32(range.row_start, row_delta)?,
        row_end: shift_u32(range.row_end, row_delta)?,
        col_start: range.col_start,
        col_end: range.col_end,
    })
}

fn translate_ast_by_offset(ast: &Ast, row_delta: i32, col_delta: i32) -> Option<Ast> {
    Some(match ast {
        Ast::Number(n) => Ast::Number(*n),
        Ast::Text(s) => Ast::Text(s.clone()),
        Ast::Name(name) => Ast::Name(name.clone()),
        Ast::Ref(addr) => Ast::Ref(translate_cell_addr_by_offset(addr, row_delta, col_delta)?),
        Ast::SheetRef { sheet_id, addr } => Ast::SheetRef {
            sheet_id: *sheet_id,
            addr: translate_cell_addr_by_offset(addr, row_delta, col_delta)?,
        },
        Ast::Range(range) => Ast::Range(translate_range_by_offset(range, row_delta, col_delta)?),
        Ast::Neg(a) => Ast::Neg(Box::new(translate_ast_by_offset(a, row_delta, col_delta)?)),
        Ast::Add(a, b) => Ast::Add(
            Box::new(translate_ast_by_offset(a, row_delta, col_delta)?),
            Box::new(translate_ast_by_offset(b, row_delta, col_delta)?),
        ),
        Ast::Sub(a, b) => Ast::Sub(
            Box::new(translate_ast_by_offset(a, row_delta, col_delta)?),
            Box::new(translate_ast_by_offset(b, row_delta, col_delta)?),
        ),
        Ast::Mul(a, b) => Ast::Mul(
            Box::new(translate_ast_by_offset(a, row_delta, col_delta)?),
            Box::new(translate_ast_by_offset(b, row_delta, col_delta)?),
        ),
        Ast::Div(a, b) => Ast::Div(
            Box::new(translate_ast_by_offset(a, row_delta, col_delta)?),
            Box::new(translate_ast_by_offset(b, row_delta, col_delta)?),
        ),
        Ast::Pow(a, b) => Ast::Pow(
            Box::new(translate_ast_by_offset(a, row_delta, col_delta)?),
            Box::new(translate_ast_by_offset(b, row_delta, col_delta)?),
        ),
        Ast::Call { name, args } => Ast::Call {
            name: name.clone(),
            args: args
                .iter()
                .map(|a| translate_ast_by_offset(a, row_delta, col_delta))
                .collect::<Option<Vec<_>>>()?,
        },
    })
}

fn translate_ast(ast: &Ast, ctx: &FormulaCopyContext) -> Option<Ast> {
    Some(match ast {
        Ast::Number(n) => Ast::Number(*n),
        Ast::Text(s) => Ast::Text(s.clone()),
        Ast::Name(name) => Ast::Name(name.clone()),
        Ast::Ref(addr) => Ast::Ref(translate_addr(addr, ctx)?),
        Ast::SheetRef { sheet_id, addr } => {
            if *sheet_id == ctx.source_sheet_id {
                Ast::Ref(translate_addr(addr, ctx)?)
            } else {
                Ast::SheetRef {
                    sheet_id: *sheet_id,
                    addr: addr.clone(),
                }
            }
        }
        Ast::Range(range) => Ast::Range(translate_range(range, ctx)?),
        Ast::Neg(a) => Ast::Neg(Box::new(translate_ast(a, ctx)?)),
        Ast::Add(a, b) => Ast::Add(
            Box::new(translate_ast(a, ctx)?),
            Box::new(translate_ast(b, ctx)?),
        ),
        Ast::Sub(a, b) => Ast::Sub(
            Box::new(translate_ast(a, ctx)?),
            Box::new(translate_ast(b, ctx)?),
        ),
        Ast::Mul(a, b) => Ast::Mul(
            Box::new(translate_ast(a, ctx)?),
            Box::new(translate_ast(b, ctx)?),
        ),
        Ast::Div(a, b) => Ast::Div(
            Box::new(translate_ast(a, ctx)?),
            Box::new(translate_ast(b, ctx)?),
        ),
        Ast::Pow(a, b) => Ast::Pow(
            Box::new(translate_ast(a, ctx)?),
            Box::new(translate_ast(b, ctx)?),
        ),
        Ast::Call { name, args } => Ast::Call {
            name: name.clone(),
            args: args
                .iter()
                .map(|a| translate_ast(a, ctx))
                .collect::<Option<Vec<_>>>()?,
        },
    })
}

pub fn translate_formula_text(raw: &str, ctx: &FormulaCopyContext) -> Option<String> {
    let t = raw.trim();
    let (expr, label) =
        split_labeled_formula(t).map_or((t, None), |(expr, label)| (expr, Some(label)));
    if !expr.starts_with('=') {
        return None;
    }
    let mut parser = Parser {
        s: &expr[1..],
        i: 0,
        main_cols: ctx.main_cols,
    };
    let ast = parser.parse_expr().ok()?;
    parser.skip_ws();
    if parser.i != parser.s.len() {
        return None;
    }
    let translated = translate_ast(&ast, ctx)?;
    let mut out = format!("={}", render_ast(&translated));
    if let Some(label) = label {
        out.push_str(" -- ");
        out.push_str(label);
    }
    Some(out)
}

fn rewrite_header_template(expr: &str, row: u32) -> String {
    let mut out = String::new();
    let bytes = expr.as_bytes();
    let mut i = 0usize;
    while i < bytes.len() {
        let b = bytes[i];
        if b.is_ascii_alphabetic() {
            let start = i;
            while i < bytes.len() && bytes[i].is_ascii_alphabetic() {
                i += 1;
            }
            let token = &expr[start..i];
            if i < bytes.len() && (bytes[i].is_ascii_digit() || bytes[i] == b'(') {
                out.push_str(token);
            } else {
                out.push_str(&token.to_ascii_uppercase());
                out.push_str(&(row + 1).to_string());
            }
        } else {
            out.push(b as char);
            i += 1;
        }
    }
    out
}

fn rewrite_row_template(expr: &str, col: usize) -> String {
    let mut out = String::new();
    let mut chars = expr.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == ':' {
            let mut digits = String::new();
            while matches!(chars.peek(), Some(c) if c.is_ascii_digit()) {
                digits.push(chars.next().unwrap());
            }
            if digits.is_empty() {
                out.push(ch);
            } else {
                out.push_str(&excel_column_name(col));
                out.push_str(&digits);
            }
        } else {
            out.push(ch);
        }
    }
    out
}

fn control_formula_expr(grid: &Grid, addr: &CellAddr) -> Option<String> {
    let raw = grid.get(addr)?;
    let (expr, _label) = split_labeled_formula(raw)?;
    Some(expr.to_string())
}

fn control_formula_label(grid: &Grid, addr: &CellAddr) -> Option<String> {
    let raw = grid.get(addr)?;
    let (_expr, label) = split_labeled_formula(raw)?;
    Some(label.to_string())
}

fn templated_formula(grid: &Grid, addr: &CellAddr) -> Option<String> {
    let CellAddr::Main { row, col } = addr else {
        return None;
    };

    let header_addr = CellAddr::Header {
        row: (HEADER_ROWS - 1) as u8,
        col: (MARGIN_COLS as u32) + *col,
    };
    if let Some(expr) = control_formula_expr(grid, &header_addr) {
        return Some(format!("={}", rewrite_header_template(&expr, *row)));
    }

    let left_addr = CellAddr::Left {
        col: (MARGIN_COLS - 1) as u8,
        row: *row,
    };
    if let Some(expr) = control_formula_expr(grid, &left_addr) {
        return Some(format!("={}", rewrite_row_template(&expr, *col as usize)));
    }

    None
}

/// True if the stored cell text is a formula (`=` prefix after trim).
pub fn is_formula(raw: &str) -> bool {
    raw.trim_start().starts_with('=')
}

/// Numeric value for aggregation: formulas evaluate to a number if possible; plain text uses `f64` parse.
pub fn effective_numeric(
    grid: &Grid,
    addr: &CellAddr,
    visiting: &mut Vec<CellAddr>,
    budget: &mut usize,
) -> Option<f64> {
    let raw = grid.get(addr).unwrap_or("");
    if templated_formula(grid, addr).is_none() && !is_formula(raw) {
        return functions::parse_numeric_or_date_literal(raw);
    }
    match eval_cell(grid, addr, visiting, budget) {
        EvalResult::Number(n) if !n.is_nan() => Some(n),
        EvalResult::Text(s) => functions::parse_numeric_or_date_literal(&s),
        _ => None,
    }
}

/// Evaluate a cell (handles `=...`); used for display and dependencies.
pub fn eval_cell(
    grid: &Grid,
    addr: &CellAddr,
    visiting: &mut Vec<CellAddr>,
    budget: &mut usize,
) -> EvalResult {
    eval_cell_inner(grid, addr, visiting, budget, true)
}

fn eval_cell_with_sheet(
    grid: &Grid,
    sheet_id: u32,
    addr: &CellAddr,
    visiting: &mut Vec<(u32, CellAddr)>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    *budget = budget.saturating_sub(1);
    if *budget == 0 {
        return EvalResult::Error("LIMIT");
    }
    let raw = grid.get(addr).unwrap_or("");
    let t = raw.trim();
    if t.is_empty() {
        return EvalResult::Number(0.0);
    }
    if !t.starts_with('=') {
        return if let Some(n) = parse_number_literal(t) {
            EvalResult::Number(n)
        } else {
            EvalResult::Text(t.to_string())
        };
    }
    if visiting.iter().any(|a| a.0 == sheet_id && &a.1 == addr) {
        return EvalResult::Error("CIRC");
    }
    visiting.push((sheet_id, addr.clone()));
    let r = eval_expr_str(
        &t[1..],
        grid,
        &mut Vec::new(),
        bindings,
        budget,
        allow_templates,
    );
    visiting.pop();
    r
}

fn eval_cell_inner(
    grid: &Grid,
    addr: &CellAddr,
    visiting: &mut Vec<CellAddr>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    *budget = budget.saturating_sub(1);
    if *budget == 0 {
        return EvalResult::Error("LIMIT");
    }

    if allow_templates {
        if let Some(formula) = templated_formula(grid, addr) {
            return eval_expr_str(
                &formula[1..],
                grid,
                visiting,
                &mut Vec::new(),
                budget,
                false,
            );
        }
    }

    let raw = grid.get(addr).unwrap_or("");
    let t = raw.trim();
    if t.is_empty() {
        return EvalResult::Number(0.0);
    }
    if !t.starts_with('=') {
        return if let Some(n) = parse_number_literal(t) {
            EvalResult::Number(n)
        } else {
            EvalResult::Text(t.to_string())
        };
    }

    if let Some(expr) = control_formula_expr(grid, addr) {
        return eval_expr_str(&expr, grid, visiting, &mut Vec::new(), budget, false);
    }

    if visiting.iter().any(|a| a == addr) {
        return EvalResult::Error("CIRC");
    }

    visiting.push(addr.clone());
    let r = eval_expr_str(&t[1..], grid, visiting, &mut Vec::new(), budget, false);
    visiting.pop();
    r
}

fn eval_expr_str(
    expr: &str,
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    let mut p = Parser {
        s: expr.trim(),
        i: 0,
        main_cols: grid.main_cols(),
    };
    let ast = match p.parse_expr() {
        Ok(a) => a,
        Err(()) => return EvalResult::Error("PARSE"),
    };
    p.skip_ws();
    if p.i != p.s.len() {
        return EvalResult::Error("PARSE");
    }
    eval_ast(&ast, grid, visiting, bindings, budget, allow_templates)
}

#[derive(Clone, Debug)]
enum Ast {
    Number(f64),
    Text(String),
    Name(String),
    Ref(CellAddr),
    SheetRef {
        sheet_id: u32,
        addr: CellAddr,
    },
    /// Main grid only (`A1:B2`).
    Range(MainRange),
    Neg(Box<Ast>),
    Add(Box<Ast>, Box<Ast>),
    Sub(Box<Ast>, Box<Ast>),
    Mul(Box<Ast>, Box<Ast>),
    Div(Box<Ast>, Box<Ast>),
    Pow(Box<Ast>, Box<Ast>),
    Call {
        name: String,
        args: Vec<Ast>,
    },
}

struct Parser<'a> {
    s: &'a str,
    i: usize,
    main_cols: usize,
}

impl<'a> Parser<'a> {
    fn peek(&self) -> Option<u8> {
        self.s.as_bytes().get(self.i).copied()
    }

    fn skip_ws(&mut self) {
        while let Some(b) = self.peek() {
            if b.is_ascii_whitespace() {
                self.i += 1;
            } else {
                break;
            }
        }
    }

    fn parse_expr(&mut self) -> Result<Ast, ()> {
        self.parse_add_sub()
    }

    fn parse_add_sub(&mut self) -> Result<Ast, ()> {
        let mut left = self.parse_mul_div()?;
        loop {
            self.skip_ws();
            match self.peek() {
                Some(b'+') => {
                    self.i += 1;
                    let right = self.parse_mul_div()?;
                    left = Ast::Add(Box::new(left), Box::new(right));
                }
                Some(b'-') => {
                    self.i += 1;
                    let right = self.parse_mul_div()?;
                    left = Ast::Sub(Box::new(left), Box::new(right));
                }
                _ => break,
            }
        }
        Ok(left)
    }

    fn parse_mul_div(&mut self) -> Result<Ast, ()> {
        let mut left = self.parse_unary()?;
        loop {
            self.skip_ws();
            match self.peek() {
                Some(b'*') => {
                    self.i += 1;
                    let right = self.parse_unary()?;
                    left = Ast::Mul(Box::new(left), Box::new(right));
                }
                Some(b'/') => {
                    self.i += 1;
                    let right = self.parse_unary()?;
                    left = Ast::Div(Box::new(left), Box::new(right));
                }
                _ => break,
            }
        }
        Ok(left)
    }

    fn parse_unary(&mut self) -> Result<Ast, ()> {
        self.skip_ws();
        if self.peek() == Some(b'-') {
            self.i += 1;
            let inner = self.parse_unary()?;
            return Ok(Ast::Neg(Box::new(inner)));
        }
        self.parse_power()
    }

    fn parse_power(&mut self) -> Result<Ast, ()> {
        let left = self.parse_primary()?;
        self.skip_ws();
        if self.peek() == Some(b'^') {
            self.i += 1;
            let right = self.parse_unary()?;
            return Ok(Ast::Pow(Box::new(left), Box::new(right)));
        }
        Ok(left)
    }

    fn parse_primary(&mut self) -> Result<Ast, ()> {
        self.skip_ws();
        let b = self.peek().ok_or(())?;

        if b == b'"' {
            self.i += 1;
            let mut out = String::new();
            while let Some(ch) = self.peek() {
                self.i += 1;
                if ch == b'"' {
                    if self.peek() == Some(b'"') {
                        out.push('"');
                        self.i += 1;
                    } else {
                        return Ok(Ast::Text(out));
                    }
                } else {
                    out.push(ch as char);
                }
            }
            return Err(());
        }

        if b == b'(' {
            self.i += 1;
            let e = self.parse_expr()?;
            self.skip_ws();
            if self.peek() != Some(b')') {
                return Err(());
            }
            self.i += 1;
            return Ok(e);
        }

        // Number
        if b.is_ascii_digit()
            || (b == b'.'
                && self
                    .s
                    .get(self.i + 1..)
                    .and_then(|r| r.as_bytes().first())
                    .map_or(false, |x| x.is_ascii_digit()))
        {
            return Ok(Ast::Number(self.parse_number()?));
        }

        let rest = &self.s[self.i..];

        if rest.starts_with('#') {
            if let Some((sheet_id, addr, len)) = parse_sheet_qualified_ref(rest) {
                self.i += len;
                return Ok(Ast::SheetRef { sheet_id, addr });
            }
            return Err(());
        }

        if let Some(ch) = self.s[self.i..].chars().next() {
            if ch == 'π' {
                self.i += ch.len_utf8();
                return Ok(Ast::Name(ch.to_string()));
            }
        }

        if rest.starts_with('$') {
            if let Some((sheet_id, addr, len)) =
                parse_sheet_qualified_cell_ref_at_for_workbook(rest)
            {
                self.i += len;
                return Ok(Ast::SheetRef { sheet_id, addr });
            }
            let bytes = rest.as_bytes();
            let mut j = 1usize;
            while j < bytes.len() && (bytes[j].is_ascii_alphanumeric() || bytes[j] == b'_') {
                j += 1;
            }
            if j > 1 && j < bytes.len() && bytes[j] == b':' {
                let sheet = &rest[1..j];
                let after = &rest[j + 1..];
                if let Some((sheet_id, grid)) = workbook_lookup_sheet_ref(sheet) {
                    let (addr, len) = parse_cell_ref_at(after, grid.main_cols()).ok_or(())?;
                    self.i += j + 1 + len;
                    return Ok(Ast::SheetRef { sheet_id, addr });
                }
            }
            return Err(());
        }

        if rest.starts_with('[') || rest.starts_with(']') {
            let (addr, len) = parse_cell_ref_at(rest, 0).ok_or(())?;
            self.i += len;
            return Ok(Ast::Ref(addr));
        }

        // Letter: A1:B2, sum( … ), A1, or bare name.
        if b.is_ascii_alphabetic() || b == b'_' {
            let start = self.i;
            while self
                .peek()
                .map(|x| x.is_ascii_alphanumeric() || x == b'_')
                .unwrap_or(false)
            {
                self.i += 1;
            }
            let token = &self.s[start..self.i];
            let rest = &self.s[start..];
            if let Some((range, len)) = parse_main_range_at(rest) {
                self.i = start + len;
                return Ok(Ast::Range(range));
            }
            if self.peek() == Some(b'(') {
                let name = token.to_string();
                self.i += 1;
                let inner_end = self.scan_balanced_from(self.i)?;
                let inner = &self.s[self.i..inner_end];
                self.i = inner_end + 1;
                let args = if inner.trim().is_empty() {
                    Vec::new()
                } else {
                    split_top_level_args(inner)?
                };
                let mut arg_asts = Vec::with_capacity(args.len());
                for a in args {
                    let mut sub = Parser {
                        s: a.trim(),
                        i: 0,
                        main_cols: self.main_cols,
                    };
                    arg_asts.push(sub.parse_expr()?);
                    sub.skip_ws();
                    if sub.i != sub.s.len() {
                        return Err(());
                    }
                }
                return Ok(Ast::Call {
                    name,
                    args: arg_asts,
                });
            }
            if let Some((addr, len)) = parse_cell_ref_at(rest, self.main_cols) {
                self.i = start + len;
                return Ok(Ast::Ref(addr));
            }
            return Ok(Ast::Name(token.to_string()));
        }

        Err(())
    }

    fn parse_number(&mut self) -> Result<f64, ()> {
        let start = self.i;
        while let Some(b) = self.peek() {
            if b.is_ascii_digit() || b == b'.' {
                self.i += 1;
            } else {
                break;
            }
        }
        self.s[start..self.i].parse::<f64>().map_err(|_| ())
    }

    /// Find index of closing `)` matching an already-open `(` at depth starting at `from`.
    fn scan_balanced_from(&self, from: usize) -> Result<usize, ()> {
        let bytes = self.s.as_bytes();
        let mut depth = 1usize;
        let mut i = from;
        let mut in_string = false;
        while i < bytes.len() {
            match bytes[i] {
                b'"' => {
                    if in_string && i + 1 < bytes.len() && bytes[i + 1] == b'"' {
                        i += 1;
                    } else {
                        in_string = !in_string;
                    }
                }
                b'(' if !in_string => depth += 1,
                b')' if !in_string => {
                    depth -= 1;
                    if depth == 0 {
                        return Ok(i);
                    }
                }
                _ => {}
            }
            i += 1;
        }
        Err(())
    }
}

fn parse_sheet_qualified_ref(s: &str) -> Option<(u32, CellAddr, usize)> {
    let bytes = s.as_bytes();
    if bytes.first().copied()? != b'#' {
        return None;
    }
    let mut i = 1usize;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i == 1 || i >= bytes.len() || bytes[i] != b'!' {
        return None;
    }
    let sheet_id = std::str::from_utf8(&bytes[1..i]).ok()?.parse().ok()?;
    let (addr, len) = parse_cell_ref_at(&s[i + 1..], 0)?;
    Some((sheet_id, addr, i + 1 + len))
}

fn parse_sheet_qualified_cell_ref_at_for_workbook(s: &str) -> Option<(u32, CellAddr, usize)> {
    let bytes = s.as_bytes();
    let mut i = 1usize;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i == 1 || i >= bytes.len() || bytes[i] != b':' {
        return None;
    }
    let sheet_id = std::str::from_utf8(&bytes[1..i]).ok()?.parse().ok()?;
    let grid = workbook_lookup(sheet_id)?;
    let (addr, len) = parse_cell_ref_at(&s[i + 1..], grid.main_cols())?;
    Some((sheet_id, addr, i + 1 + len))
}

fn split_top_level_args(s: &str) -> Result<Vec<&str>, ()> {
    let mut depth = 0i32;
    let mut start = 0usize;
    let mut out = Vec::new();
    let mut in_string = false;
    for (i, c) in s.char_indices() {
        match c {
            '"' => {
                if in_string {
                    in_string = false;
                } else {
                    in_string = true;
                }
            }
            '(' if !in_string => depth += 1,
            ')' if !in_string => depth -= 1,
            ',' if depth == 0 && !in_string => {
                out.push(s[start..i].trim());
                start = i + c.len_utf8();
            }
            _ => {}
        }
    }
    out.push(s[start..].trim());
    if depth != 0 {
        return Err(());
    }
    Ok(out)
}

fn eval_ast(
    ast: &Ast,
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    match ast {
        Ast::Number(n) => EvalResult::Number(*n),
        Ast::Text(s) => EvalResult::Text(s.clone()),
        Ast::Name(name) => resolve_name(name, bindings).unwrap_or(EvalResult::Error("NAME")),
        Ast::Ref(addr) => eval_cell_inner(grid, addr, visiting, budget, allow_templates),
        Ast::SheetRef { sheet_id, addr } => {
            let Some(sheet_grid) = workbook_lookup(*sheet_id) else {
                return EvalResult::Error("SHEET");
            };
            let mut sheet_visiting: Vec<(u32, CellAddr)> = Vec::new();
            let mut sheet_bindings: Vec<(String, EvalResult)> = Vec::new();
            eval_cell_with_sheet(
                &sheet_grid,
                *sheet_id,
                addr,
                &mut sheet_visiting,
                &mut sheet_bindings,
                budget,
                allow_templates,
            )
        }
        Ast::Range(_) => EvalResult::Error("RANGE"),
        Ast::Neg(a) => match eval_ast(a, grid, visiting, bindings, budget, allow_templates) {
            EvalResult::Number(n) => EvalResult::Number(-n),
            e => e,
        },
        Ast::Add(a, b) => eval_binary(
            a,
            b,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
            |x, y| x + y,
        ),
        Ast::Sub(a, b) => eval_binary(
            a,
            b,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
            |x, y| x - y,
        ),
        Ast::Mul(a, b) => eval_binary(
            a,
            b,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
            |x, y| x * y,
        ),
        Ast::Div(a, b) => eval_binary(
            a,
            b,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
            |x, y| {
                if y == 0.0 {
                    f64::NAN
                } else {
                    x / y
                }
            },
        ),
        Ast::Pow(a, b) => eval_binary(
            a,
            b,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
            |x, y| x.powf(y),
        ),
        Ast::Call { name, args } => functions::eval_builtin(
            name,
            args,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        ),
    }
}

fn truthy(e: EvalResult) -> bool {
    match e.scalar_coerce() {
        EvalResult::Number(n) => n != 0.0 && !n.is_nan(),
        EvalResult::Text(s) => functions::parse_numeric_or_date_literal(&s)
            .map(|n| n != 0.0)
            .unwrap_or(false),
        EvalResult::Error(_) => false,
        EvalResult::Array(_) => false,
    }
}

fn eval_binary(
    a: &Ast,
    b: &Ast,
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
    f: fn(f64, f64) -> f64,
) -> EvalResult {
    let (ea, eb) = (
        eval_ast(a, grid, visiting, bindings, budget, allow_templates).scalar_coerce(),
        eval_ast(b, grid, visiting, bindings, budget, allow_templates).scalar_coerce(),
    );
    let na = match ea {
        EvalResult::Number(n) => n,
        EvalResult::Text(s) => {
            if let Some(n) = functions::parse_numeric_or_date_literal(&s) {
                n
            } else {
                return EvalResult::Error("VALUE");
            }
        }
        EvalResult::Error(e) => return EvalResult::Error(e),
        EvalResult::Array(_) => return EvalResult::Error("CALC"),
    };
    let nb = match eb {
        EvalResult::Number(n) => n,
        EvalResult::Text(s) => {
            if let Some(n) = functions::parse_numeric_or_date_literal(&s) {
                n
            } else {
                return EvalResult::Error("VALUE");
            }
        }
        EvalResult::Error(e) => return EvalResult::Error(e),
        EvalResult::Array(_) => return EvalResult::Error("CALC"),
    };
    EvalResult::Number(f(na, nb))
}

fn eval_sum(
    arg: &Ast,
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    match arg {
        Ast::Range(r) => EvalResult::Number(sum_main_range(grid, r, visiting, budget)),
        Ast::Ref(addr) => {
            let n = effective_numeric(grid, addr, visiting, budget);
            EvalResult::Number(n.unwrap_or(0.0))
        }
        Ast::Call { .. }
        | Ast::Neg(_)
        | Ast::Add(_, _)
        | Ast::Sub(_, _)
        | Ast::Mul(_, _)
        | Ast::Div(_, _)
        | Ast::Pow(_, _)
        | Ast::Text(_)
        | Ast::Name(_)
        | Ast::SheetRef { .. } => match eval_ast(
            arg,
            grid,
            visiting,
            &mut Vec::new(),
            budget,
            allow_templates,
        )
        .scalar_coerce()
        {
            EvalResult::Number(n) => EvalResult::Number(n),
            EvalResult::Text(s) => {
                if let Some(n) = parse_number_literal(&s) {
                    EvalResult::Number(n)
                } else {
                    EvalResult::Error("VALUE")
                }
            }
            EvalResult::Error(e) => EvalResult::Error(e),
            EvalResult::Array(_) => EvalResult::Error("CALC"),
        },
        Ast::Number(n) => EvalResult::Number(*n),
    }
}

fn sum_main_range(
    grid: &Grid,
    range: &MainRange,
    visiting: &mut Vec<CellAddr>,
    budget: &mut usize,
) -> f64 {
    if range.is_empty() {
        return 0.0;
    }
    let mut s = 0.0;
    for r in range.row_start..range.row_end {
        for c in range.col_start..range.col_end {
            let addr = CellAddr::Main { row: r, col: c };
            let n = effective_numeric(grid, &addr, visiting, budget).unwrap_or(0.0);
            s += n;
        }
    }
    s
}

pub fn refresh_spills(grid: &mut Grid) {
    let mut prev_followers = grid.spill_followers.clone();
    let mut prev_errors = grid.spill_errors.clone();
    for _ in 0..8 {
        grid.clear_spills();
        let mut anchors: Vec<(CellAddr, String)> = grid
            .main_cells
            .iter()
            .map(|((row, col), value)| {
                (
                    CellAddr::Main {
                        row: *row,
                        col: *col,
                    },
                    value.clone(),
                )
            })
            .collect();
        anchors.sort_by_key(|(addr, _)| match addr {
            CellAddr::Main { row, col } => (*row, *col),
            _ => (u32::MAX, u32::MAX),
        });
        for (addr, raw) in anchors {
            if !is_formula(&raw) {
                continue;
            }
            let mut visiting = Vec::new();
            let mut budget = DEFAULT_BUDGET;
            if let EvalResult::Array(rows) = eval_cell(grid, &addr, &mut visiting, &mut budget) {
                let CellAddr::Main { row: ar, col: ac } = addr else {
                    continue;
                };
                for (r, row) in rows.iter().enumerate() {
                    for (c, cell) in row.iter().enumerate() {
                        if r == 0 && c == 0 {
                            continue;
                        }
                        grid.set_spill_value(
                            CellAddr::Main {
                                row: ar + r as u32,
                                col: ac + c as u32,
                            },
                            eval_result_to_string(cell),
                        );
                    }
                }
            }
        }
        if grid.spill_followers == prev_followers && grid.spill_errors == prev_errors {
            break;
        }
        prev_followers = grid.spill_followers.clone();
        prev_errors = grid.spill_errors.clone();
    }
}

fn eval_result_to_string(result: &EvalResult) -> String {
    match result {
        EvalResult::Number(n) => {
            if n.is_nan() {
                "#NUM!".to_string()
            } else {
                format!("{n}")
            }
        }
        EvalResult::Text(s) => s.clone(),
        EvalResult::Error(e) => format!("#{e}"),
        EvalResult::Array(rows) => rows
            .first()
            .and_then(|row| row.first())
            .map(eval_result_to_string)
            .unwrap_or_else(|| "#CALC".to_string()),
    }
}

/// Display string for a cell: evaluated formula result, or raw text.
pub fn cell_effective_display(grid: &Grid, addr: &CellAddr) -> String {
    if let Some(label) = control_formula_label(grid, addr) {
        return label;
    }
    if let Some(err) = grid.spill_error(addr) {
        return format!("#{err}");
    }
    if let Some(v) = grid.spill_followers.get(addr) {
        return v.clone();
    }
    let raw = grid.get(addr).unwrap_or("");
    if templated_formula(grid, addr).is_none() && !is_formula(raw) {
        return raw.to_string();
    }
    let mut visiting = Vec::new();
    let mut budget = DEFAULT_BUDGET;
    match eval_cell(grid, addr, &mut visiting, &mut budget) {
        EvalResult::Number(n) => {
            if n.is_nan() {
                "#NUM!".to_string()
            } else {
                format!("{n}")
            }
        }
        EvalResult::Text(s) => s,
        EvalResult::Error(e) => format!("#{e}"),
        EvalResult::Array(rows) => rows
            .first()
            .and_then(|row| row.first())
            .map(eval_result_to_string)
            .unwrap_or_else(|| "#CALC".to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::grid::Grid;

    #[test]
    fn formula_add() {
        let mut g = Grid::new(2, 2);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=1+2*3".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 7.0).abs() < 1e-9),
            e => panic!("expected number {:?}", e),
        }
    }

    #[test]
    fn formula_pow() {
        let mut g = Grid::new(1, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=2^3".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "=4^0.5".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 8.0).abs() < 1e-9),
            e => panic!("expected 8 {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 1 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 2.0).abs() < 1e-9),
            e => panic!("expected 2 {:?}", e),
        }
    }

    #[test]
    fn power_is_right_associative() {
        let mut g = Grid::new(1, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=2^3^2".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 512.0).abs() < 1e-9),
            e => panic!("expected 512 {:?}", e),
        }
    }

    #[test]
    fn unary_minus_binds_weaker_than_power() {
        let mut g = Grid::new(1, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=-2^2".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n + 4.0).abs() < 1e-9),
            e => panic!("expected -4 {:?}", e),
        }
    }

    #[test]
    fn sum_range_with_formula_cells() {
        let mut g = Grid::new(2, 2);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "2".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "=A1+3".into());
        g.set(&CellAddr::Main { row: 1, col: 0 }, "=sum(A1:B1)".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 1, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 7.0).abs() < 1e-9),
            e => panic!("expected 7 {:?}", e),
        }
    }

    #[test]
    fn quoted_text_literal_parses() {
        let mut g = Grid::new(1, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=\"hi\"".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "hi"),
            e => panic!("expected text {:?}", e),
        }
    }

    #[test]
    fn quoted_text_escape_parses() {
        let mut g = Grid::new(1, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=\"a\"\"b\"".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "a\"b"),
            e => panic!("expected escaped text {:?}", e),
        }
    }

    #[test]
    fn let_binds_names() {
        let mut g = Grid::new(1, 1);
        g.set(
            &CellAddr::Main { row: 0, col: 0 },
            "=LET(x, 2, y, x + 3, x + y)".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 7.0).abs() < 1e-9),
            e => panic!("expected 7 {:?}", e),
        }
    }

    #[test]
    fn let_supports_shadowing() {
        let mut g = Grid::new(1, 1);
        g.set(
            &CellAddr::Main { row: 0, col: 0 },
            "=LET(x, 1, x, x + 2, x)".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 3.0).abs() < 1e-9),
            e => panic!("expected 3 {:?}", e),
        }
    }

    #[test]
    fn translate_formula_text_rewrites_row_refs() {
        let ctx = FormulaCopyContext {
            source_sheet_id: 1,
            target_sheet_id: 2,
            row_map: vec![1, 0],
            main_cols: 1,
        };
        let out = translate_formula_text("=A1+B2", &ctx).expect("translated formula");
        assert_eq!(out, "=(A2+B1)");
    }

    #[test]
    fn math_constants_evaluate() {
        let mut g = Grid::new(1, 3);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=sin(π)".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "=e".into());
        g.set(&CellAddr::Main { row: 0, col: 2 }, "=c".into());

        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!(n.abs() < 1e-12),
            e => panic!("expected 0 {:?}", e),
        }

        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 1 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - std::f64::consts::E).abs() < 1e-12),
            e => panic!("expected e {:?}", e),
        }

        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 2 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - SPEED_OF_LIGHT_MPS).abs() < 1e-12),
            e => panic!("expected c {:?}", e),
        }
    }

    #[test]
    fn let_can_shadow_pi_constant() {
        let mut g = Grid::new(1, 1);
        g.set(
            &CellAddr::Main { row: 0, col: 0 },
            "=LET(π, 2, π + 1)".into(),
        );

        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 3.0).abs() < 1e-9),
            e => panic!("expected 3 {:?}", e),
        }
    }

    #[test]
    fn bare_name_is_parse_error_outside_let() {
        let mut p = Parser {
            s: "x",
            i: 0,
            main_cols: 0,
        };
        assert!(p.parse_expr().is_ok());
        let mut g = Grid::new(1, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=x".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Error(e) => assert_eq!(e, "NAME"),
            e => panic!("expected NAME {:?}", e),
        }
    }

    #[test]
    fn countif_quoted_text_criteria() {
        let mut g = Grid::new(1, 3);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "b".into());
        g.set(
            &CellAddr::Main { row: 0, col: 2 },
            "=COUNTIF(A1:C1,\"a\")".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 2 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 1.0).abs() < 1e-9),
            e => panic!("expected 1 {:?}", e),
        }
    }

    #[test]
    fn xlookup_exact_match() {
        let mut g = Grid::new(1, 5);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "b".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "a".into());
        g.set(&CellAddr::Main { row: 0, col: 2 }, "20".into());
        g.set(&CellAddr::Main { row: 0, col: 3 }, "30".into());
        g.set(
            &CellAddr::Main { row: 0, col: 4 },
            "=XLOOKUP(\"a\", A1:B1, C1:D1)".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 4 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 30.0).abs() < 1e-9),
            e => panic!("expected 30 {:?}", e),
        }
    }

    #[test]
    fn xlookup_if_not_found() {
        let mut g = Grid::new(1, 4);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "10".into());
        g.set(&CellAddr::Main { row: 0, col: 2 }, "30".into());
        g.set(
            &CellAddr::Main { row: 0, col: 3 },
            "=XLOOKUP(\"z\", A1:A1, B1:B1, \"missing\")".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 3 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "missing"),
            e => panic!("expected missing {:?}", e),
        }
    }

    #[test]
    fn sequence_spills() {
        let mut g = Grid::new(1, 1);
        g.set(
            &CellAddr::Main { row: 0, col: 0 },
            "=SEQUENCE(2,2,1,1)".into(),
        );
        refresh_spills(&mut g);
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 0, col: 0 }),
            "1"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 0, col: 1 }),
            "2"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 1, col: 0 }),
            "3"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 1, col: 1 }),
            "4"
        );
    }

    #[test]
    fn unique_deduplicates() {
        let mut g = Grid::new(1, 4);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "a".into());
        g.set(&CellAddr::Main { row: 0, col: 2 }, "b".into());
        g.set(&CellAddr::Main { row: 0, col: 3 }, "=UNIQUE(A1:C1)".into());
        refresh_spills(&mut g);
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 0, col: 3 }),
            "a"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 1, col: 3 }),
            "b"
        );
    }

    #[test]
    fn filter_applies_mask() {
        let mut g = Grid::new(1, 4);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "b".into());
        g.set(&CellAddr::Main { row: 0, col: 2 }, "1".into());
        g.set(
            &CellAddr::Main { row: 0, col: 3 },
            "=FILTER(A1:B1, C1:D1)".into(),
        );
        refresh_spills(&mut g);
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 0, col: 3 }),
            "a"
        );
    }

    #[test]
    fn iferror_and_ifna() {
        let mut g = Grid::new(1, 4);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "1".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "2".into());
        g.set(
            &CellAddr::Main { row: 0, col: 2 },
            "=IFERROR(VLOOKUP(9,A1:B1,2),\"x\")".into(),
        );
        g.set(
            &CellAddr::Main { row: 0, col: 3 },
            "=IFNA(VLOOKUP(9,A1:B1,2),\"y\")".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 2 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "x"),
            e => panic!("expected x {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 3 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "y"),
            e => panic!("expected y {:?}", e),
        }
    }

    #[test]
    fn index_and_match_work() {
        let mut g = Grid::new(2, 4);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "b".into());
        g.set(&CellAddr::Main { row: 1, col: 0 }, "c".into());
        g.set(&CellAddr::Main { row: 1, col: 1 }, "d".into());
        g.set(
            &CellAddr::Main { row: 0, col: 2 },
            "=MATCH(\"c\",A1:A2,0)".into(),
        );
        g.set(
            &CellAddr::Main { row: 0, col: 3 },
            "=INDEX(A1:B2,2,2)".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 2 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 2.0).abs() < 1e-9),
            e => panic!("expected 2 {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 3 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "d"),
            e => panic!("expected d {:?}", e),
        }
    }

    #[test]
    fn countifs_sumifs_averageifs_work() {
        let mut g = Grid::new(2, 4);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "1".into());
        g.set(&CellAddr::Main { row: 1, col: 0 }, "a".into());
        g.set(&CellAddr::Main { row: 1, col: 1 }, "3".into());
        g.set(
            &CellAddr::Main { row: 0, col: 2 },
            "=COUNTIFS(A1:A2,\"a\")".into(),
        );
        g.set(
            &CellAddr::Main { row: 0, col: 3 },
            "=SUMIFS(B1:B2,A1:A2,\"a\")".into(),
        );
        g.set(
            &CellAddr::Main { row: 1, col: 2 },
            "=AVERAGEIFS(B1:B2,A1:A2,\"a\")".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 2 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 2.0).abs() < 1e-9),
            e => panic!("expected 2 {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 3 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 4.0).abs() < 1e-9),
            e => panic!("expected 4 {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 1, col: 2 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 2.0).abs() < 1e-9),
            e => panic!("expected 2 {:?}", e),
        }
    }

    #[test]
    fn sort_take_drop_choose_work() {
        let mut g = Grid::new(4, 9);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "3".into());
        g.set(&CellAddr::Main { row: 1, col: 0 }, "1".into());
        g.set(&CellAddr::Main { row: 2, col: 0 }, "2".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "b".into());
        g.set(&CellAddr::Main { row: 1, col: 1 }, "d".into());
        g.set(&CellAddr::Main { row: 0, col: 3 }, "=SORT(A1:A3)".into());
        g.set(&CellAddr::Main { row: 0, col: 4 }, "=TAKE(A1:A3,2)".into());
        g.set(&CellAddr::Main { row: 0, col: 5 }, "=DROP(A1:A3,1)".into());
        g.set(
            &CellAddr::Main { row: 0, col: 6 },
            "=CHOOSECOLS(A1:B2,2)".into(),
        );
        g.set(
            &CellAddr::Main { row: 0, col: 7 },
            "=CHOOSEROWS(A1:B2,2)".into(),
        );
        refresh_spills(&mut g);
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 0, col: 3 }),
            "1"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 1, col: 3 }),
            "2"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 2, col: 3 }),
            "3"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 0, col: 4 }),
            "3"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 1, col: 4 }),
            "1"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 0, col: 5 }),
            "1"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 1, col: 5 }),
            "2"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 0, col: 6 }),
            "b"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 1, col: 6 }),
            "d"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 0, col: 7 }),
            "1"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 0, col: 8 }),
            "d"
        );
    }

    #[test]
    fn text_functions_work() {
        let mut g = Grid::new(1, 5);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=LEN(\"abc\")".into());
        g.set(
            &CellAddr::Main { row: 0, col: 1 },
            "=LEFT(\"abcd\",2)".into(),
        );
        g.set(
            &CellAddr::Main { row: 0, col: 2 },
            "=RIGHT(\"abcd\",2)".into(),
        );
        g.set(
            &CellAddr::Main { row: 0, col: 3 },
            "=MID(\"abcd\",2,2)".into(),
        );
        g.set(
            &CellAddr::Main { row: 0, col: 4 },
            "=CONCAT(\"a\",\"b\",\"c\")".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 3.0).abs() < 1e-9),
            e => panic!("expected 3 {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 1 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "ab"),
            e => panic!("expected ab {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 2 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "cd"),
            e => panic!("expected cd {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 3 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "bc"),
            e => panic!("expected bc {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 4 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "abc"),
            e => panic!("expected abc {:?}", e),
        }
    }

    #[test]
    fn text_casing_and_replace_work() {
        let mut g = Grid::new(1, 7);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=UPPER(\"Abc\")".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "=LOWER(\"AbC\")".into());
        g.set(
            &CellAddr::Main { row: 0, col: 2 },
            "=PROPER(\"hello world\")".into(),
        );
        g.set(
            &CellAddr::Main { row: 0, col: 3 },
            "=SUBSTITUTE(\"a-b-a\",\"a\",\"x\")".into(),
        );
        g.set(
            &CellAddr::Main { row: 0, col: 4 },
            "=REPLACE(\"abcdef\",2,3,\"ZZ\")".into(),
        );
        g.set(
            &CellAddr::Main { row: 0, col: 5 },
            "=FIND(\"cd\",\"abcdef\")".into(),
        );
        g.set(
            &CellAddr::Main { row: 0, col: 6 },
            "=SEARCH(\"CD\",\"abCdEf\")".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "ABC"),
            e => panic!("expected ABC {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 1 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "abc"),
            e => panic!("expected abc {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 2 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "Hello World"),
            e => panic!("expected Hello World {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 3 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "x-b-x"),
            e => panic!("expected x-b-x {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 4 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "aZZef"),
            e => panic!("expected aZZef {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 5 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 3.0).abs() < 1e-9),
            e => panic!("expected 3 {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 6 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 3.0).abs() < 1e-9),
            e => panic!("expected 3 {:?}", e),
        }
    }

    #[test]
    fn text_formatting_works() {
        let mut g = Grid::new(1, 1);
        g.set(
            &CellAddr::Main { row: 0, col: 0 },
            "=TEXT(12.345,\"0.00\")".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "12.35"),
            e => panic!("expected 12.35 {:?}", e),
        }
    }

    #[test]
    fn date_time_functions_work() {
        let mut g = Grid::new(1, 7);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "2024-01-02".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "=A1+1".into());
        g.set(&CellAddr::Main { row: 0, col: 2 }, "=YEAR(A1)".into());
        g.set(&CellAddr::Main { row: 0, col: 3 }, "=MONTH(A1)".into());
        g.set(&CellAddr::Main { row: 0, col: 4 }, "=DAY(A1)".into());
        g.set(&CellAddr::Main { row: 0, col: 5 }, "=HOUR(NOW())".into());
        g.set(&CellAddr::Main { row: 0, col: 6 }, "=MINUTE(NOW())".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 1 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!(n > 45000.0),
            e => panic!("expected date arithmetic {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 2 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 2024.0).abs() < 1e-9),
            e => panic!("expected year {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 3 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 1.0).abs() < 1e-9),
            e => panic!("expected month {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 4 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 2.0).abs() < 1e-9),
            e => panic!("expected day {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 5 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((0.0..24.0).contains(&n)),
            e => panic!("expected hour {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 6 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((0.0..60.0).contains(&n)),
            e => panic!("expected minute {:?}", e),
        }
    }

    #[test]
    fn rand_is_deterministic_per_seed() {
        let mut g = Grid::new(1, 2);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=RAND()".into());
        g.set(
            &CellAddr::Main { row: 0, col: 1 },
            "=RANDBETWEEN(1,10)".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        let first = match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => n,
            e => panic!("expected rand {:?}", e),
        };
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        let second = match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => n,
            e => panic!("expected rand {:?}", e),
        };
        assert!((first - second).abs() < 1e-12);
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        let between = match eval_cell(&g, &CellAddr::Main { row: 0, col: 1 }, &mut v, &mut b) {
            EvalResult::Number(n) => n,
            e => panic!("expected randbetween {:?}", e),
        };
        assert!((1.0..=10.0).contains(&between));
        g.bump_volatile_seed();
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        let changed = match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => n,
            e => panic!("expected rand {:?}", e),
        };
        assert!((first - changed).abs() > 1e-12);
    }

    #[test]
    fn practical_batch_functions_work() {
        let mut g = Grid::new(3, 8);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "".into());
        g.set(&CellAddr::Main { row: 1, col: 0 }, "1".into());
        g.set(&CellAddr::Main { row: 1, col: 1 }, "text".into());
        g.set(
            &CellAddr::Main { row: 2, col: 0 },
            "=COUNTBLANK(A1:B2)".into(),
        );
        g.set(&CellAddr::Main { row: 0, col: 2 }, "=ISNUMBER(A2)".into());
        g.set(&CellAddr::Main { row: 0, col: 3 }, "=ISTEXT(B2)".into());
        g.set(&CellAddr::Main { row: 0, col: 4 }, "=ISBLANK(B1)".into());
        g.set(
            &CellAddr::Main { row: 0, col: 5 },
            "=ISERROR(XMATCH(\"z\",A1:A2))".into(),
        );
        g.set(
            &CellAddr::Main { row: 0, col: 6 },
            "=SWITCH(2,1,\"one\",2,\"two\",\"default\")".into(),
        );
        g.set(
            &CellAddr::Main { row: 0, col: 7 },
            "=CHOOSE(2,\"x\",\"y\",\"z\")".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 2, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 1.0).abs() < 1e-9),
            e => panic!("expected 1 {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 2 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 1.0).abs() < 1e-9),
            e => panic!("expected 1 {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 3 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 1.0).abs() < 1e-9),
            e => panic!("expected 1 {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 4 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 1.0).abs() < 1e-9),
            e => panic!("expected 1 {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 5 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 1.0).abs() < 1e-9),
            e => panic!("expected 1 {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 6 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "two"),
            e => panic!("expected two {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 7 }, &mut v, &mut b) {
            EvalResult::Text(s) => assert_eq!(s, "y"),
            e => panic!("expected y {:?}", e),
        }
    }

    #[test]
    fn sortby_spills_sorted_rows() {
        let mut g = Grid::new(3, 6);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "b".into());
        g.set(&CellAddr::Main { row: 1, col: 0 }, "a".into());
        g.set(&CellAddr::Main { row: 2, col: 0 }, "c".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "2".into());
        g.set(&CellAddr::Main { row: 1, col: 1 }, "1".into());
        g.set(&CellAddr::Main { row: 2, col: 1 }, "3".into());
        g.set(
            &CellAddr::Main { row: 0, col: 2 },
            "=SORTBY(A1:B3,B1:B3,1)".into(),
        );
        refresh_spills(&mut g);
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 0, col: 2 }),
            "a"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 1, col: 2 }),
            "b"
        );
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 2, col: 2 }),
            "c"
        );
    }

    #[test]
    fn sumproduct_works() {
        let mut g = Grid::new(2, 3);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "1".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "2".into());
        g.set(&CellAddr::Main { row: 1, col: 0 }, "3".into());
        g.set(&CellAddr::Main { row: 1, col: 1 }, "4".into());
        g.set(
            &CellAddr::Main { row: 0, col: 2 },
            "=SUMPRODUCT(A1:B2)".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 2 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 10.0).abs() < 1e-9),
            e => panic!("expected 10 {:?}", e),
        }
    }

    #[test]
    fn ifs_works() {
        let mut g = Grid::new(1, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=IFS(0,1,1,2)".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 2.0).abs() < 1e-9),
            e => panic!("expected 2 {:?}", e),
        }
    }

    #[test]
    fn xmatch_works() {
        let mut g = Grid::new(1, 4);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "b".into());
        g.set(&CellAddr::Main { row: 0, col: 2 }, "c".into());
        g.set(
            &CellAddr::Main { row: 0, col: 3 },
            "=XMATCH(\"b\",A1:C1)".into(),
        );
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 3 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 2.0).abs() < 1e-9),
            e => panic!("expected 2 {:?}", e),
        }
    }

    #[test]
    fn boolean_functions_work() {
        let mut g = Grid::new(1, 3);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=AND(1,2,3)".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "=OR(0,0,1)".into());
        g.set(&CellAddr::Main { row: 0, col: 2 }, "=NOT(0)".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 1.0).abs() < 1e-9),
            e => panic!("expected 1 {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 1 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 1.0).abs() < 1e-9),
            e => panic!("expected 1 {:?}", e),
        }
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 2 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 1.0).abs() < 1e-9),
            e => panic!("expected 1 {:?}", e),
        }
    }

    #[test]
    fn sheet_ref_syntax_parses() {
        let mut p = Parser {
            s: "#2!A1",
            i: 0,
            main_cols: 0,
        };
        match p.parse_expr().unwrap() {
            Ast::SheetRef { sheet_id, addr } => {
                assert_eq!(sheet_id, 2);
                assert_eq!(addr, CellAddr::Main { row: 0, col: 0 });
            }
            other => panic!("unexpected ast: {other:?}"),
        }
    }

    #[test]
    fn named_sheet_ref_syntax_parses() {
        let mut wb = WorkbookState::new();
        wb.add_sheet("Sheet2".into(), crate::ops::SheetState::new(1, 1));
        let _guard = set_eval_context(&wb);
        let mut p = Parser {
            s: "$Sheet2:A1",
            i: 0,
            main_cols: 0,
        };
        match p.parse_expr().unwrap() {
            Ast::SheetRef { sheet_id, addr } => {
                assert_eq!(sheet_id, 2);
                assert_eq!(addr, CellAddr::Main { row: 0, col: 0 });
            }
            other => panic!("unexpected ast: {other:?}"),
        }
    }

    #[test]
    fn named_sheet_title_ref_parses() {
        let mut wb = WorkbookState::new();
        wb.add_sheet("Budget".into(), crate::ops::SheetState::new(1, 1));
        let _guard = set_eval_context(&wb);
        let mut p = Parser {
            s: "$Budget:A1",
            i: 0,
            main_cols: 0,
        };
        match p.parse_expr().unwrap() {
            Ast::SheetRef { sheet_id, addr } => {
                assert_eq!(sheet_id, 2);
                assert_eq!(addr, CellAddr::Main { row: 0, col: 0 });
            }
            other => panic!("unexpected ast: {other:?}"),
        }
    }

    #[test]
    fn circular_ref() {
        let mut g = Grid::new(1, 2);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "=B1".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "=A1".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Error(e) => assert_eq!(e, "CIRC"),
            e => panic!("expected CIRC {:?}", e),
        }
    }

    #[test]
    fn if_func() {
        let mut g = Grid::new(1, 3);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "0".into());
        g.set(&CellAddr::Main { row: 0, col: 2 }, "=IF(A1,1,2)".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 2 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 2.0).abs() < 1e-9),
            e => panic!("expected 2 {:?}", e),
        }
    }

    #[test]
    fn header_template_label_is_display_only() {
        let mut g = Grid::new(2, 2);
        g.set(
            &CellAddr::Header {
                row: 25,
                col: MARGIN_COLS as u32 + 1,
            },
            "=A*2 -- POW2".into(),
        );
        g.set(&CellAddr::Main { row: 0, col: 0 }, "7".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Main { row: 0, col: 1 }),
            "14"
        );
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 1 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 14.0).abs() < 1e-9),
            e => panic!("expected 14 {:?}", e),
        }
        assert_eq!(
            cell_effective_display(
                &g,
                &CellAddr::Header {
                    row: 25,
                    col: MARGIN_COLS as u32 + 1,
                },
            ),
            "POW2"
        );
    }

    #[test]
    fn left_margin_template_can_label_rows() {
        let mut g = Grid::new(2, 2);
        g.set(&CellAddr::Left { col: 9, row: 0 }, "=:1*0.1 -- TAX".into());
        g.set(&CellAddr::Main { row: 0, col: 0 }, "10".into());
        let mut v = Vec::new();
        let mut b = DEFAULT_BUDGET;
        match eval_cell(&g, &CellAddr::Main { row: 0, col: 0 }, &mut v, &mut b) {
            EvalResult::Number(n) => assert!((n - 1.0).abs() < 1e-9),
            e => panic!("expected 1 {:?}", e),
        }
        assert_eq!(
            cell_effective_display(&g, &CellAddr::Left { col: 9, row: 0 }),
            "TAX"
        );
    }
}
