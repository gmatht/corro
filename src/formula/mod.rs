//! `=...` cell formulas: parse, evaluate, display.

use crate::addr::{excel_column_name, parse_cell_ref_at, parse_main_range_at};
use crate::grid::{CellAddr, Grid, MainRange, HEADER_ROWS, MARGIN_COLS};

const DEFAULT_BUDGET: usize = 10_000;

/// Evaluation step budget for one aggregate range scan (many cells).
pub const EVAL_BUDGET_AGG: usize = 1_000_000;

/// Result of evaluating a cell (formula or plain).
#[derive(Clone, Debug, PartialEq)]
pub enum EvalResult {
    Number(f64),
    Text(String),
    /// Display as `#msg` in the UI.
    Error(&'static str),
}

fn parse_number_literal(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    t.parse::<f64>().ok()
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
        return parse_number_literal(raw);
    }
    match eval_cell(grid, addr, visiting, budget) {
        EvalResult::Number(n) if !n.is_nan() => Some(n),
        EvalResult::Text(s) => parse_number_literal(&s),
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
            return eval_expr_str(&formula[1..], grid, visiting, budget, false);
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
        return eval_expr_str(&expr, grid, visiting, budget, false);
    }

    if visiting.iter().any(|a| a == addr) {
        return EvalResult::Error("CIRC");
    }

    visiting.push(addr.clone());
    let r = eval_expr_str(&t[1..], grid, visiting, budget, false);
    visiting.pop();
    r
}

fn eval_expr_str(
    expr: &str,
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    let mut p = Parser {
        s: expr.trim(),
        i: 0,
    };
    let ast = match p.parse_expr() {
        Ok(a) => a,
        Err(()) => return EvalResult::Error("PARSE"),
    };
    p.skip_ws();
    if p.i != p.s.len() {
        return EvalResult::Error("PARSE");
    }
    eval_ast(&ast, grid, visiting, budget, allow_templates)
}

#[derive(Clone, Debug)]
enum Ast {
    Number(f64),
    Ref(CellAddr),
    /// Main grid only (`A1:B2`).
    Range(MainRange),
    Neg(Box<Ast>),
    Add(Box<Ast>, Box<Ast>),
    Sub(Box<Ast>, Box<Ast>),
    Mul(Box<Ast>, Box<Ast>),
    Div(Box<Ast>, Box<Ast>),
    Call {
        name: String,
        args: Vec<Ast>,
    },
}

struct Parser<'a> {
    s: &'a str,
    i: usize,
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
        self.parse_primary()
    }

    fn parse_primary(&mut self) -> Result<Ast, ()> {
        self.skip_ws();
        let b = self.peek().ok_or(())?;

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

        // Region-style refs
        if rest.starts_with('~')
            || rest.starts_with('_')
            || rest.starts_with('<')
            || rest.starts_with('>')
        {
            let (addr, len) = parse_cell_ref_at(rest).ok_or(())?;
            self.i += len;
            return Ok(Ast::Ref(addr));
        }

        // Letter: A1:B2, sum( … ), or A1
        if b.is_ascii_alphabetic() {
            let rest = &self.s[self.i..];
            if let Some((range, len)) = parse_main_range_at(rest) {
                self.i += len;
                return Ok(Ast::Range(range));
            }
            let start = self.i;
            while self
                .peek()
                .map(|x| x.is_ascii_alphabetic())
                .unwrap_or(false)
            {
                self.i += 1;
            }
            let letters = &self.s[start..self.i];
            if self.peek() == Some(b'(') {
                let name = letters.to_string();
                self.i += 1;
                let inner_end = self.scan_balanced_from(self.i)?;
                let inner = &self.s[self.i..inner_end];
                self.i = inner_end + 1;
                let args = split_top_level_args(inner)?;
                let mut arg_asts = Vec::with_capacity(args.len());
                for a in args {
                    let mut sub = Parser { s: a.trim(), i: 0 };
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
            self.i = start;
            let (addr, len) = parse_cell_ref_at(&self.s[self.i..]).ok_or(())?;
            self.i += len;
            return Ok(Ast::Ref(addr));
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
        while i < bytes.len() {
            match bytes[i] {
                b'(' => depth += 1,
                b')' => {
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

fn split_top_level_args(s: &str) -> Result<Vec<&str>, ()> {
    let mut depth = 0i32;
    let mut start = 0usize;
    let mut out = Vec::new();
    for (i, c) in s.char_indices() {
        match c {
            '(' => depth += 1,
            ')' => depth -= 1,
            ',' if depth == 0 => {
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
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    match ast {
        Ast::Number(n) => EvalResult::Number(*n),
        Ast::Ref(addr) => eval_cell_inner(grid, addr, visiting, budget, allow_templates),
        Ast::Range(_) => EvalResult::Error("RANGE"),
        Ast::Neg(a) => match eval_ast(a, grid, visiting, budget, allow_templates) {
            EvalResult::Number(n) => EvalResult::Number(-n),
            e => e,
        },
        Ast::Add(a, b) => eval_binary(a, b, grid, visiting, budget, allow_templates, |x, y| x + y),
        Ast::Sub(a, b) => eval_binary(a, b, grid, visiting, budget, allow_templates, |x, y| x - y),
        Ast::Mul(a, b) => eval_binary(a, b, grid, visiting, budget, allow_templates, |x, y| x * y),
        Ast::Div(a, b) => eval_binary(a, b, grid, visiting, budget, allow_templates, |x, y| {
            if y == 0.0 {
                f64::NAN
            } else {
                x / y
            }
        }),
        Ast::Call { name, args } => {
            let u = name.to_ascii_uppercase();
            match u.as_str() {
                "SUM" => {
                    if args.len() != 1 {
                        return EvalResult::Error("ARGS");
                    }
                    eval_sum(&args[0], grid, visiting, budget, allow_templates)
                }
                "IF" => {
                    if args.len() != 3 {
                        return EvalResult::Error("ARGS");
                    }
                    let cond = eval_ast(&args[0], grid, visiting, budget, allow_templates);
                    let pick = truthy(cond);
                    if pick {
                        eval_ast(&args[1], grid, visiting, budget, allow_templates)
                    } else {
                        eval_ast(&args[2], grid, visiting, budget, allow_templates)
                    }
                }
                _ => EvalResult::Error("FUNC"),
            }
        }
    }
}

fn truthy(e: EvalResult) -> bool {
    match e {
        EvalResult::Number(n) => n != 0.0 && !n.is_nan(),
        EvalResult::Text(s) => parse_number_literal(&s).map(|n| n != 0.0).unwrap_or(false),
        EvalResult::Error(_) => false,
    }
}

fn eval_binary(
    a: &Ast,
    b: &Ast,
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    budget: &mut usize,
    allow_templates: bool,
    f: fn(f64, f64) -> f64,
) -> EvalResult {
    let (ea, eb) = (
        eval_ast(a, grid, visiting, budget, allow_templates),
        eval_ast(b, grid, visiting, budget, allow_templates),
    );
    let na = match ea {
        EvalResult::Number(n) => n,
        EvalResult::Text(s) => {
            if let Some(n) = parse_number_literal(&s) {
                n
            } else {
                return EvalResult::Error("VALUE");
            }
        }
        EvalResult::Error(e) => return EvalResult::Error(e),
    };
    let nb = match eb {
        EvalResult::Number(n) => n,
        EvalResult::Text(s) => {
            if let Some(n) = parse_number_literal(&s) {
                n
            } else {
                return EvalResult::Error("VALUE");
            }
        }
        EvalResult::Error(e) => return EvalResult::Error(e),
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
        | Ast::Div(_, _) => match eval_ast(arg, grid, visiting, budget, allow_templates) {
            EvalResult::Number(n) => EvalResult::Number(n),
            EvalResult::Text(s) => {
                if let Some(n) = parse_number_literal(&s) {
                    EvalResult::Number(n)
                } else {
                    EvalResult::Error("VALUE")
                }
            }
            EvalResult::Error(e) => EvalResult::Error(e),
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

/// Display string for a cell: evaluated formula result, or raw text.
pub fn cell_effective_display(grid: &Grid, addr: &CellAddr) -> String {
    if let Some(label) = control_formula_label(grid, addr) {
        return label;
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
