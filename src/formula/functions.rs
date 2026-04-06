use super::{
    eval_ast, eval_binary, eval_cell_inner, eval_sum, parse_number_literal, Ast, EvalResult,
};
use crate::grid::{CellAddr, Grid, MainRange};

pub(crate) fn eval_builtin(
    name: &str,
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    let u = name.to_ascii_uppercase();
    match u.as_str() {
        "SUM" => {
            if args.len() != 1 {
                return EvalResult::Error("ARGS");
            }
            eval_sum(&args[0], grid, visiting, budget, allow_templates)
        }
        "AVERAGE" => eval_numeric_aggregate(
            &args,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
            NumericAgg::Average,
        ),
        "MIN" => eval_numeric_aggregate(
            &args,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
            NumericAgg::Min,
        ),
        "MAX" => eval_numeric_aggregate(
            &args,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
            NumericAgg::Max,
        ),
        "COUNT" => eval_count(&args, grid, visiting, bindings, budget, allow_templates),
        "COUNTA" => eval_counta(&args, grid, visiting, bindings, budget, allow_templates),
        "PRODUCT" => eval_numeric_aggregate(
            &args,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
            NumericAgg::Product,
        ),
        "ABS" => eval_unary_numeric(
            &args,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
            f64::abs,
        ),
        "ROUND" => eval_round(&args, grid, visiting, bindings, budget, allow_templates),
        "MOD" => eval_binary_numeric(
            &args,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
            |x, y| x % y,
        ),
        "POWER" => {
            if args.len() != 2 {
                EvalResult::Error("ARGS")
            } else {
                eval_binary(
                    &args[0],
                    &args[1],
                    grid,
                    visiting,
                    bindings,
                    budget,
                    allow_templates,
                    |x, y| x.powf(y),
                )
            }
        }
        "SQRT" => eval_unary_numeric(
            &args,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
            f64::sqrt,
        ),
        "PI" => {
            if !args.is_empty() {
                EvalResult::Error("ARGS")
            } else {
                EvalResult::Number(std::f64::consts::PI)
            }
        }
        "SIN" => eval_unary_numeric(
            &args,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
            f64::sin,
        ),
        "COS" => eval_unary_numeric(
            &args,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
            f64::cos,
        ),
        "TRIM" => eval_trim(&args, grid, visiting, bindings, budget, allow_templates),
        "COUNTIF" => eval_countif(&args, grid, visiting, bindings, budget, allow_templates),
        "SUMIF" => eval_sumif(&args, grid, visiting, bindings, budget, allow_templates),
        "LOOKUP" => eval_lookup(&args, grid, visiting, bindings, budget, allow_templates),
        "VLOOKUP" => eval_vlookup(&args, grid, visiting, bindings, budget, allow_templates),
        "XLOOKUP" => eval_xlookup(&args, grid, visiting, bindings, budget, allow_templates),
        "LET" => eval_let(&args, grid, visiting, bindings, budget, allow_templates),
        "SEQUENCE" => eval_sequence(&args, grid, visiting, bindings, budget, allow_templates),
        "FILTER" => eval_filter(&args, grid, visiting, bindings, budget, allow_templates),
        "UNIQUE" => eval_unique(&args, grid, visiting, bindings, budget, allow_templates),
        "IF" => {
            if args.len() != 3 {
                return EvalResult::Error("ARGS");
            }
            let cond = eval_ast(&args[0], grid, visiting, bindings, budget, allow_templates);
            let pick = super::truthy(cond);
            if pick {
                eval_ast(&args[1], grid, visiting, bindings, budget, allow_templates)
            } else {
                eval_ast(&args[2], grid, visiting, bindings, budget, allow_templates)
            }
        }
        _ => EvalResult::Error("FUNC"),
    }
}

enum NumericAgg {
    Average,
    Min,
    Max,
    Product,
}

fn eval_numeric_aggregate(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
    kind: NumericAgg,
) -> EvalResult {
    if args.len() != 1 {
        return EvalResult::Error("ARGS");
    }
    let nums =
        match collect_numeric_values(&args[0], grid, visiting, bindings, budget, allow_templates) {
            Ok(nums) => nums,
            Err(e) => return EvalResult::Error(e),
        };
    match kind {
        NumericAgg::Average => {
            if nums.is_empty() {
                EvalResult::Error("DIV0")
            } else {
                EvalResult::Number(nums.iter().sum::<f64>() / nums.len() as f64)
            }
        }
        NumericAgg::Min => nums
            .into_iter()
            .min_by(|a, b| a.partial_cmp(b).unwrap())
            .map(EvalResult::Number)
            .unwrap_or(EvalResult::Number(0.0)),
        NumericAgg::Max => nums
            .into_iter()
            .max_by(|a, b| a.partial_cmp(b).unwrap())
            .map(EvalResult::Number)
            .unwrap_or(EvalResult::Number(0.0)),
        NumericAgg::Product => EvalResult::Number(nums.into_iter().fold(1.0, |acc, n| acc * n)),
    }
}

fn eval_unary_numeric(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
    f: fn(f64) -> f64,
) -> EvalResult {
    if args.len() != 1 {
        return EvalResult::Error("ARGS");
    }
    match numeric_value(eval_ast(
        &args[0],
        grid,
        visiting,
        bindings,
        budget,
        allow_templates,
    )) {
        Some(n) => EvalResult::Number(f(n)),
        None => EvalResult::Error("VALUE"),
    }
}

fn eval_binary_numeric(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
    f: fn(f64, f64) -> f64,
) -> EvalResult {
    if args.len() != 2 {
        return EvalResult::Error("ARGS");
    }
    eval_binary(
        &args[0],
        &args[1],
        grid,
        visiting,
        bindings,
        budget,
        allow_templates,
        f,
    )
}

fn eval_round(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 2 {
        return EvalResult::Error("ARGS");
    }
    let n = match numeric_value(eval_ast(
        &args[0],
        grid,
        visiting,
        bindings,
        budget,
        allow_templates,
    )) {
        Some(n) => n,
        None => return EvalResult::Error("VALUE"),
    };
    let digits = match numeric_value(eval_ast(
        &args[1],
        grid,
        visiting,
        bindings,
        budget,
        allow_templates,
    )) {
        Some(n) => n,
        None => return EvalResult::Error("VALUE"),
    };
    let factor = 10f64.powf(digits);
    EvalResult::Number((n * factor).round() / factor)
}

fn eval_trim(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 1 {
        return EvalResult::Error("ARGS");
    }
    match eval_ast(&args[0], grid, visiting, bindings, budget, allow_templates).scalar_coerce() {
        EvalResult::Number(n) => EvalResult::Text(trim_spaces(&format!("{n}"))),
        EvalResult::Text(s) => EvalResult::Text(trim_spaces(&s)),
        EvalResult::Error(e) => EvalResult::Error(e),
        EvalResult::Array(_) => EvalResult::Error("CALC"),
    }
}

fn eval_count(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 1 {
        return EvalResult::Error("ARGS");
    }
    match count_numeric_values(&args[0], grid, visiting, bindings, budget, allow_templates) {
        Ok(n) => EvalResult::Number(n as f64),
        Err(e) => EvalResult::Error(e),
    }
}

fn eval_counta(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 1 {
        return EvalResult::Error("ARGS");
    }
    match count_nonempty_values(&args[0], grid, visiting, bindings, budget, allow_templates) {
        Ok(n) => EvalResult::Number(n as f64),
        Err(e) => EvalResult::Error(e),
    }
}

fn eval_countif(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 2 {
        return EvalResult::Error("ARGS");
    }
    let Some(range) = as_main_range(&args[0]) else {
        return EvalResult::Error("RANGE");
    };
    let Ok(criteria) =
        criteria_from_ast(&args[1], grid, visiting, bindings, budget, allow_templates)
    else {
        return EvalResult::Error("VALUE");
    };
    let mut count = 0usize;
    for r in range.row_start..range.row_end {
        for c in range.col_start..range.col_end {
            let addr = CellAddr::Main { row: r, col: c };
            if criteria_matches(&criteria, grid, &addr, visiting, budget, allow_templates) {
                count += 1;
            }
        }
    }
    EvalResult::Number(count as f64)
}

fn eval_sumif(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 2 && args.len() != 3 {
        return EvalResult::Error("ARGS");
    }
    let Some(criteria_range) = as_main_range(&args[0]) else {
        return EvalResult::Error("RANGE");
    };
    let sum_range = if args.len() == 3 {
        match as_main_range(&args[2]) {
            Some(r) => r,
            None => return EvalResult::Error("RANGE"),
        }
    } else {
        criteria_range.clone()
    };
    let Ok(criteria) =
        criteria_from_ast(&args[1], grid, visiting, bindings, budget, allow_templates)
    else {
        return EvalResult::Error("VALUE");
    };
    let criteria_rows = criteria_range.row_end - criteria_range.row_start;
    let criteria_cols = criteria_range.col_end - criteria_range.col_start;
    let sum_rows = sum_range.row_end - sum_range.row_start;
    let sum_cols = sum_range.col_end - sum_range.col_start;
    if criteria_rows != sum_rows || criteria_cols != sum_cols {
        return EvalResult::Error("ARGS");
    }
    let mut sum = 0.0;
    for dr in 0..criteria_rows {
        for dc in 0..criteria_cols {
            let crit_addr = CellAddr::Main {
                row: criteria_range.row_start + dr,
                col: criteria_range.col_start + dc,
            };
            if criteria_matches(
                &criteria,
                grid,
                &crit_addr,
                visiting,
                budget,
                allow_templates,
            ) {
                let sum_addr = CellAddr::Main {
                    row: sum_range.row_start + dr,
                    col: sum_range.col_start + dc,
                };
                if let Some(n) = super::effective_numeric(grid, &sum_addr, visiting, budget) {
                    sum += n;
                }
            }
        }
    }
    EvalResult::Number(sum)
}

fn collect_numeric_values(
    arg: &Ast,
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> Result<Vec<f64>, &'static str> {
    match arg {
        Ast::Range(r) => {
            let mut out = Vec::new();
            for row in r.row_start..r.row_end {
                for col in r.col_start..r.col_end {
                    let addr = CellAddr::Main { row, col };
                    if let Some(n) = super::effective_numeric(grid, &addr, visiting, budget) {
                        out.push(n);
                    }
                }
            }
            Ok(out)
        }
        Ast::Ref(addr) => Ok(super::effective_numeric(grid, addr, visiting, budget)
            .into_iter()
            .collect()),
        _ => match eval_ast(arg, grid, visiting, bindings, budget, allow_templates).scalar_coerce()
        {
            EvalResult::Number(n) => Ok(vec![n]),
            EvalResult::Text(s) => Ok(parse_number_literal(&s).into_iter().collect()),
            EvalResult::Error(e) => Err(e),
            EvalResult::Array(_) => Err("CALC"),
        },
    }
}

fn trim_spaces(s: &str) -> String {
    s.split_whitespace().collect::<Vec<_>>().join(" ")
}

fn numeric_value(result: EvalResult) -> Option<f64> {
    match result {
        EvalResult::Number(n) => Some(n),
        EvalResult::Text(s) => parse_number_literal(&s),
        EvalResult::Error(_) => None,
        EvalResult::Array(_) => None,
    }
}

fn as_main_range(ast: &Ast) -> Option<MainRange> {
    match ast {
        Ast::Range(r) => Some(r.clone()),
        _ => None,
    }
}

fn count_numeric_values(
    arg: &Ast,
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> Result<usize, &'static str> {
    match arg {
        Ast::Range(r) => {
            let mut n = 0usize;
            for row in r.row_start..r.row_end {
                for col in r.col_start..r.col_end {
                    let addr = CellAddr::Main { row, col };
                    if super::effective_numeric(grid, &addr, visiting, budget).is_some() {
                        n += 1;
                    }
                }
            }
            Ok(n)
        }
        Ast::Ref(addr) => Ok(
            if super::effective_numeric(grid, addr, visiting, budget).is_some() {
                1
            } else {
                0
            },
        ),
        _ => match eval_ast(arg, grid, visiting, bindings, budget, allow_templates).scalar_coerce()
        {
            EvalResult::Number(n) => Ok((!n.is_nan()) as usize),
            EvalResult::Text(s) => Ok(parse_number_literal(&s).is_some() as usize),
            EvalResult::Error(e) => Err(e),
            EvalResult::Array(_) => Err("CALC"),
        },
    }
}

fn count_nonempty_values(
    arg: &Ast,
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> Result<usize, &'static str> {
    match arg {
        Ast::Range(r) => {
            let mut n = 0usize;
            for row in r.row_start..r.row_end {
                for col in r.col_start..r.col_end {
                    let addr = CellAddr::Main { row, col };
                    if grid.get(&addr).map(|s| !s.is_empty()).unwrap_or(false) {
                        n += 1;
                    }
                }
            }
            Ok(n)
        }
        Ast::Ref(addr) => Ok(grid.get(addr).map(|s| !s.is_empty()).unwrap_or(false) as usize),
        _ => match eval_ast(arg, grid, visiting, bindings, budget, allow_templates).scalar_coerce()
        {
            EvalResult::Number(_) => Ok(1),
            EvalResult::Text(s) => Ok((!s.is_empty()) as usize),
            EvalResult::Error(e) => Err(e),
            EvalResult::Array(_) => Err("CALC"),
        },
    }
}

#[derive(Clone, Copy)]
enum CriteriaOp {
    Eq,
    Ne,
    Gt,
    Ge,
    Lt,
    Le,
}

struct Criteria {
    op: CriteriaOp,
    value: String,
    numeric: Option<f64>,
}

fn compare_f64(op: CriteriaOp, left: f64, right: f64) -> bool {
    match op {
        CriteriaOp::Eq => left == right,
        CriteriaOp::Ne => left != right,
        CriteriaOp::Gt => left > right,
        CriteriaOp::Ge => left >= right,
        CriteriaOp::Lt => left < right,
        CriteriaOp::Le => left <= right,
    }
}

fn compare_str(op: CriteriaOp, left: &str, right: &str) -> bool {
    match op {
        CriteriaOp::Eq => left == right,
        CriteriaOp::Ne => left != right,
        CriteriaOp::Gt => left > right,
        CriteriaOp::Ge => left >= right,
        CriteriaOp::Lt => left < right,
        CriteriaOp::Le => left <= right,
    }
}

fn criteria_from_ast(
    ast: &Ast,
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> Result<Criteria, &'static str> {
    let raw = match eval_ast(ast, grid, visiting, bindings, budget, allow_templates).scalar_coerce()
    {
        EvalResult::Number(n) => n.to_string(),
        EvalResult::Text(s) => s,
        EvalResult::Error(e) => return Err(e),
        EvalResult::Array(_) => return Err("CALC"),
    };
    let s = raw.trim();
    let (op, rest) = if let Some(rest) = s.strip_prefix(">=") {
        (CriteriaOp::Ge, rest)
    } else if let Some(rest) = s.strip_prefix("<=") {
        (CriteriaOp::Le, rest)
    } else if let Some(rest) = s.strip_prefix("<>") {
        (CriteriaOp::Ne, rest)
    } else if let Some(rest) = s.strip_prefix('>') {
        (CriteriaOp::Gt, rest)
    } else if let Some(rest) = s.strip_prefix('<') {
        (CriteriaOp::Lt, rest)
    } else if let Some(rest) = s.strip_prefix('=') {
        (CriteriaOp::Eq, rest)
    } else {
        (CriteriaOp::Eq, s)
    };
    Ok(Criteria {
        op,
        numeric: parse_number_literal(rest),
        value: rest.to_string(),
    })
}

fn criteria_matches(
    criteria: &Criteria,
    grid: &Grid,
    addr: &CellAddr,
    visiting: &mut Vec<CellAddr>,
    budget: &mut usize,
    allow_templates: bool,
) -> bool {
    match eval_cell_inner(grid, addr, visiting, budget, allow_templates).scalar_coerce() {
        EvalResult::Number(n) => match criteria.numeric {
            Some(target) => compare_f64(criteria.op, n, target),
            None => compare_str(criteria.op, &n.to_string(), &criteria.value),
        },
        EvalResult::Text(s) => match criteria.numeric {
            Some(target) => parse_number_literal(&s)
                .map(|n| compare_f64(criteria.op, n, target))
                .unwrap_or(false),
            None => compare_str(criteria.op, &s, &criteria.value),
        },
        EvalResult::Error(_) => false,
        EvalResult::Array(_) => false,
    }
}

#[derive(Clone, Debug, PartialEq)]
enum LookupValue {
    Number(f64),
    Text(String),
}

fn eval_lookup(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 2 && args.len() != 3 {
        return EvalResult::Error("ARGS");
    }
    let lookup_value = match eval_ast(&args[0], grid, visiting, bindings, budget, allow_templates)
        .scalar_coerce()
    {
        EvalResult::Number(n) => LookupValue::Number(n),
        EvalResult::Text(s) => parse_number_literal(&s)
            .map(LookupValue::Number)
            .unwrap_or(LookupValue::Text(s)),
        EvalResult::Error(e) => return EvalResult::Error(e),
        EvalResult::Array(_) => return EvalResult::Error("CALC"),
    };
    let Some(lookup_range) = as_main_range(&args[1]) else {
        return EvalResult::Error("RANGE");
    };
    let result_range = if args.len() == 3 {
        match as_main_range(&args[2]) {
            Some(r) => r,
            None => return EvalResult::Error("RANGE"),
        }
    } else {
        lookup_range.clone()
    };

    let lookup_rows = lookup_range.row_end - lookup_range.row_start;
    let lookup_cols = lookup_range.col_end - lookup_range.col_start;
    let result_rows = result_range.row_end - result_range.row_start;
    let result_cols = result_range.col_end - result_range.col_start;
    if lookup_rows != result_rows || lookup_cols != result_cols {
        return EvalResult::Error("ARGS");
    }
    if !(lookup_rows == 1 || lookup_cols == 1) {
        return EvalResult::Error("ARGS");
    }

    let len = if lookup_rows == 1 {
        lookup_cols
    } else {
        lookup_rows
    };
    for i in 0..len {
        let lookup_addr = if lookup_rows == 1 {
            CellAddr::Main {
                row: lookup_range.row_start,
                col: lookup_range.col_start + i,
            }
        } else {
            CellAddr::Main {
                row: lookup_range.row_start + i,
                col: lookup_range.col_start,
            }
        };
        let candidate = match eval_cell_inner(grid, &lookup_addr, visiting, budget, allow_templates)
            .scalar_coerce()
        {
            EvalResult::Number(n) => LookupValue::Number(n),
            EvalResult::Text(s) => parse_number_literal(&s)
                .map(LookupValue::Number)
                .unwrap_or(LookupValue::Text(s)),
            EvalResult::Error(_) => continue,
            EvalResult::Array(_) => continue,
        };
        if lookup_value == candidate {
            let result_addr = if result_rows == 1 {
                CellAddr::Main {
                    row: result_range.row_start,
                    col: result_range.col_start + i,
                }
            } else {
                CellAddr::Main {
                    row: result_range.row_start + i,
                    col: result_range.col_start,
                }
            };
            return eval_cell_inner(grid, &result_addr, visiting, budget, allow_templates);
        }
    }
    EvalResult::Error("NA")
}

fn eval_vlookup(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 3 && args.len() != 4 {
        return EvalResult::Error("ARGS");
    }
    let lookup_value = match eval_ast(&args[0], grid, visiting, bindings, budget, allow_templates)
        .scalar_coerce()
    {
        EvalResult::Number(n) => LookupValue::Number(n),
        EvalResult::Text(s) => parse_number_literal(&s)
            .map(LookupValue::Number)
            .unwrap_or(LookupValue::Text(s)),
        EvalResult::Error(e) => return EvalResult::Error(e),
        EvalResult::Array(_) => return EvalResult::Error("CALC"),
    };
    let Some(table) = as_main_range(&args[1]) else {
        return EvalResult::Error("RANGE");
    };
    let col_index = match eval_ast(&args[2], grid, visiting, bindings, budget, allow_templates) {
        EvalResult::Number(n) if n.is_finite() && n >= 1.0 && n.fract() == 0.0 => n as u32,
        EvalResult::Text(s) => match parse_number_literal(&s) {
            Some(n) if n.is_finite() && n >= 1.0 && n.fract() == 0.0 => n as u32,
            _ => return EvalResult::Error("VALUE"),
        },
        EvalResult::Error(e) => return EvalResult::Error(e),
        _ => return EvalResult::Error("VALUE"),
    };
    let table_rows = table.row_end - table.row_start;
    let table_cols = table.col_end - table.col_start;
    if col_index == 0 || col_index > table_cols {
        return EvalResult::Error("REF");
    }
    for dr in 0..table_rows {
        let key_addr = CellAddr::Main {
            row: table.row_start + dr,
            col: table.col_start,
        };
        let candidate = match eval_cell_inner(grid, &key_addr, visiting, budget, allow_templates)
            .scalar_coerce()
        {
            EvalResult::Number(n) => LookupValue::Number(n),
            EvalResult::Text(s) => parse_number_literal(&s)
                .map(LookupValue::Number)
                .unwrap_or(LookupValue::Text(s)),
            EvalResult::Error(_) => continue,
            EvalResult::Array(_) => continue,
        };
        if lookup_value == candidate {
            let result_addr = CellAddr::Main {
                row: table.row_start + dr,
                col: table.col_start + (col_index - 1),
            };
            return eval_cell_inner(grid, &result_addr, visiting, budget, allow_templates);
        }
    }
    EvalResult::Error("NA")
}

fn eval_xlookup(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 3 && args.len() != 4 {
        return EvalResult::Error("ARGS");
    }
    let lookup_value = match eval_ast(&args[0], grid, visiting, bindings, budget, allow_templates)
        .scalar_coerce()
    {
        EvalResult::Number(n) => LookupValue::Number(n),
        EvalResult::Text(s) => parse_number_literal(&s)
            .map(LookupValue::Number)
            .unwrap_or(LookupValue::Text(s)),
        EvalResult::Error(e) => return EvalResult::Error(e),
        EvalResult::Array(_) => return EvalResult::Error("CALC"),
    };
    let Some(lookup_range) = as_main_range(&args[1]) else {
        return EvalResult::Error("RANGE");
    };
    let Some(return_range) = as_main_range(&args[2]) else {
        return EvalResult::Error("RANGE");
    };

    let lookup_rows = lookup_range.row_end - lookup_range.row_start;
    let lookup_cols = lookup_range.col_end - lookup_range.col_start;
    let return_rows = return_range.row_end - return_range.row_start;
    let return_cols = return_range.col_end - return_range.col_start;
    if lookup_rows != return_rows || lookup_cols != return_cols {
        return EvalResult::Error("ARGS");
    }
    if !(lookup_rows == 1 || lookup_cols == 1) {
        return EvalResult::Error("ARGS");
    }

    let if_not_found = if args.len() == 4 {
        Some(eval_ast(
            &args[3],
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        ))
    } else {
        None
    };

    let len = if lookup_rows == 1 {
        lookup_cols
    } else {
        lookup_rows
    };
    for i in 0..len {
        let lookup_addr = if lookup_rows == 1 {
            CellAddr::Main {
                row: lookup_range.row_start,
                col: lookup_range.col_start + i,
            }
        } else {
            CellAddr::Main {
                row: lookup_range.row_start + i,
                col: lookup_range.col_start,
            }
        };
        let candidate = match eval_cell_inner(grid, &lookup_addr, visiting, budget, allow_templates)
            .scalar_coerce()
        {
            EvalResult::Number(n) => LookupValue::Number(n),
            EvalResult::Text(s) => parse_number_literal(&s)
                .map(LookupValue::Number)
                .unwrap_or(LookupValue::Text(s)),
            EvalResult::Error(_) => continue,
            EvalResult::Array(_) => continue,
        };
        if lookup_value == candidate {
            let result_addr = if return_rows == 1 {
                CellAddr::Main {
                    row: return_range.row_start,
                    col: return_range.col_start + i,
                }
            } else {
                CellAddr::Main {
                    row: return_range.row_start + i,
                    col: return_range.col_start,
                }
            };
            return eval_cell_inner(grid, &result_addr, visiting, budget, allow_templates);
        }
    }

    if let Some(v) = if_not_found {
        return v;
    }
    EvalResult::Error("NA")
}

fn eval_let(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() < 3 || args.len().is_multiple_of(2) {
        return EvalResult::Error("ARGS");
    }
    let base_len = bindings.len();
    let mut i = 0usize;
    while i + 1 < args.len() {
        let name = match &args[i] {
            Ast::Name(s) => s.clone(),
            _ => {
                bindings.truncate(base_len);
                return EvalResult::Error("NAME");
            }
        };
        let value = eval_ast(
            &args[i + 1],
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        );
        if matches!(value, EvalResult::Error(_)) {
            bindings.truncate(base_len);
            return value;
        }
        bindings.push((name, value));
        i += 2;
    }
    let result = eval_ast(
        &args[args.len() - 1],
        grid,
        visiting,
        bindings,
        budget,
        allow_templates,
    );
    bindings.truncate(base_len);
    result
}

fn eval_sequence(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.is_empty() || args.len() > 4 {
        return EvalResult::Error("ARGS");
    }
    let rows = match eval_ast(&args[0], grid, visiting, bindings, budget, allow_templates)
        .scalar_coerce()
    {
        EvalResult::Number(n) if n.is_finite() && n >= 0.0 => n as usize,
        _ => return EvalResult::Error("VALUE"),
    };
    let cols = if args.len() >= 2 {
        match eval_ast(&args[1], grid, visiting, bindings, budget, allow_templates).scalar_coerce()
        {
            EvalResult::Number(n) if n.is_finite() && n >= 0.0 => n as usize,
            _ => return EvalResult::Error("VALUE"),
        }
    } else {
        1
    };
    let start = if args.len() >= 3 {
        match eval_ast(&args[2], grid, visiting, bindings, budget, allow_templates).scalar_coerce()
        {
            EvalResult::Number(n) => n,
            _ => return EvalResult::Error("VALUE"),
        }
    } else {
        1.0
    };
    let step = if args.len() >= 4 {
        match eval_ast(&args[3], grid, visiting, bindings, budget, allow_templates).scalar_coerce()
        {
            EvalResult::Number(n) => n,
            _ => return EvalResult::Error("VALUE"),
        }
    } else {
        1.0
    };
    let mut out = Vec::with_capacity(rows);
    for r in 0..rows {
        let mut row = Vec::with_capacity(cols);
        for c in 0..cols {
            row.push(EvalResult::Number(start + step * (r * cols + c) as f64));
        }
        out.push(row);
    }
    EvalResult::Array(out)
}

fn eval_unique(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 1 {
        return EvalResult::Error("ARGS");
    }
    let values =
        match collect_array_values(&args[0], grid, visiting, bindings, budget, allow_templates) {
            Ok(v) => v,
            Err(e) => return EvalResult::Error(e),
        };
    let mut seen = Vec::<String>::new();
    let mut out = Vec::new();
    for v in values {
        let key = eval_result_to_key(&v);
        if !seen.contains(&key) {
            seen.push(key);
            out.push(vec![v]);
        }
    }
    EvalResult::Array(out)
}

fn eval_filter(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 2 {
        return EvalResult::Error("ARGS");
    }
    let values =
        match collect_array_values(&args[0], grid, visiting, bindings, budget, allow_templates) {
            Ok(v) => v,
            Err(e) => return EvalResult::Error(e),
        };
    let mask =
        match collect_array_values(&args[1], grid, visiting, bindings, budget, allow_templates) {
            Ok(v) => v,
            Err(e) => return EvalResult::Error(e),
        };
    let mut out = Vec::new();
    for (idx, v) in values.into_iter().enumerate() {
        let keep = mask
            .get(idx)
            .and_then(|cell| cell.top_left())
            .cloned()
            .map(super::truthy)
            .unwrap_or(false);
        if keep {
            out.push(vec![v]);
        }
    }
    EvalResult::Array(out)
}

fn collect_array_values(
    arg: &Ast,
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> Result<Vec<EvalResult>, &'static str> {
    match arg {
        Ast::Range(r) => {
            let mut out = Vec::new();
            for row in r.row_start..r.row_end {
                for col in r.col_start..r.col_end {
                    let addr = CellAddr::Main { row, col };
                    out.push(eval_cell_inner(
                        grid,
                        &addr,
                        visiting,
                        budget,
                        allow_templates,
                    ));
                }
            }
            Ok(out)
        }
        Ast::Ref(addr) => Ok(vec![eval_cell_inner(
            grid,
            addr,
            visiting,
            budget,
            allow_templates,
        )]),
        _ => Ok(vec![eval_ast(
            arg,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        )]),
    }
}

fn eval_result_to_key(result: &EvalResult) -> String {
    match result {
        EvalResult::Number(n) => n.to_string(),
        EvalResult::Text(s) => s.clone(),
        EvalResult::Error(e) => format!("#{e}"),
        EvalResult::Array(rows) => rows
            .first()
            .and_then(|row| row.first())
            .map(eval_result_to_key)
            .unwrap_or_else(|| "#CALC".to_string()),
    }
}
