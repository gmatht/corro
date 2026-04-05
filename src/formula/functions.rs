use super::{
    as_main_range, compare_f64, compare_str, count_nonempty_values, count_numeric_values,
    criteria_from_ast, criteria_matches, eval_ast, eval_binary, eval_cell_inner, eval_sum,
    numeric_value, parse_number_literal, sum_main_range, Ast, CriteriaOp, EvalResult,
};
use crate::grid::{CellAddr, Grid};

pub(crate) fn eval_builtin(
    name: &str,
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
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
            budget,
            allow_templates,
            NumericAgg::Average,
        ),
        "MIN" => eval_numeric_aggregate(
            &args,
            grid,
            visiting,
            budget,
            allow_templates,
            NumericAgg::Min,
        ),
        "MAX" => eval_numeric_aggregate(
            &args,
            grid,
            visiting,
            budget,
            allow_templates,
            NumericAgg::Max,
        ),
        "COUNT" => eval_count(&args, grid, visiting, budget, allow_templates),
        "COUNTA" => eval_counta(&args, grid, visiting, budget, allow_templates),
        "PRODUCT" => eval_numeric_aggregate(
            &args,
            grid,
            visiting,
            budget,
            allow_templates,
            NumericAgg::Product,
        ),
        "ABS" => eval_unary_numeric(&args, grid, visiting, budget, allow_templates, f64::abs),
        "ROUND" => eval_round(&args, grid, visiting, budget, allow_templates),
        "MOD" => eval_binary_numeric(&args, grid, visiting, budget, allow_templates, |x, y| x % y),
        "SQRT" => eval_unary_numeric(&args, grid, visiting, budget, allow_templates, f64::sqrt),
        "PI" => {
            if !args.is_empty() {
                EvalResult::Error("ARGS")
            } else {
                EvalResult::Number(std::f64::consts::PI)
            }
        }
        "SIN" => eval_unary_numeric(&args, grid, visiting, budget, allow_templates, f64::sin),
        "COS" => eval_unary_numeric(&args, grid, visiting, budget, allow_templates, f64::cos),
        "TRIM" => eval_trim(&args, grid, visiting, budget, allow_templates),
        "COUNTIF" => eval_countif(&args, grid, visiting, budget, allow_templates),
        "SUMIF" => eval_sumif(&args, grid, visiting, budget, allow_templates),
        "IF" => {
            if args.len() != 3 {
                return EvalResult::Error("ARGS");
            }
            let cond = eval_ast(&args[0], grid, visiting, budget, allow_templates);
            let pick = super::truthy(cond);
            if pick {
                eval_ast(&args[1], grid, visiting, budget, allow_templates)
            } else {
                eval_ast(&args[2], grid, visiting, budget, allow_templates)
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
    budget: &mut usize,
    allow_templates: bool,
    kind: NumericAgg,
) -> EvalResult {
    if args.len() != 1 {
        return EvalResult::Error("ARGS");
    }
    let nums = match collect_numeric_values(&args[0], grid, visiting, budget, allow_templates) {
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
    budget: &mut usize,
    allow_templates: bool,
    f: fn(f64) -> f64,
) -> EvalResult {
    if args.len() != 1 {
        return EvalResult::Error("ARGS");
    }
    match numeric_value(eval_ast(&args[0], grid, visiting, budget, allow_templates)) {
        Some(n) => EvalResult::Number(f(n)),
        None => EvalResult::Error("VALUE"),
    }
}

fn eval_binary_numeric(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
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
        budget,
        allow_templates,
        f,
    )
}

fn eval_round(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 2 {
        return EvalResult::Error("ARGS");
    }
    let n = match numeric_value(eval_ast(&args[0], grid, visiting, budget, allow_templates)) {
        Some(n) => n,
        None => return EvalResult::Error("VALUE"),
    };
    let digits = match numeric_value(eval_ast(&args[1], grid, visiting, budget, allow_templates)) {
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
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 1 {
        return EvalResult::Error("ARGS");
    }
    match eval_ast(&args[0], grid, visiting, budget, allow_templates) {
        EvalResult::Number(n) => EvalResult::Text(trim_spaces(&format!("{n}"))),
        EvalResult::Text(s) => EvalResult::Text(trim_spaces(&s)),
        EvalResult::Error(e) => EvalResult::Error(e),
    }
}

fn eval_count(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 1 {
        return EvalResult::Error("ARGS");
    }
    match count_numeric_values(&args[0], grid, visiting, budget, allow_templates) {
        Ok(n) => EvalResult::Number(n as f64),
        Err(e) => EvalResult::Error(e),
    }
}

fn eval_counta(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 1 {
        return EvalResult::Error("ARGS");
    }
    match count_nonempty_values(&args[0], grid, visiting, budget, allow_templates) {
        Ok(n) => EvalResult::Number(n as f64),
        Err(e) => EvalResult::Error(e),
    }
}

fn eval_countif(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 2 {
        return EvalResult::Error("ARGS");
    }
    let Some(range) = as_main_range(&args[0]) else {
        return EvalResult::Error("RANGE");
    };
    let Ok(criteria) = criteria_from_ast(&args[1], grid, visiting, budget, allow_templates) else {
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
    let Ok(criteria) = criteria_from_ast(&args[1], grid, visiting, budget, allow_templates) else {
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
        _ => match eval_ast(arg, grid, visiting, budget, allow_templates) {
            EvalResult::Number(n) => Ok(vec![n]),
            EvalResult::Text(s) => Ok(parse_number_literal(&s).into_iter().collect()),
            EvalResult::Error(e) => Err(e),
        },
    }
}

fn trim_spaces(s: &str) -> String {
    s.split_whitespace().collect::<Vec<_>>().join(" ")
}
