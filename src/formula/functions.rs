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
        "LEN" => eval_len(&args, grid, visiting, bindings, budget, allow_templates),
        "LEFT" => eval_left(&args, grid, visiting, bindings, budget, allow_templates),
        "RIGHT" => eval_right(&args, grid, visiting, bindings, budget, allow_templates),
        "MID" => eval_mid(&args, grid, visiting, bindings, budget, allow_templates),
        "CONCAT" => eval_concat(&args, grid, visiting, bindings, budget, allow_templates),
        "TEXTJOIN" => eval_textjoin(&args, grid, visiting, bindings, budget, allow_templates),
        "AND" => eval_and(&args, grid, visiting, bindings, budget, allow_templates),
        "OR" => eval_or(&args, grid, visiting, bindings, budget, allow_templates),
        "NOT" => eval_not(&args, grid, visiting, bindings, budget, allow_templates),
        "IFERROR" => eval_iferror(&args, grid, visiting, bindings, budget, allow_templates),
        "IFNA" => eval_ifna(&args, grid, visiting, bindings, budget, allow_templates),
        "COUNTIF" => eval_countif(&args, grid, visiting, bindings, budget, allow_templates),
        "SUMIF" => eval_sumif(&args, grid, visiting, bindings, budget, allow_templates),
        "COUNTIFS" => eval_countifs(&args, grid, visiting, bindings, budget, allow_templates),
        "SUMIFS" => eval_sumifs(&args, grid, visiting, bindings, budget, allow_templates),
        "AVERAGEIFS" => eval_averageifs(&args, grid, visiting, bindings, budget, allow_templates),
        "LOOKUP" => eval_lookup(&args, grid, visiting, bindings, budget, allow_templates),
        "VLOOKUP" => eval_vlookup(&args, grid, visiting, bindings, budget, allow_templates),
        "XLOOKUP" => eval_xlookup(&args, grid, visiting, bindings, budget, allow_templates),
        "MATCH" => eval_match(&args, grid, visiting, bindings, budget, allow_templates),
        "INDEX" => eval_index(&args, grid, visiting, bindings, budget, allow_templates),
        "LET" => eval_let(&args, grid, visiting, bindings, budget, allow_templates),
        "SEQUENCE" => eval_sequence(&args, grid, visiting, bindings, budget, allow_templates),
        "FILTER" => eval_filter(&args, grid, visiting, bindings, budget, allow_templates),
        "UNIQUE" => eval_unique(&args, grid, visiting, bindings, budget, allow_templates),
        "SORT" => eval_sort(&args, grid, visiting, bindings, budget, allow_templates),
        "TAKE" => eval_take(&args, grid, visiting, bindings, budget, allow_templates),
        "DROP" => eval_drop(&args, grid, visiting, bindings, budget, allow_templates),
        "CHOOSECOLS" => eval_choosecols(&args, grid, visiting, bindings, budget, allow_templates),
        "CHOOSEROWS" => eval_chooserows(&args, grid, visiting, bindings, budget, allow_templates),
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

fn eval_iferror(
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
    let value = eval_ast(&args[0], grid, visiting, bindings, budget, allow_templates);
    if matches!(value, EvalResult::Error(_) | EvalResult::Array(_)) {
        eval_ast(&args[1], grid, visiting, bindings, budget, allow_templates)
    } else {
        value
    }
}

fn eval_ifna(
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
    let value = eval_ast(&args[0], grid, visiting, bindings, budget, allow_templates);
    match value {
        EvalResult::Error("NA") | EvalResult::Error("PARSE") => {
            eval_ast(&args[1], grid, visiting, bindings, budget, allow_templates)
        }
        _ => value,
    }
}

fn eval_len(
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
        EvalResult::Text(s) => EvalResult::Number(s.chars().count() as f64),
        EvalResult::Number(n) => EvalResult::Number(format!("{n}").chars().count() as f64),
        EvalResult::Error(e) => EvalResult::Error(e),
        EvalResult::Array(_) => EvalResult::Error("CALC"),
    }
}

fn eval_left(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 1 && args.len() != 2 {
        return EvalResult::Error("ARGS");
    }
    let text = match eval_ast(&args[0], grid, visiting, bindings, budget, allow_templates)
        .scalar_coerce()
    {
        EvalResult::Text(s) => s,
        EvalResult::Number(n) => n.to_string(),
        EvalResult::Error(e) => return EvalResult::Error(e),
        EvalResult::Array(_) => return EvalResult::Error("CALC"),
    };
    let n = if args.len() == 2 {
        match numeric_value(eval_ast(
            &args[1],
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        )) {
            Some(v) if v >= 0.0 => v as usize,
            _ => return EvalResult::Error("VALUE"),
        }
    } else {
        1
    };
    EvalResult::Text(text.chars().take(n).collect())
}

fn eval_right(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 1 && args.len() != 2 {
        return EvalResult::Error("ARGS");
    }
    let text = match eval_ast(&args[0], grid, visiting, bindings, budget, allow_templates)
        .scalar_coerce()
    {
        EvalResult::Text(s) => s,
        EvalResult::Number(n) => n.to_string(),
        EvalResult::Error(e) => return EvalResult::Error(e),
        EvalResult::Array(_) => return EvalResult::Error("CALC"),
    };
    let n = if args.len() == 2 {
        match numeric_value(eval_ast(
            &args[1],
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        )) {
            Some(v) if v >= 0.0 => v as usize,
            _ => return EvalResult::Error("VALUE"),
        }
    } else {
        1
    };
    let chars: Vec<char> = text.chars().collect();
    let start = chars.len().saturating_sub(n);
    EvalResult::Text(chars[start..].iter().collect())
}

fn eval_mid(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() != 3 {
        return EvalResult::Error("ARGS");
    }
    let text = match eval_ast(&args[0], grid, visiting, bindings, budget, allow_templates)
        .scalar_coerce()
    {
        EvalResult::Text(s) => s,
        EvalResult::Number(n) => n.to_string(),
        EvalResult::Error(e) => return EvalResult::Error(e),
        EvalResult::Array(_) => return EvalResult::Error("CALC"),
    };
    let start = match numeric_value(eval_ast(
        &args[1],
        grid,
        visiting,
        bindings,
        budget,
        allow_templates,
    )) {
        Some(v) if v >= 1.0 => v as usize,
        _ => return EvalResult::Error("VALUE"),
    };
    let len = match numeric_value(eval_ast(
        &args[2],
        grid,
        visiting,
        bindings,
        budget,
        allow_templates,
    )) {
        Some(v) if v >= 0.0 => v as usize,
        _ => return EvalResult::Error("VALUE"),
    };
    let chars: Vec<char> = text.chars().collect();
    let start = start.saturating_sub(1);
    EvalResult::Text(chars.into_iter().skip(start).take(len).collect())
}

fn eval_concat(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    let mut out = String::new();
    for arg in args {
        match eval_ast(arg, grid, visiting, bindings, budget, allow_templates).scalar_coerce() {
            EvalResult::Number(n) => out.push_str(&n.to_string()),
            EvalResult::Text(s) => out.push_str(&s),
            EvalResult::Error(e) => return EvalResult::Error(e),
            EvalResult::Array(_) => return EvalResult::Error("CALC"),
        }
    }
    EvalResult::Text(out)
}

fn eval_textjoin(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() < 2 {
        return EvalResult::Error("ARGS");
    }
    let delim = match eval_ast(&args[0], grid, visiting, bindings, budget, allow_templates)
        .scalar_coerce()
    {
        EvalResult::Text(s) => s,
        EvalResult::Number(n) => n.to_string(),
        EvalResult::Error(e) => return EvalResult::Error(e),
        EvalResult::Array(_) => return EvalResult::Error("CALC"),
    };
    let mut parts = Vec::new();
    for arg in &args[1..] {
        let value =
            eval_ast(arg, grid, visiting, bindings, budget, allow_templates).scalar_coerce();
        match value {
            EvalResult::Number(n) => parts.push(n.to_string()),
            EvalResult::Text(s) => {
                if !s.is_empty() {
                    parts.push(s);
                }
            }
            EvalResult::Error(e) => return EvalResult::Error(e),
            EvalResult::Array(_) => return EvalResult::Error("CALC"),
        }
    }
    EvalResult::Text(parts.join(&delim))
}

fn eval_not(
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
    EvalResult::Number(
        if super::truthy(eval_ast(
            &args[0],
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        )) {
            0.0
        } else {
            1.0
        },
    )
}

fn eval_and(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    for arg in args {
        if !super::truthy(eval_ast(
            arg,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        )) {
            return EvalResult::Number(0.0);
        }
    }
    EvalResult::Number(1.0)
}

fn eval_or(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    for arg in args {
        if super::truthy(eval_ast(
            arg,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        )) {
            return EvalResult::Number(1.0);
        }
    }
    EvalResult::Number(0.0)
}

fn eval_match(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() < 2 || args.len() > 3 {
        return EvalResult::Error("ARGS");
    }
    if args.len() == 3 {
        let mt = match numeric_value(eval_ast(
            &args[2],
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        )) {
            Some(n) => n,
            None => return EvalResult::Error("VALUE"),
        };
        if mt != 0.0 {
            return EvalResult::Error("NA");
        }
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
    let matrix =
        match collect_matrix_values(&args[1], grid, visiting, bindings, budget, allow_templates) {
            Ok(m) => m,
            Err(e) => return EvalResult::Error(e),
        };
    let mut idx = 1u32;
    for row in matrix {
        for cell in row {
            let candidate = match cell.scalar_coerce() {
                EvalResult::Number(n) => LookupValue::Number(n),
                EvalResult::Text(s) => parse_number_literal(&s)
                    .map(LookupValue::Number)
                    .unwrap_or(LookupValue::Text(s)),
                EvalResult::Error(_) => {
                    idx += 1;
                    continue;
                }
                EvalResult::Array(_) => {
                    idx += 1;
                    continue;
                }
            };
            if candidate == lookup_value {
                return EvalResult::Number(idx as f64);
            }
            idx += 1;
        }
    }
    EvalResult::Error("NA")
}

fn eval_index(
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
    let matrix =
        match collect_matrix_values(&args[0], grid, visiting, bindings, budget, allow_templates) {
            Ok(m) => m,
            Err(e) => return EvalResult::Error(e),
        };
    if matrix.is_empty() || matrix[0].is_empty() {
        return EvalResult::Error("REF");
    }
    let row = match numeric_value(eval_ast(
        &args[1],
        grid,
        visiting,
        bindings,
        budget,
        allow_templates,
    )) {
        Some(n) if n >= 1.0 => n as usize,
        _ => return EvalResult::Error("VALUE"),
    };
    let col = if args.len() == 3 {
        match numeric_value(eval_ast(
            &args[2],
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        )) {
            Some(n) if n >= 1.0 => n as usize,
            _ => return EvalResult::Error("VALUE"),
        }
    } else if matrix.len() == 1 {
        row
    } else {
        1
    };
    if row == 0 || col == 0 {
        return EvalResult::Error("REF");
    }
    if row > matrix.len() {
        return EvalResult::Error("REF");
    }
    let row_idx = row - 1;
    if col > matrix[row_idx].len() {
        return EvalResult::Error("REF");
    }
    matrix[row_idx][col - 1].clone()
}

fn eval_countifs(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() < 2 || !args.len().is_multiple_of(2) {
        return EvalResult::Error("ARGS");
    }
    let pairs =
        match collect_criteria_pairs(args, grid, visiting, bindings, budget, allow_templates) {
            Ok(p) => p,
            Err(e) => return EvalResult::Error(e),
        };
    let Some((first_range, _)) = pairs.first() else {
        return EvalResult::Error("ARGS");
    };
    let mut count = 0usize;
    for dr in 0..range_height(first_range) {
        for dc in 0..range_width(first_range) {
            let mut ok = true;
            for (range, criteria) in &pairs {
                let addr = CellAddr::Main {
                    row: range.row_start + dr,
                    col: range.col_start + dc,
                };
                if !criteria_matches(criteria, grid, &addr, visiting, budget, allow_templates) {
                    ok = false;
                    break;
                }
            }
            if ok {
                count += 1;
            }
        }
    }
    EvalResult::Number(count as f64)
}

fn eval_sumifs(
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
    let Some(sum_range) = as_main_range(&args[0]) else {
        return EvalResult::Error("RANGE");
    };
    let pairs = match collect_criteria_pairs(
        &args[1..],
        grid,
        visiting,
        bindings,
        budget,
        allow_templates,
    ) {
        Ok(p) => p,
        Err(e) => return EvalResult::Error(e),
    };
    if pairs.iter().any(|(range, _)| {
        range_height(range) != range_height(&sum_range)
            || range_width(range) != range_width(&sum_range)
    }) {
        return EvalResult::Error("ARGS");
    }
    let mut sum = 0.0;
    for dr in 0..range_height(&sum_range) {
        for dc in 0..range_width(&sum_range) {
            let mut ok = true;
            for (range, criteria) in &pairs {
                let addr = CellAddr::Main {
                    row: range.row_start + dr,
                    col: range.col_start + dc,
                };
                if !criteria_matches(criteria, grid, &addr, visiting, budget, allow_templates) {
                    ok = false;
                    break;
                }
            }
            if ok {
                let addr = CellAddr::Main {
                    row: sum_range.row_start + dr,
                    col: sum_range.col_start + dc,
                };
                if let Some(n) = super::effective_numeric(grid, &addr, visiting, budget) {
                    sum += n;
                }
            }
        }
    }
    EvalResult::Number(sum)
}

fn eval_averageifs(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    match eval_sumifs(args, grid, visiting, bindings, budget, allow_templates) {
        EvalResult::Number(sum) => {
            let sum_range = match as_main_range(&args[0]) {
                Some(r) => r,
                None => return EvalResult::Error("RANGE"),
            };
            let pairs = match collect_criteria_pairs(
                &args[1..],
                grid,
                visiting,
                bindings,
                budget,
                allow_templates,
            ) {
                Ok(p) => p,
                Err(e) => return EvalResult::Error(e),
            };
            let mut count = 0usize;
            for dr in 0..range_height(&sum_range) {
                for dc in 0..range_width(&sum_range) {
                    let mut ok = true;
                    for (range, criteria) in &pairs {
                        let addr = CellAddr::Main {
                            row: range.row_start + dr,
                            col: range.col_start + dc,
                        };
                        if !criteria_matches(
                            criteria,
                            grid,
                            &addr,
                            visiting,
                            budget,
                            allow_templates,
                        ) {
                            ok = false;
                            break;
                        }
                    }
                    if ok {
                        let addr = CellAddr::Main {
                            row: sum_range.row_start + dr,
                            col: sum_range.col_start + dc,
                        };
                        if super::effective_numeric(grid, &addr, visiting, budget).is_some() {
                            count += 1;
                        }
                    }
                }
            }
            if count == 0 {
                EvalResult::Error("DIV0")
            } else {
                EvalResult::Number(sum / count as f64)
            }
        }
        other => other,
    }
}

fn eval_sort(
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
    let mut matrix =
        match collect_matrix_values(&args[0], grid, visiting, bindings, budget, allow_templates) {
            Ok(m) => m,
            Err(e) => return EvalResult::Error(e),
        };
    let sort_index = if args.len() >= 2 {
        match numeric_value(eval_ast(
            &args[1],
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        )) {
            Some(n) if n >= 1.0 => n as usize,
            _ => return EvalResult::Error("VALUE"),
        }
    } else {
        1
    };
    let sort_order = if args.len() >= 3 {
        match numeric_value(eval_ast(
            &args[2],
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        )) {
            Some(n) if n < 0.0 => -1,
            Some(_) => 1,
            None => return EvalResult::Error("VALUE"),
        }
    } else {
        1
    };
    let by_col = if args.len() == 4 {
        super::truthy(eval_ast(
            &args[3],
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        ))
    } else {
        false
    };
    if by_col {
        transpose_matrix(&mut matrix);
    }
    let key_col = sort_index.saturating_sub(1);
    if matrix.is_empty() || key_col >= matrix[0].len() {
        return EvalResult::Error("REF");
    }
    matrix.sort_by(|a, b| {
        compare_eval_cells(&a[key_col], &b[key_col]).then(std::cmp::Ordering::Equal)
    });
    if sort_order < 0 {
        matrix.reverse();
    }
    if by_col {
        transpose_matrix(&mut matrix);
    }
    EvalResult::Array(matrix)
}

fn eval_take(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() < 2 || args.len() > 3 {
        return EvalResult::Error("ARGS");
    }
    let matrix =
        match collect_matrix_values(&args[0], grid, visiting, bindings, budget, allow_templates) {
            Ok(m) => m,
            Err(e) => return EvalResult::Error(e),
        };
    slice_take_drop(
        matrix,
        &args[1],
        args.get(2),
        grid,
        visiting,
        bindings,
        budget,
        allow_templates,
        true,
    )
}

fn eval_drop(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() < 2 || args.len() > 3 {
        return EvalResult::Error("ARGS");
    }
    let matrix =
        match collect_matrix_values(&args[0], grid, visiting, bindings, budget, allow_templates) {
            Ok(m) => m,
            Err(e) => return EvalResult::Error(e),
        };
    slice_take_drop(
        matrix,
        &args[1],
        args.get(2),
        grid,
        visiting,
        bindings,
        budget,
        allow_templates,
        false,
    )
}

fn eval_choosecols(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() < 2 {
        return EvalResult::Error("ARGS");
    }
    let matrix =
        match collect_matrix_values(&args[0], grid, visiting, bindings, budget, allow_templates) {
            Ok(m) => m,
            Err(e) => return EvalResult::Error(e),
        };
    choose_axes(
        matrix,
        &args[1..],
        grid,
        visiting,
        bindings,
        budget,
        allow_templates,
        true,
    )
}

fn eval_chooserows(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> EvalResult {
    if args.len() < 2 {
        return EvalResult::Error("ARGS");
    }
    let matrix =
        match collect_matrix_values(&args[0], grid, visiting, bindings, budget, allow_templates) {
            Ok(m) => m,
            Err(e) => return EvalResult::Error(e),
        };
    choose_axes(
        matrix,
        &args[1..],
        grid,
        visiting,
        bindings,
        budget,
        allow_templates,
        false,
    )
}

fn collect_matrix_values(
    arg: &Ast,
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> Result<Vec<Vec<EvalResult>>, &'static str> {
    match arg {
        Ast::Range(r) => {
            let mut out = Vec::new();
            for row in r.row_start..r.row_end {
                let mut out_row = Vec::new();
                for col in r.col_start..r.col_end {
                    let addr = CellAddr::Main { row, col };
                    out_row.push(
                        eval_cell_inner(grid, &addr, visiting, budget, allow_templates)
                            .scalar_coerce(),
                    );
                }
                out.push(out_row);
            }
            Ok(out)
        }
        Ast::Ref(addr) => Ok(vec![vec![eval_cell_inner(
            grid,
            addr,
            visiting,
            budget,
            allow_templates,
        )
        .scalar_coerce()]]),
        _ => match eval_ast(arg, grid, visiting, bindings, budget, allow_templates) {
            EvalResult::Array(rows) => Ok(rows),
            other => Ok(vec![vec![other]]),
        },
    }
}

fn collect_criteria_pairs(
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
) -> Result<Vec<(MainRange, Criteria)>, &'static str> {
    if args.len() < 2 || !args.len().is_multiple_of(2) {
        return Err("ARGS");
    }
    let mut out = Vec::new();
    for pair in args.chunks(2) {
        let Some(range) = as_main_range(&pair[0]) else {
            return Err("RANGE");
        };
        let criteria =
            criteria_from_ast(&pair[1], grid, visiting, bindings, budget, allow_templates)?;
        out.push((range, criteria));
    }
    let base = out[0].0.clone();
    if out.iter().any(|(r, _)| {
        range_height(r) != range_height(&base) || range_width(r) != range_width(&base)
    }) {
        return Err("ARGS");
    }
    Ok(out)
}

fn range_height(r: &MainRange) -> u32 {
    r.row_end.saturating_sub(r.row_start)
}

fn range_width(r: &MainRange) -> u32 {
    r.col_end.saturating_sub(r.col_start)
}

fn compare_eval_cells(a: &EvalResult, b: &EvalResult) -> std::cmp::Ordering {
    use std::cmp::Ordering;
    let a = a.clone().scalar_coerce();
    let b = b.clone().scalar_coerce();
    match (a, b) {
        (EvalResult::Number(x), EvalResult::Number(y)) => {
            x.partial_cmp(&y).unwrap_or(Ordering::Equal)
        }
        (EvalResult::Text(x), EvalResult::Text(y)) => x.cmp(&y),
        (EvalResult::Number(_), EvalResult::Text(_)) => Ordering::Less,
        (EvalResult::Text(_), EvalResult::Number(_)) => Ordering::Greater,
        (EvalResult::Error(x), EvalResult::Error(y)) => x.cmp(y),
        (EvalResult::Error(_), _) => Ordering::Greater,
        (_, EvalResult::Error(_)) => Ordering::Less,
        _ => Ordering::Equal,
    }
}

fn transpose_matrix(matrix: &mut Vec<Vec<EvalResult>>) {
    if matrix.is_empty() {
        return;
    }
    let rows = matrix.len();
    let cols = matrix.iter().map(|r| r.len()).max().unwrap_or(0);
    let mut out = vec![vec![EvalResult::Error("CALC"); rows]; cols];
    for (r, row) in matrix.iter().enumerate() {
        for (c, cell) in row.iter().enumerate() {
            out[c][r] = cell.clone();
        }
    }
    *matrix = out;
}

fn slice_take_drop(
    matrix: Vec<Vec<EvalResult>>,
    row_arg: &Ast,
    col_arg: Option<&Ast>,
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
    take: bool,
) -> EvalResult {
    let rows_n = match numeric_value(eval_ast(
        row_arg,
        grid,
        visiting,
        bindings,
        budget,
        allow_templates,
    )) {
        Some(n) if n.is_finite() => n as isize,
        _ => return EvalResult::Error("VALUE"),
    };
    let cols_n = if let Some(col_arg) = col_arg {
        match numeric_value(eval_ast(
            col_arg,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        )) {
            Some(n) if n.is_finite() => n as isize,
            _ => return EvalResult::Error("VALUE"),
        }
    } else {
        0
    };
    let mut out = matrix;
    if take {
        out = slice_rows(out, rows_n, true);
        if col_arg.is_some() {
            out = slice_cols(out, cols_n, true);
        }
    } else {
        out = slice_rows(out, rows_n, false);
        if col_arg.is_some() {
            out = slice_cols(out, cols_n, false);
        }
    }
    EvalResult::Array(out)
}

fn slice_rows(mut matrix: Vec<Vec<EvalResult>>, n: isize, take: bool) -> Vec<Vec<EvalResult>> {
    if matrix.is_empty() {
        return matrix;
    }
    let len = matrix.len() as isize;
    let n = if n < 0 { len + n } else { n };
    if take {
        if n >= 0 {
            matrix.truncate(n.min(len) as usize);
            matrix
        } else {
            let keep = (len + n).max(0) as usize;
            matrix.into_iter().take(keep).collect()
        }
    } else if n >= 0 {
        matrix.into_iter().skip(n.min(len) as usize).collect()
    } else {
        let keep = (len + n).max(0) as usize;
        matrix.truncate(keep);
        matrix
    }
}

fn slice_cols(matrix: Vec<Vec<EvalResult>>, n: isize, take: bool) -> Vec<Vec<EvalResult>> {
    if matrix.is_empty() {
        return matrix;
    }
    let len = matrix[0].len() as isize;
    let n = if n < 0 { len + n } else { n };
    matrix
        .into_iter()
        .map(|mut row| {
            if take {
                row.truncate(n.min(len) as usize);
                row
            } else if n >= 0 {
                row.into_iter().skip(n.min(len) as usize).collect()
            } else {
                let keep = (len + n).max(0) as usize;
                row.truncate(keep);
                row
            }
        })
        .collect()
}

fn choose_axes(
    matrix: Vec<Vec<EvalResult>>,
    args: &[Ast],
    grid: &Grid,
    visiting: &mut Vec<CellAddr>,
    bindings: &mut Vec<(String, EvalResult)>,
    budget: &mut usize,
    allow_templates: bool,
    cols: bool,
) -> EvalResult {
    let mut indices = Vec::new();
    for arg in args {
        let idx = match numeric_value(eval_ast(
            arg,
            grid,
            visiting,
            bindings,
            budget,
            allow_templates,
        )) {
            Some(n) if n.is_finite() && n != 0.0 => n as isize,
            _ => return EvalResult::Error("VALUE"),
        };
        indices.push(idx);
    }
    if cols {
        let mut out = Vec::new();
        for row in matrix {
            let mut new_row = Vec::new();
            for idx in &indices {
                let j = resolve_index(*idx, row.len());
                let Some(j) = j else {
                    return EvalResult::Error("REF");
                };
                new_row.push(row[j].clone());
            }
            out.push(new_row);
        }
        EvalResult::Array(out)
    } else {
        let mut out = Vec::new();
        for idx in indices {
            let j = match resolve_index(idx, matrix.len()) {
                Some(v) => v,
                None => return EvalResult::Error("REF"),
            };
            out.push(matrix[j].clone());
        }
        EvalResult::Array(out)
    }
}

fn resolve_index(idx: isize, len: usize) -> Option<usize> {
    if idx > 0 {
        let i = idx as usize - 1;
        (i < len).then_some(i)
    } else if idx < 0 {
        let i = len as isize + idx;
        (i >= 0).then_some(i as usize)
    } else {
        None
    }
}

fn text_from_result(result: EvalResult) -> Option<String> {
    match result {
        EvalResult::Number(n) => Some(n.to_string()),
        EvalResult::Text(s) => Some(s),
        EvalResult::Error(_) => None,
        EvalResult::Array(rows) => rows
            .into_iter()
            .next()
            .and_then(|row| row.into_iter().next())
            .and_then(text_from_result),
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
