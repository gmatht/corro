//! Book balancing by row reordering.

use crate::formula::{cell_effective_display, translate_formula_text, FormulaCopyContext};
use crate::grid::{CellAddr, Grid};
use crate::ops::SheetState;

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum BalanceDirection {
    PosToNeg,
    NegToPos,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct BalanceSourceRow {
    pub row_index: usize,
    pub amount_cents: i64,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct BalanceGroup {
    pub row_indices: Vec<usize>,
    pub total_cents: i64,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct BalanceReport {
    pub direction: BalanceDirection,
    pub amount_col: usize,
    pub groups: Vec<BalanceGroup>,
    pub leftovers: Vec<usize>,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct BalanceCopyPlan {
    pub source_sheet_id: u32,
    pub source_sheet_title: String,
    pub target_sheet_id: u32,
    pub target_title: String,
    pub amount_col: usize,
    pub row_order: Vec<usize>,
    pub unmatched_start: usize,
    pub show_unmatched_heading: bool,
    pub preserve_formulas: bool,
}

pub fn parse_amount_cents(raw: &str) -> Option<i64> {
    let t = raw.trim();
    if t.is_empty() {
        return None;
    }

    let mut s = t;
    let negative = match s.chars().next()? {
        '+' => {
            s = &s[1..];
            false
        }
        '-' => {
            s = &s[1..];
            true
        }
        _ => false,
    };

    let s = s.replace(',', "");
    let (whole, frac) = match s.split_once('.') {
        Some((whole, frac)) => (whole, frac),
        None => (s.as_str(), ""),
    };
    if whole.is_empty() || !whole.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    if !frac.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }

    let mut cents = whole.parse::<i64>().ok()?.saturating_mul(100);
    let frac_cents = match frac.len() {
        0 => 0,
        1 => frac.parse::<i64>().ok()?.saturating_mul(10),
        _ => frac.get(0..2)?.parse::<i64>().ok()?,
    };
    cents = cents.saturating_add(frac_cents);
    Some(if negative { -cents } else { cents })
}

pub fn format_amount_cents(cents: i64) -> String {
    let sign = if cents < 0 { "-" } else { "" };
    let abs = cents.abs();
    format!("{sign}{}.{:02}", abs / 100, abs % 100)
}

pub fn choose_balance_column(grid: &Grid) -> Option<usize> {
    for col in 0..grid.main_cols() {
        let mut saw_pos = false;
        let mut saw_neg = false;
        for row in 0..grid.main_rows() {
            let addr = CellAddr::Main {
                row: row as u32,
                col: col as u32,
            };
            if let Some(amount) = parse_amount_cents(grid.get(&addr).unwrap_or("")) {
                saw_pos |= amount > 0;
                saw_neg |= amount < 0;
            }
            if saw_pos && saw_neg {
                return Some(col);
            }
        }
    }
    None
}

pub fn source_rows_from_grid(grid: &Grid, col: usize) -> Vec<BalanceSourceRow> {
    let mut rows = Vec::new();
    for row in 0..grid.main_rows() {
        let amount_addr = CellAddr::Main {
            row: row as u32,
            col: col as u32,
        };
        let amount = parse_amount_cents(grid.get(&amount_addr).unwrap_or("")).unwrap_or(0);
        rows.push(BalanceSourceRow {
            row_index: row,
            amount_cents: amount,
        });
    }
    rows
}

pub fn balance_books(
    rows: &[BalanceSourceRow],
    direction: BalanceDirection,
    amount_col: usize,
) -> BalanceReport {
    let mut supply: Vec<BalanceSourceRow> = rows
        .iter()
        .filter(|row| match direction {
            BalanceDirection::PosToNeg => row.amount_cents > 0,
            BalanceDirection::NegToPos => row.amount_cents < 0,
        })
        .cloned()
        .collect();
    let mut demand: Vec<BalanceSourceRow> = rows
        .iter()
        .filter(|row| match direction {
            BalanceDirection::PosToNeg => row.amount_cents < 0,
            BalanceDirection::NegToPos => row.amount_cents > 0,
        })
        .cloned()
        .collect();

    supply.sort_by_key(|r| (r.row_index, r.amount_cents.abs()));
    demand.sort_by_key(|r| (r.row_index, r.amount_cents.abs()));

    let mut used = vec![false; demand.len()];
    let mut groups = Vec::new();

    for anchor in supply {
        let target = anchor.amount_cents.abs();
        let subset = choose_subset(&demand, &used, target);
        if let Some(indices) = subset {
            for &idx in &indices {
                used[idx] = true;
            }
            let total: i64 = indices.iter().map(|&idx| demand[idx].amount_cents).sum();
            groups.push(BalanceGroup {
                row_indices: std::iter::once(anchor.row_index)
                    .chain(indices.into_iter().map(|idx| demand[idx].row_index))
                    .collect(),
                total_cents: anchor.amount_cents + total,
            });
        }
    }

    let mut marked = vec![false; rows.len()];
    for group in &groups {
        for &row_index in &group.row_indices {
            if row_index < marked.len() {
                marked[row_index] = true;
            }
        }
    }
    let leftovers = (0..rows.len())
        .filter(|idx| !marked[*idx])
        .collect::<Vec<_>>();

    BalanceReport {
        direction,
        amount_col,
        groups,
        leftovers,
    }
}

fn choose_subset(rows: &[BalanceSourceRow], used: &[bool], target: i64) -> Option<Vec<usize>> {
    fn rec(
        rows: &[BalanceSourceRow],
        used: &[bool],
        target: i64,
        start: usize,
        current: &mut Vec<usize>,
    ) -> Option<Vec<usize>> {
        if target == 0 {
            return Some(current.clone());
        }
        if target < 0 {
            return None;
        }

        for idx in start..rows.len() {
            if used[idx] {
                continue;
            }
            let value = rows[idx].amount_cents.abs();
            if value > target {
                continue;
            }
            current.push(idx);
            if let Some(found) = rec(rows, used, target - value, idx + 1, current) {
                return Some(found);
            }
            current.pop();
        }
        None
    }

    rec(rows, used, target, 0, &mut Vec::new())
}

pub fn build_balance_report(grid: &Grid, col: usize, direction: BalanceDirection) -> BalanceReport {
    balance_books(&source_rows_from_grid(grid, col), direction, col)
}

pub fn balance_copy_plan(
    source_sheet_id: u32,
    source_sheet_title: String,
    target_sheet_id: u32,
    target_title: String,
    amount_col: usize,
    total_rows: usize,
    report: &BalanceReport,
    preserve_formulas: bool,
) -> BalanceCopyPlan {
    let row_order = row_order_from_groups(report, total_rows);
    let unmatched_start = report
        .groups
        .iter()
        .map(|group| group.row_indices.len())
        .sum::<usize>();
    let show_unmatched_heading = !report.groups.is_empty() && !report.leftovers.is_empty();
    BalanceCopyPlan {
        source_sheet_id,
        source_sheet_title,
        target_sheet_id,
        target_title,
        amount_col,
        row_order,
        unmatched_start,
        show_unmatched_heading,
        preserve_formulas,
    }
}

pub fn apply_balance_copy(source: &SheetState, target: &mut SheetState, plan: &BalanceCopyPlan) {
    let mc = source.grid.main_cols();
    let mr = source.grid.main_rows();
    let extra_row = usize::from(plan.show_unmatched_heading);
    target
        .grid
        .set_main_size(mr.saturating_add(extra_row).max(1), mc.max(1));
    target.grid.main_cells.clear();
    target.grid.left.clear();
    target.grid.right.clear();
    target.grid.clear_spills();
    target.grid.view_sort_cols.clear();

    let mut row_map = vec![0u32; mr];
    for (new_row, &old_row) in plan.row_order.iter().enumerate() {
        if old_row < row_map.len() {
            row_map[old_row] = if new_row >= plan.unmatched_start {
                (new_row + extra_row) as u32
            } else {
                new_row as u32
            };
        }
    }
    let ctx = FormulaCopyContext {
        source_sheet_id: plan.source_sheet_id,
        target_sheet_id: plan.target_sheet_id,
        row_map: row_map.clone(),
    };

    let copy_cell =
        |src: &SheetState, dst: &mut SheetState, src_addr: CellAddr, dst_addr: CellAddr| {
            if let Some(raw) = src.grid.get(&src_addr) {
                let value = if plan.preserve_formulas && raw.trim_start().starts_with('=') {
                    translate_formula_text(raw, &ctx)
                        .unwrap_or_else(|| cell_effective_display(&src.grid, &src_addr))
                } else {
                    raw.to_string()
                };
                if !value.is_empty() {
                    dst.grid.set(&dst_addr, value);
                }
            }
        };

    let total_cols = source.grid.total_cols();
    for row in 0..crate::grid::HEADER_ROWS {
        for col in 0..total_cols {
            let src_addr = CellAddr::Header {
                row: row as u8,
                col: col as u32,
            };
            let dst_addr = src_addr.clone();
            copy_cell(source, target, src_addr, dst_addr);
        }
    }

    for src_row in 0..mr {
        let dst_row = *row_map.get(src_row).unwrap_or(&(src_row as u32));
        for col in 0..mc {
            let src_addr = CellAddr::Main {
                row: src_row as u32,
                col: col as u32,
            };
            let dst_addr = CellAddr::Main {
                row: dst_row,
                col: col as u32,
            };
            copy_cell(source, target, src_addr, dst_addr);
        }
        for col in 0..crate::grid::MARGIN_COLS {
            let src_left = CellAddr::Left {
                col: col as u8,
                row: src_row as u32,
            };
            let dst_left = CellAddr::Left {
                col: col as u8,
                row: dst_row,
            };
            copy_cell(source, target, src_left, dst_left);

            let src_right = CellAddr::Right {
                col: col as u8,
                row: src_row as u32,
            };
            let dst_right = CellAddr::Right {
                col: col as u8,
                row: dst_row,
            };
            copy_cell(source, target, src_right, dst_right);
        }
    }

    if plan.show_unmatched_heading {
        let heading_row = plan.unmatched_start;
        if heading_row < target.grid.main_rows() {
            target.grid.set(
                &CellAddr::Main {
                    row: heading_row as u32,
                    col: 0,
                },
                "UNMATCHED".into(),
            );
        }
    }

    for row in 0..crate::grid::FOOTER_ROWS {
        for col in 0..total_cols {
            let src_addr = CellAddr::Footer {
                row: row as u8,
                col: col as u32,
            };
            let dst_addr = src_addr.clone();
            copy_cell(source, target, src_addr, dst_addr);
        }
    }

    target.grid.max_col_width = source.grid.max_col_width;
    target.grid.col_width_overrides = source.grid.col_width_overrides.clone();
    target.grid.volatile_seed = 0;
}

pub fn materialize_report_sheet(source: &SheetState, plan: &BalanceCopyPlan) -> SheetState {
    let mut target = SheetState::new(source.grid.main_rows(), source.grid.main_cols());
    apply_balance_copy(source, &mut target, plan);
    target
}

pub fn row_order_from_groups(report: &BalanceReport, total_rows: usize) -> Vec<usize> {
    let mut order = Vec::new();
    for group in &report.groups {
        order.extend(group.row_indices.iter().copied());
    }
    order.extend(report.leftovers.iter().copied());
    for row in 0..total_rows {
        if !order.contains(&row) {
            order.push(row);
        }
    }
    order
}

pub fn reordered_row_order(report: &BalanceReport, total_rows: usize) -> Vec<usize> {
    let mut order = Vec::new();
    for group in &report.groups {
        order.extend(group.row_indices.iter().copied());
    }
    order.extend(report.leftovers.iter().copied());
    for row in 0..total_rows {
        if !order.contains(&row) {
            order.push(row);
        }
    }
    order
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reordered_report_preserves_columns_and_translates_formulas() {
        let mut source = SheetState::new(2, 2);
        source
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "10".into());
        source
            .grid
            .set(&CellAddr::Main { row: 0, col: 1 }, "=A1".into());
        source
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "-10".into());
        source
            .grid
            .set(&CellAddr::Main { row: 1, col: 1 }, "=A2".into());

        let report = BalanceReport {
            direction: BalanceDirection::PosToNeg,
            amount_col: 0,
            groups: Vec::new(),
            leftovers: vec![1, 0],
        };
        let plan = balance_copy_plan(
            1,
            "Src".into(),
            2,
            "Dst".into(),
            0,
            source.grid.main_rows(),
            &report,
            true,
        );
        let out = materialize_report_sheet(&source, &plan);

        assert_eq!(out.grid.main_cols(), 2);
        assert_eq!(
            out.grid.get(&CellAddr::Main { row: 0, col: 0 }),
            Some("-10")
        );
        assert_eq!(
            out.grid.get(&CellAddr::Main { row: 0, col: 1 }),
            Some("=A1")
        );
        assert_eq!(out.grid.get(&CellAddr::Main { row: 1, col: 0 }), Some("10"));
        assert_eq!(
            out.grid.get(&CellAddr::Main { row: 1, col: 1 }),
            Some("=A2")
        );
    }

    #[test]
    fn unmatched_heading_is_inserted_after_matched_rows() {
        let mut source = SheetState::new(3, 1);
        source
            .grid
            .set(&CellAddr::Main { row: 0, col: 0 }, "10".into());
        source
            .grid
            .set(&CellAddr::Main { row: 1, col: 0 }, "-10".into());
        source
            .grid
            .set(&CellAddr::Main { row: 2, col: 0 }, "5".into());

        let report = build_balance_report(&source.grid, 0, BalanceDirection::PosToNeg);
        let plan = balance_copy_plan(
            1,
            "Src".into(),
            2,
            "Dst".into(),
            0,
            source.grid.main_rows(),
            &report,
            true,
        );
        let out = materialize_report_sheet(&source, &plan);

        assert_eq!(out.grid.get(&CellAddr::Main { row: 0, col: 0 }), Some("10"));
        assert_eq!(
            out.grid.get(&CellAddr::Main { row: 1, col: 0 }),
            Some("-10")
        );
        assert_eq!(
            out.grid.get(&CellAddr::Main { row: 2, col: 0 }),
            Some("UNMATCHED")
        );
        assert_eq!(out.grid.get(&CellAddr::Main { row: 3, col: 0 }), Some("5"));
    }
}
