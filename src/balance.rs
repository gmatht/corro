//! Book balancing by row reordering.

use crate::formula::{cell_effective_display, translate_formula_text, FormulaCopyContext};
use crate::grid::{CellAddr, GridBox as Grid};
use crate::ops::SheetState;
use balance_core as core;

pub use balance_core::{BalanceDirection, BalanceItem};

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct BalanceGroup {
    pub row_indices: Vec<usize>,
    pub total_cents: i64,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct BalanceSourceRow {
    pub row_index: usize,
    pub amount_cents: i64,
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
    core::parse_amount_cents(raw)
}

pub fn format_amount_cents(cents: i64) -> String {
    core::format_amount_cents(cents)
}

pub fn choose_balance_column(grid: &Grid) -> Option<usize> {
    let mut first_numeric = None;
    for col in 0..grid.main_cols() {
        let mut saw_pos = false;
        let mut saw_neg = false;
        for row in 0..grid.main_rows() {
            let addr = CellAddr::Main {
                row: row as u32,
                col: col as u32,
            };
            if let Some(raw) = grid.get(&addr) {
                if let Some(amount) = parse_amount_cents(&raw) {
                    first_numeric.get_or_insert(col);
                    saw_pos |= amount > 0;
                    saw_neg |= amount < 0;
                }
            }
            if saw_pos && saw_neg {
                return Some(col);
            }
        }
    }
    first_numeric
}

pub fn source_rows_from_grid(grid: &Grid, col: usize) -> Vec<BalanceSourceRow> {
    let mut rows = Vec::new();
    for row in 0..grid.main_rows() {
        let amount_addr = CellAddr::Main {
            row: row as u32,
            col: col as u32,
        };
        let amount = grid.get(&amount_addr).and_then(|s| parse_amount_cents(&s)).unwrap_or(0);
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
    let items: Vec<BalanceItem> = rows
        .iter()
        .map(|row| BalanceItem {
            id: row.row_index,
            amount_cents: row.amount_cents,
        })
        .collect();
    let sol = core::solve_balance_groups(&items, direction);
    BalanceReport {
        direction,
        amount_col,
        groups: sol
            .groups
            .into_iter()
            .map(|group| BalanceGroup {
                row_indices: group.item_ids,
                total_cents: group.total_cents,
            })
            .collect(),
        leftovers: sol.leftovers,
    }
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
    // Clear cell storage and spills on the target (use GridBox API).
    target.grid.clear_cells();
    target.grid.clear_spills();
    target.grid.set_view_sort_cols(Vec::new());

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
        main_cols: source.grid.main_cols(),
    };

    let copy_cell =
        |src: &SheetState, dst: &mut SheetState, src_addr: CellAddr, dst_addr: CellAddr| {
                if let Some(raw) = src.grid.get(&src_addr) {
                    let value = if plan.preserve_formulas && raw.trim_start().starts_with('=') {
                        // translate_formula_text expects &str; borrow the owned String
                        translate_formula_text(&raw, &ctx)
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
            let src_left = CellAddr::Left { col, row: src_row as u32 };
            let dst_left = CellAddr::Left { col, row: dst_row };
            copy_cell(source, target, src_left, dst_left);

            let src_right = CellAddr::Right { col, row: src_row as u32 };
            let dst_right = CellAddr::Right { col, row: dst_row };
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

    target
        .grid
        .set_max_col_width(source.grid.max_col_width());
    target
        .grid
        .set_col_width_overrides(source.grid.col_width_overrides());
    target.grid.set_volatile_seed(0);
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
    row_order_from_groups(report, total_rows)
}
