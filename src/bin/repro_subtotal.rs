use std::fs;
use std::path::PathBuf;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let repo_root = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let path = repo_root.join("subtotal.corro");
    let data = fs::read_to_string(&path)?;

    let mut wb = corro::ops::WorkbookState::new();
    let mut active = wb.sheet_id(wb.active_sheet);

    for (i, line) in data.lines().enumerate() {
        let t = line.trim();
        if t.is_empty() {
            continue;
        }
        // Try parse as workbook op first to preserve sheet-qualified SETs
        match corro::ops::parse_workbook_line(t) {
            Ok(op) => {
                corro::ops::apply_workbook_op(&mut wb, &mut active, op)?;
                continue;
            }
            Err(_) => {}
        }
        if let Some(op) = corro::ops::parse_op_line(t) {
            let opw = corro::ops::WorkbookOp::SheetOp {
                sheet_id: active,
                op,
            };
            corro::ops::apply_workbook_op(&mut wb, &mut active, opw)?;
            continue;
        }
        // Fallback: try applying directly to active sheet
        corro::ops::apply_log_line_to_workbook(t, &mut wb, &mut active)?;
    }

    // Inspect sheet 1
    let sheet_id = 1u32;
    let sheet = wb
        .sheet_mut_by_id(sheet_id)
        .ok_or("missing sheet 1")?;
    let main_cols = sheet.grid.main_cols();
    println!("After replay: main_cols = {}", main_cols);

    let total = sheet.grid.total_cols();
    println!("Non-empty header cells:");
    for r in 0..corro::grid::HEADER_ROWS {
        for c in 0..total {
            let addr = corro::grid::CellAddr::Header { row: r as u8, col: c as u32 };
            let v = sheet.grid.get(&addr);
            if !v.is_empty() {
                println!("  {} => {}", corro::addr::cell_ref_text(&addr, main_cols), v);
            }
        }
    }

    println!("Non-empty footer cells:");
    for r in 0..corro::grid::FOOTER_ROWS {
        for c in 0..total {
            let addr = corro::grid::CellAddr::Footer { row: r as u8, col: c as u32 };
            let v = sheet.grid.get(&addr);
            if !v.is_empty() {
                println!("  {} => {}", corro::addr::cell_ref_text(&addr, main_cols), v);
            }
        }
    }

    Ok(())
}
