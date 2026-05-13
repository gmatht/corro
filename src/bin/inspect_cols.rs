use std::path::Path;

fn main() {
    let path = Path::new("math.corro");
    let data = std::fs::read_to_string(path).expect("read math.corro");
    let mut workbook = corro::ops::WorkbookState::new();
    let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
    for line in data.lines() {
        let t = line.trim();
        if t.is_empty() {
            continue;
        }
        corro::ops::apply_log_line_to_workbook(t, &mut workbook, &mut active_sheet).unwrap();
    }

    let sheet = workbook.sheet_mut_by_id(active_sheet).unwrap();
    eprintln!("max_col_width = {}", sheet.grid.max_col_width());
    eprintln!("overrides:");
    for (col, w) in sheet.grid.col_width_overrides() {
        eprintln!("  col {} -> {}", col, w);
    }

    eprintln!("first 20 col widths:");
    for c in 0..20usize {
        eprintln!("  {}: {}", c, sheet.grid.col_width(c));
    }
}
