use corro::export::{DelimitedExportOptions,ExportContent};
use corro::grid::Grid;
use corro::grid::CellAddr;
use corro::grid::MARGIN_COLS;

fn main() {
    let mut g = Grid::new(1, 2);
    use corro::grid::HEADER_ROWS;
    g.set(&CellAddr::Header { row: (HEADER_ROWS - 1) as u32, col: (MARGIN_COLS + 1) as u32 }, "=A*0.1 -- TAX".into());
    g.set(&CellAddr::Header { row: (HEADER_ROWS - 1) as u32, col: (MARGIN_COLS + 2) as u32 }, "=TOTAL".into());
    g.set(&CellAddr::Left { col: MARGIN_COLS - 1, row: 0 }, "Hammers".into());
    g.set(&CellAddr::Main { row: 0, col: 0 }, "5".into());
    g.set(&CellAddr::Left { col: MARGIN_COLS - 1, row: 1 }, "=TOTAL".into());
    g.set(&CellAddr::Footer { row: 0, col: (MARGIN_COLS - 1) as u32 }, "=TOTAL".into());

    let opts = DelimitedExportOptions { content: ExportContent::Generic, include_margins: true, include_header_row: true, include_row_label_column: true };
    // Use the public helper delimited_export_matrix to derive span/rows.
    let (matrix, col_start, col_end, rows) = corro::export::delimited_export_matrix(&g.into(), &opts);
    eprintln!("matrix rows={} cols={}", matrix.len(), matrix.get(0).map(|r| r.len()).unwrap_or(0));
    eprintln!("col_start={} col_end={} rows={:?}", col_start, col_end, rows);
}
