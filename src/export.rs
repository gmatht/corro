//! TSV and CSV export for the main data region.

use crate::grid::{CellAddr, Grid, FOOTER_ROWS, HEADER_ROWS, MARGIN_COLS};
use std::io::Write;

pub fn export_tsv(grid: &Grid, out: &mut dyn Write) {
    export_delimited(grid, out, '\t', false, false);
}

pub fn export_csv(grid: &Grid, out: &mut dyn Write) {
    export_delimited(grid, out, ',', false, false);
}

pub fn export_ascii_table(grid: &Grid, out: &mut dyn Write) {
    let mc = grid.main_cols();
    let mr = grid.main_rows();
    let tc = grid.total_cols();
    let total_rows = HEADER_ROWS + mr + FOOTER_ROWS;

    let mut col_widths: Vec<usize> = vec![0; tc];
    for c in 0..tc {
        let label = col_header_label(c, mc);
        col_widths[c] = label.chars().count().max(1);
    }

    for r in 0..total_rows {
        for c in 0..tc {
            let val = cell_value_at(grid, r, c);
            col_widths[c] = col_widths[c].max(val.chars().count());
        }
    }

    let _total_width: usize = col_widths.iter().sum::<usize>() + tc + 1;
    let border: String = "+".to_string()
        + &col_widths
            .iter()
            .map(|w| "-".repeat(*w))
            .collect::<Vec<_>>()
            .join("+")
        + "+";

    let _ = writeln!(out, "{}", border);
    let mut header_line = String::from("|");
    for c in 0..tc {
        let label = col_header_label(c, mc);
        let w = col_widths[c];
        header_line.push_str(&format!(" {:>width$} |", label, width = w));
    }
    let _ = writeln!(out, "{}", header_line);
    let _ = writeln!(out, "{}", border);

    for r in 0..total_rows {
        let row_label = sheet_row_label(r, mr);
        let mut data_line = format!(
            "{:width$}|",
            row_label,
            width = row_label.chars().count().max(4)
        );
        for c in 0..tc {
            let val = cell_value_at(grid, r, c);
            let w = col_widths[c];
            data_line.push_str(&format!(" {:>width$} |", val, width = w));
        }
        let _ = writeln!(out, "{}", data_line);
        let _ = writeln!(out, "{}", border);
    }
}

pub fn export_all(grid: &Grid, out: &mut dyn Write) {
    export_delimited(grid, out, '\t', true, true);
}

pub fn export_selection(grid: &Grid, out: &mut dyn Write, rows: &[usize], cols: &[usize]) {
    if rows.is_empty() || cols.is_empty() {
        return;
    }

    for (ci, &c) in cols.iter().enumerate() {
        if ci > 0 {
            let _ = write!(out, "\t");
        }
        let label = col_header_label(c, grid.main_cols());
        let _ = write!(out, "{}", label);
    }
    let _ = writeln!(out);

    for &r in rows {
        for (ci, &c) in cols.iter().enumerate() {
            if ci > 0 {
                let _ = write!(out, "\t");
            }
            let val = cell_value_at(grid, r, c);
            let _ = write!(out, "{}", val);
        }
        let _ = writeln!(out);
    }
}

fn col_header_label(global_col: usize, main_cols: usize) -> String {
    let m = MARGIN_COLS;
    if global_col < m {
        format!("<{}", m - 1 - global_col)
    } else if global_col < m + main_cols {
        crate::addr::excel_column_name(global_col - m)
    } else {
        format!(">{}", global_col - m - main_cols)
    }
}

fn sheet_row_label(logical_row: usize, main_rows: usize) -> String {
    let hr = HEADER_ROWS;
    if logical_row < hr {
        format!("^{}", (b'Z' - logical_row as u8) as char)
    } else if logical_row < hr + main_rows {
        format!("{}", logical_row - hr + 1)
    } else {
        let fr = logical_row - hr - main_rows;
        format!("_{}", (b'A' + fr as u8) as char)
    }
}

fn cell_value_at(grid: &Grid, logical_row: usize, global_col: usize) -> String {
    let hr = HEADER_ROWS;
    let mr = grid.main_rows();
    let lm = MARGIN_COLS;
    let mc = grid.main_cols();
    let _fr = FOOTER_ROWS;

    if logical_row < hr {
        let r = logical_row as u8;
        grid.get(&CellAddr::Header {
            row: r,
            col: global_col as u32,
        })
        .unwrap_or("")
        .to_string()
    } else if logical_row < hr + mr {
        let mri = logical_row - hr;
        if global_col < lm {
            let c = (lm - 1 - global_col) as u8;
            grid.get(&CellAddr::Left {
                col: c,
                row: mri as u32,
            })
            .unwrap_or("")
            .to_string()
        } else if global_col < lm + mc {
            let mc_idx = global_col - lm;
            grid.get(&CellAddr::Main {
                row: mri as u32,
                col: mc_idx as u32,
            })
            .unwrap_or("")
            .to_string()
        } else {
            let rc = (global_col - lm - mc) as u8;
            grid.get(&CellAddr::Right {
                col: rc,
                row: mri as u32,
            })
            .unwrap_or("")
            .to_string()
        }
    } else {
        let fr_idx = logical_row - hr - mr;
        let r = fr_idx as u8;
        grid.get(&CellAddr::Footer {
            row: r,
            col: global_col as u32,
        })
        .unwrap_or("")
        .to_string()
    }
}

fn needs_csv_quoting(s: &str, delim: char) -> bool {
    s.contains(delim) || s.contains('"') || s.contains('\n') || s.contains('\r')
}

fn csv_quote(s: &str) -> String {
    format!("\"{}\"", s.replace('"', "\"\""))
}

fn export_delimited(
    grid: &Grid,
    out: &mut dyn Write,
    delim: char,
    include_headers: bool,
    include_margins: bool,
) {
    let mr = grid.main_rows();
    let mc = grid.main_cols();
    let tc = grid.total_cols();
    let hr = HEADER_ROWS;
    let lm = MARGIN_COLS;
    let _rm = MARGIN_COLS;
    let fr = FOOTER_ROWS;
    let total_rows = hr + mr + fr;

    let col_start = if include_margins { 0 } else { lm };
    let col_end = if include_margins { tc } else { lm + mc };

    if include_headers {
        for c in col_start..col_end {
            if c > col_start {
                let _ = write!(out, "{delim}");
            }
            let label = col_header_label(c, mc);
            let _ = write!(out, "{}", label);
        }
        let _ = writeln!(out);
    }

    for r in 0..total_rows {
        let mut first = true;
        for c in col_start..col_end {
            if !first {
                let _ = write!(out, "{delim}");
            }
            first = false;
            let val = cell_value_at(grid, r, c);
            if delim == ',' && needs_csv_quoting(&val, delim) {
                let _ = write!(out, "{}", csv_quote(&val));
            } else {
                let _ = write!(out, "{}", val);
            }
        }
        let _ = writeln!(out);
    }
}
