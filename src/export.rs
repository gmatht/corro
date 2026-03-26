//! TSV and CSV export for the main data region.

use crate::grid::{CellAddr, Grid};
use std::io::Write;

pub fn export_tsv(grid: &Grid, out: &mut dyn Write) {
    export_delimited(grid, out, '\t');
}

pub fn export_csv(grid: &Grid, out: &mut dyn Write) {
    export_delimited(grid, out, ',');
}

fn needs_csv_quoting(s: &str, delim: char) -> bool {
    s.contains(delim) || s.contains('"') || s.contains('\n') || s.contains('\r')
}

fn csv_quote(s: &str) -> String {
    format!("\"{}\"", s.replace('"', "\"\""))
}

fn export_delimited(grid: &Grid, out: &mut dyn Write, delim: char) {
    let mr = grid.main_rows();
    let mc = grid.main_cols();
    for r in 0..mr {
        for c in 0..mc {
            if c > 0 {
                let _ = write!(out, "{delim}");
            }
            let val = grid
                .get(&CellAddr::Main {
                    row: r as u32,
                    col: c as u32,
                })
                .unwrap_or("");
            if delim == ',' && needs_csv_quoting(val, delim) {
                let _ = write!(out, "{}", csv_quote(val));
            } else {
                let _ = write!(out, "{val}");
            }
        }
        let _ = writeln!(out);
    }
}
