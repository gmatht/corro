//! TSV and CSV export for the main data region.

use crate::grid::{CellAddr, Grid, FOOTER_ROWS, HEADER_ROWS, MARGIN_COLS};
use std::io::Write;
use zip::write::FileOptions;

pub fn export_tsv(grid: &Grid, out: &mut dyn Write) {
    export_delimited(grid, out, '\t', false, false);
}

pub fn export_csv(grid: &Grid, out: &mut dyn Write) {
    export_delimited(grid, out, ',', false, false);
}

pub fn export_ascii_table(grid: &Grid, out: &mut dyn Write) {
    let mc = grid.main_cols();
    let tc = grid.total_cols();
    let (row_start, row_end) = ascii_row_bounds(grid);
    let (col_start, col_end) = ascii_col_bounds(grid);

    let mut col_widths: Vec<usize> = vec![0; tc];
    for c in col_start..col_end {
        let label = col_header_label(c, mc);
        col_widths[c] = label.chars().count().max(1);
    }

    for r in row_start..row_end {
        for c in col_start..col_end {
            let val = rendered_value_at(grid, r, c);
            let content_w = if val.is_empty() {
                0
            } else {
                val.chars().count() + 1
            };
            col_widths[c] = col_widths[c].max(content_w);
        }
    }

    let _total_width: usize =
        col_widths[col_start..col_end].iter().sum::<usize>() + (col_end - col_start) + 1;
    let border: String = "+".to_string()
        + &col_widths[col_start..col_end]
            .iter()
            .map(|w| "-".repeat(*w))
            .collect::<Vec<_>>()
            .join("+")
        + "+";

    let _ = writeln!(out, "{}", border);
    let mut header_line = String::from("|");
    for c in col_start..col_end {
        let label = col_header_label(c, mc);
        let w = col_widths[c];
        header_line.push_str(&format!(" {:>width$} |", label, width = w));
    }
    let _ = writeln!(out, "{}", header_line);
    let _ = writeln!(out, "{}", border);

    for r in row_start..row_end {
        let row_label = sheet_row_label(r, grid.main_rows());
        let mut data_line = format!(
            "{:width$}|",
            row_label,
            width = row_label.chars().count().max(4)
        );
        for c in col_start..col_end {
            let val = rendered_value_at(grid, r, c);
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

pub fn export_odt_bytes(grid: &Grid) -> Result<Vec<u8>, std::io::Error> {
    let cursor = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let opt = FileOptions::default().compression_method(zip::CompressionMethod::Stored);

    zip.start_file("mimetype", opt)?;
    zip.write_all(b"application/vnd.oasis.opendocument.text")?;

    zip.start_file("content.xml", FileOptions::default())?;
    let xml = odt_content_xml(grid);
    zip.write_all(xml.as_bytes())?;

    zip.start_file("META-INF/manifest.xml", FileOptions::default())?;
    zip.write_all(odt_manifest_xml().as_bytes())?;

    let cursor = zip.finish()?;
    Ok(cursor.into_inner())
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
        format!("~{}", (b'Z' - logical_row as u8) as char)
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

fn rendered_value_at(grid: &Grid, logical_row: usize, global_col: usize) -> String {
    let addr = match cell_addr_at(grid, logical_row, global_col) {
        Some(addr) => addr,
        None => return String::new(),
    };
    crate::ui::format_cell_display(
        grid,
        &addr,
        crate::formula::cell_effective_display(grid, &addr),
    )
}

fn cell_addr_at(grid: &Grid, logical_row: usize, global_col: usize) -> Option<CellAddr> {
    let hr = HEADER_ROWS;
    let mr = grid.main_rows();
    let lm = MARGIN_COLS;
    let mc = grid.main_cols();

    if logical_row < hr {
        Some(CellAddr::Header {
            row: logical_row as u8,
            col: global_col as u32,
        })
    } else if logical_row < hr + mr {
        let mri = logical_row - hr;
        if global_col < lm {
            Some(CellAddr::Left {
                col: (lm - 1 - global_col) as u8,
                row: mri as u32,
            })
        } else if global_col < lm + mc {
            Some(CellAddr::Main {
                row: mri as u32,
                col: (global_col - lm) as u32,
            })
        } else {
            Some(CellAddr::Right {
                col: (global_col - lm - mc) as u8,
                row: mri as u32,
            })
        }
    } else {
        Some(CellAddr::Footer {
            row: (logical_row - hr - mr) as u8,
            col: global_col as u32,
        })
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

    for r in row_order(grid, total_rows) {
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

fn odt_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn odt_content_xml(grid: &Grid) -> String {
    let mr = grid.main_rows();
    let tc = grid.total_cols();
    let total_rows = HEADER_ROWS + mr + FOOTER_ROWS;

    let mut s = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.2"><office:body><office:text><table:table>"#,
    );

    for c in 0..tc {
        s.push_str(&format!(
            r#"<table:table-column table:number-columns-repeated="1" table:style-name="co{}"/>"#,
            c
        ));
    }

    for r in row_order(grid, total_rows) {
        s.push_str("<table:table-row>");
        for c in 0..tc {
            let val = odt_escape(&cell_value_at(grid, r, c));
            let text = if val.is_empty() {
                String::new()
            } else {
                format!(r#"<text:p>{}</text:p>"#, val)
            };
            s.push_str(&format!(
                r#"<table:table-cell office:value-type="string">{}</table:table-cell>"#,
                text
            ));
        }
        s.push_str("</table:table-row>");
    }

    s.push_str("</table:table></office:text></office:body></office:document-content>");
    s
}

fn ascii_row_bounds(grid: &Grid) -> (usize, usize) {
    let hr = HEADER_ROWS;
    let mr = grid.main_rows();
    let fr = FOOTER_ROWS;
    let mut start = 0;
    while start < hr + mr + fr && !grid.logical_row_has_content(start) {
        start += 1;
    }
    let mut end = hr + mr + fr;
    while end > start && !grid.logical_row_has_content(end - 1) {
        end -= 1;
    }
    (start, end.max(start + 1))
}

fn ascii_col_bounds(grid: &Grid) -> (usize, usize) {
    let tc = grid.total_cols();
    let mut start = 0;
    while start < tc && !grid.logical_col_has_content(start) {
        start += 1;
    }
    let mut end = tc;
    while end > start && !grid.logical_col_has_content(end - 1) {
        end -= 1;
    }
    (start, end.max(start + 1))
}

fn odt_manifest_xml() -> String {
    String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>"#,
    )
}

fn row_order(grid: &Grid, total_rows: usize) -> Vec<usize> {
    let hr = HEADER_ROWS;
    let mr = grid.main_rows();
    let fr = FOOTER_ROWS;
    let mut rows: Vec<usize> = Vec::with_capacity(total_rows);
    rows.extend(0..hr);

    rows.extend(grid.sorted_main_rows().into_iter().map(|r| hr + r));

    rows.extend((0..fr).map(|r| hr + mr + r));
    rows
}

pub fn export_sorted_tsv(grid: &Grid, out: &mut dyn Write, sort_cols: &[usize]) {
    let mr = grid.main_rows();
    let mc = grid.main_cols();
    if sort_cols.is_empty() {
        export_tsv(grid, out);
        return;
    }

    let mut rows: Vec<usize> = (0..mr).collect();
    rows.sort_by(|&a, &b| {
        for &c in sort_cols {
            let va = grid
                .get(&CellAddr::Main {
                    row: a as u32,
                    col: c as u32,
                })
                .unwrap_or("");
            let vb = grid
                .get(&CellAddr::Main {
                    row: b as u32,
                    col: c as u32,
                })
                .unwrap_or("");
            let ord = va.cmp(vb);
            if ord != std::cmp::Ordering::Equal {
                return ord;
            }
        }
        a.cmp(&b)
    });

    for (i, c) in (0..mc).enumerate() {
        if i > 0 {
            let _ = write!(out, "\t");
        }
        let _ = write!(out, "{}", col_header_label(MARGIN_COLS + c, mc));
    }
    let _ = writeln!(out);

    for r in rows {
        for c in 0..mc {
            if c > 0 {
                let _ = write!(out, "\t");
            }
            let val = grid
                .get(&CellAddr::Main {
                    row: r as u32,
                    col: c as u32,
                })
                .unwrap_or("");
            let _ = write!(out, "{}", val);
        }
        let _ = writeln!(out);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ascii_table_trims_empty_margin_columns() {
        let mut g = Grid::new(3, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "Aasdf".into());
        g.set(&CellAddr::Main { row: 2, col: 0 }, "adsf".into());
        let mut out = Vec::new();
        export_ascii_table(&g, &mut out);
        let s = String::from_utf8(out).unwrap();
        assert!(!s.contains("<9"));
        assert!(!s.contains(">9"));
        assert!(s.contains("Aasdf"));
        assert!(s.contains("adsf"));
    }

    #[test]
    fn colwidth_fixture_keeps_column_a_narrow_and_b_wide() {
        use std::path::Path;

        let data = std::fs::read_to_string(Path::new("docs/tests/colwidth.corro")).unwrap();
        let mut workbook = crate::ops::WorkbookState::new();
        let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
        for line in data.lines() {
            let t = line.trim();
            if t.is_empty() {
                continue;
            }
            crate::ops::apply_log_line_to_workbook(t, &mut workbook, &mut active_sheet).unwrap();
        }
        let sheet = workbook.sheet_mut_by_id(active_sheet).unwrap();
        for c in 0..sheet.grid.main_cols() {
            sheet
                .grid
                .fit_column_to_content(crate::grid::MARGIN_COLS + c);
        }

        assert_eq!(sheet.grid.col_width(crate::grid::MARGIN_COLS), 4);
        assert_eq!(sheet.grid.col_width(crate::grid::MARGIN_COLS + 1), 20);
    }

    #[test]
    fn ascii_export_uses_rendered_widths() {
        use std::path::Path;

        let data = std::fs::read_to_string(Path::new("subtotal.corro")).unwrap();
        let mut workbook = crate::ops::WorkbookState::new();
        let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
        for line in data.lines() {
            let t = line.trim();
            if t.is_empty() {
                continue;
            }
            crate::ops::apply_log_line_to_workbook(t, &mut workbook, &mut active_sheet).unwrap();
        }
        let sheet = workbook.sheet_mut_by_id(active_sheet).unwrap();
        for c in 0..sheet.grid.main_cols() {
            sheet.grid.set_col_width(crate::grid::MARGIN_COLS + c, None);
        }
        let mut out = Vec::new();
        export_ascii_table(&sheet.grid, &mut out);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains("TOTAL"));
        assert!(s
            .lines()
            .any(|line| line.starts_with("|") && line.contains("|")));
        assert!(!s.contains("…"));
    }
}
