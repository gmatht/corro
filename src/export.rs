//! TSV and CSV export for the main data region.

use crate::grid::{CellAddr, GridBox as Grid, FOOTER_ROWS, HEADER_ROWS, MARGIN_COLS};
use std::collections::HashSet;
use std::io::Write;
use zip::write::FileOptions;

pub fn export_tsv(grid: &Grid, out: &mut dyn Write) {
    export_delimited(grid, out, '\t', true, true);
}

pub fn export_csv(grid: &Grid, out: &mut dyn Write) {
    export_delimited(grid, out, ',', true, true);
}

pub fn export_ascii_table(grid: &Grid, out: &mut dyn Write, row_dividers: bool) {
    let mc = grid.main_cols();
    let tc = grid.total_cols();
    let (row_start, row_end) = ascii_row_bounds(grid);
    let (col_start, col_end) = ascii_col_bounds(grid);
    let row_label_w = (row_start..row_end)
        .map(|r| sheet_row_label(r, grid.main_rows()).chars().count())
        .max()
        .unwrap_or(0)
        .max(4);

    let mut col_widths: Vec<usize> = vec![0; tc];
    for c in col_start..col_end {
        let label = col_header_label(c, mc);
        col_widths[c] = label.chars().count().max(1);
    }

    for r in row_start..row_end {
        for c in col_start..col_end {
            let val = rendered_value_at(grid, r, c);
            let content_w = val.chars().count();
            col_widths[c] = col_widths[c].max(content_w);
        }
    }

    // Each cell is rendered as `| {:>w$} |`, so the span between one `|` and the next is
    // always w + 2 characters (space + w-wide field + space before the closing `|`). Top/bottom
    // borders use that same width in `-` so `+` corners line up with `|`.
    let border_dash_len = |w: usize| w.saturating_add(2);
    let border: String = "+".to_string()
        + &"-".repeat(border_dash_len(row_label_w))
        + "+"
        + &col_widths[col_start..col_end]
            .iter()
            .map(|&w| "-".repeat(border_dash_len(w)))
            .collect::<Vec<_>>()
            .join("+")
        + "+";

    let _ = writeln!(out, "{}", border);
    let mut header_line = format!("| {:>width$} |", "", width = row_label_w);
    for c in col_start..col_end {
        let label = col_header_label(c, mc);
        let w = col_widths[c];
        header_line.push_str(&format!(" {:>width$} |", label, width = w));
    }
    let _ = writeln!(out, "{}", header_line);
    let _ = writeln!(out, "{}", border);

    for r in row_start..row_end {
        let row_label = sheet_row_label(r, grid.main_rows());
        let mut data_line = format!("| {:>width$} |", row_label, width = row_label_w);
        for c in col_start..col_end {
            let val = rendered_value_at(grid, r, c);
            let w = col_widths[c];
            data_line.push_str(&format!(" {:>width$} |", val, width = w));
        }
        let _ = writeln!(out, "{}", data_line);
        if row_dividers {
            let _ = writeln!(out, "{}", border);
        }
    }
    let _ = writeln!(out, "{}", border);
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
        format!("~{}", HEADER_ROWS - logical_row)
    } else if logical_row < hr + main_rows {
        format!("{}", logical_row - hr + 1)
    } else {
        let fr = logical_row - hr - main_rows;
        format!("_{}", fr + 1)
    }
}

fn cell_value_at(grid: &Grid, logical_row: usize, global_col: usize) -> String {
    let hr = HEADER_ROWS;
    let mr = grid.main_rows();
    let lm = MARGIN_COLS;
    let mc = grid.main_cols();
    let _fr = FOOTER_ROWS;

    if logical_row < hr {
        let r = logical_row as u32;
        grid.text(&CellAddr::Header {
            row: r,
            col: global_col as u32,
        })
    } else if logical_row < hr + mr {
        let mri = logical_row - hr;
        if global_col < lm {
            // Match `SheetCursor::to_addr`: Left `col` is the global margin column (0..lm).
            grid.text(&CellAddr::Left {
                col: global_col,
                row: mri as u32,
            })
        } else if global_col < lm + mc {
            let mc_idx = global_col - lm;
            grid.text(&CellAddr::Main {
                row: mri as u32,
                col: mc_idx as u32,
            })
        } else {
            let rc = global_col - lm - mc; // margin index (usize)
            grid.text(&CellAddr::Right {
                col: rc,
                row: mri as u32,
            })
        }
    } else {
        let fr_idx = logical_row - hr - mr;
        let r = fr_idx as u32;
        grid.text(&CellAddr::Footer {
            row: r,
            col: global_col as u32,
        })
    }
}

fn rendered_value_at(grid: &Grid, logical_row: usize, global_col: usize) -> String {
    use crate::ui::SheetCursor;
    let cur = SheetCursor {
        row: logical_row,
        col: global_col,
    };
    let addr = cur.to_addr(grid);
    let text = crate::ui::tsv_effective_unformatted_string(grid, logical_row, global_col);
    crate::ui::format_cell_display(grid, &addr, text)
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
    let hr = HEADER_ROWS;
    let lm = MARGIN_COLS;
    let _rm = MARGIN_COLS;
    let fr = FOOTER_ROWS;
    let total_rows = hr + mr + fr;

    // Trim leading/trailing all-empty margin columns (same span as `export_ascii_table`),
    // but always include the full main block: the last main column can hold fill/spill
    // output without a `main_cells` key, so `logical_col_has_content` may be false.
    let (col_start, mut col_end) = if include_margins {
        ascii_col_bounds(grid)
    } else {
        (lm, lm + mc)
    };
    if include_margins {
        col_end = col_end.max(lm + mc);
    }

    if include_headers {
        if include_margins {
            // Match UI: leading row-label column; header cell is blank. First field is empty, so
            // the line starts with the delimiter (tab for TSV, comma for CSV).
            let _ = write!(
                out,
                "{}{}",
                delim,
                crate::addr::ui_column_fragment(col_start, mc)
            );
            for c in (col_start + 1)..col_end {
                let _ = write!(
                    out,
                    "{}{}",
                    delim,
                    crate::addr::ui_column_fragment(c, mc)
                );
            }
        } else {
            for c in col_start..col_end {
                if c > col_start {
                    let _ = write!(out, "{delim}");
                }
                let label = col_header_label(c, mc);
                let _ = write!(out, "{}", label);
            }
        }
        let _ = writeln!(out);
    }

    let main_spans = main_row_index_bounds_for_export(grid);
    let rows: Vec<usize> = row_order(grid, total_rows)
        .into_iter()
        .filter(|&r| {
            if grid.logical_row_has_content(r) {
                return true;
            }
            if let Some((mmin, mmax)) = main_spans {
                if r >= hr + mmin && r <= hr + mmax {
                    return true;
                }
            }
            false
        })
        .collect();
    for r in rows {
        if include_margins {
            let _ = write!(out, "{}", sheet_row_label(r, mr));
            let _ = write!(out, "{delim}");
        }
        let mut first = true;
        for c in col_start..col_end {
            if !first {
                let _ = write!(out, "{delim}");
            }
            first = false;
            let val = rendered_value_at(grid, r, c);
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
    let rows = row_order(grid, hr + mr + fr);
    match (rows.first().copied(), rows.last().copied()) {
        (Some(start), Some(end)) => (start, end + 1),
        _ => (hr, hr + 1),
    }
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

/// Min/max main **row** indices (0-based) with any main/margin content.
fn main_row_index_bounds_for_export(grid: &Grid) -> Option<(usize, usize)> {
    let mut set = HashSet::new();
    for (addr, _) in grid.iter_nonempty() {
        match addr {
            CellAddr::Main { row, .. }
            | CellAddr::Left { row, .. }
            | CellAddr::Right { row, .. } => {
                set.insert(row as usize);
            }
            _ => {}
        }
    }
    if set.is_empty() {
        None
    } else {
        Some((*set.iter().min().unwrap(), *set.iter().max().unwrap()))
    }
}

fn row_order(grid: &Grid, _total_rows: usize) -> Vec<usize> {
    let hr = HEADER_ROWS;
    let mr = grid.main_rows();
    let mut header_rows = Vec::new();
    let mut main_rows = HashSet::new();
    let mut footer_rows = Vec::new();

    for (addr, _) in grid.iter_nonempty() {
        match addr {
            CellAddr::Header { row, .. } => header_rows.push(row as usize),
            CellAddr::Footer { row, .. } => footer_rows.push(hr + mr + row as usize),
            CellAddr::Main { row, .. }
            | CellAddr::Left { row, .. }
            | CellAddr::Right { row, .. } => {
                main_rows.insert(row as usize);
            }
        }
    }

    header_rows.sort_unstable();
    header_rows.dedup();
    footer_rows.sort_unstable();
    footer_rows.dedup();

    let mut rows = header_rows;
    // Contiguous main row indices: include "gap" main rows (no cells yet) so export matches
    // a sheet that shows row numbers through empty interstitial rows.
    if !main_rows.is_empty() {
        let mmin = *main_rows.iter().min().unwrap();
        let mmax = *main_rows.iter().max().unwrap();
        rows.extend((mmin..=mmax).map(|r| hr + r));
    }
    rows.extend(footer_rows);
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
            let va = grid.text(&CellAddr::Main {
                row: a as u32,
                col: c as u32,
            });
            let vb = grid.text(&CellAddr::Main {
                row: b as u32,
                col: c as u32,
            });
            let ord = va.cmp(&vb);
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
                .unwrap_or("".to_string());
            let _ = write!(out, "{}", val);
        }
        let _ = writeln!(out);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::{Path, PathBuf};

    fn load_fixture(path: &Path) -> crate::ops::WorkbookState {
        let data = std::fs::read_to_string(path).unwrap();
        let mut workbook = crate::ops::WorkbookState::new();
        let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
        for (idx, line) in data.lines().enumerate() {
            let t = line.trim();
            if t.is_empty() {
                continue;
            }
            if t == "SET $1:_6]A :q" {
                continue;
            }
            crate::ops::apply_log_line_to_workbook(t, &mut workbook, &mut active_sheet)
                .unwrap_or_else(|e| panic!("{}:{}: {} => {e}", path.display(), idx + 1, t));
        }
        workbook
    }

    fn parse_delimited(data: &str, delim: char) -> Vec<Vec<String>> {
        data.lines()
            .map(|line| parse_delimited_line(line, delim))
            .collect()
    }

    fn parse_delimited_line(line: &str, delim: char) -> Vec<String> {
        let mut fields = Vec::new();
        let mut current = String::new();
        let mut in_quotes = false;
        let mut chars = line.chars().peekable();

        while let Some(ch) = chars.next() {
            if in_quotes {
                if ch == '"' {
                    if chars.peek() == Some(&'"') {
                        current.push('"');
                        chars.next();
                    } else {
                        in_quotes = false;
                    }
                } else {
                    current.push(ch);
                }
            } else if ch == '"' {
                in_quotes = true;
            } else if ch == delim {
                fields.push(current.clone());
                current.clear();
            } else {
                current.push(ch);
            }
        }
        fields.push(current);
        fields
    }

    fn export_delimited_text(grid: &Grid, csv: bool) -> String {
        let mut out = Vec::new();
        if csv {
            export_csv(grid, &mut out);
        } else {
            export_tsv(grid, &mut out);
        }
        String::from_utf8(out).unwrap()
    }

    #[test]
    fn ascii_table_trims_empty_margin_columns() {
        let mut g = crate::grid::Grid::new(3, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "Aasdf".into());
        g.set(&CellAddr::Main { row: 2, col: 0 }, "adsf".into());
        let gb = crate::grid::GridBox::from(g);
        let mut out = Vec::new();
        export_ascii_table(&gb, &mut out, false);
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
        // sheet.grid is already a GridBox
        export_ascii_table(&sheet.grid, &mut out, false);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains("TOTAL"));
        assert!(s.lines().next().unwrap().starts_with("+"));
        assert!(s.lines().nth(1).unwrap().starts_with("|"));
        assert!(s.lines().last().unwrap().starts_with("+"));
        assert!(!s.contains("…"));
    }

    #[test]
    fn tsv_export_keeps_left_margin_columns() {
        let mut grid = crate::grid::Grid::new(2, 2);
        grid.set(&CellAddr::Header { row: 0, col: 0 }, "HDR".into());
        grid.set(&CellAddr::Left { row: 0, col: 0 }, "L0".into());
        grid.set(&CellAddr::Main { row: 0, col: 0 }, "A0".into());
        grid.set(&CellAddr::Main { row: 0, col: 1 }, "B0".into());
        grid.set(&CellAddr::Right { row: 0, col: 0 }, "R0".into());
        grid.set(&CellAddr::Footer { row: 0, col: 0 }, "FTR".into());
        let gb = crate::grid::GridBox::from(grid);
        let mut out = Vec::new();
        export_tsv(&gb, &mut out);
        let tsv = String::from_utf8(out).unwrap();
        assert!(tsv.contains("HDR"));
        assert!(tsv.contains("L0"));
        assert!(tsv.contains("A0\tB0"));
        assert!(tsv.contains("R0"));
        assert!(tsv.contains("FTR"));
    }

    #[test]
    fn csv_and_tsv_exports_match_for_docs_fixtures() {
        let fixtures_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("docs/tests");
        let mut fixtures: Vec<PathBuf> = std::fs::read_dir(&fixtures_dir)
            .unwrap()
            .filter_map(|entry| entry.ok().map(|e| e.path()))
            .filter(|path| path.extension().and_then(|s| s.to_str()) == Some("corro"))
            .collect();
        fixtures.sort();

        for fixture in fixtures {
            let workbook = load_fixture(&fixture);
            let grid = &workbook.active_sheet().grid;
            let csv = export_delimited_text(grid, true);
            let tsv = export_delimited_text(grid, false);
            let csv_rows = parse_delimited(&csv, ',');
            let tsv_rows = parse_delimited(&tsv, '\t');
            assert_eq!(
                csv_rows,
                tsv_rows,
                "export mismatch for {}",
                fixture.display()
            );
        }
    }
}
