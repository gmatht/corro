//! ODS import/export for workbook data.

use crate::addr::excel_column_name;
use crate::formula::is_formula;
use crate::grid::{CellAddr, Grid, FOOTER_ROWS, HEADER_ROWS, MARGIN_COLS};
use crate::ops::{SheetRecord, SheetState, WorkbookSnapshot, WorkbookState};
use quick_xml::events::Event;
use quick_xml::Reader;
use std::io::{Read, Write};
use std::path::Path;
use thiserror::Error;
use zip::write::FileOptions;

#[derive(Debug, Error)]
pub enum OdsError {
    #[error("ODS: {0}")]
    Io(#[from] std::io::Error),
    #[error("ODS XML: {0}")]
    Xml(String),
    #[error("ODS archive: {0}")]
    Zip(#[from] zip::result::ZipError),
}

pub fn export_ods_bytes(grid: &Grid) -> Result<Vec<u8>, OdsError> {
    let cursor = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let opt = FileOptions::default().compression_method(zip::CompressionMethod::Stored);

    zip.start_file("mimetype", opt)?;
    zip.write_all(b"application/vnd.oasis.opendocument.spreadsheet")?;

    zip.start_file("content.xml", FileOptions::default())?;
    zip.write_all(ods_content_xml(grid).as_bytes())?;

    zip.start_file("META-INF/manifest.xml", FileOptions::default())?;
    zip.write_all(ods_manifest_xml().as_bytes())?;

    Ok(zip.finish()?.into_inner())
}

pub fn import_ods_workbook(path: &Path) -> Result<WorkbookState, OdsError> {
    let bytes = std::fs::read(path)?;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes))?;
    let mut content = String::new();
    archive
        .by_name("content.xml")?
        .read_to_string(&mut content)?;
    parse_ods_content(&content)
}

fn ods_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn ods_manifest_xml() -> String {
    String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>"#,
    )
}

fn ods_content_xml(grid: &Grid) -> String {
    let tc = ods_col_end(grid);
    let row_end = ods_row_end(grid);
    let mut s = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" office:version="1.2"><office:body><office:spreadsheet><table:table>"#,
    );

    for _ in 0..tc {
        s.push_str("<table:table-column table:number-columns-repeated=\"1\"/>");
    }

    for r in 0..row_end {
        s.push_str("<table:table-row>");
        let mut c = 0usize;
        while c < tc {
            let global_col = c;
            let cell = ods_cell_xml(grid, r, global_col);
            s.push_str(&cell);
            c += 1;
        }
        s.push_str("</table:table-row>");
    }
    s.push_str("</table:table></office:spreadsheet></office:body></office:document-content>");
    s
}

fn ods_row_end(grid: &Grid) -> usize {
    let mut end = HEADER_ROWS + grid.main_rows() + FOOTER_ROWS;
    while end > 0 && !grid.logical_row_has_content(end - 1) {
        end -= 1;
    }
    end.max(1)
}

fn ods_col_end(grid: &Grid) -> usize {
    let mut end = grid.total_cols();
    while end > 0 && !grid.logical_col_has_content(end - 1) {
        end -= 1;
    }
    end.max(1)
}

fn ods_cell_xml(grid: &Grid, logical_row: usize, global_col: usize) -> String {
    let hr = HEADER_ROWS;
    let mr = grid.main_rows();
    let lm = MARGIN_COLS;
    let mc = grid.main_cols();

    let value = if logical_row < hr {
        header_formula_or_value(grid, logical_row, global_col, mc)
    } else if logical_row < hr + mr {
        main_formula_or_value(grid, logical_row - hr, global_col, mc)
    } else {
        footer_formula_or_value(grid, logical_row - hr - mr, global_col, mc)
    };

    if value.is_empty() {
        "<table:table-cell/>".into()
    } else if value.starts_with('=') {
        format!(
            r#"<table:table-cell office:value-type="string" table:formula="of:{}"><text:p>{}</text:p></table:table-cell>"#,
            ods_escape(&value),
            ods_escape(&value)
        )
    } else {
        format!(
            r#"<table:table-cell office:value-type="string"><text:p>{}</text:p></table:table-cell>"#,
            ods_escape(&value)
        )
    }
}

fn header_formula_or_value(grid: &Grid, row: usize, global_col: usize, main_cols: usize) -> String {
    let base = grid
        .get(&CellAddr::Header {
            row: row as u8,
            col: global_col as u32,
        })
        .unwrap_or("")
        .to_string();
    if global_col < MARGIN_COLS || global_col >= MARGIN_COLS + main_cols {
        return base;
    }
    if base.trim().eq_ignore_ascii_case("TOTAL") {
        let col = excel_column_name(global_col - MARGIN_COLS);
        return format!("=SUBTOTAL(9;{col}1:{col}{})", grid.main_rows());
    }
    base
}

fn main_formula_or_value(
    grid: &Grid,
    main_row: usize,
    global_col: usize,
    main_cols: usize,
) -> String {
    let lm = MARGIN_COLS;
    let mr = grid.main_rows();
    if global_col < lm {
        let c = (lm - 1 - global_col) as u8;
        let raw = grid
            .get(&CellAddr::Left {
                col: c,
                row: main_row as u32,
            })
            .unwrap_or("")
            .to_string();
        if raw.trim().eq_ignore_ascii_case("TOTAL") {
            let start = row_total_block_start(grid, main_row as u32);
            let col = excel_column_name(0);
            return format!("=SUBTOTAL(9;{col}{}:{col}{})", start + 1, main_row + 1);
        }
        return raw;
    }
    if global_col < lm + main_cols {
        let raw = grid
            .get(&CellAddr::Main {
                row: main_row as u32,
                col: (global_col - lm) as u32,
            })
            .unwrap_or("")
            .to_string();
        if is_formula(&raw) {
            raw
        } else {
            raw
        }
    } else {
        let rc = (global_col - lm - main_cols) as u8;
        let raw = grid
            .get(&CellAddr::Right {
                col: rc,
                row: main_row as u32,
            })
            .unwrap_or("")
            .to_string();
        if raw.trim().eq_ignore_ascii_case("TOTAL") {
            return format!(
                "=SUBTOTAL(9;{}1:{}{})",
                excel_column_name(0),
                excel_column_name(main_cols - 1),
                mr
            );
        }
        raw
    }
}

fn footer_formula_or_value(
    grid: &Grid,
    footer_row: usize,
    global_col: usize,
    main_cols: usize,
) -> String {
    let raw = grid
        .get(&CellAddr::Footer {
            row: footer_row as u8,
            col: global_col as u32,
        })
        .unwrap_or("")
        .to_string();
    if raw.trim().eq_ignore_ascii_case("TOTAL") {
        return format!(
            "=SUBTOTAL(9;{}1:{}{})",
            excel_column_name(0),
            excel_column_name(main_cols - 1),
            grid.main_rows()
        );
    }
    raw
}

fn row_total_block_start(grid: &Grid, current_main_row: u32) -> u32 {
    for candidate in (0..current_main_row).rev() {
        if grid
            .get(&CellAddr::Left {
                col: (MARGIN_COLS - 1) as u8,
                row: candidate,
            })
            .is_some()
        {
            return candidate + 1;
        }
    }
    0
}

fn parse_ods_content(xml: &str) -> Result<WorkbookState, OdsError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut workbook = WorkbookState::new();
    workbook.sheets.clear();
    let mut current_sheet: Option<SheetRecord> = None;
    let mut row_idx = 0usize;
    let mut col_idx = 0usize;
    let mut pending_value = String::new();
    let mut pending_formula: Option<String> = None;
    let mut in_p = false;
    let mut in_cell = false;

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(e)) => match e.name().as_ref() {
                b"table:table" => {
                    let title = attr_value(&e, b"table:name").unwrap_or_else(|| "Sheet1".into());
                    current_sheet = Some(SheetRecord {
                        id: workbook.next_sheet_id,
                        title,
                        state: SheetState::new(1, 1),
                    });
                    workbook.next_sheet_id += 1;
                    row_idx = 0;
                }
                b"table:table-row" => {
                    col_idx = 0;
                }
                b"text:p" => {
                    in_p = true;
                    pending_value.clear();
                }
                b"table:table-cell" => {
                    in_cell = true;
                    pending_value.clear();
                    pending_formula = attr_value(&e, b"table:formula");
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table-cell" => {
                    if let Some(sheet) = current_sheet.as_mut() {
                        set_ods_cell(&mut sheet.state, row_idx, col_idx, None, "");
                    }
                    col_idx += 1;
                }
                _ => {}
            },
            Ok(Event::Text(t)) => {
                if in_p || in_cell {
                    pending_value.push_str(&String::from_utf8_lossy(t.as_ref()));
                }
            }
            Ok(Event::End(e)) => match e.name().as_ref() {
                b"text:p" => in_p = false,
                b"table:table-cell" => {
                    if let Some(sheet) = current_sheet.as_mut() {
                        set_ods_cell(
                            &mut sheet.state,
                            row_idx,
                            col_idx,
                            pending_formula.as_deref(),
                            &pending_value,
                        );
                    }
                    in_cell = false;
                    col_idx += 1;
                }
                b"table:table-row" => row_idx += 1,
                b"table:table" => {
                    if let Some(sheet) = current_sheet.take() {
                        workbook.sheets.push(sheet);
                    }
                }
                _ => {}
            },
            Ok(Event::Eof) => break,
            Err(e) => return Err(OdsError::Xml(e.to_string())),
            _ => {}
        }
        buf.clear();
    }

    if workbook.sheets.is_empty() {
        return Err(OdsError::Xml("no sheets found".into()));
    }
    for sheet in &mut workbook.sheets {
        let rows = ods_row_end_for_sheet(&sheet.state.grid);
        let cols = ods_col_end_for_sheet(&sheet.state.grid);
        sheet.state.grid.set_main_size(rows.max(1), cols.max(1));
    }
    let snapshot = WorkbookSnapshot {
        next_sheet_id: workbook.next_sheet_id,
        active_sheet_id: 1,
        sheets: workbook.sheets,
        volatile_seed: 0,
    };
    Ok(WorkbookState::from_snapshot(&snapshot))
}

fn set_ods_cell(
    state: &mut SheetState,
    row: usize,
    col: usize,
    formula: Option<&str>,
    value: &str,
) {
    if value.is_empty() && formula.is_none() {
        return;
    }
    let target = if row < HEADER_ROWS {
        CellAddr::Header {
            row: row as u8,
            col: col as u32,
        }
    } else if row < HEADER_ROWS + state.grid.main_rows() {
        let mr = row - HEADER_ROWS;
        if col < MARGIN_COLS {
            CellAddr::Left {
                col: (MARGIN_COLS - 1 - col) as u8,
                row: mr as u32,
            }
        } else if col < MARGIN_COLS + state.grid.main_cols() {
            CellAddr::Main {
                row: mr as u32,
                col: (col - MARGIN_COLS) as u32,
            }
        } else {
            CellAddr::Right {
                col: (col - MARGIN_COLS - state.grid.main_cols()) as u8,
                row: mr as u32,
            }
        }
    } else {
        let fr = row - HEADER_ROWS - state.grid.main_rows();
        CellAddr::Footer {
            row: fr as u8,
            col: col as u32,
        }
    };
    if let Some(f) = formula {
        let expr = f.strip_prefix("of:").unwrap_or(f);
        state.grid.set(&target, format!("={}", expr));
    } else {
        state.grid.set(&target, value.to_string());
    }
}

fn ods_row_end_for_sheet(grid: &Grid) -> usize {
    HEADER_ROWS + grid.main_rows() + FOOTER_ROWS
}

fn ods_col_end_for_sheet(grid: &Grid) -> usize {
    grid.total_cols()
}

fn attr_value(e: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Option<String> {
    for a in e.attributes().flatten() {
        if a.key.as_ref() == key {
            return Some(String::from_utf8_lossy(a.value.as_ref()).into_owned());
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::NamedTempFile;

    #[test]
    fn export_writes_ods_zip() {
        let mut grid = Grid::new(2, 2);
        grid.set(&CellAddr::Main { row: 0, col: 0 }, "1".into());
        let bytes = export_ods_bytes(&grid).unwrap();
        assert!(bytes.starts_with(b"PK"));
    }

    #[test]
    fn import_ods_roundtrip_basic_sheet() {
        let mut grid = Grid::new(2, 2);
        grid.set(&CellAddr::Main { row: 0, col: 0 }, "42".into());
        let bytes = export_ods_bytes(&grid).unwrap();
        let tmp = NamedTempFile::new().unwrap();
        std::fs::write(tmp.path(), bytes).unwrap();
        let workbook = import_ods_workbook(tmp.path()).unwrap();
        assert_eq!(
            workbook
                .active_sheet()
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 }),
            Some("42")
        );
    }

    #[test]
    fn export_trims_trailing_blank_rows_and_columns() {
        let mut grid = Grid::new(2, 2);
        grid.set(&CellAddr::Main { row: 0, col: 0 }, "42".into());
        let content = exported_content_xml(&grid);
        assert_eq!(content.matches("<table:table-row>").count(), 27);
        assert_eq!(content.matches("<table:table-column").count(), 11);
    }

    #[test]
    fn export_converts_total_to_subtotal_formula() {
        let mut grid = Grid::new(1, 1);
        grid.set(
            &CellAddr::Header {
                row: 0,
                col: MARGIN_COLS as u32,
            },
            "TOTAL".into(),
        );
        let content = exported_content_xml(&grid);
        assert!(content.contains(r#"table:formula="of:=SUBTOTAL(9;A1:A1)""#));
    }

    fn exported_content_xml(grid: &Grid) -> String {
        let bytes = export_ods_bytes(grid).unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut content = String::new();
        archive
            .by_name("content.xml")
            .unwrap()
            .read_to_string(&mut content)
            .unwrap();
        content
    }
}
