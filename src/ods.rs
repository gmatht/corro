//! ODS import/export for workbook data.

use crate::addr::excel_column_name;
use crate::export::ExportContent;
use crate::formula::{cell_effective_display, is_formula};
use crate::grid::{CellAddr, GridBox as Grid, HEADER_ROWS, MARGIN_COLS};
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

/// Corro re-import: if the first `table:table-row` uses a `number-rows-repeated` this large,
/// ODF table rows are treated as full logical (header + margins + main + footer). Otherwise, ODF
/// rows are mapped directly onto the main grid (interop, including compact export).
const ODS_GLOBAL_LAYOUT_MIN_FIRST_ROW_REPEAT: u64 = 1_000_000;
/// A single ODF `table:number-rows-repeated` must not make the sheet length exceed what Calc allows.
const LIBREOFFICE_MAX_SHEET_ROWS: usize = 1_048_576;
const CORRO_ODS_LAYOUT_PATH: &str = "corro-ods-layout";

/// How to interpret 0-based ODF table `row_idx` when writing into the Corro grid. Corro exports
/// include a `corro-ods-layout` sidecar; files without it use the first row's
/// [ODS_GLOBAL_LAYOUT_MIN_FIRST_ROW_REPEAT] as before (LibreOffice / interop ODS).
enum OdsTableLayout {
    Odf,
    Rebase { min: usize },
    Compact { physical_to_logical: Vec<usize> },
}

impl OdsTableLayout {
    /// Use full logical (header + margins + main + footer) mapping; see [set_ods_cell].
    fn use_global_ods_addr(&self, odf_inferred: bool) -> bool {
        match self {
            OdsTableLayout::Odf => odf_inferred,
            OdsTableLayout::Rebase { .. } | OdsTableLayout::Compact { .. } => true,
        }
    }

    fn odf_to_logical(&self, odf_table_row: usize) -> usize {
        match self {
            OdsTableLayout::Odf => odf_table_row,
            OdsTableLayout::Rebase { min } => min.saturating_add(odf_table_row),
            OdsTableLayout::Compact {
                physical_to_logical,
            } => physical_to_logical
                .get(odf_table_row)
                .copied()
                .unwrap_or(odf_table_row),
        }
    }
}

fn corro_ods_layout_file_bytes(
    use_compact: bool,
    min_r: usize,
    ordered_logical_rows: &[usize],
) -> String {
    if use_compact {
        let mut s = String::from("v1\ncompact\n");
        for r in ordered_logical_rows {
            use std::fmt::Write;
            let _ = writeln!(s, "{}", r);
        }
        s
    } else {
        format!("v1\nrebase\n{}\n", min_r)
    }
}

fn parse_corro_ods_layout_str(buf: &str) -> OdsTableLayout {
    let lines: Vec<&str> = buf.lines().map(str::trim).filter(|l| !l.is_empty()).collect();
    if lines.first().copied() != Some("v1") {
        return OdsTableLayout::Odf;
    }
    match lines.get(1).copied() {
        Some("rebase") => {
            if let Some(n) = lines.get(2).and_then(|s| s.parse().ok()) {
                OdsTableLayout::Rebase { min: n }
            } else {
                OdsTableLayout::Odf
            }
        }
        Some("compact") => {
            let rows: Vec<usize> = lines
                .get(2..)
                .unwrap_or(&[])
                .iter()
                .filter_map(|l| l.parse().ok())
                .collect();
            if rows.is_empty() {
                OdsTableLayout::Odf
            } else {
                OdsTableLayout::Compact {
                    physical_to_logical: rows,
                }
            }
        }
        _ => OdsTableLayout::Odf,
    }
}

pub fn export_ods_bytes(grid: &Grid) -> Result<Vec<u8>, OdsError> {
    export_ods_bytes_with_options(grid, ExportContent::Formulas)
}

/// `ExportContent::Formulas` preserves `table:formula` where applicable; `Values` writes static cells
/// from computed display (no formulas in the ODF file).
pub fn export_ods_bytes_with_options(
    grid: &Grid,
    content: ExportContent,
) -> Result<Vec<u8>, OdsError> {
    let tc = ods_col_end(grid);
    let rows = ods_row_order(grid);
    let max_rebase_run = ods_max_rebase_blank_run(&rows);
    let use_compact = max_rebase_run >= LIBREOFFICE_MAX_SHEET_ROWS;
    let content_xml = ods_content_xml_with_rows_and_mode(
        grid,
        tc,
        rows.clone(),
        use_compact,
        max_rebase_run,
        content,
    );
    let min_r = *rows
        .first()
        .expect("ods_row_order is never empty");
    let sidecar = corro_ods_layout_file_bytes(use_compact, min_r, &rows);

    let cursor = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(cursor);
    let opt = FileOptions::default().compression_method(zip::CompressionMethod::Stored);

    zip.start_file("mimetype", opt)?;
    zip.write_all(b"application/vnd.oasis.opendocument.spreadsheet")?;

    zip.start_file("content.xml", FileOptions::default())?;
    zip.write_all(content_xml.as_bytes())?;

    zip.start_file(CORRO_ODS_LAYOUT_PATH, FileOptions::default())?;
    zip.write_all(sidecar.as_bytes())?;

    zip.start_file("META-INF/manifest.xml", FileOptions::default())?;
    zip.write_all(ods_manifest_with_corro_layout().as_bytes())?;

    Ok(zip.finish()?.into_inner())
}

pub fn import_ods_workbook(path: &Path) -> Result<WorkbookState, OdsError> {
    let bytes = std::fs::read(path)?;
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes))?;
    let mut content = String::new();
    archive
        .by_name("content.xml")?
        .read_to_string(&mut content)?;
    let mut sidecar = String::new();
    let layout = if let Ok(mut f) = archive.by_name(CORRO_ODS_LAYOUT_PATH) {
        if f.read_to_string(&mut sidecar).is_ok() {
            parse_corro_ods_layout_str(&sidecar)
        } else {
            OdsTableLayout::Odf
        }
    } else {
        OdsTableLayout::Odf
    };
    parse_ods_content_with_layout(&content, layout)
}

fn ods_escape(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

fn ods_manifest_with_corro_layout() -> String {
    String::from(concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>"#,
        r#"
<manifest:file-entry manifest:full-path="corro-ods-layout" manifest:media-type="text/plain"/>"#,
        r#"
</manifest:manifest>"#
    ))
}

fn ods_column_styles_xml(grid: &Grid, tc: usize) -> String {
    let mut s = String::new();
    for c in 0..tc {
        let width_cm = ods_column_width_cm(grid.col_width(c));
        s.push_str(&format!(
            r#"<style:style style:name="co{c}" style:family="table-column"><style:table-column-properties style:column-width="{width_cm:.2}cm"/></style:style>"#
        ));
    }
    s
}

fn ods_column_width_cm(char_width: usize) -> f32 {
    let chars = char_width.max(1) as f32;
    (chars * 0.20 + 0.20).max(0.45)
}

fn ods_col_end(grid: &Grid) -> usize {
    let mut end = grid.total_cols();
    while end > 0 && !ods_logical_col_has_content(grid, end - 1) {
        end -= 1;
    }
    end.max(1)
}

fn ods_logical_col_has_content(grid: &Grid, col: usize) -> bool {
    grid.logical_col_has_content(col)
}

fn ods_row_order(grid: &Grid) -> Vec<usize> {
    let hr = HEADER_ROWS;
    let mr = grid.main_rows();
    let mut rows = Vec::new();
    for (addr, _) in grid.iter_nonempty() {
        let logical_row = match addr {
            CellAddr::Header { row, .. } => row as usize,
            CellAddr::Footer { row, .. } => hr + mr + row as usize,
            CellAddr::Main { row, .. }
            | CellAddr::Left { row, .. }
            | CellAddr::Right { row, .. } => hr + row as usize,
        };
        rows.push(logical_row);
    }
    rows.sort_unstable();
    rows.dedup();
    if rows.is_empty() {
        rows.push(0);
    }
    rows
}

/// Max single `table:number-rows-repeated` in the rebase-anchored (min-row) ODF run we would emit
/// (leading gap + any gap between used logical rows). If that exceeds [LIBREOFFICE_MAX_SHEET_ROWS],
/// we use compact ODF: one ODF table row per used logical row (no large blank run).
fn ods_max_rebase_blank_run(rows: &[usize]) -> usize {
    if rows.is_empty() {
        return 0;
    }
    let min_r = rows[0];
    let mut next_rel = 0usize;
    let mut max_run = 0usize;
    for &r in rows {
        let rel = r.saturating_sub(min_r);
        if rel > next_rel {
            max_run = max_run.max(rel - next_rel);
        }
        next_rel = rel.saturating_add(1);
    }
    max_run
}

fn ods_content_xml_with_rows_and_mode(
    grid: &Grid,
    tc: usize,
    rows: Vec<usize>,
    use_compact: bool,
    _max_rebase_run: usize,
    export_content: ExportContent,
) -> String {
    let column_styles = ods_column_styles_xml(grid, tc);
    let mut s = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" office:version="1.2"><office:automatic-styles>"#,
    );

    s.push_str(&column_styles);
    s.push_str("</office:automatic-styles><office:body><office:spreadsheet><table:table>");

    for c in 0..tc {
        s.push_str(&format!(
            r#"<table:table-column table:style-name="co{c}"/>"#
        ));
    }

    if use_compact {
        for r in &rows {
            s.push_str("<table:table-row>");
            let mut c = 0usize;
            while c < tc {
                s.push_str(&ods_cell_xml(grid, *r, c, export_content));
                c += 1;
            }
            s.push_str("</table:table-row>");
        }
    } else {
        let min_r = *rows
            .first()
            .expect("ods_row_order is never empty");
        let mut next_row = 0usize;
        for r in rows {
            let rel = r.saturating_sub(min_r);
            if rel > next_row {
                let repeated = rel - next_row;
                s.push_str(&format!(
                    r#"<table:table-row table:number-rows-repeated="{repeated}"><table:table-cell table:number-columns-repeated="{tc}"/></table:table-row>"#
                ));
            }
            s.push_str("<table:table-row>");
            let mut c = 0usize;
            while c < tc {
                s.push_str(&ods_cell_xml(grid, r, c, export_content));
                c += 1;
            }
            s.push_str("</table:table-row>");
            next_row = rel + 1;
        }
    }
    s.push_str("</table:table></office:spreadsheet></office:body></office:document-content>");
    s
}

fn ods_cell_xml(
    grid: &Grid,
    logical_row: usize,
    global_col: usize,
    export_content: ExportContent,
) -> String {
    let hr = HEADER_ROWS;
    let mr = grid.main_rows();
    let mc = grid.main_cols();
    let addr = ods_cell_addr(grid, logical_row, global_col);
    let raw = grid.text(&addr);
    let display = cell_effective_display(grid, &addr);

    let value = if logical_row < hr {
        header_formula_or_value(grid, logical_row, global_col, mc)
    } else if logical_row < hr + mr {
        main_formula_or_value(grid, logical_row - hr, global_col, mc)
    } else {
        footer_formula_or_value(grid, logical_row - hr - mr, global_col, mc)
    };

    if value.is_empty() && raw.is_empty() {
        return "<table:table-cell/>".into();
    }
    if export_content == ExportContent::Values {
        return ods_cell_xml_values_only(&display, &value, &raw);
    }
    if value.starts_with('=') || is_formula(&raw) {
        let formula = if value.starts_with('=') { value } else { raw };
        let formula = ods_formula_expr(&formula).unwrap_or(formula);
        let value_attrs = match display.trim().parse::<f64>() {
            Ok(n) => format!(r#" office:value-type="float" office:value="{n}""#),
            Err(_) => r#" office:value-type="string""#.to_string(),
        };
        format!(
            r#"<table:table-cell{} table:formula="of:{}"><text:p>{}</text:p></table:table-cell>"#,
            value_attrs,
            ods_escape(&formula),
            ods_escape(&display)
        )
    } else {
        format!(
            r#"<table:table-cell office:value-type="string"><text:p>{}</text:p></table:table-cell>"#,
            ods_escape(if display.is_empty() { &value } else { &display })
        )
    }
}

/// Static ODF cell: `display` is preferred (evaluated for formulas), then `value` / `raw`.
fn ods_cell_xml_values_only(display: &str, value: &str, raw: &str) -> String {
    let show = if !display.is_empty() {
        display
    } else if !value.is_empty() {
        value
    } else {
        raw
    };
    if show.trim().is_empty() {
        return "<table:table-cell/>".into();
    }
    if let Ok(n) = show.trim().parse::<f64>() {
        return format!(
            r#"<table:table-cell office:value-type="float" office:value="{n}"><text:p>{}</text:p></table:table-cell>"#,
            ods_escape(show)
        );
    }
    format!(
        r#"<table:table-cell office:value-type="string"><text:p>{}</text:p></table:table-cell>"#,
        ods_escape(show)
    )
}

fn ods_cell_addr(grid: &Grid, logical_row: usize, global_col: usize) -> CellAddr {
    let hr = HEADER_ROWS;
    let mr = grid.main_rows();
    let lm = MARGIN_COLS;
    let mc = grid.main_cols();

    if logical_row < hr {
        CellAddr::Header {
            row: logical_row as u32,
            col: global_col as u32,
        }
    } else if logical_row < hr + mr {
        let main_row = (logical_row - hr) as u32;
        if global_col < lm {
            CellAddr::Left {
                col: lm - 1 - global_col,
                row: main_row,
            }
        } else if global_col < lm + mc {
            CellAddr::Main {
                row: main_row,
                col: (global_col - lm) as u32,
            }
        } else {
            CellAddr::Right {
                col: global_col - lm - mc,
                row: main_row,
            }
        }
    } else {
        CellAddr::Footer {
            row: (logical_row - hr - mr) as u32,
            col: global_col as u32,
        }
    }
}

fn ods_formula_expr(raw: &str) -> Option<String> {
    let t = raw.trim();
    let expr = t.strip_prefix('=')?;
    let expr = expr
        .split_once(" -- ")
        .map_or(expr, |(expr, _)| expr)
        .trim();
    if expr.is_empty() {
        None
    } else {
        Some(format!("={expr}"))
    }
}

fn header_formula_or_value(grid: &Grid, row: usize, global_col: usize, main_cols: usize) -> String {
    let base = grid.text(&CellAddr::Header {
        row: row as u32,
        col: global_col as u32,
    });
    if global_col < MARGIN_COLS || global_col >= MARGIN_COLS + main_cols {
        return base;
    }
    if let Some(code) = subtotal_code_for_label(&base) {
        let col = excel_column_name(global_col - MARGIN_COLS);
        return format!("=SUBTOTAL({code};{col}1:{col}{})", grid.main_rows());
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
        let c = lm - 1 - global_col;
        let raw = grid.text(&CellAddr::Left {
            col: c,
            row: main_row as u32,
        });
        if let Some(code) = subtotal_code_for_label(&raw) {
            let start = row_total_block_start(grid, main_row as u32);
            let col = excel_column_name(0);
            return format!("=SUBTOTAL({code};{col}{}:{col}{})", start + 1, main_row + 1);
        }
        return raw;
    }
    if global_col < lm + main_cols {
        let raw = grid.text(&CellAddr::Main {
            row: main_row as u32,
            col: (global_col - lm) as u32,
        });
        if is_formula(&raw) {
            raw
        } else {
            raw
        }
    } else {
        let rc = global_col - lm - main_cols;
        let raw = grid.text(&CellAddr::Right {
            col: rc,
            row: main_row as u32,
        });
        if let Some(code) = subtotal_code_for_label(&raw) {
            return format!(
                "=SUBTOTAL({code};{}1:{}{})",
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
    let raw = grid.text(&CellAddr::Footer {
        row: footer_row as u32,
        col: global_col as u32,
    });
    if let Some(code) = subtotal_code_for_label(&raw) {
        return format!(
            "=SUBTOTAL({code};{}1:{}{})",
            excel_column_name(0),
            excel_column_name(main_cols - 1),
            grid.main_rows()
        );
    }
    raw
}

fn subtotal_code_for_label(raw: &str) -> Option<u8> {
    match raw.trim().to_ascii_uppercase().as_str() {
        "TOTAL" | "SUM" => Some(9),
        "MEAN" | "AVERAGE" | "AVG" => Some(1),
        "COUNT" => Some(2),
        "MAX" | "MAXIMUM" => Some(4),
        "MIN" | "MINIMUM" => Some(5),
        _ => None,
    }
}

fn row_total_block_start(grid: &Grid, current_main_row: u32) -> u32 {
    for candidate in (0..current_main_row).rev() {
        if grid
            .get(&CellAddr::Left {
                col: MARGIN_COLS - 1,
                row: candidate,
            })
            .is_some()
        {
            return candidate + 1;
        }
    }
    0
}

fn parse_ods_content_with_layout(
    xml: &str,
    table_layout: OdsTableLayout,
) -> Result<WorkbookState, OdsError> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut buf = Vec::new();
    let mut workbook = WorkbookState::new();
    workbook.sheets.clear();
    // `new()` left `next_sheet_id` at 2; the first imported sheet would incorrectly get id 2.
    workbook.next_sheet_id = 1;
    let mut current_sheet: Option<SheetRecord> = None;
    let mut odf_uses_global_logical: Option<bool> = None;
    let mut row_idx = 0usize;
    let mut col_idx = 0usize;
    let mut pending_value = String::new();
    let mut pending_formula: Option<String> = None;
    let mut in_p = false;
    let mut in_cell = false;
    let mut current_row_repeat = 1usize;
    let mut table_cell_num_cols = 1usize;

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
                    odf_uses_global_logical = None;
                }
                b"table:table-row" => {
                    col_idx = 0;
                    current_row_repeat = attr_value(&e, b"table:number-rows-repeated")
                        .and_then(|n| n.parse().ok())
                        .unwrap_or(1);
                    if odf_uses_global_logical.is_none() {
                        odf_uses_global_logical = Some(
                            (current_row_repeat as u64) >= ODS_GLOBAL_LAYOUT_MIN_FIRST_ROW_REPEAT,
                        );
                    }
                }
                b"text:p" => {
                    in_p = true;
                    pending_value.clear();
                }
                b"table:table-cell" => {
                    in_cell = true;
                    pending_value.clear();
                    pending_formula = attr_value(&e, b"table:formula");
                    table_cell_num_cols = attr_value(&e, b"table:number-columns-repeated")
                        .and_then(|n| n.parse().ok())
                        .unwrap_or(1)
                        .max(1);
                }
                _ => {}
            },
            Ok(Event::Empty(e)) => match e.name().as_ref() {
                b"table:table-cell" => {
                    let n = attr_value(&e, b"table:number-columns-repeated")
                        .and_then(|n| n.parse().ok())
                        .unwrap_or(1)
                        .max(1);
                    if let Some(sheet) = current_sheet.as_mut() {
                        let g = table_layout
                            .use_global_ods_addr(odf_uses_global_logical == Some(true));
                        let logical = table_layout.odf_to_logical(row_idx);
                        for c_off in 0..n {
                            set_ods_cell(
                                &mut sheet.state,
                                logical,
                                col_idx + c_off,
                                None,
                                "",
                                g,
                            );
                        }
                        col_idx += n;
                    } else {
                        col_idx += n;
                    }
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
                        let n = table_cell_num_cols;
                        let g = table_layout
                            .use_global_ods_addr(odf_uses_global_logical == Some(true));
                        let logical = table_layout.odf_to_logical(row_idx);
                        for c_off in 0..n {
                            set_ods_cell(
                                &mut sheet.state,
                                logical,
                                col_idx + c_off,
                                pending_formula.as_deref(),
                                &pending_value,
                                g,
                            );
                        }
                        col_idx += n;
                    } else {
                        col_idx += table_cell_num_cols;
                    }
                    in_cell = false;
                }
                b"table:table-row" => row_idx += current_row_repeat,
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
    let active_sheet_id = workbook
        .sheets
        .first()
        .map(|s| s.id)
        .unwrap_or(1);
    let snapshot = WorkbookSnapshot {
        next_sheet_id: workbook.next_sheet_id,
        active_sheet_id,
        sheets: workbook.sheets,
        volatile_seed: 0,
    };
    Ok(WorkbookState::from_snapshot(&snapshot))
}

fn set_ods_cell(
    state: &mut SheetState,
    odf_table_row: usize,
    odf_table_col: usize,
    formula: Option<&str>,
    value: &str,
    odf_uses_global_logical: bool,
) {
    if value.is_empty() && formula.is_none() {
        return;
    }
    // Interop: ODF/Calc tables are a plain rectangle — map directly to the main grid. (Do not add
    // HEADER_ROWS / MARGIN_COLS; the old `row < … + main_rows()` check misclassified ODF row 2+
    // as footer while `main_rows` was still 1.)
    if !odf_uses_global_logical {
        let target = CellAddr::Main {
            row: odf_table_row as u32,
            col: odf_table_col as u32,
        };
        if let Some(f) = formula {
            let expr = f.strip_prefix("of:").unwrap_or(f);
            state.grid.set(&target, format!("={}", expr));
        } else {
            state.grid.set(&target, value.to_string());
        }
        return;
    }

    // Corro round-trip export: (row, col) are full logical coordinates (huge index rows, etc.)
    let (row, col) = (odf_table_row, odf_table_col);
    let target = if row < HEADER_ROWS {
        CellAddr::Header {
            row: row as u32,
            col: col as u32,
        }
    } else if row < HEADER_ROWS + state.grid.main_rows() {
        let mr = row - HEADER_ROWS;
        if col < MARGIN_COLS {
            CellAddr::Left {
                col: MARGIN_COLS - 1 - col,
                row: mr as u32,
            }
        } else if col < MARGIN_COLS + state.grid.main_cols() {
            CellAddr::Main {
                row: mr as u32,
                col: (col - MARGIN_COLS) as u32,
            }
        } else {
            CellAddr::Right {
                col: col - MARGIN_COLS - state.grid.main_cols(),
                row: mr as u32,
            }
        }
    } else {
        let fr = row - HEADER_ROWS - state.grid.main_rows();
        CellAddr::Footer {
            row: fr as u32,
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
    grid.iter_nonempty()
        .filter_map(|(addr, _)| match addr {
            CellAddr::Main { row, .. }
            | CellAddr::Left { row, .. }
            | CellAddr::Right { row, .. } => Some(row as usize + 1),
            _ => None,
        })
        .max()
        .unwrap_or_else(|| grid.main_rows())
}

/// Used after import to clip main extent — must be **main** width, not `total_cols()` (which
/// includes both margins) or `set_main_size` will expand the main area by ~1400 columns.
fn ods_col_end_for_sheet(grid: &Grid) -> usize {
    grid.main_cols()
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
    use std::io::{ErrorKind, Read};
    use std::process::Command;
    use tempfile::tempdir;
    use tempfile::NamedTempFile;

    #[test]
    fn export_writes_ods_zip() {
        let mut grid = crate::grid::Grid::new(2, 2);
        grid.set(&CellAddr::Main { row: 0, col: 0 }, "1".into());
        let gb = crate::grid::GridBox::from(grid);
        let bytes = export_ods_bytes(&gb).unwrap();
        assert!(bytes.starts_with(b"PK"));
    }

    #[test]
    fn import_ods_roundtrip_basic_sheet() {
        let mut grid = crate::grid::Grid::new(2, 2);
        grid.set(&CellAddr::Main { row: 0, col: 0 }, "42".into());
        let gb = crate::grid::GridBox::from(grid);
        let bytes = export_ods_bytes(&gb).unwrap();
        let tmp = NamedTempFile::new().unwrap();
        std::fs::write(tmp.path(), bytes).unwrap();
        let workbook = import_ods_workbook(tmp.path()).unwrap();
        assert_eq!(
            workbook
                .active_sheet()
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("42")
        );
    }

    /// LibreOffice/Calc: first `table:table-row` has a small (or no) `number-rows-repeated`; ODF
    /// 0,0 should become corro main (0,0), not a `~N` header row.
    #[test]
    fn import_interop_ods_first_cell_goes_to_main() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.2">
<office:body><office:spreadsheet>
<table:table table:name="Sheet1">
<table:table-row>
<table:table-cell><text:p>hello</text:p></table:table-cell>
</table:table-row>
</table:table>
</office:spreadsheet></office:body>
</office:document-content>
"#;
        let workbook = parse_ods_content_with_layout(xml, OdsTableLayout::Odf).unwrap();
        assert_eq!(
            workbook
                .active_sheet()
                .grid
                .get(&CellAddr::Main { row: 0, col: 0 })
                .as_deref(),
            Some("hello")
        );
    }

    /// ODF row 1+ must stay in the main grid, and `set_main_size` must not use `total_cols` (margins + main).
    #[test]
    fn import_interop_ods_row2_and_few_main_cols() {
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" office:version="1.2">
<office:body><office:spreadsheet>
<table:table table:name="S">
<table:table-row><table:table-cell><text:p>a1</text:p></table:table-cell><table:table-cell><text:p>b1</text:p></table:table-cell></table:table-row>
<table:table-row><table:table-cell><text:p>a2</text:p></table:table-cell><table:table-cell><text:p>b2</text:p></table:table-cell></table:table-row>
</table:table>
</office:spreadsheet></office:body>
</office:document-content>
"#;
        let workbook = parse_ods_content_with_layout(xml, OdsTableLayout::Odf).unwrap();
        let s = workbook.active_sheet();
        assert_eq!(
            s.grid.get(&CellAddr::Main { row: 1, col: 1 }).as_deref(),
            Some("b2")
        );
        assert_eq!(s.grid.main_cols(), 2);
        assert!(s.grid.get(&CellAddr::Footer { row: 0, col: 0 }).is_none());
    }

    #[test]
    fn export_trims_trailing_blank_rows_and_columns() {
        let mut grid = crate::grid::Grid::new(2, 2);
        grid.set(&CellAddr::Main { row: 0, col: 0 }, "42".into());
        let gb = crate::grid::GridBox::from(grid);
        let content = exported_content_xml(&gb);
        assert_eq!(content.matches("<table:table-row>").count(), 1);
        assert!(
            !content.contains(&format!(
                r#"<table:table-row table:number-rows-repeated="{}">"#,
                HEADER_ROWS
            )),
            "a single ~1e9 blank run exceeds Calc row limits; export should rebase"
        );
        let bytes = export_ods_bytes(&gb).unwrap();
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
        let mut sidecar = String::new();
        archive
            .by_name(CORRO_ODS_LAYOUT_PATH)
            .unwrap()
            .read_to_string(&mut sidecar)
            .unwrap();
        assert!(sidecar.contains("rebase"));
        assert_eq!(
            content.matches("<table:table-column").count(),
            MARGIN_COLS + 1
        );
    }

    #[test]
    fn export_converts_total_to_subtotal_formula() {
        let mut grid = crate::grid::Grid::new(1, 1);
        grid.set(
            &CellAddr::Header {
                row: 0,
                col: MARGIN_COLS as u32,
            },
            "TOTAL".into(),
        );
        let gb = crate::grid::GridBox::from(grid);
        let content = exported_content_xml(&gb);
        assert!(content.contains(r#"table:formula="of:=SUBTOTAL(9;A1:A1)""#));
    }

    #[test]
    fn export_translates_other_aggregate_labels() {
        let mut grid = crate::grid::Grid::new(1, 1);
        grid.set(
            &CellAddr::Footer {
                row: 0,
                col: (MARGIN_COLS + 0) as u32,
            },
            "MAX".into(),
        );
        let gb = crate::grid::GridBox::from(grid);
        let content = exported_content_xml(&gb);
        assert!(
            content.contains(r#"table:formula="of:=SUBTOTAL(4;A1:A1)""#),
            "{}",
            content
        );
    }

    #[test]
    fn export_emits_column_styles() {
        let mut grid = crate::grid::Grid::new(2, 2);
        grid.set(&CellAddr::Main { row: 0, col: 0 }, "42".into());
        grid.set(&CellAddr::Left { row: 0, col: 0 }, "X".into());
        let gb = crate::grid::GridBox::from(grid);
        let content = exported_content_xml(&gb);
        assert!(content.contains(r#"style:style style:name="co0" style:family="table-column""#));
        assert!(content.contains(r#"style:column-width=""#));
        assert!(content.contains(r#"table:style-name="co0""#));
    }

    #[test]
    fn subtotal_fixture_roundtrips_through_ods() {
        let workbook = workbook_from_fixture("subtotal.corro");
        let grid = &workbook.active_sheet().grid;
        let content = exported_content_xml(grid);
        assert!(content.contains(r#"table:formula="of:=SUBTOTAL("#));

        let bytes = export_ods_bytes(grid).unwrap();
        let tmp = NamedTempFile::new().unwrap();
        std::fs::write(tmp.path(), bytes).unwrap();
        let reimported = import_ods_workbook(tmp.path()).unwrap();

        let sheet = reimported.active_sheet();
        let mut saw_formula = false;
        let mut saw_subtotal = false;
        scan_grid(&sheet.grid, |value| {
            if value.starts_with('=') {
                saw_formula = true;
            }
            if value.contains("SUBTOTAL") {
                saw_subtotal = true;
            }
        });
        assert!(
            saw_formula,
            "expected at least one formula cell to survive roundtrip"
        );
        assert!(saw_subtotal, "expected subtotal formulas to be preserved");
    }

    #[test]
    fn subtotal_fixture_opens_in_libreoffice_if_available() {
        let Some(soffice) = libreoffice_binary() else {
            return;
        };

        let workbook = workbook_from_fixture("subtotal.corro");
        let grid = &workbook.active_sheet().grid;
        let dir = tempdir().unwrap();
        let ods_path = dir.path().join("subtotal.ods");
        std::fs::write(&ods_path, export_ods_bytes(grid).unwrap()).unwrap();
        let out_dir = dir.path().join("out");
        std::fs::create_dir(&out_dir).unwrap();

        let status = match Command::new(&soffice)
            .args([
                "--headless",
                "--nologo",
                "--nolockcheck",
                "--nodefault",
                "--convert-to",
                "xlsx",
                "--outdir",
                out_dir.to_str().unwrap(),
                ods_path.to_str().unwrap(),
            ])
            .status()
        {
            Ok(status) => status,
            Err(err) if err.kind() == ErrorKind::NotFound => {
                eprintln!("LibreOffice not found; skipping ODS smoke test");
                return;
            }
            Err(err) => panic!("failed to run LibreOffice: {err}"),
        };
        assert!(status.success(), "LibreOffice conversion failed");

        let xlsx = out_dir.join("subtotal.xlsx");
        assert!(xlsx.exists(), "LibreOffice did not write xlsx output");
        let mut archive =
            zip::ZipArchive::new(std::io::Cursor::new(std::fs::read(&xlsx).unwrap())).unwrap();
        let mut sheet_xml = String::new();
        archive
            .by_name("xl/worksheets/sheet1.xml")
            .unwrap()
            .read_to_string(&mut sheet_xml)
            .unwrap();
        assert!(sheet_xml.contains("SUBTOTAL"));
        assert!(sheet_xml.contains("<f"));
    }

    fn workbook_from_fixture(name: &str) -> WorkbookState {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join(name);
        let data = std::fs::read_to_string(path).unwrap();
        let mut workbook = WorkbookState::new();
        let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
        for line in data.lines() {
            let t = line.trim();
            if t.is_empty() {
                continue;
            }
            crate::ops::apply_log_line_to_workbook(t, &mut workbook, &mut active_sheet).unwrap();
        }
        workbook
    }

    fn scan_grid<F: FnMut(&str)>(grid: &Grid, mut f: F) {
        for (_, value) in grid.iter_nonempty() {
            if !value.is_empty() {
                f(&value);
            }
        }
    }

    fn libreoffice_binary() -> Option<String> {
        for name in ["soffice", "libreoffice"] {
            let Ok(output) = Command::new(name).arg("--version").output() else {
                continue;
            };
            if output.status.success() {
                return Some(name.to_string());
            }
        }
        None
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
