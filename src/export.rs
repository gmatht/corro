//! TSV and CSV export for the main data region.

use crate::formula;
use crate::grid::{CellAddr, GridBox as Grid, FOOTER_ROWS, HEADER_ROWS, MARGIN_COLS};
use std::collections::HashSet;
use std::io::Write;
use zip::write::FileOptions;

/// Whether delimited/ASCII/selection export emits computed display text or stored cell text (`=…`).
#[derive(Clone, Copy, Debug, Eq, PartialEq, Default)]
pub enum ExportContent {
    /// Evaluated + formatted (matches the main grid / TSV golden files).
    #[default]
    Values,
    /// Raw storage: formula text where present, labels as stored.
    Formulas,
    /// Labeled ` -- ` column headers, rows use interop `=…` (comma-separated args; see plan).
    Generic,
}

/// Options for tab/comma (and "export all") delimited text.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct DelimitedExportOptions {
    /// If true, emit a first line of column names (`A` / `[A` / `]B` with margins, etc.).
    pub include_header_row: bool,
    /// If true, include margin/header/footer *regions*; if false, main block only.
    pub include_margins: bool,
    /// If true, prefix each data line with the sheet row label and the delimiter, and when
    /// `include_header_row` the header has an empty first field for that row-address column (if the
    /// width style is margin-style, or a leading delimiter for main-block-only export). If false,
    /// data starts with the first data column. Independent of `include_margins` (main-only export
    /// can still show row `1`/`2` in the first field). Same idea as
    /// `AsciiTableOptions::include_row_label_column`.
    pub include_row_label_column: bool,
    /// Computed display vs stored `=…` text; see [ExportContent].
    pub content: ExportContent,
}

impl Default for DelimitedExportOptions {
    fn default() -> Self {
        Self {
            include_header_row: true,
            include_margins: true,
            include_row_label_column: true,
            content: ExportContent::default(),
        }
    }
}

pub fn export_tsv(grid: &Grid, out: &mut dyn Write) {
    export_tsv_with_options(grid, out, &DelimitedExportOptions::default());
}

pub fn export_tsv_with_options(
    grid: &Grid,
    out: &mut dyn Write,
    options: &DelimitedExportOptions,
) {
    export_delimited(grid, out, '\t', options);
}

pub fn export_csv(grid: &Grid, out: &mut dyn Write) {
    export_csv_with_options(grid, out, &DelimitedExportOptions::default());
}

pub fn export_csv_with_options(
    grid: &Grid,
    out: &mut dyn Write,
    options: &DelimitedExportOptions,
) {
    export_delimited(grid, out, ',', options);
}

/// Pad/truncate to `width` by Unicode scalar values; right-align (pads on the left) with `pad`.
fn ascii_field(s: &str, width: usize, pad: char) -> String {
    if width == 0 {
        return String::new();
    }
    let w = s.chars().count();
    if w > width {
        s.chars().rev().take(width).collect::<String>().chars().rev().collect()
    } else {
        let n = width - w;
        std::iter::repeat(pad)
            .take(n)
            .chain(s.chars())
            .collect()
    }
}

/// Space (U+0020) vs em space (U+2003) for padding between ASCII table pipes.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Default)]
pub enum AsciiInterCellSpace {
    #[default]
    Space,
    EmSpace,
}

/// Rule for the line between the column-label row and the first data row.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Default)]
pub enum AsciiHeaderDataSeparator {
    /// A full `+---+` line under the column labels (default when a label row is present).
    #[default]
    FullBorder,
    /// No border between label row and data; data follows the label line immediately.
    None,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct AsciiTableOptions {
    /// When false, export only the main data block: rows `hr..hr+main_rows`, cols A..Z (main) — no
    /// margin/header/footer cell columns or rows in the text table.
    pub include_margins: bool,
    /// Draw extra rules in the main block: a horizontal line of `=` above and below the main row
    /// range, and `+` at the left and right of the main block on each main data row to meet those
    /// lines (in addition to the normal outer `+---+` table border).
    pub data_frame: bool,
    /// First column: sheet row labels (`1`, `2`, `~1`, …). When false, the table starts with
    /// `| A | B |` (or margin columns) with no left gutter. Distinct from [`Self::include_column_label_row`].
    pub include_row_label_column: bool,
    pub include_column_label_row: bool,
    pub row_dividers: bool,
    pub inter_cell_space: AsciiInterCellSpace,
    pub header_data_separator: AsciiHeaderDataSeparator,
    /// Same meaning as [DelimitedExportOptions::content].
    pub content: ExportContent,
}

impl Default for AsciiTableOptions {
    fn default() -> Self {
        Self {
            include_margins: true,
            data_frame: false,
            include_row_label_column: true,
            include_column_label_row: true,
            row_dividers: false,
            inter_cell_space: AsciiInterCellSpace::Space,
            header_data_separator: AsciiHeaderDataSeparator::FullBorder,
            content: ExportContent::default(),
        }
    }
}

fn ascii_pre(opts: &AsciiTableOptions) -> char {
    match opts.inter_cell_space {
        AsciiInterCellSpace::Space => ' ',
        AsciiInterCellSpace::EmSpace => '\u{2003}',
    }
}

/// Append one cell: `pre` + right-aligned `text` in `w` (pad) + pre + `|`.
fn ascii_push_cell(s: &mut String, pre: char, pad: char, text: &str, w: usize) {
    s.push(pre);
    s.push_str(&ascii_field(text, w, pad));
    s.push(pre);
    s.push('|');
}

/// `+---...---+` — optional row-label run (first block) is `-`. When `use_equals_in_main`, column
/// runs for `c in main_c0..main_c1` use `=`, which matches the `data_frame` inner horizontals.
fn ascii_border_line(
    with_row_gutter: bool,
    col_start: usize,
    col_end: usize,
    row_label_w: usize,
    col_widths: &[usize],
    main_c0: usize,
    main_c1: usize,
    use_equals_in_main: bool,
) -> String {
    let border_dash_len = |w: usize| w.saturating_add(2);
    let mut s = String::new();
    s.push('+');
    if with_row_gutter {
        s.push_str(&"-".repeat(border_dash_len(row_label_w)));
        s.push('+');
    }
    for c in col_start..col_end {
        let w = col_widths[c];
        let dch = if use_equals_in_main && (main_c0..main_c1).contains(&c) {
            '='
        } else {
            '-'
        };
        s.push_str(
            &std::iter::repeat(dch)
                .take(border_dash_len(w))
                .collect::<String>(),
        );
        s.push('+');
    }
    s
}

pub fn export_ascii_table_with_options(
    grid: &Grid,
    out: &mut dyn Write,
    options: &AsciiTableOptions,
) {
    let mc = grid.main_cols();
    let tc = grid.total_cols();
    let m = MARGIN_COLS;
    let hr = HEADER_ROWS;
    let mr = grid.main_rows();

    let (mut row_start, mut row_end) = ascii_row_bounds(grid);
    let (mut col_start, mut col_end) = ascii_col_bounds(grid);

    if !options.include_margins {
        col_start = m;
        col_end = (m + mc).min(grid.total_cols());
        col_end = col_end.max(col_start.saturating_add(1));
        row_start = hr;
        row_end = hr + mr;
    }

    let cell_content = options.content;

    let main_c0 = m.max(col_start);
    let main_c1 = (m + mc).min(col_end);
    let frame_active = options.data_frame && main_c0 < main_c1;

    let first_main_r = (row_start..row_end)
        .find(|&r| (hr..hr + mr).contains(&r));
    let last_main_r = (row_start..row_end)
        .rfind(|&r| (hr..hr + mr).contains(&r));

    let generic_rebase = if cell_content == ExportContent::Generic {
        let (dr, dc) = ascii_generic_rebase(
            col_start,
            col_end,
            row_start,
            row_end,
            first_main_r,
            last_main_r,
            frame_active,
            options,
        );
        Some((dr, dc))
    } else {
        None
    };

    let row_label_w = if options.include_row_label_column {
        (row_start..row_end)
            .map(|r| sheet_row_label(r, grid.main_rows()).chars().count())
            .max()
            .unwrap_or(0)
            .max(4)
    } else {
        0
    };
    let with_row_gutter = options.include_row_label_column;

    let mut col_widths: Vec<usize> = vec![0; tc];
    for c in col_start..col_end {
        let label = col_header_label_for_export(grid, c, mc, cell_content);
        col_widths[c] = label.chars().count().max(1);
    }

    for r in row_start..row_end {
        for c in col_start..col_end {
            let val = export_cell_text(grid, r, c, cell_content, generic_rebase);
            let content_w = val.chars().count();
            col_widths[c] = col_widths[c].max(content_w);
        }
    }

    // Each cell is rendered as `| {:>w$} |`, so the span between one `|` and the next is
    // always w + 2 characters (space + w-wide field + space before the closing `|`). Top/bottom
    // borders use that same width in `-` so `+` corners line up with `|`.
    let border: String = ascii_border_line(
        with_row_gutter,
        col_start,
        col_end,
        row_label_w,
        &col_widths,
        main_c0,
        main_c1,
        false,
    );
    let frame_line = if frame_active {
        Some(ascii_border_line(
            with_row_gutter,
            col_start,
            col_end,
            row_label_w,
            &col_widths,
            main_c0,
            main_c1,
            true,
        ))
    } else {
        None
    };

    let pre = ascii_pre(options);
    let pad = pre;

    let _ = writeln!(out, "{}", border);

    if options.include_column_label_row {
        let mut header_line = String::new();
        header_line.push('|');
        if with_row_gutter {
            ascii_push_cell(&mut header_line, pre, pad, "", row_label_w);
        }
        for c in col_start..col_end {
            let label = col_header_label_for_export(grid, c, mc, cell_content);
            let w = col_widths[c];
            ascii_push_cell(&mut header_line, pre, pad, &label, w);
        }
        let _ = writeln!(out, "{}", header_line);
        if matches!(options.header_data_separator, AsciiHeaderDataSeparator::FullBorder) {
            let _ = writeln!(out, "{}", border);
        }
    }

    for r in row_start..row_end {
        if frame_active && Some(r) == first_main_r {
            if let Some(ref fl) = frame_line {
                let _ = writeln!(out, "{}", fl);
            }
        }

        let in_main = (hr..hr + mr).contains(&r);
        let row_label = sheet_row_label(r, grid.main_rows());
        let mut data_line = String::new();
        data_line.push('|');
        if with_row_gutter {
            ascii_push_cell(&mut data_line, pre, pad, &row_label, row_label_w);
        }
        for c in col_start..col_end {
            if frame_active && in_main && c == main_c0 {
                if !options.include_row_label_column && col_start == main_c0 {
                    if data_line.starts_with('|') {
                        data_line.remove(0);
                        data_line.insert(0, '+');
                    }
                } else if data_line
                    .as_bytes()
                    .last()
                    .is_some_and(|&b| b == b'|')
                {
                    data_line.pop();
                    data_line.push('+');
                }
            }
            let val = export_cell_text(grid, r, c, cell_content, generic_rebase);
            let w = col_widths[c];
            ascii_push_cell(&mut data_line, pre, pad, &val, w);
            if frame_active
                && in_main
                && main_c0 < main_c1
                && c + 1 == main_c1
            {
                if data_line
                    .as_bytes()
                    .last()
                    .is_some_and(|&b| b == b'|')
                {
                    data_line.pop();
                    data_line.push('+');
                }
            }
        }

        let _ = writeln!(out, "{}", data_line);
        if options.row_dividers {
            let _ = writeln!(out, "{}", border);
        }
        if frame_active && Some(r) == last_main_r {
            if let Some(ref fl) = frame_line {
                let _ = writeln!(out, "{}", fl);
            }
        }
    }
    let _ = writeln!(out, "{}", border);
}

/// Renders a text table. For backward compatibility, [`export_ascii_table`] fixes `row_dividers`
/// only; use [`export_ascii_table_with_options`] for full control.
pub fn export_ascii_table(grid: &Grid, out: &mut dyn Write, row_dividers: bool) {
    let mut o = AsciiTableOptions::default();
    o.row_dividers = row_dividers;
    export_ascii_table_with_options(grid, out, &o);
}

pub fn export_all(grid: &Grid, out: &mut dyn Write) {
    export_all_with_options(grid, out, &DelimitedExportOptions::default());
}

pub fn export_all_with_options(grid: &Grid, out: &mut dyn Write, options: &DelimitedExportOptions) {
    export_delimited(grid, out, '\t', options);
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

pub fn export_selection(
    grid: &Grid,
    out: &mut dyn Write,
    rows: &[usize],
    cols: &[usize],
    options: &DelimitedExportOptions,
) {
    if rows.is_empty() || cols.is_empty() {
        return;
    }

    let include_header_row = options.include_header_row;
    let content = options.content;
    let generic_rebase = if content == ExportContent::Generic {
        let (dr, dc) = selection_generic_rebase(cols, include_header_row, rows);
        Some((dr, dc))
    } else {
        None
    };

    if include_header_row {
        for (ci, &c) in cols.iter().enumerate() {
            if ci > 0 {
                let _ = write!(out, "\t");
            }
            let label = col_header_label_for_export(grid, c, grid.main_cols(), content);
            let _ = write!(out, "{}", label);
        }
        let _ = writeln!(out);
    }

    for &r in rows {
        for (ci, &c) in cols.iter().enumerate() {
            if ci > 0 {
                let _ = write!(out, "\t");
            }
            let val = export_cell_text(grid, r, c, content, generic_rebase);
            let _ = write!(out, "{}", val);
        }
        let _ = writeln!(out);
    }
}

/// TSV/CSV main-only / selection header token: `<`/`>` margins or `A`/`B` (column letters).
/// Generic mode shows `TAX` etc. on the bottom control header row, not in this synthetic label line.
fn col_header_label_for_export(
    _grid: &Grid,
    global_col: usize,
    main_cols: usize,
    _content: ExportContent,
) -> String {
    let m = MARGIN_COLS;
    if global_col < m {
        format!("<{}", m - 1 - global_col)
    } else if global_col < m + main_cols {
        crate::addr::excel_column_name(global_col - m)
    } else {
        format!(">{}", global_col - m - main_cols)
    }
}

/// With margins: [A / B / ]C (same in Generic; labeled-column titles appear on the `~1` control row).
fn delimited_marginal_header_token(
    _grid: &Grid,
    global_col: usize,
    main_cols: usize,
    _content: ExportContent,
) -> String {
    crate::addr::ui_column_fragment(global_col, main_cols)
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

/// ODF `;` → Excel `,` in function call lists.
fn interop_excel_list_separators(s: &str) -> String {
    s.replace(';', ",")
}

fn finish_generic_interop(s: String, rebase: Option<(i32, i32)>) -> String {
    let Some((d_row, d_col)) = rebase else {
        return s;
    };
    if s.trim_start().starts_with('=') {
        formula::rebase_interop_formula_row_col(&s, d_row, d_col)
    } else {
        s
    }
}

/// Deltas for [`formula::rebase_interop_formula_row_col`]: the exported TSV/CSV/ASCII file’s
/// top-left cell (line 0, field 0) is treated as Excel A1, so all `=…` refs shift by these.
fn delimited_generic_rebase(
    col_start: usize,
    col_end: usize,
    include_header_row: bool,
    include_row_label_column: bool,
    main_rows: &[usize],
) -> (i32, i32) {
    // Row labels (~1, 1, 2, …) are a field *before* the col loop: count them in d_col.
    // Left margin `<…` / `[A` style cols are the first c in (col_start..MARGIN_COLS) in the loop;
    // `position(== MARGIN_COLS)` is exactly that many.
    let d_col = (if include_row_label_column { 1 } else { 0 })
        + (col_start..col_end)
            .position(|c| c == MARGIN_COLS)
            .map(|i| i as i32)
            .unwrap_or(0);
    let base = if include_header_row { 1 } else { 0 };
    let d_row = main_rows
        .iter()
        .position(|&r| r == HEADER_ROWS)
        .map(|j| base + j as i32)
        .unwrap_or(0);
    (d_row, d_col)
}

fn selection_generic_rebase(
    cols: &[usize],
    include_header_row: bool,
    rows: &[usize],
) -> (i32, i32) {
    let d_col = cols
        .iter()
        .position(|&c| c == MARGIN_COLS)
        .map(|i| i as i32)
        .unwrap_or(0);
    let base = if include_header_row { 1 } else { 0 };
    let d_row = rows
        .iter()
        .position(|&r| r == HEADER_ROWS)
        .map(|j| base + j as i32)
        .unwrap_or(0);
    (d_row, d_col)
}

/// Same row/col semantics as delimited, matching [`export_ascii_table_with_options`]'s `writeln!` order.
fn ascii_generic_rebase(
    col_start: usize,
    col_end: usize,
    row_start: usize,
    row_end: usize,
    first_main_r: Option<usize>,
    last_main_r: Option<usize>,
    frame_active: bool,
    options: &AsciiTableOptions,
) -> (i32, i32) {
    let d_col = (if options.include_row_label_column { 1 } else { 0 })
        + (col_start..col_end)
            .position(|c| c == MARGIN_COLS)
            .map(|i| i as i32)
            .unwrap_or(0);
    if !(row_start..row_end).contains(&HEADER_ROWS) {
        return (0, d_col);
    }
    let hr = HEADER_ROWS;
    let mut d_row: i32 = 0;
    d_row = d_row.saturating_add(1);
    if options.include_column_label_row {
        d_row = d_row.saturating_add(1);
        if matches!(
            options.header_data_separator,
            AsciiHeaderDataSeparator::FullBorder
        ) {
            d_row = d_row.saturating_add(1);
        }
    }
    for r in row_start..hr {
        if frame_active && first_main_r == Some(r) {
            d_row = d_row.saturating_add(1);
        }
        d_row = d_row.saturating_add(1);
        if options.row_dividers {
            d_row = d_row.saturating_add(1);
        }
        if frame_active && last_main_r == Some(r) {
            d_row = d_row.saturating_add(1);
        }
    }
    if frame_active && first_main_r == Some(hr) {
        d_row = d_row.saturating_add(1);
    }
    (d_row, d_col)
}

fn generic_interop_text(
    grid: &Grid,
    logical_row: usize,
    global_col: usize,
    rebase: Option<(i32, i32)>,
) -> Option<String> {
    if logical_row + 1 == HEADER_ROWS
        && (MARGIN_COLS..(MARGIN_COLS + grid.main_cols())).contains(&global_col)
    {
        let main_c = global_col - MARGIN_COLS;
        if let Some(lab) = formula::main_column_label_from_header(grid, main_c) {
            return Some(lab);
        }
    }
    use crate::ui::SheetCursor;
    let cur = SheetCursor {
        row: logical_row,
        col: global_col,
    };
    let addr = cur.to_addr(grid);

    if let Some(tf) = formula::export_templated_formula(grid, &addr) {
        let s = interop_excel_list_separators(
            &crate::ods::ods_labeled_prefix_strip_to_formula(&tf).unwrap_or(tf),
        );
        return Some(finish_generic_interop(s, rebase));
    }

    let v = crate::ods::cell_export_value_string(grid, logical_row, global_col);
    if !v.is_empty() {
        if let Some(st) = crate::ods::ods_labeled_prefix_strip_to_formula(&v) {
            return Some(finish_generic_interop(
                interop_excel_list_separators(&st),
                rebase,
            ));
        }
    }
    if let Some(raw) = grid.get(&addr) {
        if formula::is_formula(&raw) {
            if let Some(st) = crate::ods::ods_labeled_prefix_strip_to_formula(&raw) {
                return Some(finish_generic_interop(
                    interop_excel_list_separators(&st),
                    rebase,
                ));
            }
        }
    }
    None
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

fn export_cell_text(
    grid: &Grid,
    logical_row: usize,
    global_col: usize,
    content: ExportContent,
    generic_rebase: Option<(i32, i32)>,
) -> String {
    match content {
        ExportContent::Values => rendered_value_at(grid, logical_row, global_col),
        ExportContent::Formulas => cell_value_at(grid, logical_row, global_col),
        ExportContent::Generic => {
            generic_interop_text(grid, logical_row, global_col, generic_rebase)
                .unwrap_or_else(|| rendered_value_at(grid, logical_row, global_col))
        }
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
    options: &DelimitedExportOptions,
) {
    let include_headers = options.include_header_row;
    let include_margins = options.include_margins;
    let row_key_col = options.include_row_label_column;
    let content = options.content;
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
            if row_key_col {
                // Match UI: leading row-label column; header cell is blank. First field is empty,
                // so the line starts with the delimiter (tab for TSV, comma for CSV).
                let _ = write!(
                    out,
                    "{}{}",
                    delim,
                    delimited_marginal_header_token(grid, col_start, mc, content)
                );
                for c in (col_start + 1)..col_end {
                    let _ = write!(
                        out,
                        "{}{}",
                        delim,
                        delimited_marginal_header_token(grid, c, mc, content)
                    );
                }
            } else {
                let _ = write!(
                    out,
                    "{}",
                    delimited_marginal_header_token(grid, col_start, mc, content)
                );
                for c in (col_start + 1)..col_end {
                    let _ = write!(
                        out,
                        "{}{}",
                        delim,
                        delimited_marginal_header_token(grid, c, mc, content)
                    );
                }
            }
        } else {
            if row_key_col {
                let _ = write!(out, "{delim}");
            }
            for c in col_start..col_end {
                if c > col_start {
                    let _ = write!(out, "{delim}");
                }
                let label = col_header_label_for_export(grid, c, mc, content);
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
    let generic_rebase = if content == ExportContent::Generic {
        let (dr, dc) = delimited_generic_rebase(
            col_start,
            col_end,
            include_headers,
            row_key_col,
            &rows,
        );
        Some((dr, dc))
    } else {
        None
    };
    for r in rows {
        if row_key_col {
            let _ = write!(out, "{}", sheet_row_label(r, mr));
            let _ = write!(out, "{delim}");
        }
        let mut first = true;
        for c in col_start..col_end {
            if !first {
                let _ = write!(out, "{delim}");
            }
            first = false;
            let val = export_cell_text(grid, r, c, content, generic_rebase);
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

    #[test]
    fn delimited_omit_header_row_starts_with_data() {
        let mut grid = crate::grid::Grid::new(2, 2);
        grid.set(&CellAddr::Main { row: 0, col: 0 }, "V1".into());
        grid.set(&CellAddr::Main { row: 0, col: 1 }, "V2".into());
        let gb = crate::grid::GridBox::from(grid);
        let opts = DelimitedExportOptions {
            include_header_row: false,
            include_margins: false,
            include_row_label_column: false,
            ..Default::default()
        };
        let mut out = Vec::new();
        export_tsv_with_options(&gb, &mut out, &opts);
        let tsv = String::from_utf8(out).unwrap();
        let first = tsv.lines().next().expect("at least one line");
        assert_eq!(first, "V1\tV2", "first line should be data, not A/B");
    }

    #[test]
    fn delimited_main_only_can_keep_row_key_without_margins() {
        let mut grid = crate::grid::Grid::new(2, 2);
        grid.set(&CellAddr::Main { row: 0, col: 0 }, "V1".into());
        grid.set(&CellAddr::Main { row: 0, col: 1 }, "V2".into());
        let gb = crate::grid::GridBox::from(grid);
        let opts = DelimitedExportOptions {
            include_header_row: false,
            include_margins: false,
            include_row_label_column: true,
            ..Default::default()
        };
        let mut out = Vec::new();
        export_tsv_with_options(&gb, &mut out, &opts);
        let tsv = String::from_utf8(out).unwrap();
        let first = tsv.lines().next().expect("at least one line");
        assert!(
            first.starts_with("1\t"),
            "main-only TSV with row# on should start with row label: {first:?}"
        );
        assert!(first.contains("V1"));
    }

    #[test]
    fn generic_tsv_uses_tsv_header_and_interop_formula() {
        use crate::grid::HEADER_ROWS;

        let mut grid = crate::grid::Grid::new(2, 2);
        grid.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: MARGIN_COLS as u32 + 1,
            },
            "=A*0.1 -- TAX".into(),
        );
        grid.set(&CellAddr::Main { row: 0, col: 0 }, "100".into());
        let gb = crate::grid::GridBox::from(grid);
        let opts = DelimitedExportOptions {
            content: ExportContent::Generic,
            include_margins: false,
            include_header_row: true,
            include_row_label_column: false,
            ..Default::default()
        };
        let mut out = Vec::new();
        export_tsv_with_options(&gb, &mut out, &opts);
        let tsv = String::from_utf8(out).unwrap();
        let lines: Vec<_> = tsv.lines().collect();
        assert_eq!(lines[0], "A\tB", "synthetic A/B; labeled header text is on the ~1 control row");
        // `row_order` lists non-empty header rows before the main block, so the control-strip row
        // appears as an extra data line before the first main data row.
        assert!(
            lines
                .iter()
                .any(|l| *l == "\tTAX" || l.ends_with("\tTAX")),
            "TAX replaces =A*0.1 on the control row, got {lines:?}"
        );
        let main_row = lines
            .iter()
            .find(|l| l.starts_with("100\t"))
            .expect("main data row with 100 in column A");
        // d_row=2, d_col=0: two body lines (control strip + main) before the first data row, no
        // left columns in this main-only export, so =A1*0.1 => =A3*0.1 (file A1 = top-left).
        assert!(
            (main_row.contains("A3*0.1") && main_row.contains('='))
                || main_row.contains("(A3*0.1)"),
            "TAX column rebased, got {main_row:?}"
        );
    }

    #[test]
    fn generic_right_margin_uses_excel_list_sep() {
        use crate::grid::HEADER_ROWS;

        let mut grid = crate::grid::Grid::new(1, 2);
        grid.set(
            &CellAddr::Header {
                row: (HEADER_ROWS - 1) as u32,
                col: MARGIN_COLS as u32 + 1,
            },
            "MAX".into(),
        );
        grid.set(&CellAddr::Main { row: 0, col: 0 }, "3".into());
        grid.set(&CellAddr::Right { col: 0, row: 0 }, "MAX".into());
        let gb = crate::grid::GridBox::from(grid);
        let opts = DelimitedExportOptions {
            content: ExportContent::Generic,
            include_margins: true,
            include_header_row: true,
            include_row_label_column: true,
            ..Default::default()
        };
        let mut out = Vec::new();
        export_tsv_with_options(&gb, &mut out, &opts);
        let tsv = String::from_utf8(out).unwrap();
        assert!(tsv.contains("=SUBTOTAL(4,"), "expect comma, got {tsv}");
    }

    #[test]
    fn selection_omit_header_row_is_data_first() {
        use crate::grid::HEADER_ROWS;

        let mut grid = crate::grid::Grid::new(2, 2);
        grid.set(&CellAddr::Main { row: 0, col: 0 }, "a".into());
        let gb = crate::grid::GridBox::from(grid);
        let m = MARGIN_COLS;
        let mut out = Vec::new();
        export_selection(
            &gb,
            &mut out,
            &[HEADER_ROWS],
            &[m, m + 1],
            &DelimitedExportOptions {
                include_header_row: false,
                ..Default::default()
            },
        );
        let s = String::from_utf8(out).unwrap();
        let first = s.lines().next().expect("at least one line");
        assert_eq!(first, "a\t", "first line is data, not column labels");

        let mut out2 = Vec::new();
        export_selection(
            &gb,
            &mut out2,
            &[HEADER_ROWS],
            &[m, m + 1],
            &DelimitedExportOptions {
                include_header_row: true,
                ..Default::default()
            },
        );
        let s2 = String::from_utf8(out2).unwrap();
        let first2 = s2.lines().next().expect("at least one line");
        assert_eq!(first2, "A\tB", "with header, first line is A/B for two main columns");
    }

    #[test]
    fn ascii_ems_space_in_cell_glue() {
        let mut g = crate::grid::Grid::new(3, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "x".into());
        let gb = crate::grid::GridBox::from(g);
        let em = '\u{2003}';
        let o = AsciiTableOptions {
            inter_cell_space: AsciiInterCellSpace::EmSpace,
            ..Default::default()
        };
        let mut out = Vec::new();
        export_ascii_table_with_options(&gb, &mut out, &o);
        let s = String::from_utf8(out).unwrap();
        assert!(s.contains(em), "expected em U+2003 in output");
    }

    #[test]
    fn ascii_omit_row_label_column_starts_with_column_not_row_numbers() {
        let mut g = crate::grid::Grid::new(3, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "x".into());
        let gb = crate::grid::GridBox::from(g);
        let o = AsciiTableOptions {
            include_row_label_column: false,
            include_margins: false,
            ..Default::default()
        };
        let mut out = Vec::new();
        export_ascii_table_with_options(&gb, &mut out, &o);
        let s = String::from_utf8(out).unwrap();
        let data_line = s.lines().find(|l| l.contains("x")).expect("data row");
        assert!(
            !data_line.contains("  1  ") && !data_line.contains("| 1 |"),
            "no row-number gutter: {data_line}"
        );
        assert!(data_line.contains("x"));
    }

    #[test]
    fn ascii_omit_column_label_row_goes_straight_to_data() {
        let mut g = crate::grid::Grid::new(3, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "val".into());
        let gb = crate::grid::GridBox::from(g);
        let o = AsciiTableOptions {
            include_column_label_row: false,
            ..Default::default()
        };
        let mut out = Vec::new();
        export_ascii_table_with_options(&gb, &mut out, &o);
        let s = String::from_utf8(out).unwrap();
        let lines: Vec<_> = s.lines().collect();
        assert!(lines.len() >= 2);
        // After top border, first line is a data row (row label + cell), not a column-letter row.
        assert!(
            lines[1].contains("val"),
            "line after top border should be data, got: {:?}",
            lines[1]
        );
    }

    #[test]
    fn ascii_header_data_separator_none_drops_line_between_label_and_data() {
        let mut g = crate::grid::Grid::new(3, 1);
        g.set(&CellAddr::Main { row: 0, col: 0 }, "z".into());
        let gb = crate::grid::GridBox::from(g);
        let o_full = AsciiTableOptions {
            header_data_separator: AsciiHeaderDataSeparator::FullBorder,
            ..Default::default()
        };
        let o_none = AsciiTableOptions {
            header_data_separator: AsciiHeaderDataSeparator::None,
            ..Default::default()
        };
        let mut out_full = Vec::new();
        let mut out_none = Vec::new();
        export_ascii_table_with_options(&gb, &mut out_full, &o_full);
        export_ascii_table_with_options(&gb, &mut out_none, &o_none);
        let full = String::from_utf8(out_full).unwrap();
        let none = String::from_utf8(out_none).unwrap();
        assert!(
            full.lines().count() > none.lines().count(),
            "FullBorder should add a line under labels: full={} none={}",
            full.lines().count(),
            none.lines().count()
        );
    }
}
