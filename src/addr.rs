//! Shared cell-address parsing (Excel columns, global column suffixes, single-cell refs).

use crate::grid::{CellAddr, HEADER_ROWS};

/// Parse Excel-style column name `A`..`ZZZ` → 0-based main column index.
pub fn parse_excel_column(name: &str) -> Option<u32> {
    let mut n: u32 = 0;
    for b in name.bytes() {
        if !b.is_ascii_uppercase() {
            return None;
        }
        n = n.checked_mul(26)?.checked_add((b - b'A') as u32 + 1)?;
    }
    Some(n - 1)
}

/// 0-based main column index → Excel column letters.
pub fn excel_column_name(main_col_index: usize) -> String {
    let mut n = main_col_index + 1;
    let mut s = String::new();
    while n > 0 {
        n -= 1;
        s.push((b'A' + (n % 26) as u8) as char);
        n /= 26;
    }
    s.chars().rev().collect()
}

/// Margin label (`A` nearest the main grid/right edge, up to `ZZ`).
pub fn mirror_margin_column_name(margin_col_index: usize, left_side: bool) -> String {
    // Map the margin_col_index (0..MARGIN_COLS-1) into a letter sequence.
    // If left_side is true, mirror the index (so 0 -> last, as in previous
    // behavior for small margins).
    let max = crate::grid::MARGIN_COLS;
    let idx = margin_col_index.min(max.saturating_sub(1));
    let mapped = if left_side {
        max.saturating_sub(1).saturating_sub(idx)
    } else {
        idx
    };
    // Use excel-style column naming for the mapped index (0 -> A, 25 -> Z,
    // 26 -> AA, ...). Reuse excel_column_name which is 0-based.
    excel_column_name(mapped)
}

/// UI-style column fragment for display and formulas.
pub fn ui_column_fragment(global_col: usize, main_cols: usize) -> String {
    let m = crate::grid::MARGIN_COLS;
    if global_col < m {
        format!("[{}", mirror_margin_column_name(global_col, true))
    } else if global_col < m + main_cols {
        excel_column_name(global_col - m)
    } else {
        format!(
            "]{}",
            mirror_margin_column_name(global_col - m - main_cols, false)
        )
    }
}

/// UI-style row label for the left gutter (`~N`, `1`, `_N`).
pub fn ui_row_label(logical_row: usize, main_rows: usize) -> String {
    let hr = crate::grid::HEADER_ROWS;
    if logical_row < hr {
        format!("~{}", hr - logical_row)
    } else if logical_row < hr + main_rows {
        format!("{}", logical_row - hr + 1)
    } else {
        let fr = logical_row - hr - main_rows;
        format!("_{}", fr + 1)
    }
}

/// Convert a logical sheet cursor (`row`, global `col`) to a concrete cell address.
pub fn sheet_cursor_to_addr(
    logical_row: usize,
    global_col: usize,
    main_rows: usize,
    main_cols: usize,
) -> CellAddr {
    let hr = crate::grid::HEADER_ROWS;
    if logical_row < hr {
        CellAddr::Header {
            row: logical_row as u32,
            col: global_col as u32,
        }
    } else if logical_row < hr + main_rows {
        let main_row = logical_row - hr;
        if global_col < crate::grid::MARGIN_COLS {
            CellAddr::Left {
                col: global_col,
                row: main_row as u32,
            }
        } else if global_col < crate::grid::MARGIN_COLS + main_cols {
            CellAddr::Main {
                row: main_row as u32,
                col: (global_col - crate::grid::MARGIN_COLS) as u32,
            }
        } else {
            CellAddr::Right {
                col: global_col - crate::grid::MARGIN_COLS - main_cols,
                row: main_row as u32,
            }
        }
    } else {
        CellAddr::Footer {
            row: (logical_row - hr - main_rows) as u32,
            col: global_col as u32,
        }
    }
}

/// Convert a concrete cell address to a logical sheet cursor (`row`, global `col`).
pub fn addr_to_sheet_cursor(addr: &CellAddr, main_rows: usize, main_cols: usize) -> (usize, usize) {
    let row_col = match addr {
        CellAddr::Header { row, col } => (*row as usize, *col as usize),
        CellAddr::Footer { row, col } => (
            crate::grid::HEADER_ROWS + main_rows + *row as usize,
            *col as usize,
        ),
        CellAddr::Main { row, col } => (
            crate::grid::HEADER_ROWS + *row as usize,
            crate::grid::MARGIN_COLS + *col as usize,
        ),
        CellAddr::Left { col, row } => (crate::grid::HEADER_ROWS + *row as usize, *col as usize),
        CellAddr::Right { col, row } => (
            crate::grid::HEADER_ROWS + *row as usize,
            crate::grid::MARGIN_COLS + main_cols + *col as usize,
        ),
    };
    row_col
}

/// Parse a column fragment at the start of a cell ref.
pub fn parse_ui_column_fragment(s: &str, main_cols: usize) -> Option<(u32, usize)> {
    if let Some(rest) = s.strip_prefix('[') {
        let col_len = rest.chars().take_while(|c| c.is_ascii_uppercase()).count();
        if col_len == 0 {
            return None;
        }
        let col = parse_mirror_margin_column_name(&rest[..col_len], true)?;
        return Some((col as u32, 1 + col_len));
    }
    if let Some(rest) = s.strip_prefix(']') {
        let col_len = rest.chars().take_while(|c| c.is_ascii_uppercase()).count();
        if col_len == 0 {
            return None;
        }
        let col = parse_mirror_margin_column_name(&rest[..col_len], false)?;
        return Some((
            crate::grid::MARGIN_COLS as u32 + main_cols as u32 + col as u32,
            1 + col_len,
        ));
    }
    let col_len = s.chars().take_while(|c| c.is_ascii_uppercase()).count();
    if col_len == 0 {
        return None;
    }
    let col = parse_excel_column(&s[..col_len])?;
    Some((crate::grid::MARGIN_COLS as u32 + col, col_len))
}

/// Back-compat alias for the UI-style column fragment.
pub fn ui_column_name(global_col: usize, main_cols: usize) -> String {
    ui_column_fragment(global_col, main_cols)
}

/// Parse a sheet id prefix like `$12` at the start of `s`.
pub fn parse_sheet_id_prefix_at(s: &str) -> Option<(u32, usize)> {
    let bytes = s.as_bytes();
    if bytes.first().copied()? != b'$' {
        return None;
    }
    let mut i = 1usize;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i == 1 {
        return None;
    }
    let sheet_id = std::str::from_utf8(&bytes[1..i]).ok()?.parse().ok()?;
    Some((sheet_id, i))
}

/// Parse a sheet-qualified cell ref like `$2:A1` at the start of `s`.
pub fn parse_sheet_qualified_cell_ref_at(
    s: &str,
    main_cols: usize,
) -> Option<(u32, CellAddr, usize)> {
    let (sheet_id, prefix_len) = parse_sheet_id_prefix_at(s)?;
    let rest = s.get(prefix_len..)?;
    let rest = rest.strip_prefix(':')?;
    let (addr, addr_len) = parse_cell_ref_at(rest, main_cols)?;
    Some((sheet_id, addr, prefix_len + 1 + addr_len))
}

pub(crate) fn parse_mirror_margin_column_name(name: &str, left_side: bool) -> Option<usize> {
    // Accept multi-letter uppercase sequences and parse them like Excel
    // columns, then map according to left_side mirroring.
    if name.is_empty() || !name.chars().all(|c| c.is_ascii_uppercase()) {
        return None;
    }
    let parsed = parse_excel_column(name)? as usize; // 0-based
    if parsed >= crate::grid::MARGIN_COLS {
        return None;
    }
    let mapped = if left_side {
        crate::grid::MARGIN_COLS - 1 - parsed
    } else {
        parsed
    };
    Some(mapped)
}

/// Parse one cell reference at the start of `s` (no leading whitespace).
/// Returns `(address, byte length consumed)`.
pub fn parse_cell_ref_at(s: &str, main_cols: usize) -> Option<(CellAddr, usize)> {
    let bytes = s.as_bytes();
    if bytes.is_empty() {
        return None;
    }

    let (prefix, rest, prefix_len) = match bytes[0] {
        b'[' => (Some(true), &s[1..], 1usize),
        b']' => (Some(false), &s[1..], 1usize),
        _ => (None, s, 0usize),
    };

    let col_len = {
        let len = rest.chars().take_while(|c| c.is_ascii_uppercase()).count();
        if len == 0 {
            return None;
        }
        len
    };
    let col_name = &rest[..col_len];
    let after = &rest[col_len..];

    if let Some(marker) = after.chars().next().filter(|c| *c == '~' || *c == '_') {
        let row_digits = after[1..]
            .chars()
            .take_while(|c| c.is_ascii_digit())
            .count();
        if row_digits == 0 {
            return None;
        }
        let row_num: usize = after[1..1 + row_digits].parse().ok()?;
        let row = if marker == '~' {
            if row_num == 0 || row_num > crate::grid::HEADER_ROWS {
                return None;
            }
            (crate::grid::HEADER_ROWS - row_num) as u32
        } else {
            if row_num == 0 || row_num > crate::grid::FOOTER_ROWS {
                return None;
            }
            (row_num - 1) as u32
        };
        let col = match prefix {
            Some(true) => parse_mirror_margin_column_name(col_name, true)? as u32,
            Some(false) => parse_mirror_margin_column_name(col_name, false)
                .map(|c| (crate::grid::MARGIN_COLS + main_cols + c as usize) as u32)
                .or_else(|| Some(parse_excel_column(col_name)?))?,
            // Unprefixed Excel letters always refer to the main/data region
            // column. Map to a global column index by adding the left-margin
            // offset.
            None => crate::grid::MARGIN_COLS as u32 + parse_excel_column(col_name)?,
        };
        return Some((
            if marker == '~' {
                CellAddr::Header { row, col }
            } else {
                CellAddr::Footer { row, col }
            },
            prefix_len + col_len + 1 + row_digits,
        ));
    }

    let row_digits = after.chars().take_while(|c| c.is_ascii_digit()).count();
    if row_digits == 0 {
        return None;
    }
    let row_num: u32 = after[..row_digits].parse().ok()?;
    if row_num == 0 {
        return None;
    }
    let addr = match prefix {
        Some(true) => CellAddr::Left {
            col: parse_mirror_margin_column_name(col_name, true)?,
            row: row_num - 1,
        },
        Some(false) => CellAddr::Right {
            col: parse_mirror_margin_column_name(col_name, false)?,
            row: row_num - 1,
        },
        None => CellAddr::Main {
            row: row_num - 1,
            // Always map unprefixed Excel letters to main/data columns.
            col: parse_excel_column(col_name)?,
        },
    };
    Some((addr, prefix_len + col_len + row_digits))
}

pub fn cell_ref_text(addr: &CellAddr, main_cols: usize) -> String {
    match addr {
        CellAddr::Header { row, col } => {
            let row = HEADER_ROWS - *row as usize;
            if (*col as usize) < crate::grid::MARGIN_COLS {
                format!(
                    "[{}~{}",
                    mirror_margin_column_name(*col as usize, true),
                    row
                )
            } else if (*col as usize) < crate::grid::MARGIN_COLS + main_cols {
                format!(
                    "{}~{}",
                    excel_column_name(*col as usize - crate::grid::MARGIN_COLS),
                    row
                )
            } else {
                format!(
                    "]{}~{}",
                    mirror_margin_column_name(
                        *col as usize - crate::grid::MARGIN_COLS - main_cols,
                        false
                    ),
                    row
                )
            }
        }
        CellAddr::Footer { row, col } => {
            let row = *row as usize + 1;
            if (*col as usize) < crate::grid::MARGIN_COLS {
                format!("[{}_{row}", mirror_margin_column_name(*col as usize, true))
            } else if (*col as usize) < crate::grid::MARGIN_COLS + main_cols {
                format!(
                    "{}_{row}",
                    excel_column_name(*col as usize - crate::grid::MARGIN_COLS)
                )
            } else {
                format!(
                    "]{}_{row}",
                    mirror_margin_column_name(
                        *col as usize - crate::grid::MARGIN_COLS - main_cols,
                        false
                    )
                )
            }
        }
        CellAddr::Main { row, col } => format!("{}{}", excel_column_name(*col as usize), row + 1),
        CellAddr::Left { col, row } => format!(
            "[{}{}",
            mirror_margin_column_name(*col as usize, true),
            row + 1
        ),
        CellAddr::Right { col, row } => format!(
            "]{}{}",
            mirror_margin_column_name(*col as usize, false),
            row + 1
        ),
    }
}

pub fn sheet_qualified_cell_ref_text(sheet_id: u32, addr: &CellAddr, main_cols: usize) -> String {
    format!("${sheet_id}:{}", cell_ref_text(addr, main_cols))
}

/// Parse `A1:B2` at start of `s`; both ends must be main cells. Returns range + consumed length.
pub fn parse_main_range_at(s: &str) -> Option<(crate::grid::MainRange, usize)> {
    let (a, na) = parse_cell_ref_at(s, 0)?;
    let CellAddr::Main { row: ra, col: ca } = a else {
        return None;
    };
    let rest = s.get(na..)?;
    let rest = rest.strip_prefix(':')?;
    let (b, nb) = parse_cell_ref_at(rest, 0)?;
    let CellAddr::Main { row: rb, col: cb } = b else {
        return None;
    };
    let r0 = ra.min(rb);
    let r1 = ra.max(rb);
    let c0 = ca.min(cb);
    let c1 = ca.max(cb);
    let range = crate::grid::MainRange {
        row_start: r0,
        row_end: r1 + 1,
        col_start: c0,
        col_end: c1 + 1,
    };
    Some((range, na + 1 + nb))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a1_roundtrip() {
        let (a, n) = parse_cell_ref_at("A1", 1).unwrap();
        assert_eq!(n, 2);
        assert_eq!(a, CellAddr::Main { row: 0, col: 0 });
    }

    #[test]
    fn main_range() {
        let (r, n) = parse_main_range_at("B2:A1").unwrap();
        assert_eq!(n, 5);
        assert_eq!(r.row_start, 0);
        assert_eq!(r.row_end, 2);
        assert_eq!(r.col_start, 0);
        assert_eq!(r.col_end, 2);
    }

    #[test]
    fn legacy_special_refs_parse() {
        assert_eq!(parse_cell_ref_at("A~1", 1).unwrap().1, 3);
        assert_eq!(parse_cell_ref_at("A_1", 1).unwrap().1, 3);
        assert_eq!(parse_cell_ref_at("[A1", 1).unwrap().1, 3);
        assert_eq!(parse_cell_ref_at("]A1", 1).unwrap().1, 3);
    }

    #[test]
    fn left_margin_is_mirrored_from_the_main_grid() {
        assert_eq!(mirror_margin_column_name(0, true), "ZZ");
        assert_eq!(
            mirror_margin_column_name(crate::grid::MARGIN_COLS - 1, true),
            "A"
        );
        assert_eq!(
            parse_cell_ref_at("[A1", 1).unwrap().0,
            CellAddr::Left {
                col: crate::grid::MARGIN_COLS - 1,
                row: 0
            }
        );
    }

    #[test]
    fn sheet_qualified_cell_refs_parse() {
        let (sheet_id, addr, len) = parse_sheet_qualified_cell_ref_at("$12:A5", 1).unwrap();
        assert_eq!(sheet_id, 12);
        assert_eq!(addr, CellAddr::Main { row: 4, col: 0 });
        assert_eq!(len, 6);
    }

    #[test]
    fn parses_corners_and_footers() {
        assert_eq!(
            parse_cell_ref_at("A_3", 4).unwrap().0,
            CellAddr::Footer {
                row: 2,
                col: crate::grid::MARGIN_COLS as u32
            }
        );
        assert_eq!(
            parse_cell_ref_at("[A_3", 4).unwrap().0,
            CellAddr::Footer {
                row: 2,
                col: (crate::grid::MARGIN_COLS - 1) as u32
            }
        );
        assert_eq!(
            parse_cell_ref_at("]A~3", 4).unwrap().0,
            CellAddr::Header {
                row: (crate::grid::HEADER_ROWS - 3) as u32,
                col: (crate::grid::MARGIN_COLS + 4) as u32
            }
        );
    }

    #[test]
    fn parses_boundary_header_footer_rows() {
        assert_eq!(
            parse_cell_ref_at("A~999999999", 1).unwrap().0,
            CellAddr::Header {
                row: 0,
                col: crate::grid::MARGIN_COLS as u32
            }
        );
        assert_eq!(
            parse_cell_ref_at("A_999999999", 1).unwrap().0,
            CellAddr::Footer {
                row: 999_999_998,
                col: crate::grid::MARGIN_COLS as u32
            }
        );
        assert!(parse_cell_ref_at("A~1000000000", 1).is_none());
        assert!(parse_cell_ref_at("A_1000000000", 1).is_none());
    }

    #[test]
    fn ui_column_fragment_roundtrip() {
        let main_cols = 3usize;
        let cols = [
            crate::grid::MARGIN_COLS - 1,
            crate::grid::MARGIN_COLS,
            crate::grid::MARGIN_COLS + 1,
            crate::grid::MARGIN_COLS + main_cols,
        ];
        for col in cols {
            let frag = ui_column_fragment(col, main_cols);
            let (parsed, n) = parse_ui_column_fragment(&frag, main_cols).unwrap();
            assert_eq!(n, frag.len());
            assert_eq!(parsed as usize, col);
        }
    }

    #[test]
    fn ui_row_label_regions() {
        let main_rows = 2usize;
        assert_eq!(ui_row_label(0, main_rows), format!("~{}", crate::grid::HEADER_ROWS));
        assert_eq!(ui_row_label(crate::grid::HEADER_ROWS, main_rows), "1");
        assert_eq!(ui_row_label(crate::grid::HEADER_ROWS + main_rows, main_rows), "_1");
    }

    #[test]
    fn cursor_addr_roundtrip_across_regions() {
        let main_rows = 3usize;
        let main_cols = 4usize;
        let addrs = [
            CellAddr::Header {
                row: 0,
                col: crate::grid::MARGIN_COLS as u32,
            },
            CellAddr::Left {
                col: crate::grid::MARGIN_COLS - 1,
                row: 1,
            },
            CellAddr::Main { row: 2, col: 3 },
            CellAddr::Right { col: 0, row: 2 },
            CellAddr::Footer {
                row: 0,
                col: (crate::grid::MARGIN_COLS + 1) as u32,
            },
        ];
        for addr in addrs {
            let (row, col) = addr_to_sheet_cursor(&addr, main_rows, main_cols);
            let back = sheet_cursor_to_addr(row, col, main_rows, main_cols);
            assert_eq!(back, addr);
        }
    }
}
