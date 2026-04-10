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

/// 10-column margin label (`A` nearest the main grid, `J` farthest).
pub fn mirror_margin_column_name(margin_col_index: usize, left_side: bool) -> String {
    let idx = margin_col_index.min(9);
    let idx = if left_side { 9 - idx } else { idx };
    ((b'A' + idx as u8) as char).to_string()
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

/// Parse a column fragment at the start of a cell ref.
pub fn parse_ui_column_fragment(s: &str, main_cols: usize) -> Option<(u32, usize)> {
    if let Some(rest) = s.strip_prefix('[') {
        let col_len = rest.chars().take_while(|c| c.is_ascii_uppercase()).count();
        if col_len != 1 {
            return None;
        }
        let col = parse_mirror_margin_column_name(&rest[..col_len], true)?;
        return Some((col as u32, 1 + col_len));
    }
    if let Some(rest) = s.strip_prefix(']') {
        let col_len = rest.chars().take_while(|c| c.is_ascii_uppercase()).count();
        if col_len != 1 {
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

pub(crate) fn parse_mirror_margin_column_name(name: &str, left_side: bool) -> Option<u8> {
    let mut chars = name.chars();
    let ch = chars.next()?;
    if chars.next().is_some() || !ch.is_ascii_uppercase() {
        return None;
    }
    let idx = (ch as u8 - b'A') as usize;
    if idx > 9 {
        return None;
    }
    Some(if left_side {
        (9 - idx) as u8
    } else {
        idx as u8
    })
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

    let col_len = if prefix.is_some() {
        let len = rest.chars().take_while(|c| c.is_ascii_uppercase()).count();
        if len != 1 {
            return None;
        }
        len
    } else {
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
            (crate::grid::HEADER_ROWS - row_num) as u8
        } else {
            if row_num == 0 || row_num > crate::grid::FOOTER_ROWS {
                return None;
            }
            (row_num - 1) as u8
        };
        let col = match prefix {
            Some(true) => parse_mirror_margin_column_name(col_name, true)? as u32,
            Some(false) => {
                (crate::grid::MARGIN_COLS
                    + main_cols
                    + parse_mirror_margin_column_name(col_name, false)? as usize)
                    as u32
            }
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
        assert_eq!(mirror_margin_column_name(0, true), "J");
        assert_eq!(mirror_margin_column_name(9, true), "A");
        assert_eq!(
            parse_cell_ref_at("[A1", 1).unwrap().0,
            CellAddr::Left { col: 9, row: 0 }
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
            CellAddr::Footer { row: 2, col: 0 }
        );
        assert_eq!(
            parse_cell_ref_at("[A_3", 4).unwrap().0,
            CellAddr::Footer { row: 2, col: 9 }
        );
        assert_eq!(
            parse_cell_ref_at("]A~3", 4).unwrap().0,
            CellAddr::Header { row: 23, col: 15 }
        );
    }
}
