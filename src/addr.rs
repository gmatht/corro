//! Shared cell-address parsing (Excel columns, global column suffixes, single-cell refs).

use crate::grid::CellAddr;

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
pub fn parse_sheet_qualified_cell_ref_at(s: &str) -> Option<(u32, CellAddr, usize)> {
    let (sheet_id, prefix_len) = parse_sheet_id_prefix_at(s)?;
    let rest = s.get(prefix_len..)?;
    let rest = rest.strip_prefix(':')?;
    let (addr, addr_len) = parse_cell_ref_at(rest)?;
    Some((sheet_id, addr, prefix_len + 1 + addr_len))
}

fn parse_mirror_margin_column_name(name: &str, left_side: bool) -> Option<u8> {
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
pub fn parse_cell_ref_at(s: &str) -> Option<(CellAddr, usize)> {
    let bytes = s.as_bytes();
    if bytes.is_empty() {
        return None;
    }

    // Main: A1, AA10
    if bytes[0].is_ascii_uppercase() {
        let mut col_end = 0;
        while col_end < bytes.len() && bytes[col_end].is_ascii_uppercase() {
            col_end += 1;
        }
        if col_end == 0 || col_end >= bytes.len() {
            return None;
        }
        let col_name = &s[..col_end];
        let rest = &s[col_end..];
        let row_digits = rest.chars().take_while(|c| c.is_ascii_digit()).count();
        if row_digits == 0 {
            return None;
        }
        let row_str = &rest[..row_digits];
        let row_num: u32 = row_str.parse().ok()?;
        if row_num == 0 {
            return None;
        }
        let col = parse_excel_column(col_name)?;
        let consumed = col_end + row_digits;
        return Some((
            CellAddr::Main {
                row: row_num - 1,
                col,
            },
            consumed,
        ));
    }

    // Header/footer: ~<row>[COL] / _<row>[COL], where COL is absolute sheet column letters.
    if bytes[0] == b'~' || bytes[0] == b'_' {
        let is_header = bytes[0] == b'~';
        let rest = &s[1..];
        let row_digits = rest.chars().take_while(|c| c.is_ascii_digit()).count();
        if row_digits > 0 {
            let row_num: usize = rest[..row_digits].parse().ok()?;
            let row = if is_header {
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

            let after = &rest[row_digits..];
            if after.is_empty() {
                return Some((
                    if is_header {
                        CellAddr::Header { row, col: 0 }
                    } else {
                        CellAddr::Footer { row, col: 0 }
                    },
                    1 + row_digits,
                ));
            }
            let col_digits = after.chars().take_while(|c| c.is_ascii_uppercase()).count();
            if col_digits == 0 {
                return None;
            }
            let col = parse_excel_column(&after[..col_digits])?;
            return Some((
                if is_header {
                    CellAddr::Header { row, col }
                } else {
                    CellAddr::Footer { row, col }
                },
                1 + row_digits + col_digits,
            ));
        }
    }

    // Mirrored margins: [J1..[A1 / ]A1..]J1.
    if bytes[0] == b'[' || bytes[0] == b']' {
        let left_side = bytes[0] == b'[';
        let rest = &s[1..];
        let col_len = rest.chars().take_while(|c| c.is_ascii_uppercase()).count();
        if col_len != 1 {
            return None;
        }
        let col = parse_mirror_margin_column_name(&rest[..col_len], left_side)?;
        let after = &rest[col_len..];
        let row_digits = after.chars().take_while(|c| c.is_ascii_digit()).count();
        if row_digits == 0 {
            return None;
        }
        let main_row: u32 = after[..row_digits].parse().ok()?;
        if main_row == 0 {
            return None;
        }
        let consumed = 1 + col_len + row_digits;
        return Some((
            if left_side {
                CellAddr::Left {
                    col,
                    row: main_row - 1,
                }
            } else {
                CellAddr::Right {
                    col,
                    row: main_row - 1,
                }
            },
            consumed,
        ));
    }

    None
}

/// Parse `A1:B2` at start of `s`; both ends must be main cells. Returns range + consumed length.
pub fn parse_main_range_at(s: &str) -> Option<(crate::grid::MainRange, usize)> {
    let (a, na) = parse_cell_ref_at(s)?;
    let CellAddr::Main { row: ra, col: ca } = a else {
        return None;
    };
    let rest = s.get(na..)?;
    let rest = rest.strip_prefix(':')?;
    let (b, nb) = parse_cell_ref_at(rest)?;
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
        let (a, n) = parse_cell_ref_at("A1").unwrap();
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
        assert_eq!(parse_cell_ref_at("~1A").unwrap().1, 3);
        assert_eq!(parse_cell_ref_at("_1A").unwrap().1, 3);
        assert_eq!(parse_cell_ref_at("[A1").unwrap().1, 3);
        assert_eq!(parse_cell_ref_at("]A1").unwrap().1, 3);
    }

    #[test]
    fn left_margin_is_mirrored_from_the_main_grid() {
        assert_eq!(mirror_margin_column_name(0, true), "J");
        assert_eq!(mirror_margin_column_name(9, true), "A");
        assert_eq!(
            parse_cell_ref_at("[A1").unwrap().0,
            CellAddr::Left { col: 9, row: 0 }
        );
    }

    #[test]
    fn sheet_qualified_cell_refs_parse() {
        let (sheet_id, addr, len) = parse_sheet_qualified_cell_ref_at("$12:A5").unwrap();
        assert_eq!(sheet_id, 12);
        assert_eq!(addr, CellAddr::Main { row: 4, col: 0 });
        assert_eq!(len, 6);
    }
}
