//! Shared cell-address parsing (Excel columns, global column suffixes, single-cell refs).

use crate::grid::{CellAddr, MARGIN_COLS};

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

/// Resolve `COL` after `^X,` or `_X,`: `<n`, `>n`, or main column letters.
pub fn resolve_global_col(col_part: &str) -> Option<u32> {
    let bytes = col_part.as_bytes();
    if bytes.is_empty() {
        return None;
    }
    if bytes[0] == b'<' {
        let n: u32 = col_part[1..].parse().ok()?;
        return Some((MARGIN_COLS as u32).checked_sub(1)?.checked_sub(n)?);
    }
    if bytes[0] == b'>' {
        let n: u32 = col_part[1..].parse().ok()?;
        return Some(n);
    }
    if bytes[0].is_ascii_uppercase() {
        let col = parse_excel_column(col_part)?;
        return Some(MARGIN_COLS as u32 + col);
    }
    None
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

    // Header: ^X or ^X,COL (legacy logs also use `:` and raw global column numbers)
    if bytes[0] == b'^' && bytes.len() >= 2 {
        let letter = bytes[1];
        if !letter.is_ascii_uppercase() {
            return None;
        }
        let row = b'Z' - letter;
        if s.len() == 2 {
            return Some((CellAddr::Header { row, col: 0 }, 2));
        }
        if !matches!(bytes.get(2), Some(b',') | Some(b':')) {
            return None;
        }
        let col_part = &s[3..];
        let cf = parse_global_col_fragment(col_part)?;
        return Some((CellAddr::Header { row, col: cf.col }, 3 + cf.consumed));
    }

    // Footer: _X or _X,COL (legacy logs also use `:` and raw global column numbers)
    if bytes[0] == b'_' && bytes.len() >= 2 {
        let letter = bytes[1];
        if !letter.is_ascii_uppercase() {
            return None;
        }
        let row = letter - b'A';
        if s.len() == 2 {
            return Some((CellAddr::Footer { row, col: 0 }, 2));
        }
        if !matches!(bytes.get(2), Some(b',') | Some(b':')) {
            return None;
        }
        let col_part = &s[3..];
        let cf = parse_global_col_fragment(col_part)?;
        return Some((CellAddr::Footer { row, col: cf.col }, 3 + cf.consumed));
    }

    // Left: <N,R (legacy logs also use `:`)
    if bytes[0] == b'<' {
        let rest = &s[1..];
        let sep = rest.find(|c| c == ',' || c == ':')?;
        let margin_col: u8 = rest[..sep].parse().ok()?;
        let after = &rest[sep + 1..];
        let row_digits = after.chars().take_while(|c| c.is_ascii_digit()).count();
        if row_digits == 0 {
            return None;
        }
        let main_row: u32 = after[..row_digits].parse().ok()?;
        if main_row == 0 {
            return None;
        }
        let consumed = 1 + sep + 1 + row_digits;
        return Some((
            CellAddr::Left {
                col: margin_col,
                row: main_row - 1,
            },
            consumed,
        ));
    }

    // Right: >N,R (legacy logs also use `:`)
    if bytes[0] == b'>' {
        let rest = &s[1..];
        let sep = rest.find(|c| c == ',' || c == ':')?;
        let margin_col: u8 = rest[..sep].parse().ok()?;
        let after = &rest[sep + 1..];
        let row_digits = after.chars().take_while(|c| c.is_ascii_digit()).count();
        if row_digits == 0 {
            return None;
        }
        let main_row: u32 = after[..row_digits].parse().ok()?;
        if main_row == 0 {
            return None;
        }
        let consumed = 1 + sep + 1 + row_digits;
        return Some((
            CellAddr::Right {
                col: margin_col,
                row: main_row - 1,
            },
            consumed,
        ));
    }

    None
}

struct ColFragment {
    col: u32,
    consumed: usize,
}

fn parse_global_col_fragment(s: &str) -> Option<ColFragment> {
    let bytes = s.as_bytes();
    if bytes.is_empty() {
        return None;
    }
    if bytes[0] == b'<' {
        let rest = &s[1..];
        let n_digits = rest.bytes().take_while(|b| b.is_ascii_digit()).count();
        if n_digits == 0 {
            return None;
        }
        let n: u32 = rest[..n_digits].parse().ok()?;
        let col = (MARGIN_COLS as u32).checked_sub(1)?.checked_sub(n)?;
        return Some(ColFragment {
            col,
            consumed: 1 + n_digits,
        });
    }
    if bytes[0] == b'>' {
        let rest = &s[1..];
        let n_digits = rest.bytes().take_while(|b| b.is_ascii_digit()).count();
        if n_digits == 0 {
            return None;
        }
        let v: u32 = rest[..n_digits].parse().ok()?;
        return Some(ColFragment {
            col: v,
            consumed: 1 + n_digits,
        });
    }
    if bytes[0].is_ascii_digit() {
        let n_digits = s.bytes().take_while(|b| b.is_ascii_digit()).count();
        if n_digits == 0 {
            return None;
        }
        let col: u32 = s[..n_digits].parse().ok()?;
        return Some(ColFragment {
            col,
            consumed: n_digits,
        });
    }
    if bytes[0].is_ascii_uppercase() {
        let n = s.bytes().take_while(|b| b.is_ascii_uppercase()).count();
        let name = &s[..n];
        let col = parse_excel_column(name)?;
        return Some(ColFragment {
            col: MARGIN_COLS as u32 + col,
            consumed: n,
        });
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
        assert_eq!(parse_cell_ref_at("^A:10").unwrap().1, 5);
        assert_eq!(parse_cell_ref_at("_B:10").unwrap().1, 5);
        assert_eq!(parse_cell_ref_at("<0:1").unwrap().1, 4);
        assert_eq!(parse_cell_ref_at(">0:1").unwrap().1, 4);
    }
}
