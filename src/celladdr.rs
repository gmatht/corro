//! High-level cell/address abstraction.
//!
//! This module introduces a compact, expressive representation for cell
//! references used by higher-level code: RowRegion (HEADER/DATA/FOOTER) and
//! ColRegion (LEFT/DATA/RIGHT/GLOBAL). The numeric payloads are 1-based and
//! correspond to the textual form (`~1` -> Header(1), `C` -> Data(3)).
//!
//! This file provides conversion helpers to/from the existing
//! `crate::grid::CellAddr` so we can migrate incrementally without changing
//! Grid's storage layout.

use crate::grid::{CellAddr, HEADER_ROWS, MARGIN_COLS, MarginIndex};
use std::fmt;

/// Row region with 1-based textual index.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum RowRegion {
    Header(u32), // ~1..~N (1-based)
    Data(u32),   // main data row 1-based
    Footer(u32), // _1.._N (1-based)
}

/// Column region with 1-based indices for ``Data`` and ``Global`` using
/// human/textual numbers; left/right margins are MarginIndex (usize).
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ColRegion {
    Left(MarginIndex),    // < margin index (usize)
    Data(u32),   // A=1, B=2, ... (1-based main-column index)
    Right(MarginIndex),   // > margin index (usize)
    Global(u32), // absolute global column index
}

/// Rich cell reference (parsed form). `raw` contains the original column
/// fragment (letters) if available and `prefix` holds '[' or ']' when used.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct CellRef {
    pub row: RowRegion,
    pub col: ColRegion,
    /// Raw column fragment (letters) when present; useful for serialization.
    pub raw_col_fragment: Option<String>,
    /// Optional prefix character: Some('[') or Some(']') or None.
    pub prefix: Option<char>,
}

impl fmt::Display for RowRegion {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            RowRegion::Header(n) => write!(f, "~{}", n),
            RowRegion::Data(n) => write!(f, "{}", n),
            RowRegion::Footer(n) => write!(f, "_{}", n),
        }
    }
}

impl fmt::Display for ColRegion {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            ColRegion::Left(i) => write!(
                f,
                "[{}",
                super::addr::mirror_margin_column_name(*i, true)
            ),
            ColRegion::Right(i) => write!(
                f,
                "]{}",
                super::addr::mirror_margin_column_name(*i, false)
            ),
            ColRegion::Data(n) => write!(f, "{}", super::addr::excel_column_name(*n as usize - 1)),
            ColRegion::Global(g) => write!(f, "{}", super::addr::excel_column_name(*g as usize)),
        }
    }
}

impl CellRef {
    /// Create a textual (log) representation for this parsed cell reference.
    /// `main_cols` is used when deciding whether a Data column sits inside
    /// the main region vs. the right margin when formatting global forms.
    pub fn to_log_text(&self, _main_cols: usize) -> String {
        let col_text = match &self.col {
            ColRegion::Left(i) => format!(
                "[{}",
                super::addr::mirror_margin_column_name(*i, true)
            ),
            ColRegion::Right(i) => format!(
                "]{}",
                super::addr::mirror_margin_column_name(*i, false)
            ),
            ColRegion::Data(n) => super::addr::excel_column_name(*n as usize - 1),
            ColRegion::Global(g) => super::addr::excel_column_name(*g as usize),
        };
        match &self.row {
            RowRegion::Header(r) => format!("{}~{}", col_text, r),
            RowRegion::Footer(r) => format!("{}_{r}", col_text),
            RowRegion::Data(r) => format!("{}{}", col_text, r),
        }
    }

    /// Convert to the canonical grid::CellAddr using the provided main_cols
    /// hint (needed to compute Right-margin global columns).
    pub fn to_grid_addr(&self, main_cols: usize) -> CellAddr {
        match (&self.row, &self.col) {
            (RowRegion::Header(r), ColRegion::Left(i)) => CellAddr::Header {
                row: ((*r as usize).saturating_sub(1) as u8).saturating_sub(0),
                col: *i as u32,
            },
            (RowRegion::Header(r), ColRegion::Right(i)) => CellAddr::Header {
                row: (HEADER_ROWS - (*r as usize)) as u8,
                col: (MARGIN_COLS + main_cols + *i as usize) as u32,
            },
            (RowRegion::Header(r), ColRegion::Data(n)) => CellAddr::Header {
                row: (HEADER_ROWS - (*r as usize)) as u8,
                col: (MARGIN_COLS + (*n as usize - 1)) as u32,
            },
            (RowRegion::Header(r), ColRegion::Global(g)) => CellAddr::Header {
                row: (HEADER_ROWS - (*r as usize)) as u8,
                col: *g as u32,
            },

            (RowRegion::Footer(r), ColRegion::Left(i)) => CellAddr::Footer {
                row: (*r as u8 - 1),
                col: *i as u32,
            },
            (RowRegion::Footer(r), ColRegion::Right(i)) => CellAddr::Footer {
                row: (*r as u8 - 1),
                col: (MARGIN_COLS + main_cols + *i as usize) as u32,
            },
            (RowRegion::Footer(r), ColRegion::Data(n)) => CellAddr::Footer {
                row: (*r as u8 - 1),
                col: (MARGIN_COLS + (*n as usize - 1)) as u32,
            },
            (RowRegion::Footer(r), ColRegion::Global(g)) => CellAddr::Footer {
                row: (*r as u8 - 1),
                col: *g as u32,
            },

            (RowRegion::Data(rr), ColRegion::Left(i)) => CellAddr::Left {
                col: *i,
                row: (*rr as u32 - 1),
            },
            (RowRegion::Data(rr), ColRegion::Right(i)) => CellAddr::Right {
                col: *i,
                row: (*rr as u32 - 1),
            },
            (RowRegion::Data(rr), ColRegion::Data(n)) => CellAddr::Main {
                row: (*rr as u32 - 1),
                col: (*n as u32 - 1),
            },
            (RowRegion::Data(rr), ColRegion::Global(g)) => {
                // Interpret Global as a global column index: map to left/mid/right
                let gc = *g as usize;
                if gc < MARGIN_COLS {
                    CellAddr::Left {
                        col: gc,
                        row: (*rr as u32 - 1),
                    }
                } else if gc < MARGIN_COLS + main_cols {
                    CellAddr::Main {
                        row: (*rr as u32 - 1),
                        col: (gc - MARGIN_COLS) as u32,
                    }
                } else {
                    CellAddr::Right {
                        col: (gc - MARGIN_COLS - main_cols),
                        row: (*rr as u32 - 1),
                    }
                }
            }
        }
    }

    /// Build a CellRef from an existing grid::CellAddr (useful for serializing).
    pub fn from_grid(addr: &CellAddr, main_cols: usize) -> CellRef {
        match addr {
            CellAddr::Header { row, col } => {
                let c = *col as usize;
                if c < MARGIN_COLS {
                    CellRef {
                        row: RowRegion::Header((HEADER_ROWS - *row as usize) as u32),
                        col: ColRegion::Left(*col as usize),
                        raw_col_fragment: Some(super::addr::mirror_margin_column_name(
                            *col as usize,
                            true,
                        )),
                        prefix: Some('['),
                    }
                } else if c < MARGIN_COLS + main_cols {
                    CellRef {
                        row: RowRegion::Header((HEADER_ROWS - *row as usize) as u32),
                        col: ColRegion::Data((c - MARGIN_COLS) as u32 + 1),
                        raw_col_fragment: Some(super::addr::excel_column_name(c - MARGIN_COLS)),
                        prefix: None,
                    }
                } else {
                    CellRef {
                        row: RowRegion::Header((HEADER_ROWS - *row as usize) as u32),
                        col: ColRegion::Right((c - MARGIN_COLS - main_cols) as usize),
                        raw_col_fragment: Some(super::addr::mirror_margin_column_name(
                            c - MARGIN_COLS - main_cols,
                            false,
                        )),
                        prefix: Some(']'),
                    }
                }
            }
            CellAddr::Footer { row, col } => {
                let c = *col as usize;
                if c < MARGIN_COLS {
                    CellRef {
                        row: RowRegion::Footer(*row as u32 + 1),
                        col: ColRegion::Left(*col as usize),
                        raw_col_fragment: Some(super::addr::mirror_margin_column_name(
                            *col as usize,
                            true,
                        )),
                        prefix: Some('['),
                    }
                } else if c < MARGIN_COLS + main_cols {
                    CellRef {
                        row: RowRegion::Footer(*row as u32 + 1),
                        col: ColRegion::Data((c - MARGIN_COLS) as u32 + 1),
                        raw_col_fragment: Some(super::addr::excel_column_name(c - MARGIN_COLS)),
                        prefix: None,
                    }
                } else {
                    CellRef {
                        row: RowRegion::Footer(*row as u32 + 1),
                        col: ColRegion::Right((c - MARGIN_COLS - main_cols) as usize),
                        raw_col_fragment: Some(super::addr::mirror_margin_column_name(
                            c - MARGIN_COLS - main_cols,
                            false,
                        )),
                        prefix: Some(']'),
                    }
                }
            }
            CellAddr::Main { row, col } => CellRef {
                row: RowRegion::Data(*row + 1),
                col: ColRegion::Data(*col + 1),
                raw_col_fragment: Some(super::addr::excel_column_name(*col as usize)),
                prefix: None,
            },
            CellAddr::Left { col, row } => CellRef {
                row: RowRegion::Data(*row + 1),
                col: ColRegion::Left(*col),
                raw_col_fragment: Some(super::addr::mirror_margin_column_name(*col as usize, true)),
                prefix: Some('['),
            },
            CellAddr::Right { col, row } => CellRef {
                row: RowRegion::Data(*row + 1),
                col: ColRegion::Right(*col),
                raw_col_fragment: Some(super::addr::mirror_margin_column_name(
                    *col as usize,
                    false,
                )),
                prefix: Some(']'),
            },
        }
    }

    /// Parse a cell reference at the start of `s` and return a parsed
    /// CellRef (1-based indices) plus the number of bytes consumed.
    /// Unprefixed Excel column names are interpreted as Data columns
    /// (map to main/data-region columns). Bracketed prefixes (`[`, `]`)
    /// and explicit Global forms are supported as before.
    pub fn parse_at(s: &str) -> Option<(CellRef, usize)> {
        let bytes = s.as_bytes();
        if bytes.is_empty() {
            return None;
        }

        let (prefix, rest, prefix_len) = match bytes[0] {
            b'[' => (Some('['), &s[1..], 1usize),
            b']' => (Some(']'), &s[1..], 1usize),
            _ => (None, s, 0usize),
        };

        // Accept one-or-more uppercase letters for the column fragment
        // regardless of whether a bracket prefix was present. The older
        // implementation required exactly one letter when a bracket was
        // used; that prevented multi-letter mirror margin names like "AA".
        let col_len = {
            let len = rest.chars().take_while(|c| c.is_ascii_uppercase()).count();
            if len == 0 {
                return None;
            }
            len
        };
        let col_name = &rest[..col_len];
        let after = &rest[col_len..];

        // header/footer form like C~1 or A_1
        if let Some(marker) = after.chars().next().filter(|c| *c == '~' || *c == '_') {
            let row_digits = after[1..]
                .chars()
                .take_while(|c| c.is_ascii_digit())
                .count();
            if row_digits == 0 {
                return None;
            }
            let row_num: usize = after[1..1 + row_digits].parse().ok()?;
            let row_region = if marker == '~' {
                // Header: keep textual 1-based row number
                RowRegion::Header(row_num as u32)
            } else {
                RowRegion::Footer(row_num as u32)
            };

            let col_region = match prefix {
                Some('[') => {
                    let idx = super::addr::parse_mirror_margin_column_name(col_name, true)?;
                    ColRegion::Left(idx)
                }
                Some(']') => {
                    // Try mirror name first; otherwise fall back to treating the
                    // token as an Excel column name (legacy behavior).
                    if let Some(idx) = super::addr::parse_mirror_margin_column_name(col_name, false)
                    {
                        ColRegion::Right(idx)
                    } else {
                        let excel_idx = super::addr::parse_excel_column(col_name)?; // 0-based
                        ColRegion::Global(excel_idx as u32)
                    }
                }
                Some(_) | None => {
                    // Treat unprefixed Excel letters in header/footer forms as
                    // Data columns (map into the sheet's main/data region).
                    // The historic "Global" interpretation is removed; an
                    // unprefixed column always refers to the main/data
                    // column named by the letters.
                    let excel_idx = super::addr::parse_excel_column(col_name)?; // 0-based
                    ColRegion::Data(excel_idx as u32 + 1)
                }
            };

            return Some((
                CellRef {
                    row: row_region,
                    col: col_region,
                    raw_col_fragment: Some(col_name.to_string()),
                    prefix,
                },
                prefix_len + col_len + 1 + row_digits,
            ));
        }

        // A1-style main cell or prefixed margin cell
        let row_digits = after.chars().take_while(|c| c.is_ascii_digit()).count();
        if row_digits == 0 {
            return None;
        }
        let row_num: u32 = after[..row_digits].parse().ok()?;
        let cref = match prefix {
            Some('[') => CellRef {
                row: RowRegion::Data(row_num),
                col: ColRegion::Left(super::addr::parse_mirror_margin_column_name(
                    col_name, true,
                )?),
                raw_col_fragment: Some(col_name.to_string()),
                prefix: Some('['),
            },
            Some(']') => {
                // if mirror parse fails, treat as Global token
                if let Some(idx) = super::addr::parse_mirror_margin_column_name(col_name, false) {
                    CellRef {
                        row: RowRegion::Data(row_num),
                        col: ColRegion::Right(idx),
                        raw_col_fragment: Some(col_name.to_string()),
                        prefix: Some(']'),
                    }
                } else {
                    // fallback: treat as data with excel column name
                    let excel_idx = super::addr::parse_excel_column(col_name)?;
                    CellRef {
                        row: RowRegion::Data(row_num),
                        col: ColRegion::Data(excel_idx as u32 + 1),
                        raw_col_fragment: Some(col_name.to_string()),
                        prefix: Some(']'),
                    }
                }
            }
            Some(_) | None => {
                let excel_idx = super::addr::parse_excel_column(col_name)?;
                CellRef {
                    row: RowRegion::Data(row_num),
                    // Always interpret unprefixed Excel letters as Data columns.
                    col: ColRegion::Data(excel_idx as u32 + 1),
                    raw_col_fragment: Some(col_name.to_string()),
                    prefix: None,
                }
            }
        };

        Some((cref, prefix_len + col_len + row_digits))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::grid::Grid;

    #[test]
    fn roundtrip_main_cell() {
        let g = Grid::new(2, 3);
        let addr = CellAddr::Main { row: 1, col: 2 };
        let cref = CellRef::from_grid(&addr, g.main_cols());
        assert_eq!(cref.row, RowRegion::Data(2));
        assert_eq!(cref.col, ColRegion::Data(3));
        let back = cref.to_grid_addr(g.main_cols());
        assert_eq!(back, addr);
    }

    #[test]
    fn roundtrip_header_cell() {
        let mut g = Grid::new(1, 1);
        g.set(&CellAddr::Header { row: 25, col: 12 }, "TOTAL".into());
        let cref = CellRef::from_grid(&CellAddr::Header { row: 25, col: 12 }, g.main_cols());
        assert!(matches!(cref.col, ColRegion::Data(_)));
        let back = cref.to_grid_addr(g.main_cols());
        assert_eq!(back, CellAddr::Header { row: 25, col: 12 });
    }
}
