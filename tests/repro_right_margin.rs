use corro::grid::{CellAddr, ColumnAddr, GridBox as Grid, HEADER_ROWS};

/// A control formula in the right-margin key header (`]A~1`) templates the
/// `]A` column itself — the same header→column relationship a main column
/// has. It must NOT be mirrored onto the last main column (that made the
/// value appear in the last data column while `]A` stayed blank).
#[test]
fn right_margin_header_templates_its_own_column_only() {
    let mut g = Grid::from(corro::grid::Grid::new(2, 3));
    let mc = g.main_cols();

    let header_addr = CellAddr::Header {
        row: (HEADER_ROWS - 1) as u32,
        col: ColumnAddr::Right(0),
    };
    g.set(&header_addr, "=B".into());
    assert_eq!(g.get(&header_addr).as_deref(), Some("=B"));

    // `]A`'s own data rows are templated from `]A~1`.
    let right_addr = CellAddr::Right { row: 0, col: 0 };
    assert_eq!(
        corro::formula::export_templated_formula(&g, &right_addr),
        Some("=B1".into())
    );

    // The last main column must stay untouched by a right-margin header.
    let main_addr = CellAddr::Main {
        row: 0,
        col: (mc - 1) as u32,
    };
    assert_eq!(
        corro::formula::export_templated_formula(&g, &main_addr),
        None
    );
}
