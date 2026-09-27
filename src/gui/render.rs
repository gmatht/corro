use std::collections::HashMap;

use crate::grid::{CellAddr, ColumnAddr, GridBox, NumberFormat};
use crate::ops::AggFunc;
use crate::ui_core;
use unicode_width::UnicodeWidthStr;

use super::compute::{self, right_col_agg, CellDisplayStyle};

/// Abstract sink for cell data produced by `fill_cells`.
/// Each backend implements this to route data to its native rendering system.
pub trait CellSink {
    fn set_cell(&mut self, display_row: u32, display_col: u32, text: &str);
    fn set_cell_style(&mut self, display_row: u32, display_col: u32, style: CellDisplayStyle);
    fn set_raw_cell(&mut self, display_row: u32, display_col: u32, text: &str);
    fn set_cursor(&mut self, display_row: u32, display_col: u32);
}

/// Populate a CellSink with cell data for the given viewport.
/// Generic version shared by pancurses and canvas backends.
#[allow(clippy::too_many_arguments)]
pub fn fill_cells(
    sink: &mut dyn CellSink,
    display_rows: &[usize],
    col_ixs: &[usize],
    col_widths: &HashMap<usize, usize>,
    g: &GridBox,
    hr: usize,
    mr: usize,
    mc: usize,
    lm: usize,
    data_width: usize,
    display_cursor_row: usize,
    display_cursor_col: usize,
    row_agg_func: &[Option<AggFunc>],
) {
    for (ri, &logical_row) in display_rows.iter().enumerate() {
        let main_row = if logical_row >= hr && logical_row < hr + mr {
            Some((logical_row - hr) as u32)
        } else {
            None
        };
        let footer_row_idx = if logical_row >= hr + mr {
            Some((logical_row - hr - mr) as u32)
        } else {
            None
        };
        let row_agg = row_agg_func[ri];

        let mut col_ix = 0usize;
        while col_ix < col_ixs.len() {
            let c = col_ixs[col_ix];
            let addr = cell_addr_for_coords(logical_row, hr, mr, c, lm, mc);

            let cw = g.col_width(c).max(1);

            let rca = right_col_agg(g, c);
            let is_cursor_cell = logical_row == display_cursor_row && c == display_cursor_col;

            let cell_info = compute::compute_cell_info(
                g, &addr, is_cursor_cell, row_agg, main_row, footer_row_idx,
                rca, c, lm, mc, mr,
            );

            let formatted = &cell_info.formatted;
            let fw = formatted.width();
            let align = ui_core::effective_cell_align(g, &addr, formatted);

            let allow_spill = fw > cw
                && (align.is_none() || align == Some(crate::grid::TextAlign::Left))
                && !cell_info.is_agg_cell;

            let mut did_spill = false;
            if allow_spill {
                let mut next_ix = col_ix + 1;
                let mut total_spill_gaps = cw;
                let mut narrow_spill = cw;
                let mut beyond_boundary = false;

                while next_ix < col_ixs.len() {
                    let c_next = col_ixs[next_ix];
                    let prev_vp = next_ix - 1;
                    let prev_col = col_ixs[prev_vp];
                    let trailing = ui_core::inter_column_trailing_after_data_cell(
                        prev_vp, prev_col, col_ixs, lm, mc, col_ixs.contains(&(lm + mc)),
                    );
                    match trailing {
                        ui_core::InterColumnTrailing::AsciiSpace => {
                            total_spill_gaps += 1;
                            if !beyond_boundary {
                                narrow_spill += 1;
                            }
                        }
                        ui_core::InterColumnTrailing::PipeAndSpace => {
                            total_spill_gaps += 2;
                            if !beyond_boundary {
                                narrow_spill += 2;
                            }
                        }
                        _ => {}
                    }
                    if !cell_display_at(g, logical_row, hr, mr, c_next, lm, mc).trim().is_empty() {
                        break;
                    }
                    if c_next == lm {
                        let render_w = ui_core::visible_cols_render_width(g, col_ixs);
                        let right_gap = data_width.saturating_sub(render_w);
                        total_spill_gaps = total_spill_gaps.saturating_add(right_gap);
                        narrow_spill = narrow_spill.saturating_add(right_gap);
                        break;
                    }
                    if c_next == lm + mc {
                        beyond_boundary = true;
                        let cw_rm = *col_widths.get(&c_next).unwrap_or(&4);
                        total_spill_gaps += cw_rm;
                        if next_ix + 1 < col_ixs.len() {
                            let t = ui_core::inter_column_trailing_after_data_cell(
                                next_ix, c_next, col_ixs,
                                lm, mc, col_ixs.contains(&(lm + mc)),
                            );
                            match t {
                                ui_core::InterColumnTrailing::AsciiSpace => total_spill_gaps += 1,
                                ui_core::InterColumnTrailing::PipeAndSpace => total_spill_gaps += 2,
                                _ => {}
                            }
                        }
                        let mut wide_ix = next_ix + 1;
                        while wide_ix < col_ixs.len() {
                            let cw_more = *col_widths.get(&col_ixs[wide_ix]).unwrap_or(&4);
                            total_spill_gaps += cw_more;
                            if wide_ix + 1 < col_ixs.len() {
                                let t = ui_core::inter_column_trailing_after_data_cell(
                                    wide_ix, col_ixs[wide_ix], col_ixs,
                                    lm, mc, col_ixs.contains(&(lm + mc)),
                                );
                                match t {
                                    ui_core::InterColumnTrailing::AsciiSpace => total_spill_gaps += 1,
                                    ui_core::InterColumnTrailing::PipeAndSpace => total_spill_gaps += 2,
                                    _ => {}
                                }
                            }
                            wide_ix += 1;
                        }
                        let render_w = ui_core::visible_cols_render_width(g, col_ixs);
                        let right_gap = data_width.saturating_sub(render_w);
                        total_spill_gaps = total_spill_gaps.saturating_add(right_gap);
                        break;
                    }
                    let cw_next = *col_widths.get(&c_next).unwrap_or(&4);
                    total_spill_gaps += cw_next;
                    if !beyond_boundary {
                        narrow_spill += cw_next;
                    }
                    next_ix += 1;
                }

                if !beyond_boundary && next_ix >= col_ixs.len() {
                    let render_w = ui_core::visible_cols_render_width(g, col_ixs);
                    let right_gap = data_width.saturating_sub(render_w);
                    narrow_spill = narrow_spill.saturating_add(right_gap);
                    total_spill_gaps = narrow_spill;
                }

                let use_wide = fw > narrow_spill;
                let pipe_gap = if !use_wide && beyond_boundary { 2 } else { 0 };
                let pad_spill = if use_wide { total_spill_gaps } else { narrow_spill.saturating_sub(pipe_gap) };

                if pad_spill > cw {
                    did_spill = true;
                    let store_text = if formatted.trim().is_empty() {
                        String::new()
                    } else {
                        formatted.clone()
                    };
                    // Always store, even when empty: the sink is a persistent
                    // (display_row, display_col) map, so skipping empty cells
                    // leaves the *previous* frame's text visible once the
                    // viewport scrolls (e.g. the footer's "TOTAL" key cell
                    // lingering on a data row that now occupies that screen
                    // line). Refill must be authoritative for every visible
                    // position.
                    sink.set_cell(ri as u32, c as u32, &store_text);
                    sink.set_cell_style(ri as u32, c as u32, cell_info.style);
                    if !store_text.is_empty() {
                        if let Some(ref raw_val) = cell_info.raw_value {
                            sink.set_raw_cell(ri as u32, c as u32, raw_val);
                        }
                    }

                    let advance_to = if use_wide && beyond_boundary {
                        col_ixs.len()
                    } else {
                        next_ix
                    };

                    for skip in (col_ix + 1)..advance_to {
                        let skip_col = col_ixs[skip];
                        sink.set_cell(ri as u32, skip_col as u32, "");
                        sink.set_cell_style(ri as u32, skip_col as u32, CellDisplayStyle::Default);
                    }
                    col_ix = advance_to;
                }
            }

            if !did_spill {
                let display_text = if fw > cw {
                    truncate_and_align(g, &addr, formatted, cw, align)
                } else {
                    ui_core::align_cell_display(formatted.to_string(), cw, align)
                };

                let store_text = if display_text.trim().is_empty() {
                    String::new()
                } else {
                    display_text
                };
                // Always store (see the spill branch): an empty cell must
                // overwrite the previous frame's text at this display
                // position, or stale content survives a scroll.
                sink.set_cell(ri as u32, c as u32, &store_text);
                sink.set_cell_style(ri as u32, c as u32, cell_info.style);
                if !store_text.is_empty() {
                    if let Some(ref raw_val) = cell_info.raw_value {
                        sink.set_raw_cell(ri as u32, c as u32, raw_val);
                    }
                }
                col_ix += 1;
            }
        }
    }
}

/// Truncate and align text to fit within a column width.
fn truncate_and_align(
    g: &GridBox,
    addr: &CellAddr,
    formatted: &str,
    cw: usize,
    align: Option<crate::grid::TextAlign>,
) -> String {
    let cell_fmt = g.format_for_addr(addr);
    let rational_hint = if matches!(cell_fmt.number, None | Some(NumberFormat::Rational | NumberFormat::DecimalGeneric))
        && ui_core::would_ellipsis_hide_decimal_point(formatted, cw)
    {
        crate::formula::effective_numeric(g, addr, &mut Vec::new(), &mut 10_000usize)
            .map(|n| n.to_f64())
            .filter(|v| v.is_finite())
    } else {
        None
    };
    let exp_preferred = if ui_core::would_ellipsis_hide_decimal_point(formatted, cw) {
        ui_core::exponential_numeric_display_with_hint(formatted, cw, rational_hint)
    } else {
        None
    };
    let inner = exp_preferred
        .or_else(|| ui_core::shrink_numeric_display(formatted, cw))
        .or_else(|| ui_core::exponential_numeric_display(formatted, cw))
        .unwrap_or_else(|| ui_core::truncate_with_ellipsis(formatted, cw));
    ui_core::align_cell_display(inner, cw, align)
}

/// Get the effective display value for a cell at a given (logical_row, global_col).
fn cell_display_at(g: &GridBox, logical_row: usize, hr: usize, mr: usize, c: usize, lm: usize, mc: usize) -> String {
    let addr = cell_addr_for_coords(logical_row, hr, mr, c, lm, mc);
    crate::formula::cell_effective_display(g, &addr)
}

/// Build a CellAddr from logical row/col coordinates.
fn cell_addr_for_coords(logical_row: usize, hr: usize, mr: usize, c: usize, lm: usize, mc: usize) -> CellAddr {
    if logical_row < hr {
        let hdr_row = logical_row as u32;
        if c < lm {
            CellAddr::Header { row: hdr_row, col: ColumnAddr::Left(c) }
        } else if c < lm + mc {
            CellAddr::Header { row: hdr_row, col: ColumnAddr::Main((c - lm) as u32) }
        } else {
            CellAddr::Header { row: hdr_row, col: ColumnAddr::Right(c - lm - mc) }
        }
    } else if logical_row < hr + mr {
        let main_row = (logical_row - hr) as u32;
        if c < lm {
            CellAddr::Left { row: main_row, col: c }
        } else if c < lm + mc {
            CellAddr::Main { row: main_row, col: (c - lm) as u32 }
        } else {
            CellAddr::Right { row: main_row, col: c - lm - mc }
        }
    } else {
        let ftr_row = (logical_row - hr - mr) as u32;
        if c < lm {
            CellAddr::Footer { row: ftr_row, col: ColumnAddr::Left(c) }
        } else if c < lm + mc {
            CellAddr::Footer { row: ftr_row, col: ColumnAddr::Main((c - lm) as u32) }
        } else {
            CellAddr::Footer { row: ftr_row, col: ColumnAddr::Right(c - lm - mc) }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    /// Recording sink: keeps the same persistent-map semantics the native
    /// backends use (insert per position), so a stale entry survives unless
    /// `fill_cells` overwrites it.
    #[derive(Default)]
    struct RecordingSink {
        cells: HashMap<(u32, u32), String>,
        styles: HashMap<(u32, u32), CellDisplayStyle>,
    }

    impl CellSink for RecordingSink {
        fn set_cell(&mut self, r: u32, c: u32, text: &str) {
            self.cells.insert((r, c), text.to_string());
        }
        fn set_cell_style(&mut self, r: u32, c: u32, style: CellDisplayStyle) {
            self.styles.insert((r, c), style);
        }
        fn set_raw_cell(&mut self, _r: u32, _c: u32, _t: &str) {}
        fn set_cursor(&mut self, _r: u32, _c: u32) {}
    }

    fn fill(g: &GridBox, display_rows: &[usize], cursor_row: usize, sink: &mut RecordingSink) {
        let lm = crate::grid::MARGIN_COLS;
        let mc = g.main_cols();
        let mr = g.main_rows();
        let hr = crate::grid::HEADER_ROWS;
        let col_ixs: Vec<usize> = (0..(lm + mc + crate::grid::MARGIN_COLS)).collect();
        let col_widths: HashMap<usize, usize> =
            col_ixs.iter().map(|&c| (c, g.col_width(c).max(1))).collect();
        let ragg = compute::compute_row_agg_func(g, display_rows, hr, mr);
        fill_cells(
            sink, display_rows, &col_ixs, &col_widths, g, hr, mr, mc, lm, 80,
            cursor_row, lm, &ragg,
        );
    }

    /// Regression: as the viewport scrolls, the screen position that used to
    /// be the footer row (carrying the seeded margin-key `TOTAL`) becomes a
    /// data row. `fill_cells` must overwrite that position with the now-empty
    /// cell, or the stale "TOTAL" lingers on a data row — the reported
    /// "scrolling shows TOTAL in rows that shouldn't have it" bug.
    #[test]
    fn refill_overwrites_stale_cells_after_scroll() {
        let mut sheet = crate::ops::SheetState::new_seeded();
        // Three main rows so the viewport can scroll over them.
        sheet.grid.set_main_size(3, 1);
        let g = &sheet.grid;
        let hr = crate::grid::HEADER_ROWS;
        let mr = g.main_rows();

        // Frame 1: cursor in the footer — its display rows hold the margin
        // key ("TOTAL") at the left-margin position.
        let footer_rows: Vec<usize> = (0..5).map(|i| hr + mr + i).collect();
        let mut sink = RecordingSink::default();
        let cursor_footer = hr + mr;
        fill(g, &footer_rows, cursor_footer, &mut sink);
        let total_at = (0u32, (crate::grid::MARGIN_COLS - 1) as u32);
        assert_eq!(
            sink.cells.get(&total_at).map(|s| s.trim()),
            Some("TOTAL"),
            "frame 1: display row 0 is the footer row, whose margin key is TOTAL"
        );

        // Frame 2: same screen positions, but now they are plain data rows.
        let data_rows: Vec<usize> = (0..5).map(|i| hr + i).collect();
        fill(g, &data_rows, hr, &mut sink);
        assert_eq!(
            sink.cells.get(&total_at).map(|s| s.trim()),
            Some(""),
            "frame 2: the reused screen row must be cleared (stale TOTAL \
             would otherwise show on a data row)"
        );
    }

    /// The cursor style must be re-emitted each frame at the cursor's screen
    /// position (the reported "selection disappears on scroll").
    #[test]
    fn refill_reemits_cursor_style_every_frame() {
        let mut sheet = crate::ops::SheetState::new_seeded();
        sheet.grid.set_main_size(3, 1);
        let g = &sheet.grid;
        let hr = crate::grid::HEADER_ROWS;
        let rows: Vec<usize> = (0..5).map(|i| hr + i).collect();

        let mut sink = RecordingSink::default();
        fill(g, &rows, hr + 2, &mut sink);
        assert_eq!(
            sink.styles.get(&(2, crate::grid::MARGIN_COLS as u32)),
            Some(&CellDisplayStyle::Cursor),
            "cursor cell must be styled Cursor"
        );
        // Scrolling to a different row must move the cursor style, not drop it.
        fill(g, &rows, hr, &mut sink);
        assert_eq!(
            sink.styles.get(&(2, crate::grid::MARGIN_COLS as u32)),
            Some(&CellDisplayStyle::Default),
            "the old cursor position must be reset to Default"
        );
        assert_eq!(
            sink.styles.get(&(0, crate::grid::MARGIN_COLS as u32)),
            Some(&CellDisplayStyle::Cursor),
            "the new cursor position must carry Cursor"
        );
    }
}

/// Regression: a control formula in the right-margin key header (`]A~1`)
/// must render in the **`]A` column**, not the last main column. Reported:
/// typing `=A*B -- AB` in `]A~1` displayed `A*B` (its value) in the last
/// data column instead of in `]A`.
#[cfg(test)]
mod right_margin_display_tests {
    use super::*;

    #[test]
    fn right_margin_header_template_renders_in_the_right_margin_column() {
        let mut g = crate::grid::GridBox::from(crate::grid::Grid::new(3, 3));
        let hr = crate::grid::HEADER_ROWS;
        let lm = crate::grid::MARGIN_COLS;
        let mr = g.main_rows();
        let mc = g.main_cols();

        // A1 = 3, B1 = 4; the header of ]A holds the template.
        g.set(&CellAddr::Main { row: 0, col: 0 }, "3".into());
        g.set(&CellAddr::Main { row: 0, col: 1 }, "4".into());
        let right_a_hdr = CellAddr::Header { row: (hr - 1) as u32, col: ColumnAddr::Right(0) };
        g.set(&right_a_hdr, "=A*B -- AB".into());

        let row0 = hr;                    // first data row
        let last_main_col = lm + mc - 1;  // C
        let right_a_col = lm + mc;        // ]A

        // ]A shows the computed value A1*B1 = 12.
        assert_eq!(
            cell_display_at(&g, row0, hr, mr, right_a_col, lm, mc),
            "12",
            "]A must display its own header template's value"
        );
        // The last main column stays empty -- no leak.
        assert_eq!(
            cell_display_at(&g, row0, hr, mr, last_main_col, lm, mc),
            "",
            "the ]A header must not render in the last data column"
        );
        // The ]A header cell itself shows its label (looked up by its true
        // address; the header band renders row 0 as the visible header).
        assert_eq!(
            crate::formula::cell_effective_display(&g, &right_a_hdr),
            "AB"
        );
    }
}

/// A [`CellSink`] that writes the viewport into an rswidgets `GridView`.
///
/// This is the **one** sink both hosts share. Before it, each backend had its
/// own: `SpreadsheetSink` written against the pancurses widget
/// (`pnc_backend.rs`), and `GuiCanvasSink` written against four `HashMap`s
/// (`gui_backend.rs`). Those two differed only in *where* they put the cells —
/// and the GUI's four maps (`cells`, `styles`, `raw_values`, `cursor_pos`) are
/// exactly the fields `SpreadsheetModel` already has, so the difference was not
/// real. `GridView` is that model plus the surface to draw it on, so routing
/// both backends through this adapter removes the last genuine duplication
/// between them.
///
/// The style conversion lives here rather than in each backend because the u8
/// style bits are the model's vocabulary (mirroring
/// `CellDisplayStyle::to_pancurses_style`), not a backend's.
pub struct GridSink<'a> {
    grid: &'a rswidgets::gridview::GridView,
}

impl<'a> GridSink<'a> {
    pub fn new(grid: &'a rswidgets::gridview::GridView) -> Self {
        GridSink { grid }
    }

    /// The style-bit conversion, shared so a backend cannot drift from the
    /// model's numbering.
    pub fn style_bits(style: CellDisplayStyle) -> u8 {
        match style {
            CellDisplayStyle::Default => 0,
            CellDisplayStyle::Cursor => 1,
            CellDisplayStyle::Aggregate => 2,
            CellDisplayStyle::FooterAggregate => 3,
            CellDisplayStyle::Selected => 4,
            CellDisplayStyle::ActiveHeader => 5,
            CellDisplayStyle::InactiveHeader => 6,
            CellDisplayStyle::Hyperlink => 7,
        }
    }
}

impl CellSink for GridSink<'_> {
    fn set_cell(&mut self, display_row: u32, display_col: u32, text: &str) {
        self.grid.set_cell(display_row, display_col, text);
    }
    fn set_cell_style(&mut self, display_row: u32, display_col: u32, style: CellDisplayStyle) {
        self.grid.set_cell_style(display_row, display_col, Self::style_bits(style));
    }
    fn set_raw_cell(&mut self, display_row: u32, display_col: u32, text: &str) {
        self.grid.set_raw_cell(display_row, display_col, text);
    }
    fn set_cursor(&mut self, display_row: u32, display_col: u32) {
        self.grid.set_cursor(display_row, display_col);
    }
}

#[cfg(test)]
mod grid_sink_tests {
    use super::*;
    use rswidgets::gridview::GridView;

    /// The adapter must forward all four `CellSink` operations into the grid's
    /// model, and map the style bits to the model's own numbering.
    #[test]
    fn grid_sink_forwards_into_the_model() {
        let grid = GridView::new(4, 3);
        {
            let mut sink = GridSink::new(&grid);
            sink.set_cell(1, 2, "hi");
            sink.set_cell_style(1, 2, CellDisplayStyle::Hyperlink);
            sink.set_raw_cell(1, 2, "hi");
            sink.set_cursor(1, 2);
        }
        let m = grid.model();
        let m = m.borrow();
        assert_eq!(Some(&"hi".to_string()), m.cells.get(&(1, 2)));
        assert_eq!(Some(&7u8), m.cell_styles.get(&(1, 2)), "hyperlink is style 7");
        assert_eq!(Some(&"hi".to_string()), m.raw_cells.get(&(1, 2)));
        assert_eq!((1, 2), (m.cursor_row, m.cursor_col));
    }

    /// Every style must map, and the numbering must match
    /// `CellDisplayStyle::to_pancurses_style` — this is the contract the
    /// renderer reads, so a silent renumbering would repaint every cell.
    #[test]
    fn style_bits_match_the_models_numbering() {
        for (style, bits) in [
            (CellDisplayStyle::Default, 0u8),
            (CellDisplayStyle::Cursor, 1),
            (CellDisplayStyle::Aggregate, 2),
            (CellDisplayStyle::FooterAggregate, 3),
            (CellDisplayStyle::Selected, 4),
            (CellDisplayStyle::ActiveHeader, 5),
            (CellDisplayStyle::InactiveHeader, 6),
            (CellDisplayStyle::Hyperlink, 7),
        ] {
            assert_eq!(bits, GridSink::style_bits(style));
            #[cfg(feature = "pancurses")]
            assert_eq!(bits, style.to_pancurses_style(), "must agree with the existing map");
        }
    }

    /// `fill_cells` must drive the shared sink end to end: the generic
    /// renderer writes through it into the model, which is what both hosts
    /// need.
    ///
    /// NOTE on `hr`: `grid::HEADER_ROWS` is a **sentinel** (999_999_999), not a
    /// count — it is the logical row index of the topmost header, used so
    /// headers can grow upward without renumbering. `display_rows` are absolute
    /// logical rows starting at `hr`, and `mr` is the *count* of main rows. So
    /// `compute_row_agg_func(g, display_rows, hr, mr)` gets `hr` as the
    /// sentinel and `mr` as the count, which is why passing a small `hr` here
    /// would classify every row as a header. Getting this wrong is not a
    /// silent mis-render: a non-sentinel `hr` makes the `lr < hr + mr` branch
    /// run for ~10^9 rows and the test hangs. That trap is worth a comment.
    #[test]
    fn fill_cells_writes_through_the_shared_sink() {
        use std::collections::HashMap;

        let mut sheet = crate::ops::SheetState::new_seeded();
        sheet.grid.set_main_size(3, 2);
        let g = &sheet.grid;
        let hr = crate::grid::HEADER_ROWS; // sentinel, not a count
        let mr = g.main_rows();
        let mc = g.main_cols();

        // Absolute logical rows: the header band, then the main rows.
        let display_rows: Vec<usize> = (0..mr).map(|i| hr + i).collect();
        let col_ixs: Vec<usize> = (0..mc).collect();
        let widths: HashMap<usize, usize> =
            col_ixs.iter().map(|&c| (c, g.col_width(c).max(1))).collect();
        let agg = compute::compute_row_agg_func(g, &display_rows, hr, mr);

        let grid = GridView::new(display_rows.len() as u32, col_ixs.len() as u32);
        {
            let mut sink = GridSink::new(&grid);
            fill_cells(
                &mut sink, &display_rows, &col_ixs, &widths, g,
                hr, mr, mc, 0, 4096, display_rows[0], 0, &agg,
            );
        }

        // Read the model in ONE borrow: holding a temporary `Ref` (as
        // `&m.borrow().cells` does) across a second `borrow()` deadlocks the
        // RefCell, which is what this test is careful not to do.
        let m = grid.model();
        let m = m.borrow();
        assert!(
            !m.cells.is_empty(),
            "fill_cells must have written cells through the shared sink"
        );
        // The cursor was forwarded as well (fill_cells calls set_cursor).
        assert_eq!((0u32, 0u32), (m.cursor_row, m.cursor_col));
    }
}
