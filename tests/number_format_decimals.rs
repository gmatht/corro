//! Decimal / Fixed / Currency formats on formula results.
//!
//! `NumberFormat::DecimalGeneric` was a no-op (it returned the text
//! unchanged), and Fixed/Currency couldn't handle a formula's exact-rational
//! display (`=2/6` shows as `1/3`, which does not parse as `f64`), so
//! formatting `=2/6` never produced a decimal. All backends share
//! `ui_core::format_cell_display`, so one fix covers TUI, GUI and pancurses.

use corro::grid::{CellAddr, CellFormat, GridBox, NumberFormat};

fn displayed(text: &str, fmt: Option<NumberFormat>) -> String {
    let addr = CellAddr::Main { row: 0, col: 0 };
    let mut g = GridBox::from(corro::grid::Grid::new(1, 1));
    g.set(&addr, text.to_string());
    corro::formula::refresh_spills(&mut g);
    let raw = corro::formula::cell_effective_display(&g, &addr);
    g.set_cell_format(
        addr.clone(),
        CellFormat {
            number: fmt,
            align: None,
        },
    );
    corro::ui_core::format_cell_display(&g, &addr, raw)
}

/// Format ▸ Number ▸ Decimal must turn `=2/6` into a decimal, not leave `1/3`.
#[test]
fn decimal_generic_converts_exact_rational_to_decimal() {
    assert_eq!(
        displayed("=2/6", Some(NumberFormat::DecimalGeneric)),
        "0.3333333333",
        "Decimal on =2/6 must print a decimal"
    );
    assert_eq!(
        displayed("=1/3", Some(NumberFormat::DecimalGeneric)),
        "0.3333333333"
    );
    assert_eq!(
        displayed("=-1/3", Some(NumberFormat::DecimalGeneric)),
        "-0.3333333333",
        "negative rationals keep their sign"
    );
}

/// The unformatted and Rational views keep the exact fraction.
#[test]
fn default_and_rational_views_keep_the_fraction() {
    assert_eq!(displayed("=2/6", None), "1/3");
    assert_eq!(
        displayed("=2/6", Some(NumberFormat::Rational)),
        "1/3",
        "Rational format keeps the exact form"
    );
}

/// Fixed and Currency share the same evaluator-literal parsing, so they now
/// work on formula results too.
#[test]
fn fixed_and_currency_apply_to_formula_results() {
    assert_eq!(
        displayed("=2/6", Some(NumberFormat::Fixed { decimals: 2 })),
        "0.33"
    );
    assert_eq!(
        displayed("=2/6", Some(NumberFormat::Currency { decimals: 2 })),
        "0.33"
    );
    assert_eq!(
        displayed("0.5", Some(NumberFormat::Fixed { decimals: 2 })),
        "0.50",
        "plain literals still format as before"
    );
}

/// Terminating decimals and non-numbers are untouched.
#[test]
fn terminating_and_text_cells_are_untouched() {
    assert_eq!(displayed("=2/4", Some(NumberFormat::DecimalGeneric)), "0.5");
    assert_eq!(displayed("hello", Some(NumberFormat::DecimalGeneric)), "hello");
    assert_eq!(displayed("hello", Some(NumberFormat::Fixed { decimals: 2 })), "hello");
}
