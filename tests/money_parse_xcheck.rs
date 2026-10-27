//! Cross-check: the verified accept/reject shape of `parse_amount_cents`
//! against the real `balance_core::parse_amount_cents`.
//!
//! `proofs/money_parse.rs` proves the structure of the parse over bytes. The
//! real function differs in one deliberate way: it strips every ',' before
//! parsing, which is why `",1"` parses as 1 rather than being rejected. This
//! test pins that difference down instead of pretending it away.
//!
//! This test writes nothing.

use corro::balance::parse_amount_cents;

#[test]
fn accepts_plain_integers() {
    assert_eq!(parse_amount_cents("1"), Some(100));
    assert_eq!(parse_amount_cents("0"), Some(0));
    assert_eq!(parse_amount_cents("1234"), Some(123_400));
}

#[test]
fn accepts_one_or_two_fraction_digits() {
    assert_eq!(parse_amount_cents("1.5"), Some(150));
    assert_eq!(parse_amount_cents("1.25"), Some(125));
    // more than two digits are truncated, not rounded
    assert_eq!(parse_amount_cents("1.255"), Some(125));
}

#[test]
fn accepts_signs() {
    assert_eq!(parse_amount_cents("-1.5"), Some(-150));
    assert_eq!(parse_amount_cents("+1.5"), Some(150));
}

#[test]
fn rejects_malformed() {
    for c in ["", " ", "abc", ".5", "1.2.3", "-", "+", "1.5x"] {
        assert_eq!(parse_amount_cents(c), None, "{c:?} should be rejected");
    }
}

/// The verified model has no comma pre-pass; the real function strips every
/// comma first. That is why the real one accepts `",1"` -- the comma goes away
/// and what is left is `"1"`.
#[test]
fn commas_are_stripped_before_parsing() {
    assert_eq!(parse_amount_cents("1,234.56"), Some(123_456));
    assert_eq!(parse_amount_cents("1,"), Some(100));
    assert_eq!(parse_amount_cents(",1"), Some(100));
    // without the pre-pass these last two would have an empty whole part and
    // be rejected, which is what proofs/money_parse.rs proves about its own
    // input language.
}

/// A trailing '.' parses as a zero fraction.
#[test]
fn trailing_dot_is_an_empty_fraction() {
    assert_eq!(parse_amount_cents("1."), Some(100));
}
