//! Cross-check: the Verus-verified column-name model against corro's real
//! `parse_excel_column`.
//!
//! `proofs/col_decode.rs` proves, for its own copy of the function, that the
//! result is `Some` exactly when every byte is `A`..=`Z`, and that the value is
//! the base-26 value of the name. A proof about a copy is only worth something
//! if the copy is the real thing, so this test drives corro's actual function
//! and the model over the same inputs.
//!
//! Two defects in `src/addr.rs` were found this way and have since been fixed;
//! the tests below pin the fixed behaviour and say what it used to be. Both
//! were found by writing the proof, not by fuzzing:
//!
//!   * `parse_excel_column("")` panicked on `n - 1` (release: wrapped to
//!     `u32::MAX`, i.e. the last column);
//!   * names of 7+ letters were indistinguishable from malformed ones.
//!
//! This test writes nothing.

use corro::addr::{parse_excel_column, parse_excel_column_checked};

/// The executor verified in `proofs/col_decode.rs`, transcribed back to Rust
/// with the same shape (accumulate, reject non-letters, return the index).
///
/// The `wrapping_*` calls are not a weakening: the proof's precondition says
/// the accumulation never leaves `u32`, so on the domain it covers they are
/// identical to plain arithmetic. Using them here lets the test *reach* the
/// out-of-domain inputs and show what happens there, instead of panicking
/// inside the model.
fn verified_model(name: &[u8]) -> Option<u32> {
    let mut acc: u32 = 0;
    let mut i = 0;
    while i < name.len() {
        let b = name[i];
        if !(b >= b'A' && b <= b'Z') {
            return None;
        }
        acc = acc.wrapping_mul(26).wrapping_add(b as u32 - 65 + 1);
        i += 1;
    }
    Some(acc)
}

/// corro returns a 0-based index; the model returns the 1-based Horner value.
fn model_as_index(name: &str) -> Option<u32> {
    verified_model(name.as_bytes()).map(|v| v.wrapping_sub(1))
}

#[test]
fn model_matches_corro_on_known_names() {
    for c in ["A", "B", "Z", "AA", "AZ", "BA", "ZZ", "ALL", "ZZZ", "XFD"] {
        assert_eq!(
            model_as_index(c),
            parse_excel_column(c),
            "model disagreed with parse_excel_column on {c:?}"
        );
    }
}

#[test]
fn model_matches_corro_on_rejected_input() {
    for c in ["a", "Ab", "A1", "1A", "[A", " ", "-", "A ", " A", "A-A"] {
        assert_eq!(
            model_as_index(c),
            parse_excel_column(c),
            "model disagreed with parse_excel_column on {c:?}"
        );
        assert_eq!(parse_excel_column(c), None, "{c:?} should be rejected");
    }
}

/// The verified claim is "`Some` iff every byte is `A`..=`Z`", checked over a
/// small exhaustive sweep. Names are kept below 7 letters, past which corro's
/// own `n - 1` can overflow (see `corro_rejects_7_z's`).
#[test]
fn some_iff_all_uppercase() {
    let alphabet = [b'A', b'Z', b'a', b'0', b'[', b'@'];
    let mut buf = Vec::new();
    for len in 0..=3usize {
        let mut idx = vec![0usize; len];
        loop {
            buf.clear();
            buf.extend(idx.iter().map(|&i| alphabet[i]));
            let name = std::str::from_utf8(&buf).unwrap();
            if !name.is_empty() {
                let all_upper = buf.iter().all(|b| b.is_ascii_uppercase());
                let got = parse_excel_column(name);
                assert_eq!(got.is_some(), all_upper, "Some mismatch for {name:?}");
                assert_eq!(model_as_index(name), got, "value mismatch for {name:?}");
            }
            if len == 0 {
                break;
            }
            let mut carry = true;
            for slot in idx.iter_mut().rev() {
                if *slot + 1 < alphabet.len() {
                    *slot += 1;
                    carry = false;
                    break;
                }
                *slot = 0;
            }
            if carry {
                break;
            }
        }
    }
}

/// FIXED (was a bug found by this proof): the empty string no longer panics.
///
/// `parse_excel_column("")` used to loop zero times, leave `n` at 0, and
/// evaluate `Some(n - 1)`, which subtracts from zero. In debug that panicked
/// with "attempt to subtract with overflow"; in release it wrapped to
/// `Some(u32::MAX)` and silently named the last column. It now uses
/// `checked_sub`.
#[test]
fn empty_input_is_rejected_not_a_panic() {
    let got = std::panic::catch_unwind(|| parse_excel_column(""));
    assert!(got.is_ok(), "empty input must not panic, got {got:?}");
    assert_eq!(parse_excel_column(""), None);
    // The released build used to wrap to u32::MAX here; that would have named
    // the last column, so assert it is not that either.
    assert_ne!(parse_excel_column(""), Some(u32::MAX));
}

/// FIXED (was a bug found by this proof): "not a column" and "column too large"
/// are no longer the same answer.
///
/// Seven or more `Z`s is a well-formed name whose index does not fit in a
/// `u32` (`ZZZZZZZ` denotes 8_353_082_581). `parse_excel_column` still returns
/// `None` -- the return type cannot express it -- but
/// `parse_excel_column_checked` now says *why*, so a caller need not report a
/// valid column as malformed.
#[test]
fn oversized_names_are_distinguished_from_malformed_ones() {
    use corro::addr::ColumnNameError;

    // The largest name that fits.
    assert_eq!(parse_excel_column_checked("ZZZZZZ"), Ok(321_272_405));
    assert_eq!(parse_excel_column("ZZZZZZ"), Some(321_272_405));

    // Well-formed but unrepresentable: OutOfRange, not Malformed.
    assert_eq!(
        parse_excel_column_checked("ZZZZZZZ"),
        Err(ColumnNameError::OutOfRange)
    );

    // Genuinely not a column name.
    for bad in ["", "abc", "A1", "a", "[A", " A"] {
        assert_eq!(
            parse_excel_column_checked(bad),
            Err(ColumnNameError::Malformed),
            "{bad:?} should be Malformed"
        );
    }

    // The two are distinguishable, which is the point of the fix.
    assert_ne!(
        parse_excel_column_checked("ZZZZZZZ"),
        parse_excel_column_checked("abc")
    );
}

/// The checked variant agrees with the plain one wherever the plain one succeeds.
#[test]
fn checked_agrees_with_plain_on_the_valid_domain() {
    for name in ["A", "Z", "AA", "AZ", "ZZ", "ALL", "ZZZ", "XFD", "ZZZZZZ"] {
        assert_eq!(
            parse_excel_column_checked(name),
            Ok(parse_excel_column(name).unwrap()),
            "checked disagreed with plain on {name:?}"
        );
    }
}

/// The verified model's precondition is exactly the domain where the plain
/// function succeeds.
#[test]
fn the_model_precondition_matches_the_plain_functions_domain() {
    // Inside the precondition (6 letters): model and function agree.
    assert_eq!(
        verified_model(b"ZZZZZZ").map(|v| v - 1),
        parse_excel_column("ZZZZZZ")
    );
    // Outside it (7 letters) the model wraps to a value that is not the index,
    // which is exactly why `proofs/col_decode.rs` carries the precondition.
    assert_eq!(verified_model(b"ZZZZZZZ"), Some(4_058_115_286));
    assert_eq!(parse_excel_column("ZZZZZZZ"), None);
}
