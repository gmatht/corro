//! Cross-check: the Verus-verified column-name model against corro's real
//! `parse_excel_column`.
//!
//! `proofs/col_decode.rs` proves, for its own copy of the function, that the
//! result is `Some` exactly when every byte is `A`..=`Z`, and that the value is
//! the base-26 value of the name. A proof about a copy is only worth something
//! if the copy is the real thing, so this test drives corro's actual function
//! and the model over the same inputs.
//!
//! Where they disagree, this file asserts the disagreement rather than hiding
//! it -- see `corro_panics_on_empty_input` and `corro_rejects_7_z's`. Both are
//! bugs in `src/addr.rs`, found by writing the proof, not by fuzzing.
//!
//! This test writes nothing.

use corro::addr::parse_excel_column;

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

/// BUG (found by this proof): the empty string panics.
///
/// `parse_excel_column("")` loops zero times, `n` stays 0, and `Some(n - 1)`
/// subtracts from zero. The verified model returns `Some(u32::MAX)`; corro
/// aborts in debug and wraps to `Some(u32::MAX)` in release.
#[test]
fn corro_panics_on_empty_input() {
    let got = std::panic::catch_unwind(|| parse_excel_column(""));
    assert!(
        got.is_err(),
        "expected the `n - 1` underflow; got {got:?}. If this now passes, \
         the empty-string bug was fixed and the model should be re-checked"
    );
}

/// BUG (found by this proof): 7 or more `Z`s are rejected although the name is
/// valid.
///
/// The Horner value of `ZZZZZZZ` is 8_353_082_582, and the index it denotes is
/// 8_353_082_581; neither fits in a `u32`, so `checked_mul`/`checked_add`
/// return `None`. That is arguably correct -- but the function returns `None`
/// for a well-formed name, which callers cannot distinguish from a malformed
/// one. `ZZZZZZ` (6 letters, index 321_272_405) is the largest that fits.
#[test]
fn corro_rejects_7_zs() {
    assert_eq!(parse_excel_column("ZZZZZZ"), Some(321_272_405));
    assert_eq!(parse_excel_column("ZZZZZZZ"), None);
    // The model wraps rather than returning None, so the precondition in
    // `proofs/col_decode.rs` (`pow26(len+1) <= u32::MAX`) is exactly the gap:
    // outside it the two functions are not claimed to agree, and here they
    // visibly do not.
    // The Horner value 8_353_082_582 wrapped mod 2^32 is 4_058_115_286 -- not the
    // true index, which is why the model is only correct inside the
    // precondition.
    assert_eq!(verified_model(b"ZZZZZZZ"), Some(4_058_115_286));
    assert_eq!(parse_excel_column("ZZZZZZZ"), None);
}
