//! Cross-check: `excel_column_name` against `parse_excel_column`.
//!
//! `excel_column_name` now has a byte-level core
//! (`src/addr.rs::excel_column_name_bytes`) wrapped by a `String`-returning
//! function, so the allocation count dropped without changing any caller. These
//! tests pin the two things that refactor could have broken: the *output* and
//! the round-trip.
//!
//! The proof of the byte-level core is `proofs/col_encode.rs`; this test is the
//! half that checks it against the real function.
//!
//! This test writes nothing.

use corro::addr::{excel_column_name, parse_excel_column};

/// The names must be exactly what they were before the refactor. These are the
/// boundaries of the bijective base-26 scheme: single letters, the rollover to
/// two letters, the rollover to three, and the largest index a `u32` holds.
#[test]
fn known_names_are_unchanged() {
    for (idx, want) in [
        (0usize, "A"),
        (1, "B"),
        (25, "Z"),
        (26, "AA"),
        (27, "AB"),
        (51, "AZ"),
        (52, "BA"),
        (701, "ZZ"),
        (702, "AAA"),
        (16_383, "XFD"),
        (321_272_405, "ZZZZZZ"),
    ] {
        assert_eq!(excel_column_name(idx), want, "index {idx}");
    }
}

/// Every index in a broad range round-trips back to itself.
#[test]
fn roundtrip_over_a_broad_range() {
    for i in 0..70_000usize {
        let name = excel_column_name(i);
        assert_eq!(
            parse_excel_column(&name),
            Some(i as u32),
            "{i} -> {name:?} did not parse back"
        );
    }
}

/// Names are unique -- the scheme is a bijection, so two indices never collide.
///
/// (They are *not* in lexicographic order: "AA" sorts before "Z" in ASCII even
/// though index 26 > 25, because the scheme is bijective base 26, not plain.
/// Asserting lexicographic order here would be testing a property the encoding
/// does not have.)
#[test]
fn names_are_unique() {
    let mut seen = std::collections::HashSet::new();
    for i in 0..70_000usize {
        let name = excel_column_name(i);
        assert!(seen.insert(name.clone()), "{name:?} produced twice (at {i})");
    }
}

/// Length is the bijective base-26 digit count.
///
/// An index `i` needs `k` letters exactly when `S_k <= i < S_{k+1}`, where
/// `S_k = 26^0 + ... + 26^(k-1)` is the offset of the first k-letter name.
#[test]
fn lengths_follow_the_digit_count() {
    for i in 0..70_000usize {
        let n = excel_column_name(i).len();
        // smallest k with S_{k+1} > i, counting from k = 1
        let (mut lo, mut hi) = (0usize, 26usize); // S_1 = 0, S_2 = 26
        let mut k = 1usize;
        while !(lo <= i && i < hi) {
            lo = hi;
            hi = hi.saturating_mul(26) + 26; // S_{k+2} = 26*S_{k+1} + 26
            k += 1;
        }
        assert_eq!(n, k, "index {i} should be {k} letters, got {n}");
    }
}
