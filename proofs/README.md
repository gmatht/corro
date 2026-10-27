# Verus proofs for corro

Three files verify, plus a cross-check test that drives the real corro code.
Everything else in this directory is scratch and is not claimed.

Checked with `verus 0.2026.10.18.7d82994` (bundled z3 4.16.0):

```
$ verus --triggers-mode silent proofs/arith.rs
verification results:: 12 verified, 0 errors

$ verus --triggers-mode silent proofs/col_codec.rs
verification results:: 8 verified, 0 errors

$ verus --triggers-mode silent proofs/col_decode.rs
verification results:: 19 verified, 0 errors

$ cargo test --test col_decode_xcheck
running 5 tests ... test result: ok. 5 passed; 0 failed
```

## `arith.rs` — the integer lemmas

Z3 does not multiply a hypothesis by a constant on its own, and it will not
chain a product identity into a bound. Both are trivial for a human and
neither comes free here, so the shapes that *do* go through are collected in
one place: `mul_mono_const` (by induction on the gap), `mul_nonneg`,
`expand26`, `expand10`, `comm26`, `le_of_le_eq`, `le_add_const`.

Every kernel proof in corro will need these; they are why a "one-line"
inequality takes a lemma.

## `col_decode.rs` — the column codec, proved

`src/addr.rs:34 parse_excel_column`, transcribed to `&[u8]` and verified end to
end: `Some` iff every byte is `A`..=`Z`, and on success the value is the
base-26 value of the name.

The hard part was the bound. Both of the obvious statements are **false**:

| claim | counterexample |
|---|---|
| `val <= 26^len - 1` | `"Z"` = 26, and `26^1 - 1 = 25` |
| `val <= 26^len` | `"ZZ"` = 702, and `26^2 = 676` |

The exact maximum is `(26^(len+1) - 26)/25`, and the bound that *composes*
under the recursion is

```
25 * val(s) + 26  <=  26^(len+1)
```

because `25*(26t + d) + 26 = 650t + 25d + 26 <= 650t + 676 = 26*(25t+26)`.
With that, the induction is linear at every step and closes.

`col_codec.rs` (8 obligations) keeps the earlier, smaller spec of the same
value; it is superseded by `col_decode.rs` for the loop, and is retained
because it states the value function in the digit-at-front form that reads
more like the source.

## `tests/col_decode_xcheck.rs` — the proof checked against corro

A proof about a copy is only worth something if the copy is the real thing.
This test drives corro's actual `parse_excel_column` and the verified model
over the same inputs: known names, rejected input, and an exhaustive sweep of
`{A,Z,a,0,[,@}` up to length 3, asserting `Some` iff all-uppercase.

It found two bugs in `src/addr.rs`, both asserted rather than papered over:

- **`parse_excel_column("")` panics.** `n` stays 0 and `Some(n - 1)` subtracts
  from zero.
- **`"ZZZZZZZ"` and longer are rejected** although well formed: the Horner
  value overflows a `u32`. `"ZZZZZZ"` (index 321_272_405) is the largest that
  fits.

The proof's precondition `pow26(len+1) <= u32::MAX` is exactly the domain where
the two agree, so neither bug is inside what is proved — which is also why the
precondition is worth stating explicitly rather than hiding.

## Why this matters for corro

`cargo test` cannot fail when a spec stops holding, so verification has to be
its own CI step naming the real source path. A file that nothing `mod`-declares
is invisible to `cargo` and visible to `verus`, so the step must name the file
explicitly.

## Files

| file | status |
|---|---|
| `arith.rs` | verifies: 12 obligations |
| `col_decode.rs` | verifies: 19 obligations; the full decoder |
| `col_codec.rs` | verifies: 8 obligations (smaller spec of the same value) |
| `col_probe.rs` | superseded by `col_decode.rs`; kept as the failure record |
| `tiny.rs` | timing floor for a trivial file (~3 s wall) |
| `scratch/` | exploratory, not claimed |
| `VERUS_FIT_REPORT.md` | the earlier measurement report, unchanged |
