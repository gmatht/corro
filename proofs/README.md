# Verus proofs for corro

Eight files verify, plus cross-check tests that drive the real corro code.
Everything else in this directory is scratch and is not claimed.

Checked with `verus 0.2026.10.18.7d82994` (bundled z3 4.16.0):

```
$ verus --triggers-mode silent proofs/arith.rs
verification results:: 12 verified, 0 errors

$ verus --triggers-mode silent proofs/col_codec.rs
verification results:: 8 verified, 0 errors

$ verus --triggers-mode silent proofs/col_decode.rs
verification results:: 19 verified, 0 errors

$ verus --triggers-mode silent proofs/split_digits.rs
verification results:: 4 verified, 0 errors

$ verus --triggers-mode silent proofs/money_parse.rs
verification results:: 4 verified, 0 errors

$ verus --triggers-mode silent proofs/balance_subset.rs
verification results:: 4 verified, 0 errors

$ verus --triggers-mode silent proofs/shift_guards.rs
verification results:: 6 verified, 0 errors

$ verus --triggers-mode silent proofs/keyword_table.rs
verification results:: 10 verified, 0 errors

$ cargo test --test col_decode_xcheck --test split_digits_xcheck \
             --test money_parse_xcheck --test balance_subset_xcheck
running 5 tests ... test result: ok. 5 passed; 0 failed
running 5 tests ... test result: ok. 5 passed; 0 failed
running 6 tests ... test result: ok. 6 passed; 0 failed
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

## `split_digits.rs` — the fill-inference splitter

`src/extrapolate/mod.rs:398 split_trailing_digits`, over bytes: the scan
returns exactly `len - trailing_digit_run`, and `None` exactly when there is
no trailing digit. Supporting lemma `ndigits_split` carries the loop
invariant.

Not proved: that the byte at the boundary is a non-digit (the run is maximal).
The loop condition makes it true; stating it needs another induction.

This is the half of fill inference that is pure byte work. Its caller
`infer_suffix_fill` adds `parse::<i64>` and `format!`, neither of which has a
vstd spec -- the same wall the money parser hits.

## `money_parse.rs` — the money parser's shape, not its value

`balance-core/src/lib.rs:7 parse_amount_cents`: an optional sign, a non-empty
digit run, then optionally a `.` and a digit run reaching the end. Proved:
whenever it accepts, the parts it identified really are digit runs with a
non-empty whole part.

**Not proved:** the returned cents value. That needs `parse::<i64>`,
`saturating_mul`/`saturating_add` and the `replace(',', "")` pre-pass — three
things with no vstd spec. Also no round-trip, since `format_amount_cents` uses
`format!`.

This is the honest shape of the money work: the control flow is provable and
the arithmetic is reachable only through shims that would each be a proof
obligation in disguise.

## `balance_subset.rs` — book balancing cannot lie about "no match"

`balance-core/src/lib.rs:125 choose_subset` searches for a subset of the
remaining rows whose amounts sum exactly to an anchor's amount. The property
that matters is soundness: when it reports nothing, there really is nothing.
A search that silently gave up would balance a book wrongly and say nothing.

Proved: `prepend_ok` (adding a smaller index in front keeps a subset valid) and
`no_subset_step` -- if nothing in `[start, n)` sums to `target`, then for any
usable index `i` nothing in `(i, n)` sums to `target - vals[i]`. That is
exactly the step the recursion takes before moving to the next candidate.

**Not proved:** the executable recursion, the `used[]` mask, the outer loop over
anchors, or that groups are disjoint. Those need `Vec`, `sort_by_key` and
iterator chains rewritten to loops.

## `shift_guards.rs` — inserted-row repair cannot corrupt `$`-locked refs

`src/formula/mod.rs:626 shift_if_unlocked_row` and `:638 shift_if_unlocked_col`
decide whether a cell reference moves when rows or columns are inserted. Getting
this wrong silently rewrites a user's formulas, so the properties are:

- a locked coordinate is returned unchanged;
- one above the insertion gap is unchanged (rows after the gap keep their number);
- one at or below the gap moves by exactly `delta`;
- `None` exactly when the result leaves the `u32` range;
- the `i64` intermediate cannot overflow, so that range test is the only place a
  value is ever refused.

**Not proved:** `translate_cell_addr_by_offset` (:1145), which dispatches over
`CellAddr` variants and composes these shifts with `ColumnAddr::from_global`.

One thing worth knowing when reading the tests: `gap_start` is a **0-based**
main-row index while formula text is 1-based, so `A5` is index 4 and shifts only
when `gap_start <= 4`. Getting that backwards makes a correct function look broken.

## `keyword_table.rs` — the margin-keyword precedence rules

`src/ops/mod.rs:35 margin_key_agg_func` decides whether a margin label is an
aggregate. The interesting content is the *precedence* between four input
forms, and that is what is proved:

- a leading apostrophe makes the label literal text — never an aggregate,
  whatever follows;
- a single `=` prefix accepts **only** `TOTAL`, which is what keeps `=MIN` and
  `=SUM` usable as ordinary formulas;
- `==WORD` and a bare `WORD` share one table;
- nothing outside the eleven literals is ever accepted.

**Not proved:** the trim and upper-case steps (`str::trim`,
`to_ascii_uppercase`), which the source applies before the table. The
cross-check test covers the composed behaviour.

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
its own CI step naming the real source path. The two `*_xcheck` tests are the
other half: they are what notice when the *copy* in `proofs/` drifts from the
function in `src/`. A file that nothing `mod`-declares
is invisible to `cargo` and visible to `verus`, so the step must name the file
explicitly.

## Files

| file | status |
|---|---|
| `arith.rs` | verifies: 12 obligations |
| `col_decode.rs` | verifies: 19 obligations; the full decoder |
| `col_codec.rs` | verifies: 8 obligations (smaller spec of the same value) |
| `split_digits.rs` | verifies: 4 obligations; the digit splitter |
| `money_parse.rs` | verifies: 4 obligations; parse shape only |
| `balance_subset.rs` | verifies: 4 obligations; subset-sum soundness |
| `shift_guards.rs` | verifies: 6 obligations; insert-repair shift guards |
| `keyword_table.rs` | verifies: 10 obligations; keyword precedence table |
| `col_probe.rs` | superseded by `col_decode.rs`; kept as the failure record |
| `tiny.rs` | timing floor for a trivial file (~3 s wall) |
| `scratch/` | exploratory, not claimed |
| `VERUS_FIT_REPORT.md` | the earlier measurement report, unchanged |
