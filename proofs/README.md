# Verus proofs for corro

Thirteen files verify (93 obligations), plus eleven cross-check suites (56 tests)
that drive the real corro code. Everything else in this directory is scratch and
is not claimed.

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

$ verus --triggers-mode silent proofs/agg_sparse_order.rs
verification results:: 3 verified, 0 errors

$ verus --triggers-mode silent proofs/cursor_roundtrip.rs
verification results:: 4 verified, 0 errors

$ verus --triggers-mode silent proofs/cellref_parse.rs
verification results:: 7 verified, 0 errors

$ verus --triggers-mode silent proofs/decimal_parse.rs
verification results:: 2 verified, 0 errors

$ verus --triggers-mode silent proofs/col_encode.rs
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

## `agg_sparse_order.rs` — why the sparse aggregate path agrees with dense

`src/agg/mod.rs:60 fold_numbers_summable` has two paths. The dense one visits
every cell of the range in row-major order; the sparse one walks
`plan.stored_sorted`, the stored cells filtered to the range and sorted. They
agree only if **filtering preserves the visit order**.

Proved: `subseq_mem` (every element of a subsequence occurs in the source) and
`subseq_sorted` (a subsequence of a row-major sorted list is row-major sorted).
Together they say the sparse path sees exactly the sequence the dense path
would — which matters because aggregates are order-sensitive: floating-point
addition is not associative.

**Not proved:** that skipped empty cells contribute nothing, that
`summable_numeric` agrees between paths, or any `Number` arithmetic. This file
is about the *index order*, the part reachable without `Grid`,
`Box<dyn GridImpl>` or closures.

## `cursor_roundtrip.rs` — the sheet cursor and the cell address are inverses

`src/addr.rs:115 sheet_cursor_to_addr` and `:159 addr_to_sheet_cursor`. The UI
converts between them on every cursor move and every address display, so a
mismatch puts the selection or the status line on the wrong cell.

Two layers, both proved:

- `ColumnAddr::from_global` / `to_global` (`src/grid/mod.rs:222,234`) invert each
  other in both directions, given the region bounds. The left margin is globals
  `[0, 702)`, the main grid `[702, 702+main_cols)`, the right margin everything
  above.
- `cursor_to_addr` / `addr_to_cursor` over the five regions (Header, Left, Main,
  Right, Footer) compose to the identity.

`regions_are_distinct` is in the file so the round-trip is not vacuously true:
all five regions are reachable and map to distinct address kinds.

Modelled over `int` rather than `usize`/`u32`. The `as u32` casts in the source
are therefore outside the proof — the cross-check test covers those, and a
deliberately injected off-by-one in `to_global` is caught by it.

## `cellref_parse.rs` — where a cell ref ends

`src/addr.rs:282 parse_cell_ref_at` returns `(addr, locks, n)`, and every caller
slices the formula source at `n` to find the next token. Off by one and `A5+A6`
mis-tokenises. So the property worth proving is not the address value (that is
`col_decode.rs`) but that **`n` is exactly the letters and digits consumed, and
the byte after it is not a digit**.

Proved: `letter_run`/`digit_run` count maximal runs and stop where the run ends;
a run never passes the end of the input; `a1_len` is in range and the byte at it
is not a digit.

**Not proved:** the header/footer (`~`/`_`) branch, the `$`-lock prefix rules, or
`cell_ref_text` (which is `format!`-bound). The cross-check test covers those.

## `decimal_parse.rs` — which literals count as numbers

`src/formula/number.rs:462 parse_number_literal`. Proved is the accept/reject
shape, the part with no `BigInt` in it: non-empty, at most one `.`, every byte a
digit or `.`.

A quirk worth knowing, because it looks like a bug and is not: **`"."` and
`"-."` are accepted and parse as zero.** Neither side of the `.` being empty is
rejected, and `parse_decimal_rational` turns two empty parts into 0.

**Not proved:** the numeric value (`BigInt`/`BigRational`), the `trim` and the
complex-literal pre-pass, and `rational_to_formula_literal` — so no round-trip.

While cross-checking this, deleting the multi-dot guard in
`parse_number_literal` turned out to change no observable behaviour: the
downstream `parse_decimal_rational` rejects a second `.` anyway. Recorded in
`Bugs.txt`, since the guard fires (7 of 9 multi-dot inputs reach it) but is
redundant rather than load-bearing.

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

## Bugs these proofs found

Two real defects in `src/addr.rs`, both now fixed. Recorded here because the
*route* is the reusable part: neither came from fuzzing, review, or a failing
user report.

**The method.** Prove the kernel, then cross-check the proof against the real
function. A proof forces the input domain to be written down explicitly
(`pow26(len+1) <= u32::MAX`); every input the proof *excludes* becomes a
question — "why is that excluded, and what does the code do there?" Asking that
of an empty name produced an immediate panic on the first test written against
the real function.

| defect | mechanism | how the proof exposed it |
|---|---|---|
| `parse_excel_column("")` panicked | `Some(n - 1)` with `n = 0`; debug panics, release wraps to `u32::MAX` (the last column) | the domain had to exclude the empty name |
| 7+ letter names indistinguishable from malformed | the index of `ZZZZZZZ` is 8_353_082_581, which does not fit the `u32` return | the precondition `pow26(len+1) <= u32::MAX` had to be *justified* |

Both were latent: every caller guards `col_len == 0` first, and the `u32`
ceiling is real (six letters is the largest that fits). That is exactly why
neither showed up in testing — the code never reached them.

Each fix is pinned by a test that fails when the fix is reverted, checked by
reverting it. See `Bugs.txt` for the full mechanism of each.

**The caveat, stated plainly.** This found bugs in `addr.rs` because `addr.rs`
is the one kernel with a proof sharp enough to have a stated domain. The other
ten proofs do not imply their kernels are bug-free — only that the property
each one states holds.

## `col_encode` — the encoder's digit sequence, partly proved

`src/addr.rs::excel_column_name_bytes` is a *bijective* base-26 counter: each
step does `n -= 1` before taking `n % 26`. That is not a base conversion, and
both specs I first wrote were wrong — caught only by running the model against
data, not by the prover:

| claim | counterexample |
|---|---|
| `enc(m)` = base-26 digits of `m − 1` | `enc(27) = [0,0]`, base-26 of 26 is `[0,1]` |
| `sum(d_j · 26^j) = m − 1` | fails for 19,973 of the first 20,000 |

The characterisation that holds recurses exactly as the loop does:

```
digits(x) = []                                       if x == 0
            digits((x − 1)/26).push((x − 1) mod 26)  otherwise
```

validated against the real loop for every `x < 50000`.

**Proved** (`col_encode.rs`, 10 obligations): the recurrence, its length, that a
positive value has at least one digit, four concrete values (`A`, `Z`, `AA`,
`AB`), and that the loop terminates having written exactly `digits(x).len()`
bytes.

**Not proved:** that the bytes written *equal* `digits(x)`. The loop writes
MSB-first into a buffer tail, i.e. the reverse of `digits`, and that reversal
invariant is not closed — the loop is proved to have the right *length*, not the
right *content*. So the round-trip against `parse_excel_column` rests on
`tests/column_name_xcheck.rs` (70,000 indices), not on the proof.

## Why this matters for corro

`cargo test` cannot fail when a spec stops holding: `proofs/` is not a crate, no
file in it is reachable from a `mod` tree, and `cargo` therefore never compiles
them. So verification is `.github/workflows/verus.yml`, which names each proof
file explicitly. The `*_xcheck` suites are the other half: they are what notice
when a copy in `proofs/` drifts from the function in `src/`.

The job pins Verus by version rather than tracking `latest`, and sets
`RUSTUP_TOOLCHAIN=1.99.1-...` **only** on the step that runs `verus` -- the test
step must use the repo's own `rust-toolchain.toml` (`nightly`), or `src/` would
silently be built on a toolchain the repo does not declare.

Checked by running the job's own steps on a clean `RUSTUP_HOME` (only the pinned
toolchain installed, Verus fetched from the release URL): 11 proofs verify and
9 suites pass. Also checked that the job *fails*: breaking one `ensures` in
`cellref_parse.rs` makes the run exit 1 with that file named. A file that nothing `mod`-declares
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
| `agg_sparse_order.rs` | verifies: 3 obligations; sparse-path visit order |
| `cursor_roundtrip.rs` | verifies: 4 obligations; cursor <-> address inverse |
| `cellref_parse.rs` | verifies: 7 obligations; cell-ref consumed length |
| `decimal_parse.rs` | verifies: 2 obligations; decimal accept/reject shape |
| `col_encode.rs` | verifies: 10 obligations; encoder digit spec + loop length |
| `col_probe.rs` | superseded by `col_decode.rs`; kept as the failure record |
| `tiny.rs` | timing floor for a trivial file (~3 s wall) |
| `scratch/` | exploratory, not claimed |
| `VERUS_FIT_REPORT.md` | the earlier measurement report, unchanged |
## `col_encode` — attempted, not closed

`src/addr.rs::excel_column_name_bytes` is now a byte-level core with a
`String` wrapper, so it is a candidate for the inverse-direction proof. It is
**not proved**. Recording why, so the next attempt does not repeat the work.

The encoder is a *bijective* base-26 counter: each step does `n -= 1` before
taking `n % 26`. The obvious specs are wrong, and I checked each against data
rather than trusting it:

| claim | counterexample |
|---|---|
| `enc(m)` = base-26 digits of `m - 1` | `enc(27) = [0,0]`, but base-26 of 26 is `[0,1]` |
| `sum(d_j * 26^j) = m - 1` | fails for 19,973 of the first 20,000 |

The correct characterisation, validated against every `m < 20000`, is

```
S_k       = (26^k - 1) / 25            (offset of the first k-letter name)
n_k       = (m - S_k) / 26^k
digit_k   = (n_k - 1) % 26
```

`proofs/scratch/enc3.rs` proves `25 * S_k == 26^k - 1`. What is missing is one
integer-division step, `(n_k - 1)/26 == (m - S_k - 26^k)/26^(k+1)`, which needs
a Euclidean-division lemma that `scratch/enc5.rs` and `enc6.rs` did not
close. That is the open item; it is arithmetic, not a modelling problem.

What *is* pinned meanwhile: `tests/column_name_xcheck.rs` (4 tests) checks the
byte core against `parse_excel_column` over 70,000 indices, that names are
unique, and that lengths follow the bijective digit count. Mutation-checked:
`(n + 1) % 26` in the digit fails 2 of the 4.

## Why this matters for corro

`cargo test` cannot fail when a spec stops holding: `proofs/` is not a crate, no
file in it is reachable from a `mod` tree, and `cargo` therefore never compiles
them. So verification is `.github/workflows/verus.yml`, which names each proof
file explicitly. The `*_xcheck` suites are the other half: they are what notice
when a copy in `proofs/` drifts from the function in `src/`.

The job pins Verus by version rather than tracking `latest`, and sets
`RUSTUP_TOOLCHAIN=1.99.1-...` **only** on the step that runs `verus` -- the test
step must use the repo's own `rust-toolchain.toml` (`nightly`), or `src/` would
silently be built on a toolchain the repo does not declare.

Checked by running the job's own steps on a clean `RUSTUP_HOME` (only the pinned
toolchain installed, Verus fetched from the release URL): 11 proofs verify and
9 suites pass. Also checked that the job *fails*: breaking one `ensures` in
`cellref_parse.rs` makes the run exit 1 with that file named. A file that nothing `mod`-declares
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
| `agg_sparse_order.rs` | verifies: 3 obligations; sparse-path visit order |
| `cursor_roundtrip.rs` | verifies: 4 obligations; cursor <-> address inverse |
| `cellref_parse.rs` | verifies: 7 obligations; cell-ref consumed length |
| `decimal_parse.rs` | verifies: 2 obligations; decimal accept/reject shape |
| `col_encode.rs` | verifies: 10 obligations; encoder digit spec + loop length |
| `col_probe.rs` | superseded by `col_decode.rs`; kept as the failure record |
| `tiny.rs` | timing floor for a trivial file (~3 s wall) |
| `scratch/` | exploratory, not claimed |
| `VERUS_FIT_REPORT.md` | the earlier measurement report, unchanged |
