# Verus proofs for corro

Two files verify today. Everything else in this directory is scratch and is
not claimed.

Checked with `verus 0.2026.10.18.7d82994` (bundled z3 4.16.0):

```
$ /opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/arith.rs
verification results:: 12 verified, 0 errors

$ /opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/col_codec.rs
verification results:: 8 verified, 0 errors
```

## `arith.rs` — the integer lemmas

Z3 does not multiply a hypothesis by a constant on its own, and it will not
chain a product identity into a bound. Both are trivial for a human and
neither comes free here, so the shapes that *do* go through are collected in
one place: `mul_mono_const` (by induction on the gap), `mul_nonneg`,
`expand26`, `expand10`, `comm26`, `le_of_le_eq`, `le_add_const`.

Every kernel proof in corro will need these; they are the reason a "one-line"
inequality takes a lemma.

## `col_codec.rs` — the column codec, partially

Established:

- `val_at` / `val` define the bijective base-26 value of a column name;
- `val_at` is non-negative on a valid name (`val_at_nonneg`);
- `val_at` unfolds exactly one digit (`val_at_unfold`);
- a valid name's value is `< 26^len` — *not* proved; the file stops short.

**Not established, and this is the honest limit of the file:** the upper bound
and the executable decoder loop. A valid name of `k` letters has value below
`26^k`; the tight bound needs a case split (when the last digit is `Z` and the
prefix is maximal the value reaches exactly `pow26(k)`), and the file carries
neither. The executable `parse_excel_column` is therefore **not** verified.

## What closing the codec would take

The missing piece is one lemma, `val_at_upper`, in this shape:

```rust
pub proof fn val_at_upper(s: Seq<u8>, i: int)
    requires valid(s), 0 <= i <= s.len(),
    ensures val_at(s, i) < pow26((s.len() - i) as nat),
```

The induction step is `d + 26*t` with `d <= 26` and `t < pow26(k)`. It splits
into two cases that *both* verify standalone (`scratch/try_CE1.rs`, 4
obligations):

- `d <= 25` — `d + 26t <= 25 + 26*(P-1) = 26P - 1`;
- `t <= P - 2` — `d + 26t <= 26 + 26*(P-2) = 26P - 26`.

The remaining case, `d == 26 && t == pow26(k) - 1`, gives exactly
`26*pow26(k)`. So the bound that actually holds is `val <= pow26(len)` (not
`pow26(len) - 1`) — `'Z'` is 26 and `26^1 - 1 = 25`, which is the counterexample
that cost the most time here. With that statement the induction closes
*except* for threading the `+ 26` slack through the recursive call: several
attempts in `scratch/` got the step lemma alone (`try_CA1.rs` and `try_CD1.rs`
both verify with 5 obligations)
but not the full recursion. That is the next hour of work, not a research
problem — but it is unfinished, and `col_probe.rs` still reports
`3 verified, 3 errors`.

## Why this matters for corro

`cargo test` cannot fail when a spec stops holding, so verification has to be
its own CI step naming the real source path. A file that nothing `mod`-declares
is invisible to `cargo` and visible to `verus`, so the step must name the file
explicitly.

## Files

| file | status |
|---|---|
| `arith.rs` | verifies: 12 obligations |
| `col_codec.rs` | verifies: 8 obligations (codec spec + non-negativity) |
| `col_probe.rs` | not finished; kept as the record of the failing obligations |
| `tiny.rs` | timing floor for a trivial file (~3 s wall) |
| `scratch/` | exploratory, not claimed |
| `VERUS_FIT_REPORT.md` | the earlier measurement report, unchanged |
