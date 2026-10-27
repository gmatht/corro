# Can Verus prove parts of corro? — measured, not guessed

Tooling used for every result below: `verus 0.2026.10.18.7d82994`, bundled
`z3 4.16.0`, toolchain `1.99.1-x86_64-unknown-linux-gnu`.
corro itself is on unpinned `nightly` (`rust-toolchain.toml`).

Every block in §1 is pasted from a real run in this worktree; nothing is
paraphrased. §2 is what those runs mean, §3 is the short answer.

---

## 1. What the runs actually printed

### 1.1 The whole corro library compiles on Verus's pinned toolchain

```
$ RUSTUP_TOOLCHAIN=1.99.1-x86_64-unknown-linux-gnu cargo check --lib --offline
    Finished `dev` profile [unoptimized + debuginfo] target(s) in 0.57s
```

(the first, cold run of the same command took `33.52s`; it re-checks `corro
v0.7.0` and every dependency, then is incremental.)

This is the single biggest strategic unknown, and the answer is favourable:
a proof can be aimed at a real source file inside the real workspace, rather
than at a kernel retyped into a separate crate. corro compiles on the
toolchain Verus is pinned to.

### 1.2 A real kernel, in place, in a workspace crate — `parse_amount_cents`

`balance-core/src/lib_verus_spike.rs` (a new file inside the existing
`balance-core` crate, not a copy outside the tree). The spec and the first
three statements of the real function:

```rust
pub assume_specification [str::trim] (s: &str) -> (t: &str);

pub fn parse_amount_cents(raw: &str) -> (r: Option<i64>)
    ensures
        r.is_some() ==> exists|w: Seq<char>, f: Seq<char>|
            #![trigger cents_of(w, f)]
            all_digits(w) && all_digits(f) && w.len() >= 1
            && r.unwrap() as int == cents_of(w, f),
{
    let t = raw.trim();
    if t.is_empty() {
        return None;
    }
    let mut s = t;
    // ---- Verus stops HERE. The rest of the real body is below, verbatim,
    // because str::replace has no vstd spec -- see proofs/VERUS_FIT_REPORT.md.
    None
}
```

The 15 statements Verus never reaches are kept verbatim below that point, so
the file is honest about how little of the function it covers.

Verus accepts this much:

```
$ RUSTUP_TOOLCHAIN=1.99.1-x86_64-unknown-linux-gnu cargo check -p balance-core --offline
    Finished `dev` profile [unoptimized + debuginfo] target(s) in 0.30s

$ /opt/verus-.../verus --triggers-mode silent balance-core/src/lib_verus_spike.rs
verification results:: 2 verified, 0 errors
```

Two facts here, and both matter:

- the Verus spec is valid Rust, so it lives **in** `balance-core/src/` with no
  shadow copy to drift out of sync with the real function;
- `str::trim` is not in vstd, and `assume_specification [str::trim]` is
  accepted (the shim costs one line).

### 1.3 Where that same function stops

Uncommenting the next line of the real body:

```
error: `alloc::str::impl&%5::replace` is not supported (note: you may be able to add a Verus specification to this function with `assume_specification`) (note: the vstd library provides some specification for the Rust std library, but it is currently limited)
  --> balance-core/src/lib_verus_spike.rs:44:13
```

The same run, with comments and everything after `replace` left in, reported
all eight blockers at once:

```
error: `alloc::str::impl&%5::replace` is not supported ...
error: `core::str::impl&%0::split_once` is not supported ...
error: `core::char::methods::impl&%0::is_ascii_digit` is not supported ...
error: `core::num::impl&%3::saturating_mul` is not supported ...
error: `core::num::error::ParseIntError` is not supported ...
error: `core::str::impl&%0::parse` is not supported ...
error: `core::str::impl&%0::get` is not supported ...
error: `core::num::impl&%3::saturating_add` is not supported ...
```

### 1.4 An integer-shaped kernel: `parse_excel_column`

`proofs/col_probe.rs` — corro's bijective base-26 decode
(`src/addr.rs:34`), with `val()`/`valid()`/`val_push()` written as
`vstd` specs. There is no `String`, no `parse`, no `format!` in the parse
direction, so this is the best case Verus can be given in corro:

```
$ /opt/verus-.../verus --triggers-mode silent proofs/col_probe.rs
error: postcondition not satisfied
error: postcondition not satisfied
error: invariant not satisfied at end of loop body
error: possible arithmetic underflow/overflow
verification results:: 3 verified, 3 errors
error: aborting due to 4 previous errors
```

Three functions verify; four obligations fail. `val_push` is the append
lemma `val(s.push(b)) == 26*val(s) + digit(b)` — its body is a single call to
`len_push`, and it still does not go through. Nothing here is unusual work for
a Verus user, but the *shape* is: every lemma needs its own induction, and
some elementary integer identities have to be handed over in the exact form
Z3 will accept.

### 1.5 Cheapness per proof obligation

```
$ /usr/bin/time -f "%e s wall" verus proofs/tiny.rs     # 2 trivial obligations
verification results:: 2 verified, 0 errors 3.56 s wall
verification results:: 2 verified, 0 errors 2.95 s wall
verification results:: 2 verified, 0 errors 3.32 s wall
```

(three consecutive runs). Three seconds of fixed overhead on a two-line file
sets the floor for any CI job that verifies anything at all.

---

## 2. What this means for corro

### 2.1 The boundary is `str`, not `unsafe`

corro's small pure kernels are not blocked by pointers, threads or I/O —
they are blocked by `str`/`String` methods. `trim`, `replace`, `split_once`,
`parse`, `format!`, `is_ascii_digit`, `saturating_mul` all lack vstd specs,
and they appear everywhere in exactly the functions a proof would want:

| file | `str`/`char`/num-op call sites |
|---|---|
| `src/extrapolate/mod.rs` | 53 |
| `src/addr.rs` | 32 |
| `balance-core/src/lib.rs` | 21 |
| `src/formula/number.rs` | 16 |
| `src/agg/helpers.rs` | 12 |
| `src/celladdr.rs` | 9 |

Two ways through, and they are not equal:

- **Shims** (`assume_specification`): one line each, but a shim that says
  only "returns some string" (as §1.2's `trim` does) buys nothing — the
  postcondition about the parsed value cannot be closed through it. A shim
  that says what the function *means* is a proof obligation in disguise.
- **Byte-level rewriting**: replace `&str` with `&[u8]` in the kernel. This
  is what makes `parse_excel_column` tractable (§1.4), and it is the only
  honest route for anything whose correctness is about characters.

The cost is therefore not "add Verus", it is "rewrite the kernels over
`&[u8]`". That is a real refactor of code that currently passes its tests.

### 2.2 The in-place arrangement is the good news — with one caveat

Because corro builds on Verus's toolchain (§1.1) and specs are valid Rust
(§1.2), a proof can sit next to the function it proves:

```
balance-core/src/lib.rs              <- unchanged body
balance-core/src/lib_verus_spike.rs  <- spec + proof, same crate
```

`cargo` compiles only what a `mod` tree reaches; `rust_verify` checks the
file it is handed. So verification has to be wired into CI *as its own step*,
pointing at the real source path — nothing in `cargo test` can fail when a
spec stops holding. That is the failure mode to guard from day one.

(One hazard seen in this very spike, worth knowing: the file was created
between two `cargo` runs and neither `cargo check -p balance-core` nor
`cargo check --lib` mentioned it. A file that nothing `mod`-declares is
invisible to `cargo` and visible to `verus`, which is another reason the
`verus` invocation must name the file explicitly rather than trusting a
build to pick it up.)

### 2.3 The honest ceiling

What is provable today, in descending order of return:

1. **Sparse/dense aggregate equivalence** (`src/agg/mod.rs`,
   `fold_numbers_summable` vs `count_numeric_cells` sparse fast path). This
   is the one correctness property with real user-visible consequence — a
   wrong fast path silently changes a displayed SUM — and the kernel part is
   an index-set lemma that can be written over `Seq<(u32,u32)>` with no
   `str` in sight.
2. **Address round-trips** (`addr_to_sheet_cursor ∘ sheet_cursor_to_addr =
   id`). Integer-valued, five regions, no strings in the interesting part.
3. **`parse_excel_column`** (§1.4) — small, self-contained, and finishes
   with ordinary proof work.

What is *not* provable without a rewrite: `Number` arithmetic
(`BigRational` + uninterpreted `f64`), the formula evaluator
(`thread_local!` caches, `Rc<Ast>`, `Drop`), anything with `format!`
(which is most display code), and the ODS/zip layer.

All three of those are *proposals*, not results. None was attempted in this
spike; the only kernel actually put to Verus is `parse_excel_column` (§1.4),
and it is unfinished.

### 2.4 The thing to budget for

§1.4 shows the shape: specs discharge quickly, then the arithmetic grinds.
Verus/Z3 refused several elementary integer facts outright, and the workaround
each time was to hand it the fact in the form it can use
(`rec_le(a, b-1, m)` by induction for `a*m <= b*m`; `nonlinear_arith` for
`26*(P-1) + 26 = 26*P`). Expect to spend most of the time on lemmas about
multiplication, not on corro.

---

## 3. Answer

**Yes, but not for the code as written, and not for much of it.**

- The toolchain fits: corro builds on it, and proofs live in-tree next to
  the real functions with no copied code. That half of the integration risk
  is gone.
- The blocker is `str`. corro's small kernels are written in idiomatic
  string code, and vstd specifies almost none of it, so a proof of
  `parse_amount_cents` as it stands would be built on eight trusted stubs.
- Start where strings are absent: the sparse/dense aggregate equivalence
  first, then the address round-trip, then `parse_excel_column`. If those
  three land, the case for touching the string kernels with `&[u8]` becomes
  worth making; if they do not, no amount of shimming will help.

Not attempted here and not claimed: `excel_column_name` (needs `String`
building), `choose_subset` (`Vec`/`sort_by_key`/iterators), and the log codec
(`format!`).
