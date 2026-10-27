# proofs/ — Verus spikes

Nothing here is part of the cargo build. Each file is checked explicitly:

```
/opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/col_probe.rs
/opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent proofs/tiny.rs
/opt/verus-0.2026.10.18.7d82994/verus --triggers-mode silent balance-core/src/lib_verus_spike.rs
```

- `col_probe.rs` — `src/addr.rs parse_excel_column` as a `vstd`-shaped kernel.
  Not finished: 3 verified, 3 errors.
- `tiny.rs` — timing floor for a trivial file.
- `../balance-core/src/lib_verus_spike.rs` — the in-place arrangement, showing
  a spec + shim accepted next to the real crate.

Read `VERUS_FIT_REPORT.md` first; it states what each run printed and what
that means, and it does not claim anything the runs do not show.
