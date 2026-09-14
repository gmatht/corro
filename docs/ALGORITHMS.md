# Corro algorithms & data structures audit

> Audit + implementation record. §§1–3 are the original analysis; §6 records
> what was implemented (all behaviour-preserving: full lib suite + targeted
> regression tests green) and what was deliberately deferred. Line references
> are to `src/grid/mod.rs`, `src/formula/mod.rs`, `src/ops/mod.rs`, `src/agg/`,
> `src/export.rs`, `src/ui_core.rs`, `src/ui/mod.rs`, `src/io/mod.rs`.

## 1. Storage model (what is appropriate today)

- The sheet is **sparse**: `Grid.main_cells: HashMap<(u32,u32), String>`,
  plus `left: HashMap<(u32, MarginIndex), String>`,
  `right: …`, `header/footer: HashMap<(u32, ColumnAddr), String>`
  (`src/grid/mod.rs`, `struct Grid`). Absent key = empty.
- That is the **right** choice for an unbounded logical sheet
  (`HEADER_ROWS = FOOTER_ROWS = 999_999_999`, `MARGIN_COLS = 702`,
  `total_cols() = 1404 + main_cols`): a dense `Vec<Vec<…>>` would be
  impossible, and point lookup is expected `O(1)`.
- Formats/widths are also sparse maps (`col_width_overrides`,
  `col_all/data/special_formats: HashMap<usize, CellFormat>`,
  `cell_formats: HashMap<CellAddr, CellFormat>`). Fine.
- `ColumnAddr::{Left,Main,Right}` is objective (stable across main resize)
  for header/footer keys — good; global-column `usize` keys used for widths
  and formats are *not* stable and must be remapped on every resize
  (`remap_main_col_layout_for_resize`, `remap_formats_for_resize`). The
  remaps themselves are `O(m)` in the size of the override maps — fine.

**What is missing is secondary indexes.** Every "does row/col r/c contain
anything?" or "what is the extent?" query is answered by a **full scan of a
primary map** (`O(n)`, `n` = number of stored cells), and callers then loop
that query over all rows/columns/extents. That is where every `O(n²)`-class
behaviour below comes from. The maps themselves are not the problem; the
absence of row/column occupancy bookkeeping is.

Related allocation smell: `GridImpl::iter_nonempty()` materialises a
`Vec<(CellAddr, String)>` — cloning **every stored string** — on each call,
and `get_owned()` clones per lookup. Hot paths call `iter_nonempty()` once
per row (see §2.4) or several times per frame (see §2.5), so CPU *and*
allocator cost scale with `n` far more often than necessary. A borrowing
iterator / `get(&self) -> Option<&str>` fast path for internal callers would
cut the constant factor everywhere below (no complexity change, but large).

## 2. Hotspots: current vs. attainable complexity

Notation: `n` = stored non-empty cells; `R` = `extent_main_rows`;
`C` = `extent_main_cols`; `T = 2·MARGIN_COLS + C` (≈1404 minimum);
`F` = formula cells; `V` = visible cells per frame.

### 2.1 `logical_row/col_has_content` — `O(n)` each, called in loops

Implementation (`src/grid/mod.rs`):

- `logical_row_has_content(r)` — `.keys().any(...)` over `header`,
  then `main_cells` + `left` + `right`, then `footer`. `O(n)` per call.
- `logical_col_has_content(c)` — `.keys().any(...)` over `header`
  (with a `to_global` per key), then one margin/main map, then `footer`.
  `O(n)` per call.

Callers that multiply it:

| Caller | Cost today | Should be |
|---|---|---|
| `Grid::col_width()` calls `logical_col_has_content` (`grid/mod.rs`). Render calls `col_width` per visible column per frame | `O(cols_visible · n)` per frame | `O(1)` with col occupancy counts |
| `main_col_window` (`ui_core.rs`) loops `0..mc` × `logical_col_has_content` | `O(C·n)` per viewport computation | `O(C)` / `O(occupied)` |
| `right_nonblank_end` (`ui_core.rs`) walks up to 702 margin cols × `logical_col_has_content` | `O(702·n)` | `O(1)` / scan of occupied set |
| `ui/mod.rs` viewport search `find(... logical_row_has_content ...)` + `logical_col_has_content` loop | `O(R·n + C·n)` | `O(1)` each |
| `ascii_col_bounds` (`export.rs`): `while` from both ends over up to `T` columns × `O(n)` | `O(T·n)` per export (T ≥ 1404 always) | `O(1)`–`O(log n)` bounds from index |

### 2.2 `content_width_for_column` + `auto_fit_column` on every write — quadratic import

- `content_width_for_column` (`grid/mod.rs`) scans header + footer maps
  fully and then loops `0..extent_main_rows` doing a HashMap `get` per row:
  `O(n_hdr + n_ftr + R)` per call.
- `Grid::set()` calls `auto_fit_column` → `content_width_for_column` on
  **every single non-empty write** (main, left, right, header, footer paths
  all do it).
- Consequence: importing `N` cells into a column with `R` rows costs
  `O(N·R)` — e.g. `import_delimited` (`io/mod.rs`, one `set` per field)
  re-scans the whole column per field. The same applies to paste/fill loops
  (`FillRange`, `RelFillRange`, `CopyFromTo` in `ops/mod.rs`) and to the
  per-cell `grid.set` inside duplicate/duplicate-range apply paths, which
  additionally call `resize_header_footer_width()` (two full `retain`
  scans) per cell.

What *should* happen: width = max over one column is a textbook incremental
aggregate — maintain per-column max (and handle shrink by lazy recompute /
refcount on the winning width), or at minimum defer auto-fit out of the
per-cell path (fit once per batch/import/column, or only when the written
value is wider than the cached width and recompute-from-scratch only when a
narrowing write removes the max).

### 2.3 `sorted_main_rows` recomputed from scratch per frame — `O(R log R)` with expensive comparator

`Grid::sorted_main_rows` (`grid/mod.rs`) builds `0..extent_main_rows` and
sorts with a comparator that does `grid.get(...)` (HashMap + `String` clone
via `get_owned` in some paths) per compared spec column, plus
`compare_sort_values` string/number parsing per comparison:
`O(R log R · S)` HashMap lookups, `S` = sort specs. It is called from
`visible_row_indices` (`ui_core.rs`), both `ui/mod.rs` render paths
(`main_order`), export row ordering, and status code — i.e. potentially
**several times per frame**, with no cache/invalidation (no generation
counter; `view_sort_cols` change vs. data change are not distinguished).
Recommendation: cache the order + a dirty bit bumped only by data/sort-spec
writes (or by a grid revision counter), not recomputed per viewport call.
Unsorted fast path (`view_sort_cols.is_empty()` → `0..R` vec) is already
cheap but still allocates per call; callers could share one.

### 2.4 Export / row-list construction is `O(rows · n)` — full scan per row

`delimited_table_col_span_and_rows` (`export.rs`): for **each** candidate row
it iterates `grid.iter_nonempty()` (which itself clones all `n` strings)
and breaks on first match — worst case `O(rows·n)` string-clone-heavy work,
run on top of `ascii_col_bounds` (`O(T·n)`, §2.1) and `row_order` /
`main_row_index_bounds_for_export` (each another full `iter_nonempty` scan).
`row_order` then fills *contiguous* main ranges (gap rows included), and
`odt_content_xml` emits `rows × T` cells with `T ≥ 1404` — i.e. ODT output
size/compute scales with the margin width even for a 1-column sheet.
With row/col occupancy sets (§3), row-list construction becomes
`O(occupied)` and bounds become `O(1)`/`O(log n)`.

### 2.5 Formula evaluation has no memoisation — repeated parse + re-eval per frame

- `cell_effective_display` (`formula/mod.rs`) re-**parses** the formula
  (`Parser::parse_expr`, fresh AST, fresh allocations) and recursively
  evaluates referenced cells on **every call**; render paths call it per
  visible cell per frame (`ui/mod.rs` has ~10 call sites in the render
  loop; `gui/render.rs` similarly), and `visible_row_indices`-adjacent code
  evaluates cells just to measure/normalise text. So one frame costs
  `O(V · eval_cost)` with `eval_cost` itself unbounded by sharing: a column
  of `=A1+…` chains re-walks shared precedents per cell (chain of length `L`
  → `O(L)` per cell → `O(L²)` per column per frame), and aggregates
  re-evaluate every cell in their range (§2.6).
- Cycle detection is a linear `visiting.iter().any(...)` scan
  (`formula/mod.rs`, several sites) — `O(depth)` per step, depth ≤ 128, so
  bounded and fine; the fix is caching, not the cycle check.
- `refresh_spills` (`formula/mod.rs`): up to **8 fixpoint passes**; each
  pass collects + sorts **all** nonempty cells (`O(n log n)` + clone of all
  strings), then `eval_cell` on every formula cell with a fresh
  `DEFAULT_BUDGET = 10_000`. It early-exits via a `spills_dirty` flag
  (good), but any single `set` marks everything stale, so interactive typing
  pays the full multi-pass scan per keystroke. A dependency-aware
  (only re-evaluate array-formula anchors / dependants) or at minimum
  generation-memoised design would cut this to the affected subset.
- `repair_all_formulas_after_main_row/col_insert` (`formula/mod.rs`,
  invoked from `MoveRow/ColRange`, duplicate/delete paths in `ops/mod.rs`):
  iterates **all** nonempty cells, re-parses every `=…` string and rewrites
  it — `O(n · parse_cost)`, plus each resulting `grid.set` re-marks spills
  stale and re-auto-fits (§2.2). Touching all formulas is asymptotically
  necessary for an insert (references below the gap move), but the constant
  factor (full re-parse + per-cell set overhead) dominates; batching the
  writes (single stale-mark, single fit pass) is the cheap win.

### 2.6 Aggregates re-scan + re-evaluate their range per aggregate cell

`collect_numbers_summable` / `count_numeric_cells` (`agg/mod.rs`) loop the
dense `rows × cols` area and call `summable_numeric`/`effective_numeric`
(= parse + evaluate) **per cell**. `compute_aggregate` calls the collector
once per aggregate invocation, and each footer/margin aggregate cell
(`footer_special_col_aggregate`, `left/right_margin_*_aggregate` in
`agg/helpers.rs`, `export.rs` aggregate-formula helpers) invokes it over its
own source range — so `A` aggregates over a shared column of `N` cells cost
`O(A·N·eval_cost)`. Mean collects then folds (fine); Min/Max collect a full
`Vec<Number>` just to take min/max (streaming fold would halve memory, minor).
`median_aggregate` sorts — `O(N log N)`, optimal for exact median; note only
that it also pays the collect cost first. Dense-range iteration (`sum_main_range`
in `formula/mod.rs` has the same shape) is `O(area)` even when the range is
sparse — acceptable for bounded ranges, but a large `SUM(A1:Z100000)`-style
range iterates millions of empty addresses each doing a HashMap lookup; an
occupied-row/col index could bound iteration to stored cells.

### 2.7 Structural moves are dense in a sparse grid

- `move_main_rows` (`grid/mod.rs`): builds the permutation, then for **every**
  `(new_pos, c)` with `c` in `0..extent_main_cols` does a HashMap `get`
  (main), and for every `(new_pos, mc)` with `mc` in `0..702` a `get` on
  left **and** right. Cost `O(R·C + R·1404)` HashMap lookups even when the
  sheet holds a handful of cells. Same shape mirrored in `move_main_cols`
  (`O(R·C)` gets + header/footer remap). Correct fix: iterate the **stored**
  cells once (`O(n)`) and remap keys through the permutation (`new_pos =
  position[old]`), touching only what exists.
- `DuplicateRowRange` / `DuplicateColRange` (`ops/mod.rs`): collect phase is
  proportional to the range (dense cell-by-cell `get`, including margin
  loops over all 702 margin cols per row for row-duplication) — mostly
  proportional to copied data, acceptable, except the 702-wide margin sweep
  per row should iterate stored margin cells of those rows instead.
- `Grid::set` with empty value → `shrink_to_content()`: scans **all** of
  `main_cells` + header + footer + left + right (`O(n)`) to recompute both
  extents — on **every single-cell clear**. Clearing `k` cells (e.g. a column
  delete implemented as clears, or user holding Delete) costs `O(k·n)`.
  `set_main_size`'s `retain` passes are `O(n)` each but happen once per
  explicit resize — fine.

### 2.8 Parse-per-keystroke and string churn (constant-factor, but pervasive)

`encode/decode_log_value`, `addr_text`/`cell_ref_text`, formula
`translate_*` (parse → AST → rewrite → render string) are all linear-time and
fine individually, but they sit inside the loops above (repair-all,
duplicate-paste with `translate_formula_text_by_offset` per cell, log replay
in `io/mod.rs` which applies ops one by one, each paying §§2.2/2.5 costs).
Log import therefore inherits the per-`set` quadratic terms: replaying a
`SIZE` + `N` SETs + interleaved auto-fits is `O(N·R + N·n_spillscan)`.
Batching import (suspend auto-fit + spill refresh until end of replay) is the
standard fix and does not change semantics.

## 3. Should we use reference counting per row/col for blank row/col?

**Yes — with one design caveat about which key you count by.** It directly
converts every `O(n)` occupancy probe in §2.1 (and the export/ODT/viewport
derivations) into `O(1)`, and additionally gives `O(1)` shrink decisions
(§2.7) and `O(occupied)` bounds enumeration instead of `O(extent)` scans.

Recommended shape:

- Keep the sparse primary maps as-is. Add per-region occupancy:
  - `main_row_count: HashMap<u32, usize>` / `main_col_count: HashMap<u32, usize>`
    (or a single `row_count`/`col_count` per region: main, left, right,
    header-by-row, footer-by-row),
  - plus the *logical* rollups the queries actually ask: logical-row
    occupancy = main-row-content ∪ margins ∪ header/footer-row-content, and
    logical-col occupancy keyed by the **objective** column identity, not the
    volatile global index.
- The caveat: global column indexes shift on main resize (right-margin
  globals move; `ColumnAddr::Main(i)` does not). So per-**global**-col
  refcounts would need a remap on every resize — the same remap pass widths
  already pay, so piggybacking is possible, but cleaner is to count by
  `(Region, index)` / `ColumnAddr` and resolve to global only at query time.
  Header/footer are already keyed by `ColumnAddr` — count them the same way.
- Update discipline: increment on transition empty→non-empty, decrement on
  the reverse, inside the single choke point (`Grid::set` and the batch
  paths `set_main_size`/`move_*`/duplicate/delete, which should apply
  deltas rather than rebuild). Empty-string stores already remove the key —
  hook the same branch. Cost per write stays `O(1)` amortised; memory is
  `O(occupied rows + occupied cols)`, negligible.
- Decide and document whether **spill followers** count as content. Today
  `get()` returns the spill value first, so spills already *look* like
  content to most readers while the occupancy scans (`keys().any` on primary
  maps) ignore them — an inconsistency worth pinning down when the counts
  are introduced (recommend: spills do not affect stored-content counts;
  viewport/export already handle spill overflow separately).
- Consider `BTreeSet<u32>` (per region/dimension) *alongside* counts rather
  than instead of them: counts give `O(1)` emptiness; ordered sets give
  min/max/next-occupied for bounds, viewport windows and export spans in
  `O(log n)` without scanning `0..T` / `0..R`. Either is a large win over
  today; both together are a few hundred bytes of bookkeeping.

What refcounts do **not** fix (do these separately): per-frame formula
re-evaluation (§2.5 — needs memoisation/dependency tracking), the
`content_width` max-tracking (§2.2 — needs a per-column width cache, which
pairs naturally with the col counts), dense structural moves (§2.7 — iterate
stored cells), and the `iter_nonempty`-clones-everything pattern (§1 —
borrowed access).

## 4. Prioritised recommendations

1. **Row/col occupancy index** — DONE. `Grid` carries per-region
   row/col counts (`main_row_counts`, `left_col_counts`,
   `header_col_counts`, …), maintained incrementally in `set` and
   rebuilt after bulk ops (`set_main_size`, `move_*`, `clear_cells`).
   `logical_row/col_has_content` are `O(log n)` probes. Header/footer
   columns are keyed by objective `ColumnAddr`; the col probe additionally
   checks the `Main(i)` spelling for out-of-extent absolute globals
   (`Main(i)` keeps global `MARGIN_COLS+i` while narrow) — the one edge
   the naive `from_global` inverse misses, caught by
   `occupancy_sees_out_of_extent_main_header_cols`.
2. **Defer auto-fit / cache column widths** — DONE (batching half). New
   `suspend_auto_fit`/`resume_auto_fit` (nesting-safe counter +
   `auto_fit_pending` set) and `set_many`; wired through `import_delimited`
   (also deleting a no-op pre-fit loop over empty columns), all `FillRange`/
   `RelFillRange`/`CopyFromTo`/duplicate-paste arms in `ops`, both
   `repair_all_formulas_after_*` functions, the log-replay loops (RAII
   `ReplayFitGuard`), the snapshot loader, and ODS table import. Final
   widths are identical (each column's last fit already saw final content).
   The per-write `resize_header_footer_width` retain scans in `set` were
   removed outright (proven no-ops: stored header/footer cells always
   satisfy the predicate). The remaining per-write `O(R)` scan is gone too:
   a `col_content_widths` index (global col → width → count) is maintained
   incrementally in `set` (overwrite adjusts both buckets, so narrowing
   writes shrink the max correctly) and read in `O(log W)` by
   `content_width_for_column`/`auto_fit_column`, with byte-identical results
   to the old scan (same `chars + 1` terms, same floor of 4, same `None`
   when empty). Rebuilt only where global-column identities shift
   (`set_main_size`, `move_main_cols`, main-width growth — which shifts
   right-margin globals — and `clear_cells`); row moves need no rebuild.
3. **Sparse structural moves** — DONE. `move_main_rows/cols` remap stored
   keys through the permutation in `O(n)`; row/col/header/footer occupancy
   is permuted alongside (no rebuild).
4. **Incremental shrink** — DONE. `shrink_to_content` reads maxima from the
   occupancy index (`next_back`, `O(log n)`) instead of scanning all maps.
5. **Formula evaluation caching + batch spill refresh** (§2.5) — DONE.
   Parse cache first (previous pass): thread-local `PARSE_CACHE` (cap 2048,
   clear-on-cap) serves every parse site; sound because the parse is a pure
   function of the string, so no invalidation is ever needed and failures
   are never cached. Dead `main_cols` params removed as a consequence.
   Then result memoisation: an always-on thread-local `EVAL_MEMO` keyed by
   `(grid id, sheet?, addr, allow_templates)`, consulted in place at both
   `eval_cell_inner` and `eval_cell_with_sheet` (so every caller — display,
   aggregates, VLOOKUP/SORT, spills — shares). Valid by construction:
   every grid/spill mutation hook clears it (`mark_spills_stale`,
   `bump_data_revision`, `set_eval_context`), cycle results (`CIRC`) and
   budget exhaustion (`LIMIT`) are never inserted, lookups are skipped
   while the cell is on the live visiting stack, cell-level evaluation
   always starts with empty bindings, and `RAND` is deterministic in
   `(seed, addr)` (seed bumps clear). Wall-clock `NOW`/`TODAY` freeze per
   pass — matching Excel recalc semantics. Only budget-*exhaustion* edge
   behaviour can differ (memoised values are fully evaluated, i.e. strictly
   more correct). Cap 8192, clear-on-cap, retention purely performance.
   Plus a `formula_cell_count` index so `refresh_spills` with zero formulas
   is O(1) (clear + flag, identical end state to the fixpoint loop).
6. **Aggregate sharing** — MOSTLY DONE. Sum/Mean/Min/Max stream via
   `fold_numbers_summable` (no `Vec` alloc; Min/Max keep `min_by`/`max_by`
   last-tie-wins semantics exactly). Dense area loops (`fold_`, `count_`,
   `sum_main_range`, the `references_all_empty` range arm) take a sparse
   path via the new `MainRangeEvalPlan` when the area dwarfs stored content:
   stored cells plus spill followers, sorted row-major, which preserves
   evaluation order, budget consumption and visiting-stack behaviour exactly
   (skipped empty cells provably consume no budget and parse non-numeric).
   Ranges with applicable header/row templates fall back to the dense loop;
   `COUNTIFS`/`SUMIFS`/`AVERAGEIFS` stay dense deliberately (criteria can
   match blanks). Cross-aggregate shared evaluation remains open.
7. **Borrowed iteration** (§1) — DONE. New `for_each_nonempty(&CellAddr,
   &str)` (same visit order as `iter_nonempty`) with a borrowing `Grid`
   implementation; `iter_nonempty` itself is now a thin wrapper over it.
   Migrated the filter-then-clone hot paths: `refresh_spills` anchor
   collection (clones formula cells only), both repair-all scans, and the
   trait defaults (`content_summary`, `main_range_eval_plan`,
   `stored_main_count`). One-shot export/ODS/balance callers keep the owned
   API.
8. **Bounded/indexed range aggregation** (§2.6) — DONE via the plan in 6
   (row-skip + stored-union + template-gated dense fallback).
9. **Single `sorted_main_rows` cache + revision counter** — DONE.
   `data_revision` rides on `mark_spills_stale` (every content/extent
   mutation funnels through it; width/format setters don't, so they keep
   the cache valid) plus explicit bumps in the spill setters (spills
   surface through `get`, which the comparator reads). `Grid::PartialEq`
   is now manual and compares observable state only (index/cache/revision
   excluded).
10. **ODT/export column span from the occupancy index** — DONE.
    `row_order`, `main_row_index_bounds_for_export` and
    `delimited_table_col_span_and_rows` consume the new
    `GridContentSummary` (one index read, `O(1)` set memberships) instead
    of per-row full scans; outputs are identical by construction (same
    sets, same predicates). `odt_content_xml` now emits the trimmed content
    span (`odt_col_span` = `ascii_col_bounds`, the same rule every other
    export uses) instead of all `0..total_cols` margin columns, expanded to
    cover spill followers (which render via `get` but are not stored
    content — otherwise a spill-only trailing column would vanish).
    Column styles keep global numbering; empty grids still emit one column.
    Measured on a 2-cell sheet: 5601 → 815 ODT bytes on disk (~7×, despite
    zip), ~400× less XML generated. Safe to do explicitly: the exporter has
    no in-repo callers (public API for external consumers) and no
    round-trip path depends on full-width ODT (ODS round-trips use the
    sidecar-layout ODS exporter, untouched).

## 7. Implementation notes, second pass (sparse ranges, borrowed iteration, parse cache)

## 8. Implementation notes, third pass (result memo, formula counter, width index, ODT)

- Result memo (`EVAL_MEMO`): wrappers keep the names `eval_cell_inner` /
  `eval_cell_with_sheet` so all ~15 existing callers (display, aggregates,
  lookups, spills) share hits unchanged; bodies renamed `*_uncached`.
  Clear-hooks: `Grid::mark_spills_stale`, `Grid::bump_data_revision`
  (spill setters), `set_eval_context`. No scopes or generations — validity
  by construction. Tests: budget-untouched second eval (no timing),
  precedent-edit invalidation, repeated CIRC reporting (cached CIRC must
  never leak into a fresh stack), plus the full suite guarding the wrapper
  refactor. The memo test for cycles also caught a mis-addressed test
  fixture (B2 vs A2), fixed as test-wrong.
- Formula counter: `note_formula_replace` in all five `set` branches
  (insert/overwrite/remove), `recount_formulas` after bulk drops, moves
  untouched (permutations preserve counts). `refresh_spills` early path
  asserts identical end state (empty maps + refreshed flag).
- Width index (`col_content_widths`): write-time bucket maintenance with
  overwrite adjust; growth-triggered rebuilds where globals shift. Tests
  cross-check against the deleted full scan after every op class.
- ODT: `odt_col_span` + three structural tests (span/numbering/values,
  spill-only column, degenerate empty grid).

- `MainRange::{area, contains}` helpers; `MainRangeEvalPlan
  {stored_sorted, has_template}` built by `Grid::main_range_eval_plan`
  (direct map access, no clones) with a scanning default on the trait.
  Sparse/fast-path selection rule: area `> 4 × stored_main_count`, bounding
  plan-build overhead to a fraction of the dense cost it replaces.
- Template detection (`range_has_template`) mirrors the three sources in
  `templated_formula` (column header, right-margin mirror, row key cell)
  and is conservative by construction: it may force the dense fallback when
  the template expression is inert, never the reverse. Includes the
  out-of-extent `Main(c)` header spelling (same edge as the col probe).
- Spill followers are union members of the sparse plan (they surface through
  `get`), so spilled values aggregate identically on both paths; the union
  is deduplicated (a cell with both a stored value and a spill evaluates
  once, as the dense loop does).
- New tests: plan contents/order/template flags (header/row/aggregate-label
  negatives included), sparse-vs-dense aggregate equivalence over a
  26×500 range for all six functions, spill inclusion in sparse sums,
  parse-cache hit/miss/no-negative-caching, cached-vs-uncached eval
  equality, and width-index-vs-brute-scan cross-checks (bucket totals,
  narrowing overwrites, growth-shifted right margins, moves/resizes).

## 6. Implementation notes, first pass (occupancy, batching, moves, cache)

- `Grid`'s new fields are derived/transient only; no log or file format
  changed, no `GridImpl` method was removed (three additive methods with
  default bodies: `content_summary`, `suspend/resume_auto_fit`, `set_many`).
- `set_many`/suspend batching preserves final widths: a column's last
  per-write fit already observed final content, and `auto_fit_column` is a
  no-op on empty columns, so clearing never leaves a different width behind.
- Regression tests added in `src/grid/mod.rs`: occupancy-vs-brute-force
  cross-checks across writes/overwrites/clears/moves/resizes
  (`assert_occupancy_matches_brute_force`, including count totals),
  out-of-extent header columns, stepwise shrink, sort-cache invalidation
  across set/move/resize (plus width-change cache retention),
  `set_many`-vs-sequential width equivalence, and suspend/resume nesting
  semantics. The out-of-extent test caught a genuine first-version probe
  bug (footer-`Right(0)`/header-`Main(2)` columns), fixed before merge.

## 5. Verdict on the question asked

- Are there `O(n²)` algorithms where `O(n)` (or `O(1)`) should do? **Yes —
  several, all of the same family**: an `O(n)` full-map scan (occupancy
  probe, width scan, shrink scan, spill anchor collect, per-row export
  filter) nested inside a loop over rows, columns, cells, keystrokes, or
  fixpoint passes. Individually each scan is linear and innocent; composed,
  import/replay/edit/render/export scale as `N·R`, `T·n`, `rows·n`,
  `V·eval`, `passes·n log n`.
- Are the data structures appropriate? **The primary store is; the index
  layer is missing.** `HashMap` sparse cells + sparse width/format maps are
  right for the unbounded sheet. What should be added is derived,
  incrementally maintained state: per-row/per-col occupancy (refcounts and/or
  ordered sets), cached column widths, cached sort order, and eventually a
  formula dependency/memoisation layer.
- Should we refcount rows/cols for blank row/col determination?
  **Yes**, keyed by objective region/index (not volatile global column),
  maintained in the write choke points, with spills explicitly excluded
  (or included — but decided and documented), ideally paired with an ordered
  occupied-set for bounds queries.

## 9. Post-optimization re-audit (fresh analysis; no code changed)

This section re-asks the original three questions against the current code
(occupancy + width + formula-count indexes, sort/parse/result caches,
batched auto-fit, sparse moves/ranges, ODT span trim all landed per §§4–8).
Findings below were verified by reading the code, not by running it; the
trait forwarding was explicitly re-checked (`impl GridImpl for Grid`
delegates `content_summary`, `main_range_eval_plan`, `stored_main_count`,
`formula_cell_count` to the index-backed inherent methods, so `GridBox`-only
callers really do get the fast paths — a silent-fallback bug here would
have reintroduced every scan it was meant to kill).

### 9.1 The three questions, re-answered

- **O(n²) where O(n) should do?** The systemic family from §2 (an `O(n)`
  full-map scan nested in a loop over rows/columns/cells/keystrokes) is
  gone: probes are `O(log n)`, width reads `O(log W)`, moves/shrinks
  `O(n)`-sparse or `O(log n)`, imports/replays fit once per column,
  range scans go sparse past `4×` stored density, formula parses happen
  once per distinct string, and cell results memoise across a pass.
  What remains is itemised in §9.3 — each entry is inherent (must touch
  every formula), bounded (fixpoint ≤ 8, dirty-gated), deliberate (blank
  matching), rare-path (save/sorted-export), or allocation-only.
- **Are the data structures appropriate?** Yes, with one standing con.
  Primary `HashMap` sparse cells plus derived `BTreeMap`/`HashMap` counts,
  width buckets, a formula counter, and three caches is the right shape for
  a sparse unbounded sheet. The con is structural, not performance: **five
  derived structures must now agree across ~ten mutation paths**
  (occupancy, widths, formula count, sort revision + memo clearing in
  `set`, `set_main_size`, moves, grows, clears, spill setters). A missed
  path fails *silently* (wrong blanks/widths/values, suite still green
  unless a cross-check covers it). Mitigation in place is the
  brute-force cross-check tests (`assert_*_matches_brute_force`, bucket
  totals, memo invalidation tests); the rule for future work is that **any
  new mutation path must update all five, with a cross-check test**.
  Memory overhead is ~2–4× the key storage (counts + buckets + caches);
  negligible next to the cell strings themselves.
- **Reference counting for blank row/col?** Closed: yes, and done.
  Counts (not booleans, so shared rows/cols survive partial clears),
  keyed by objective region/index, spills deliberately excluded, with
  `BTreeMap` row keys supplying the ordered-set half of the §3
  recommendation for free (`next_back` maxima for shrink, sorted keys for
  spans). Unordered `HashMap` col counts are sufficient — no col operation
  needs ordering today. No further structure needed here.

### 9.2 Fixed-cost table (before → after)

| Path | Before | After |
|---|---|---|
| `logical_row/col_has_content` | `O(n)` scan | `O(log n)` probe |
| `col_width` / viewport / bounds loops | `O(cols·n)`, `O(T·n)` | `O(cols)`, `O(T)` probes |
| per-write auto-fit | `O(R)` scan | `O(log W)` bucket read |
| import / replay / paste / repair fits | `O(N·R)` | one fit per touched column |
| `move_main_rows/cols` | `O(R·C)` / `O(R·1404)` gets | `O(n)` key remap |
| per-clear shrink | `O(n)` full scan | `O(log n)` index maxima |
| `sorted_main_rows` | recomputed per caller per frame | once per generation |
| export row lists | `O(rows·n)` clones | one summary + `O(1)` memberships |
| aggregates / `SUM` over sparse ranges | `O(area)` evals | stored∪spill union, same order/budget |
| formula parse | per eval/translate | once per distinct string |
| cell re-evaluation | per caller per frame | memoised within unchanged state |
| spill refresh, formula-free sheets | collect+sort `O(n log n)` | `O(1)` counter check |
| ODT emission | `rows × T` cells (T ≥ 1404) | content span (+spill cover) |

### 9.3 Remaining costs, honestly graded

1. **Per-frame clone-all residuals (the largest remainder).** Three sites
   still materialise every cell value per frame just to collect row-id
   sets: `visible_row_indices` in `ui/mod.rs` (~1715) and `ui_core.rs`
   (~613), and `view_row_order` in `ui/mod.rs` (~4056). Probes and evals
   underneath are now cheap, so this is allocation-only overhead — but it
   is still `O(n)` string clones per frame on the hot path. The fix is
   mechanical and safe (`for_each_nonempty` or one shared
   `content_summary`), deliberately left for a future sweep.
2. **`rendered_width_for_column` (ui_core ~529 + App method ~5568).**
   Full clone-all scan plus per-row *evaluated* display per column on fit
   paths. Cannot use the stored-width index (needs evaluated values); the
   result memo covers the eval share. Fit-scoped, not per-frame — acceptable.
3. **One `content_summary` per export call site.** Delimited export builds
   2–3 summaries (`delimited_table…`, `row_order`, bounds) at `O(occ)`
   each. Micro-item: thread a single instance through.
4. **Spill fixpoint when formulas exist.** Bounded (8), dirty-gated,
   borrowed anchors, memoised evals — the remaining cost is re-evaluating
   array anchors per keystroke. A dependency graph is still not recommended
   (see §4 item 5 reasoning); the formula counter already removes the
   empty case.
5. **`repair_all_formulas_after_*` is `O(F)` per insert.** Inherent — every
   stored formula must be examined — and now parse-cached, borrowed, and
   batch-fitted. No further action.
6. **Memo clearing is the new dominant per-write term.** Every `set()`
   clears the eval memo (`O(memo size)`, bounded by the 8192 cap).
   Correct and small, but it is the price of validity-by-construction; if
   it ever profiles, the alternative is a dirty flag checked at lookup
   (stale-hits-possible design — strictly worse safety). Same note covers
   the parse/result cap-clear thrash scenario (pathological only).
7. **Sort comparator cost per cache miss** (`grid.get` + value parsing per
   comparison, `O(R log R)`). Cached per generation; misses coincide with
   real mutations. Acceptable.
8. **Rare paths stay dense:** `save_workbook` iterates `R×C` with `get`
   (io ~228), `export_sorted_tsv` clones per comparison. Infrequent,
   bounded by sheet size, not worth index machinery.
9. **`COUNTIFS`/`SUMIFS`/`AVERAGEIFS` stay dense — deliberately.**
   Criteria can match blank cells, so sparse-skipping needs per-criterion
   blank-match analysis with Excel-compatible `""`/`=`/`<>` semantics.
   Wrong here means wrong numbers. No action.
10. **Accepted behaviour deltas** (all documented at implementation time,
    restated for the record): `NOW`/`TODAY` freeze per pass (matches Excel
    recalc); memo hits skip budget consumption (differs only under
    exhaustion, toward more-correct); `min_by`/`max_by` last-tie-wins
    preserved exactly in streaming; ODT span trim was explicit, not silent.

### 9.4 Recommended next sweeps (1–2 implemented; 3 stands)

1. ~~Migrate the three per-frame clone-all sites~~ DONE (§9.3.1): both
   `visible_row_indices` copies and `view_row_order` now build their
   header/footer row sets from one `content_summary()` instead of cloning
   every cell. Output-identical by construction (same sets; every path
   sorts/dedups before use); covered by the full render/parity suite.
2. ~~Thread one `GridContentSummary` per export operation~~ DONE
   (§9.3.3): `main_row_index_bounds_for_export` and `row_order` take the
   summary as a parameter; delimited/ODT/ascii paths build it once.
   Guarded by a hand-computed order test (gap-fill, left-only span
   membership, Main|Right-only bounds) plus the exact-matrix export suite.
3. Profile before anything else. The neighbouring candidates (per-pass
   memo scoping refinements, inverted row→cols iteration sets,
   cross-aggregate gather phases) all cost bug surface for gains the
   current numbers no longer justify without measurements.

### 9.5 Win32 input correctness (toolkit layer; app untouched)

Shifted characters were unreachable on the NWG backend: raw Win32 VKs
were forwarded, so Shift+9 typed `9` instead of `(`. Fixed in
`rustxWidgets/rswidgets` (verified end to end: `=(1+2)` commits and
evaluates to 3):

- `translate_vk` maps each KEYDOWN through `ToUnicodeEx` with the live
  keyboard state (layout-aware Shift pairs and capitals). Ctrl/Alt-held
  keys keep raw VKs so accelerators and menus still dispatch.
- Translated chars that numerically collide with app-dispatched VKs
  (`!"#$%&'()*` are `VK_PRIOR`..`VK_DOWN`, `.` is `VK_DELETE`) are
  *declined*, not passed on: passing `(` (0x28) would commit the edit
  and move down as `VK_DOWN`. The native control inserts the char and
  the change event resyncs app state.
- `vk_produces_wm_char` (pure, unit-tested) drives WM_CHAR suppression
  with set-not-arm semantics, so consumed keys don't double (`==` for
  `=`) and non-producing keys can't eat the next char via a stale flag
  (numpad adjusted for NumLock).
- Full Shift/Ctrl/Alt modifier mask (`modifier_state`; canvases decline
  pure-Ctrl so Ctrl+C never types `c`), and `Entry::set_text` moves the
  caret to end — `SetWindowText` leaves it at 0, which parked native
  insertions at the front (`(=`), failed the resync guard, and got
  clobbered by the next sync.

### 9.6 Fresh-document parity: CORRO_TEMPLATE + TOTAL seeds on GUI

Fresh TUI docs resolve content as template workbook (`CORRO_TEMPLATE`)
else built-in margin TOTAL seeds (`SheetState::new_seeded`). The GUI
backend ignored both (plain `WorkbookState::new`), so no-file GUI starts
were truly blank while the TUI showed the TOTAL row. Now
`gui::App::new_with_paths` mirrors the TUI exactly for genuinely fresh
docs (no paths): template load, seeded fallback with the same
"opened blank instead" status note on failure; load paths still replay
onto plain blank. The env lookup moved to `crate::io`
(`template_path_from_env`, beside `load_workbook_template`) because
`crate::ui` is ratatui-gated and invisible to gui-only builds.
`tests/gui_fresh_template.rs` pins seeds, TUI/GUI seed equality,
template load, and bad-template fallback. Negative result recorded:
a 4-column radio grid was tried for the Special Char picker and
reverted — GTK radio arrows navigate spatially, so Down*2 landed on
index 8 instead of 2; the single column keeps Down*n → nth choice on
every backend.
