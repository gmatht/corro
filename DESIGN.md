# Corro design notes

Corro is a Rust terminal UI (TUI) spreadsheet-like tool built around an **append-only operation log** and a workbook model with multiple sheets. Multiple running instances can follow the same file and converge by **tailing and applying ops**; the newest op for a given cell wins.

This document summarizes the architecture and key decisions implemented so far.

## Goals

- **Spreadsheet-ish editing in a TUI**: navigate, edit cells, display a small viewport, and switch between sheets with a bottom tab bar.
- **Collaborative-ish via filesystem**: append-only text log as the source of truth; instances watch and apply new lines.
- **Structural ops**: move row/column ranges without rewriting the whole file.
- **Formulas**: cells whose value starts with `=` evaluate for display and for numeric range aggregation.
- **Special rows/columns**: margin labels (`SUM`, `=TOTAL`, …) drive computed totals over main data; the stored total directive is `=TOTAL` and the TUI shows `TOTAL` (see below).
- **Sparse “infinite” sheet**: unbounded logical size without allocating huge dense grids.
- **Workbook tabs**: stable numeric sheet IDs, per-sheet titles, and cross-sheet formula references.

Non-goals (currently):

- Full Excel compatibility (full function set, formatting, etc.).
- Robust multi-writer conflict resolution beyond “last writer wins per cell”.
- Performance tuning for very large logs.

## Data model: five regions

The sheet is conceptually split into **five regions**:

- **Header**: up to `999999999` fixed logical rows indexed `0…999999998` internally. Header references use `~`; display/reference text numbers them from `~999999999` at the top to `~1` nearest the main region.
- **Footer**: up to `999999999` fixed logical rows indexed `0…999999998` internally. Footer references use `_`; display/reference text numbers them from `_1` nearest the main region to `_999999999` at the bottom.
- **Left margin**: `MARGIN_COLS` fixed columns on the left of the main region. `MARGIN_COLS` is currently `26 * 26` (676). Margin references use `[` plus mirrored Excel-style names, so the column nearest the main region is `[A`, then `[B`, etc. outward.
- **Right margin**: another 676 fixed columns on the right of the main region. Right-margin references use `]A`, `]B`, etc. from the main region outward.
- **Main**: the spreadsheet body, addressed by `(row, col)` with Excel-like column letters `A`, `B`, ..., and 1-based row numbers in user-facing references.

Internally, columns in header/footer and rendering code use one **global column index**:

- `0 .. MARGIN_COLS` is the left margin.
- `MARGIN_COLS .. MARGIN_COLS + extent_main_cols` is the main region.
- `MARGIN_COLS + extent_main_cols .. total_cols()` is the right margin.

Addresses are represented by `CellAddr` in `src/grid/mod.rs`:

- `Header { row, col }`, `Footer { row, col }` use header/footer row index plus global column index.
- `Main { row, col }` uses main-relative row and column indices.
- `Left { col, row }` and `Right { col, row }` use margin-relative column plus main row.

`GridBox` wraps a `Box<dyn GridImpl>`. Most callers use the boxed abstraction, while the current concrete implementation is `Grid`.

## Storage: current sparse hybrid

The current concrete `Grid` is sparse across editable cell regions:

- **Main cells**: `HashMap<(u32, u32), String>`, keyed by main row and main column. Absent keys are empty cells. This is similar to Gnumeric which also uses a hashmap, though I am not sure it is optimal.
- **Left margin cells**: `HashMap<(u32, MarginIndex), String>`, keyed by main row and left-margin column.
- **Right margin cells**: `HashMap<(u32, MarginIndex), String>`, keyed by main row and right-margin column.
- **Header/footer cells**: sparse `HashMap<(u32, u32), String>` maps keyed by special-row index and global column. Absent keys are empty cells.
- **Column widths**: default `max_col_width` plus sparse per-global-column overrides in `HashMap<usize, usize>`.
- **View sort**: `Vec<SortSpec>` containing global main-column indices and descending flags.
- **Formatting**: sparse maps for all-column, data-column, special-column, and exact-cell overrides.
- **Formula spills**: transient `spill_followers: HashMap<CellAddr, String>` and `spill_errors: HashMap<CellAddr, &'static str>`, layered over stored cell values by `get()`.
- **Volatile formulas**: `volatile_seed: u64`, bumped to invalidate/recompute volatile formula output.

The main region has a logical extent:

- `extent_main_rows: u32`
- `extent_main_cols: u32`

These extents:

- start at at least `1×1`
- grow when setting non-empty main or margin cells outside the current row extent
- grow when setting non-empty main cells outside the current column extent
- grow when navigating off the bottom/right edge of main (see UI rules)
- can be set explicitly via `SetMainSize`, which prunes main and margin cells beyond the new row/column extent

When the main column count changes, header/footer rows and global-column metadata are remapped so the right margin remains anchored after the main region.

### Current structure tradeoffs

Pros:

- Sparse main and margin storage keeps mostly-empty sheets cheap.
- `HashMap` gives simple expected O(1) lookup and update for random cell edits.
- Sparse header/footer rows allow very large logical header/footer limits without allocating empty rows.
- `GridImpl`/`GridBox` leaves room to experiment with alternate storage without rewriting every caller.

Cons:

- Header/footer logical row positions can be very far apart, so export paths must preserve sparse row positions without scanning every blank row.
- Row and column moves rebuild maps by scanning logical extents, which is wasteful for large sparse sheets.
- Row/column content checks scan sparse maps instead of using maintained indexes.
- Global-column metadata needs careful remapping whenever main columns grow, shrink, or move.
- `HashMap` iteration order is nondeterministic, so callers that need stable output must sort or otherwise normalize.

## Possible alternative grid structures

### Extent tree with global-column dense rows

Use a row-ordered extent tree, similar to a rope or B+ tree, for the main row region. Each node stores the number of rows below it and cached aggregate summaries. Leaf nodes store row spans. Materialized rows use a global column dictionary that maps sparse logical columns to compact dense column indexes.

This combines two ideas:

- the tree answers “which row is logical row N?” and supports row-range movement by split/splice
- the global column dictionary answers “which dense slot represents logical column C?” and gives compact row-local access for table-shaped data

Sketch:

```rust
struct ExtentGrid {
    root: RowNode,
    columns: ColumnDict,
    extent_main_rows: u32,
    extent_main_cols: u32,
}

struct ColumnDict {
    logical_to_dense: HashMap<GlobalCol, DenseCol>,
    dense_to_logical: Vec<GlobalCol>,
    live: BitSet,
}

enum RowNode {
    Internal {
        summary: NodeSummary,
        children: Vec<RowNode>,
    },
    Leaf {
        summary: NodeSummary,
        spans: Vec<RowSpan>,
    },
}

enum RowSpan {
    Empty { len: u32 },
    Rows(Vec<Row>),
}

struct Row {
    cells: RowCells,
    stats: RowStats,
}

enum RowCells {
    Empty,
    SparseSmall(Vec<(DenseCol, CellSlot)>),
    Dense {
        cells: Vec<CellSlot>,
        present: BitSet,
    },
}

struct NodeSummary {
    rows: u32,
    non_empty_cells: u64,
    cols: HashMap<DenseCol, ColStats>,
}

struct ColStats {
    non_empty: u32,
    numeric: u32,
    total: f64,
    min: f64,
    max: f64,
}
```

Rows should be stored adaptively:

- `Empty` for blank rows.
- `SparseSmall` for rows with only a few populated cells.
- `Dense` for rows with enough populated cells, or rows accessed often enough, that dense indexing is cheaper than scanning/searching sparse pairs.

The row's dense vector may be shorter than the number of globally known non-empty columns. Missing dense slots are treated as empty. This keeps rows with one populated cell from paying for every active column in the sheet.

Internal and leaf summaries are recomputed from children or rows:

- `rows` is the order-statistic count used for row navigation.
- `non_empty_cells` supports bounds detection and export trimming.
- `cols` stores per-column summaries only for columns present in that subtree.
- `total`, `min`, `max`, and `numeric` support footer and aggregate rows without scanning every cell in the target range.

Range aggregate evaluation can descend the tree, combine fully-covered node summaries, and inspect only boundary leaves. This makes vertical aggregates over large row ranges much cheaper than the current `collect_numbers` scan. Horizontal row aggregates still use the row's own `RowStats` or scan that row's populated cells.

Pros:

- Preserves sparse/infinite row behavior by representing blank row ranges as extents.
- Makes row moves a tree split/splice problem instead of rebuilding all cell maps.
- Gives footer and aggregate rows fast access to cached per-column stats.
- Dense row storage is cache-friendly for table-shaped data.
- Sparse node summaries and adaptive rows avoid the worst case where every row allocates every globally active column.

Cons:

- Considerably more complex than the current `HashMap` grid.
- Every cell edit must update row stats and all ancestor summaries.
- Column deletion/compaction needs tombstones or a global dense-index remap.
- Column moves need a logical ordering layer so dense indexes do not have to be rewritten everywhere.
- Formula-backed numeric values require dependency invalidation before they can safely participate in cached summaries.
- `MEDIAN` is not supported by `total`/`min`/`max` summaries; it needs a scan or a heavier ordered-value summary.

### Fully sparse cell map

Use one `HashMap<CellAddr, CellData>` for all regions, with separate extents and metadata maps.

Pros:

- One storage path for all cells, including header/footer.
- Memory scales with non-empty cells only.
- Easier to serialize, diff, and iterate all cells uniformly.

Cons:

- Rendering header/footer rows needs many sparse lookups unless row/column indexes are added.
- Region-specific invariants become less obvious.
- Exact `CellAddr` keys are larger and can be slower than tuple keys for hot main-cell access.

### Row-oriented sparse grid

Use `BTreeMap<u32, Row>` or `HashMap<u32, Row>`, where each row stores sparse columns such as `BTreeMap<u32, CellData>` or `HashMap<u32, CellData>`.

Pros:

- Efficient row rendering and row moves.
- Natural fit for line-oriented operations, exports, and row-content checks.
- Rows can own their margin and main cells together.

Cons:

- Column moves and column scans touch many rows.
- Empty-row cleanup and row metadata need discipline.
- Header/footer still need either special rows or separate storage.

### Column-oriented sparse grid

Use `HashMap<u32, Column>` or `BTreeMap<u32, Column>`, where each column stores sparse rows.

Pros:

- Efficient column formatting, width calculation, sorting inputs, and column moves.
- Natural place for column metadata such as width and format.
- Column-content checks are cheap.

Cons:

- Row rendering and row moves touch many columns.
- Row-oriented exports and formulas over ranges can become less cache-friendly.
- Margins and header/footer still need special handling or encoded columns.

### Chunked sparse tiles

Split the main region into fixed-size tiles, such as 32x32 or 64x64, stored in a `HashMap<(tile_row, tile_col), Tile>`. Each tile can be dense or internally sparse.

Pros:

- Good locality for viewport rendering and rectangular range formulas.
- Sparse at large scale while avoiding one allocation/hash entry per cell.
- Tiles can later support caching, dirty flags, or compression.

Cons:

- More complex addressing and move logic.
- Row/column insert or move across tile boundaries is harder.
- Poor fit for very skinny or very wide sparse sheets unless tile shape is tuned.

### Dense `Vec<Vec<CellData>>`

Store the whole logical sheet densely.

Pros:

- Very simple indexing and rendering.
- Fast contiguous scans for small to medium dense sheets.
- Row and column moves are straightforward vector operations.

Cons:

- Does not match the sparse “infinite” design goal.
- Memory grows with every blank cell in the logical extent.
- Large cursor movements or imported sparse data can allocate huge empty regions.

### Log-derived overlay

Keep the append-only operation log as the primary structure and maintain indexes or materialized views for the viewport, formulas, and exports.

Pros:

- Aligns directly with Corro’s append-only file format.
- Can support undo/history and conflict inspection naturally.
- Materialized views can be tuned independently for UI and formula workloads.

Cons:

- Requires invalidation and replay machinery.
- Harder to reason about current-cell lookup without indexes.
- More moving parts than the current in-memory `Grid`.

## Operation log (append-only text)

All edits are represented as short text lines and appended one per line.
Older JSON log lines are still accepted on replay for compatibility.

Key op types (see `src/ops/mod.rs`):

- `SetCell { addr, value }`: set a cell’s raw string value (plain text or a formula beginning with `=`).
- `SetMainSize { main_rows, main_cols }`: set main extent.
- `MoveRowRange { from, count, to }`: reorder main rows (and associated margin cells).
- `MoveColRange { from, count, to }`: reorder main columns (and shift header/footer cells for main columns).

Replay applies ops in order to rebuild `SheetState` (currently: `grid` only).

**Breaking change:** older logs that used `set_aggregate` are no longer accepted; use `SetCell` with `=SUM(…)` (or other formulas) instead.

### Concurrency model

- A process appends ops to the file it has opened.
- Other processes watch the file (via `notify`) and apply newly appended lines by tailing from a stored byte offset.
- Conflicts are resolved by log order: if two writers set the same cell, the later line wins on replay.

## Formulas (`=…`)

- Stored as normal cell text. After a leading `=`, the rest is parsed as an expression (`src/formula/mod.rs`).
- **Display**: the TUI shows the **evaluated** result (or `#CIRC`, `#PARSE`, etc.); **edit** mode shows the raw string.
- **Operators**: `+ - * /`, parentheses, unary `-`.
- **References** (see `src/addr.rs`): main `A1`, headers `^A` / `^A,A1`, footers `_B` / `_B,A1`, margins `<0` / `<0,1`, `>0` / `>0,1` (same rules as `ADDR: value` shorthand in the UI).
- **Ranges** (main only): `A1:B2` (rectangle between corners, inclusive).
- **Functions**: `SUM(range|expr)`, `IF(cond, a, b)` (condition is “truthy” if it evaluates to a non-zero number). Function names are case-insensitive.
- **Cycles**: recursive references are detected; evaluation returns `#CIRC`.
- **Budget**: evaluation is step-limited to avoid pathological graphs.

## Aggregates and special rows/columns

There is **no** separate `SetAggregate` op. Numeric aggregation over a main-region rectangle uses the same rules as formulas: each cell contributes its **effective numeric value** (plain number or formula result).

**Special header/footer behavior** (UI only): certain margin cells label a row or column as `SUM`, `MEAN`, etc. The TUI then **computes** that aggregate over the corresponding main strip using `compute_aggregate` in `src/agg/mod.rs` (`AggregateDef` + `AggFunc` are **not** persisted; they are constructed at render time). Those totals treat formula cells like any other cell: the formula’s numeric value is used.

Supported aggregate names: `SUM` (bare), `=TOTAL` (stored; displayed as `TOTAL`), `MEAN`/`AVERAGE`/`AVG`, `MEDIAN`, `MIN`/`MINIMUM`, `MAX`/`MAXIMUM`, `COUNT`.

## TUI: viewport and navigation

### Viewport (current behavior)

The sheet view is tiny (BUG!):

- The sheet viewport is **3×3 logical cells** (`VIEW_DIM = 3`).
- It includes a top “column header” line plus 3 data lines.

It SHOULD instead show 3 *additional* blank lines for the main data, and a < > _ and/or ^ line if the user moves outside the displayed main data (e.g. with arrow keys). On a large real world spread sheet the viewport should be as big as will fit on the screen.

The viewport is currently centered around the cursor when possible, but is clamped so that the UI does not show unbounded emptiness:
It SHOULD be possible to move the cursor without moving the whole sheet.

- It finds the first/last “interesting” row/col (has content, or is the cursor).
- It limits the window to show at most `MAX_EDGE_BLANK = 2` blank rows/cols beyond the interesting bounds, while keeping the cursor visible.

This is the mechanism that implements: “use a sparse matrix for ‘infinite’ size, but limit the number of blank rows/cols you show to the user.”

### Cursor and growth rules

The cursor is tracked in **logical** sheet coordinates (header + main + footer; left + main + right).

Navigation keys:

- Movement: arrows and `hjkl`
- Quit: `q` / `Ctrl+Q`
- Help: `?`

Main extent growth on navigation:

- Moving **Down** from the **last main row** grows `extent_main_rows` by 1.
- Moving **Right** from the **last sheet column** grows `extent_main_cols` by 1 (and header/footer width).

Cursor “clamp” keeps the cursor within `HEADER + main + FOOTER` rows and within `total_cols`.

### Editing and selection

Modes:

- Normal
- Edit (Enter / `e`)
- Visual selection (`v`)
- Aggregate picker (`a`)
- Open file path (`o`)
- Help (`?`)

Range moves:

- `r`: move selected full main rows
- `c`: move selected full main columns

## File layout

Core modules:

- `src/grid/mod.rs`: regions, addressing, sparse storage, row/col moves.
- `src/addr.rs`: shared Excel / `^` / `_` / `<` / `>` reference parsing.
- `src/formula/mod.rs`: formula parse and evaluation.
- `src/ops/mod.rs`: op schema, serialization, replay onto `SheetState`.
- `src/io/mod.rs`: load/tail/apply ops; append writer; file watcher integration.
- `src/agg/mod.rs`: aggregate evaluation and `cell_display` (raw string).
- `src/ui/mod.rs`: ratatui app loop, input handling, and rendering.

## Known limitations / next steps

- The UI is deliberately tiny and not yet a usable spreadsheet.
- The log format is not versioned; migrations are not handled.
- Single-sheet logs remain append-only text. Multi-sheet workbooks save as workbook snapshots with sheet headers and per-sheet cell lines.
- Formula language is minimal; errors are coarse (`#PARSE`, `#CIRC`, …).
