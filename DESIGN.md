# Corro design notes

Corro is a Rust terminal UI (TUI) spreadsheet-like tool built around an **append-only operation log**. Multiple running instances can follow the same file and converge by **tailing and applying ops**; the newest op for a given cell wins.

This document summarizes the architecture and key decisions implemented so far.

## Goals

- **Spreadsheet-ish editing in a TUI**: navigate, edit cells, display a small viewport.
- **Collaborative-ish via filesystem**: append-only text log as the source of truth; instances watch and apply new lines.
- **Structural ops**: move row/column ranges without rewriting the whole file.
- **Formulas**: cells whose value starts with `=` evaluate for display and for numeric range aggregation.
- **Special rows/columns**: margin labels (`SUM`, `TOTAL`, …) drive computed totals over main data (see below).
- **Sparse “infinite” sheet**: unbounded logical size without allocating huge dense grids.

Non-goals (currently):

- Full Excel compatibility (full function set, formatting, etc.).
- Robust multi-writer conflict resolution beyond “last writer wins per cell”.
- Performance tuning for very large logs.

## Data model: five regions

The sheet is conceptually split into **five regions**:

- **Header**: 26 fixed rows indexed `0…25` internally; row labels in the TUI use **`^Z` (top) through `^A` (bottom)** — letter derived as `(Z - logical_row)`.
- **Footer**: 26 fixed rows; labels **`_A` (top) through `_Z` (bottom)** (`_` + `(A + row_index)`).
- **Left margin**: 10 fixed columns labeled `<0`…`<9` in column headers (see `col_header_label`), indexed by main row.
- **Right margin**: 10 fixed columns labeled `>0`…`>9`.
- **Main**: the spreadsheet body, addressed by `(row, col)` with Excel-like column letters `A`…

Addresses are represented by `CellAddr` in `src/grid/mod.rs`:

- `Header { row, col }`, `Footer { row, col }` use **global column indices**.
- `Main { row, col }` uses **main-relative indices**.
- `Left { col, row }` and `Right { col, row }` use main row + margin column.

## Storage: sparse main + sparse margins

To support “infinite” size without dense allocation, the grid uses sparse maps:

- **Main cells**: `HashMap<(u32, u32), String>`
- **Left margin**: `HashMap<(u32, u8), String>`
- **Right margin**: `HashMap<(u32, u8), String>`

Header/footer remain dense `Vec<Vec<String>>` sized to the full visible width.

The main region has a logical extent:

- `extent_main_rows: u32`
- `extent_main_cols: u32`

These extents:

- start at at least `1×1`
- grow when setting non-empty cells outside current extents
- grow when navigating off the bottom/right edge of main (see UI rules)
- can be set explicitly via `SetMainSize` (prunes cells beyond extent)

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

Supported aggregate names: `SUM`/`TOTAL`, `MEAN`/`AVERAGE`/`AVG`, `MEDIAN`, `MIN`/`MINIMUM`, `MAX`/`MAXIMUM`, `COUNT`.

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
- No durability guarantees beyond append; no compaction/snapshotting.
- Formula language is minimal; errors are coarse (`#PARSE`, `#CIRC`, …).
