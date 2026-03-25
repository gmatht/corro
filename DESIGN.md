# Corro design notes

Corro is a Rust terminal UI (TUI) spreadsheet-like tool built around an **append-only operation log**. Multiple running instances can follow the same file and converge by **tailing and applying ops**; the newest op for a given cell wins.

This document summarizes the architecture and key decisions implemented so far.

## Goals

- **Spreadsheet-ish editing in a TUI**: navigate, edit cells, display a small viewport.
- **Collaborative-ish via filesystem**: append-only JSONL file as the source of truth; instances watch and apply new lines.
- **Structural ops**: move row/column ranges without rewriting the whole file.
- **Aggregates**: define computed cells like SUM/MEAN/etc over a main-region range.
- **Sparse “infinite” sheet**: unbounded logical size without allocating huge dense grids.

Non-goals (currently):

- Full spreadsheet features (formulas, recalculation graph, formatting, etc.).
- Robust multi-writer conflict resolution beyond “last writer wins per cell”.
- Performance tuning for very large logs.

## Data model: five regions

The sheet is conceptually split into **five regions**:

- **Header**: 26 fixed rows labeled `^0`…`^25` (displayed as `^25`→`^0` top-to-bottom).
- **Footer**: 26 fixed rows labeled `_0`…`_25` (displayed as `_25`→`_0` top-to-bottom).
- **Left margin**: 10 fixed columns labeled `<A`…`<J`, indexed by main row.
- **Right margin**: 10 fixed columns labeled `>A`…`>J`, indexed by main row.
- **Main**: the spreadsheet body, addressed by `(row, col)` with Excel-like column letters `A..`.

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

## Operation log (append-only JSONL)

All edits are represented as ops serialized to JSON and appended as one line (JSONL).

Key op types (see `src/ops/mod.rs`):

- `SetCell { addr, value }`: set a cell’s raw string value.
- `SetAggregate { addr, def }`: store aggregate definition at a cell address.
- `SetMainSize { main_rows, main_cols }`: set main extent.
- `MoveRowRange { from, count, to }`: reorder main rows (and associated margin cells).
- `MoveColRange { from, count, to }`: reorder main columns (and shift header/footer cells for main columns).

Replay applies ops in order to rebuild `SheetState` (`grid` + `aggregates`).

### Concurrency model

- A process appends ops to the file it has opened.
- Other processes watch the file (via `notify`) and apply newly appended lines by tailing from a stored byte offset.
- Conflicts are resolved by log order: if two writers set the same cell, the later line wins on replay.

## Aggregates

Aggregates are stored as `AggregateDef { func, source }` keyed by an output `CellAddr`.

`source` is a `MainRange` describing an inclusive-exclusive rectangle over **main** cells.

Aggregate evaluation currently scans the range on demand and parses numeric cells (trim + `f64` parse).

Supported functions:

- `SUM`, `MEAN`, `MEDIAN`, `MIN`, `MAX`, `COUNT`

## TUI: viewport and navigation

### Viewport (current behavior)

The sheet view is tiny (BUG!):

- The sheet viewport is **3×3 logical cells** (`VIEW_DIM = 3`).
- It includes a top “column header” line plus 3 data lines.

It SHOULD instead show 3 *additional* blank lines for the main data, and a < > _ and/or ^ line if the user moves outside the displayed main data (e.g. with arrow keys). On a large real world spread sheet the viewport should be as big as will fit on the screen.

The viewport is currently centered around the cursor when possible, but is clamped so that the UI does not show unbounded emptiness:
It SHOULD be possible to move the cursor without moving the whole sheet.

- It finds the first/last “interesting” row/col (has content, has aggregates, or is the cursor).
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
- `src/ops/mod.rs`: op schema, serialization, replay onto `SheetState`.
- `src/io/mod.rs`: load/tail/apply ops; append writer; file watcher integration.
- `src/agg/mod.rs`: aggregate evaluation and display helpers.
- `src/ui/mod.rs`: ratatui app loop, input handling, and rendering.

## Known limitations / next steps

- The UI is deliberately tiny and not yet a usable spreadsheet.
- The log format is not versioned; migrations are not handled.
- No durability guarantees beyond append; no compaction/snapshotting.
- Aggregate evaluation is naive (range scanning each time).

