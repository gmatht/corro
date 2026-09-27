# 0.7.0 – Windows 95 support (WIP); new release artifacts

## Pinch to zoom, and platform-aware drag on the grid
- **Pinch to zoom** for the grid widget on Android and iOS. `GridView` gained
  `zoom`/`set_zoom`/`zoom_by`/`reset_zoom`/`on_zoom`/`zoom_limits`, and the
  underlying `SpreadsheetModel` a `zoom` factor that its metrics
  (`char_w`/`row_h`/`header_h`/`row_label_w`), the renderer, `col_x`/`col_width`
  and `cell_at` all honour — so what is drawn and what a finger hits scale
  together and cannot drift apart. Clamped to 0.4x..4x.
- Android wires it through `ScaleGestureDetector` (`SheetView.nativeZoom`) and
  iOS through `UIPinchGestureRecognizer` (`corro_ios_canvas_zoom`); both feed
  the same span-ratio API in `gui_backend`, so one pinch behaves identically on
  both platforms. A double tap resets to 1x.
- **Platform-aware drag**: a *mouse* drag extends the selection (desktop
  behaviour), while a *finger* drag pans the sheet — and only selects after a
  long press arms it (the phone's substitute for shift-click). This is
  `GridView`'s `PointerKind`/`DragOutcome` state machine, mirrored for corro's
  live sheet in `gui_backend::mobile_gesture`; the gesture policy lives in
  Rust, so the hosts only report the raw stream plus the platform facts (is
  this a finger? has it held still?).
- The pixel→cell conversion for touch scrolling moved into Rust, where the live
  metrics are: a pinch changes them, so a cell size cached on the Java/ObjC side
  would pan by the stale amount after a zoom. The hosts keep only the sub-cell
  remainder.
- Docs: `ANDROID_GUIDELINES.md` §3.1 and `IOS_GUIDELINES.md` §3.1 document the
  gesture contract and the exports; the widget-level API is
  `rswidgets::gridview::GridView::pointer_down/move/up/long_press/cancel`.
- Fixed `host_density()`'s mobile signatures (they promised `f64` while the
  adapters return `Option<f64>`), which had left the Android/iOS/macOS targets
  failing to compile.

## Clicking a row/column header selects that row/column (LibreOffice parity)
- A click in the gutter chrome — the `A`/`B`/… column strip or the `1`/`2`/…
  row-label strip — now selects the whole main-body row/column, as LibreOffice
  does. Clicking the top-left corner box selects the whole sheet. The padlock
  affordance still owns its own hit test (toggling a pin never moves the
  selection), exactly as before.
- The extent is one shared rule: `ui_core::main_row_selection_span` /
  `main_col_selection_span` define which cells a selected row/column covers.
  ratatui's row/column commands store that span directly; the GUI stores the
  active cell plus `SelectionKind` and the renderer expands to the same span
  (locked by a render test), so the two backends cannot disagree.
- The existing rectangle model could not express this: a full-width rectangle
  would light up *every* column header when a row was selected. `SelectionKind`
  (previously private to `ui/mod.rs`) moved to `crate::grid` beside
  `SheetCursor` and is now carried by `CoreApp`, so the GUI renderer narrows
  header coverage and body highlight per axis (`Rows` → row gutters only,
  `Cols` → column headers only). Plain navigation collapses the anchor and
  resets the kind; extending a selection keeps it (ratatui parity).
- The active cell stays in the clicked row/column at the previously active
  other-axis position (LibreOffice keeps the cell there), while the *extent*
  comes from the shared span and the kind — so the formula bar and the next
  keystroke stay about the row/column the user clicked.
- Fixed: the corner box set the anchor and the cursor but left
  `last_row`/`last_col` on the anchor, and the renderer reads those as the
  *cursor* — so the visible selection collapsed to A1 even though a whole-sheet
  span had been stored. All three gutter selections now go through
  `apply_selection`, which reports the cursor cell the caller must record;
  the rule is unit-tested.
- **Repeat-click cycling:** clicking the *same* header again narrows the
  selection to that target's data cells (the non-blank main span in that
  row/column; the sheet's used range for the corner), and a third click widens
  it back. The cycle is keyed by the clicked target, so a click on a different
  header — or any navigation, body click or padlock toggle — starts the next
  selection at full again.

## Deduplication sweep: two more items removed, and the rest is boundary not duplication
- **Orphaned doc comments** in `pnc_backend.rs`: `about_body`/`help_body` bodies
  were deleted when the About/Help text moved into the shared presenter, but
  their doc comments were left dangling. Removed.
- **`GuiMenuPresenter::show_dialog` no longer sniffs dialog text.** It compared
  the body against `ui_core::about_page_body()` to choose between the About and
  Help widgets — which would silently open the *wrong* dialog if that shared
  wording ever changed. Replaced with a typed
  `actions::InfoDialog { About, HelpFull, Other(title, body) }`, so each backend
  selects its widget from the variant instead of parsing the content. The
  recorder test follows the type.
- Swept the remaining candidates and found no further duplication worth
  removing:
  * the six `MenuPresenter` methods are an **interface boundary** — they differ
    in mechanism by construction (`set_status` updates status plus the formula
    bar's trailing text on a terminal vs `sync_chrome_labels` on the GUI;
    `show_dialog` uses one SGR box vs three bespoke GTK widgets; `show_keybinds`
    lists terminal keys vs GTK's). Merging them would need capability queries,
    which is worse than six honest implementations.
  * `pnc_backend`'s four remaining local helpers (`cursor_column_letter`,
    `splice_special_pick`, `push_anchor_to_widget`, `show_info_dialog`) are thin
    terminal-specific plumbing over already-shared decisions — e.g.
    `cursor_column_letter` is a 4-line wrapper over `addr::excel_column_name`,
    and the special-char splice decision already lives in
    `special_picker::SpecialSplice` precisely because each backend splices at
    its own caret.
  * the GUI's geometry helpers (`char_w`, `row_h`, `font_size`, …) have no
    counterpart in the terminal host, which paints through the toolkit's cell
    grid rather than computing pixels. Checked, not assumed.
  * constants (`MARGIN_COLS`, `HEADER_ROWS`, 131 references) are already
    imported from `crate::grid` by both.

## `pnc_backend`'s cell sink deliberately retained (transition investigated, rejected)
- The GUI's duplicate `CellSink` was deleted (see below). The matching question
  for the terminal — retype the widget's display grid to `SpreadsheetModel` so
  both hosts use `GridSink` — was investigated and **rejected**; the reasoning is
  written up in `docs/WHY_PNC_AND_GUI_RS.md` §6a.
- The case *for* it measured well: `pancurses.rs` opens the widget's display
  grid with `PcWidgetKind::Spreadsheet { grid, .. }` **136 times** (~19,600
  chars, median line 134) for only ~1.5 field accesses each, which is a genuine
  defect. But the fix targets the wrong layer: that grid is
  `rswidgets::core::Grid`, a **public toolkit type** (public module, all-public
  fields), and `SpreadsheetModel` is the *renderer's* model carrying corro's
  chrome (`formula_bar_address`, `border_title`, `tab_titles`, `menu_text`).
  Retyping would invert toolkit↔host layering permanently and drag
  host-specific concepts into a generic widget — for a one-time line saving.
- Also recorded: three of that grid's fields (`top_row`, `left_col`,
  `col_width`) *look* dead (no writer in `pancurses.rs` beyond an initialiser)
  but are public API — a scroll origin and a default column width — that any
  grid-widget consumer may use. Pruning public API on one consumer's usage is
  not a cleanup.
- The real fix, if the boilerplate is attacked: give the widget's display state
  a handle or accessor *inside* `pancurses.rs`, which collapses the
  destructures with no public-API or layering change.
- Grep trap for this area, also recorded: **two unrelated types are named
  `Grid`** — `rswidgets::core::Grid` and `corro::grid::Grid`. Searches spanning
  both silently mix them.

## The GUI host now uses the shared `CellSink` (one of two sinks deleted)
- **`GuiCanvasSink` is gone.** The previous entry added `render::GridSink` but
  did not wire either backend to it, so both old sinks were still live; that
  entry's "the duplication is removed" was premature and is corrected here.
- `gui_backend` now fills a `GridView` through `GridSink`, the same adapter the
  terminal host is moving to, and `render_to` reads the *model* instead of a
  private pair of `HashMap`s. Two consequences beyond the deletion:
  * `GuiCanvasSink`'s `raw_values` and `cursor_pos` were **write-only** (no
    reads anywhere), so those values now live in the model where a host can
    actually consult them;
  * the style now round-trips bit → enum → bit, so
    `CellDisplayStyle::from_style_bits` was added as the inverse of
    `to_pancurses_style`, with a test that every style survives the round trip
    and that an unknown bit degrades to `Default` rather than panicking a
    repaint.
- Remaining: `pnc_backend`'s `SpreadsheetSink`. Swapping it is **not** a
  like-for-like change — the widget stores `core::Grid` while `GridView` wraps
  `SpreadsheetModel`, and `GridView::from_grid` copies, so a naive swap would
  trade 24 lines of duplication for a per-frame copy. It needs the widget to
  store the model directly, which is a change to the terminal renderer's
  storage and belongs in its own commit with the parity tests as the gate.
- Verification note: a `render_to` change is gated by the GUI render tests, not
  only the unit tests. `gui_edit_parity` reports the **same 2 failures with and
  without** this change (both single-threaded, same 16 passes) — pre-existing
  and unrelated; without `--test-threads=1` the set is flaky (2–5), which is why
  the comparison was made single-threaded.

## The two `CellSink`s are now one (`render::GridSink`)
- **The last genuine feature duplication between the two hosts is removed.**
  Each backend had its own `CellSink`: `SpreadsheetSink` (24 lines, writing into
  the pancurses widget) and `GuiCanvasSink` (28 lines, writing into four
  `HashMap`s). They differed only in *where* they put the cells — and the GUI's
  four maps (`cells`, `styles`, `raw_values`, `cursor_pos`) are exactly the
  fields `SpreadsheetModel` already has, so the difference was never real.
- `render::GridSink` is the one adapter, writing into a `GridView` (the model
  plus the surface to draw it on). `GridView` gained `set_raw_cell`/`raw_cell`
  so it covers the raw-value half; the style-bit conversion lives in the sink
  because the u8 style numbering is the *model's* vocabulary, and a test asserts
  it agrees with `CellDisplayStyle::to_pancurses_style` so the two cannot drift.
- Writing the tests found two real hazards worth recording:
  * `grid::HEADER_ROWS` is a **sentinel** (999_999_999), not a count — it is the
    logical row of the topmost header. Passing it where a *count* is expected
    does not mis-render, it makes `compute_row_agg_func` iterate ~10^9 rows and
    the caller hangs. Diagnosed by taking a stack dump of the hung test
    (`compute_row_agg_func` with `hr=999999999` in the frame, `display_rows`
    built as `0..hr+mr`). The comment on the new test records the trap.
  * Holding a temporary `Ref` across a second `borrow()` (`&m.borrow().cells`
    then `m.borrow()`) deadlocks the `RefCell` — the test now reads the model in
    one borrow.
- Note: `GuiCanvasSink`'s `raw_values` and `cursor_pos` were **write-only** (no
  reads anywhere in the file), so routing through the model both unifies the
  sinks and puts those two values somewhere they are actually consulted.

## Loop ownership is now a queryable capability, not an undocumented split
- The full reasoning — what is shared, what is not, the measurements, and why
  the terminal loop is deliberately not unified — is written up in
  `docs/WHY_PNC_AND_GUI_RS.md`.
- Adding `BackendApp::owns_event_loop()`. Every backend already implements
  `run()`, so the *entry point* is uniform; what differs is who dispatches. Two
  shapes exist and both are correct for their environment:
  * **delegating** (default `false`) — a toolkit or OS framework owns dispatch
    and calls back: GTK (`g_main_loop_run`), NWG (`GetMessageW`), and
    WASM/Android/iOS/macOS, whose `run` returns immediately because the browser
    / Activity / UIKit owns the loop. These quit by signalling the owner
    (`quit_main_loop` reaches through a published loop pointer).
  * **owning** (`true`) — nothing else wants to dispatch for us, so the backend
    reads input directly and *is* the loop: pancurses (`getch`), zork (REPL),
    ratatui (crossterm). No loop pointer, no pumping, no deferred redraw.
- A host that must behave differently can now ask instead of using `cfg` — and
  the default is `false`, so a backend that delegates needs no boilerplate.
- This is deliberately **not** unified into one model. Forcing the delegating
  contract on a terminal would mean inventing a loop pointer whose only possible
  caller is the code already running, a `present()`-style context pump for
  widgets that have no unmapped state, and a frame-clock hop that adds latency
  to a character display. wxWidgets draws the same line (`wxApp` vs
  `wxAppConsole`), which is why "the wx-like contract" isn't the universal one.
- Tested: a source-level assertion that the three owning backends report `true`
  and that GTK/NWG use the default `false`, so a future backend cannot silently
  pick the wrong model.

## Feature parity measured: both backends now pass every probe row
- **The capability probe reports "every probed step works" on both GTK3 and
  pancurses, row for row.** The report previously carried a misleading row: the
  window-key callback was checked with `live()`, but a key can only be delivered
  by a running event loop and the probe starts none — so it printed `INERT` on
  *every* backend, i.e. an artefact of the probe dressed up as a capability gap.
  It now records acceptance and labels the row "fires only with a running loop".
- So the honest statement is no longer "1 non-verdict row each" but: the two
  backends are **equal on every measured step**. What still differs is not
  capability but the **loop-ownership model** — `gui_backend` hands control to
  GTK (`try_quit` reaches through the loader into `g_main_loop_quit`, `present()`
  manually pumps `g_main_context_iteration` so widgets become mappable, and
  events can arrive mid-setup) while `pnc_backend` owns its loop outright
  (`while running { getch() }`, no publishing, no pumping, no async mapping).
  That is the wx-like contract versus a console-app contract, and it is the
  part that is not a gap to close.

## A real `Entry` on pancurses: caret, focus, and change firing
- **The terminal entry now satisfies the whole `Entry` contract**, not just the
  text half. `get_position`/`set_position` were stubs ("terminal entries report
  no caret: callers keep their own"), `grab_focus` was a no-op and `has_focus`
  always returned false — so a host had to keep a parallel caret copy that could
  drift from the widget's. The widget had in fact maintained a real caret all
  along (`PcWidgetKind::Entry { cursor }`, advanced by the arrow/Backspace/
  Delete handlers); it simply was not exposed.
- Exposed via `get_entry_cursor`/`set_entry_cursor`/`entry_grab_focus`/
  `entry_has_focus`, with two contracts the naive version would have missed:
  the caret **clamps** to the buffer (a stale position from a host would
  otherwise panic on the next keystroke) and **snaps to a char boundary** (a
  mid-codepoint caret would corrupt the text on the next removal). Focus uses
  the tree's existing single `focus_id`, so `grab_focus` followed by
  `has_focus` is now consistent instead of always false.
- Because firing is now real, the change-suppression contract is load-bearing:
  `set_text_suppressing_changed` sets a flag `connect_changed` checks, and a
  test asserts **both** directions (an ordinary set fires; a suppressed one does
  not; firing resumes) — otherwise removing the firing would make the
  suppression test pass vacuously. This is what the GUI backend's 13 call sites
  (preset edit buffers, movie replay, Insert Date/Time) depend on.
- Nine new tests (six caret/focus, two through the adapter, one integration
  through the generic wrapper); the `caret_override` shim in `common.rs` is
  retained for the backends that genuinely cannot report a caret (WASM, iOS,
  macOS, Android, zork) and its comments corrected to say so.

## `Entry::connect_changed` now fires on pancurses
- **The last functional gap in the terminal canvas/entry surface is closed.**
  `set_entry_text` mutated the buffer but never invoked the handlers registered
  via `connect_changed` — a contract the GUI backends get from GTK's `changed`
  signal, and the one `gui_backend`'s formula bar depends on. It is also what
  the capability probe had been reporting as `INERT`.
- The dispatch takes each closure out, runs it, and moves it straight back
  (`Callback` is a `Box<dyn FnMut()>` and so not `Clone`), preserving the
  persistent semantics a `connect_*` name promises — a change fires on every
  edit, not just the first. Handlers may re-enter the widget tree (that is the
  point of the hook), so they run outside the state borrow.
- Three tests: one change → one call, a second change fires again, a handler
  that writes another widget does not deadlock, and an unknown id is inert.
- Measured effect: the `pnc_tree_probe` report is now **1 unavailable step**,
  matching GTK3's own count (that one row is the window-key callback, which
  needs a live event loop and is labelled a non-verdict). Earlier in this
  changelog the count was 4.

## Terminal size comes from the toolkit, not the host
- **`pnc_backend` no longer issues its own `TIOCGWINSZ`.** `rswidgets::core`
  already had `terminal_size()` (unix `ioctl`, and — unlike the inline copy —
  a Windows `GetConsoleScreenBufferInfo` path). Added
  `terminal_size_with_override()` beside it, which encodes the precedence a
  testable host needs: `CORRO_TERM_COLS`/`_ROWS`, then a real query, then
  `$COLUMNS`/80x50. `pnc_backend` now calls that in both places, deleting ~30
  lines and its `libc` use.
- This also fixes a real inconsistency: the movie-replay copy of the same
  geometry block skipped the `TIOCGWINSZ` query entirely, so a `--movie`
  replay laid itself out at 80x50 unless the env override was set — while the
  comment directly above it claimed the derivation was shared "so a movie frame
  is laid out identically to a live session". It now is.
- Scope note, because it corrects an earlier claim in this file's history:
  of the three concerns `pnc_backend` was said to need terminal-specific
  branches for, **two are already abstracted** — terminal size (above) and the
  idle marker (a generic `set_after_redraw_callback` hook, not a terminal
  concept). Only `suspend_terminal`/`resume_terminal` is genuinely
  terminal-only: `core.rs` and `common.rs` contain no `endwin`/`def_prog_mode`
  and should not, since a GUI has no terminal to restore. So the honest count
  of non-abstractable terminal concerns is **one**, not three.

## Terminal canvases are now real: draw callbacks, blitter, scrolled windows
- **`Canvas::set_draw_callback` works on pancurses.** It was a stub, and it is
  the *only* rendering path `gui_backend` has — so it was the concrete blocker
  for ever hosting the GUI tree on a terminal. Now: a `CANVAS_DRAW` table holds
  callbacks per canvas id, `paint_canvases()` runs them during redraw with a
  `PancursesDrawContext` over a `CellGrid` sized to the canvas's own rect, and
  the grid is reduced to absolutely-positioned ANSI (`CellGrid::to_ansi`) and
  flushed through the backend's existing `emit_sgr` — including its Win9x
  curses fallback. Four tests cover invocation, rect sizing, clearing, and that
  a zero-sized canvas is skipped rather than painted empty.
- `CellGrid` gained the emit path it never had (`rgb_to_palette_index`,
  `row_attr_runs`, `to_ansi`, `blit_to_window`) — previously it was produced and
  read only by tests. Colour reduces to the xterm-256 cube plus grey ramp, and
  deliberately not the 16 user-configurable system colours, so output depends on
  the model rather than the terminal theme.
- **`App::new_scrolled_window` gained its pancurses (and zork) arm.** Every GUI
  backend had one; the terminal backends fell through to
  `Err("scrolled windows not supported on this backend")`, which is where
  `gui_backend`'s construction failed first. A terminal has no scrollbar chrome,
  but it must still satisfy the portable contract so a host that nests its
  canvas in one can be built unchanged.
- Measured effect: the `pnc_tree_probe` capability report went from **4
  unavailable steps to 2**, and `new_scrolled_window` / `scrolled.set_child`
  now pass on pancurses (GTK3 reports 1 non-verdict row). The two that remain
  are `Entry::connect_changed` never firing, and the event-loop inversion —
  neither of which is a stub any more, they are genuinely different designs.

## `CellGrid` → terminal blitter (the missing half of the terminal render path)
- **`CellGrid` is now consumable, not just assertable.** It previously held
  per-cell RGB with no way to reach a screen — it was produced by
  `pancurses_draw` and read only by tests (`row_strings`). That is what made
  `Canvas::set_draw_callback` impossible to implement on a terminal, and it was
  the concrete blocker for ever dropping `pnc_backend.rs`.
- `rgb_to_palette_index` reduces the renderer's resolved RGB to xterm-256:
  the 6×6×6 cube for colours, the 24-step grey ramp for greys (the cube has no
  true greys, so mapping grey there would look tinted), and deliberately *not*
  the 16 system colours, whose RGB is user-configurable and would make output
  depend on the terminal theme rather than the model.
- `row_attr_runs` batches a row into maximal same-attribute runs, so the
  emitter switches SGR per colour change rather than per cell. A blank glyph
  with a coloured background keeps its own run — the background *is* the
  content for cell fills.
- `CellGrid::to_ansi(origin_y, origin_x)` emits absolutely-positioned SGR
  (CUP per run, reset per row) so a caller can blit into any region without
  tracking cursor state; an empty grid emits nothing rather than clearing.
  `CellGrid::blit_to_window` is the curses fallback for terminals with no ANSI
  support, taking the position callback because colour-pair tables are
  per-screen state a grid cannot own.
- Ten tests cover cube exactness at the axis values, grey-ramp use, run
  merging/splitting, blank-with-fill, absolute positioning and reset, the empty
  case, the curses fallback, and the end-to-end model → paint → ANSI pipeline.
  Writing them surfaced two real facts worth recording: `SpreadsheetModel::new`
  defaults to spreadsheet chrome (so body cell (0,0) is model (1,1), the
  convention `GridView` exists to hide), and a column whose layout width pushes
  it past the grid width is clipped **silently** by the renderer.

## Capability probe: which backends can host the GUI widget tree
- New example `rswidgets/examples/pnc_tree_probe.rs` builds the *same* widget
  tree `gui_backend.rs` builds and reports, step by step, what a backend
  provides and what is inert. It exists because "can pancurses just be another
  `gui_backend` UI?" is recurring and its answer is non-obvious: the tree almost
  builds, then fails at one factory, or builds with every callback inert.
  Reporting each step (rather than propagating the first `?`) makes the gap
  measurable.
- Current output — GTK3 vs pancurses:

  | step | GTK3 | pancurses |
  | --- | --- | --- |
  | `new_scrolled_window` | ok | **ERR: no arm in `App::new_scrolled_window`** |
  | `entry.connect_changed` fires | ok | **INERT** |
  | `canvas.set_draw_callback` | accepted | accepted but a no-op in the adapter |

  Two rows are deliberately *not* verdicts: the canvas draw callback and the
  window key callback only fire once a frame clock / event loop is running,
  which the probe does not start, and `try_quit` reports "main loop not
  running" by design. They are labelled as such in the output so they cannot be
  misread as backend deficiencies.
- Also recorded (a correction to an earlier claim in this file's own history):
  adopting `GridView` in `pnc_backend` is **not** the bounded win it looks
  like. `GridView` renders through `pancurses_draw` into a `CellGrid`, and a
  `CellGrid` holds RGB per cell with **no emit-to-screen path** — it is
  produced only for parity assertions (`row_strings`) and is never consumed
  outside tests. `pnc_backend` renders by emitting SGR from its redraw
  callback. So swapping `GridView` in would replace a working renderer with one
  that has no output, until someone writes a `CellGrid`→window blitter (colour
  quantisation to the terminal palette, attribute runs, cursor placement). That
  blitter is the real prerequisite, and it is a feature, not a refactor.
- Conclusion it documents: pancurses cannot host `gui_backend` unchanged, for
  three structural reasons — `App::new_*` is a per-target dispatch table with a
  missing pancurses arm for scrolled windows; ~54 adapter methods are no-ops,
  including the three `gui_backend`'s rendering depends on; and the event loop
  is inverted (`gui_backend` has none, `pnc_backend` is one).

## One `MenuDispatch` consumer instead of one per backend
- **The menu-result protocol now has a single implementation.** Both backends
  handled the identical nine `MenuDispatch` variants with their own ~60-line
  match, differing only in *how* a dialog/picker/prompt is shown — the decision
  (`dispatch_menu_action`) was already shared, the presentation was duplicated.
- `actions::MenuPresenter` names the five primitives a host must supply
  (`show_dialog`, `show_list_picker`, `open_prompt`, `begin_edit_with`,
  `show_keybinds`, `set_status`), and `actions::present_menu_dispatch` maps
  every variant to them in one place. `pnc_backend`'s match collapsed to a
  single call — it now has **zero** `MenuDispatch::` references.
- `show_keybinds` is a separate primitive rather than a shared body on purpose:
  a terminal lists terminal keys and GTK lists GTK's, so there is no correct
  common string. The `About`/`HelpFull` bodies *are* shared, via
  `ui_core::about_page_body`/`help_page_body`, so those cannot drift.
- The GUI keeps bespoke handling for the variants that are presentation
  *decisions* it must own (`save_as` opens a real file dialog, the pickers
  anchor to a cell, Replay switches key-routing mode); the shared presenter
  covers those whose result is pure presentation. Both paths are exercised by
  tests, and two new ones assert every variant routes to its primitive with its
  payload intact and that the pickers draw their items from the shared modules
  rather than re-listing them.
- Also: `pnc_backend`'s startup frame now derives its viewport with
  `Viewport::recompute` instead of hand-computing column widths, row labels,
  column layout and aggregate descriptors — the same four things
  `viewport::build` already produces. The startup frame can no longer drift
  from every later frame, which is exactly the bug class that duplication
  invited.

## GTK4 key press/release handling moved into the adapter
- **The host no longer knows about GTK4's double key delivery.** GTK4 emits
  BOTH `key-pressed` and `key-released` from the same `GtkEventControllerKey`;
  on some GTK4/WSLg versions the release is routed through the pressed handler,
  so an application hooking only `key-pressed` saw every key twice. The GTK
  adapter used to rely on a `capture_handled` handshake between its CAPTURE and
  BUBBLE controllers, which was not sufficient — `gui_backend.rs` compensated
  with three consecutive-keyval dedup trackers gated on `feature = "gtk4"`.
  Those heuristics could not distinguish a genuine `Right,Right` from a
  press+release, which is precisely why they had to be feature-gated.
- The adapter now registers `GtkEventControllerKey::key-released` itself
  (`swallow_key_releases_gtk4`, wired into all four GTK4 key registrations:
  window CAPTURE, window BUBBLE, entry, canvas) and returns `GDK_EVENT_STOP`
  without invoking the host, so **every** host — terminal, GTK3, GTK4, nwg,
  wasm, mobile — receives a pure press stream.
- `gtk_dynamic_loader` gains `connect_key_released` and a
  `gtk_compat_trampoline_key_released` (mirroring the key-pressed pair) so the
  release signal can be connected like any other.
- Consequently `gui_backend.rs` loses all three dedup fields
  (`last_dedup_key`, `dedup_count`, `last_keyval_dedup`), their initializers,
  the prime-the-tracker block, both skip sites, and the `gtk4` cfg cluster
  entirely: 72 → 59 cfg-gated lines, and **zero** remaining
  `feature = "gtk4"` gates.
- **The GTK4 key path is now empirically tested.** `canvas_display` gained
  `key_released_is_a_distinct_signal_from_key_pressed`, which emits
  `key-pressed` and `key-released` on one `GtkEventControllerKey` (as a physical
  key does) and asserts the release never reaches the press handler. It was
  validated by injecting the bug into the release trampoline (delivering the
  release to the press closure): the test fails with `left: 1, right: 2`. So it
  detects a regression rather than passing vacuously. This also confirms the
  premise of the fix — `key-released` really is a distinct signal, so filtering
  it is sufficient.
- `gtk_dynamic_loader` gained `EventControllerKey::emit_key`, so any test can
  drive the key pipeline without a compositor able to deliver synthetic input.
- **`canvas_display` must now run with `--test-threads=1`,** and its docs say
  why: GTK4's type system is not thread-safe. Under the default parallel
  harness two tests initialising GTK concurrently segfault *inside
  libgtk-4* — confirmed under gdb as `SIGSEGV` in `libgtk-4.so.1` ←
  `g_type_class_ref` ← `g_object_new`. This was previously mis-filed as a
  floating-ref bug; it is a threading rule, and the fix is to serialise the
  test, not to add locking we cannot enforce for arbitrary callers.
- Five tests pin the invariant: the loader trampoline delivers the keyval and
  tolerates a null closure (2), the filter helper exists and every
  `connect_key_pressed` site has a matching filter (2, as a source-level count
  assertion so a new key registration cannot silently skip it), plus a
  headless render check.

## A real Grid view widget (`rswidgets::gridview`)
- **New module `gridview`**: the data model, the renderer and the pointer
  contract assembled into one widget a host can use without touching a
  `Canvas` itself. Before this, rswidgets had the pieces but not the assembly —
  `core::Grid` is a *model* with no widget (deliberately: `core.rs` keeps
  `crate::Grid` meaning the active backend's layout container), the backend
  `Grid` widget is a packing container (`attach` only), and `spreadsheet::paint`
  renders a model through any `DrawContext`. Every host therefore re-derived the
  same glue: corro's `gui_backend` wires `canvas.on_click(handle_click)` and
  recomputes the viewport per frame, and `pnc_backend` recomputes it in three
  separate callbacks.
- `GridView` owns a `SpreadsheetModel` and exposes body coordinates
  (`set_cell(0, 0, …)`, `set_cursor`, `cell`) so a host never adds the
  header/margin offsets; `GridView::new` configures *no* chrome, while
  `set_row_counts`/`set_grid_config` add a header band or margin column when
  wanted, and `GridView::from_grid` imports a `core::Grid`.
- The pointer side is the important half: `hit_test(x, y)` maps a **pixel**
  (what every backend already produces — the terminal canvas converts its cell
  hit for us) through the renderer's own layout, and returns a typed
  `GridHit` (`Cell` / `ColumnHeader` / `RowHeader` / `Outside`) so a host can
  distinguish selecting a column from moving the cursor. `click_at` runs the
  host callbacks and moves the cursor by default; a callback that re-enters the
  grid never trips a `RefCell` double borrow.
- Because rendering and hit-testing share one layout, a click can no longer
  disagree with what was painted. Eight tests cover the mapping (gutter, header
  band, per-column widths, out-of-range), the callbacks, and rendering into both
  a `DrawContext` and a terminal cell grid.

## One mouse API across backends (pancurses and GUI)
- **The pancurses canvas now implements `on_click`, `on_click_button` and
  `on_motion`, the same three methods the GTK/NWG/gtk4 canvases expose.** They
  were no-ops on the terminal adapter, so `corro`'s GUI backend had a click
  path (`canvas.on_click(handle_click)`) while the pancurses backend had none.
  The terminal converts its cell hit into `DrawContext` pixel coordinates
  (`CHAR_W`/`ROW_H` per cell, relative to the canvas origin), so a host's
  click math is backend-independent instead of branching on the backend.
- `on_click_button` takes precedence when both handlers are registered, matching
  GTK (where the two share one signal connection). Motion is delivered with
  button 0 for a bare position report and with the held button otherwise.
- The toolkit keeps the handlers in a side table (`canvas_on_click` /
  `canvas_on_click_button` / `canvas_on_motion` / `canvas_clear_pointer_handlers`)
  rather than on the widget node, so the node type stays cloneable.
- Canvas clicks are resolved before the generic widget dispatch, so a
  canvas-painted UI hit-tests at pixel resolution exactly like GTK; a canvas
  with no handlers falls through to the built-in button/checkbox/entry/grid
  handling untouched.
- `corro`'s pancurses backend now asks for capture and uses `set_mouse_hook` to
  handle what the widget cannot know (a header-row click selects its column),
  letting the toolkit's default dispatch move the cell cursor; capture is
  released on exit so the terminal keeps its native selection afterwards.
- New tests cover the routing (button handler wins over plain, motion reaches
  `on_motion` with button 0, clicks outside the canvas are not delivered) and
  the cell→pixel conversion.

## Optional mouse support in the pancurses backend
- **The terminal backend can now take pointer input — opt in with
  `set_mouse_enabled(true)`.** Capture stays off by default, because a
  mouse-capturing terminal swallows drag events and loses its native text
  selection; the existing startup `mousemask(0, None)` is unchanged.
- Pointer events are decoded into a backend-neutral
  `MouseEvent`/`MouseAction` pair (`Pressed`, `Released`, `Clicked`,
  `DoubleClicked`, `TripleClicked`, `Moved`, `Wheel`, `Position`) with the
  Shift/Ctrl/Alt flags and the raw `bstate` preserved. One decoder serves both
  link targets, since ncurses (mouse v2) and PDCurses share the button-group
  and modifier bit layout. A PDCurses `BUTTONn_MOVED` record (which aliases the
  triple-click bit) is only read as motion when `PDC_MOUSE_MOVED` is also set,
  so drags and triple clicks stay distinguishable.
- Clicks are hit-tested against the widget tree and dispatch the same way the
  keyboard does: buttons fire, checkboxes and radio buttons toggle (radio
  selection follows the group), text fields focus, and a click on a grid maps
  through the renderer's own cell layout (header band, row-label gutter, and
  per-column widths) to move the cell cursor — committing an in-progress edit
  and firing the cursor-move callbacks so hosts stay in sync.
- Wheel notches scroll the focused grid by three rows (the PageUp/PageDown
  step), scrolling the viewport and refreshing the formula bar.
- `set_mouse_hook` lets a host observe or consume events before the built-in
  dispatch; `mouse_is_enabled` reports the opt-in without a terminal.
- The decode and hit-test logic is pure and unit-tested (mask layout, wheel,
  modifiers, motion, gutter/column/row mapping), so the renderer and the
  click mapping cannot drift apart silently. See `KEYBINDINGS.md` for the
  gesture table.

## Demo mode on the GUI backends
- **`--movie` now works on the GUI backends**, not just the ratatui terminal.
  The workbook parsing and op application are shared (`src/gui/movie.rs`), so a
  movie reads identically whichever UI replays it; only the frame-painting half
  is backend-specific (`src/gui/gui_movie.rs` for the widget backends,
  `run_pancurses_movie` for pancurses).
- `scripts/gui_movie.py` records a movie from the running window under a
  throwaway X server and encodes it with ffmpeg, so the demo video is
  reproducible from a checkout; a recording is checked in at
  `dist/corro-gui-movie.mp4`. `scripts/demo_movie.py` builds the shipped demo
  from title cards plus those replays.
- Movie pacing flags (`--movie-typing-cps`, `--movie-confirm-ms`,
  `--movie-menu-hold-ms`) now parse on the CLI for every backend rather than
  only being honored by the ratatui path.
- **A movie replay no longer writes to the log it is reading.** Every GUI
  commit path appends to `app.core.path`, and the replay left the movie's own
  file bound, so each recording appended its steps to the fixture — the demo
  workbooks grew on every run. The replay now detaches the file (as the TUI
  replayer already did) while keeping it as the display source.
- **The GUI movie animates typing, visibly.** It advanced one whole *step* per
  tick, so every value appeared instantly and then sat motionless for the entire
  hold — a long margin label like `--- Belmont ---` occupied the formula bar for
  ten seconds with nothing happening. `--movie-typing-cps` was parsed and
  published but never used by the driver (`char_delay` had no callers). The
  driver now reveals a value one character per tick, writing the growing text
  into the formula entry (as the TUI does by re-entering edit mode with the
  partial buffer each character) and applying the edit when the last character
  lands. The partial buffer is dropped at the commit so the bar shows the
  committed cell value rather than the last typed run.
- **The GUI movie moves the cursor before it types.** The cursor only moved when
  a step was applied, so the grid's edit overlay painted the growing text on the
  *previous* step's cell instead of the cell being typed. The TUI moves to the
  target address first (`movie_move_cursor_to_addr`) and then types, which is
  what puts the preview on the right cell; the driver now does the same, so the
  margin cell fills in alongside the formula bar.
- **The GUI movie opens on a blank sheet and types into the right cell.** Two
  problems, both fixed. The driver was armed before the window was mapped, so
  the first steps were applied while the window was still being built and a
  recording opened on a sheet that already had content; there is now a one
  second lead-in showing the untouched (genuinely empty) sheet. And the cursor
  was parked on the *previous* step's cell while a value was typed, because the
  cursor only moved when a step was applied — the typing is drawn on the cursor
  cell, so the value appeared to be entered somewhere else and only jumped home
  at the commit. The cursor now moves onto the target cell one tick before the
  first character. The margin-cell case needed `step_cursor` to resolve
  `SetCellRef` the same way the step itself does; hand-resolving it conflated
  the margin cell `[A1` with main `A1`.
- **Recording speed is now real.** Both capture scripts used an `xwd` +
  `convert` loop, and at 2400x820 each takes ~500ms, so the loop delivered
  about **1 fps** however high a rate was requested — a 32-second session
  produced ~32 frames and played back as 2.7 seconds. They now record with
  ffmpeg's `x11grab`, which captures continuously at the requested rate. The
  shipped demo is also half its previous speed throughout (typing, step holds,
  card durations and the collaboration timings), and is 298s rather than 60s.
- **The two-window segments keep their true shape.** They are still distorted if
  the pair is not scaled by the same factor on both axes: scaling 2400x820 to
  the video width alone left each window 2x too wide, and the 600x410 the demo
  then settled on squashed each one 2x too *narrow*. The pair is now scaled
  uniformly (2400x820 -> 1200x410, so one window is 600x410 against a true 1.50
  aspect) and centred vertically.
- **The demo replays at double speed.** The shipped video had become sluggish:
  3.25 chars/sec with an 800ms per-step hold and a 2800ms menu hold. The typing
  rate and every hold move together (`--cps 6.5`, `--confirm-ms 400`,
  `--menu-hold-ms 1400`, two-window `--tempo 1.0`, cards 3s), roughly halving
  the replay wall-clock. Capture and output both go to 16/s so the
  per-character animation stays visible at the faster pace, and a long workbook
  (`main.corro`) no longer hits the capture frame cap, which had been silently
  truncating it.
- **Concurrent editing demos.** `scripts/two_window_movie.py` records two
  front-ends editing one file (`--pair gui-gui` or `gui-tui`), showing each
  window's commit appear in the other. Both windows are scripted via
  `CORRO_EDIT_SCRIPT` and apply their edits through the ordinary commit path
  (`ui_core::edit_script_from_env`; the GUI arms it from its timer, the TUI from
  its loop), because driving them with synthetic keystrokes needs pixel
  calibration that silently writes to the wrong cell when it drifts.
- `scripts/demo_movie.py` builds the shipped demo video from title cards plus a
  replay of each featured workbook, recorded from the running window. The
  `main.corro` card notes that its `#NAME`/`#PARSE`/`#CIRC` cells are deliberate
  broken-formula fixtures rather than rendering faults. The
  typing speed is a flag (`--cps`), and the demo uses a deliberately readable
  6.5 characters/second (a quarter of the old 26) so the typed values can
  actually be followed in the recording.
- **Movie mode is now the normal UI.** `--gui --movie` opens the same window,
  builds the same widget tree and runs the same draw callbacks as an
  interactive session; a periodic timer applies one movie step per tick instead
  of waiting for a keystroke (`gui_backend::arm_movie_driver`). The window
  closes itself when the script ends, like the TUI replayer quits. The parallel
  raster renderer is gone — it had drifted from the window four times (blank
  frame, overlapping glyphs, missing margin shading, missing text) because it
  was a second implementation of rendering. `scripts/gui_movie.py` now records
  the live window under Xvfb instead of producing frames from that renderer.
- The GUI movie painter no longer draws its own copy of the sheet. It now
  calls the same body renderer the live canvas uses
  (`gui_backend::render_grid_body`), so margin shading, cell colours, the
  cursor ring, selection and gutter padlocks cannot drift from the window —
  the movie had lost the grey margin bands precisely because it was a second
  implementation. The viewport is likewise built by the shared controller.
- Fixed two frame-rendering defects in the GUI movie painter: the sheet was
  sized with a *column count* where the viewport expects a *character width*,
  which trimmed it down to the margin columns and left most of the frame blank
  (the "huge empty spaces" the window itself had before it sized its viewport
  from the live canvas); and glyphs were drawn at the 7.2px cell-layout pitch,
  which made 2x text overlap into an unreadable smear. Columns are now
  stretched to span the frame and text uses its own advance.

## Changes
- **Saving an unsaved document** no longer drops the auto‑added TOTALs. The untitled log (which Save renames verbatim onto the destination) was written as a bare header, so the in‑memory margin seeds — which have no committed edit op — vanished on reopen. It is now a complete serialization (all regions, seeds and LINKs included).
- **Windows release exes** build with the GUI feature so Explorer double‑clicks open the NWG window. The exe stays console‑subsystem (so cmd/PowerShell wait for the TUI until \`Ctrl+Q\`); a throwaway Explorer console is detected (only our process attached) and hidden/freed at startup. Previously every modern Windows exe was console‑subsystem with no GUI to fall back to, or (briefly) linked GUI‑subsystem, which made the shell return to its prompt immediately while the TUI kept drawing into the same console.
- **Windows now redirects stderr** to the debug log (like the existing Unix `dup2`), so debug/instrumentation traces can never overwrite the TUI; `%LOCALAPPDATA%\corro\debug.log` is used when no `CORRO_DEBUG_LOG`/`XDG_STATE_HOME`/`HOME` is set.
- **pancurses+ratatui+gui** now combine on Linux and Windows (triple combo builds and runs every UI). `rswidgets` resolves its default backend native‑first as documented (`pancurses` stays reachable via `backends::pancurses::init()`); the NWG adapter stack is available alongside pancurses on Windows, and corro's native GUI code names the platform adapter + common wrappers explicitly instead of the flipping root prelude. Previously the combo failed to compile (17 errors Linux, 19 Windows) and Linux `--gui` died with “loader not initialized”.
- **Android: the UI can be built and inspected without an emulator.** `cargo run --example android-ui --features gui` builds the exact widget tree the Android path roots in the Activity (menu bar from the shared menu definition, formula bar, sheet canvas, tab strip, status line) and exits before the event loop, so a layout change can be checked in seconds instead of an APK round‑trip. `src/gui/android_backend.rs` and the `gui` feature's rustdoc now describe the Android host explicitly.
- **Android resources are generated by the build.** `android/corro/build.rs` runs `rswidgets::android_generator` (the source is `#[path]`-included from the rswidgets checkout, so no `[build-dependencies]` entry — adding one breaks cargo feature unification and the host crate fails on `crate::backends::init`) and writes the Material 3 theme, palette, strings, adaptive launcher icon and manifest into the project root that contains `app/`. Additive, so hand‑edited resources are never overwritten; `build_apk.sh` passes `--features generate-android-resources`, and `RSWIDGETS_ANDROID_PROJECT` still overrides the root.

### Release artifacts for this version
- **Linux**: `gcorro.gz` (GTK GUI, runtime‑dlopen) and `corro-no-gui.gz` (terminal UI), both linked for glibc 2.17, alongside `corro.gz`.

#### WIP / not shipped
- **Windows 95 support**. The pancurses TUI backend renders and takes input on Win95 (PDCurses win32 console flavour, ANSI APIs; wait‑based input poll; key‑code mapping for non‑WIDE PDCurses), and the native NWG GUI builds for Win95 (`/SUBSYSTEM:WINDOWS`, `WinMain` entry). This is **NOT** part of the 0.7.0 release: it cannot be built on CI (rust9x toolchain + VC6 libs are local‑only) and is not gated by the test suite, so no `corro-no-gui-win95.exe.zip` / `gcorro-win95.exe.zip` is published for 0.7.0. Treat the `i586‑rust9x` targets as experimental.
- **Busybox‑style argv[0] dispatch** (`gcorro*` prefers the GUI, `pcorro*` the pancurses UI; explicit flags always win) works on the platforms built above, but the Win95 exes that motivated it are WIP.
- **Win95 runtime note** (for when it does ship): rust9x exes need `unicows.dll` next to the exe (or in `C:\WINDOWS\SYSTEM`); it is not bundled (see `README95`).

---

*Generated from the project’s changelog.*
