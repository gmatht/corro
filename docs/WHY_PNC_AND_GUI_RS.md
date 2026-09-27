# Why `pnc_backend.rs` and `gui_backend.rs` both exist

Status: **descriptive** — this documents the current architecture and the
reasoning that produced it. It records why the two backends have not been
merged, what was measured to decide that, and which parts *were* merged.

See also: `Principles.txt` (project invariants), `../CHANGELOG.md` (the changes
described here), `rustxWidgets/rswidgets/examples/pnc_tree_probe.rs` (the
measurement tool).

## 1. The question

`src/gui/` contains two large UI modules:

| module | lines | hosts |
| --- | --- | --- |
| `gui_backend.rs` | 6531 | GTK3, GTK4, NWG (Win32), WASM, Android, iOS, macOS |
| `pnc_backend.rs` | 1512 | pancurses (terminal) |

The obvious question is why the terminal is not "just another UI supported by
`gui_backend.rs`", the way GTK and NWG are. This document answers that, and
corrects several wrong answers given along the way.

## 2. Short answer

They are not two implementations of one thing. They are **two hosts for one
application**, differing in *who dispatches events*:

* `gui_backend.rs` is the shared host for backends where **a framework owns the
  event loop** and calls back into the widgets.
* `pnc_backend.rs` is the host for backends where **the backend is the loop** —
  it reads input directly.

Everything above that — what the app *decides* — is shared. Everything below it
— how a decision is *presented* — is not, and cannot be.

## 3. What is already shared

The application logic lives in backend-free modules, and both hosts call it:

| shared layer | what it owns | `pnc_backend` | `gui_backend` |
| --- | --- | --- | --- |
| `gui/viewport.rs` | visible rows/cols, column widths, row labels, agg funcs | 8 uses | 1 |
| `gui/render.rs` (`fill_cells`, `CellSink`) | cell text/styles/cursor into any sink | 6 | 2 |
| `gui/compute.rs` | per-cell display style and aggregates | 7 | 8 |
| `ui_core.rs` | layout constants, formatting, alignment | 16 | 24 |
| `gui/actions.rs` (`MenuPresenter`, `dispatch_menu_action`) | what a menu action *means* | 4 | 4 |

The `MenuDispatch` consumer was the largest duplicated block (a ~60-line match
in each backend over the same nine variants). It is now one implementation:
`pnc_backend` has **zero** remaining `MenuDispatch::` references and implements
only the five presentation primitives of `actions::MenuPresenter`.

## 4. What is genuinely different

### 4.1 Loop ownership — the irreducible difference

Both backends implement `BackendApp::run()`, so the *entry point* is uniform.
What differs is dispatch. `BackendApp::owns_event_loop()` makes this queryable:

| backend | `owns_event_loop()` | what `run()` does |
| --- | --- | --- |
| GTK3 / GTK4 | `false` | `g_main_loop_run` — GTK dispatches |
| NWG | `false` | `GetMessageW` — Win32 dispatches |
| WASM | `false` | returns immediately — the browser dispatches |
| Android / iOS / macOS | `false` | returns immediately — Activity / UIKit / AppKit dispatches |
| **pancurses** | **`true`** | blocks in `getch()` — it *is* the loop |
| zork | `true` | REPL read |
| ratatui | `true` | crossterm event loop |

Where a framework owns the loop, `rswidgets` must also *reach into* it to quit:

* a `MAIN_LOOP` atomic in `backends/gtk.rs`, **and**
* `loader.main_loop: Mutex<usize>` in the loader,

both existing purely so `quit_main_loop()` can signal a loop the caller does not
own. `gui_backend` never calls `g_main_loop_run`; it calls `rxapp.try_quit()`
(line 339), which is the wxWidgets pattern — *the toolkit owns the loop, the app
only signals it*.

There is a recorded bug from this design: a loop that has `quit_main_loop`
called on it *before* `run()` makes `g_main_loop_run` return immediately, so the
app produces `WINDOW_DRAWN=true` with no output and no keystroke dispatch.
`backends/gtk.rs` now always creates a fresh loop for this reason.

A second consequence is `present()`, which is not a simple show. It calls
`gtk_widget_show_all` + `gtk_window_present`, then **manually pumps
`g_main_context_iteration`** (with `g_main_context_pending` and wall-clock
budgets) because on GTK4/Wayland the window and children are not mapped until
the main loop processes the configure event — and `grab_focus()` silently fails
until then. `gui_backend` therefore has code written around "a loop may be
running underneath me": clicks can arrive mid-setup, and there is a
`StartupEdit` decision made from state observed *after* `present()` returns.

None of that has an analogue on a terminal. A curses widget is fully specified
the moment it is created: no unmapped state, no async acknowledgement, no
focus that silently fails.

### 4.2 Presentation

A terminal has an SGR box, a list popup drawn over the grid, and a formula-bar
text prompt. GTK has modal windows, native menus popped at root coordinates,
and a real `Entry` with a caret. `MenuPresenter` names the five primitives
(`show_dialog`, `show_list_picker`, `open_prompt`, `begin_edit_with`,
`show_keybinds`, `set_status`) so the *decision* is shared while the *widget* is
not.

`show_keybinds` is deliberately its own primitive rather than a shared body: a
terminal lists terminal keys, GTK lists GTK's. A common string would be wrong
for one of them. The `About`/`Help` bodies *are* shared, via
`ui_core::about_page_body`/`help_page_body`, so those cannot drift apart.

### 4.3 Terminal-only operations

`suspend_terminal()` / `resume_terminal()` (`def_prog_mode`/`endwin`) exist so
the external editor can take the real screen. `core.rs` and `common.rs` contain
no equivalent and should not: a GUI has no terminal to restore, and a portable
no-op would be an API that lies about what it does.

That is the complete list of terminal-only *operations*: one pair, ~10 lines.

Two things previously believed to be terminal-only turned out **not** to be:

* **Terminal size.** `rswidgets::core::terminal_size()` already existed (unix
  `ioctl(TIOCGWINSZ)` plus a Windows `GetConsoleScreenBufferInfo` path — a
  superset of the inline copy in `pnc_backend`, which lacked the Windows half).
  `pnc_backend` now calls `terminal_size_with_override()`, which encodes the
  precedence a testable host needs: `CORRO_TERM_COLS`/`_ROWS` → real query →
  `$COLUMNS`/80×50.
* **The idle marker.** `set_after_redraw_callback` is a generic after-frame
  hook, not a terminal concept.

## 5. The measurement

`rustxWidgets/rswidgets/examples/pnc_tree_probe.rs` builds **the same widget tree
`gui_backend` builds** and reports, per step, what a backend provides. It exists
because arguments about this kept being wrong; a number is better.

It reports identical results on GTK3 and pancurses — every row passes:

```
  new_window                             ok
  new_box(vertical)                      ok
  new_box(horizontal)                    ok
  new_label                              ok
  new_entry                              ok
  entry.set_hexpand                      ok
  entry.connect_changed                  ok
  entry changed callback                 ok   (fired 1x)
  entry.get_text after set_text          ok   (round-trips)
  new_canvas                             ok
  canvas.set_can_focus                   ok
  canvas.set_draw_callback               ok
  new_scrolled_window                    ok
  scrolled.set_child                     ok
  window key callback accepted           ok   (fires only with a running loop)
  try_quit exists (needs a running loop) ok
```

Getting there required closing four real gaps, all found by the probe rather
than by reasoning:

1. `Canvas::set_draw_callback` was a no-op on pancurses — and it is the *only*
   rendering path `gui_backend` has.
2. `CellGrid` had no emit-to-screen path (`to_ansi`, `blit_to_window` added),
   so there was nothing for a terminal draw callback to paint into.
3. `App::new_scrolled_window()` had no pancurses arm and fell through to
   `Err("scrolled windows not supported on this backend")` — where
   `gui_backend`'s construction failed first.
4. `Entry::connect_changed` was registered but never fired, and
   `get_position`/`set_position`/`has_focus` were stubs, so a host had to keep a
   parallel caret it could not verify.

Two probe rows are deliberately *not* verdicts and are labelled as such: the
canvas draw and window key callbacks only fire under a running frame clock or
event loop, which the probe does not start, and `try_quit` reports "main loop
not running" by design. An earlier version of the probe printed `INERT` for the
key callback on *every* backend — an artefact dressed up as a capability gap.

## 6. What would it take to delete `pnc_backend.rs`?

Not "more abstraction". The remaining work would be:

1. **Done:** terminal draw callbacks, the `CellGrid` blitter, the scrolled-window
   arm, a real `Entry`.
2. **Then:** give `gui_backend` a terminal-aware mode for `TIOCGWINSZ`
   (already portable — it just needs no work) and `suspend_terminal` (~10 lines).
3. **And:** invert the loop — either restructure `gui_backend` so it does not
   own the loop for *any* backend, or give pancurses a framework-shaped loop it
   has no use for.

Step 3 is where "delete the file" stops being a refactor. Forcing the delegating
contract on a terminal means inventing a loop pointer whose only possible caller
is the code already running, a `present()`-style context pump for widgets that
have no unmapped state, and a frame-clock hop that adds pure latency to a
character display.

wxWidgets draws the same line itself: `wxApp` (GUI, framework-owned loop) versus
`wxAppConsole` (console, own loop). The "wx-like contract" is therefore not the
universal one — it is the contract for backends that have a framework to
delegate to.

## 6a. One duplication deliberately left in place, and why

`gui_backend`'s private cell sink (`GuiCanvasSink`) was deleted — it and
`pnc_backend`'s `SpreadsheetSink` were the same adapter written twice, differing
only in where they stored cells, and `render::GridSink` now serves the GUI.
`pnc_backend`'s copy (24 lines) is **deliberately retained**.

The obvious-looking follow-up is to make the terminal widget store a
`SpreadsheetModel` too, so both hosts use `GridSink`. That was investigated and
rejected. The reasoning is worth recording because the measurements all pointed
the *other* way:

- the terminal widget's display grid is `rswidgets::core::Grid`, and
  `pancurses.rs` opens it with `PcWidgetKind::Spreadsheet { grid, .. }`
  **136 times** (~19,600 characters, median line 134 chars) for only ~1.5 field
  accesses per destructure — a real defect worth fixing;
- three of its fields (`top_row`, `left_col`, `col_width`) have no writer in
  `pancurses.rs` besides an initialiser, so they *look* dead.

They are not dead. They are fields of a **public toolkit type**
(`rswidgets::core` is `pub` and every field is `pub`), and they are a scroll
origin and a default column width — API any consumer of a grid widget may
reasonably use. They look unused only because corro drives scrolling through
`column_layout` and `col_ixs`. Pruning public API on the evidence of one
consumer's usage is not a cleanup.

`SpreadsheetModel`, meanwhile, is the **renderer's** model and carries corro's
chrome (`formula_bar_address`, `border_title`, `tab_titles`, `menu_text`).
Retyping the widget's field to it would make the generic toolkit depend on the
host's rendering model — a permanent layering inversion — and would drag
host-specific concepts into the shared widget layer.

**The defect is real; the proposed fix targets the wrong layer.** The 136
destructures exist because the display state is an enum payload, which is
fixable *inside* `pancurses.rs` by giving that state a handle or accessor. That
achieves the same win with no public-API change and no layering cost. If the
boilerplate is to be attacked, that is the change to make — not the retype.

There is also a trap for anyone grepping this area: **two unrelated types are
named `Grid`** — `rswidgets::core::Grid` (the toolkit's display grid) and
`corro::grid::Grid` (the workbook's source of truth, wrapped by `GridBox`).
Searches that span both will silently mix them, which is how the "dead fields"
conclusion was first reached.

## 7. Consequences for contributors

* **Adding a widget capability?** Add it to the portable surface
  (`common.rs` / `core.rs` / `BackendApp`) with a real implementation per
  backend, and make the probe report it. Do not add a `cfg` branch to a host.
* **Adding a menu action?** Add a `MenuDispatch` variant in `actions.rs`, map it
  once in `present_menu_dispatch`, and implement the primitive it needs in each
  `MenuPresenter`. `pnc_backend` should not grow another `MenuDispatch` match.
* **Adding an event-loop behaviour?** Check `owns_event_loop()` first. If it is
  `false`, the framework is in charge and you must signal rather than block.
* **Tempted to add `cfg(target_os)` to `gui_backend.rs`?** Check whether the
  capability belongs in an adapter first. `gui_backend.rs` currently has 59
  cfg-gated lines out of 6531 (0.9%) — that ratio is the health metric.

## 8. Summary

| | |
| --- | --- |
| Shared | the application: viewport, rendering, computation, formatting, menu decisions |
| Not shared, and should not be | presentation widgets, and event-loop ownership |
| Measured | both backends pass every capability probe row |
| Remaining `pnc_backend` unique surface | ~10 lines (`suspend_terminal`), plus its own loop |
| Why the loop is not unified | a terminal has nothing to delegate to; wx draws the same line |
