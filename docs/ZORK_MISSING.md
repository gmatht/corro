# zork backend — status vs. GTK / NWG / pancurses

Status: **no open parity gaps.** The enumeration this file used to carry is
done; what remains is stated below, with the reason.
Last updated: 2026-09-30, at `zork/gaps` (9c43fbc4).

Verify with:

```
cargo build -p rswidgets --no-default-features --features zork
cargo test  -p rswidgets --no-default-features --features zork
```

137 lib, 27 harness, 80 adapter, 1 doc test. Plus corro's own: 468 lib
(`cargo test --no-default-features --features zork --lib`) and 34 in
`tests/zork_feature_parity.rs`. `.github/workflows/backends.yml` builds
and tests all of these, and `--examples`, on every push that touches the
crate.

The `zork` backend (`rustxWidgets/rswidgets/src/backends_zork_adapter.rs`
+ `src/backends/zork/`) is a headless, in-memory widget model with a typed
test harness and a text REPL. It is a first-class backend: `core.rs` and
`common.rs` already have `#[cfg(feature = "zork")]` arms for it everywhere.

## How "no open gaps" was measured

Method-level gaps are produced by mechanically diffing the public inherent
method surface of the zork adapter against `backends_gtk_adapter_impl.rs`,
`backends_nwg_adapter.rs` and `backends_pancurses_adapter.rs`, then
subtracting what `common.rs` / `core.rs` actually require. The script that
produces the numbers is in the repo, so the claim is re-checkable rather
than asserted:

```
$ python3 scripts/surface_diff.py
gtk: 181 methods, 126 free functions
nwg: 150 methods, 119 free functions
panc: 162 methods, 150 free functions
zork: 296 methods, 232 free functions

methods zork lacks (3):
  DropDown.diagnostics                 (gtk)
  DropDown.has_size_request_symbol     (gtk)
  DropDown.is_gtk4                     (gtk)

free functions zork lacks (7):
  create_dialog_button   (nwg)
  diagnostics             (gtk)      -- version probes, below
  has_size_request_symbol (gtk)
  is_gtk4                 (gtk)
  new                     (gtk)      -- a `Box::new` false positive
  pump_main_context       (gtk)      -- [-], the REPL *is* the loop
  quit_main_loop          (gtk, nwg) -- [-], likewise
```

Those three methods and three of the seven free functions are GTK's
*version probes* — "am I GTK4?", "does `gtk_widget_set_size_request`
exist?" — and answering them on a backend that is neither would be a
fabrication, not a feature. `create_dialog_button` is NWG-only and is
about a Win32 dialog control, not a model operation. `new` is the script
seeing a `Box::new` in GTK's source, not a definition.

After those, the *method* surface has no gap at all, and the free-function
surface has none that is not either a version probe or a documented `[-]`.
Everything the old file listed is now implemented; the sections below
record how each was closed and what it cost.

zork has *more* methods than GTK because a headless model has state GTK
holds in the compositor (the redraw cascade, the layout callback, the
virtual clock) and needs somewhere to put it. Those 67 zork-only methods
are not a compatibility surface — see §5.

---

## 1. What was closed in the `zork/gaps` pass

Six commits. Each is independently revertible; the order is the order the
work happened in.

### 1.1 The model had nowhere to record any of it — `75d9d455`

Every remaining gap was a gap in the **model** first. A cursor, an edit
buffer, a column layout, a tab strip, a timer deadline, an application's
action set: there was no field for any of them, so every adapter method
could only discard its argument. `model.rs` grows four things:

| Type | What it holds | Reached through |
|------|---------------|-----------------|
| `SheetChrome` | cursor, anchor, edit buffer, grid config, row counts, column layout, row labels, tabs, menu/status text, formula bar | `sheet_set_*` |
| `SheetInput` | the four spreadsheet input handlers | `sheet_add_*_callback` |
| `Timer` + a **virtual** clock | deadlines, intervals, one-shot vs repeating | `timeout_add_*` / `advance_timers` |
| `app_actions` / `layout_cbs` / `key_callbacks` | the action registry, the layout callback, the process-wide key table | `application_add_action` etc. |

`SheetChrome` defaults to "nothing configured", so a spreadsheet created
and immediately snapshotted is byte-identical to one that was never
touched — the same discipline `ZorkProps` already followed, and the reason
a snapshot diff can distinguish "set to the default" from "never set".

`ZorkKind::name()` is now the single place a kind's name is spelled. It had
grown two copies of the same 20-arm `match` (the snapshot's `kind` string
and the tree dump's prefix), which is exactly the pair of views that quietly
disagree — the failure mode of S27 below.

### 1.2 `draw_rgba_image` — `b24abf9a`

S26 said this was unimplemented on six of eight backends. That was true,
and it was worst on the **recorder**: a recorder's job is to record, and
refusing a blit it could perfectly well record meant a draw callback that
drew an image left *no trace at all* — so neither a test nor a driver could
tell whether it happened. `headless::DrawOp::Image` now carries the
destination, the source extent, the `scale` and the samples verbatim, and
`RecordingDrawContext::draw_rgba_image` records it and returns `true`.

The one refusal kept is a `pixels` buffer too short for
`width * height * 4`: there the op could not be recorded faithfully, and
returning `true` is precisely the "promised a draw that never happened"
answer the trait's contract exists to prevent. A zero-sized image is a
valid blit and is recorded.

`tests/zork_feature_parity.rs` had a test asserting the old `false`. It is
replaced by two — one driving a blit through a real canvas draw callback end
to end, one pinning the short-buffer refusal — so both halves of the
contract are asserted rather than only the half that was easy.

### 1.3 The REPL — `8bab5aa1`

Three defects, all found by *running* it, none by reading it.

**R6 — `type <n>`.** The verb could only write to the Entry you were
standing on, so driving a form with two entries meant walking to each in
turn. It now takes an optional target: `type 3 hello` writes straight into
item 3; `type 3` with no text still opens the sub-prompt.

Doing that exposed a trap. `type` read its first word as a target
unconditionally, so `type hello there` on a non-Entry answered *there is
no "hello" here* — a complaint about the reader's text that hid the actual
problem. A word is now read as a target only when it resolves to something;
a bare number or direction that names nothing is still reported as a
targeting mistake, because there it really is one.

**The number on screen was not the number `go` accepted.** The "You see:"
list is printed in one order (directions, then the parent's other children,
then this node's own) and `resolve_number_target` walked a second,
differently-ordered list (all siblings, then all children). The reader was
shown `3. Entry` and `3` resolved to something else. Four computations of
the same thing — the direction lines, the numbered list, `navigate_find`,
and the resolver — are now `navigate_dirs` plus `visible_from`, with the
display and the resolver both derived from it. This is the same class of
bug as S27, and fixing it is why `type <n>` is usable at all.

**`dump` wrote an empty file.** `ZorkApp::adopt_singleton` *moves* the model
out of the thread-local and leaves an empty one behind, so the verb went
through the facade's singleton and dumped nothing while reporting "0 nodes"
for a session that had just built a hundred widgets. `debug_dump_tree_of`
takes the state directly, and the entry points return `io::Result` so
"could not write" is distinguishable from "wrote an empty tree".

Two more, found the same way: a dump line ran `#2Book1` together (the
label is now always quoted, even when empty), and `sheet` refused because
you were standing on the Window rather than the Spreadsheet — which is
where every session starts.

New verbs, or the new surface is reachable only from Rust: `sheet`
(`show`/`cursor`/`cell`/`edit`/`commit`/`stop-edit`), `timers`, and `wait`.
Each prompt also advances the virtual clock by 50 ms, so a timer a host
installed actually fires instead of silently never coming due.

### 1.4 Harness — `6c5ef1cb`

The harness is the *recommended* way to test against zork, so leaving the
new surface reachable only from the adapter and the REPL would have made it
untestable without a thread-local. `Harness` gains the chrome, the input
hooks, the timers, the application registry, the key table, the layout
callback, the redraw cascade and the tree dump; `Spreadsheet` and a new
`Application` handle gain their own methods.

### 1.5 `Canvas::on_gesture` — `9c43fbc4`

The last real method GTK and NWG had. One callback for hover, drag, buttons
and scroll, which `core::DragTracker` turns into a drag without each caller
reassembling one from three hooks and working out the click-versus-drag
threshold.

The gestures are synthesised from the `pointer_*` entry points the raw hooks
already use, so a host that registers both gets both from one event. That
needed one piece of new state, `PointerHooks::down_button`, because a
`Gesture::Drag` has to report *which* button it belongs to and the raw
`on_motion` does not carry one.

---

## 2. Defects found while doing the above

Four, all of which would have shipped. Recorded because the *kind* of each
is the argument for the test that found it.

| # | Defect | Symptom | Found by |
|---|--------|---------|----------|
| Z4 | A timer handler that touched a widget panicked `RefCell already borrowed` | `advance_timers` ran each handler while the facade held the singleton's `RefCell`. A timer body touching a widget is the *normal* case — a poll that redraws, a tick that activates an action — not an exotic one. Same defect class as the `with_state` bug that killed the first corro run (Z1), in the one place written since. | `timers::a_timer_handler_may_touch_the_model` |
| Z5 | Advancing 200 ms across a 50 ms timer fired it **once**, not four | The re-arm was measured from the end of the jump, silently skipping every interval in between. | `timers::a_long_jump_fires_a_repeating_timer_once_per_interval` |
| Z6 | The REPL's printed number was not the number it resolved | See §1.3. | `type 3` on a session that printed `3. Entry` |
| Z7 | `dump` reported "0 nodes" for a 100-widget session | The REPL's `adopt_singleton` leaves the thread-local empty. | running it |

Z5 **inverted** a policy the first version documented, so the reasoning is
in the code: "catch up, do not replay" is right for a debounced poll and
wrong for a main loop reaching successive timeouts, and the latter is what
a model of `g_timeout_add` has to be.

## 3. Contracts settled, not just implemented

These are the ones that are easy to get subtly wrong, and are therefore
asserted rather than left to a comment:

* `set_cursor` **clamps to the sheet's known extent** — and leaves the value
  alone when the extent is unknown (`main_rows == 0`), so a spreadsheet
  driven before its geometry is configured still remembers where it was put.
  A cursor one row past the last row is a state the widget can never be in.
* `set_editing(true, "", _)` **seeds the buffer from the cursor cell**, not
  from empty. An empty seed would mean F2-then-commit silently blanks the
  cell — the failure mode a caller would only find by losing data.
* `set_tab_data` **clamps the active index** to the list, for the same
  reason `set_dropdown_selected` drops an out-of-range selection.
* A cursor-move handler **sees the clamped position**, because the clamp and
  the notification are two borrows and the notification runs second.
* A commit handler **sees a consistent model**: the mutation lands under one
  borrow, the notification under a second, released one.
* `draw_rgba_image` **refuses a short buffer** and records nothing — the
  only refusal, and the one the contract is written to prevent.

---

## 4. The pre-existing enumeration, resolved

Kept so the claims can be checked against the code rather than taken on
faith. `model.rs` is 3889 lines, `harness.rs` 1229, `repl.rs` 1401,
`facade.rs` 1054, the adapter 2193.

| # | Item | State |
|---|------|-------|
| B1–B6 | §0: the `zork` feature did not compile (38 errors) | `[x]` fixed in 701f961b, and `.github/workflows/backends.yml` exists so nothing rots again |
| M1–M6 | §1: missing model node kinds | `[x]` `ZorkKind` has 20 variants; `Canvas`/`Overlay`/`ScrolledWindow`/`Fixed`/`Application`/`Spreadsheet` all carry real state |
| §2.1 | `common_types_mod!` requirements | `[x]` |
| §2.2 | `common_types` cfg-tail methods | `[x]` |
| §2.3 P1–P34 | cross-backend parity methods | `[x]` every one; see §1 for the two that were not in the list at all (`on_gesture`, the redraw cascade) |
| §2.4 F1 | `create_fixed` | `[x]` |
| §2.4 F2 | `create_application`, `add_action` | `[x]` `75d9d455` |
| §2.4 F3 | `create_spreadsheet` + 12 `spreadsheet_set_*` | `[x]` `75d9d455` (see F9) |
| §2.4 F4 | `timeout_add_once`, `timeout_add_repeating` | `[x]` `75d9d455`; found Z4, Z5 |
| §2.4 F5 | `pump_main_context`, `quit_main_loop` | `[-]` zork's REPL **is** the loop (`owns_event_loop() == true`); there is no loop pointer to publish. Same answer as the old file. The `quit` path itself is real: `ZorkState::quit` sets `running = false` and the REPL polls it. |
| §2.4 F6 | `open_file_filtered` / `save_file_filtered` | `[-]` unchanged, and unchanged in reasoning: a headless backend has no chooser, so the honest answer is "cancelled". The REPL is text; a real filter picker is a terminal-dialog project. |
| §2.4 F7 | `set_window_pos`, `mark95a`, `translate_vk`, … | `[-]` Win32/Win95 only |
| §2.4 F8 | `debug_dump_native_tree`, `window_queue_redraw_cascades_to_children` | `[x]` `75d9d455`; the REPL's `dump` found Z7 |
| §2.4 F9 | 15 `set_*` spreadsheet calls | `[x]` `75d9d455` — `SheetChrome` plus one setter per field |
| §2.4 F10 | `add_key_callback`, `add_goto_callback`, `add_cursor_move_callback`, `add_commit_edit_callback` | `[x]` `75d9d455` — `SheetInput`, each fired with the borrow released |
| S1–S5 | no geometry / visibility / expansion / margins / classes | `[x]` `ZorkProps` (pre-existing) |
| S6–S11 | focus, caret, editability, scroll | `[x]` |
| S12 | no draw callback / draw surface | `[x]` |
| S13 | no pointer/key routing | `[x]`, and extended with the `Gesture` stream in `9c43fbc4` |
| S14–S16 | lossy menu model, selection does not dispatch, radio groups | `[x]` |
| S17 | dialog responses absent | `[x]` |
| S18 | no `Clone` on 5 adapter types | `[x]` |
| S19 | `MenuBar::insert_action_group` | `[x]` |
| S20 | radio group lost in `core.rs` | `[x]` |
| S21 | no event loop / quit plumbing | `[x]` for the model side (`quit` sets `running`); `[-]` for `present()`, which has nothing to present to |
| S22 | `owns_event_loop` parity | `[-]` correct as-is |
| S23 | snapshot too thin | `[x]` — it now carries `props`, and also `chrome`, `sheet_hooks`, `timers`, `now_ms` and `app_actions` |
| S24 | `set_child` destructive-but-buggy | `[x]` `append_child` is a genuine append; re-parenting cannot duplicate |
| S25 | no pass-through | `[x]` |
| S26 | `draw_rgba_image` | `[x]` `b24abf9a` |
| S27 | submenu nodes duplicated | `[x]` 59e3a0de; and §1.3 removed the second, worse instance of the same bug in the numbered list |
| H1–H9 | harness gaps | `[x]` (pre-existing) and extended in `6c5ef1cb` |
| R1–R5 | REPL gaps | `[x]` (pre-existing) |
| R6 | `type` needs a target | `[x]` `8bab5aa1`; found Z6 on the way |

### Explicitly out of scope (`[-]`)

* Win32/Win95 shims (F7), `pump_main_context` / `quit_main_loop` (F5 — the
  REPL is the loop), and the four file-dialog functions (F6 — no chooser
  exists to open). Each of these has a *reason* rather than a `todo`, and
  the reason is in the code at the site.
* The `--features wasm` **without** `--no-default-features` build: still
  18 errors, and still not a wasm problem (the default `gtk` feature cfgs
  out the `crate::common::*` types the wasm adapter names). The supported
  invocations are clean at `-D warnings`; see `docs/WASM_MISSING.md`.
* **`cargo test -p rswidgets --no-default-features --features pancurses`
  does not compile its two test binaries** — 22 errors, all of them
  `core::InputEvent` and the `key::*` constants the tests import and a
  non-pancurses feature set does not provide. Verified byte-identical on
  `main` (the sorted error lists `comm` to nothing present on only one
  side), so it is pre-existing and untouched here. `backends.yml`'s
  pancurses job builds the library and corro and runs corro's `--lib`
  tests, which do pass; it does not claim the rswidgets test binaries
  build.
* CSS-vs-`pdcurses-sys` feature conflicts: a pre-existing, unrelated
  workspace resolution failure (`rustxWidgets/Cargo.lock` has a registry
  `pdcurses-sys 0.7.1` while the repo-root `[patch.crates-io]` points
  `pdcurses-sys` at `vendor/pdcurses-sys`). Build zork from the
  **repo-root** workspace, which has a correct lock file.

---

## 5. What is still not true

Stated plainly, because a document that only lists wins is a document you
cannot plan from.

* **corro on zork is a model driver, not an interactive spreadsheet.** There
  is no display, so no key routing to the grid, no clicking cells, no
  scrolling. The text REPL explores and edits the widget tree, and `sheet`
  now drives the grid's cursor, cells and edit state — which is a great deal
  more than it could do before, and still not a terminal front-end.
* **`ZorkState` is single-threaded and in-memory.** No persistence, no
  concurrency, and the harness and REPL each take a whole `&mut` state. A
  model test cannot exercise two threads touching widgets, because there is
  nothing there to exercise.
* **The virtual clock is not a clock.** Timers fire when a driver advances
  it. A host that never advances it registers a timer that never fires —
  the same situation as registering one and leaving the main loop before it
  elapses, but with no error to tell you.
* **The four GTK version probes stay unanswered**, and should: `is_gtk4` on
  a backend that is not GTK would be a fabrication.
* **67 of 296 zork methods are zork-only.** They exist because a headless
  model needs somewhere to put state GTK keeps in the compositor. They are
  not a compatibility surface, and nothing outside the backend should need
  them — the shared code is supposed to go through `common.rs`.
