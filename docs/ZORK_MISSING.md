# Zork backend — missing features vs. GTK / NWG (and pancurses)

Status: **enumeration complete**, implementation in progress.
Last updated: 2026-09-28.

The `zork` backend (`rustxWidgets/rswidgets/src/backends_zork_adapter.rs` +
`src/backends/zork/`) is a headless, in-memory widget model with a typed test
harness and a text REPL. It is a first-class backend: `core.rs` and
`common.rs` already have `#[cfg(feature = "zork")]` arms for it everywhere.

This document enumerates **every** gap found, ranked by severity, and tracks
which are implemented. Method-level gaps were produced by mechanically diffing
the public inherent-method surface of the zork adapter against
`backends_gtk_adapter_impl.rs`, `backends_nwg_adapter.rs` and
`backends_pancurses_adapter.rs`, then subtracting the surface that
`common.rs` / `core.rs` actually require.

Legend: `[ ]` todo · `[~]` in progress · `[x]` done · `[-]` deliberately not
applicable to a headless backend (with a reason).

---

## 0. Blocker: the `zork` feature does not compile

`cargo build -p rswidgets --no-default-features --features zork` fails with
**38 errors**. The backend has been dead code: nothing in CI, `run.sh`,
`build95*.sh` or `.github/workflows/` ever enables `zork`, so the breakage was
invisible. Consequences:

| # | Error class | Where | Why |
|---|-------------|-------|-----|
| B1 | `E0425` missing fns | `backends_zork_adapter.rs:732,736,740` | `create_canvas` / `create_overlay` / `create_scrolled_window` call `crate::backends::zork::*`, but the model has no such node kinds. `Canvas`/`Overlay`/`ScrolledWindow` are **empty stubs with no node**. |
| B2 | `E0119` conflicting `Clone` | `backends_zork_adapter.rs:701` | `ScrolledWindow` has `#[derive(Clone)]` *and* a hand-written `impl Clone`. |
| B3 | `E0308` `expected Result<_, Error>, found ()` | `core.rs:1665,1704,1764,1805,1845,1943,1983,2024,2088` | The 9 `#[cfg(feature = "zork")]` `new_*` arms in `core.rs` were never written — each `new_window`/`new_box`/`new_label`/`new_entry`/`new_canvas`/`new_menu`/`new_simple_action`/`new_menubar`/`new_dialog` has no zork block, so under `--features zork` the function body evaluates to `()`. |
| B4 | `E0277` missing `Clone` | `common.rs:83,102,104,106,108` | `common_types_mod!` derives `Clone` on `WidgetBox`/`Menu`/`SimpleAction`/`MenuBar`/`Dialog`; zork's `BoxWidget`, `Menu`, `SimpleAction`, `MenuBar`, `Dialog` do not implement it. |
| B5 | `E0599` missing methods | `common.rs` (18 sites) | `common_types_mod!` unconditionally requires a set of methods zork's adapter lacks — see §2. |
| B6 | dead feature gate | `common.rs:344` | `mod common_types { common_types_mod!(); }` for zork omits the `Canvas::on_key` / `ScrolledWindow::*` impls every other backend cfg-block adds, so even after B1–B5 the wrapper types lack `Canvas::on_key` and `ScrolledWindow::set_child/set_policy/set_vexpand/scroll_to/on_scroll`. |

Everything in §1–§3 is downstream of B1–B6: they cannot be implemented until
the feature builds.

---

## 1. Missing model node kinds (`ZorkKind`)

The zork `ZorkKind` enum covers 14 kinds. GTK/NWG/pancurses expose 20. Six
container/widget kinds exist as adapter stubs but have **no** model node, so
every method on them is a no-op that silently discards information.

| # | Missing kind | Adapter type exists? | GTK | NWG | panc |
|---|--------------|---------------------|-----|-----|------|
| M1 | `Canvas` | yes (stub, id never allocated) | ✅ | ✅ | ✅ |
| M2 | `Overlay` | yes (stub) | ✅ | ✅ | ✅ |
| M3 | `ScrolledWindow` | yes (stub) | ✅ | ✅ | ✅ |
| M4 | `Fixed` | no | ✅ | — | — |
| M5 | `Application` | no | ✅ | — | — |
| M6 | `Spreadsheet` | no | ✅ | — | ✅ |

M1–M3 are **required to compile** (B1). M4–M6 are `core.rs`-reachable on
GTK/pancurses only; lower priority.

## 2. Missing adapter methods, by group

### 2.1 `common_types_mod!` requirements (build blockers — B5)

These are called unconditionally from `common.rs`, so zork must provide them:

| Type | Required methods | zork has |
|------|------------------|----------|
| `Window` | `set_child_box` | ❌ |
| `BoxWidget` | `set_child_hexpand`, `set_child_vexpand`, `set_hexpand` | ❌ (also needs `Clone`) |
| `Label` | `set_fixed_width`, `set_margin_start`, `set_margin_top`, `raw_handle` | ❌ |
| `Entry` | `set_hexpand`, `set_vexpand` | ❌ |
| `SimpleAction` | `connect_activate` | ❌ |
| `MenuBar` | `activate_submenu_by_mnemonic`, `popup_submenu_by_mnemonic_at`, `activate_submenu_item_by_mnemonic`, `insert_action_group`, `handle_mnemonic_key`, `handle_menu_key`, `menu_active`, `menu_close` | ❌ |
| `Dialog` | `set_transient_for` | ❌ |
| `BoxWidget`/`Menu`/`SimpleAction`/`MenuBar`/`Dialog` | `Clone` | ❌ |

### 2.2 `common_types` cfg-tail methods (missing on zork — B6)

| Type | Methods present on GTK/NWG/wasm/android/ios/macos, absent on zork |
|------|------------------------------------------------------------------|
| `Canvas` | `on_key(Box<dyn FnMut(u32) -> bool>)` |
| `ScrolledWindow` | `set_child`, `set_policy`, `set_vexpand`, `scroll_to(hval,hupper,hpage,vval,vupper,vpage)`, `on_scroll(Box<dyn FnMut(bool,f64)>)` |
| `TextView` | `set_hexpand`, `set_vexpand` (GTK/NWG both have these) |
| `TextView` | `append_text` — zork *has* it, but as O(n) read-modify-write (NWG is too) |
| `Entry` | `on_key` (NWG) |

### 2.3 Cross-backend parity methods (present on GTK **and** NWG, missing on zork)

Real, useful, and implementable in a headless model:

| # | Type | Method | GTK | NWG | panc |
|---|------|--------|-----|-----|------|
| P1 | `Button` | `add_class` / `remove_class` | ✅ | ✅ | ✅ |
| P2 | `Button` | `set_size_request` | ✅ | ✅ | ✅ |
| P3 | `Button` | `set_font_style` | ✅ | ✅ | ✅ |
| P4 | `Button` | `set_visible` | ✅ | — | — |
| P5 | `Button` | `set_hexpand` / `set_vexpand` | ✅ | — | — |
| P6 | `BoxWidget` | `set_visible`, `set_size_request`, `set_hexpand`, `set_vexpand` | ✅ | partial | partial |
| P7 | `BoxWidget` | `set_child_hexpand` / `set_child_vexpand` | ✅ | ✅ | ✅ |
| P8 | `BoxWidget` | `layout(x,y,w,h)` | ❌ (GTK) | — | — |
| P9 | `BoxWidget` | `request_layout` | — | ✅ | — |
| P10 | `Grid` | `set_visible`, `set_size_request`, `set_hexpand`, `set_vexpand` | ✅ | — | — |
| P11 | `Grid` | `layout()` | ❌ (GTK) | — | — |
| P12 | `Label` | `set_halign` / `set_valign` | — | ✅ | — |
| P13 | `Label` | `set_size_request`, `set_hexpand`, `set_vexpand` | ✅ | — | — |
| P14 | `Label` | `add_class`/`remove_class` | ❌ (GTK Label lacks it; zork has it) | ❌ | — |
| P15 | `CheckButton` | `set_label` / `get_label` | — | ✅ | — |
| P16 | `CheckButton` | `set_visible`, `set_size_request`, `set_hexpand`, `set_vexpand` | ✅ | — | — |
| P17 | `RadioButton` | `set_label` / `get_label`, `grab_focus` | grab_focus ✅ | set/get_label ✅ | — |
| P18 | `RadioButton` | `set_visible`, `set_size_request`, `set_hexpand`, `set_vexpand` | ✅ | partial | — |
| P19 | `TextView` | `set_visible`, `set_hexpand`, `set_vexpand` | ✅ | — | — |
| P20 | `TextView` | `set_editable` / `get_buffer` | — | ✅ | — |
| P21 | `DropDown` | `grab_focus`, `set_visible`, `set_size_request`, `set_hexpand`, `set_vexpand`, `set_offset` | ✅ | ✅ | partial |
| P22 | `Canvas` | `set_hexpand`, `set_vexpand`, `set_margin_start`, `set_margin_top` | ✅ | partial | — |
| P23 | `Canvas` | `clear_draw_callback` | — | — | ✅ |
| P24 | `Overlay` | `set_hexpand`, `set_vexpand` | ✅ | ✅ | — |
| P25 | `Overlay` | `set_overlay_pass_through` (zork has it, GTK/NWG do not) | — | — | — |
| P26 | `ScrolledWindow` | `set_size_request`, `set_hexpand`, `set_vexpand` | ✅ | partial | partial |
| P27 | `Dialog` | `set_visible`, `set_size_request` | — | ✅ | — |
| P28 | `Dialog` | `get_content_area`, `mark_destroyed`, `set_default_response` | ✅ | set_default_response ✅ | — |
| P29 | `Dialog` | `bind_esc_dismiss`, `collect`, `layout_dialog`, `run` | — | ✅ | — |
| P30 | `Window` | `set_layout_cb`, `start_repeating_timer` | — | ✅ | — |
| P31 | `Menu` | `append_item`, `append_section` | ✅ | ✅ | — |
| P32 | `Menu` | `append_with_shortcut`, `append_submenu_with_shortcut`, `append_separator`, `append_key`, `append_check`, `append_radio` | — | — | ✅ |
| P33 | `SimpleAction` | `on_activate` in addition to `connect_activate` | `connect_activate` ✅ | ✅ | ✅ |
| P34 | `CheckButton`/`RadioButton` | `on_toggle` (zork has it; GTK/NWG name it `connect_toggled`) | — | — | — |

### 2.4 Free-function gaps

| # | Function | GTK | NWG | panc | Notes |
|---|----------|-----|-----|------|-------|
| F1 | `create_fixed` | ✅ | — | — | needs M4 |
| F2 | `create_application`, `add_action` | ✅ | — | — | needs M5 |
| F3 | `create_spreadsheet` (+ 12 `spreadsheet_set_*`) | ✅ | — | ✅ | needs M6 |
| F4 | `timeout_add_once`, `timeout_add_repeating` | ✅ | `start_repeating_timer` | — | GTK-only |
| F5 | `pump_main_context`, `quit_main_loop` | ✅ | `quit_main_loop` | — | zork's REPL **is** the loop (`owns_event_loop() == true`), so this is `[-]` |
| F6 | `open_file_filtered` / `save_file_filtered` | ✅ | ✅ | — | zork has only unfiltered `open_file`/`save_file`, both hard-coded `Ok(None)` |
| F7 | `set_window_pos`, `mark95a`, `mark95xy`, `translate_vk`, `modifier_state`, … | — | ✅ | — | Win32/Win95 only → `[-]` |
| F8 | `debug_dump_native_tree`, `window_queue_redraw_cascades_to_children` | — | ✅ | — | native handle introspection → `[-]` |
| F9 | `create_dialog_button`, `set_menu_text`, `set_border_title`, `set_status_text`, `set_formula_bar`, `set_formula_bar_trailing`, `set_grid_config`, `set_column_layout`, `set_row_counts`, `set_row_labels`, `set_tab_data`, `set_cursor`, `set_editing`, `set_raw_cell`, `get_cell`, `set_cell`, `set_cell_style` | — | partial | ✅ | spreadsheet/textview surface |
| F10 | `add_key_callback`, `add_goto_callback`, `add_cursor_move_callback`, `add_commit_edit_callback` | — | — | ✅ | spreadsheet input model |

## 3. Model-level semantic gaps (not just missing methods)

These are behavioural: the model has no way to *represent* the state, so no
adapter method can be more than a no-op until the model grows.

| # | Gap | Detail |
|---|-----|--------|
| S1 | No geometry | Nothing stores a rect. `BoxWidget::layout(x,y,w,h)`, `Grid::attach(l,t,w,h)`, `set_size_request`, `set_default_size`, `set_offset` all discard their arguments. `Grid`'s `cols`/`rows` are hard-coded `0` and never written. |
| S2 | No visibility | `set_visible` is a no-op on every type. `show_all` on `Overlay` does nothing. |
| S3 | No expansion flags | `set_hexpand`/`set_vexpand`/`set_child_hexpand`/`set_child_vexpand` are no-ops. |
| S4 | No margins | `set_margin_start`/`set_margin_top` (Label, Entry, Canvas) are no-ops. |
| S5 | No CSS classes | `add_class`/`remove_class` are no-ops; a `classes: Vec<String>` per node is the obvious model. |
| S6 | No focus | `set_focus` is an empty fn. `grab_focus` is a no-op, `has_focus` hard-coded `false`, `connect_focus_in/out_event` register an *untyped* callback. No `focused` node id. |
| S7 | No alignment | `set_xalign`, `set_halign`, `set_valign`, `set_fixed_width` are no-ops. |
| S8 | No font style | `set_font_style` is a no-op on Button. |
| S9 | No caret | `Entry` has a `cursor` field, but `get_position` hard-codes `None` and `set_position` is a no-op — the field is written once by `set_entry_text` and never read. |
| S10 | No editability | `TextView::set_editable` / `get_buffer` absent. |
| S11 | No scroll state | `ScrolledWindow` has no `scroll_to`/`on_scroll`; nothing records a scroll position or range. |
| S12 | No draw callback / draw surface | `Canvas::set_draw_callback` discards the closure; `force_draw`, `set_content_size`, `queue_redraw` are no-ops. Zork has **no** `DrawContext` implementation at all (only the dependency-free `backends::headless` recording context exists). |
| S13 | No pointer/key event routing | `on_click`, `on_click_button`, `on_motion`, `on_release`, `on_key`, `on_key_raw`, `connect_button_press` all accept a closure and never store it. `screen_origin` is hard-coded `None`. |
| S14 | Menu model is lossy | `MenuItemData` has only `label`/`action`/`submenu`. Missing: accelerator/shortcut, kind (normal/separator/check/radio), checked state, section headers, `enabled`. `menu_append_submenu` copies submenu items by value at append time, so later edits to the submenu are lost. `MenuBar` has no mnemonic state, no "menu is open" flag, no selection cursor. |
| S15 | Menu selection does not dispatch | `Harness::select` validates the index then fires the *node's* callbacks — it never routes to the `SimpleAction` the item names, and `ZorkState::click` on a `Menu`/`MenuBar` fires all callbacks indiscriminately. REPL `select` only prints. |
| S16 | Radio groups are not enforced | `create_radiobutton(group_id, ..)` stores `group_id` but `set_radiobutton_checked` never clears siblings. GTK/NWG enforce mutual exclusion. |
| S17 | Dialog responses absent | `add_button`/`connect_response` are no-ops; no response id, no `default_response`, no `transient_for` parent, no `mark_destroyed`. |
| S18 | No `Clone` on 5 adapter types | `BoxWidget`, `Menu`, `SimpleAction`, `MenuBar`, `Dialog` (B4). |
| S19 | `MenuBar::insert_action_group` | zork's `MenuBar` has no such method; `Window::insert_action_group` exists but is a no-op. |
| S20 | Radio group lost in `core.rs` | `core.rs::create_radiobutton` does `let _ = group; create_radiobutton(None, label)` — the group argument is discarded even though the adapter supports it. |
| S21 | No event loop / quit plumbing | `ZorkState::quit` sets `running = false` but nothing observes it except the REPL. `Window::present`/`Dialog::present`/`close` are no-ops. |
| S22 | No `owns_event_loop` parity issue | REPL returns `true` — correct, but `ZorkState` is also used by the harness, which has no loop at all. `[-]` (correct as-is) |
| S23 | Snapshot is too thin | `SnapshotNode` has no geometry/visibility/classes/focus/size-request/caret, so a snapshot regression test cannot detect a regression in any of §3's properties. `Snapshot` also omits `prev_location` and any scroll/expansion state. |
| S24 | `set_child` is destructive-but-buggy | `set_child` removes the child from *all* parents' `children` lists but only fixes up the **new** parent's link; the old parent's list is cleaned, but `add_node` also pushes to the parent, so a node can end up duplicated. And `append_child` is an alias of `set_child` — it *replaces* rather than appends, despite the name. |
| S25 | No `Fixed`/`Overlay` pass-through | `Overlay::add_overlay` and `set_child` are the same code path; pass-through is a no-op. |

## 4. Harness (`backends::zork::harness`) gaps

The harness is the *recommended* way to test against zork, so its surface
matters as much as the adapter's.

| # | Gap |
|---|-------|
| H1 | No `create_canvas` / `create_overlay` / `create_scrolled_window` handles |
| H2 | No geometry handles: `set_size_request`, `set_offset`, `layout`, `attach` |
| H3 | No `set_visible` / `set_hexpand` / `set_vexpand` / `set_margin_*` / `add_class` |
| H4 | No `set_focus` / `has_focus` / `grab_focus` on any handle |
| H5 | No `set_position`/`get_position` on `Entry` (the model has the field) |
| H6 | `Harness::select` does not dispatch to the action the menu item names (S15) |
| H7 | No `Harness::key_*` / `click_at(x,y)` / `motion` / `release` actions (S13) |
| H8 | No `Harness::add_class`/`classes` getters for snapshot assertions (S5) |
| H9 | `snapshot_json` cannot express the §3 properties (S23) |

## 5. REPL (`backends::zork::repl`) gaps

| # | Gap |
|---|-------|
| R1 | No verbs for any §3 property — nothing can `look` at geometry, visibility, focus or classes |
| R2 | `select` only prints the action name; it never fires the named `SimpleAction`'s callbacks |
| R3 | No verb to set properties (`set title`, `hide`, `show`, `focus`) |
| R4 | `toggle` on a `RadioButton` does not clear the rest of the group (S16) |
| R5 | `click` on a non-Button falls through to "You can't click that", even for `Entry`/`CheckButton`/`DropDown` which *do* have click semantics |
| R6 | `type` requires standing on the `Entry`; no way to target a specific entry |

---

## Implementation order

1. **Unblock the build** — B1..B6. Nothing else can be verified until
   `cargo build --features zork` is green. Add the six missing model node kinds
   (M1–M3 at minimum), give the 5 adapter types `Clone` (S18), delete the
   duplicate `impl Clone`, write the 9 `core.rs` `new_*` arms, add the missing
   `common_types_mod!` methods, and give zork the same `common_types` cfg-tail
   (`Canvas::on_key`, `ScrolledWindow::*`) every other backend has.
2. **Model properties** (S1–S8, S23) — add a `ZorkProps` struct (rect, visible,
   expand flags, margins, classes, font, align) plus `focus`, `caret`,
   `editable`, `scroll`. One struct keeps `ZorkNode` from growing 15 fields and
   makes the snapshot serializable in one go.
3. **Wire the adapter** (P1–P34) to those properties.
4. **Event routing** (S13) — store click/motion/release/key callbacks on
   `Canvas` and add harness actions to fire them.
5. **Menu semantics** (S14–S16) — extend `MenuItemData`, fix `set_child`
   (S24), make `select` dispatch (S15, H6, R2), enforce radio groups.
6. **Dialog responses** (S17).
7. **Harness + REPL** (H1–H9, R1–R6).
8. **Spreadsheet / Fixed / Application** (M4–M6, F1–F3, F9–F10) — optional;
   only reachable from GTK/pancurses `core.rs` paths today.

### Explicitly out of scope (`[-]`)

* Win32/Win95 shims (F7), native-handle introspection (F8), `pump_main_context`
  / `quit_main_loop` (F5 — zork's REPL is the loop), CSS-vs-`pdcurses-sys`
  feature conflicts (a pre-existing, unrelated workspace resolution failure:
  `rustxWidgets/Cargo.lock` has a registry `pdcurses-sys 0.7.1` while the
  repo-root `[patch.crates-io]` points `pdcurses-sys` at `vendor/pdcurses-sys`.
  Build zork from the **repo-root** workspace, which has a correct lock file).
