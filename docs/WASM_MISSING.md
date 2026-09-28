# WASM backend — missing features vs. GTK / NWG

Status: **implemented**. Every row below is either `[x]`, or `[-]` with a
reason. The `wasm32` target now compiles, for both the widget crate and the
app.
Last updated: 2026-09-28.

The `wasm32` backend (`rustxWidgets/rswidgets/src/backends_wasm_adapter.rs`,
1759 lines) renders the shared GUI pipeline as DOM. It is nominally a
first-class backend: `core.rs` and `common.rs` have `#[cfg(target_arch =
"wasm32")]` arms everywhere. This document enumerates **every** gap found,
ranked by severity, and tracks which are implemented.

Method: mechanically diff the public inherent-method surface of the WASM
adapter against `backends_gtk_adapter_impl.rs` (GTK, 1823 lines) and
`backends_nwg_adapter.rs` (NWG/Win32, 3536 lines), then subtract the surface
that `common.rs` / `core.rs` actually require.

### A note on the method, written after the fact

To be fair to the diff: it was right. Every method §1 and §3 list as missing
was missing, and §0's error count (34) was within one of the 35 the compiler
produced when the enumeration was finally checked — a surface diff against a
file that is nominally a first-class backend is good evidence.

What it could not reach is the honest limit of the method. The diff compares
the *adapter* against GTK and NWG; it does not compile anything, and so it
never saw the seven cfg errors in the app crate itself (`corro_main` compiled
for a platform with no CLI, `libc::EXDEV` named where `libc` is excluded,
`ios_backend` absent while six call sites referenced it). Those were on no
list — not because the diff was careless, but because a diff of one file
against two others cannot find a fourth. `gui_loop.sh` would have caught them
if it had run; it is guarded on the target being installed for the `nightly`
toolchain, which does not have it, so it printed `skip` and exited 0 — a green
run that proved nothing.

Two consequences, both now in place:

- The build is checked in CI. `.github/workflows/wasm.yml` pins `+stable`,
  installs the target, sets `-D warnings`, builds the app (not just checks
  it), and asserts the emitted `.wasm` has the right magic number. The
  app-crate cfg gates are only compiled by that job's second step, so
  "rswidgets builds" alone would still be blind to them.
- The lesson is in §0: **check the build before enumerating gaps.** The order
  here was inverted — enumerate, then implement — and it worked only because
  the enumeration happened to be thorough. For a backend with no CI coverage,
  "does it even compile?" is a question with a one-line answer and a large
  payoff, and it should be the first question.

Legend: `[ ]` todo · `[~]` in progress · `[x]` done · `[-]` deliberately not
applicable to a browser DOM (with a reason).

---

## 0. [x] The `wasm32` target compiles

Both `cargo check -p rswidgets --target wasm32-unknown-unknown
--no-default-features` and the app's own build (`--target wasm32-unknown-unknown
--features wasm --no-default-features`, which is what `gui_loop.sh:390`
checks) are clean. The history of this section, kept because it explains why
the rest of the document reads as it does:

`cargo +stable check -p rswidgets --target wasm32-unknown-unknown
--no-default-features` fails with **34 errors** (34 as counted by hand from the
tables below; the compiler reports 35 — the extra one is a downstream
consequence of another row, so the two counts are consistent). The backend is
dead code the
same way `zork` was before §0 of `ZORK_MISSING.md`: no CI job builds it.
(The `nightly` toolchain in `rust-toolchain.toml` has no `wasm32-unknown-unknown`
std installed, so the check must be run `+stable`.)

| # | Error class | Where | Why |
|---|-------------|-------|-----|
| B1 | `E0599` × 1 | `common.rs:119` | `common_types_mod!` calls `Window::resize` for **every** backend. The WASM `Window` (wasm 138–225) has no `resize`; it only has `set_default_size`, which under CSS is *advisory* in exactly the way the wrapper's own doc comment warns about. |
| B2 | `E0599` × 2 | `common.rs:122,123` | Same cause: `Window::hwnd` and `Window::set_child_box` are required unconditionally. The WASM `Window` has neither (it has `set_child(&impl AsElement)` instead). |
| B3 | `E0599` × 3 | `common.rs:141,145,150` | `Label::set_fixed_width` / `set_margin_start` / `set_margin_top`. |
| B4 | `E0425` missing value | `core.rs:2050` | `App::new_menubar`'s `#[cfg(target_arch = "wasm32")]` arm passes `action_group`, but the function parameter is named `_action_group` (`core.rs:2031`). |
| B5 | `E0425` × 3 | `backends_wasm_adapter.rs:1427,1442,1576` | `HtmlCanvasElement` is used in three type positions but is **not** in the adapter's `use web_sys::{…}` list (wasm 10-13). `common.rs`'s Cargo.toml *does* enable the `web-sys` feature; the import is the only gap. |
| B6 | `E0119` conflicting `Clone` | `backends_wasm_adapter.rs:1696` | `ScrolledWindow` has `#[derive(Clone)]` (1696) *and* a hand-written `impl Clone` (1713). |
| B7 | `E0592` duplicate definition | `common.rs:452` | `set_vexpand` is declared twice in the wasm `common_types` block (451 and 452). |
| B8 | `E0277` missing `Clone` | `common.rs:108` | `common_types_mod!` derives `Clone` on `Dialog`; the WASM `Dialog` (wasm 922) does not implement it. |
| B9 | `E0599` × 8 | `common.rs:297-334` | `impl MenuBar` in the WASM adapter is an **empty block** (wasm 785-787), so all eight methods `common_types_mod!` requires are absent. |
| B10 | `E0599` × 6 | `common.rs:215,217,232,233,247,248` | `Entry::{set_vexpand, set_visible, set_halign, set_valign, set_margin_start, set_margin_top}` — all called unconditionally by the wrapper, none present in the adapter. |
| B11 | `E0599` | `common.rs:346` | `Dialog::set_transient_for`. |
| B12 | `E0599` | `common.rs:153` | `Label::raw_handle` — the WASM `Label` *does* implement `core::Widget` (wasm 327), but the `common_types` module never imports the trait, so the method is not in scope. |
| B13 | `E0599` | `common.rs:357` | `ScrolledWindow::as_ref` — the WASM `ScrolledWindow` has no `impl AsRef<*mut c_void>`. |
| B14 | `E0425` × 3 | `backends_wasm_adapter.rs:668,1544,1560` | `store_closure` is called three times but **never defined** — the `Canvas`/`Entry` DOM listeners that were meant to own their `Closure` are uncompilable. |
| B15 | `E0599` | `backends_wasm_adapter.rs:1396` | `CanvasRenderingContext2d::measure_text` needs the `CanvasTextMetrics`/`TextMetrics` `web-sys` feature, which `common.rs`'s Cargo.toml does not enable. Without it `text_extents_styled` cannot measure at all. |

All fifteen are fixed. Everything in §1–§5 was downstream of them: nothing
could be verified until the target built.

The app crate had its own seven, in cfg gates that never accounted for
`--features wasm` being neither `gui` nor a desktop target — `ios_backend`
not compiled for wasm while six `log_ios` calls referenced it, `libc::EXDEV`
named where `libc` is excluded, and `corro_main`/`parse_args` compiled for a
platform that has no CLI. Those are fixed too, so the wasm build that
`gui_loop.sh` reports on is real.

---

## 1. Missing `common_types_mod!` requirements (build blockers — B1–B13)

Called unconditionally from `common.rs`, so WASM must provide them:

All now present. Notes on the two that are not a straight port:

| Type | Required methods | WASM has |
|------|------------------|----------|
| `Window` | `resize`, `hwnd`, `set_child_box` | ✅ (`resize` sets `min-*` as well as `width`/`height`, since a CSS width on a flex child is only a request; `hwnd` is `null_mut`, the honest answer) |
| `Label` | `set_fixed_width`, `set_margin_start`, `set_margin_top`, `raw_handle` (B12: import only) | ✅ (`set_fixed_width` sets `min-width` too, and clears both on `None` — corro pins the formula-bar address and status labels precisely so changing text cannot resize them) |
| `Entry` | `set_vexpand`, `set_visible`, `set_halign`, `set_valign`, `set_margin_start`, `set_margin_top` | ✅ |
| `MenuBar` | `activate_submenu_by_mnemonic`, `popup_submenu_by_mnemonic_at`, `activate_submenu_item_by_mnemonic`, `insert_action_group`, `handle_mnemonic_key`, `handle_menu_key`, `menu_active`, `menu_close` | ✅ — `impl MenuBar` was an empty block; all eight are now real over a `MenuEntry` registry. `insert_action_group` stays a no-op (see §7). |
| `Dialog` | `set_transient_for` | ✅ — the browser's answer to "parent-modal" is `show_modal`, which blocks the rest of the page. The parent handle itself is not meaningful and is ignored. |
| `ScrolledWindow` | `AsRef<*mut c_void>` | ✅ |
| `Dialog` | `Clone` | ✅ (derived; `HtmlDialogElement` is a handle) |

## 2. Missing `common_types` cfg-tail methods

| Type | Methods present on GTK/NWG, absent on WASM |
|------|--------------------------------------------|
| `ScrolledWindow` | `scroll_to(hval,hupper,hpage,vval,vupper,vpage)` — **hard no-op** at `common.rs:448`; `on_scroll` — **hard no-op** at `common.rs:449` |

`set_child` / `set_policy` / `set_vexpand` exist as no-ops on WASM but the
*adapter* has real versions, so the common layer deliberately discards them
(this is the one place where the wasm cfg-tail is more inert than it needs to
be — see §5.3).

## 3. Cross-backend parity methods (present on GTK **and/or** NWG, missing or stubbed on WASM)

### 3.1 `Window`

| # | Method | GTK | NWG | WASM | Notes |
|---|--------|-----|-----|------|-------|
| W1 | `on_close` | ✅ 103 | ✅ 533 | **no-op** (198) | GTK wires `close-request`, NWG `WM_CLOSE`. The DOM's `beforeunload` is the analogue. |
| W2 | `on_event_key` | ✅ 123 | ✅ 530 | **partial** (200) | Registers `keydown` on `document`, not the element; and forwards **only** `alt_key` (`state |= 0x8`), never Shift (1) or Ctrl (4). GTK's callback receives the full GDK modifier mask. |
| W3 | `queue_redraw` | ✅ 69 | ✅ 510 | **no-op** (199) | GTK `queue_draw`, NWG `RedrawWindow(RDW_ALLCHILDREN)`. |
| W4 | `present` | ✅ 60 | ✅ 412 | **no-op** (184) | Acceptable: the div is appended to `body` at creation (227). `[-]` |
| W5 | `resize` | ✅ | ✅ | **absent** (B1) | → `elem.style().width/height`. |
| W6 | `hwnd` | ✅ 86 | ✅ 434 | **absent** (B2) | `[-]` — no OS handle. Return `null_mut()`, matching zork/ios/android/macos. |
| W7 | `set_child_box` | ✅ 58 | ✅ 384 | **absent** (B2) | → forward to the existing `set_child`. |
| W8 | `set_layout_cb` | ❌ | ✅ 397 | ❌ | NWG-only `WM_SIZE` hook. `[-]` — CSS flexbox does the layout. |
| W9 | `start_repeating_timer` | ❌ (free fn) | ✅ 449 | ✅ | Backs `core::add_periodic_tick`. → `setInterval`. |

### 3.2 `Button`

| # | Method | GTK | NWG | WASM | Notes |
|---|--------|-----|-----|------|-------|
| B-1 | `set_visible` | ✅ 254 | ✅ | ✅ | `style.display`. |
| B-2 | `set_size_request` | ✅ 253 | ✅ 730 | ✅ | CSS width/height. |
| B-3 | `set_hexpand` / `set_vexpand` | ✅ 255-256 | ✅ | ✅ | `flex-grow` / `align-self`. |
| B-4 | `set_font_style(weight, italic)` | ✅ 257 | ✅ 741 | ✅ | `font-weight` / `font-style`. |
| B-5 | `add_class` / `remove_class` | ✅ 258-259 | ✅ 771 (**no-op**) | ✅ | `classList` — the WASM `Label`/`Entry` already do this. |

### 3.3 `Label`

| # | Method | GTK | NWG | WASM | Notes |
|---|--------|-----|-----|------|-------|
| L1 | `set_fixed_width(Option<i32>)` | ✅ 306 | ✅ 845 | ✅ | **High-value.** Backed by `common::Label::set_fixed_width` (common.rs:136). `None` must clear the pin (`width:auto`), not just ignore. |
| L2 | `set_margin_start(px)` | ✅ 290 | ✅ 929 (**no-op**) | ✅ | `margin-left` (LTR). |
| L3 | `set_margin_top(px)` | ✅ 315 | ✅ 930 (**no-op**) | ✅ | `margin-top`. |
| L4 | `set_hexpand` / `set_vexpand` | ✅ 282-283 | ✅ | ✅ | `flex-grow` / `align-self`. |
| L5 | `set_size_request(w,h)` | ✅ 284 | ✅ | ✅ | CSS width/height. |
| L6 | `set_halign` / `set_valign` | ❌ | ✅ | ✅ | NWG-only; → `justify-self` / `align-self`. |
| L7 | `set_xalign` | ✅ 317 | ✅ 911 | ✅ 371 | **Not missing** — `text-align`. |
| L8 | `set_markup` | ✅ 280 | ✅ 907 | ✅ 357 | **Not missing** — `innerHTML`. |
| L9 | `add_class` / `remove_class` | ✅ 278-279 | ❌ | ✅ 349/353 | **Not missing.** |

### 3.4 `BoxWidget` / `Grid`

| # | Method | GTK | NWG | WASM | Notes |
|---|--------|-----|-----|------|-------|
| X1 | `BoxWidget::set_size_request` | ✅ 336 | ✅ | ✅ |  |
| X2 | `BoxWidget::set_vexpand` | ✅ 337 | ✅ 1236 (**no-op**) | ✅ |  |
| X3 | `BoxWidget::set_visible` | ✅ 339 | ✅ | ✅ |  |
| X4 | `BoxWidget::layout(x,y,w,h)` | ❌ | ✅ 1070 | ❌ | NWG-only absolute layout. `[-]` — CSS flexbox. |
| X5 | `BoxWidget::request_layout` | ❌ | ✅ 1225 | ❌ | `[-]` — same reason. |
| X6 | `Grid::set_visible` | ✅ 368 | ✅ | ✅ |  |
| X7 | `Grid::set_hexpand` / `set_vexpand` | ✅ 369-370 | ✅ | ✅ |  |
| X8 | `Grid::set_size_request` | ✅ 371 | ✅ | ✅ |  |

### 3.5 `Entry`

| # | Method | GTK | NWG | WASM | Notes |
|---|--------|-----|-----|------|-------|
| E1 | `has_focus` | ✅ 400 | ✅ 1460 | **returns `false`** (621) | `document.activeElement === elem` is available. Reachable via `common::Entry::has_focus` (common.rs:234); corro uses it to decide push-vs-append, so the stub silently changes editing behaviour. |
| E2 | `get_position` / `set_position` | ✅ 384/387 | ✅ 1402/1417 | `None` / no-op (652-653) | `selectionStart`/`setSelectionRange`. Behaviour is currently preserved by `common::Entry::caret_override`, so **low severity** — but only because the wrapper compensates. |
| E3 | `set_visible` | ✅ 411 | ✅ 1457 | ✅ |  |
| E4 | `set_vexpand` | ✅ 413 | ✅ 1506 (**no-op**) | ✅ |  |
| E5 | `set_margin_start` / `set_margin_top` | ✅ 407-408 | ✅ 1509/1521 | ✅ | Both reachable from `common.rs:242-243`. |
| E6 | `set_halign` / `set_valign` | ✅ 409-410 | ✅ 1533-1534 (**no-op**) | ✅ | Both reachable from `common.rs:227-228`. |
| E7 | `on_key` | ❌ | ✅ 1502 | ❌ | NWG `Box`-form. `common::Entry::on_key` exists only on NWG (common.rs:403), so this is adapter-only. |
| E8 | `set_hexpand` | ✅ 412 | ✅ 1505 (**no-op**) | **no-op** (581) | Not a gap — no-op on NWG too. |

### 3.6 `DropDown`

| # | Method | GTK | NWG | WASM | Notes |
|---|--------|-----|-----|------|-------|
| D1 | `set_active(Option<u32>)` | ✅ 653 | ✅ 1765 | **`set_active(u32)`** (1051) | **Signature mismatch, not a stub.** GTK/NWG take `Option` so a caller can *clear* the selection; WASM cannot express "no selection". A `<select>` does this with `selectedIndex = -1`. |
| D2 | `grab_focus` | ✅ 656 | ✅ 1770 | ✅ | `elem.focus()`. |
| D3 | `set_visible` | ✅ 666 | ✅ 1804 | ✅ |  |
| D4 | `set_hexpand` / `set_vexpand` | ✅ 667-668 | ✅ 1788-1789 (**no-op**) | ✅ |  |
| D5 | `set_size_request` | ✅ 669 | ✅ 1791 | ✅ |  |
| D6 | `set_offset(x,y)` | ✅ 671 | ✅ 1819 | ✅ | **Functional gap**: how an in-overlay dropdown is placed. WASM's `Overlay::add_overlay` sets `position:absolute` + `z-index:10` (1633) but never an offset, so an overlaid dropdown lands at the overlay's top-left corner. |
| D7 | `active() -> Option<u32>` | ❌ | ✅ 1779 | ✅ | → `selected_index() >= 0`. |
| D8 | `diagnostics` / `is_gtk4` / `has_size_request_symbol` | ✅ 658-661 | ❌ | ❌ | `[-]` — GTK-version/symbol probes. |

### 3.7 `CheckButton` / `RadioButton`

| # | Method | GTK | NWG | WASM | Notes |
|---|--------|-----|-----|------|-------|
| C1 | `set_visible` | ✅ 696/721 | ✅ | ✅ |  |
| C2 | `set_size_request` | ✅ 699/724 | ✅ | ✅ |  |
| C3 | `set_hexpand` / `set_vexpand` | ✅ 697-698, 722-723 | ✅ 1908-1909 (**no-op**) | ✅ |  |
| C4 | `set_label` / `get_label` | ❌ | ✅ 1862-1863, 1903-1904 | ✅ | NWG-only (set label post-construction). → mutate the wrapper's `Text` node. |
| C5 | `RadioButton::grab_focus` | ✅ 717 | ✅ 1892 | ✅ |  |

### 3.8 `TextView`

| # | Method | GTK | NWG | WASM | Notes |
|---|--------|-----|-----|------|-------|
| T1 | `set_visible` | ✅ 747 | ✅ | ✅ |  |
| T2 | `set_hexpand` / `set_vexpand` | ✅ 748-749 | ❌ | ✅ 1310/1318 | **Not missing.** |
| T3 | `set_editable` | ❌ | ✅ 1942 | ✅ | → `elem.set_read_only(!editable)`. |
| T4 | `get_buffer` | ❌ | ✅ 1939 | ❌ | NWG-only; exposes the stored `changed` callback. |

### 3.9 `Dialog`

| # | Method | GTK | NWG | WASM | Notes |
|---|--------|-----|-----|------|-------|
| G1 | `set_transient_for` | ✅ 625 | ✅ 1996 (centres) | ✅ | Reachable from `common.rs:341`. → `HtmlDialogElement::show_modal()`. |
| G2 | `set_default_response` | ✅ 628 | ✅ 2127 | ✅ | Reachable only from adapter-direct callers (no `common.rs` wrapper). → mark the button `autofocus`. |
| G3 | `mark_destroyed` | ✅ 636 | ❌ | ❌ | `[-]` — GTK widget-lifecycle flag. |
| G4 | `set_visible` | ❌ | ✅ 2229 | ✅ | NWG-only (`ShowWindow`) → `show()`/`close()`. |
| G5 | Escape-dismiss | ✅ native | ✅ 2168 | ✅ | `<dialog>` fires `cancel` for Escape. A modal dialog opened by `set_transient_for` is closed by the browser's own Escape handling; a non-modal one needs a `cancel` listener calling `close()`. Low value on a page, where the browser already provides the gesture. |
| G6 | `layout_dialog` / `run` | ❌ | ✅ 2096 / 1991 (returns 0) | ❌ | `[-]` low value. |

### 3.10 `MenuBar`

`impl MenuBar` in the WASM adapter is an **empty block** (wasm 785-787). Every
method `common.rs:285-331` requires is absent — which is a second build
blocker the `#[cfg]` arm masks, since the wasm `common_types` block adds no
`MenuBar` impls and the shared `common_types_mod!` does call them.

| # | Method | GTK | NWG | WASM | Notes |
|---|--------|-----|-----|------|-------|
| M1 | `activate_submenu_by_mnemonic` | ✅ 536 | ✅ 3157 (**false**) | ✅ |  |
| M2 | `popup_submenu_by_mnemonic_at` | ✅ 544 | ✅ 3166 (**false**) | ✅ | **Functional gap** — the sheet-tab context menu. GTK is the only real impl. |
| M3 | `activate_submenu_item_by_mnemonic` | ✅ 547 | ✅ 3169 (**false**) | ✅ |  |
| M4 | `handle_mnemonic_key` | ✅ 553 | ✅ 3173 (**false**) | ✅ |  |
| M5 | `handle_menu_key` | ✅ 556 | ✅ 3176 (**false**) | ✅ |  |
| M6 | `menu_active` | ✅ 559 | ✅ 3179 (**false**) | ✅ |  |
| M7 | `menu_close` | ✅ 562 | ✅ 3182 (**no-op**) | ✅ |  |
| M8 | `insert_action_group` | ✅ 550 | ✅ 3172 (**no-op**) | ❌ | `[-]` — already a no-op on NWG. |

**Note**: WASM's `create_menubar` (791-897) builds *real* `<button>` elements
with click handlers, so the *behaviour* exists — only the keyboard-activation
API is missing. An `Alt+letter` keydown handler over the menu items plus an
`open` flag covers M1/M3/M4/M5/M6/M7 without any new platform API.

### 3.11 `Canvas`

| # | Method | GTK | NWG | WASM | Notes |
|---|--------|-----|-----|------|-------|
| V1 | `on_click_button` | ✅ 1110 | ✅ 2466 (**no-op**) | **hard no-op** (1508) | The comment says "the DOM handler does not yet forward button/state", but `MouseEvent::button()` is right there. **Real gap**: right-click / middle-click on the canvas is impossible, so no canvas context menu. |
| V2 | `on_motion` | ✅ 1195 | ✅ 2472 (**no-op**) | **hard no-op** (1514) | **Real gap**: no hover, no drag. `mousemove` is available. |
| V3 | `on_release` | ✅ 1239 | ✅ 2477 (**no-op**) | **hard no-op** (1528) | **Real gap**: press/release cannot be paired, so no drag-end. `mouseup` is available. |
| V4 | `screen_origin` | ✅ 1341 | ✅ 2486 | **returns `None`** (1523) | `getBoundingClientRect()` is available. This is what forces `open_sheet_context_menu` to fall back to an *unpositioned* popup. |
| V5 | `set_hexpand` / `set_vexpand` | ✅ 1403-1404 | ✅ 2428-2430 (**no-op**) | ✅ |  |
| V6 | `set_margin_start` / `set_margin_top` | ✅ 1407-1408 | ✅ | ✅ |  |
| V7 | `on_key` signature | ✅ `(u32,u32)` | ✅ `(u32)` | ✅ `(u32)` (1547) | GTK is the odd one out; `common::Canvas::on_key` (common.rs:362) papers over it. Not a gap. |
| V8 | `set_can_focus` / `force_draw` | ✅ 1390 / 991 | ✅ 2503-2504 (**no-op**) | ✅ 1571-1572 (**no-op**) | Not a gap — no-op on NWG too. |

### 3.12 `ScrolledWindow`

| # | Method | GTK | NWG | WASM | Notes |
|---|--------|-----|-----|------|-------|
| S1 | `scroll_to` | ✅ 1440 | ✅ 2878 | ✅ | **Fixed.** The 6-float model is *cells*, not pixels (`sync_scrollbars` in `gui_backend.rs` passes cell counts), while `scrollLeft`/`scrollTop` are pixels, so the cell fraction is scaled by the scrollable extent — `scrollHeight - clientHeight`, which is exactly the distance the offset ranges over. A non-scrollable axis is a no-op rather than a jump to 0. |
| S2 | `on_scroll` | ✅ 1471 | ✅ 2892 | ✅ | **Fixed** — a real `scroll` listener reporting the scrolled *fraction*, which is the shape `scroll_to_cursor` already clamps. GTK reports an adjustment value in its own units; here those units are the fraction, which is the same thing. |
| S3 | `set_size_request` | ✅ 1515 | ✅ 2770 | ✅ | The viewport box. A CSS width is right here without a `min-width` companion: a scroll container's size is its visible box and the scrollbar appears inside it. |
| S4 | `set_policy` | ✅ 1432 | ✅ 2871 (**no-op**) | ✅ 1720 (`i32`, not `u32`) | Signature divergence, not a gap. |
| S5 | `set_vexpand` | ✅ 1514 | ✅ 2896 (**no-op**) | ✅ 1734 | Not a gap. |
| S6 | `set_child` (cfg-tail) | ✅ | ✅ | **discarded** (`common.rs:448`) | The shared GUI calls `set_child(&crate::common::Canvas)` (a `Canvas`, not a `Canvas` *handle*), so the wasm cfg-tail cannot forward to the adapter's `set_child(&impl AsElement)` as-is. Needs a `Canvas`-taking shim, like the android/ios/macos arms already have. |

### 3.13 `Overlay`

| # | Method | GTK | NWG | WASM | Notes |
|---|--------|-----|-----|------|-------|
| O1 | `set_hexpand` / `set_vexpand` | ✅ 1539-1574 | ✅ 2738-2761 | ✅ 1666/1674 | Not a gap. |
| O2 | `set_margin_start` / `set_margin_top` | ✅ | ✅ | ✅ | GTK's `GtkOverlay` supports per-child margins. |

### 3.14 `DrawContext`

| # | Method | GTK | NWG | WASM | Notes |
|---|--------|-----|-----|------|-------|
| R1 | `draw_rgba_image` | ✅ **840** (overrides; straight→premultiplied for Cairo) | ⚠️ trait default (`core.rs:928`) | ⚠️ trait default | **The only trait method GTK overrides that WASM does not.** `putImageData` + `drawImage` make it implementable. |

All 8 *required* trait methods (`fill_rect`, `stroke_rect`, `draw_text_styled`,
`text_extents_styled`, `clear`, `save`, `restore`, `clip`) are present on all
three. GTK also redundantly overrides the `draw_text`/`text_extents` defaults;
that is not a gap.

## 4. Free-function gaps

| # | Function | GTK | NWG | WASM | Notes |
|---|----------|-----|-----|------|-------|
| F1 | `create_fixed` + `Fixed::put` | ✅ 1605 | ❌ | ❌ | GTK absolute-positioning container. `[-]` — CSS `position:absolute` inside `Overlay` (WASM's `Overlay` already sets `position:relative`, wasm 1682). |
| F2 | `create_application` + `Application::{register,as_ptr,add_action}` | ✅ 593-607 | ❌ | ❌ | `[-]` — WASM's own `register_action` (wasm 95) plus the JS click handlers in `create_menubar` replace the `GApplication`/`GActionGroup` mechanism entirely. Needed by `App::ensure_action_group` (core.rs:2128) and `register_action` (2144), both of which are no-ops off GTK. |
| F3 | `create_spreadsheet` + 12 methods | ✅ 1692-1730 | ❌ | ❌ | Adapter-direct only; `core.rs` panes it to pancurses/zork. |
| F4 | `timeout_add_once` | ✅ 759 | ❌ | ✅ | `setTimeout`. |
| F5 | `timeout_add_repeating` | ✅ 767 | ❌ (via W9) | ✅ | `setInterval`; the `wasm32` exemption in `core::add_periodic_tick` is gone (see W9). |
| F6 | `pump_main_context` | ✅ 1756 | ❌ | ❌ | `[-]` — the browser owns the event loop; `core::pump_events` (core.rs:2191) is already a no-op off GTK. |
| F7 | `open_file` / `open_file_filtered` | ✅ 1625-1664 | ✅ 3385-3389 | `[-]` + `✅ open_file_content` | **The path form cannot exist in a browser** and stays `Ok(None)`: the picker is modal and async, and yields a `File` with contents, never a path a later `std::fs` call can open. The capability is provided as `open_file_content(accept, on_chosen)` — hidden `<input type=file>`, gesture-triggered `.click()`, `FileReader`, callback. `src/io` gained `load_workbook_text` so content reaches the shared parser. |
| F8 | `save_file` / `save_file_filtered` | ✅ 1637-1664 | ✅ 3410-3418 | `[-]` + `✅ save_file_content` | Same reasoning. `save_file_content(filename, text)` — `Blob` + synthetic `<a download>` click. The browser owns the destination and never reports it, so no path is set and no "Saved to X" is claimed. |
| F9 | `quit_main_loop` | ✅ 1743 | ✅ 3442 | ✅ 32 | Not a gap (`document.title = "CORRO_QUIT"` + `win.close()`). |
| F10 | `debug_dump_native_tree` | ❌ | ✅ 3450 | ❌ | `[-]` — HWND walk. |
| F11 | `Appendable::collect_hwnds` | ❌ | ✅ 996 | ❌ | `[-]` — multi-HWND fan-out. |
| F12 | `is_gtk4` / `mark_destroyed` / `has_size_request_symbol` | ✅ | ❌ | ❌ | `[-]` — GTK-internal. |

## 5. Behavioural gaps (not just missing methods)

| # | Gap | Resolution |
|---|-----|------------|
| G1 | `Window::on_event_key` modifier mask is truncated | **Fixed.** One `gdk_modifier_mask` helper now reads Shift/Control/Alt, and every callback in the adapter uses it (or its `MouseEvent` twin) instead of Alt-only or a literal `0`. `Canvas::on_key_raw` is registered directly rather than through `on_key`, which discarded the state by construction. |
| G2 | `Canvas::on_click` reports only the primary button | **Fixed.** `on_click_button`, `on_motion` and `on_release` were comment-only bodies; all three now forward `MouseEvent::button()` and the modifier mask on `mousedown`/`mousemove`/`mouseup`. |
| G3 | `ScrolledWindow` cfg-tail is more inert than the adapter | **Fixed.** The `common.rs` wasm tail forwarded all five instead of discarding them (and declared `set_vexpand` twice). `set_child` needed a `Canvas`-taking shim, because the shared GUI passes a `Canvas` while the adapter takes `&impl AsElement`. |
| G4 | `DropDown` has no `Option` selection | **Fixed.** `set_active(Option<u32>)` maps to `selectedIndex` including `-1`; `active() -> Option<u32>` reads it back. `set_offset` added via `transform: translate`, which composes with the position `add_overlay` established. |
| G5 | `Entry::has_focus` always false | **Fixed** — `document.activeElement`. This was not neutral: it is what chooses push-vs-append while typing, so always-false made every keystroke push. The caret (`get_position`/`set_position`) was `None`/inert and now uses `selectionStart`/`setSelectionRange`, converting between the DOM's UTF-16 offsets and the character indices the shared API documents. |
| G6 | `MenuBar` submenu dropdowns ignore keyboard state | **Fixed.** All eight required methods are implemented over a new `MenuEntry` registry (label, toggle, dropdown, items) that `create_menubar` now fills: Alt+letter opens, a printable key activates the item it starts, Escape closes, Up/Down move the selection, and choosing an item closes the menu. |
| G7 | `TextView::append_text` is O(document) | **Not fixed, and not fixable here.** Read-modify-write is the only thing a `<textarea>` offers; the same limitation is documented in `ZORK_MISSING.md` for the terminal backend. A high-rate log pane should reimplement against the native handle. |
| G8 | No `ScrolledWindow::on_scroll`, so the reentrancy guard was untested | **Fixed.** `on_scroll` is a real `scroll` listener reporting the scrolled *fraction*, which is the shape `scroll_to_cursor` already clamps; `scroll_to` scales the cell-index domain by the scrollable extent. The `syncing_scroll` guard in `gui_backend.rs` now has something to guard. |

## 6. Verified NOT missing (a WASM equivalent exists — do not implement)

| GTK / NWG item | WASM equivalent |
|----------------|-----------------|
| `Window::set_default_size` | CSS `width`/`height` px (wasm 186) |
| `Label::set_xalign` | CSS `text-align` (wasm 371) |
| `Label::set_markup` | `innerHTML` (wasm 357) |
| `Label::add_class` / `remove_class` | `classList` (wasm 349/353) |
| `Entry::add_class` / `remove_class` | `classList` (wasm 640/644) |
| `Entry::set_width_chars` | `input.size` (wasm 568) |
| `Entry::set_size_request` | CSS width/height (wasm 572) |
| `Entry::set_hexpand` | no-op on NWG too (wasm 581) |
| `CheckButton`/`RadioButton` `is_active`/`set_active`/`connect_toggled` | `input.checked` + `change` listener (wasm 1125-1147, 1201-1223) |
| `Dialog::get_content_area` / `append_content_area` | inner `<div>` (wasm 981/985) |
| `Dialog::present` / `close` / `connect_response` / `add_button` | `dialog.show()`/`close()` (wasm 989/998) |
| `Overlay::{set_child, add_overlay, remove, show_all, set_overlay_pass_through}` | same names, real DOM/CSS (wasm 1624-1659) |
| `Canvas::{set_draw_callback, queue_redraw, set_size_request, set_content_size, set_visible, grab_focus, on_click, on_key_raw}` | same names (wasm 1465-1570) |
| `TextView::{set_text, get_text, set_wrap_mode, set_size_request, set_hexpand, set_vexpand, append_text}` | same names (wasm 1284-1337) |
| `Menu::{append, append_submenu}` | same (wasm 736/748) |
| `SimpleAction::connect_activate` | `register_action` registry (wasm 904) |
| `Grid::attach` | CSS `grid-column`/`grid-row` (wasm 492) |
| `BoxWidget::{append, set_child_vexpand, set_child_hexpand}` | same (wasm 431/435/446) |
| `Button::{on_click, emit_clicked}` | `click` listener / `elem.click()` (wasm 274/289) |
| all 17 `create_*` factories | all present (wasm 227-1748) |
| `DrawContext`'s 8 required methods | all present (wasm 1371-1423) |
| `Entry::get_position`/`set_position` (behaviour) | shimmed by `common::Entry::caret_override` (common.rs:170-181) |
| `Window::on_event` | no-op on NWG (529) too |
| `Canvas::set_can_focus` / `force_draw` | no-op on NWG (2503/2504) too |
| `BoxWidget::set_hexpand` | no-op on NWG (1237) too |

## 7. WASM-only (the reverse direction)

| Item | Note |
|------|------|
| `append_wasm_output_line` (48) | Test-harness output sink (`window.__corro_output`). |
| `AsElement` trait (86) + its impls | The DOM's typing mechanism; GTK/NWG use `AsRef<*mut c_void>`. |
| `Orientation::as_flex_direction` (127) | Maps to CSS `flex-direction`. |
| `RadioButton` group via `input.name` (1213-1215) | Mutual exclusion falls out of the HTML radio-group semantics — a genuinely *better* answer than GTK's `set_group`, and worth copying into any backend that has to do it manually. |

---

## What was done

Ten commits on `wasm/impl`, in the order below. Each is independently
compilable; the build fix had to come first because nothing else could be
verified until the target built.

1. **`55456605` — unblock the build.** B1–B15: the 18 `common_types_mod!`
   methods, the empty `impl MenuBar` filled, the duplicate `impl Clone` and
   `set_vexpand` removed, `store_closure` defined, `HtmlCanvasElement`
   imported, `TextMetrics`/`DomRect` enabled, `_action_group` renamed, and
   `AsRef` implemented on `ScrolledWindow`. 35 errors → 0. Two restored
   methods also fixed behaviour rather than only adding surface:
   `Entry::has_focus` was hard-coded `false` (so every keystroke pushed), and
   the caret was `None`/inert.
2. **`40c58c85` — the app crate.** Seven cfg errors that had nothing to do
   with rswidgets and that made the build `gui_loop.sh:390` actually reports
   on fail.
3. **`585cc366` — the CSS property layer and canvas pointer events.** Every
   one-line property gap across all ten widget types, plus the four canvas
   methods that were comment-only stubs.
4. **`01a7be6c` — `draw_rgba_image` and timers.** `draw_rgba_image` was the
   only `DrawContext` method GTK overrides that WASM did not, so every caller
   fell through to the trait default and got `false` (an image preview on a
   page showed nothing). `add_periodic_tick` was in `core.rs`'s no-op list
   "because the browser owns the event loop" — true of the dispatch, false of
   the timer, so an app that armed a tick silently never ran.
5. **`0aaa45ff` — file dialogs.** See F7/F8: the path form cannot exist in a
   browser, so Open and Save were reimplemented over content, with
   `load_workbook_text` in `src/io` so the shared parser still does the work.

### One gap deliberately left

`TextView::append_text` is read-modify-write over the whole buffer, because
that is the only thing a `<textarea>` offers. It is G7 above, and the same
limitation is documented for the terminal backend in `ZORK_MISSING.md`.

### Still absent, by nature of a browser

`hwnd`, `insert_action_group`, `pump_main_context`, `is_gtk4`,
`has_size_request_symbol`, `mark_destroyed`, `Fixed`, `Application`,
`Spreadsheet`, `create_fixed`, `create_application`, and every Win32 shim
(`debug_dump_native_tree`, `Appendable::collect_hwnds`) have no DOM
equivalent. Each keeps the existing no-op / `null_mut()` pattern.

`Fixed` and the absolute-positioning it provides *is* reachable, via CSS
`position: absolute` inside `Overlay` (which already sets
`position: relative`) — so a caller that needs it should use the DOM types
directly rather than waiting for a `Fixed` wrapper. `Spreadsheet` and
`create_application` are reachable only from GTK `core.rs` arms today;
wiring them for WASM is a separate decision from making the backend itself
complete.

### The file-dialog shape, for a reader who expected paths

`open_file`/`save_file` keep their signatures and return `Ok(None)` on wasm.
This is not an oversight left over from the enumeration: a browser cannot
return a path from a synchronous function, because the picker is async and
yields a `File`. Returning a name and then handing it to `std::fs` would
produce a "no such file" at open time and a silent no-op at save time — a
worse failure than saying so. `open_file_content` and `save_file_content` are
the shapes a browser can satisfy, and the File menu uses them.

### Explicitly out of scope (`[-]`)

* Win32/Win95 shims, `Appendable`, `BoxWidget::layout`/`request_layout`,
  `debug_dump_native_tree` — all HWND-specific.
* `hwnd`, `insert_action_group`, `pump_main_context`, `is_gtk4`,
  `has_size_request_symbol`, `mark_destroyed`, `Fixed`, `Application` — no
  browser equivalent; keep the existing no-op / `null_mut()` pattern.
* `Spreadsheet` / `create_fixed` / `create_application` — reachable only from
  GTK `core.rs` arms today; wiring them into `core.rs` for WASM is a
  separate decision from making the backend itself complete.
