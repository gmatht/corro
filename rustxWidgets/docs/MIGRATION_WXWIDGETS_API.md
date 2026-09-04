# Migration plan: a wxWidgets-like API for rustxwidgets

**Goal:** reshape rustxwidgets' API along wxWidgets' proven designs, while
keeping Rust idioms (native primitives, closures, modules, memory safety).
Where the two conflict, prefer the wxWidgets *shape* over our handrolled
versions; where they don't, use native Rust.

## Naming: no `rx*` prefix

wxWidgets uses the `wx` prefix because C++ has no namespaces. Rust has
modules, so the prefix is redundant — `rustxwidgets::BoxSizer` is already
namespaced. **Adopt wxWidgets' type names, drop the prefix:**

| Current (handrolled) | wxWidgets name | Proposed |
|---|---|---|
| `BoxWidget` | `wxBoxSizer` | `BoxSizer` |
| `Grid` (layout) | `wxGridSizer` / `wxFlexGridSizer` | `GridSizer` / `FlexGridSizer` |
| `Entry` | `wxTextCtrl` | `TextCtrl` (or keep `Entry`) |
| `CheckButton` | `wxCheckBox` | `CheckBox` |
| `DropDown` | `wxChoice` | `Choice` |
| `TextView` | `wxTextCtrl` (multiline) | `TextView` (keep) |
| `Canvas` | `wxGLCanvas` / `wxPanel` | `Canvas` (keep) |
| `Menu` / `MenuBar` / `Dialog` / `RadioButton` | same | keep |

The `Sizer` hierarchy is the one real rename: `BoxWidget`/`Grid` become
`Sizer` + `BoxSizer` + `GridSizer` + `FlexGridSizer`.

## Design principles

1. **Native Rust *types* (structs/enums) when all else is equal.** `String`
   not `wxString`, `Option`/`Result` not sentinel values — and **typed
   structs, not bare primitives**: `Point { x, y }`, `Size { w, h }`,
   `Colour { r, g, b, a }` instead of passing bare `u32`/`f64` pairs. Bare
   numbers throw away the type information wxWidgets encodes in
   `wxPoint`/`wxSize`/`wxColour` (which field is x vs y, r vs g vs b); Rust
   structs encode it with public fields, derives, and pattern matching.
2. **Favor wxWidgets designs over handrolled ones.** Where wxWidgets has a
   proven shape (sizers, event propagation, standard dialogs, client data,
   standard menu items, tree-shaped menus), adopt it rather than inventing a
   thinner version.
3. **Keep Rust idioms where they're strictly better.** Closures over event
   tables, modules over prefixes, `Rc`/`RefCell` over raw pointers.

## Naming: `Point`, not `rxPoint`/`wxPoint`

- `wxPoint` — misleading: implies a wxWidgets type. We borrow wxWidgets'
  *design*, not its types.
- `rxPoint` — redundant: the `wx` prefix exists because C++ has no
  namespaces; Rust modules already namespace `rustxwidgets::Point`.
- `Point` — idiomatic. If an app has its own `Point`, it aliases:
  `use rustxwidgets::Point as RxPoint;`.

The type safety comes from the struct fields, not the prefix — all three
names would provide it equally, so pick the idiomatic one.

## Where the proposed design differs from wxWidgets

| Area | wxWidgets | rustxwidgets (proposed) | Why |
|---|---|---|---|
| **Data types** | `wxString`, `wxPoint`, `wxSize`, `wxColour` | Rust types: `String`, `Point { x, y }`, `Size { w, h }`, `Colour { r, g, b, a }` | Rust structs encode the same type info as wx classes, with Rust ergonomics |
| **Events** | Event tables + handlers | Closures + propagation | Closures are idiomatic; propagation added for the wxWidgets benefit |
| **Menu IDs** | Integer IDs (`wxID_ABOUT`) | String action names | No ID registry; app-layer enum gives compile-time safety (see STRINGS_VS_INTEGER_IDS.md) |
| **Naming** | `wx*` prefix | Modules | Rust namespaces |
| **Backends** | C++ native (wxMSW, wxGTK, wxOSX) | Native per-platform (GTK, NWG, pancurses) | No C++ FFI; dlopen philosophy |
| **Ownership** | Raw pointers, manual lifetime | `Rc`/`RefCell`, RAII | Memory safety |
| **Threading** | `wxThread`, `wxMutex` | Single-threaded main loop (later: optional) | Simplicity; add only if needed |
| **Resources** | `wxXRC` (XML UI) | None (later: declarative UI) | Deferred |
| **i18n / a11y** | `wxLocale`, accessibility | None (later) | Deferred |

## Resolved decisions

1. **First-class actions — yes, minimal registry in Phase 1.** Add
   `Action { name, enabled, checked, on_activate }` with a registry
   (`register_action`, `set_action_enabled`, `set_action_checked`).
   `MenuItem` keeps the action-name string; the registry is additive and
   gives enable/disable + shared state across menu items, toolbars, and
   keybindings. The standard vocabulary becomes predefined actions.
2. **App lifecycle — keep the current model.** The app drives the loop; no
   wxApp::OnInit/OnExit formalization. Deliberate: a TUI-first toolkit
   doesn't need the framework-owned lifecycle.
3. **Standard dialogs on pancurses — scoped now.** Phase 4 includes the TUI
   versions: file-path prompt (exists), message box (dialog overlay exists),
   color picker (swatch grid), font picker (list). GTK/NWG use native
   dialogs.
4. **BackendApp trait — formalize incrementally.** Each phase extends the
   trait contract; document responsibilities as they land so backends don't
   drift (the per-backend `Menu` types were exactly that drift).
5. **Toolkit testing — headless unit tests.** Add a headless-based test suite
   for the model (menu tree, sizers, event propagation) alongside the phases,
   independent of any rendering backend.

## Phases

### Phase 0 — Conventions
- Adopt the naming table above (no `rx*` prefix; wxWidgets type names).
- Adopt the design principles as a documented policy.

### Phase 1 — Universal menu model (in progress)
- `rustxwidgets::Menu` + `MenuItem` in the core; every backend converts it.
  **Done**: `MenuItem` (Action/Submenu) in the core; pancurses + adapter
  convert the same model; corro builds one `menu_bar()` tree.
- Add item kinds: `Separator`, `Check { checked }`, `Radio { group }`.
  **Done**: all three kinds + `Action { shortcut }` in the model; pancurses
  renders them; adapter has `append_separator`/`append_check`/`append_radio`.
- Add accelerators to the model (`MenuItem::Action { label, action, shortcut }`).
  **Done**: `shortcut: Option<String>` on `Action`.
- **Express menus as a nested tree** via a `menu!` macro — the tree *is* the
  menu, no separate submenu constants:

  ```rust
  menu! {
      "File" => [
          "Open file" => open,
          "Save as"   => save_as,
          "Export"    => [ "TSV" => export_tsv, "CSV" => export_csv, /* … */ ],
          "Exit"      => quit,
      ],
      "Edit" => [ "Cut" => cut, "Copy" => copy, /* … */ ],
  }
  ```

  The macro distinguishes an action (a name) from a submenu (a nested list)
  recursively; it replaces the separate `FILE_MENU`/`EXPORT_MENU`/… constants
  in corro's `menu.rs`.  **Done**: `menu_bar()` in corro's `menu.rs`.
- Add a **standard item vocabulary** — a documented convention of generic
  action strings any app may use, so backends can give them consistent,
  platform-correct behavior without knowing the app:
  - `about`, `exit`, `open`, `save`, `save_as`, `new`, `close`, `help`
  - `cut`, `copy`, `paste`, `find`, `replace`, `undo`, `redo`
  - Rationale: (1) platform semantics (wxID_ABOUT/wxID_EXIT move to the macOS
    system menu — a backend must recognize them by name); (2) backend default
    behavior (`exit` → close window, `open`/`save` → native file dialogs);
    (3) consistency + parity testing.
  - The vocabulary is the *generic subset*; app-specific actions
    (`insert_date`, `balance_books`, …) are NOT in it. This is the
    toolkit-purity line: the toolkit knows about `exit`, not corro's menu.
  **Done**: see `docs/STANDARD_ACTIONS.md`.
- **First-class actions** — `Action { name, enabled, checked }` registry
  (`register_action`, `set_action_enabled`, `set_action_checked`).
  **Done**: `rustxwidgets::Action` + pancurses `ACTION_REGISTRY`.
- **Tests:** menu parity across pancurses/ratatui; item-kind rendering.
  **Done**: `menu_file_parity_with_ratatui`; rustxwidgets lib unit tests for
  item kinds + action registry.

### Phase 2 — Event propagation
- Add a small `Event` enum + `CallbackResult { Handled, Skip }` return; the
  toolkit walks the parent chain when a callback returns `Skip` (mirrors
  `wxEvent::Skip`).
- **Opt-in**: existing callbacks return `()` and are unaffected; only
  callbacks that return `CallbackResult` participate in propagation.
- **Tests:** a parent container intercepts a child's event; skip passes it up.

### Phase 3 — Sizers (layout)
- Replace `BoxWidget`/`Grid` with `Sizer` + `BoxSizer` + `GridSizer` +
  `FlexGridSizer`.
- Add weights, borders, and expansion flags (mirrors `wxSizerFlags`).
- Keep the existing layout backends working (pancurses computes rects).
- **Tests:** layout parity across backends for a fixed window size.

  **Done**: `Sizer`/`SizerChild`/`SizerFlags`/`Align` in the core;
  `PcWidgetKind::Sizer` + `layout_sizer_inner` (box weights, grid cells) in
  the pancurses backend; adapter `create_box_sizer`/`create_grid_sizer`/
  `create_flex_grid_sizer`/`sizer_add`; App-level methods; unit tests
  (`box_sizer_lays_out_children_by_weight`, `grid_sizer_arranges_children_in_cells`).
  **Pending**: migrate the GTK-path consumers (`gui_backend.rs`, `dialogs.rs`)
  from `BoxWidget`/`Grid` to `Sizer` — untestable without the `gui` feature,
  so deferred; `BoxWidget`/`Grid` remain for backward compatibility.

### Phase 4 — Standard dialogs
- Core APIs: file open/save, message box, color picker, font picker
  (mirrors `wxFileDialog`, `wxMessageBox`, `wxColourDialog`, `wxFontDialog`).
- Backends: GTK/NWG use native dialogs; pancurses uses the existing TUI
  prompt/dialog overlay.
- **Tests:** each dialog opens and returns a value in each backend.

### Phase 5 — Client data
- `set_data`/`get_data` on widgets and menu items (mirrors
  `wxWindow::SetClientData`, `wxMenu::SetClientData`).
- **Tests:** attach/retrieve data round-trip.

### Phase 6 — Drawing
- Extend `DrawContext` toward `wxDC`/`wxGraphicsContext` coverage: paths,
  transforms, text measurement, images.
- pancurses `Canvas` remains a no-op (TUI has no canvas) — documented.

### Phase 7+ — Deferred
- Declarative UI (wxXRC-like), threading, i18n, accessibility. Only if a
  concrete need appears.
- **Spreadsheet → Grid refactor.** Keep the `Spreadsheet` widget for now (no
  real harm); when it is refactored toward a generic `Grid` (wxGrid-like),
  move the corro-shaped fields (formula bar, menu text, status text) to the
  app layer. Same toolkit-purity rule as menus.
- **Rust-native GUI backend (optional).** Add iced (retained, closest to the
  widget-tree model) or egui (immediate mode, most popular) as an optional
  backend. Value: no system dependencies (static binaries), macOS + wasm in
  one backend, pure-Rust ecosystem. Cost: non-native look/feel, limited
  accessibility, `wgpu` dependency weight. **No core design change** — the
  universal menu model, node tree, and event system are backend-agnostic; a
  Rust-native backend is just another converter. The only nuance is the
  rendering model (immediate vs on-demand re-emit), which is a backend
  concern. Land it after the universal menu and sizers, so the converter has
  a stable model to target.

## Migration order

1. Phase 0 + 1 (menu model) — unblocks the current pancurses/ratatui parity
   work and the "menu usable by every backend" goal.
2. Phase 3 (sizers) — biggest user-visible API change; do early while the
   widget set is small.
3. Phase 2, 4, 5 (events, dialogs, client data) — additive, low risk.
4. Phase 6+ — as needed.

## Risks

- **Sizer migration** touches every layout call site; do it in one pass with
  the parity tests as the guard.
- **Event propagation** must not break the existing closure model; make it
  opt-in (a `skip` return) so existing callbacks are unaffected.
- **Standard item vocabulary** must not hardcode app menus into the toolkit —
  the vocabulary is a *convention* (documented strings), not a corro-specific
  structure (see AGENTS.md "Toolkit purity").
