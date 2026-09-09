# Menu item IDs: strings vs integers

**Status:** design decision — rswidgets uses **string action names** as the
universal menu-item ID.

## Context

rswidgets is moving to a single, backend-agnostic menu model
(`rswidgets::Menu` + `rswidgets::MenuItem`) that every backend renders
with its own native widgets. A menu item needs an **ID** — the identity used
for dispatch, state queries (`enable`, `is_checked`), and cross-backend parity
testing. The two candidate ID types are:

- **String** — `MenuItem::Action { label, action: "open" }`. The action name
  *is* the ID.
- **Integer** — `MenuItem::Action { label, id: 2001 }` (the wxWidgets model:
  `wxID_ABOUT`, `#define ID_SOMETHING 2001`).

The ID is **per menu item, not per backend**: one item has one ID, and it is
the same ID on every backend. A backend may *translate* that ID into its
platform's native identity internally, but that translation is private plumbing,
never a second ID scheme.

## Native ID preferences by backend

| Backend | Native menu identity | Prefers |
|---|---|---|
| GTK / gtk4-rs | `GAction` **names** (strings) | string |
| NWG (Windows) | `WM_COMMAND` **IDs** (ints) | int |
| Android | `MenuItem.getItemId()` (ints) | int |
| pancurses | pass-through callback | either |
| zork / wasm / ratatui / headless | pass-through | either |

## Per-backend effect

**GTK / gtk4-rs** — *strings are free, ints cost.* GTK's `GMenu` model addresses
items by `GAction` name (`"app.open"`). With string IDs the action string maps
directly to the GAction name — zero conversion. With integer IDs the backend
must synthesize a name per int (`format!("app.action_{id}")`) and keep a
reverse map. (wxGTK does exactly this internally: wxID → GAction name.)

**NWG (Windows)** — *ints are free, strings cost.* Windows menus dispatch
`WM_COMMAND` with an integer command ID; accelerators use ints too. With
integer IDs the item's ID *is* the command ID — zero conversion. With string
IDs the backend allocates an int per string and keeps a
`HashMap<String, u16>` for dispatch. (The classic wxMSW model: wxID → command
ID.)

**Android** — *ints are free, strings cost.* `onOptionsItemSelected` delivers
an int item ID. Same story as Windows: strings need a map, ints don't.

**pancurses** — *neutral.* The backend fires `menu_action_callback(String)` —
it passes the ID through. With strings the callback carries the string
directly; with ints it carries the int and the app matches on it. Either
works; strings are slightly more self-documenting in the callback.

**zork / wasm / ratatui / headless** — *neutral.* The model is data; dispatch
is by index or callback. The ID type does not change the backend's work.

## The tradeoff

It is not about performance — a `HashMap<String, u16>` is trivial. It is about
**which backends you optimize for**:

- **Strings** (chosen): GTK is free; Windows/Android pay a small internal map.
  Self-documenting, typo-safe at the app layer (the app can define an enum
  that maps to strings), and the primary backends (GTK, pancurses) are
  string-native.
- **Integers** (wxWidgets): Windows/Android are free; GTK pays a
  name-synthesis map. Requires ID bookkeeping at the app layer (the
  enum/`#define` problem: manual allocation, collision risk, split
  definitions).

Either choice leaves exactly one family of backends doing a conversion;
strings just pick the family that is currently least important.

## Why strings win for rswidgets

1. **The two backends that matter most today — GTK on Linux and pancurses for
   the TUI — are both string-native.** No conversion on the primary paths.
2. **No ID registry.** The action name is self-documenting and needs no
   allocation, collision management, or `#define` bookkeeping.
3. **The app layer gets compile-time safety anyway.** The app defines its own
   enum (`enum CorroMenuAction { Open, SaveAs, … }`) and maps it to strings
   (`impl From<CorroMenuAction> for &str`). Dispatch is exhaustive and
   typo-free without the toolkit needing integer IDs.
4. **The string is the shared contract for parity testing.** The walk test
   compares pancurses vs ratatui by string (menu labels, formula bar), so the
   ID type that tests see is the string.

## The three-layer model

```
app enum (compile-time safety)  →  action string (the item's ONE ID)  →  backend's native handle (private)
```

- The **app enum** is the compile-time front door; it never reaches backends.
- The **action string** is the universal ID: one per item, identical on every
  backend, used for dispatch, state queries, and tests.
- The **backend's native handle** (command ID, GAction name, item ID) is
  private translation inside each backend's `create_menubar` conversion.

This mirrors wxWidgets: `wxMenu` is universal, and each backend (wxMSW,
wxGTK, wxOSX) does its own internal ID handling underneath.
