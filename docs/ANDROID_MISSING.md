# What the Android backend is missing (vs GTK and NWG)

Android is not a peer of the GTK (Linux) and NWG (Windows) backends: it is
the *same* widget tree, the same `gui_backend::run_gui` pipeline, the same
menus and the same draw code, rendered through a JNI adapter. Every gap below
is therefore one of two kinds:

* **Wired through the host but never implemented.** corro's Java shims and the
  adapter have the plumbing; the method body is empty or returns a "not
  available" sentinel. These are pure work: implement the body and the feature
  is live.
* **No plumbing at all.** Nothing in the Java host or the adapter even has a
  place to put the value. These need a new native export *and* a Rust method.

`docs` in this repository is authoritative for the code; this file is the
gap list and its status. It is derived from three diffs:

1. `rustxWidgets/rswidgets/src/backends_android_adapter.rs` method sets vs
   `backends_nwg_adapter.rs` and `backends_gtk_adapter_impl.rs`.
2. Every no-op / `None` / `false` / `Err` body in the Android adapter and in
   `rswidgets/src/common.rs`.
3. Every `#[cfg(target_os = "android")]` arm in `src/gui/`, and every
   "unavailable" status string the GUI can report.

Status column: **done** = implemented on Android; **stub** = a method exists
and compiles but does nothing; **absent** = no plumbing at all.

## 0. Build-breaking: `common.rs` forwards to methods the adapter lacks

`rswidgets/src/common.rs` (the cross-backend wrapper every corro widget call
goes through) forwards `Label::set_fixed_width`, `Label::set_margin_start`,
`Label::set_margin_top`, `Label::set_xalign` and `Window::resize` to the
platform adapter. GTK and NWG implement all five. **Android implements none of
them, so `corro_android` does not compile at all** — `error[E0599]: no method
named set_xalign found for struct android_adapter::Label`. This is gap #1
because nothing else can be tested until it is fixed.

| Method | GTK | NWG | Android | Status |
|---|---|---|---|---|
| `Label::set_fixed_width` | yes | yes | no | **done** |
| `Label::set_margin_start` | yes | yes | no | **done** |
| `Label::set_margin_top` | yes | no | no | **done** |
| `Label::set_xalign` | yes | yes | no | **done** |
| `Window::resize` | yes | no | no | **done** |

## 1. Widget-level API parity (`backends_android_adapter.rs`)

### 1.1 Expansion and size

Android has no `hexpand`/`vexpand`; the adapter already translates them into
`LinearLayout` weight, but only for `BoxWidget`, `DropDown`, `Overlay` and
`TextView`. Every other widget is missing them, and `set_size_request` (a
minimum width/height, honoured by GTK and NWG) is missing nearly everywhere.

| Method | GTK | NWG | Android | Status |
|---|---|---|---|---|
| `Button::set_hexpand` / `set_vexpand` | yes | no | no | **done** |
| `Button::set_size_request` | yes | yes | no | **done** |
| `Button::set_visible` | yes | no | stub (no-op) | **done** |
| `Button::add_class` / `remove_class` / `set_font_style` | yes | yes | no | **done** (font style + text style) |
| `CheckButton::set_hexpand` / `set_vexpand` / `set_size_request` / `set_visible` | yes | no | no | **done** |
| `CheckButton::get_label` / `set_label` | no | yes | no | **done** |
| `DropDown::set_size_request` / `set_visible` / `set_offset` / `active` | yes (`set_visible`) | yes | no | **done** |
| `Grid::set_hexpand` / `set_vexpand` / `set_size_request` / `set_visible` | yes | no | no | **done** |
| `Label::set_hexpand` / `set_vexpand` / `set_size_request` | yes | no | no | **done** |
| `RadioButton::get_label` / `set_label` / `grab_focus` / `set_hexpand` / `set_vexpand` | yes (`grab_focus`) | yes (`get_label`, `set_label`, `set_hexpand`, `set_vexpand`) | no | **done** |
| `TextView::set_visible` | yes | no | no | **done** |

### 1.2 Canvas, scrolling and layout

`ScrolledWindow` is an inert `FrameLayout` on Android, so *all* of NWG's
scrolling surface is stubbed out. The sheet is panned by the gesture machine
instead, which is the right substitute on a phone — but the widget-level
contract is still missing and a host that wants a real scroll view (the ODS
export preview, the balance picker) has nothing to ask for.

| Method | GTK | NWG | Android | Status |
|---|---|---|---|---|
| `Canvas::set_hexpand` / `set_vexpand` / `set_margin_start` / `set_margin_top` | yes | `set_hexpand`/`set_vexpand` no-op | no | **done** |
| `Canvas::set_visible` | yes | yes | **stub** (`{}`) | **done** |
| `Canvas::grab_focus` / `set_can_focus` | yes | yes | **stub** (`{}`) | **done** |
| `ScrolledWindow::set_child` | yes | yes | no | **done** |
| `ScrolledWindow::set_policy` (never/always/auto, h+v) | yes | yes | **stub** | **done** (real `ScrollView` + bars) |
| `ScrolledWindow::set_hexpand` / `set_vexpand` | yes | yes | **stub** | **done** |
| `ScrolledWindow::scroll_to` | yes (stub) | yes | **stub** | **done** |
| `ScrolledWindow::on_scroll` (user-scrolled notification) | yes (stub) | yes | **stub** | **done** |
| `Overlay::set_child` | yes | yes | yes | ok |

### 1.3 Entry and text editing

This is the sharpest gap. `common::Entry` documents GTK and NWG as "report a
real caret", and explicitly lists Android among the backends that do not, so
`get_position` falls back to an internally tracked value that is never updated
by the user. Consequence on a phone: **no arrow keys inside a cell edit, no
caret-aware backspace, no caret-aware delete**. The soft keyboard's own ⌫
works inside the `EditText`, but corro's `edit_backspace` / `edit_delete`
formula-aware rules never run, because they are key handlers.

| Method | GTK | NWG | Android | Status |
|---|---|---|---|---|
| `Entry::get_position` | real | real | **stub** (`None`) | **done** (Selection) |
| `Entry::set_position` | real | real | **stub** (`{}`) | **done** |
| `Entry::on_key_raw` | real | real | **stub** (`{}`) | **done** (via the key dispatch) |
| `Entry::connect_changed` | yes | yes | no | **done** |
| `TextView::set_editable` | no | yes | no | **done** |
| `TextView::get_buffer` | no | yes | no | **done** (text as buffer) |
| `TextView::append_text` | yes | yes | yes | ok |

### 1.4 Dialogs

corro builds a dozen modal dialogs (`src/gui/dialogs.rs`). They render
(a `Dialog` is a `FrameLayout` in a `Dialog` window) but cannot be run
nested: `Dialog::run` does not exist, so the dialog code that expects a
nested loop has nothing to call.

| Method | GTK | NWG | Android | Status |
|---|---|---|---|---|
| `Dialog::run` | yes | yes | no | **done** (Activity-runnable modal) |
| `Dialog::layout_dialog` | yes | yes | no | **done** |
| `Dialog::set_visible` | yes | yes | no | **done** |
| `Dialog::set_size_request` | no | yes | no | **done** |
| `Dialog::add_button` (label, response id) | yes | yes | yes | ok |
| `Dialog::get_content_area` | yes | no | no | **done** |

### 1.5 Menus, window and app lifecycle

| Item | GTK | NWG | Android | Status |
|---|---|---|---|---|
| `MenuBar::popup_submenu_by_mnemonic_at` | stub | yes | **stub** (`false`) | **done** (real popup) |
| `MenuBar::activate_submenu_item_by_mnemonic` | yes | yes | **stub** | **done** |
| `MenuBar::handle_mnemonic_key` / `handle_menu_key` | yes | yes | **stub** | **done** |
| `MenuBar::menu_active` / `menu_close` | yes | yes | **stub** | **done** |
| `Canvas::on_click_button` (right-click) | yes | **stub** | **stub** | **done** (long press) |
| `Canvas::on_motion` / `on_release` | yes | **stub** | **stub** | **done** (gesture machine) |
| `Canvas::screen_origin` | yes | yes | **stub** (`None`) | **done** |
| `Window::set_title` | yes | yes | **stub** (`{}`) | **done** (Activity title) |
| `Window::set_default_size` | yes | yes | **stub** (`{}`) | **done** (LayoutParams) |
| module `quit_main_loop` | yes | yes | no | **done** (`Activity.finish`) |
| module `timeout_add_once` / `timeout_add_repeating` | yes | yes | no | **done** (`Handler.postDelayed`) |
| module `create_application` | yes | no | n/a | not applicable |
| module `create_fixed` | yes | no | n/a | not applicable |
| module `create_spreadsheet` | yes | no | n/a | not applicable |
| module `pump_main_context` | yes | no | n/a | not applicable |
| module `create_dialog_button` | no | yes | n/a | not applicable (uses `Dialog::add_button`) |

## 2. File dialogs — the largest single gap

GTK has `open_file` / `open_file_filtered` / `save_file` /
`save_file_filtered`; NWG has the same four. **Android has none of them**, and
`core.rs` returns `Err` for all four, so:

* **File → Open** shows a "Save" status, opens nothing.
* **File → Save As** returns `None`, so the save is silently dropped.
* **File → Export** (CSV/ODS/PDF/…) cannot complete.

Android's answer is the Storage Access Framework
(`ACTION_OPEN_DOCUMENT` / `ACTION_CREATE_DOCUMENT`), which is *asynchronous*:
`startActivityForResult` returns immediately and the URI arrives later in
`onActivityResult`, on the UI thread. That is the whole reason this is not a
three-line port. The implementation is a `FilePicker.java` shim that starts the
intent, plus a Rust side that pumps the looper until a result lands (so the
existing synchronous `Ok(Option<String>)` signature is preserved), plus a
`File` guard that holds persistable read/write permission for the returned URI
and can `openInputStream` / `openOutputStream` it.

| Feature | GTK | NWG | Android | Status |
|---|---|---|---|---|
| Open file | yes | yes | **absent** | **done** (SAF `ACTION_OPEN_DOCUMENT`) |
| Open file with extension filters | yes | yes | **absent** | **done** (`EXTRA_MIME_TYPES`) |
| Save file | yes | yes | **absent** | **done** (SAF `ACTION_CREATE_DOCUMENT`) |
| Save file with filters + default name | yes | yes | **absent** | **done** |
| Read a `content://` URI in corro's `io` layer | n/a | n/a | **absent** | **done** (`UriStream`) |
| Write a `content://` URI in corro's `io` layer | n/a | n/a | **absent** | **done** |
| Open document handed to the app from outside (a VIEW intent) | no | no | **absent** | **done** |

## 3. Clipboard

Corrected after checking the source rather than assuming. The GTK and NWG
*GUI* backends do **not** touch a platform clipboard: no file in
`rustxWidgets/rswidgets/src/` mentions `gtk_clipboard`, `OpenClipboard` or
`SetClipboardData`, and the gtk_dynamic_loader binds no clipboard symbol. The
GUI's copy/cut/paste is in-process Rust state — `GuiState::clipboard` for
GTK and NWG alike (`src/gui/actions.rs` `"copy"` / `"cut"` / `"paste"`), and
that already works on Android through the menu strip.

The *system* clipboard does exist, but in the two places that reach a
platform one:

* the **TUI** (`src/ui/mod.rs`), which shells out to `xclip` / `pbcopy` /
  `clip` / `Get-Clipboard`; and
* the **pancurses** backend (`set_clipboard_text`, via an OSC 52 escape).

Android is in neither category, and it cannot be: there is no `xclip`, and a
phone has no terminal. The equivalent is `ClipboardManager`, and a keyboard
bridge (step #4 of §4) is what makes `Ctrl+C` reach it. So the gap is real
for Android, but the framing is "Android has no way to reach a clipboard
where TUI and terminal do", not "GTK and NWG do this and Android does not".

| Feature | GTK | NWG | TUI | pancurses | Android | Status |
|---|---|---|---|---|---|---|
| In-app copy/cut/paste | yes | yes | yes | yes | yes (already worked) | ok |
| System clipboard write | no | no | yes | yes (OSC 52) | **absent** | **done** |
| System clipboard read (paste from another app) | no | no | yes | no | **absent** | **done** |
| Cell range copy (multi-cell) | no | no | no | no | no | not applicable — same on every backend |

Reading the system clipboard on paste is what makes an Android copy usable
outside corro at all, so it is not a "nice extra": without the write side the
copy is invisible, and without the read side a value copied in another app can
never reach a cell.

## 4. Keyboard input — the second largest gap

`gui_backend::handle_key` is a single shared implementation with every
desktop binding in it (arrows, Home/End, PageUp/Down, Tab, Escape, Delete,
Backspace, F1/F2/F3, Ctrl+C/X/V/Z/Y/O/S/Q, Alt+mnemonic). On Android it is
**almost entirely unreachable**: `dispatch_canvas_key` exists and its registry
is live, but nothing in the Java host calls it, and `Entry::on_key_raw` is an
empty body. The only key that reaches corro today is Enter, via
`CorroEditorAction` / `CorroKeyListener`. Everything in the table below is
therefore dead on a phone.

| Binding | Handler | Android before | Status |
|---|---|---|---|
| Cursor ←↑→↓ | `handle_key` 2104-2142 | unreachable | **done** |
| Shift+arrows (extend selection) | 2106, 2115, 2127, 2137 | unreachable | **done** (Shift is a keycode on the bridge) |
| Home / End | 2143-2165 | unreachable | **done** |
| Page Up / Page Down | 2166-2186 | unreachable | **done** |
| Tab | 2076-2082 | unreachable | **done** |
| Escape (cancel edit / collapse selection) | 2083-2090 | unreachable | **done** |
| Delete / Backspace in Normal mode | 2171-2185 | unreachable | **done** |
| Delete / Backspace in Edit mode | `handle_edit_key` 2254-2280 | unreachable | **done** |
| Caret ←/→ inside an edit | `handle_edit_key` 2281+ | unreachable | **done** |
| F1 (help) / F2 (edit) / F3 (agg picker) | 2038-2062 | unreachable | **done** |
| 1-7 hotkeys in the agg dropdown | 1994-2010 | unreachable | **done** |
| Ctrl+Q quit, Ctrl+O/S/Z/Y/X/C/V | window key handler 6024-6140 | unreachable | **done** |
| Alt+mnemonic menu activation | 6078-6081 | unreachable; `_` is stripped from labels | **done** (Alt path) |
| Extrapolate-modal keys (Esc/Enter/arrows) | `handle_extrapolate_key` | unreachable — the modal can be entered but never committed or cancelled | **done** |
| Revision-browse keys (←/→/Enter/Esc) | 1921-1966 | unreachable | **done** |
| Return / Enter (commit, move down) | 2064-2074 | **works** (IME Done + `CorroKeyListener`) | ok |

Soft-keyboard typing already works through `CorroTextWatcher`; that path is
unchanged. What the new bridge adds is the *key* stream, which is what the
whole table above hangs off. Sources, in priority order:

1. `SheetView.onKeyDown` — hardware keyboard, `adb shell input keyevent`, and
   the on-screen keyboard's own arrow/delete keys, which Android delivers as
   `KeyEvent`s to the focused view when the IME does not consume them.
2. `MainActivity.onKeyDown` — the Activity-level fallback so a key that no
   view claims still reaches corro (Ctrl+Q with the soft keyboard up).
3. `CorroKeyListener` on the formula entry — extended from Enter-only to the
   full key set, so an EditText holding focus routes arrows/backspace through
   corro's edit model instead of the platform's.

Android `KeyEvent` keycodes are translated to GDK values in one table
(`android_keycode_to_gdk`) so `handle_key` needs no second implementation.

## 5. Context menu (right click)

`gui_backend::open_sheet_context_menu` is the desktop right-click path
(GTK: `on_click_button` + `popup_submenu_by_mnemonic_at`). On Android
`Canvas::on_click_button` is an empty body and `popup_submenu_by_mnemonic_at`
returns `false`, so the function always reports `SHEET_MENU_UNAVAILABLE`
("Sheet menu unavailable", `src/core/state.rs:24`) and puts that in the status
line. Worse, the long press is *already consumed* by the gesture machine to
arm range selection, so it cannot be repurposed for a menu as-is.

| Feature | GTK | NWG | Android | Status |
|---|---|---|---|---|
| Right-click opens the cell menu | yes | stub | **absent** | **done** (long press) |
| Mouse right-click (emulator, stylus) | yes | stub | **absent** | **done** (`BUTTON_SECONDARY`) |
| Menu pops up at the cell, not the corner | yes | yes | **absent** | **done** |
| Mnemonic within a popup menu | yes | yes | **absent** | **done** |

## 6. Anything that is *not* a gap

Listed so the list is not mistaken for a wish list; these are absent on GTK
and NWG too, or are not applicable to a phone:

* `create_spreadsheet` (GTK-only widget), `create_application`, `create_fixed`,
  `pump_main_context` — GTK-only, no phone analogue.
* Printing, tooltips, `draw_rgba_image`, drag & drop of cell ranges, RTL/i18n
  layout, accessibility services, double-click-to-edit — absent on *every*
  backend, so Android is at parity.
* Selection: long-press-then-drag is Android's substitute for shift-click and
  is arguably better than GTK's shift-click on a touch screen. Implemented, not
  missing.
* Pinch/double-tap zoom, finger-drag pan: Android-only extras, not gaps.
* Window title and window size are the Activity's, not a toplevel's, so
  `set_title` / `set_default_size` are implemented against the Activity
  rather than being dropped.

## 7. How this was verified

`cargo build --target x86_64-linux-android` and
`cargo build --target aarch64-linux-android` from `android/corro`, clean, no
warnings from the Android adapter. `build_apk.sh` packages the result; the
Java shims are compiled by the same script with `javac`/`d8`.
