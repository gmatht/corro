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
three-line port. corro's file dialogs are synchronous — they are called from
a menu action inside a native call, on the UI thread, with no opportunity to
return and be called back — so the thread has to wait, and the wait is
signalled by the platform's own `onActivityResult` delivery.

Two further things the naive port gets wrong, and which decide whether the
feature works at all:

* **A `content://` URI is not a path.** `std::fs` cannot open one and
  `Path::exists` is false for it, so handing the result on unchanged loads
  nothing. The *import* direction therefore copies the document once into
  private storage and returns a real path, which every existing loader works
  with unchanged. The *save* direction keeps the URI and routes it through
  the content resolver, because there is no other handle to write to — hence
  `SaveTarget` (path or document) rather than a `PathBuf`.
* **The permission lapses.** The picker's grant is per-document and does not
  survive a reboot unless `FLAG_GRANT_PERSISTABLE_URI_PERMISSION` was
  requested and taken; without it the next save fails with a
  `SecurityException` that is very hard to attribute later.

The honest cost, documented at the call site: a path save is atomic (temp
sibling + rename) and a document save is not, because a document provider
exposes a single stream with no sibling to rename over.

| Feature | GTK | NWG | Android | Status |
|---|---|---|---|---|
| Open file | yes | yes | **absent** | **done** (`CorroFile.open`, `ACTION_OPEN_DOCUMENT`) |
| Open file with extension filters | yes | yes | **absent** | **done** (`mimes_from_filters` → `EXTRA_MIME_TYPES`) |
| Save file | yes | yes | **absent** | **done** (`CorroFile.create`, `ACTION_CREATE_DOCUMENT`) |
| Save file with filters + default name | yes | yes | **absent** | **done** (`Intent.EXTRA_TITLE`) |
| Read a picked document | n/a | n/a | **absent** | **done** (`materialize_document` → a real `PathBuf`) |
| Write to a picked document | n/a | n/a | **absent** | **done** (`SaveTarget::Document` → `write_document`) |
| Open document handed to the app from outside | no | no | **absent** | **done** (VIEW/SEND intent filters + `onNewIntent`) |

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

* `cargo ndk -t x86_64 build` in `android/corro` — clean, no warnings from
  the Android adapter or the app-level glue.
* `cargo ndk -t arm64-v8a build` — clean, and `build_apk.sh arm64-v8a` produces
  a signed APK with a real `ELF 64-bit LSB shared object, ARM aarch64` and all
  25 JNI exports in it.

  This took two fixes, and both are worth recording because the first one made
  the second invisible:

  1. **`build_apk.sh` never actually built arm64.** Its case statement set
     `TARGET="aarch64"` and passed that to `cargo ndk -t`, but cargo-ndk wants
     the *ABI* name (`arm64-v8a`), not the rustc triple's first component. The
     build line piped cargo through `tail -1`, so the resulting
     `invalid value 'aarch64' for '--target <TARGET>'` was discarded and the
     script exited silently. The correct line, and the `tail` removal, are
     both in that commit.
  2. **With a real arm64 build finally attempted, 76 type errors appeared** in
     `gtk_dynamic_loader`. Its `extern "C"` signatures spelled C strings
     `*const i8`, while the wrappers pass `CString::as_ptr()`, which is
     `*const c_char` — and `c_char` is `i8` on x86 and ARM but `u8`
     elsewhere. So the crate could not typecheck on *either* Android ABI; it
     only compiled on x86_64 because there the two happened to agree, which
     is why the emulator hid it. The signatures are now `*const u8` and the
     call sites cast. The pointee's signedness is not part of any ABI's
     calling convention, so the FFI boundary is unaffected; the reasoning is
     recorded at the top of `symbols.rs` where the next reader will hit it.
* `cargo build --features gui` (desktop, Linux/GTK) — clean. Several changes
  here are additions to `common.rs` and the GTK3 loader rather than to the
  Android adapter, and those had to keep the desktop building.
* `cargo build --features gui-core` — 19 errors, all pre-existing and
  identical on the base commit (`common::Canvas`, `common::Window`,
  `backends::init` unresolved when no GUI backend is selected).
* Merged into `main` brought one new instance of the §0 class: `995b34cc`
  added `BoxWidget::set_size_request` to `common.rs` for NWG, and Android
  did not have it, so the Android build broke again with the same
  `E0599` shape. Implemented, along with the sibling
  `set_vexpand` / `set_visible` that §1.1 lists. Worth recording as a
  pattern rather than an anecdote: **any new method on `common.rs` needs
  the Android adapter updated in the same change, and the compiler will
  say so, which is the one useful property of this being a hard error
  rather than a silent no-op.**
* `cargo test --features gui` — 824/824 unit tests pass. The
  `gui_edit_parity` integration suite has 15–17 failures **on the base commit
  too** (it needs a headless X display); the same set fails before and after
  this work, and this branch is one test better than base.
* `./build_apk.sh` — produces a signed APK. Checked in the artifact: all 19
  `#[no_mangle]` JNI exports are present in `libcorro_android.so`, all 11
  host shim classes are in the dex, and the packaged manifest carries the
  VIEW and SEND intent filters.
* The Java is also compiled directly against `android-34` with `javac`, which
  is what caught two real API mistakes: `MessageQueue.next()` is not public
  API, and `OnScrollChangeListener` has a 5-argument shape, not the 7 one a
  more familiar widget has. Both would have been runtime failures on a
  device.

## 8. What is still not implemented, and why

Kept short so this file is not mistaken for "nothing remains":

* `create_spreadsheet`, `create_application`, `create_fixed`,
  `pump_main_context`, `create_dialog_button` — GTK- or NWG-only widget and
  lifecycle concepts with no phone analogue. `ScrolledWindow` plus a
  `Canvas` is corro's answer, and it works.
* `DropDown::diagnostics` / `has_size_request_symbol` / `is_gtk4` — GTK4
  loader introspection, for telling which symbols a dlopen'd libgtk
  actually has. Android binds no symbols, so there is nothing to introspect.
* `Dialog::mark_destroyed` — a GTK lifetime guard. Android's `dismiss` and
  the `GlobalRef` keep-alive cover the same ground.
* Window `set_default_size` / `resize` are honoured against the root layout,
  but a phone's window size is decided by the system window manager and the
  device, so these are requests, not guarantees. That is the platform's
  answer, not a gap in the implementation.
