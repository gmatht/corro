# Win95 patches (vendored `native-windows-gui 1.0.13`)

Upstream `native-windows-gui` statically imports four Win32 APIs that do not
exist on Windows 95. A single statically-linked missing import makes the
*whole executable* refuse to load on Win95 (`Error Starting Program: linked
to missing export ...`), so this vendored copy replaces them with
Win95-compatible equivalents:

| API | Needs | Used in | Replacement |
| --- | ----- | ------- | ----------- |
| `GetAncestor(hwnd, GA_ROOT)` | Win98+ | `src/win32/mod.rs` message loop | `get_root_window()` walks the parent chain with `GetParent` (Win3.1+) |
| `MonitorFromWindow` | Win98+ | `src/win32/monitor.rs` | primary-display geometry via `GetSystemMetrics(SM_CXSCREEN/SM_CYSCREEN)` |
| `GetMonitorInfoW` | Win98+ | `src/win32/monitor.rs` | same fallback (fills `MONITORINFO` manually, `MONITORINFOF_PRIMARY`) |
| `CreateActCtxW` / `ActivateActCtx` | WinXP+ | `src/win32/mod.rs` `enable_visual_styles()` | function is now a no-op (visual styles don't exist on Win95 anyway) |

Verified with `objdump -p` that the resulting `nwg-test.exe` imports nothing
newer than what the stock Windows 95 `USER32/KERNEL32/GDI32/COMCTL32/OLE32/
SHELL32` export (checked against the DLLs from the Win95 disk image).

## Win95 lies about wide-char API exports

The first build still failed on Win95 (`System class creation failed`) even
though every import resolved: **Windows 95 exports many `*W` functions as
stubs that return 0 without doing anything.** `RegisterClassExW` is one of
them — the import binds fine, the call silently does nothing. Any `*W`
API the app actually executes at runtime must therefore be replaced with
its ANSI (`*A`) variant, which is the real implementation on Win9x.

Conversion applied throughout the core (`src/win32/`):

| Site | W version | A version |
| ---- | --------- | --------- |
| `base_helper.rs` | `GetModuleHandleW` | `GetModuleHandleA` |
| `window.rs` | `CreateWindowExW`, `RegisterClassExW`/`WNDCLASSEXW`, `LoadCursorW`, `DefWindowProcW`, `PostMessageW`, `GetClassNameW`, `SendMessageW` | ANSI equivalents, `WNDCLASSEXA` |
| `window_helper.rs` | `Get/SetWindowTextW`, `SendMessage/PostMessageW`, `Get/SetWindowLongW`, `GetClassInfoExW` | ANSI equivalents |
| `win32/mod.rs` | `GetMessageW`, `PeekMessageW`, `IsDialogMessageW`, `DispatchMessageW`, `PostMessageW` | ANSI equivalents |
| `message_box.rs` | `MessageBoxW` | `MessageBoxA` |
| `controls/label.rs`, `controls/text_input.rs`, `controls/combo_box.rs` | `DrawTextW(DT_CALCRECT)` (line-height measuring in the `WM_NCCALCSIZE` handler) | `GetTextMetricsA` (`tmHeight + tmExternalLeading`). `DrawTextW` is NOT in Win95's working GDI W subset — it is a no-op stub, which zeroed the rect, collapsed the client area to 0 and made every label vanish after any `SetWindowPos` (i.e. after any relayout) |
| `controls/list_view.rs`, app-side `lv_insert_column_direct` | `LVM_INSERTITEMW`, `LVM_SETITEMW`, `LVM_GETITEMW`, `LVM_INSERTCOLUMNW`, `HDM_GETITEMW`, `HDM_SETITEMW` (+ `LVITEMW`/`LVCOLUMNW`/`HDITEMW`) | ANSI message equivalents (`LVM_*A`, `HDM_*A`, `LVITEMA`, `LVCOLUMNA`, `HDITEMA`) — Win95 comctl32 has no W-message support, so columns/rows silently never appeared |
| `events.rs`, `controls/*.rs` (`WM_NCCALCSIZE` lParam deref) | `&mut *(lParam as *mut NCCALCSIZE_PARAMS)` (aligned reference) | `ptr::read_unaligned` / `ptr::write_unaligned` — Win95 USER passes 2-byte-aligned 16-bit-heap pointers; an aligned reference panics in debug builds ("misaligned pointer dereference") |
| `events.rs` (`MinMaxInfo`, `WM_GETMINMAXINFO` lParam) | same aligned deref | same unaligned copy-in/copy-out |

ANSI string handling: `base_helper.rs` gained `to_ansi(&str) -> Vec<u8>`
(ASCII, non-ASCII becomes `?`) and `from_ansi(&[u8]) -> String`; every W
site passes `to_ansi(...).as_ptr()` (NUL-terminated, size is safe) and
wide-character helpers decode with `from_ansi`.

## Win95 can fail `CoInitialize` (E_FAIL) — init must not abort

On a Win95 installation bootstrapped from the pristine `SYSTEM.1ST`
registry (no full OEM setup pass over it), `OLE32.CoInitialize(NULL)`
returns `E_FAIL` (0x80004005) instead of `S_OK`. Upstream `nwg::init()`
treated any CoInitialize failure as fatal, which killed every app right
after window creation (splash → exit; with `panic="abort"` and no panic
hook, the death is completely silent).

The ANSI build uses no COM-dependent feature (no OLE drag & drop, no
shell COM, no common-dialogs via COM), so `src/win32/mod.rs` now maps a
failed `CoInitialize` to `Ok(())` and simply runs without COM. If a
future feature needs COM, re-visit this (e.g. fall back to
`OleInitialize`, or run `DCOM98` in the VM to repair OLE registration).

Left unconverted because the test app never executes them (harmless on
Win95 since they are simply never called): `menu.rs`, `events.rs`
(`DragQueryFileW`), `LoadImageW` call sites, `font.rs`
(`EnumFontFamiliesExW`), `resources/embed.rs` `LoadLibraryW("Msftedit")`
(the null result is discarded).
