# Old-Windows status (Win95 / ReactOS) — 2026-09-15

This file tracks the Win95 + ReactOS bring-up. **Current priority is
Win10/11**: the GUI grid renders blank there too (see §4), so modern
Windows is being fixed first; the notes below are the state of the
legacy investigation when it was parked.

## 1. What works

- **Win95 TUI (`corro-win95.exe`, pancurses/PDCurses): fully interactive.**
  Renders the sheet, arrows navigate, letters edit, round-trip
  B1→B2→A2→A1 verified in the Win95 QEMU VM. Fixes: `KEY_OFFSET`
  gated on `target_family = "rust9x"` (0x100 vs 0xec00), steady-state
  input polls `WaitForSingleObject` on a fresh `CONIN$` (Peek/
  `GetNumberOf` report empty on 9x), `SetConsoleMode(0x18)`.
- **Builds**: `build95.sh` (TUI, `--pancurses`) and `build95-gui.sh`
  (GUI, `/SUBSYSTEM:WINDOWS,4.0`, `WinMainCRTStartup`) produce
  `dist/corro-win95.exe` / `dist/corro-win95-gui.exe` (`gcorro.exe`
  inner name via argv[0] dispatch). `.cargo/config.toml` stays
  CONSOLE; GUI linker flags live only in `build95-gui.sh`.
- **Win95 constraints learned the hard way**: no env vars on 9x
  (`GetEnvironmentStringsA` never reaches Win32), `std::fs` broken
  (CreateFileW stub → OS error 120) so all diagnostics use raw
  `CreateFileA` shims (`mark95`/`append_file_raw`), default panic
  hook dies in TLS so `install_raw_panic_hook95()` writes
  `c:\panic95.log` TLS-free, `InitCommonControlsEx` faults on 9x
  (skipped), blocking `GetMessageW` spuriously returns `WM_QUIT`
  (replaced with `PeekMessageW` + `MsgWaitForMultipleObjects` poll).

## 2. Win95 GUI: blocked on a GPF (deferred to 0.7.1)

- `GCORRO.EXE` opens its window (poll loop works, `v-gui9.png`) but
  GPFs on the **3rd** message dispatched through the ANSI
  subclass-emulation trampoline in
  `vendor/native-windows-gui/src/win32/window.rs`. No Rust panic
  (`panic95.log` empty) — a hard fault inside message dispatch.
- Next step (not started): rebuild with the `tm <hex>` message-id log
  (`SUBCLASS_M95 < 8`), restage, boot, transcribe the crashing
  message id, fix the trampoline/parenting OOB.

## 3. ReactOS 0.4.16 GUI: window opens, then dies (~1 min), grid blank

Test loop: QEMU LiveCD + `tools.iso` carrying `GCORRO.EXE`, QMP
`:4446`, screenshots `dist/ros-*.png`, fresh `c:\gcorro.log` per run
(`del` first — stale logs caused real confusion; always `del`,
always verify `fix<N>` identity, always `quit`+fresh-boot after
restaging the ISO).

- **Startup death → root-caused and fixed.** NWG's `WM_NCCALCSIZE`
  vertical-centering handlers (`text_input.rs`, `label.rs`: measure
  font, shrink client rect) recurse **infinitely** against ReactOS's
  EDIT/STATIC proc — thousands of `ni`/`nl` marks with zero `WM_SIZE`
  in between, ending in stack overflow / silent death. The
  `WM_SIZE`→`SetWindowPos` loop theory was falsified first (no
  `sz`/`sk` during the flood).
- **Current diagnostic state (TEMPORARY, rust9x-gated):** the
  NCCALCSIZE handler bodies early-return `None` (marks kept), build
  identity `fix2`. With that, the **"corro 0.7.0" window opens and
  idles healthily** (`sk` suppression, periodic `layot` passes).
- **Blocker R1 — spontaneous death while idling.** Window gone
  between ~50s and ~90s after launch (`ros-am5-0..8`), absent from
  `tasklist`, no panic, log tail ends cleanly. Suspects: timer-driven
  path (caret blink/repaint), `notify` watcher, poll loop. Needs
  per-callback heartbeat marks to catch the killer.
- **Blocker R2 — grid never paints.** Client area stays blank white
  (menu frame exists, sheet canvas doesn't render). **This also
  reproduces on Win11**, so it is now the top-priority bug and is
  being chased as a modern-Windows issue, not a ReactOS one.
- **Blocker R3 — the NCCALCSIZE no-op is a hack.** Vertical text
  centering is disabled on all rust9x builds. The real fix must
  converge with ReactOS's proc (adjust-once / OS detection) and be
  re-validated on Win95 (§2).
- Harness friction (not product): `scripts/qmp.py` keymap is
  UK-tuned for the Win95 guest; ReactOS is US, so `|` and `"`
  mistype — use screenshots, avoid pipes/quotes.

## 4. Win10/11 GUI: grid blank — ROOT-CAUSED AND FIXED (2026-09-15)

No bisect was needed: direct instrumentation found it in one session.

- **Symptom**: `corro.exe`/`gcorro.exe` window opens (menu bar draws)
  but the client area stays blank white — no formula bar, no grid.
  Reproduced under Wine 9 + Xvfb (which renders Notepad perfectly,
  so the harness is sound) with `x86_64-pc-windows-gnu --features
  gui` builds screenshotted via `import`.
- **Red herrings falsified**: misparenting (the Win32 tree is perfect:
  13 HWNDs, all visible, sane rects), missing subclass handlers,
  broken DrawContext. Final geometry was correct yet nothing
  painted — not even native EDIT/STATIC controls, and the canvas
  never received `WM_PAINT`.
- **Root cause — stale zero/negative sizes win the startup race**
  (`rustxWidgets/rswidgets/src/backends_nwg_adapter.rs`):
  1. `pump_events()` is a **no-op on Windows**, so the entire setup
     message storm queues up and is drained only when `run()`
     starts — degenerate early sizes are processed LAST.
  2. `BoxWidget::layout()` at width 0 computes `fixed = w - 10 =
     -10` and calls `SetWindowPos(child, …, -10, …)`; Wine/Windows
     clamps the size to 0 (child crushed invisible).
  3. The "airtight cascade" synthetic `SendMessageW(WM_SIZE)` packs
     the negative size as unsigned, so `-10` wraps to **65526** and
     nested children are flung ~65000px off-screen (observed:
     `child=0x1007e at 65513,5`, entry `65442x18`).
  4. End state is 0-size windows: `InvalidateRect` on a 0x0 window
     yields an empty update region, so `WM_PAINT` never fires and
     nothing ever re-invalidates — white screen forever.
- **Fix (kept, principled — not TEMP):** clamp `w`/`h` at
  `layout()` entry; clamp spans before `SetWindowPos`; send the
  synthetic `WM_SIZE` only when `cw > 0 && ch > 0`; drop
  zero-size `WM_SIZE` in the toplevel, box, and scrolled-window
  handlers (a zero-size window has nothing to lay out; the next
  real size repairs). Verified: `DRAW_CALLBACK` fires, full grid
  (menu, formula bar, headers, cells, selection, TOTALs) renders
  under Wine (`dist/wine64-10.png`, `wine64-11.png`).
- All rust9x TEMP diagnostics remain `#[cfg(target_family =
  "rust9x")]`-gated and compile out of modern-Windows builds.
- Still open: real Win11 eyeball check (Wine proves the Win32
  path; only font coverage differs — Wine shows tofu for some
  margin glyphs), GUI keyboard/mouse interaction test.
