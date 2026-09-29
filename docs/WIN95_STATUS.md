# corro's GUI on Windows 95 — where it stands

Status: **starts, creates its window and canvas, paints once at 1×1, then
aborts.** The grid is not visible. This is a bring-up note, not a claim that
it works.

Raw evidence from the first successful instrumented run is kept beside this
file:

* `win95-gcorro-first-run.log` — the app's own diagnostic log
  (`gcorro.log`, read back off the guest disk with guestfish)
* `win95-gcorro-panic.log` — the panic hook's record (`gcorro.log.panic`)

Reproduce with `./run95-gui.sh --no-build --no-shutdown` and then

    export LIBGUESTFS_BACKEND=direct
    guestfish --ro -a "$WIN95_SESSION" -m /dev/sda1 download /gcorro.log /tmp/g.log

## What works

The Win95 bring-up path itself is sound:

* The exe loads. `winmain` is the first marker, and everything after it runs,
  so the VC6 GUI CRT, the `WinMain` entry and the `unicows.dll` import are all
  fine. **`unicows.dll` had to be added to the image** — it was absent, and
  without it the loader fails before any window exists.
* The GUI toolkit initialises: `coinit`, `preicex`, `icex`, `prewinclass`,
  `wclass`, `tabs`, `canvas`, `winpost`.
* The window is created and `present()` reports a real client area of
  **620×420** (`pres 0000026c 000001a4`).
* **The canvas really paints.** `cvpaint` fires three times, `gdidc`/`gdibm`
  return valid objects, the draw callback is registered (`hascb 1`), and
  `BitBlt` succeeds. GDI is not the problem.

## What fails

Two distinct problems, in order.

### 1. The layout hands every child a zero width

    layot 00000000 000186a0     <- box 1: width 0
    layot 00000000 00030d40     <- box 2: width 0
    layot 0000026c 00049584     <- box 3: width 620 (the root)
    ...
    preSW 000007d0 00000000     <- every child positioned at width 0
    wsviz 00000000 00000000     <- ...and with WS_VISIBLE clear

The root box gets the correct 620. The boxes created *before* `present()` are
laid out from a window that does not exist yet, so NWG reports them as 0×0, and
that zero is what their children inherit. This is the same shape as the
ReactOS bug fixed in `11c34262` — a child that has never been shown reads as
hidden, is given zero size, and is then skipped by the very `set_window_pos`
that would have shown it — but here it never repairs, because the next thing
that happens kills the process.

The canvas therefore paints into a **1×1** region:

    cvsz  00000001 00000001

which is why the window is grey with no content: there is nothing to draw into.

### 2. The app then aborts on a stderr write

    PANIC std/src/io/stdio.rs: failed printing to stderr:
          FormatMessageW() returned error 120

A `println!`/`eprintln!` anywhere on that path panics because Win95's
`WriteConsoleW` is a stub and `FormatMessageW` cannot describe the failure.
`win9x_redirect_console_output` exists to prevent exactly this by pointing
`STD_OUTPUT_HANDLE`/`STD_ERROR_HANDLE` at a file — but it is called from
`WinMain` *before* `corro_main`, so the redirect is in place; something is
still reaching the real console handle, or the redirect's own failure path
prints.

The crash is what turns a cosmetic zero-size layout into a dead app: without
it the process would have kept running and a later layout pass with a real
size would likely have fixed the grid.

## Next step

The layout is the thing to fix, and the fix is already known from the ReactOS
side — a child must not be treated as hidden (or sized from a window that has
never been shown) just because its own style bit says so. The ReactOS fix
routes `set_visible` through an explicit registry; Win95 needs the same
discipline plus a guard so a box laid out at 0×0 does not pass that zero on to
its children as a final answer.

The stderr panic should be fixed first, though, because until the process
survives there is no evidence that the layout repair converges.

## Notes on the environment

* `/root/vm/win95-flat.qcow2` is **shared and was locked** by another
  operator's session during this work. It was not touched. The harness builds
  its own base image and stages everything there.
* ScanDisk ran on every boot until the images were split into a fsck'd `BASE`
  and a throwaway `SESSION` overlay; with that in place a run boots straight to
  the desktop.
* Another Win95 VM is often running on this host with a different VNC/QMP
  port. Only this harness's own pid is ever a kill target.
