# 0.7.0 – Windows 95 support (WIP); new release artifacts

## Demo mode on the GUI backends
- **`--movie` now works on the GUI backends**, not just the ratatui terminal.
  The workbook parsing and op application are shared (`src/gui/movie.rs`), so a
  movie reads identically whichever UI replays it; only the frame-painting half
  is backend-specific (`src/gui/gui_movie.rs` for the widget backends,
  `run_pancurses_movie` for pancurses).
- `scripts/gui_movie.py` records a movie from the running window under a
  throwaway X server and encodes it with ffmpeg, so the demo video is
  reproducible from a checkout; a recording is checked in at
  `dist/corro-gui-movie.mp4`. `scripts/demo_movie.py` builds the shipped demo
  from title cards plus those replays.
- Movie pacing flags (`--movie-typing-cps`, `--movie-confirm-ms`,
  `--movie-menu-hold-ms`) now parse on the CLI for every backend rather than
  only being honored by the ratatui path.
- **A movie replay no longer writes to the log it is reading.** Every GUI
  commit path appends to `app.core.path`, and the replay left the movie's own
  file bound, so each recording appended its steps to the fixture — the demo
  workbooks grew on every run. The replay now detaches the file (as the TUI
  replayer already did) while keeping it as the display source.
- **Concurrent editing demos.** `scripts/two_window_movie.py` records two
  front-ends editing one file (`--pair gui-gui` or `gui-tui`), showing each
  window's commit appear in the other. Both windows are scripted via
  `CORRO_EDIT_SCRIPT` and apply their edits through the ordinary commit path
  (`ui_core::edit_script_from_env`; the GUI arms it from its timer, the TUI from
  its loop), because driving them with synthetic keystrokes needs pixel
  calibration that silently writes to the wrong cell when it drifts.
- `scripts/demo_movie.py` builds the shipped demo video from title cards plus a
  replay of each featured workbook, recorded from the running window. The
  `main.corro` card notes that its `#NAME`/`#PARSE`/`#CIRC` cells are deliberate
  broken-formula fixtures rather than rendering faults. The
  typing speed is a flag (`--cps`), and the demo uses a deliberately readable
  6.5 characters/second (a quarter of the old 26) so the typed values can
  actually be followed in the recording.
- **Movie mode is now the normal UI.** `--gui --movie` opens the same window,
  builds the same widget tree and runs the same draw callbacks as an
  interactive session; a periodic timer applies one movie step per tick instead
  of waiting for a keystroke (`gui_backend::arm_movie_driver`). The window
  closes itself when the script ends, like the TUI replayer quits. The parallel
  raster renderer is gone — it had drifted from the window four times (blank
  frame, overlapping glyphs, missing margin shading, missing text) because it
  was a second implementation of rendering. `scripts/gui_movie.py` now records
  the live window under Xvfb instead of producing frames from that renderer.
- The GUI movie painter no longer draws its own copy of the sheet. It now
  calls the same body renderer the live canvas uses
  (`gui_backend::render_grid_body`), so margin shading, cell colours, the
  cursor ring, selection and gutter padlocks cannot drift from the window —
  the movie had lost the grey margin bands precisely because it was a second
  implementation. The viewport is likewise built by the shared controller.
- Fixed two frame-rendering defects in the GUI movie painter: the sheet was
  sized with a *column count* where the viewport expects a *character width*,
  which trimmed it down to the margin columns and left most of the frame blank
  (the "huge empty spaces" the window itself had before it sized its viewport
  from the live canvas); and glyphs were drawn at the 7.2px cell-layout pitch,
  which made 2x text overlap into an unreadable smear. Columns are now
  stretched to span the frame and text uses its own advance.

## Changes
- **Saving an unsaved document** no longer drops the auto‑added TOTALs. The untitled log (which Save renames verbatim onto the destination) was written as a bare header, so the in‑memory margin seeds — which have no committed edit op — vanished on reopen. It is now a complete serialization (all regions, seeds and LINKs included).
- **Windows release exes** build with the GUI feature so Explorer double‑clicks open the NWG window. The exe stays console‑subsystem (so cmd/PowerShell wait for the TUI until \`Ctrl+Q\`); a throwaway Explorer console is detected (only our process attached) and hidden/freed at startup. Previously every modern Windows exe was console‑subsystem with no GUI to fall back to, or (briefly) linked GUI‑subsystem, which made the shell return to its prompt immediately while the TUI kept drawing into the same console.
- **Windows now redirects stderr** to the debug log (like the existing Unix `dup2`), so debug/instrumentation traces can never overwrite the TUI; `%LOCALAPPDATA%\corro\debug.log` is used when no `CORRO_DEBUG_LOG`/`XDG_STATE_HOME`/`HOME` is set.
- **pancurses+ratatui+gui** now combine on Linux and Windows (triple combo builds and runs every UI). `rswidgets` resolves its default backend native‑first as documented (`pancurses` stays reachable via `backends::pancurses::init()`); the NWG adapter stack is available alongside pancurses on Windows, and corro's native GUI code names the platform adapter + common wrappers explicitly instead of the flipping root prelude. Previously the combo failed to compile (17 errors Linux, 19 Windows) and Linux `--gui` died with “loader not initialized”.
- **Android: the UI can be built and inspected without an emulator.** `cargo run --example android-ui --features gui` builds the exact widget tree the Android path roots in the Activity (menu bar from the shared menu definition, formula bar, sheet canvas, tab strip, status line) and exits before the event loop, so a layout change can be checked in seconds instead of an APK round‑trip. `src/gui/android_backend.rs` and the `gui` feature's rustdoc now describe the Android host explicitly.
- **Android resources are generated by the build.** `android/corro/build.rs` runs `rswidgets::android_generator` (the source is `#[path]`-included from the rswidgets checkout, so no `[build-dependencies]` entry — adding one breaks cargo feature unification and the host crate fails on `crate::backends::init`) and writes the Material 3 theme, palette, strings, adaptive launcher icon and manifest into the project root that contains `app/`. Additive, so hand‑edited resources are never overwritten; `build_apk.sh` passes `--features generate-android-resources`, and `RSWIDGETS_ANDROID_PROJECT` still overrides the root.

### Release artifacts for this version
- **Linux**: `gcorro.gz` (GTK GUI, runtime‑dlopen) and `corro-no-gui.gz` (terminal UI), both linked for glibc 2.17, alongside `corro.gz`.

#### WIP / not shipped
- **Windows 95 support**. The pancurses TUI backend renders and takes input on Win95 (PDCurses win32 console flavour, ANSI APIs; wait‑based input poll; key‑code mapping for non‑WIDE PDCurses), and the native NWG GUI builds for Win95 (`/SUBSYSTEM:WINDOWS`, `WinMain` entry). This is **NOT** part of the 0.7.0 release: it cannot be built on CI (rust9x toolchain + VC6 libs are local‑only) and is not gated by the test suite, so no `corro-no-gui-win95.exe.zip` / `gcorro-win95.exe.zip` is published for 0.7.0. Treat the `i586‑rust9x` targets as experimental.
- **Busybox‑style argv[0] dispatch** (`gcorro*` prefers the GUI, `pcorro*` the pancurses UI; explicit flags always win) works on the platforms built above, but the Win95 exes that motivated it are WIP.
- **Win95 runtime note** (for when it does ship): rust9x exes need `unicows.dll` next to the exe (or in `C:\WINDOWS\SYSTEM`); it is not bundled (see `README95`).

---

*Generated from the project’s changelog.*
