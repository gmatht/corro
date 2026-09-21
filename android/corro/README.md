corro on Android — live view
============================
A VNC server serves display :99, where a live view of the Android emulator
shows corro running.

Connect:   <host-ip>:5999   (no password: -SecurityTypes None)
  host ip: 172.25.38.120  (container eth0; from the docker host, forward/map port 5999)
  display: :99  (1280x800)
  VNC listen: 0.0.0.0:5999

App:       com.corro/.MainActivity   (APK: /tmp/corro_apk_build/signed.apk)
Emulator:  AVD corro_avd (Android 13, x86_64, port 5554)
Logs:      adb logcat -s corro rswidgets
Screenshot: adb exec-out screencap -p > shot.png
Taps:      adb shell input tap X Y
Scroll:    adb shell input swipe 540 1800 540 500 500   (slow drag = more rows)

Mirroring the emulator
----------------------
    ./mirror.sh          # then connect a VNC viewer to <host>:5999

`mirror.sh` launches the app, then loops: `adb exec-out screencap` -> a PNG
-> an `feh --reload` window on `:99`, which the VNC server mirrors.

Do NOT use scrcpy here. The emulator's MediaCodec encoder dies after a while
("Encoding error: IllegalStateException"); the scrcpy client process stays
alive but stops receiving frames, so the VNC viewer sits on a stale frame
forever and the app *looks* frozen when only the mirror is. `mirror.sh`
avoids the device encoder entirely, so the view stays live.

If the window looks frozen, the check that distinguishes the two cases is:

    # device side: does a tap change anything on the device?
    adb exec-out screencap -p > a.png; adb shell input tap 540 1200; sleep 3
    adb exec-out screencap -p > b.png; cmp -s a.png b.png || echo "device is live"

    # mirror side: does the VNC-visible window change too?
    tail -3 /tmp/live_loop.log      # frame counter must keep advancing

Building
--------
    ./build_apk.sh [x86_64|arm64-v8a] [install]

The script builds the `corro_android` cdylib for the NDK target with
`--features generate-android-resources`, then packages the Java sources,
`app/src/main/res` and the `.so` with the raw toolchain (javac/d8/aapt2).

This crate is NOT in the corro workspace: it targets the Android NDK with
its own feature set, so build it from here (or via the script), not with a
workspace-wide `cargo build`. It only compiles for an Android target — the
JNI exports and `corro::gui::android_backend` are android-gated.

Touch interaction
-----------------
Android has no native scrolling for the sheet: the grid is drawn by Rust
into a plain `Canvas`, and the adapter's `ScrolledWindow` is an inert
`FrameLayout`. `SheetView.onTouchEvent` therefore classifies the gesture
itself and hands whole-cell counts to Rust:

* **tap** (`ACTION_UP` without exceeding touch slop) — moves the cursor to
  the tapped cell;
* **drag** — scrolls the viewport.

The drag path accumulates the pixel delta from where the gesture *started*
(not from where touch slop was exceeded — discarding the pre-slop travel
loses the first row of every gesture), converts it to rows/columns using
`nativeCellSize()`, and carries the sub-cell remainder into the next event
so a slow drag still scrolls smoothly. `nativeCellSize()` exists so the
conversion uses the very metrics Rust rendered with; a constant hardcoded
in Java would drift with density or font changes.

`scroll_viewport_by_cells` in `gui_backend.rs` moves the *cursor*, because
this GUI has no independent scroll offset: `displayed_rows`/`displayed_cols`
derive the viewport from `app.core.cursor` (`prev_start` is always 0), so
the cursor is the viewport origin. That keeps one code path for "viewport
moved" and keeps the selected cell visible for free.

Layout notes
------------
* The **menu strip is pinned** to the top: it is inserted at index 0 of the
  root `LinearLayout` with `WRAP_CONTENT` height and zero weight
  (`pin_child_at_top`), so the weighted sheet expands beneath it instead of
  squeezing it. Grid scrolling moves only grid content, so the menu stays
  visible for the whole session.
* Row height is `ROW_H_BASE × metrics_scale()`, where `metrics_scale` is the
  display density (2.625 on a 420dpi device). Grid geometry (`row_h`,
  `char_w`) and the glyphs drawn through `Paint.setTextSize` are therefore in
  the same unit and agree — text fills its row instead of clipping.
* `Canvas::replay_size` deliberately ignores `set_size_request`'s
  placeholder. corro asks for 1x1 (Android measures children itself), and
  replaying the draw closure at 1x1 makes it cache a **one-row** viewport,
  which the next real `onDraw` then renders as a single enormous row filling
  the canvas — a grid that looks empty. Replaying at a plausible default
  instead keeps the pre-layout draw harmless.
* On mobile the body is grown to fill the viewport (`maintain_extent`), plus
  one extra column for each side's margin border. Without that the body keeps
  its minimal size while the viewport still spends its first columns on the
  left margin, so an empty sheet renders as a small white patch in a field of
  margin grey. The result is a white body framed by a one-cell margin border
  on all four sides, which is what an empty sheet should look like.

Android resources
-----------------
`build.rs` runs `rswidgets::android_generator` (the source is
`#[path]`-included from the rswidgets checkout, so there is no
`[build-dependencies]` entry) and writes the Material 3 theme, palette,
strings, adaptive launcher icon and manifest into this directory —
the root that contains `app/`. Generation is additive: every file this
project already ships is left alone, so it only fills gaps and never
fights hand edits. Delete a resource to have it recreated.

Off by default; enable it with the feature, or point the generator at
another project entirely:

    cargo ndk -t x86_64 build --features generate-android-resources
    RSWIDGETS_ANDROID_PROJECT=/path/to/other/project cargo ndk -t x86_64 build

See `rustxWidgets/docs/ANDROID_GUIDELINES.md` for what rswidgets generates
vs. what the app must supply (Material AAR, the Java shims, packaging).

Previewing the UI without an emulator
-------------------------------------
The widget tree this crate roots in the Activity is the same one the
desktop GUI builds. To look at it on a host (much faster than an APK
round-trip):

    cargo run --example android-ui --features gui      # from the repo root
