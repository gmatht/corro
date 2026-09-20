# Android guidelines (rswidgets)

What the Android backend requires from a host app, what rswidgets
generates for you, and the rules the backend's own code follows. The code
is authoritative when this document drifts.

## 1. What rswidgets generates for you

`rswidgets::android_generator` writes the Android resources the backend
expects, **only when they are missing**:

| File (under `<project>/app/src/main/`) | Contents |
|---|---|
| `res/values/themes.xml` | `<package>.Theme`, parent `Theme.Material3.DayNight.NoActionBar` |
| `res/values/colors.xml` | brand seed / surface / icon-background colours |
| `res/values/strings.xml` | `app_name` label |
| `res/mipmap-anydpi-v26/ic_launcher.xml` | adaptive icon (background colour + foreground drawable) |
| `res/drawable/rswidgets_icon_foreground.xml` | neutral grid glyph; replace with your own artwork |
| `AndroidManifest.xml` | package, `minSdk 24` / `targetSdk 36`, theme, launcher activity |

Two ways to run it:

* **Host-app `build.rs`** (recommended): call
  `rswidgets::android_generator::run()` from the crate that owns the
  Android project. It reads `RSWIDGETS_ANDROID_PROJECT` (default
  `../android/corro` relative to that crate) and never fails the build.
  `android/corro/build.rs` is the worked example: because a build script
  cannot depend on the crate whose `src/` it lives in, it
  `#[path]`-includes `rustxWidgets/rswidgets/src/android_generator.rs`
  directly (the same source rswidgets' own `build.rs` includes), so it
  needs no `[build-dependencies]` entry at all. It sets
  `RSWIDGETS_ANDROID_PROJECT` to its own manifest directory (the root
  *containing* `app/`) unless the caller exported one, and is enabled by
  `--features generate-android-resources` (which `build_apk.sh` passes).

  Do **not** add rswidgets as a `[build-dependencies]` entry of a crate
  that also depends on it as a library: cargo then hands the library unit
  the build dependency's feature set (`default-features = false`, i.e. no
  `gtk`) while resolving its build-script output from the library edge,
  and the host crate fails on `crate::backends::init`. Reproduced with
  cargo 1.100.0-nightly and `cargo -vv`; including the source avoids it.
* **Manually / CI**: `AndroidProject::new(root).generate()` returns the
  paths it created, so a release script can report or verify them.

Generation is additive: an app that ships its own theme, manifest or icon
keeps them untouched. Delete a generated file to have it recreated.

### What rswidgets does *not* do

* It does not fetch the **Material Components AAR**. Add
  `implementation "com.google.android.material:material:<version>"` (or
  the equivalent AAR) to the app's build config — the generated theme
  parents against `Theme.Material3.*`, which lives there.
* It does not run Gradle, `aapt2`, `d8` or `apksigner`, and it does not
  sign or package anything. `android/corro/build_apk.sh` in this repo is
  an example of the raw-toolchain path; a Gradle/AGP project works the
  same way as long as the resources above exist.
* It does not generate the Java/Kotlin side. The app supplies
  `MainActivity` and the shims named in §3.

## 2. Building the tree without Android

The Android path builds the *same* widget tree as every other rswidgets
host — only the root differs (the Activity's content view instead of a
toplevel). To inspect that tree on a desktop, build it against the normal
backend and skip the event loop; `corro/examples/android_ui.rs` does this
for corro:

```text
cargo run --example android-ui --features gui
```

It creates the menu bar from the shared menu definition, the formula bar
(address label + `fx` label + expanding entry + status label), the sheet
canvas, the tab strip and the status line — in the order
`gui_backend::run_gui` appends them — then presents and returns. Nothing
Android-specific is compiled, so this runs anywhere the GUI feature builds
(use `xvfb-run` on a headless machine). Useful for checking that a
layout/order change is what you meant before rebuilding the APK, where the
edit-build-install cycle is far slower.

## 3. The Java/Kotlin side the backend calls into

| Class | Contract |
|---|---|
| `MainActivity` | `nativeInit(Activity, LinearLayout)`: calls `init_with_layout`, then runs the app. |
| Canvas view (default class `com.corro.SheetView`) | `(Context, long canvasId)` ctor; `onDraw(Canvas)` → `nativeOnDraw(id, canvas, w, h)`; `onTouchEvent` → `nativeOnTouch(id, x, y)`. Register the class with `set_sheet_view_class`. |
| Text watcher (`com.corro.CorroTextWatcher`) | `(long viewPtr)` ctor; `afterTextChanged` → `nativeEntryChanged(viewPtr)`. |
| Editor action (`com.corro.CorroEditorAction`) | `(long viewPtr)` ctor; `onEditorAction` → `nativeEntryActivate(viewPtr)` for IME Done/Enter. |
| Key listener (`com.corro.CorroKeyListener`) | `(long viewPtr)` ctor; `onKey` → `nativeEntryActivate(viewPtr)` for hardware/adb Enter. |

Missing shim classes degrade to no-ops (the widget tree still builds and
renders); the typing/commit path simply stays inactive.

Every `native*` declaration needs its own `#[no_mangle]` export in the
cdylib — including when two shims share one Rust dispatch. A missing export
only shows up at runtime as `UnsatisfiedLinkError`.

## 4. Threading

* Everything runs on the Android UI thread: `onDraw`, touch, text watchers,
  editor actions, and every Rust closure they call.
* `JNIEnv` is wrapped for `&self` access where the `DrawContext` trait
  demands it (`JniDrawContext.env` is a `RefCell<&mut JNIEnv>`). Always
  `try_borrow_mut` and degrade (skip the primitive, fall back to the
  estimate) rather than panicking.
* The `Send` impls on the callback registries are justified only by the
  registry mutex plus single-threaded dispatch. Never invoke them from
  another thread.

## 5. Handles and registries

* Widgets are `#[repr(transparent)]` over a raw `jobject` kept alive by a
  `GlobalRef` in `KEEP_ALIVE`. Never store a local ref; reconstruct
  `JObject::from_raw` only for the duration of one JNI call.
* Canvases are keyed by **canvas id** (`u64`); entries by view pointer;
  menu actions by name.
* Resolve the callback pointer under the registry lock, **drop the lock,
  then invoke**: callbacks re-enter (`queue_redraw` from a draw closure),
  and holding the lock across the call deadlocks.
* Resolve app (APK) classes through the cached app `ClassLoader`
  (`load_app_class`). `JNIEnv::find_class` on an attached thread uses the
  system loader and cannot see app classes.

## 6. Layout and text

* Android has no expand flags: `set_hexpand` / `set_vexpand` /
  `set_width_chars` are recorded and honoured by `BoxWidget::append` as
  `LinearLayout` weight and a minimum width. Canvases and expanding
  children grow; chrome wraps.
* An empty `EditText` measures zero: expanding children get a minimum width
  so the formula entry stays tappable.
* `text_extents_styled` measures with `Paint.measureText` (+
  `ascent`/`descent`) using the same size/weight the draw path uses, and
  falls back to the monospace estimate on any failure.
* `drawText` takes a baseline: offset by `ascent` to keep the shared
  top-left convention.

## 7. Input paths differ

* Soft-keyboard typing arrives only via the `TextWatcher`.
* IME Done/Enter arrives via `OnEditorActionListener`.
* Hardware and `adb shell input keyevent` arrive as raw key events —
  attach the key listener too, or Enter works only from the IME.
* Because soft typing has no key event, the shared edit state has to adopt
  it explicitly (android-gated in `on_formula_entry_changed`): entry text
  that differs from the committed cell value starts an edit; programmatic
  formula refreshes set entry == cell and are ignored.

## 8. Debugging

`logcat_rs` (tag `rswidgets`) and the host app's own logcat helper are
best-effort and silently dropped before init. Keep permanent call sites to
lifecycle events; per-frame or per-keystroke logging is diagnosis-only and
must be removed before committing.

Diagnose in order: listener attached in logcat → dispatch fired → pixels
(`adb exec-out screencap -p`). Dispatch missing with the listener present
means the event never reached the view (focus/delivery), not a Rust bug.
