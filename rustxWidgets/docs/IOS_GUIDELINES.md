# iOS guidelines (rswidgets)

What the iOS backend requires from a host app, what the host must generate or
ship, and the rules the backend's own code follows. The code is authoritative
when this document drifts; `src/backends/ios.rs` and
`src/backends_ios_adapter.rs` are the implementation, and
`../ANDROID_GUIDELINES.md` is the sibling document this one is modelled on.

## 0. Deployment targets, in one table

| Target | Rust target | Status |
|---|---|---|
| iPhone simulator, Apple silicon / Intel | `aarch64-apple-ios-sim`, `x86_64-apple-ios` | supported; what CI can build without a device |
| iPhone/iPad, 64-bit (iOS 12+) | `aarch64-apple-ios` | supported; the shipping configuration |
| iPhone 5/5c, 32-bit (iOS 7.1.2) | `armv7s-apple-ios` | **tier-3**: needs `-Zbuild-std` + a pinned toolchain + an old Xcode/SDK, see §9 |

`aarch64-apple-ios` and `aarch64-apple-ios-sim` are **tier-2**; a plain
`rustup target add` suffices. `armv7s-apple-ios` is **tier-3**: no prebuilt
`std`, so the build needs nightly plus
`-Zbuild-std=std,panic_abort` and a `rust-src` component.

## 1. What rswidgets does *not* do

Unlike Android (whose `android_generator` writes a theme, palette, icon and
manifest), iOS has **no resource generator**, and that is deliberate: an iOS
app is defined by an Xcode project and an `Info.plist`, both of which are
already reproducible build inputs and neither of which the Rust side can
sensibly synthesise. rswidgets generates nothing for iOS. The host supplies:

* the **Xcode project** (`ios/corro/app/`) — target, deployment version,
  bundle id, code signing;
* **`Info.plist`** — `UILaunchScreen`/launch images, supported orientations,
  document types if the app opens files;
* the **ObjC/Swift shims** named in §3;
* the **icons** (an asset catalog), which are a designer input, not a
  generated one.

What rswidgets *does* own is the whole Rust-side widget tree, the draw
replay, and the input dispatch on top of the shims.

## 2. Building the tree without iOS

The iOS path builds the *same* widget tree as every other rswidgets host —
only the root differs (the view controller's view instead of a toplevel). To
inspect that tree on a desktop, build it against the normal backend and skip
the event loop; `corro/examples/ios_ui.rs` does this:

```text
cargo run --example ios-ui --features gui-mobile-host
```

It creates the menu model, the formula bar (address label + `fx` label +
expanding entry + status label), the sheet canvas, the tab strip and the
status line — in the order `gui_backend::run_gui` appends them — then presents
and returns, printing the iOS menu model the host will render. Nothing
iOS-specific is compiled, so this runs anywhere the GUI feature builds (use
`xvfb-run` on a headless machine). It is the fast way to check a layout/order
change before the slow Xcode build.

## 3. The ObjC/Swift side the backend calls into

Everything below lives in the app target (see `ios/corro/app/`), because UIKit
must be able to call it *inside* the app bundle. Each name is fixed: the Rust
side resolves it with `objc_getClass` and sends the listed selector.

| Class | Contract |
|---|---|
| `CorroSceneDelegate` (or the app delegate) | On view load, calls the cdylib's `corro_ios_root_ready(root, viewController)`. That is the whole bootstrap: it ends in `rswidgets::backends::ios::init_with_root`. |
| `SheetView : UIView` | Property/selector `corroSetCanvasId:` (called by `create_canvas`); `drawRect:` → `corro_ios_canvas_draw(canvas_id, ctx, w, h)`; `touchesEnded:` → `corro_ios_canvas_click(canvas_id, x, y)`; optional `corroSetContentSize:` for the viewport size. Without it, canvases are plain `UIView`s (the tree still builds, nothing draws). |
| `CorroIosTarget` | Class factory `targetWithCallbackId:` returning a target that, on its `corroFired:` action, calls `corro_ios_callback(callback_id)`. This is the `RustCallback.java` equivalent. |
| `CorroIosText` | `measure:font:size:slant:weight:` → a `malloc`ed `CGRect*` (or null) with the text width/height; `drawText:ctx:font:x:y:size:r:g:b:a:slant:weight:` draws into the live `CGContextRef`. See §5 for why text is delegated. |
| `CorroIosAlert` | `corroNewAlert` returns a dialog object; `corroAddAction:` adds a button; `corroPresentDialog:` on the view controller presents it. The class decides whether to use `UIAlertController` (iOS 8+) or `UIAlertView` (iOS 7). |
| `CorroIosStackView` (older-target fallback) | Implements `addArrangedSubview:` and `corroSetSpacing:` for the pre-iOS-9 case (`UIStackView` is iOS 9+). Only needed if the app targets iOS 8 or below. |
| Text-field delegate | `editingChanged` → `corro_ios_entry_changed(viewPtr)`; `textFieldShouldReturn:` → `corro_ios_entry_activate(viewPtr)`; begin/end editing → `corro_ios_entry_focus(viewPtr, gained)`. These are the four entry signals. |
| Menu bar / overflow button | Reads the model from `corro_ios_menu_model` (see §4) and dispatches a chosen item through `corro_ios_menu_action(name)`. |

**A missing shim degrades, never crashes.** Every resolution above returns
null when the class is absent, and the callers log once and become no-ops:
the widget tree still builds, the sheet just does not draw, or buttons are
inert, or dialogs never present. That is how `examples/ios_ui.rs` runs on a
desktop at all.

## 4. The menu model (why the host builds the menu)

A phone cannot show corro's six text menus the way the desktop toolbar does.
As on Android, the adapter's `create_menubar` has **no view** behind it: it
carries the model (labels + `app.*` action names). corro flattens that model
with `gui::ios_backend::menu_model()` and the host renders it — a `UIMenu` on
a navigation-bar button (iOS 14+), a `UIAlertController` action sheet (8+), or
a `UIActionSheet` (7). Selecting an item calls back with the same `app.<name>`
string the desktop menu registers, which `dispatch_mobile_menu_action` routes
through the one dispatch table.

The menu *state* is published on the main thread when `run_gui` starts
(`MOBILE_MENU_STATE`, shared with Android), because a native menu can hold a
string but not a Rust closure.

## 5. Drawing and text

* The draw model is **identical to Android**: the backend replays the
  registered `DrawContext` closure through `dispatch_draw`, which Core
  Graphics implements one primitive at a time (`fill_rect` →
  `CGContextFillRect`, `clip` → `CGContextClipToRect`, and so on). The
  closure is keyed by **canvas id**, not by view pointer.
* `CGFloat` is **`f64` on arm64 and `f32` on armv7s**. The adapter's
  `objc_msgSend` wrappers select the right signature at compile time
  (`target_pointer_width`), and the `CGRect` passed to Core Graphics is
  field-laid-out for the target. This is the single most important ABI
  difference between the two builds; it is why there is no generic
  `msg_send` helper in the adapter.
* **Text is delegated to the host shim.** Measuring with
  `boundingRectWithSize:options:attributes:` (iOS 7+) is not available on
  iOS 6 and earlier, and the old `sizeWithFont:` is deprecated — so the
  version choice lives in the shim, not in Rust. Any failure (missing shim,
  unrenderable string) falls back to the monospace estimate
  (`EST_CHAR_W = 7.2`, `EST_LINE_H = 1.2`), so layout never divides by zero
  and a host without the shim still gets a correctly-sized tree.
* `drawText` takes a baseline on both platforms: the shim offsets by the
  font's ascent so the shared top-left convention holds.

## 6. Layout, metrics and files

* **No expand flags.** `set_hexpand`/`set_vexpand`/`set_width_chars` are
  recorded on the handle (`WidgetMeta`) and honoured by `BoxWidget::append`:
  an expanding or canvas child is flagged flexible (`corroSetFlex:`), and a
  `set_width_chars(n)` request becomes a minimum width (an empty
  `UITextField` measures zero, so without it the formula entry is
  untappable). This mirrors Android's `LinearLayout` weight.
* **Density.** `UIScreen.scale` (1.0 / 2.0 / 3.0) is exposed as
  `display_density()`; corro multiplies its pixel metrics by it
  (`gui_backend::metrics_scale`) so the grid is legible on a phone without a
  single desktop call site changing. Apple's floor for body text is 11pt and
  the touch target floor is 44pt — both are used as constants in the adapter
  (`MIN_TOUCH_PT`).
* **No file watching.** `notify` is cfg'd off for iOS (`src/io/mod.rs`): a
  sandboxed app sees only its own container, and documents arrive through the
  picker, not a watched path. `LogWatcher` falls back to the size poll the
  wasm build uses. `open_file`/`save_file` return `Ok(None)` — a document
  picker is the host's business.

## 7. Threading and lifecycle

* **Everything runs on the main thread**: `drawRect:`, touches, the text-field
  delegate, and any Rust closure they call. The callback registries are
  `Mutex`-guarded and their `Send` impls are justified only by that plus
  single-threaded dispatch — never invoke them from another thread.
* Lock discipline is the Android rule: resolve the callback pointer under the
  lock, **drop the lock, then invoke**, because callbacks re-enter
  (`queue_redraw` from a draw closure) and holding the lock deadlocks.
* `IosApp::run()` returns immediately. `UIApplicationMain` never returns, so
  the app *is* the event loop; corro leaks its `App` for the same reason
  Android does (closures hold raw pointers that outlive the call).
* `App::quit()` does **nothing** on iOS, by design: Apple rejects apps that
  terminate themselves. "Quit" is the host's decision (save, then leave the
  scene). It logs and returns.
* `applicationDidBecomeActive` / `…WillResignActive` / `…WillTerminate` are
  the host's to forward. The backend does not poll in the background — a
  suspended app must not be running a timer.

## 8. Handles and registries

* Widgets are `#[repr(transparent)]` over a raw Objective-C object pointer,
  retained by us (`objc_retain` on every handle we create) and listed in
  `KEEP_ALIVE`. Never release a handle: widgets are shared across closures,
  and a view removed from the hierarchy must not deallocate while a callback
  still points at it. (UIKit owns the hierarchy itself; our retain is for the
  Rust-side lifetime.)
* Metadata (`WidgetMeta`: kind, text, expand flags, min chars, canvas id,
  requested and laid-out size) is kept **beside** the objects, keyed by handle
  pointer, rather than in associated objects — ObjC associated storage is one
  more runtime API to get wrong on 32-bit.
* Canvases are keyed by **canvas id** (`u64`, from 1); entries by view
  pointer; button callbacks by `u64` id; menu actions by name. Two size maps
  exist on purpose, exactly as on Android: `size_request` (what Rust asked
  for, often a 1x1 placeholder) and `laid_out` (what the host reported from
  `drawRect:`). The placeholder must never shrink a real laid-out size, or
  the sheet viewport collapses to one row.

## 8b. Can Zig replace Xcode? (no — but it is useful for one thing)

Short answer: **no**. Zig 0.16 can compile C and `.m` for `-target aarch64-ios`,
but it ships no Apple SDK, and the app needs four things only an SDK provides.
Measured on this machine, in the order they bite:

| Step | Result with Zig alone |
|---|---|
| `zig cc -target aarch64-ios -c foo.c` (no frameworks) | **works** — produces a Mach-O arm64 object |
| `#include <UIKit/UIKit.h>` | **fails** — `'UIKit/UIKit.h' file not found`; zig ships no headers |
| `.m` with `-fobjc-arc`, `#import <Foundation/Foundation.h>` | **fails** — no Foundation headers |
| Linking the Rust `corro_ios` staticlib (`-lobjc -framework CoreGraphics -liconv`) | **fails** — `unable to find dynamic system library 'objc'`; zig bundles **only** `libSystem.tbd` |
| A launchable `.app` (bundle layout, `Info.plist` embedding, codesign) | not something zig does at all |

Two details worth knowing, because they are not obvious:

* **Zig's bundled `libSystem.tbd` has no iOS slice.** Its `targets:` line is
  `[ x86_64-macos, x86_64-maccatalyst, arm64e-macos, arm64e-maccatalyst ]`.
  So even `-lSystem` for an iOS link is resolving against macOS stubs.
* **`-mios-version-min` cannot be honoured through Zig.** Zig's darwin libc is
  its own `libSystem.tbd`, so the deployment target is whatever that stub says,
  not what the flag asks for. That is exactly the constraint this project cares
  about most (§0), which makes Zig a poor fit for the iOS 7.1.2 path in
  particular.

If you *did* supply an SDK (`SDKROOT=/path/to/iPhoneOS.sdk`), `zig cc` becomes a
usable clang driver for it — but at that point the SDK is doing the work Xcode
was doing, and you have not removed the Apple dependency, only Xcode's UI.

**What Zig genuinely does buy, and what this repo now does:** building the Rust
side without a linker at all. `cargo rustc --target aarch64-apple-ios --lib
--crate-type rlib` needs no SDK and no `cc`, because an rlib is an archive of
objects rather than a linked image. That is the closest thing to a
"cross-compile without Xcode" story here, and it is what
`scripts/check_*_ios.sh` already do in `--emit=metadata` form. Anything past
that — a `.a`, a `.dylib`, an app — needs a linker and therefore an SDK.

## 8c. Where the simulator run actually stands

`ios/corro/scripts/verify_all.sh` passes, and the app **builds, installs,
launches and runs its full Rust startup** on a hosted macOS runner. The console
capture shows the whole sequence, which is the fastest way to see where a
future break is:

```
[rswidgets] ios_main: init_with_root
[rswidgets] ios backend initialised
[rswidgets] ios_main: building the App
[rswidgets] ios_main: load_initial
[rswidgets] ios_main: entering run_gui
[rswidgets] run_gui: crash handlers installed
[rswidgets] run_gui: App::init
[rswidgets] run_gui: App::init ok
[rswidgets] run_gui: new_window
[rswidgets] run_gui: new_window ok
[rswidgets] run_gui: new_box
[rswidgets] run_gui: new_box ok
[rswidgets] corro: menu model ready (6 menus, 64 items)
[rswidgets] ios_main: run_gui returned
PHASE: before_set_draw_callback
DRAW_CALLBACK called: w=1 h=1
PHASE: after_set_draw_callback
PHASE: about_to_present
PHASE: after_present
```

**Two known gaps**, both visible in that output:

1. **The canvas is 1x1.** `DRAW_CALLBACK called: w=1 h=1` is the placeholder
   from `canvas.set_size_request(1, 1)`, so the backend never learns the real
   laid-out size. On Android that arrives from `View.onDraw`; on iOS the host
   `SheetView` must report it (`corro_ios_canvas_size` does not exist yet), so
   the sheet would render one row stretched over the screen — the same symptom
   Android had before `CANVAS_SIZE` was wired up.
2. **The process does not stay up.** It runs the whole startup and then exits
   within ~15 s. `run_gui` returns normally (`IosApp::run` is documented as
   returning immediately because UIKit owns the loop), no Rust code calls
   `process::exit` on that path (both calls are inside `save_before_quit`), and
   no crash report is produced for the app — so this is not a panic. The next
   step is to find what terminates it; a `UIApplication` delegate logging in
   `applicationWillTerminate:` plus `NSLog` from `main` after
   `UIApplicationMain` returns would settle whether UIKit is being torn down or
   the process is being reaped.

Neither gap is a mystery in kind: both are "the host has not supplied something
the backend needs", which is the same category as the five missing selectors
that crashed `create_box` (see §3).

## 9. Debugging, and the iOS 7.1.2 path

Debug in order, mirroring the Android section: shim present → dispatch fired →
pixels.

* **System log.** `log_apple`/`log_ios` writes to the process's **stderr** (tagged `[rswidgets]`), which the simulator console already captures:
  `xcrun simctl spawn booted log stream --predicate 'eventMessage CONTAINS
  "rswidgets"'` (or Console.app). Keep permanent call sites to lifecycle
  events; per-frame logging is diagnosis-only and must be removed before
  committing.
* **Screenshots.** `xcrun simctl io booted screenshot shot.png`; taps via
  `xcrun simctl io booted tap X Y` (Xcode 15+) or `idb`.
* **Crash reports.** `objc_msgSend` with a wrong signature is the classic iOS
  crash and it does *not* produce a nice message — it lands in the wrong
  register. If a build crashes immediately on a call that works elsewhere,
  suspect the signature (especially `CGFloat`) before the logic.
* **iOS 7.1.2 / armv7s.** The Rust side of this path is real and checked
  (the adapter compiles for `armv7s-apple-ios`), but building a runnable
  binary needs: a pinned pre-1.82-ish nightly or a nightly with `rust-src`,
  `-Zbuild-std=std,panic_abort`, an Xcode old enough to accept
  `-mios-version-min=7.1.2` (Xcode 6/7 era), and the iOS 7.1 SDK for the
  deployment-target minimum. Apple no longer ships or licenses that SDK, so
  this is a "if you have the hardware and the archived toolchain" path — the
  recommended shipping configuration is arm64 / iOS 12+. Pre-iOS-8 idioms the
  code already accounts for: `UIAlertView` instead of `UIAlertController`
  (§3), the `CorroIosStackView` fallback for pre-iOS-9 `UIStackView` (§3),
  and the `sizeWithFont:` text path (§5).
* **Testing elsewhere: the four kinds of "online simulator".** They are not
  interchangeable, and only one of them solves "I have no Mac":

  | Kind | Examples | Takes | Can it build? | Useful here? |
  |---|---|---|---|---|
  | Real-device farm | LambdaTest App Live, BrowserStack App Live, AWS Device Farm | a **signed** `.ipa` | no | yes, as a *smoke test* after a build exists; no iOS 7 devices remain, and live sessions are interactive (scripted runs need App Automation + credentials) |
  | Cloud app streaming | Appetize.io | an **unsigned simulator `.app` zip** | no | yes in principle — the only category needing no Apple developer account — but it still needs a Mac to produce the binary |
  | Cloud macOS CI | GitHub Actions `macos-14`, Codemagic, Bitrise | the repo | **yes** (Xcode + `xcodebuild` + `simctl`) | **this is the answer**: it compiles the ObjC, boots the simulator and screenshots it |
  | Virtualised iOS on ARM | Corellium | either | n/a | rarely the right tool (cost, availability, legal posture) |

  The gate for all of them is the same: *something* must produce a build first.
  A farm or a streaming simulator cannot; only a Mac (or a rented one) can.

* **`.github/workflows/ios.yml` is written for exactly that** — but note the
  caveat: **it has never been executed.** It is a first draft, reviewed on a
  host with no Xcode, no simulator and no GitHub access. It builds the app on
  `macos-14`, boots a simulator, launches corro, screenshots the first frame,
  asserts the process is still alive (catching an `objc_msgSend` signature
  crash, which is the one iOS-specific failure mode that gives no backtrace),
  and uploads the log lines under the `rswidgets`/`corro` tags. Simulator
  builds need no signing identity, so the workflow needs no secrets. Its
  companion `rust-ios-check` job runs the Linux cfg checks first, so a failure
  points at the port rather than the Xcode plumbing.

  What *is* verified about it: the YAML parses and every `run` step passes
  `bash -n` (`ios/corro/scripts/verify_all.sh` checks both). What is not: that
  the steps actually succeed on a runner. Expect to iterate on the toolchain
  steps at first run, since `ios/corro` is a standalone workspace and the
  cache/component setup is the likeliest place to need adjusting.

  It cannot cover iOS 7.1.2: no hosted runner carries the archived SDK (§0).
