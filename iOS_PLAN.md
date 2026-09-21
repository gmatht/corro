# iOS support for corro + rswidgets — findings and status

Status: **implemented** (arm64 + simulator verified by type-check on this Linux
host; the Xcode/ObjC half needs macOS to build — see §5). The findings below
were written before the code; §5 records what was actually built and how each
piece is verified.

Two claims in the obvious-but-wrong answer this file started from are corrected
here, because both are still the first thing people assume:

## 0. Corrections to the obvious-but-wrong answer

| Claim | Reality in this repo |
|---|---|
| "Enable `winit`'s iOS feature" | There is **no winit**. rswidgets has hand-written backends (`backends_*_adapter.rs`) over raw platform APIs; corro's TUI is `crossterm`/`ratatui`. iOS needs a **new backend**, like the Android one. |
| "Add `armv7-apple-ios` + `-mios-version-min=7.1.2` for iOS 7" | The **tier-3** name in this toolchain is **`armv7s-apple-ios`** (`arch: "arm"`, `llvm-target: armv7s-apple-ios`, `cpu: swift`, `+vfp4,+neon,no-divide`). Correct target for iPhone 5/5c; iPhone 4/4S additionally need plain `armv7` (not in this rustc's list) and are below the 7.1.2 minimum anyway. The 32-bit Apple *tier-2* targets (`armv7-apple-ios`, `i386-apple-ios`, `armv7s-apple-ios`) were dropped from tier-2 in Rust 1.82, but this nightly (**1.100.0-nightly, 2026-09-13**) still lists `armv7s-apple-ios` and `i386-apple-ios` as tier-3, and tier-3 targets are supported **only via `-Zbuild-std`** with a rust-src component — no `rustup target add` binary. Verified: `rustc --print target-list` here prints `armv7s-apple-ios`, `i386-apple-ios`, `aarch64-apple-ios`, `aarch64-apple-ios-sim`, `x86_64-apple-ios`. None are installed (`rustup target list --installed` has only the linux/windows/android ones), and the build host is linux — so no iOS build is possible here at all; it needs a macOS host with Xcode + the iOS SDK. |

Also: iOS 7.1.2 as a *deployment target* does not mean "it will look right on a
4" screen with no safe areas"; it means the armv7 slice and the pre-UIKit-8 SDK
surface. The SDK side is the real blocker (see §3).

## 1. What the Android port already gives us (the template)

Android is a *complete* worked example of exactly this port, and it is much
closer to iOS than GTK/NWG are:

* `rustxWidgets/rswidgets/src/backends_android_adapter.rs` — widget adapter:
  `#[repr(transparent)]` raw handles (`jobject`) kept alive by `GlobalRef`s,
  callbacks in id-keyed registries, text measured via `Paint.measureText` with a
  monospace fallback (`estimate_extents`).
* `rustxWidgets/rswidgets/src/backends/android.rs` — JNI env/activity/root
  layout statics, `KEEP_ALIVE`, app `ClassLoader` cache.
* `rustxWidgets/docs/ANDROID_GUIDELINES.md` — the contract: what the backend
  expects from the host (Java shims), threading (all on the UI thread),
  registry/lock rules, layout rules (no expand flags → `LinearLayout` weight),
  input paths (typing arrives via `TextWatcher`, not key events).
* `android/corro/{build.rs,build_apk.sh,app/...}` — host crate + Java shims
  (`MainActivity`, `SheetView`, `CorroTextWatcher`, `CorroEditorAction`,
  `CorroKeyListener`, `MenuStrip`, `RustCallback`).
* `src/gui/android_backend.rs` — corro side: `native_init` JNI export, leaks the
  `App` (platform drives the loop), `install_menu_strip` (phone can't show 6 text
  menus → inline quick actions + overflow popup).
* `examples/android_ui.rs` — build the *same* widget tree on a desktop to check
  layout without an APK round-trip.

**The equivalent iOS port is:** `backends_ios_adapter.rs` (UIKit over
Objective-C runtime / `#[no_mangle] extern "C"` callbacks) + `backends/ios.rs`
(statics, handle keep-alive) + `docs/IOS_GUIDELINES.md` + `ios/corro/`
(Swift/ObjC host, Xcode project) + `src/gui/ios_backend.rs` +
`examples/ios_ui.rs`.

The good news: **UIKit maps onto the existing trait set far better than
Android does.** `UIView`, `UILabel`, `UIButton`, `UITextField`, `UIScrollView`,
`UIStackView` are 1:1 with the `Widget` trait, and `UIFont`/boundingRect
replaces the `Paint.measureText` estimate with a real measurement.

## 2. Would the rswidgets **API** need adjusting?

Mostly no. The adapter *is* the API boundary; the Android port proves the
existing trait surface absorbs a touch-based native backend without breaking
callers:

* No `Touch` event variant was added for Android (there is no such variant in
  `core::Event` today — check before asserting otherwise). Android's taps are
  delivered as *canvas clicks* (`nativeOnTouch` → same path as a mouse click) and
  typing as text-changed. iOS does exactly the same: `touchesEnded` →
  canvas click; `UITextField` `editingChanged` → `nativeEntryChanged`
  (`CorroTextWatcher` equivalent); `textFieldShouldReturn` → `nativeEntryActivate`
  (`CorroEditorAction` equivalent). **No API change required.**

So the API changes are limited to:

1. **Cfg gates** — every `backends_*` module is `#[cfg(target_os = ...)]`; add
   `#[cfg(all(target_os = "ios", not(feature = "zork")))]` arms in
   `backends/mod.rs`, `lib.rs`, `Cargo.toml` target tables. Android's
   `cfg(target_os = "android")` arms are the exact pattern.
2. **Feature plumbing** — an `ios` feature in `rswidgets/Cargo.toml` and corro's
   `gui` already implies `gtk`/`headless`; corro's `gui` feature currently pulls
   `rswidgets/gtk`, which must not apply to iOS. Needs a target-specific feature
   set (Android already dodges this because `android/corro/Cargo.toml` depends on
   rswidgets with `default-features = false`).
3. **Genuinely new API (optional, additive)** — a `backends::ios` module with
   `init_with_root(view_controller)` mirroring `init_with_layout`, plus a
   way to register the canvas class analogue (`set_sheet_view_class` on Android
   → iOS: the `UIView` subclass name). Additive only.

The only *semantic* API addition worth considering is a touch **long-press /
scroll** gesture: Android's `SheetView` only forwards `ACTION_DOWN`, so there is
no two-finger scroll or pinch on Android either. If iOS is to be better, add a
`CanvasGesture` callback to the registry (additive), not a change to `Event`.

## 3. The actual hard parts (in order of severity)

1. **iOS 7.1.2 deployment target is nearly unbuildable in 2026.**
   * Apple's toolchain: Xcode 6+ is long gone. Current Xcode/iOS SDK refuses
     `-mios-version-min=7.1.2` (minimum is typically iOS 12 on Xcode 15/16).
     You would need an ancient Xcode (≤ 9 for iOS 7 headers, realistically
     Xcode 6 with iOS 8 SDK) on an old macOS — or the **iOS 7.1 SDK itself**,
     which is not redistributable.
   * Rust: 32-bit Apple targets were demoted out of tier-2 in 1.82. On this
     nightly they survive as **tier-3** (`armv7s-apple-ios`, `i386-apple-ios`)
     which means `-Zbuild-std=std,panic_abort` + `rust-src`, nightly-only — fine,
     since `rust-toolchain.toml` already pins nightly. `armv7-apple-ios` (plain
     armv7, iPhone 4/4S) is *not* in the list at all, so those devices are out
     unless a pre-1.82 toolchain is pinned. See §0.
   * **Recommended compromise:** ship **arm64, deployment target iOS 12/13+**
     (what current Xcode/Rust actually support), and if the iOS-7 constraint is
     a real requirement, budget for a pinned toolchain + old Xcode + armv7
     slice, plus `Info.plist` shims for pre-iOS-8 behaviours. Say which one is
     wanted *before* building anything; the code is the same, only the SDK
     plumbing differs.
2. **No windowing system.** iOS has no toplevel window: the app owns a
   `UIWindow` + root `UIViewController`. corro's `App::init`/`Backend::run`
   (blocking loop) must not be used; mirror `run_android_default` — leak the
   `App`, let UIKit drive callbacks (`UIApplicationMain` never returns).
   `set_default_size`, `set_title`, `hwnd` all become no-ops exactly as in the
   Android adapter.
3. **Drawing.** No Cairo/GTK. Options: `CGContext` (Core Graphics) directly via
   raw FFI — mirrors Android's `Canvas` path and keeps the `DrawContext` replay
   model; or Metal/CoreText for text. CG is the smaller change and matches the
   backend's primitive set (`drawText` baseline offset = Android's ascent fix).
4. **Objective-C glue without a Rust ObjC dependency.** The Android side uses
   the `jni` crate; the honest iOS equivalent is the `objc2` family
   (`objc2`, `objc2-ui-kit`, `objc2-foundation`, `objc2-core-graphics`) — much
   more ergonomic than hand-rolled `objc_msgSend`, and cfg-gated so hosts are
   unaffected. Hand-rolling is possible but is a lot of `sel!`-style unsafe.
5. **App lifecycle.** `applicationDidBecomeActive` / `…WillResignActive` /
   `…WillTerminate` must be forwarded (Android's backend ignores lifecycle;
   iOS will suspend you if you keep a timer running in background). Also the
   software keyboard (first-responder handling) for the formula entry.
6. **Files.** No `notify` on iOS (corro's `cfg(not(wasm32))` dep pulls it in) —
   file watching must be cfg-gated off, and open/save goes through
   `UIDocumentPicker`/`UIDocument` rather than paths. Same shape as the Android
   "files arrive later via content URIs" note.
7. **Packaging.** `.ipa` needs codesigning/provisioning; for testing, a
   simulator build needs none. `cargo-lipo`/`cargo-xcode` or a hand-written
   `xcodebuild` wrapper produces the static lib/framework the Xcode project
   links, mirroring how `build_apk.sh` feeds the `.so` into the APK.

## 4. Lambdatest — can we test there?

**Short answer: not the simulator build, and not for iOS 7.**

* `applive.lambdatest.com/app` runs **real devices** and supports uploading your
  own IPA (so a *signed* `arm64` build can be installed and tapped on). That is
  genuinely useful for the modern arm64 target, and it is the only part of your
  question that works as asked.
* It **cannot run a Xcode-simulator binary** (different arch/ABI, no simulator
  runtime on a physical device) and its device list is current-iOS only — no
  iOS 7.1.2 devices remain in any cloud farm. So it cannot validate the
  iOS-7.1.2 claim.
* Because you are logged in on Brave, the session works, but **automation is the
  catch**: the live app is interactive-only. For scripted runs you want
  Lambdatest **App Automation** (`desiredCapabilities` `platformName: iOS`,
  `app: <lt://…>`), which needs the LT API/access key, not the browser session.
  The IPA must be re-signed to an ad-hoc/dev provisioning profile for the farm
  (LT documents an "upload with re-signing" flow).
* Practical loop for this repo: `xcodebuild` a **simulator** build → run in
  `simctl` locally for the fast iteration (the analogue of the Android emulator
  + VNC setup in `android/corro/README.md`), and use Lambdatest App Live only
  for the real-device smoke test of a signed build.

## 5. What was actually built (status)

Everything below is in the tree and verified as described. **No macOS/Xcode is
needed to check the Rust half**; producing a runnable `.app` does need it.

### rswidgets (`rustxWidgets/rswidgets/`)

| File | What it is |
|---|---|
| `src/backends/ios.rs` (≈780 lines) | ObjC runtime wrappers (typed `objc_msgSend` per ABI shape, `objc_getClass`, `sel_registerName`, `objc_retain`, `NSString` marshal/unmarshal), handle `KEEP_ALIVE`, callback registry, `WidgetMeta` registry, `display_scale()`, `log_ios()`, `IosApp`/`init_backend`. |
| `src/backends_ios_adapter.rs` (≈2300 lines) | The widget surface: `Window`, `BoxWidget` (UIStackView), `Label`, `Button`, `Entry`, `Canvas` + CoreGraphics `CgDrawContext`, `Grid`, `DropDown`, `CheckButton` (UISwitch), `RadioButton`, `Dialog`, `TextView`, `ScrolledWindow`, `Overlay`, `Menu`, `MenuBar`, `SimpleAction`, every `create_*` factory, `display_density()`. Plus `dispatch_draw`/`dispatch_canvas_click`/`dispatch_canvas_key`/`dispatch_text_changed`/`dispatch_entry_activate`/`dispatch_focus`. |
| `src/lib.rs`, `src/backends/mod.rs` | `#[cfg(target_os = "ios")]` module + `init` re-export arms (priority: native backend wins over pancurses, like Android). |
| `src/common.rs`, `src/core.rs` | `platform_module!`/`common_types_mod!`/`App::new_*` arms for iOS, `IosOrientation`, and `App::quit()` documented as a deliberate no-op. |
| `Cargo.toml` | `ios` feature. **No new dependencies at all** — see the comment there for why (the armv7s/old-SDK path is exactly where a heavy ObjC binding stack breaks). |
| `docs/IOS_GUIDELINES.md` | The host contract, mirroring `ANDROID_GUIDELINES.md`. |

### corro (`src/`, `examples/`, `Cargo.toml`)

| File | What it is |
|---|---|
| `src/gui/ios_backend.rs` | `ios_main(root, vc)`, `run_ios_default()` (leaks the `App`, like Android), `menu_model()`, `run_menu_action_by_name`, `install_menu_model`, `log_ios`. |
| `src/gui/gui_backend.rs` | iOS arms for `metrics_scale()` (`UIScreen.scale`), the padlock hit target, `maintain_extent` viewport fill, and soft-keyboard text-changed adoption; the Android menu-state machinery generalised to `MOBILE_MENU_STATE`/`dispatch_mobile_menu_action` shared by both platforms. |
| `src/gui/mod.rs`, `src/lib.rs` | `gui::ios_backend` exposed under `gui-mobile`. |
| `src/io/mod.rs` | `notify` compiled out for iOS; `LogWatcher` falls back to the wasm size-poll (recorded in the iOS guidelines §6). |
| `Cargo.toml` | `gui-mobile` (shared GUI pipeline, no GTK) + `gui-mobile-host` (plus a desktop backend, for the examples). `ios-ui` example registered. |
| `examples/ios_ui.rs` | Builds the same tree on a desktop and prints the menu model. |

### ios/corro (the host)

| File | What it is |
|---|---|
| `src/lib.rs` | 9 `extern "C"` exports: root-ready, canvas draw/click/key, entry changed/activate/focus, generic callback, menu action + a three-call menu-model walk. |
| `app/CorroBridge.h` | The matching header, so a shim cannot drift from an export silently. |
| `app/CorroIosShims.m` | `SheetView` (drawRect/touchesBegan/pressesBegan with the modifier mask), `CorroIosTarget` (callback trampoline), `CorroIosText` (iOS 7+ / iOS 6 text API branching), `CorroIosAlert` + a `UIViewController` category (UIAlertController / UIAlertView). |
| `app/AppDelegate.{h,m}`, `app/SceneDelegate.{h,m}` | Scene (iOS 13+) and legacy (≤12) window paths, lifecycle forwarding. |
| `app/CorroViewController.{h,m}` | Bootstraps Rust, installs the text-field delegate, and builds the `UIMenu` from the model (with a pre-iOS-14 action-sheet fallback). |
| `app/Info.plist`, `app/LaunchScreen.storyboard` | Launch storyboard (prevents letterboxing), scene manifest, orientations, document sharing. |
| `build_ios.sh` | `cargo -Zbuild-std` staticlib → `xcodebuild`; `sim`/`device`, `--ipa`, `--run`. |
| `gen_xcodeproj.sh` | Generates `app/Corro.xcodeproj` (deterministic 24-hex ids) when missing; the result is checked in. |
| `README.md` | How to build, debug, and test on LambdaTest. |

### How each claim is verified

| Claim | Verified by | Result |
|---|---|---|
| The iOS backend compiles (arm64 simulator) | `cargo check --target aarch64-apple-ios-sim -Zbuild-std` | passes |
| It compiles for devices (arm64) | `cargo check --target aarch64-apple-ios …` | passes |
| It compiles for 32-bit iOS 7.1.2 devices | `cargo check --target armv7s-apple-ios …` | passes (this is also what exercises the `CGFloat = f32` code path) |
| corro's iOS paths compile | same three targets via a scratch crate with `features = ["gui-mobile"]` | passes |
| The host cdylib compiles | `cargo check` in `ios/corro` for `aarch64-apple-ios-sim` | passes |
| No new clippy findings | `cargo clippy` on the iOS targets vs. the desktop baseline | no iOS-specific findings |
| Desktop is untouched | `cargo check` (TUI), `--features gui`, `-p rswidgets`, `cargo test` | 663 + 29 tests pass |
| The widget tree and menu model are right | `cargo run --example ios-ui --features gui-mobile-host` | prints the tree + 6 menus / 64 items |
| The Xcode project is well-formed | the generator's output is ID-complete and all ids are 24 hex chars | checked |
| **The ObjC/Swift compiles and the app runs** | **needs macOS + Xcode** | **not checked here** |

### What remains (step #7, deferred)

1. Build the app on macOS: `./build_ios.sh sim --run`, then fix whatever the
   ObjC compiler finds (there is no way to type-check ObjC without an SDK).
2. Confirm the first frame renders (SheetView + `CorroIosText`) and that the
   formula entry round-trips: type → `corro_ios_entry_changed` → commit.
3. iOS 7.1.2 specifically: pin a toolchain with `rust-src`, obtain an Xcode
   that accepts `-mios-version-min=7.1.2` plus the archived iOS 7.1 SDK. Until
   then the 32-bit claim is "compiles", not "runs".
4. Optional: a signed arm64 IPA + LambdaTest App Live smoke test.


## 6. The order the work was done in

1. Decide **arm64 / iOS 12+** vs **armv7+arm64 / iOS 7.1.2** (this determines
   toolchain + Xcode + SDK). Answered by building *both* Rust paths and
   documenting the SDK blocker, so no decision blocks the code: the same
   source serves both and only the toolchain differs.
2. `rswidgets`: `cfg(target_os = "ios")` arms + `ios` feature; `backends_ios_adapter.rs`
   with just `Window`/`Box`/`Label`/`Button`/`Entry`/`Canvas` and CG draw;
   `docs/IOS_GUIDELINES.md` mirroring the Android one.
3. `examples/ios_ui.rs` + a simulator host app (Swift `AppDelegate` +
   `ViewController`, a `SheetView` UIView subclass) to build the tree and see it.
4. `corro`: `src/gui/ios_backend.rs` (leak App, `native_init` as
   `#[no_mangle] extern "C"`, menu strip analogue), cfg-gate `notify`.
5. `ios/corro/` host crate + Xcode project + `build_ios.sh`.
6. Lambdatest App Live (signed arm64 IPA) / App Automation for scripted smoke.
