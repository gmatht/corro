corro on iOS
============

The iPhone/iPad host for corro: an Xcode app (ObjC shims + a checked-in
generated project) that links the `corro_ios` Rust static library built from
this directory.

    ios/corro/
    ├── Cargo.toml           corro_ios cdylib/staticlib (standalone workspace)
    ├── src/lib.rs           the extern "C" entry points the shims call
    ├── build_ios.sh         Rust staticlib + xcodebuild (+ .ipa / simulator run)
    ├── gen_xcodeproj.sh     generates app/Corro.xcodeproj when missing
    └── app/
        ├── AppDelegate.{h,m}       lifecycle, legacy (pre-scene) window
        ├── SceneDelegate.{h,m}     iOS 13+ window + navigation controller
        ├── CorroViewController.{h,m}  bootstraps Rust, owns the menu bar
        ├── CorroIosShims.m         SheetView, text, alert, callback target
        ├── CorroBridge.h           the C surface (matches src/lib.rs)
        ├── Info.plist, LaunchScreen.storyboard
        └── Corro.xcodeproj/        generated, checked in

Verifying without a Mac
----------------------

Every check that does not need Xcode is scripted, so the Rust half of this
port can be verified on Linux (or in CI):

    scripts/verify_all.sh          # all of them: iOS cfg checks + desktop regressions

The individual pieces:

    scripts/check_rswidgets_ios.sh   # rswidgets iOS cfg, sim + arm64 + armv7s
    scripts/check_corro_ios.sh       # corro iOS cfg, same three targets
    scripts/check_host_ios.sh        # this crate's cdylib
    scripts/probe_zig_ios.sh         # why Zig cannot replace Xcode (and what it can do)

They use `-Zbuild-std` (metadata only — no link, so no iOS SDK) and therefore
prove the *Rust* compiles for all three targets, including the 32-bit armv7s
path. What they cannot cover is compiling the ObjC and running the app: that
is `build_ios.sh sim --run` on macOS, and it is the one thing still untested
here (`rustxWidgets/docs/IOS_GUIDELINES.md` §9).

Building
--------

    ./build_ios.sh sim --run      # simulator, no signing, install + launch
    ./build_ios.sh device --ipa   # arm64 device build + signed .ipa

`sim` needs only Xcode. `device` additionally needs a signing identity and
provisioning profile (`DEVELOPMENT_TEAM=<teamid>`), because iOS will not run an
unsigned binary on real hardware.

The Rust half is built with `-Zbuild-std=std,panic_abort`: the iOS targets have
no prebuilt `std` in the pinned toolchain, so nightly plus a `rust-src`
component is required. `IOS_DEPLOYMENT_TARGET` (default `12.0`) sets both the
Rust `-mios-version-min` link argument and the Xcode build setting.

Deployment target: 12.0 by default, and that is the honest floor for a
*shippable* build (current Xcode refuses much older). The code is written so
the same source can go lower — `UIAlertView` instead of `UIAlertController`,
a `CorroIosStackView` fallback for pre-iOS-9 `UIStackView`, `sizeWithFont:`
for pre-iOS-7 text — but actually building for iOS 7.1.2 needs
`armv7s-apple-ios` (tier-3), an Xcode old enough to accept
`-mios-version-min=7.1.2`, and the archived iOS 7.1 SDK. See
`rustxWidgets/docs/IOS_GUIDELINES.md` §9 for the full picture and the table of
what is and is not checked in this repo.

Previewing the UI without Xcode
-------------------------------

The widget tree this app roots in the view controller is the same one every
other backend builds. To look at it on a workstation (much faster than an
Xcode round trip):

    cargo run --example ios-ui --features gui-mobile-host   # from the repo root

It prints the tree structure and the menu model the app renders, without
compiling a single line of iOS-specific code.

Logs
----

Rust-side messages go to `NSLog`, i.e. the device/simulator console:

    xcrun simctl spawn booted log stream --predicate 'eventMessage CONTAINS "rswidgets"'

Screenshots and taps:

    xcrun simctl io booted screenshot shot.png
    xcrun simctl io booted tap X Y          # Xcode 15+

Debug order (same as Android's): is the shim present → did dispatch fire →
are there pixels. `IOS_GUIDELINES.md` §9 has the details, including the
`objc_msgSend`-signature crash that is the one iOS-specific failure mode worth
knowing by heart.

Testing in the cloud (no Mac of your own)
----------------------------------------

Four different things get called "an online simulator"; only one of them
removes the macOS requirement, and it is not the device farm:

| Kind | Takes | Can build? |
|---|---|---|
| Real-device farm (LambdaTest App Live, BrowserStack, AWS Device Farm) | a **signed** `.ipa` | no |
| Cloud app streaming (Appetize.io) | an **unsigned simulator `.app`** | no |
| **Cloud macOS CI (GitHub Actions `macos-14`, Codemagic, Bitrise)** | the repo | **yes** |
| Virtualised iOS (Corellium) | either | n/a |

`.github/workflows/ios.yml` is the third row, checked in: it builds the app on
a hosted Mac, boots a simulator, launches corro, screenshots the first frame,
verifies the process survived launch (an `objc_msgSend` signature error crashes
without a usable backtrace, so "is it still alive" is the assertion that
matters), and uploads the `rswidgets`/`corro` log lines. Simulator builds need
no signing, so no secrets are required — just run the workflow.

That is also the fastest way to get the screenshots `docs/ios/README.md`
describes as missing: the workflow uploads them as a build artifact.

It still cannot cover iOS 7.1.2 — no hosted runner has the archived SDK.

Testing on a device farm
------------------------

LambdaTest App Live (`applive.lambdatest.com/app`) can install a **signed
arm64 IPA** and is the right tool for a real-device smoke test:

1. `./build_ios.sh device --ipa` (or archive/export from Xcode with an
   ad-hoc/development profile).
2. Upload the `.ipa`; LambdaTest re-signs it for the farm's devices.
3. For scripted runs use App Automation (`platformName: iOS`, `app: lt://…`),
   which needs the LT access key — the live session in a browser is
   interactive only.

What it cannot do: run a simulator build, or test iOS 7 (no such devices remain
in any farm). Fast iteration belongs in `simctl` locally.

What is not here
----------------

* **No resource generator.** Unlike Android (`rswidgets::android_generator`
  writing a theme, palette, icon and manifest), an iOS app is defined by its
  Xcode project, `Info.plist` and asset catalog — all reproducible build
  inputs the Rust side cannot sensibly synthesise. Icons are a designer input.
* **No file watching.** `notify` is compiled out for iOS; `LogWatcher` polls
  the file size instead, like the wasm build. A sandboxed app sees only its own
  container, and documents arrive through the picker.
* **No cross-compilation from Linux.** `cargo check` for the iOS cfg paths
  works anywhere (see `IOS_GUIDELINES.md` §9); producing an app needs macOS.
  Zig does not change this — it has no Apple SDK — but `scripts/probe_zig_ios.sh`
  demonstrates exactly where it stops, and `IOS_GUIDELINES.md` §8b explains
  both that and the one thing that genuinely works without an SDK (building the
  Rust side as an rlib, no linker involved).
