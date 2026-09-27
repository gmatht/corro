corro on macOS
==============

The desktop host for corro: the `corro_macos` static library that an AppKit app
links against, and the `Corro*` shims that bridge Objective-C to it.

    macos/corro/
    ├── Cargo.toml           corro_macos staticlib (standalone workspace)
    ├── src/lib.rs           the 20 corro_macos_* extern "C" entry points
    ├── build.rs             runs rswidgets' apple_generator to emit the shims
    └── app/
        ├── CorroMacBridge.h        the C surface (diffed against src/lib.rs)
        ├── CorroGeneratedShims.h   GENERATED - signatures the host adopts
        ├── CorroGeneratedShims.m   GENERATED - forwarding shims
        └── (CorroMacShims.m, the app delegate, the Xcode project)

What is here and what is not
----------------------------

**Here**: everything the Rust side owns. The `extern "C"` surface the shims
call, the generated forwarding shims (`CorroMacTarget`, `CorroMacAlert`, the
`NSView` layout category), and the signature declarations for the two
hand-written classes a host must supply.

**Not here**: the app bundle, the Xcode project, the icons, and the two
hand-written classes themselves — `CorroSheetView` (the canvas view) and
`CorroMacText` (the `NSFont` measurer/drawer). Their bodies are *behaviour*,
not forwarding: `isFlipped`, coordinate conversion, the laid-out-size report,
the font resolution table. A generator cannot validate them without an SDK, and
a wrong body there is a layout bug rather than a crash. See
`rustxWidgets/docs/MACOS_GUIDELINES.md` §5.

This is the same split the iOS host (`ios/corro`) makes, and the reason is the
same: the Rust side cannot synthesise a bundle, and an `NSApplication` needs a
window server, so a *runnable* `.app` needs a Mac with Xcode.

Verifying without a Mac
-----------------------

Everything that does not need Xcode is scripted, so the whole Rust half is
verifiable on Linux and in CI (`.github/workflows/macos.yml`):

    scripts/check_bridge.sh              # CorroMacBridge.h vs the Rust exports

    ../scripts/check_corro_macos.sh       # corro::gui::macos_backend
    ../../rustxWidgets/scripts/check_rswidgets_macos.sh   # the AppKit adapter

    cargo test --features gui-macos-host --test macos_pipeline_preview

That last one is the macOS twin of `tests/ios_pipeline_preview.rs`: corro's
half of the host contract, headless. It covers the menu model a real `NSMenu`
is built from, every `app.<action>` name mapping back to a real menu action,
the macOS and iOS models being the *same* model, and the pointer metrics
(cell size, pinch clamp) a mouse host depends on.

The one thing Linux cannot check is the ABI, because `cargo check` does not
link. `.github/workflows/macos.yml`'s `build-macos` job builds the staticlib on
a Mac and links an Objective-C file that references every `corro_macos_*`
symbol — a missing export is a link error, and a changed signature is the
runtime `objc_msgSend` corruption the generator exists to prevent.

Regenerating the shims
----------------------

The generated pair is checked in, and CI fails if it is stale (a stale `.m` is
a *runtime* crash, not a compile error, so nothing else would notice):

    RSWIDGETS_APPLE_PROJECT="$PWD" cargo +nightly build \
      --target aarch64-apple-darwin -Zbuild-std=std,panic_abort \
      --features generate-apple-shims

Generation is additive: an existing file is never overwritten, so a
hand-edited `.m` survives and "delete it to regenerate" is the escape hatch.
