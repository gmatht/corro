//! Generate the ObjC forwarding shims a macOS app needs, from the same
//! generator rswidgets ships — the Apple counterpart of
//! `android/corro/build.rs`, and the platform twin of `ios/corro/build.rs`.
//!
//! A macOS app needs a handful of Objective-C classes that live *in the app
//! bundle* (`CorroMacTarget`, `CorroMacAlert`, `CorroMacText`, the layout
//! category on `NSView`) because the Rust backend reaches them through the
//! Objective-C runtime. Most are pure forwarding: the Rust side already names
//! the selector and its exact signature, so restating it by hand is two
//! independent declarations of one ABI — and a drift between them is not a
//! compile error but a runtime crash (the mismatched-`objc_msgSend` register
//! corruption `MACOS_GUIDELINES.md` §5 warns about).
//!
//! `rustxWidgets/rswidgets/src/apple_generator.rs` holds the signatures as
//! data and emits the `.h`/`.m`; this script just runs it.
//!
//! Dependency-free by construction: `#[path]`-including the generator compiles
//! the very source downstream apps call as
//! `rswidgets::apple_generator::run()` (the vendored-crate pattern
//! `MACOS_GUIDELINES.md` §1 documents). It is also included by rswidgets' own
//! `build.rs`, so the two cannot drift.
//!
//! Off by default: the generator runs only when `RSWIDGETS_APPLE_PROJECT`
//! names an Apple project root. `--features generate-apple-shims` points it at
//! this crate's project root, so `build_macos.sh` can just enable that feature.
//!
//! What is NOT generated, deliberately: `CorroSheetView` (the canvas class)
//! and `CorroMacText` (the text measurer). Their bodies are behaviour —
//! `isFlipped`, coordinate conversion, the laid-out-size report, the
//! `NSFont` resolution table — and a wrong body there is a layout bug rather
//! than a crash. Both stay in `app/CorroMacShims.m`; the generated header
//! declares their signatures so a mismatch is caught at compile time.

#[path = "../../rustxWidgets/rswidgets/src/apple_generator.rs"]
mod apple_generator;

fn main() {
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-changed=../../rustxWidgets/rswidgets/src/apple_generator.rs");
    println!(
        "cargo:rerun-if-env-changed={}",
        apple_generator::PROJECT_ENV
    );

    // `--features generate-apple-shims` sets the variable to this crate's
    // project root, so build_macos.sh does not have to export it.
    if cfg!(feature = "generate-apple-shims")
        && std::env::var(apple_generator::PROJECT_ENV).is_err()
    {
        // The project root is this crate's own directory
        // (macos/corro -> macos/corro).
        let root = std::env::var("CARGO_MANIFEST_DIR").unwrap_or_else(|_| ".".into());
        std::env::set_var(apple_generator::PROJECT_ENV, root);
    }

    // The macOS app's configuration. The canvas and text classes are
    // hand-written and richer than the generated defaults here (the `NSFont`
    // table, the diagnostic tracing), so this app generates only the
    // forwarding shims + the declarations. That is the `override` half of
    // "generated with manual overrides": delete a file and the generator
    // fills it in. See rswidgets::apple_generator::CanvasViewConfig /
    // TextShimConfig for every knob.
    let mut config = apple_generator::ShimConfig::for_platform(apple_generator::Platform::Macos);
    config.canvas = apple_generator::CanvasViewConfig::macos("CorroSheetView");
    config.text = apple_generator::TextShimConfig::macos("CorroMacText");
    config.emit_canvas = false;
    config.emit_text = false;

    if let Some(root) = apple_generator::project_root_from_env() {
        if let Err(e) =
            apple_generator::generate_with(&root, apple_generator::Platform::Macos, &config)
        {
            println!("cargo:warning=rswidgets Apple shim generation failed ({e})");
        }
    } else {
        println!(
            "cargo:warning=rswidgets Apple shim generation skipped; set {} to an Apple project root",
            apple_generator::PROJECT_ENV
        );
    }
    println!(
        "cargo:rerun-if-env-changed={}",
        apple_generator::PROJECT_ENV
    );
}
