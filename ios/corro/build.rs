//! Generate the ObjC forwarding shims this app needs, from the same generator
//! rswidgets ships — the Apple counterpart of `android/corro/build.rs`.
//!
//! The iOS target needs a handful of Objective-C classes that live *in the app
//! bundle* (`CorroIosTarget`, `CorroIosAlert`, the layout category on `UIView`)
//! because the Rust backend reaches them through the Objective-C runtime. They
//! are pure forwarding: the Rust side already names the selector and its exact
//! signature, so restating it by hand is two independent declarations of one
//! ABI — and a drift between them is not a compile error but a runtime crash
//! (the mismatched-`objc_msgSend` corruption `IOS_GUIDELINES.md` warns about).
//!
//! `rustxWidgets/rswidgets/src/apple_generator.rs` holds the signatures as
//! data and emits the `.h`/`.m`; this script just runs it, exactly as
//! `android/corro/build.rs` runs `android_generator`.
//!
//! Dependency-free by construction: `#[path]`-including the generator compiles
//! the very source downstream apps call as `rswidgets::apple_generator::run()`
//! (the vendored-crate pattern `ANDROID_GUIDELINES.md` §1 documents). It is
//! also included by rswidgets' own `build.rs`. Unlike a
//! `[build-dependencies]` entry on rswidgets, this has no cargo
//! feature-unification side effect on the library the app links.
//!
//! Off by default: the generator runs only when `RSWIDGETS_APPLE_PROJECT`
//! names an Apple project root. `--features generate-apple-shims` points it at
//! this crate's project root, so `build_ios.sh` can just enable that feature.
//!
//! What is NOT generated, deliberately: `SheetView` (the canvas class) and
//! `CorroIosText` (the text measurer). Their bodies are behaviour — coordinate
//! conversion, the laid-out-size report, and the SDK version table
//! (`boundingRectWithSize:` vs `sizeWithFont:`) — and a wrong body there is a
//! layout bug rather than a crash. Both stay in `app/CorroIosShims.m`; the
//! generated header declares their signatures so a mismatch is caught at
//! compile time.

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
    // project root, so build_ios.sh does not have to export it.
    if cfg!(feature = "generate-apple-shims") && std::env::var(apple_generator::PROJECT_ENV).is_err() {
        // The project root is this crate's parent (ios/corro -> ios/corro).
        let root = std::env::var("CARGO_MANIFEST_DIR").unwrap_or_else(|_| ".".into());
        std::env::set_var(apple_generator::PROJECT_ENV, root);
    }

    // The iOS app's configuration: the canvas needs the fill-the-superview
    // fallback (it attaches the sheet inside a container, not a stack view) and
    // deploys below the 13.4 that UIKey needs, so the key path is guarded.
    // See rswidgets::apple_generator::CanvasViewConfig for every knob.
    let mut config = apple_generator::ShimConfig::for_platform(apple_generator::Platform::Ios);
    config.canvas = apple_generator::CanvasViewConfig::ios("SheetView");
    config.text = apple_generator::TextShimConfig::ios("CorroIosText");
    // The canvas and text classes are already hand-written and richer than the
    // generated defaults here (the version table, the diagnostic tracing), so
    // this app generates only the forwarding shims + the declarations. That is
    // the `override` half of "generated with manual overrides": delete a file
    // and the generator fills it in.
    config.emit_canvas = false;
    config.emit_text = false;

    if let Some(root) = apple_generator::project_root_from_env() {
        if let Err(e) = apple_generator::generate_with(&root, apple_generator::Platform::Ios, &config) {
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
