//! Generate the Android resources an rswidgets app needs, when the app has
//! not provided its own.
//!
//! This script lives in the rswidgets crate, so it cannot `use rswidgets`
//! (a crate's build script runs before, and independently of, the crate).
//! The generator therefore lives in `src/android_generator.rs` for downstream
//! apps to call from *their* build script, and this script includes the same
//! source directly so an app that vendors rswidgets still gets resources
//! without adding a build-dependency.
//!
//! Opt-in: set `RSWIDGETS_ANDROID_PROJECT` to the Android project root (the
//! directory containing `app/`). Without the variable nothing is written, so
//! a plain desktop build is unaffected. Generation is additive: existing
//! files are never overwritten.
//!
//! See `docs/ANDROID_GUIDELINES.md`.

#[path = "src/android_generator.rs"]
mod android_generator;

fn main() {
    println!("cargo:rerun-if-changed=build.rs");
    println!("cargo:rerun-if-changed=src/android_generator.rs");
    android_generator::run();
}
