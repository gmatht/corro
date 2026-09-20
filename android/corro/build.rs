//! Generate the Android resources this app needs, from the same generator
//! rswidgets ships.
//!
//! The APK's resources must exist before `build_apk.sh` runs `aapt2` over
//! `app/src/main/res`; generating them from the build script is what makes
//! the icon, theme, palette, strings and manifest reproducible rather than
//! hand-copied. Generation is additive: any file this app already ships is
//! left untouched, so this script only fills gaps — delete a resource to
//! have it recreated.
//!
//! Dependency-free by construction: `#[path]`-including
//! `rustxWidgets/rswidgets/src/android_generator.rs` compiles the very
//! generator downstream apps call as `rswidgets::android_generator::run()`
//! (see `rustxWidgets/docs/ANDROID_GUIDELINES.md` §1, which documents this
//! pattern for vendored crates). The same source is also included by
//! rswidgets' own build.rs, so there is one implementation and no version
//! skew — and, unlike a `[build-dependencies]` entry on rswidgets, no
//! cargo feature-unification side effect on the library corro links (see
//! the note at the bottom of Cargo.toml).
//!
//! Off by default: the generator runs only when `RSWIDGETS_ANDROID_PROJECT`
//! names an Android project root (the variable rswidgets itself uses).
//! `--features generate-android-resources` sets it to this crate's project
//! root, so `build_apk.sh` can just enable that feature; exporting the
//! variable yourself works the same way and can point at another project.
//!
//! See `rustxWidgets/docs/ANDROID_GUIDELINES.md` §1.

#[path = "../../rustxWidgets/rswidgets/src/android_generator.rs"]
mod android_generator;

use std::path::{Path, PathBuf};

fn main() {
    println!("cargo:rerun-if-changed=build.rs");
    println!(
        "cargo:rerun-if-env-changed={}",
        android_generator::PROJECT_ENV
    );

    // The feature is the opt-in for *this* project; an explicit variable
    // always wins so the same script can target another tree.
    let requested = match std::env::var_os(android_generator::PROJECT_ENV) {
        Some(explicit) => PathBuf::from(explicit),
        None if cfg!(feature = "generate-android-resources") => match project_root() {
            Some(root) => root,
            None => {
                println!(
                    "cargo:warning=corro_android: CARGO_MANIFEST_DIR unset; \
                     cannot resolve the Android project root"
                );
                return;
            }
        },
        None => {
            println!(
                "cargo:warning=corro_android: Android resource generation off \
                 (set {} or build with --features generate-android-resources)",
                android_generator::PROJECT_ENV
            );
            return;
        }
    };

    std::env::set_var(android_generator::PROJECT_ENV, &requested);
    println!(
        "cargo:warning=corro_android: generating Android resources into {}",
        requested.display()
    );
    // Prints each created path as a warning and never fails the build (a
    // missing or read-only project is not a Rust build error).
    android_generator::run();
}

/// Android project root = the directory that *contains* `app/` (that is what
/// `rswidgets::android_generator` expects: it writes
/// `<root>/app/src/main/...`). This crate sits *inside* that root as
/// `<root>/src`, so the manifest directory's parent is the root — i.e.
/// `android/corro/Cargo.toml` -> `android/corro`, which holds `app/`.
fn project_root() -> Option<PathBuf> {
    let manifest = std::env::var_os("CARGO_MANIFEST_DIR")?;
    Some(Path::new(&manifest).to_path_buf())
}
