#!/bin/bash
# Type-check corro's macOS cfg paths (gui::macos_backend and the shared
# gui_backend compiled against the AppKit adapter) without a macOS host.
#
# Same shape as ios/corro/scripts/check_corro_ios.sh: a scratch crate depending
# on corro with the `gui-macos` feature set, checked with `-Zbuild-std` so the
# target's core/std exist. No link, so no Apple SDK.
#
# This is the check that substantiates the "one corro source, many backends"
# claim for macOS: it compiles the *same* `gui_backend` (spreadsheet, menus,
# dialogs, key handling) and the *same* menu model that the iOS, Android, GTK,
# Windows and TUI builds compile — only the rswidgets adapter underneath is
# AppKit.
#
# `gui-macos` (not `gui-mobile`): `gui` resolves `rswidgets/gtk` +
# `gtk_dynamic_loader`, which cannot link on an Apple target, so the GTK
# feature is exactly the thing a Mac must not enable. `gui-macos` is the
# desktop-native feature that puts corro on the AppKit adapter and nothing
# else. An iOS target is still checked below as a regression guard: the two
# share no corro-side file, but they do share rswidgets' apple runtime module.
#
# Usage: scripts/check_corro_macos.sh [target ...]
#   default targets: x86_64-apple-darwin, aarch64-apple-darwin, and an iOS
#   target as a regression guard (the two share no corro-side file, but they do
#   share rswidgets' apple runtime module)
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
CORRO_ROOT="$(cd "$HERE/.." && pwd)"
TARGETS=("${@:-x86_64-apple-darwin aarch64-apple-darwin aarch64-apple-ios-sim}")
# shellcheck disable=SC2206
TARGETS=(${TARGETS[@]})

for TARGET in "${TARGETS[@]}"; do
  # An iOS target needs rswidgets/ios (the UIKit adapter), a macOS one needs
  # rswidgets/macos. Both are no-op features off-target, so selecting the
  # wrong one is harmless — but selecting the RIGHT one is the point, and
  # `gui-macos` alone would not pull in the UIKit adapter on an iOS target.
  case "$TARGET" in
    *-apple-ios*) FEATURES='["gui-mobile"]' ;;
    *)            FEATURES='["gui-macos"]' ;;
  esac
  SCRATCH="${TMPDIR:-/tmp}/corro_macos_check_$TARGET"
  rm -rf "$SCRATCH"; mkdir -p "$SCRATCH/src"
  cat > "$SCRATCH/Cargo.toml" <<TOML
[package]
name = "corro_macos_check"
version = "0.0.0"
edition = "2021"

[workspace]

[dependencies]
corro = { path = "$CORRO_ROOT", default-features = false, features = $FEATURES }
TOML
  cat > "$SCRATCH/src/lib.rs" <<'RS'
//! Compile-only harness for corro's macOS cfg paths.
//!
//! Both imports are deliberate: `macos_backend` is the module under test, and
//! `ios_backend` is the iOS twin, which `macos_backend` re-exports its menu
//! model and action dispatcher from. If that re-export ever stops compiling,
//! this harness fails with it.
#[allow(unused_imports)]
use corro::gui::macos_backend as _macos;
#[allow(unused_imports)]
use corro::gui::ios_backend as _ios;
RS
  echo "=== corro $TARGET (features: $FEATURES) ==="
  (cd "$SCRATCH" && cargo +nightly check --target "$TARGET" -Zbuild-std=std,panic_abort 2>&1 | tail -5)
done
