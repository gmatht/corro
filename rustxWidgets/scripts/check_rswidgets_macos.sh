#!/bin/bash
# Type-check rswidgets' macOS (AppKit) cfg paths without a macOS host or an
# Apple SDK, and re-check the iOS paths that share the same `apple` runtime
# module.
#
# Why a scratch crate instead of `cargo check` in the workspace:
#   * workspace `[patch.crates-io]` pulls a native `links = "pdcurses"` dep in,
#     which cannot resolve for an Apple target;
#   * a standalone crate with `default-features = false, features = ["macos"]`
#     gets exactly the macOS feature set and nothing else.
#
# `-Zbuild-std` supplies core/std for the target, and `cargo check` is metadata
# only — no link, so no SDK is needed. This is the check that proves the Rust
# half of the AppKit port; producing a runnable .app needs macOS
# (docs/MACOS_GUIDELINES.md).
#
# The AppKit adapter and the iOS adapter both build on `backends/apple.rs`, so
# this script deliberately also checks an iOS target: a change to the shared
# module that breaks one platform must not pass on the other.
#
# Usage: rustxWidgets/scripts/check_rswidgets_macos.sh [target ...]
#   default targets: x86_64-apple-darwin, aarch64-apple-darwin, and the iOS
#   simulator target (shared-module regression guard)
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
RSW="$(cd "$HERE/../rswidgets" && pwd)"
TARGETS=("${@:-x86_64-apple-darwin aarch64-apple-darwin aarch64-apple-ios-sim}")
# shellcheck disable=SC2206
TARGETS=(${TARGETS[@]})

for TARGET in "${TARGETS[@]}"; do
  case "$TARGET" in
    *-apple-ios*) FEATURES="ios"; ENTRY="backends_ios_adapter"; CORE="backends::ios" ;;
    *)            FEATURES="macos"; ENTRY="backends_macos_adapter"; CORE="backends::macos" ;;
  esac
  SCRATCH="${TMPDIR:-/tmp}/rsw_macos_check_$TARGET"
  rm -rf "$SCRATCH"; mkdir -p "$SCRATCH/src"
  cat > "$SCRATCH/Cargo.toml" <<TOML
[package]
name = "rswidgets_macos_check"
version = "0.0.0"
edition = "2021"

[workspace]

[dependencies]
rswidgets = { path = "$RSW", default-features = false, features = ["$FEATURES"] }
TOML
  cat > "$SCRATCH/src/lib.rs" <<RS
//! Compile-only harness: with an Apple target, this compiles the platform
//! adapter and the shared \`backends/apple\` runtime it builds on.
#[allow(unused_imports)]
use rswidgets::$ENTRY as _adapter;
#[allow(unused_imports)]
use rswidgets::$CORE as _core;
#[allow(unused_imports)]
use rswidgets::backends::apple as _apple;
RS
  echo "=== rswidgets $TARGET (features: $FEATURES) ==="
  (cd "$SCRATCH" && cargo +nightly check --target "$TARGET" -Zbuild-std=std,panic_abort 2>&1 | tail -5)
done
