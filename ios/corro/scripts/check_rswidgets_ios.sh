#!/bin/bash
# Type-check rswidgets' iOS cfg paths without a macOS host or an iOS SDK.
#
# Why a scratch crate instead of `cargo check` in the workspace:
#   * workspace `[patch.crates-io]` pulls a native `links = "pdcurses"` dep in,
#     which cannot resolve for an Apple target;
#   * a standalone crate with `default-features = false, features = ["ios"]`
#     gets exactly the iOS feature set and nothing else.
#
# `-Zbuild-std` supplies core/std for the target, and `cargo check` is metadata
# only — no link, so no SDK is needed. This is the check that proves the Rust
# half of the port; producing an app needs macOS (docs/IOS_GUIDELINES.md §9).
#
# Usage: scripts/check_rswidgets_ios.sh [target ...]
#   default targets: all three (simulator, arm64 device, 32-bit armv7s)
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
CORRO_ROOT="$(cd "$HERE/../../.." && pwd)"
RSW="$CORRO_ROOT/rustxWidgets/rswidgets"
TARGETS=("${@:-aarch64-apple-ios-sim aarch64-apple-ios armv7s-apple-ios}")
# shellcheck disable=SC2206
TARGETS=(${TARGETS[@]})

for TARGET in "${TARGETS[@]}"; do
  SCRATCH="${TMPDIR:-/tmp}/rsw_ios_check_$TARGET"
  rm -rf "$SCRATCH"; mkdir -p "$SCRATCH/src"
  cat > "$SCRATCH/Cargo.toml" <<TOML
[package]
name = "rswidgets_ios_check"
version = "0.0.0"
edition = "2021"

[workspace]

[dependencies]
rswidgets = { path = "$RSW", default-features = false, features = ["ios"] }
TOML
  cat > "$SCRATCH/src/lib.rs" <<'RS'
//! Compile-only harness: with `target_os = "ios"`, this compiles every iOS
//! cfg path in rswidgets. Nothing runs; the crate only has to type-check.
#[allow(unused_imports)]
use rswidgets::backends_ios_adapter as _adapter;
#[allow(unused_imports)]
use rswidgets::backends::ios as _core;
RS
  echo "=== rswidgets $TARGET ==="
  (cd "$SCRATCH" && cargo +nightly check --target "$TARGET" -Zbuild-std=std,panic_abort 2>&1 | tail -5)
done
