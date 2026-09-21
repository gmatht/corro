#!/bin/bash
# Type-check corro's iOS cfg paths (gui::ios_backend et al) without macOS.
#
# Same shape as check_rswidgets_ios.sh: a scratch crate depending on corro with
# the `gui-mobile` feature set — exactly what ios/corro uses — checked with
# `-Zbuild-std` so the target's core/std exist. No link, so no iOS SDK.
#
# Usage: scripts/check_corro_ios.sh [target ...]
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
CORRO_ROOT="$(cd "$HERE/../../.." && pwd)"
TARGETS=("${@:-aarch64-apple-ios-sim aarch64-apple-ios armv7s-apple-ios}")
# shellcheck disable=SC2206
TARGETS=(${TARGETS[@]})

for TARGET in "${TARGETS[@]}"; do
  SCRATCH="${TMPDIR:-/tmp}/corro_ios_check_$TARGET"
  rm -rf "$SCRATCH"; mkdir -p "$SCRATCH/src"
  cat > "$SCRATCH/Cargo.toml" <<TOML
[package]
name = "corro_ios_check"
version = "0.0.0"
edition = "2021"

[workspace]

[dependencies]
corro = { path = "$CORRO_ROOT", default-features = false, features = ["gui-mobile"] }
TOML
  cat > "$SCRATCH/src/lib.rs" <<'RS'
//! Compile-only harness for corro's iOS cfg paths.
#[allow(unused_imports)]
use corro::gui::ios_backend as _ios;
RS
  echo "=== corro $TARGET ==="
  (cd "$SCRATCH" && cargo +nightly check --target "$TARGET" -Zbuild-std=std,panic_abort 2>&1 | tail -5)
done
