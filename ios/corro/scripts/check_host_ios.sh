#!/bin/bash
# Type-check the corro_ios host cdylib (this crate) without macOS.
#
# Usage: scripts/check_host_ios.sh [target]
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
HOST="$(cd "$HERE/.." && pwd)"
TARGET="${1:-aarch64-apple-ios-sim}"
echo "=== corro_ios $TARGET ==="
(cd "$HOST" && cargo +nightly check --target "$TARGET" -Zbuild-std=std,panic_abort 2>&1 | tail -5)
