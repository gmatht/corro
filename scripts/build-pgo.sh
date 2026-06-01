#!/usr/bin/env bash
# Build the release binary with the checked-in PGO profile.
# Profile at pgo/merged.profdata is used when present.
set -euo pipefail
cd "$(dirname "$0")"

PGO_PROFILE="$(pwd)/pgo/merged.profdata"
PROFILE="${PROFILE:-release}"

if [ ! -f "$PGO_PROFILE" ]; then
  echo "ERROR: no PGO profile found at $PGO_PROFILE"
  echo "Generate it: cargo pgo build && ./target/x86_64-unknown-linux-gnu/release/corro tmp_AG.corro && cargo pgo optimize"
  echo "Then copy the merged profile: cp target/pgo-profiles/merged.profdata pgo/"
  exit 1
fi

echo "=== PGO: Building release with $PGO_PROFILE ==="
RUSTFLAGS="-C profile-use=$PGO_PROFILE" RUSTC_WRAPPER= cargo build --profile "$PROFILE" 2>&1 | tail -3

BIN="target/$PROFILE/corro"
if [ ! -f "$BIN" ]; then
  # cargo-pgo may output to target/x86_64-unknown-linux-gnu/release/
  ALT="target/x86_64-unknown-linux-gnu/$PROFILE/corro"
  if [ -f "$ALT" ]; then
    BIN="$ALT"
  fi
fi
echo "PGO build complete: $BIN"
