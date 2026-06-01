#!/usr/bin/env bash
# Build the tiny profile and strip the binary.
set -euo pipefail
cd "$(dirname "$0")"
cargo build --profile tiny "$@"
strip target/tiny/corro
echo "  stripped target/tiny/corro"
ls -lh target/tiny/corro
