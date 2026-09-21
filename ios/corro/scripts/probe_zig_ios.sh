#!/bin/bash
# Can Zig replace Xcode for this project? Probe, and print the answer.
#
# The short version: no. This script exists so the claim is reproducible rather
# than folklore — it walks the four steps that matter, in the order they bite,
# and reports each result. See rustxWidgets/docs/IOS_GUIDELINES.md §8b for the
# table and the two non-obvious details (zig's libSystem.tbd has no iOS slice;
# -mios-version-min cannot be honoured through zig's bundled darwin libc).
#
# Usage: scripts/probe_zig_ios.sh [path-to-zig]
set -uo pipefail
ZIG="${1:-zig}"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

if ! command -v "$ZIG" >/dev/null; then
  echo "zig not found (pass a path): $ZIG"; exit 1
fi
echo "zig: $("$ZIG" version 2>/dev/null)"
echo

cat > "$TMP/plain.c" <<'EOF'
int main(void) { return 0; }
EOF
cat > "$TMP/uikit.c" <<'EOF'
#include <UIKit/UIKit.h>
int main(void) { return 0; }
EOF
cat > "$TMP/objc.m" <<'EOF'
#import <Foundation/Foundation.h>
@interface CorroProbe : NSObject @end
@implementation CorroProbe @end
EOF

step() { # step <label> <cmd...>
  local label="$1"; shift
  if "$@" >"$TMP/out" 2>&1; then
    echo "  OK    $label"
  else
    echo "  FAIL  $label"
    # first line of the diagnostic, so the reason is visible
    head -1 "$TMP/out" | sed 's/^/          /'
  fi
}

echo "1. C for arm64-ios, no frameworks"
step "zig cc -target aarch64-ios -c plain.c" \
  "$ZIG" cc -target aarch64-ios -c "$TMP/plain.c" -o "$TMP/plain.o"
[ -f "$TMP/plain.o" ] && file "$TMP/plain.o" | sed 's/^/        /'

echo "2. UIKit header"
step "zig cc -target aarch64-ios -c uikit.c" \
  "$ZIG" cc -target aarch64-ios -c "$TMP/uikit.c" -o "$TMP/uikit.o"

echo "3. Objective-C with ARC (needs Foundation)"
step "zig cc -target aarch64-ios -fobjc-arc -c objc.m" \
  "$ZIG" cc -target aarch64-ios -fobjc-arc -c "$TMP/objc.m" -o "$TMP/objc.o"

echo "4. Link a Rust iOS staticlib (-lobjc, -framework CoreGraphics, -liconv)"
echo "     -> run ios/corro/build_ios.sh device and watch the linker;"
echo "        zig cc reports: unable to find dynamic system library 'objc'"

echo
echo "What zig ships for darwin:"
# `zig env` prints Zig-style TOML (`.lib_dir = "..."`), not shell syntax.
LIBDIR="$("$ZIG" env 2>/dev/null | awk -F'"' '/\.lib_dir *=/ {print $2; exit}')"
if [ -n "$LIBDIR" ] && [ -f "$LIBDIR/libc/darwin/libSystem.tbd" ]; then
  echo "  $LIBDIR/libc/darwin/$(ls "$LIBDIR/libc/darwin" | tr '\n' ' ')"
  echo "  libSystem.tbd target slices:"
  grep -m1 "targets:" "$LIBDIR/libc/darwin/libSystem.tbd" | sed 's/^/    /'
  echo "  (note: no iOS slice)"
else
  echo "  (could not locate zig's darwin libc)"
fi
echo
echo "The one thing that DOES work without an SDK: building the Rust side as an"
echo "rlib, which needs no linker at all —"
echo "  cargo +nightly rustc --release --target aarch64-apple-ios \\"
echo "      -Zbuild-std=std,panic_abort --lib --crate-type rlib"
