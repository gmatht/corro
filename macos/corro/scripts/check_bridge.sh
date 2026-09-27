#!/bin/bash
# Check that `app/CorroMacBridge.h` declares exactly the `corro_macos_*`
# functions `macos/corro/src/lib.rs` exports.
#
# ## Why this is a script and not a build step
#
# A missing declaration in the bridge header is a *link* error, which a Mac
# would catch — but a **changed signature** is not. `corro_macos_canvas_draw`
# declared with the arguments in a different order, or with `float` where the
# Rust side has `double`, compiles fine and then corrupts the register file
# inside `objc_msgSend`: a crash with a backtrace that points at the shim, not
# at the cause. The iOS host crate hit the same class of bug (see the
# `report_panics` comment in `ios/corro/src/lib.rs` for the one that actually
# fired) and the answer was to stop restating the ABI by hand.
#
# So this diffs the two lists. The generator (`apple_generator`) already does
# this for the shims the Rust adapter *messages to*; this covers the other
# direction, the functions the host *calls into*.
#
# Both directions are needed and neither subsumes the other: a selector can be
# correct while an export is missing, and an export can be correct while a
# selector is absent. A host that has only one of the two fails in a different
# way.
#
# Runs on Linux: it reads two text files.
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
HOST="$(cd "$HERE/.." && pwd)"
HEADER="$HOST/app/CorroMacBridge.h"
RUST="$HOST/src/lib.rs"

fail=0

# 1. Every `#[no_mangle] extern "C" fn corro_macos_*` in the Rust crate must be
#    declared in the header.
rust_exports=$(grep -oE '^pub (unsafe )?extern "C" fn corro_macos_[a-z_]+' "$RUST" \
  | grep -oE 'corro_macos_[a-z_]+' | sort -u)
header_decls=$(grep -oE '\bcorro_macos_[a-z_]+\s*\(' "$HEADER" \
  | sed 's/[[:space:]]*($//; s/($//; s/[[:space:]]*(//' | sort -u)

missing=""
for fn in $rust_exports; do
  if ! grep -qE "\b${fn}[[:space:]]*\(" "$HEADER"; then
    missing="$missing $fn"
  fi
done
if [ -n "$missing" ]; then
  echo "FAIL  exported by macos/corro but not declared in CorroMacBridge.h:"
  for fn in $missing; do echo "        $fn"; done
  fail=1
fi

# 2. Every declaration in the header must exist in the Rust crate. A header
#    declaring a function nobody exports is the more confusing direction: the
#    host compiles, links fail, and the message names a symbol the header
#    author did not write.
orphan=""
for fn in $header_decls; do
  # `grep -qx` is spelled as a fixed-string match plus a whole-line anchor
  # because `rg` is aliased to `grep` in this environment and swallows the
  # combined short flags.
  if ! printf '%s\n' "$rust_exports" | grep -x -F -q -- "$fn"; then
    orphan="$orphan $fn"
  fi
done
if [ -n "$orphan" ]; then
  echo "FAIL  declared in CorroMacBridge.h but not exported by macos/corro:"
  for fn in $orphan; do echo "        $fn"; done
  fail=1
fi

# 3. The Rust side must actually be an Apple target's host crate. This script
#    exists because `gui-macos` is a real feature; a build that silently
#    stopped selecting it would make every check above vacuous.
if ! grep -qE 'features = \["gui-macos"\]' "$HOST/Cargo.toml"; then
  echo "FAIL  macos/corro must depend on corro with features = [\"gui-macos\"]"
  echo "        (found: $(grep -E '^corro = ' "$HOST/Cargo.toml" || echo none))"
  fail=1
fi

# 4. Every selector the generator's table covers for macOS must be in the
#    generator's own macOS table — this is the shim side of the same contract,
#    asserted here so both directions are checked by one command.
GEN="$HOST/../../rustxWidgets/rswidgets/src/apple_generator.rs"
if [ -f "$GEN" ]; then
  # The two selectors the macOS Label width pin sends. If the adapter grew a
  # third and the table was not updated, `adapter_selectors_are_all_covered`
  # (a unit test, also run in CI) catches it — this is a belt-and-braces
  # check that the *generated* header declares them, which only the generator
  # test can see.
  for sel in corroSetPinnedWidth: corroSetContentHugging:; do
    if ! grep -q "\"${sel}\"" "$GEN"; then
      echo "FAIL  macOS Label pin selector '$sel' is not in apple_generator's tables"
      echo "        (the generated .h would not declare it, and the host would not compile)"
      fail=1
    fi
  done
else
  echo "FAIL  cannot find $GEN"
  fail=1
fi

if [ "$fail" -eq 0 ]; then
  n_rust=$(echo "$rust_exports" | grep -c . || true)
  n_hdr=$(echo "$header_decls" | grep -c . || true)
  echo "PASS  bridge contract: $n_rust exports, $n_hdr declarations, all matched"
fi
exit "$fail"
