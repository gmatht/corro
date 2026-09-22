#!/bin/bash
# Every custom ObjC selector the Rust backends send must exist in a shim.
#
# Objective-C raises `unrecognized selector sent to instance` for a missing
# method - it does NOT ignore the message. A missing selector therefore crashes
# the app, and it is invisible to every Rust-side check. This cost several
# simulator runs to find (create_box sends `corroSetSpacing:`, which nothing
# implemented), so it is checked here, on Linux, in a second.
#
# Covers BOTH Apple backends. The iOS side has a host (`ios/corro`) with real
# shim files to check against; the macOS side has no host yet, so for it the
# check is the other half of the same contract: the *generator* must be able to
# emit every selector the macOS adapter sends. That is what makes a macOS host
# a "run the generator" step rather than a "reinvent the shims" step, and it
# fails here if the macOS adapter grows a selector the generator does not know.
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
CORRO_ROOT="$(cd "$HERE/../../.." && pwd)"

CORRO_ROOT="$CORRO_ROOT" python3 - <<'PY'
import os, re, sys

root = os.environ["CORRO_ROOT"]
adapter = open(os.path.join(root, "rustxWidgets/rswidgets/src/backends_ios_adapter.rs")).read()
# BOTH files implement this contract: CorroGeneratedShims.m is the generator's
# output (the forwarding shims) and CorroIosShims.m is the hand-written half
# (canvas and text measurer). Checking only one reports false positives for
# everything the other owns.
shim = ""
for rel in ("ios/corro/app/CorroGeneratedShims.m", "ios/corro/app/CorroIosShims.m"):
    path = os.path.join(root, rel)
    if os.path.exists(path):
        shim += open(path).read()

# Selectors the backend sends, in the `corro*` namespace it owns.
sent = set(re.findall(r'"(corro[A-Za-z]+:?)"', adapter))

implemented = set()
for m in re.finditer(r"[-+]\s*\([^)]*\)\s*([A-Za-z_][A-Za-z0-9_]*)", shim):
    implemented.add(m.group(1))
for m in re.finditer(r"([A-Za-z_][A-Za-z0-9_]*):\s*\(", shim):
    implemented.add(m.group(1))

missing = [s for s in sorted(sent)
           if s not in implemented and s.rstrip(":") not in implemented]
if missing:
    print("selectors the backend sends with no shim implementation:")
    for m in missing:
        print("   ", m)
    sys.exit(1)
print(f"iOS: all {len(sent)} custom selectors implemented")

# ---------------------------------------------------------------------------
# macOS: no host crate exists yet, so check the contract from the other end.
# ---------------------------------------------------------------------------
macos_adapter = open(
    os.path.join(root, "rustxWidgets/rswidgets/src/backends_macos_adapter.rs")
).read()
generator = open(
    os.path.join(root, "rustxWidgets/rswidgets/src/apple_generator.rs")
).read()

# Every selector the macOS adapter sends, plus the two text-shim selectors the
# generator declares with parameter names (`measure:(id)text font:(id)...`), so
# compare on the colon-separated pieces.
sent_macos = set()
for m in re.finditer(r'"([a-zA-Z][A-Za-z0-9]*(?::[a-z][A-Za-z0-9]*)*:?)"', macos_adapter):
    sel = m.group(1)
    if sel.startswith("corro") or sel in (
        "targetWithCallbackId:", "corroFired:",
        "measure:font:size:slant:weight:",
        "drawText:ctx:font:x:y:size:r:g:b:a:slant:weight:",
    ):
        sent_macos.add(sel)

def in_generator(sel):
    if sel in generator:
        return True
    # The emitted declaration splits the selector across parameters, so a
    # prefix match on the first piece plus the same number of pieces is the
    # right test.
    pieces = [p for p in sel.split(":") if p]
    return all(f"{p}:" in generator for p in pieces) if len(pieces) > 1 else sel in generator

missing_macos = [s for s in sorted(sent_macos) if not in_generator(s)]
if missing_macos:
    print("selectors the macOS adapter sends that the generator cannot emit")
    print("(a macOS host would crash on these - add them to apple_generator.rs):")
    for m in missing_macos:
        print("   ", m)
    sys.exit(1)
print(f"macOS: all {len(sent_macos)} custom selectors the adapter sends are generatable")
PY
