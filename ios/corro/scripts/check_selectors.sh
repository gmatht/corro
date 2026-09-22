#!/bin/bash
# Every custom ObjC selector the Rust backend sends must exist in the shim.
#
# Objective-C raises `unrecognized selector sent to instance` for a missing
# method - it does NOT ignore the message. A missing selector therefore crashes
# the app, and it is invisible to every Rust-side check. This cost several
# simulator runs to find (create_box sends `corroSetSpacing:`, which nothing
# implemented), so it is checked here, on Linux, in a second.
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
CORRO_ROOT="$(cd "$HERE/../../.." && pwd)"

CORRO_ROOT="$CORRO_ROOT" python3 - <<'PY'
import os, re, sys

root = os.environ["CORRO_ROOT"]
adapter = open(os.path.join(root, "rustxWidgets/rswidgets/src/backends_ios_adapter.rs")).read()
shim = open(os.path.join(root, "ios/corro/app/CorroIosShims.m")).read()

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
print(f"all {len(sent)} custom selectors implemented")
PY
