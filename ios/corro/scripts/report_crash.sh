#!/bin/bash
# Print what actually went wrong when the app dies on launch.
#
# An .ips crash report is a JSON header line followed by a JSON payload. The
# fields that explain the failure are `exception`, `termination` and the
# faulting thread's first frames - a plain `head` shows only bookkeeping
# (bundle id, pid, paths), which is how an earlier CI run "reported" a crash
# without saying anything useful.
#
# Usage: report_crash.sh [udid]
set -uo pipefail
UDID="${1:-}"

echo "--- crash report ---"
LATEST="$(ls -t ~/Library/Logs/DiagnosticReports/*.ips 2>/dev/null | head -1)"
if [ -z "$LATEST" ]; then
  echo "(no .ips report - the process may have exited rather than crashed)"
else
  echo "== $LATEST"
  python3 - "$LATEST" <<'PY'
import json, sys
lines = open(sys.argv[1]).read().split("\n", 1)
hdr = json.loads(lines[0]) if lines[0].strip() else {}
body = {}
if len(lines) > 1 and lines[1].strip():
    try:
        body = json.loads(lines[1])
    except json.JSONDecodeError:
        print("(payload not parseable)")
merged = dict(hdr)
merged.update(body)
for k in ("exception", "termination", "asi", "vmregioninfo"):
    if k in merged:
        print(k, "=", json.dumps(merged[k])[:800])
faulting = body.get("faultingThread")
threads = body.get("threads") or []
if isinstance(faulting, int) and faulting < len(threads):
    print("--- faulting thread (first frames) ---")
    for fr in (threads[faulting].get("frames") or [])[:16]:
        print("   ", fr.get("symbol") or fr.get("imageOffset"))
PY
fi

echo "--- simulator log (errors) ---"
if [ -n "$UDID" ]; then
  xcrun simctl spawn "$UDID" log show --last 3m --style compact \
    --predicate 'process == "corro"' 2>/dev/null \
    | grep -iE "error|fault|exception|abort|terminat|crash|assert|fatal|library not loaded" \
    | head -25 || true
fi
