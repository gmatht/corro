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
# Filter by process name. Taking simply the newest .ips picked up unrelated
# system components (AccessibilityControlsExtension on one run), and its
# EXC_BREAKPOINT was then read as though it were corro's - which sent an
# investigation after a crash that had nothing to do with this app.
LATEST=""
for f in $(ls -t ~/Library/Logs/DiagnosticReports/*.ips 2>/dev/null | head -20); do
  case "$(basename "$f")" in
    corro-*|corro_*|*/corro-*) LATEST="$f"; break ;;
  esac
done
if [ -z "$LATEST" ]; then
  echo "(no corro .ips report - the process exited without crashing)"
  echo "recent reports (other processes, for context):"
  ls -t ~/Library/Logs/DiagnosticReports/*.ips 2>/dev/null | head -5 | sed 's/^/    /'
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
else:
    print("(faultingThread not usable:", faulting, "of", len(threads), "threads)")
# Belt and braces: the raw head of the report, so the frames are visible
# whatever shape the payload takes (the structured extraction came back empty
# once, which is worse than useless).
print("--- raw report head ---")
for line in open(sys.argv[1]).read().splitlines()[:1]:
    pass
# Print the faulting thread raw: the exact frame list matters more than any
# parse, and a structured extraction that silently yields nothing is worse than
# no extraction at all.
ft = body.get("faultingThread")
images = body.get("usedImages") or []
threads = body.get("threads") or []
if isinstance(ft, int) and ft < len(threads):
    th = threads[ft]
    print("--- faulting thread raw ---")
    for fr in (th.get("frames") or [])[:20]:
        idx = fr.get("imageIndex")
        img = images[idx].get("name") if isinstance(idx, int) and idx < len(images) else "?"
        print("   ", img, fr.get("symbol") or fr.get("imageOffset") or fr)
print("--- report keys ---")
print("   ", sorted(body.keys())[:40])
PY
fi

echo "--- simulator log (errors) ---"
if [ -n "$UDID" ]; then
  xcrun simctl spawn "$UDID" log show --last 3m --style compact \
    --predicate 'process == "corro"' 2>/dev/null \
    | grep -iE "error|fault|exception|abort|terminat|crash|assert|fatal|library not loaded" \
    | head -25 || true
fi
