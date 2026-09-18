#!/usr/bin/env bash
# Fail if a linked Linux binary needs a glibc newer than the promised floor.
#
#   tools/check-glibc-floor.sh <binary> [max=2.17]
#
# The release links through tools/zig-cc-gnu217 so the same asset runs on old
# distros incl. CentOS 7 (glibc 2.17). Verifying that requires comparing the
# *numeric* symbol versions, not a regex: the previous inline check
# (`GLIBC_2\.([3-9][0-9]|[0-9]{3,})`) only matched 2.30+/2.300+, so a binary
# needing 2.18-2.29 passed a gate that claimed to catch it. Versions are also
# three-part (`GLIBC_2.2.5`), which that regex cannot express at all.
set -euo pipefail

BIN="${1:?usage: check-glibc-floor.sh <binary> [max-version]}"
MAX="${2:-2.17}"

[ -f "$BIN" ] || { echo "error: no such file: $BIN" >&2; exit 2; }

# Highest GLIBC_x.y[.z] symbol the binary needs, by numeric version compare.
HIGHEST=$(
  objdump -T "$BIN" 2>/dev/null \
    | grep -oE 'GLIBC_[0-9]+(\.[0-9]+)*' \
    | sed 's/^GLIBC_//' \
    | sort -u \
    | python3 -c '
import sys
def key(s):
    return tuple(int(p) for p in s.strip().split("."))
vers = [l.strip() for l in sys.stdin if l.strip()]
print(max(vers, key=key) if vers else "")
'
)

if [ -z "$HIGHEST" ]; then
  echo "error: no GLIBC symbol versions found in $BIN (not a dynamically linked ELF?)" >&2
  exit 2
fi

if python3 -c "
import sys
def key(s): return tuple(int(p) for p in s.split('.'))
a, b = key('$HIGHEST'), key('$MAX')
sys.exit(0 if a > b else 1)
"; then
  echo "error: $BIN requires GLIBC_$HIGHEST but the floor is GLIBC_$MAX" >&2
  exit 1
fi

echo "ok: $BIN needs at most GLIBC_$HIGHEST (floor GLIBC_$MAX)"
