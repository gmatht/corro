#!/usr/bin/env bash
# Package the Windows 95 release zips for corro 0.7.0:
#
#   gcorro-win95.exe.zip       (native nwg GUI; inner file: gcorro.exe)
#   corro-no-gui-win95.exe.zip (pancurses console TUI; inner file: corro.exe)
#
# The exes themselves are built by ./build95-gui.sh --release and
# ./build95.sh --pancurses --release (rust9x toolchain + VC6 libs + zig are
# local-only; they cannot run on CI runners, so this packaging step is also
# local). This script only verifies + zips whatever is in dist/.
#
# Each zip also carries README95.txt (unicows.dll requirement — it cannot be
# bundled for licensing reasons).
set -euo pipefail
cd "$(dirname "$0")/.."

GUI_SRC="dist/corro-win95-gui.exe"
TUI_SRC="dist/corro-win95.exe"
[ -f "$GUI_SRC" ] || { echo "ERROR: $GUI_SRC missing — run ./build95-gui.sh --release first" >&2; exit 1; }
[ -f "$TUI_SRC" ] || { echo "ERROR: $TUI_SRC missing — run ./build95.sh --pancurses --release first" >&2; exit 1; }

# Sanity: rust9x emits DllCharacteristics 0x8140, which the Win95 loader
# rejects ("not a valid Win32 application"); both build scripts zero it.
# Subsystem must be 2 (GUI) for gcorro, 3 (console) for corro.
# SKIP_PE_CHECK=1 bypasses this (packaging self-test only, never releases).
if [ "${SKIP_PE_CHECK:-0}" != 1 ]; then
python3 - "$GUI_SRC" "$TUI_SRC" <<'EOF'
import struct, sys
for path, want_subsys in [(sys.argv[1], 2), (sys.argv[2], 3)]:
    d = open(path, "rb").read()
    pe = struct.unpack("<I", d[0x3c:0x40])[0]
    subsys, dlc = struct.unpack_from("<HH", d, pe + 24 + 68)
    tag = "OK " if (subsys == want_subsys and dlc == 0) else "FAIL"
    print(f"{tag} {path}: subsystem={subsys} (want {want_subsys}) DllCharacteristics=0x{dlc:04x}")
    if tag == "FAIL":
        sys.exit(1)
EOF
fi

VERSION=$(grep -m1 '^version' Cargo.toml | sed -E 's/.*"([^"]+)".*/\1/')
cat > dist/README95.txt <<EOF
corro $VERSION for Windows 95 (i586)
====================================

gcorro.exe            Native Windows GUI (needs no console window).
corro.exe             Console spreadsheet UI (runs in a DOS box).

REQUIREMENT: unicows.dll (Microsoft Layer for Unicode) must sit next to
the exe, or in C:\\WINDOWS\\SYSTEM. It cannot be bundled here for licensing
reasons. Get it from Microsoft's "Microsoft Layer for Unicode on Windows
95/98/Me Systems" redistributable (unicows.dll). The Wine project's
unicows.dll replacement also works.

Keys in the console UI: arrows move, letters edit, Enter commits, Esc quits.
EOF

rm -f dist/gcorro-win95.exe.zip dist/corro-no-gui-win95.exe.zip
cp "$GUI_SRC" dist/gcorro.exe
cp "$TUI_SRC" dist/corro.exe
( cd dist && zip -9 gcorro-win95.exe.zip gcorro.exe README95.txt )
( cd dist && zip -9 corro-no-gui-win95.exe.zip corro.exe README95.txt )
rm -f dist/gcorro.exe dist/corro.exe
echo "-- release zips:"
ls -lh dist/gcorro-win95.exe.zip dist/corro-no-gui-win95.exe.zip
md5sum dist/gcorro-win95.exe.zip dist/corro-no-gui-win95.exe.zip
