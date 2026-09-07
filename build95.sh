#!/usr/bin/env bash
# Build corro for Windows 95.
#
#   ./build95.sh              build (incremental, win95 profile)
#   ./build95.sh --release    optimized build
#
# Output: dist/corro-win95.exe
#
# Uses the custom `win95` profile: dev-like, but panic = "abort" (unwind
# support needs at least VC8 (VS2005) and we link the VC6 CRT — an unwinding
# panic aborts with "abnormal program termination" on Win95).
#
# Does two things:
#   1. cargo +rust9x build with the i586-rust9x-windows-msvc target.
#      (i586, NOT i686: Windows 95 never enables CR4.OSFXSR, so any SSE
#      instruction faults regardless of the virtual/physical CPU model.)
#      Requires the rust9x link setup in .cargo/config.toml (vintage MSVC
#      toolset) and the Win95 entry-point patches in src/main.rs (no_main:
#      std::rt init hangs on Win95; GetCommandLineA args shim) — those are
#      already in the source, cfg-gated to the rust9x msvc targets.
#   2. Patches DllCharacteristics in the PE optional header to 0.
#      rust9x emits 0x8140, which makes the Windows 95 loader reject the
#      whole executable with "not a valid Win32 application".
#
# Runtime note: on Win9x/ME the exe needs unicows.dll next to it (or in
# C:\WINDOWS\SYSTEM); run95.sh makes sure it is on the VM image.
set -euo pipefail
cd "$(dirname "$0")"

TARGET=i586-rust9x-windows-msvc
PROFILE=win95
ARGS=()
for a in "$@"; do
  case "$a" in
    --release) ARGS+=(--profile release); PROFILE=release ;;
    *) ARGS+=("$a") ;;
  esac
done

cargo +rust9x build --offline --profile "$PROFILE" --target "$TARGET" "${ARGS[@]+"${ARGS[@]}"}"

SRC="target/$TARGET/$PROFILE/corro.exe"
[ -f "$SRC" ] || { echo "ERROR: build output missing: $SRC" >&2; exit 1; }
mkdir -p dist
python3 - "$SRC" dist/corro-win95.exe <<'EOF'
import struct, sys
# Win95 loader: zero DllCharacteristics (rust9x emits 0x8140). NOTE the
# layout: optional-header offset 68 = Subsystem (MUST keep — corro is a
# CONSOLE app; zeroing 68 makes the Win95 loader fall back to the DOS stub
# with "This program cannot be run in DOS mode"), 70 = DllCharacteristics.
d = bytearray(open(sys.argv[1], "rb").read())
pe = struct.unpack("<I", d[0x3c:0x40])[0]
struct.pack_into("<H", d, pe + 24 + 70, 0)
open(sys.argv[2], "wb").write(d)
EOF

echo "-- built dist/corro-win95.exe ($(stat -c%s dist/corro-win95.exe) bytes)"
md5sum dist/corro-win95.exe
