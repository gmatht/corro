#!/usr/bin/env bash
# Build corro's nwg (native Windows GUI) backend for Windows 95.
#
#   ./build95-gui.sh              build (incremental, win95 profile)
#   ./build95-gui.sh --release    optimized build
#
# Output: dist/corro-win95-gui.exe
#
# This is the GUI counterpart of build95.sh (which builds the console TUI).
# Differences from the console build:
#
#   1. `--features gui,gui-subsystem` instead of `pancurses`. The `gui`
#      feature selects rswidgets' nwg backend on Windows; `gui-subsystem`
#      selects the WinMain entry point in src/main.rs.
#   2. The link uses /SUBSYSTEM:WINDOWS instead of the /SUBSYSTEM:CONSOLE set
#      in .cargo/config.toml (which is there for the TUI builds and must stay
#      for them). A GUI-subsystem PE is entered through WinMainCRTStartup ->
#      WinMain; see src/main.rs.
#   3. DllCharacteristics is zeroed (Win95 rejects rust9x's 0x8140 with
#      "not a valid Win32 application"). Subsystem is left alone.
#
# The exe is named gcorro* because corro's argv[0] dispatch prefers the GUI
# when the program name starts with "gcorro" (see argv0_ui in src/main.rs), so
# the GUI backend is selected without needing --gui.
#
# Runtime note: as for any rust9x binary on Win9x/ME, unicows.dll must be
# next to the exe (or in C:\WINDOWS\SYSTEM). run95-gui.sh handles that.
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

# GUI subsystem for this build only.
#
# This must NOT go through plain $RUSTFLAGS / CARGO_ENCODED_RUSTFLAGS: both
# *replace* the target rustflags from .cargo/config.toml, losing the /LIBPATH
# entries to the vintage MSVC import libs (the link then fails with "could not
# open 'kernel32.lib'"). The per-target variable below replaces them too, so
# the flags are re-derived from the config with only the subsystem swapped.
# /SUBSYSTEM:WINDOWS makes the CRT enter through WinMain (see src/main.rs);
# the console TUI builds in build95.sh keep /SUBSYSTEM:CONSOLE.
TARGET_ENV_VAR="CARGO_TARGET_$(echo "$TARGET" | tr 'a-z-' 'A-Z_')_RUSTFLAGS"
export "$TARGET_ENV_VAR"="$(python3 - "$TARGET" <<'PYFLAGS'
import re, sys
target = sys.argv[1]
cfg = open(".cargo/config.toml").read()
key = "[target.'cfg(all(target_family = \"rust9x\", target_env = \"msvc\"))']"
block = cfg.split(key)[1]
array = block.split("rustflags = [")[1].split("]")[0]
flags = re.findall(r"'(.*?)'", array)
# Drop the console subsystem and add the GUI one (lld/rustc honour the last).
flags = [f for f in flags if not f.startswith("-Clink-arg=/SUBSYSTEM")]
flags.append("-Clink-arg=/SUBSYSTEM:WINDOWS,4.0")
# GUI entry point: WinMainCRTStartup -> WinMain (src/main.rs gui-subsystem).
flags.append("-Clink-arg=/ENTRY:WinMainCRTStartup")
# GUI CRT variant FIRST: a full copy of VC98Lib whose LIBCMT.LIB has the
# console startup object (crt0.obj) removed, so only the GUI one
# (wincrt0.obj) can satisfy /ENTRY:WinMainCRTStartup. Listing the original
# VC98Lib after it is harmless (all other libs resolve there), but it must
# not come first or the console crt0.obj is picked up and collides.
flags = ["-Clink-arg=/LIBPATH:/opt/msvc-toolchains/vc6-gui"] + flags
sys.stdout.write(" ".join(flags))
PYFLAGS
)"

RUSTC_BOOTSTRAP=1 cargo +rust9x build --profile "$PROFILE" --target "$TARGET" \
  -Zbuild-std=std,panic_abort \
  --no-default-features --features gui,gui-subsystem --bin corro \
  "${ARGS[@]+"${ARGS[@]}"}"

SRC="target/$TARGET/$PROFILE/corro.exe"
[ -f "$SRC" ] || { echo "ERROR: build output missing: $SRC" >&2; exit 1; }
mkdir -p dist
# `gcorro.exe`: argv[0] dispatch selects the GUI backend.
python3 - "$SRC" dist/corro-win95-gui.exe <<'EOF'
import struct, sys
# Win95 loader: zero DllCharacteristics (rust9x emits 0x8140). NOTE the
# layout: optional-header offset 68 = Subsystem (MUST keep — here it is
# 2 = GUI, which is what we want for this build), 70 = DllCharacteristics.
d = bytearray(open(sys.argv[1], "rb").read())
pe = struct.unpack("<I", d[0x3c:0x40])[0]
subsystem = struct.unpack_from("<H", d, pe + 24 + 68)[0]
if subsystem != 2:
    print(f"WARNING: expected GUI subsystem (2), got {subsystem}", file=sys.stderr)
struct.pack_into("<H", d, pe + 24 + 70, 0)
open(sys.argv[2], "wb").write(d)
EOF

# Copy under the gcorro name too: argv[0] dispatch picks the GUI backend.
cp dist/corro-win95-gui.exe dist/gcorro.exe

echo "-- built dist/corro-win95-gui.exe ($(stat -c%s dist/corro-win95-gui.exe) bytes)"
md5sum dist/corro-win95-gui.exe
