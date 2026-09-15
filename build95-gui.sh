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
    --release) PROFILE=release ;;
    *) ARGS+=("$a") ;;
  esac
done

# GUI CRT member fix: this LIBCMT.LIB contains FOUR startup objects defining
# the same symbols (___app_type, ___error_mode, __aexit_rtn, __aenvptr,
# __wenvptr, __amsg_exit): dllcrt0.obj (DLL entry), 0352 (wmain entry),
# wincrt0.obj (GUI entry), 0360 (wwinmain entry). Proven by lld /VERBOSE:
# some input references __amsg_exit and the archive index lists a non-GUI
# startup member first, so lld loads it; /ENTRY then loads wincrt0.obj too
# — duplicate-symbol link failure. Nothing in our inputs references any
# DLL/wmain/wwinmain-unique symbol (exhaustively scanned), so drop the
# three non-GUI startup objects from a build-time copy (the machine-global
# toolchain is untouched and target/ is git-ignored). Deletion is by
# *defined entry symbol*, not member number (robust to lib rebuilds).
GUI_LIBDIR="$(pwd)/target/$TARGET/gui-crt"
mkdir -p "$GUI_LIBDIR"
if [ ! -f "$GUI_LIBDIR/LIBCMT.LIB" ]; then
  cp /opt/msvc-toolchains/vc6-gui/LIBCMT.LIB "$GUI_LIBDIR/LIBCMT.LIB"
  for sym in __DllMainCRTStartup@12 _wmainCRTStartup _wWinMainCRTStartup; do
    MEMBER="$(llvm-nm --print-armap "$GUI_LIBDIR/LIBCMT.LIB" | grep -E "^${sym} in .*\.obj" | sed -E 's/.* in //')"
    [ -n "$MEMBER" ] || { echo "ERROR: startup member for $sym not found" >&2; exit 1; }
    llvm-ar d "$GUI_LIBDIR/LIBCMT.LIB" "$MEMBER" || exit 1
  done
  # the GUI entry object must survive, as the SOLE ___app_type owner
  # (no `grep -q` here: with pipefail it SIGPIPEs llvm-nm on match and
  #  falsely fails precisely when the entry IS present)
  llvm-nm --print-armap "$GUI_LIBDIR/LIBCMT.LIB" | grep "_WinMainCRTStartup in " > /dev/null || { echo "ERROR: wincrt0 lost from LIBCMT.LIB" >&2; exit 1; }
  OWNERS="$(llvm-nm --print-armap "$GUI_LIBDIR/LIBCMT.LIB" | grep -cE "^___app_type in ")"
  [ "$OWNERS" = 1 ] || { echo "ERROR: $OWNERS ___app_type owners remain" >&2; exit 1; }
  # rust-lld looks up "libcmt.lib" verbatim on this case-sensitive
  # filesystem (cf. the lowercase symlinks in sdk71a/Lib); without the
  # lowercase alias it silently skips this directory and links the
  # original (broken) LIBCMT.LIB instead.
  ln -sf LIBCMT.LIB "$GUI_LIBDIR/libcmt.lib"
fi

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
export "$TARGET_ENV_VAR"="$(python3 - "$TARGET" "$GUI_LIBDIR" <<'PYFLAGS'
import re, sys, os
target = sys.argv[1]
gui_libdir = os.path.abspath(sys.argv[2])
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
# GUI CRT variant FIRST: the build-time copy above (dllcrt0.obj removed so
# only wincrt0.obj can provide the GUI startup symbols), then the original
# VC98Lib after it (all other libs resolve there). It must not come first
# or the DLL startup object is picked up and collides.
flags = ["-Clink-arg=/LIBPATH:/opt/msvc-toolchains/vc6-gui"] + flags
flags = ["-Clink-arg=/LIBPATH:" + gui_libdir] + flags
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
# layout: optional-header offset 68 = Subsystem, 70 = DllCharacteristics.
# Subsystem is FORCED to 2 (GUI) here: the per-target RUSTFLAGS above are
# *concatenated* with .cargo/config.toml's (not a replace), so the config's
# /SUBSYSTEM:CONSOLE,4.0 rides along last and wins the link; the explicit
# /ENTRY:WinMainCRTStartup is unaffected (entry resolution ignores the
# subsystem), so stamping the header field post-link yields a correct GUI
# exe (no console window). scripts/package-win95-release.sh re-verifies it.
d = bytearray(open(sys.argv[1], "rb").read())
pe = struct.unpack("<I", d[0x3c:0x40])[0]
subsystem = struct.unpack_from("<H", d, pe + 24 + 68)[0]
if subsystem != 2:
    print(f"note: link produced subsystem {subsystem}; stamping GUI (2)", file=sys.stderr)
struct.pack_into("<H", d, pe + 24 + 68, 2)
struct.pack_into("<H", d, pe + 24 + 70, 0)
open(sys.argv[2], "wb").write(d)
EOF

# Copy under the gcorro name too: argv[0] dispatch picks the GUI backend.
cp dist/corro-win95-gui.exe dist/gcorro.exe

echo "-- built dist/corro-win95-gui.exe ($(stat -c%s dist/corro-win95-gui.exe) bytes)"
md5sum dist/corro-win95-gui.exe
