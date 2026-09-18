#!/usr/bin/env bash
# Regenerate the link-time stub DLLs for the Windows/ARM64 (gnullvm) build.
#
# Background: rustc's aarch64-pc-windows-gnullvm target links system DLLs
# DIRECTLY (no import libs) and searches for the DLL files themselves before
# invoking the linker. There is no aarch64 MinGW toolchain in Ubuntu's repos
# and no Windows/ARM64 box here, so we fabricate link-time-only stubs: real
# PE32+ Aarch64 DLLs (built with zig) exporting the full symbol set lifted
# from the host x86_64 MinGW import libs (symbol names are arch-independent).
# At runtime on a real device the loader binds these names to the genuine
# system DLLs; the stubs are never shipped (MSIX carries only gcorro.exe).
#
# Only three DLLs are needed -- everything else resolves via zig's bundled
# import libs: msvcrt + synchronization (full exports from MinGW import
# libs), windows.0.52.0 (empty; the windows crate's import lib satisfies the
# link, rustc just demands the file exists).
#
#   tools/make-armstubs.sh [destdir]   (default: /tmp/armstubs)
# Then build with:
#   export CARGO_TARGET_AARCH64_PC_WINDOWS_GNULLVM_LINKER=$PWD/tools/zig-cc-aarch64-windows
#   RUSTFLAGS="-C link-arg=-mwindows -L <destdir>" \
#       cargo build --release --target aarch64-pc-windows-gnullvm --features gtk
set -euo pipefail
DEST="${1:-/tmp/armstubs}"
MINGW_LIB=/usr/x86_64-w64-mingw32/lib
mkdir -p "$DEST"
cd "$DEST"

gen_c() { # gen_c <importlib.a> <stem>  -> <stem>_stub.c
    local lib="$1" stem="$2"
    x86_64-w64-mingw32-nm --defined-only --format=posix "$lib" | python3 -c "
import sys
code, data = set(), set()
for line in sys.stdin:
    p = line.split()
    if len(p) < 2: continue
    name, typ = p[0], p[1]
    if not name.replace('_','').isalnum(): continue
    if name.startswith('__imp_'):
        if typ in ('D','I','R','B','C'): data.add(name[6:])
    elif typ == 'T': code.add(name)
    elif typ in ('D','B','R','C'): data.add(name)
data -= code
with open('$stem' + '_stub.c', 'w') as f:
    for s in sorted(code): f.write(f'__declspec(dllexport) void {s}(void){{}}\n')
    for s in sorted(data): f.write(f'__declspec(dllexport) int {s};\n')
print('$stem', 'funcs:', len(code), 'vars:', len(data))
"
}

gen_c "$MINGW_LIB/libmsvcrt.a" msvcrt
gen_c "$MINGW_LIB/libsynchronization.a" synchronization
# windows.0.52.0 is virtual (the windows crate's import lib satisfies the
# link; rustc just demands the file exists). An empty C file links reliably;
# a .def-only input makes zig fall back to the system linker and fail.
echo 'void win_armstub_anchor(void){}' > windows_stub.c

# Drop symbols zig's own CRT already defines (else duplicate-symbol at link).
for stem in msvcrt synchronization; do
    for _ in $(seq 1 30); do
        errs=$(zig cc -target aarch64-windows-gnu -shared -fno-builtin \
            -o "$stem.dll" "${stem}_stub.c" 2>&1 || true)
        dups=$(echo "$errs" | grep -o 'duplicate symbol: [^ ]*' | cut -d' ' -f3 | sort -u || true)
        [ -z "$dups" ] && break
        # shellcheck disable=SC2086
        grep -vwE "$(echo $dups | tr ' ' '|')" "${stem}_stub.c" > "${stem}_stub.c.tmp"
        mv "${stem}_stub.c.tmp" "${stem}_stub.c"
    done
    zig cc -target aarch64-windows-gnu -shared -fno-builtin \
        -o "$stem.dll" "${stem}_stub.c"
    echo "stub $stem.dll OK"
done
zig cc -target aarch64-windows-gnu -shared -fno-builtin \
    -o windows.0.52.0.dll windows_stub.c
echo "stub windows.0.52.0.dll OK"
ls -la "$DEST"/*.dll
