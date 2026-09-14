#!/usr/bin/env bash
# Build the Windows NWG GUI binary (gcorro.exe) via MinGW cross-compile.
#
#   scripts/build-gcorro.sh [dest] [--release]
#
# Builds `corro` for x86_64-pc-windows-gnu with the GTK feature set (the
# NWG backend is selected on Windows via cfg) and stages it as gcorro.exe.
# Default dest is /mnt/c/Temp/gcorro.exe (C:\Temp\gcorro.exe), the path the
# PowerShell probe scripts use. Pass another path to stage elsewhere.
#
# Debug (default) keeps debug_asserts on — they caught a real startup crash
# (BoxWidget::set_child_vexpand pre-append call). Use --release for a
# quieter, optimized binary.
#
# If the destination is locked by a running gcorro.exe, the script tries
# once to stop it (best effort via powershell.exe interop) and retries the
# copy before giving up with a clear message.
set -euo pipefail
cd "$(dirname "$0")/.."

PROFILE=debug
ARGS=()
DEST=""
for arg in "$@"; do
    case "$arg" in
        --release) PROFILE=release ;;
        -h|--help)
            echo "Usage: $0 [dest] [--release]"
            echo "  dest defaults to /mnt/c/Temp/gcorro.exe"
            exit 0
            ;;
        *) DEST="$arg" ;;
    esac
done
if [ -z "$DEST" ]; then
    DEST="/mnt/c/Temp/gcorro.exe"
fi

if ! rustup target list --installed 2>/dev/null | grep -q "x86_64-pc-windows-gnu"; then
    echo "error: rust target x86_64-pc-windows-gnu is not installed" >&2
    echo "  run: rustup target add x86_64-pc-windows-gnu" >&2
    exit 1
fi

if [ "$PROFILE" = "release" ]; then
    ARGS+=(--release)
fi
cargo build --target x86_64-pc-windows-gnu --features gtk "${ARGS[@]}"

SRC="target/x86_64-pc-windows-gnu/$PROFILE/corro.exe"
if ! cp "$SRC" "$DEST" 2>/dev/null; then
    echo "destination locked ($DEST); trying to stop a running gcorro.exe..."
    if command -v powershell.exe >/dev/null 2>&1; then
        powershell.exe -NoProfile -Command \
            "Get-Process gcorro -ErrorAction SilentlyContinue | Stop-Process -Force" >/dev/null 2>&1 || true
        sleep 2
    fi
    if ! cp "$SRC" "$DEST" 2>/dev/null; then
        echo "error: cannot write $DEST — is gcorro.exe still running?" >&2
        exit 1
    fi
fi
echo "staged $DEST"
ls -lh "$DEST"
