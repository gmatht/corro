#!/usr/bin/env bash
# WSL wrapper for running cargo/test commands inside WSL
# with DISPLAY=:0 set correctly (avoids 'timeout DISPLAY=:0' syntax error).
#
# Usage:
#   ./scripts/wsl-wrapper.sh cargo +nightly build --features gui
#   ./scripts/wsl-wrapper.sh cargo +nightly test --features gui --test test_tiny5
#
# This sets DISPLAY=:0 inside the WSL bash -c string so external test
# runners that wrap with `timeout DISPLAY=:0` (wrong syntax) don't fail.

set -eo pipefail

WSL_PWD="/mnt/$(pwd | sed 's|^/\([a-z]\)/|\1/|')"
exec wsl.exe bash -c "cd '$WSL_PWD' && DISPLAY=:0 $*"
