#!/usr/bin/env bash
# GUI backend test runner.
# Tests GTK3/GTK4 (via WSL) and NWG (Windows native) backends.
#
# Usage:
#   ./gui_loop.sh                  # Run all available backend tests
#   ./gui_loop.sh --nwg-only      # Run only NWG (Windows native) tests
#   ./gui_loop.sh --gtk-only      # Run only GTK tests (requires WSL)
#   ./gui_loop.sh --list          # List available test scenarios

set -eo pipefail
cd "$(dirname "$0")"

PASS=0
FAIL=0
SKIP=0

pass() { PASS=$((PASS+1)); echo "  PASS: $1"; }
fail() { FAIL=$((FAIL+1)); echo "  FAIL: $1"; }
skip() { SKIP=$((SKIP+1)); echo "  SKIP: $1"; }

# ---------------------------------------------------------------------------
# Prerequisite checks
# ---------------------------------------------------------------------------

HAS_WSL=false
HAS_WSL_BASH=false
HAS_PYTHON=false
PYTHON=""
HAS_CARGO=false
HAS_WSL_CARGO=false

if command -v wsl.exe &>/dev/null; then
    HAS_WSL=true
    # Wrap wsl.exe to prevent MSYS2 path conversion (which would convert
    # WSL paths like /mnt/d/... to D:\..., breaking the command).
    wsl() { MSYS2_ARG_CONV_EXCL="*" wsl.exe "$@"; }
    if wsl bash --version &>/dev/null; then
        HAS_WSL_BASH=true
    fi
fi

if command -v python3 &>/dev/null; then
    HAS_PYTHON=true
    PYTHON="python3"
elif command -v python &>/dev/null; then
    HAS_PYTHON=true
    PYTHON="python"
fi

if command -v cargo &>/dev/null || which cargo &>/dev/null 2>&1; then
    HAS_CARGO=true
fi

if [ "$HAS_WSL" = true ] && [ "$HAS_WSL_BASH" = true ]; then
    if wsl bash -c "command -v cargo &>/dev/null" 2>/dev/null; then
        HAS_WSL_CARGO=true
    fi
fi

echo "=== GUI Backend Test Suite ==="
echo "WSL: $HAS_WSL  Python: $HAS_PYTHON  Cargo: $HAS_CARGO"
echo ""

# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

if [ "$HAS_CARGO" = true ]; then
    echo "--- Building GUI backend ---"
    if cargo +nightly build --features gui 2>&1; then
        pass "Build (gui features)"
    else
        fail "Build (gui features)"
    fi
    echo ""

    # Restore test data files from git before running tests.
    # The NWG replayer can modify subtotal-tiny.corro (it launches
    # corro.exe --gui on that file, which appends SET commands).
    # Also ensure test_rec5.corro is byte-identical to subtotal-tiny.corro
    # (the committed version may be out of sync).
    # The committed HEAD may itself have accumulated replayer artifacts,
    # so we truncate to the canonical 21 lines after checkout.
    git checkout -- docs/tests/subtotal-tiny.corro test_rec5.corro 2>/dev/null || true
    # The committed file has blank lines between SET commands (42 lines total).
    # head -n 42 preserves all 21 SET/FILL commands + 21 blank lines.
    # Using fewer lines (e.g. head -n 21) would truncate to only 11 SET
    # commands because of the interleaved blank lines.
    for f in docs/tests/subtotal-tiny.corro test_rec5.corro; do
        if [ -f "$f" ]; then
            head -n 42 "$f" > /tmp/$(basename "$f").clean 2>/dev/null
            cp -f /tmp/$(basename "$f").clean "$f" 2>/dev/null || true
            chmod +w "$f" 2>/dev/null || true
            rm -f /tmp/$(basename "$f").clean 2>/dev/null || true
        fi
    done
    # On Windows, chmod +w may not clear the read-only attribute.
    # Use attrib -r as a fallback (available via cmd.exe in Git Bash).
    if [[ $OSTYPE == "msys" || $OSTYPE == "cygwin" || -n "${WINDIR:-}" ]]; then
        for f in docs/tests/subtotal-tiny.corro test_rec5.corro; do
            if [ -f "$f" ]; then
                # Git Bash converts /d/GitHub/... to D:\GitHub\... for native Windows exes
                cmd.exe /c "attrib -R $f" 2>/dev/null || true
            fi
        done
    fi
    # Ensure test_rec5.corro is writable and has consistent line endings.
    # The git checkout may leave it read-only with CRLF line endings, which
    # can cause "output mismatch" in the check_vals test.
    if [ -f test_rec5.corro ]; then
        chmod +w test_rec5.corro 2>/dev/null || true
    fi

    # ---------------------------------------------------------------------------
    # Rust integration tests (NWG on Windows, gtk on Linux)
    # ---------------------------------------------------------------------------

    echo "--- Running recording replay tests (test_tiny5 / test_tiny6) ---"
    # NOTE: No --features gui here — these tests only use ui::App (TUI) which
    # doesn't need the gui feature. The gui feature links native-windows-gui
    # which causes the test process to hang after completion on Windows.
    if cargo +nightly test --test test_tiny5 2>&1; then
        pass "recrec5 (test_tiny5)"
    else
        fail "recrec5 (test_tiny5)"
    fi

    if cargo +nightly test --test test_tiny6 2>&1; then
        pass "recrec6 (test_tiny6)"
    else
        fail "recrec6 (test_tiny6)"
    fi
    echo ""

    # ---------------------------------------------------------------------------
    # GUI-specific Rust tests
    # ---------------------------------------------------------------------------

    echo "--- Running GUI-specific Rust tests ---"
    for t in check_vals gui_enter_text_creates_file check_agg_gui check_gui_imports quit_alt_f_q; do
        if cargo +nightly test --features gui --test "$t" 2>&1; then
            pass "$t"
        else
            fail "$t"
        fi
    done
    echo ""
else
    skip "Rust build and tests (cargo not available)"
    echo ""
fi

# ---------------------------------------------------------------------------
# NWG replayer tests (Windows only)
# ---------------------------------------------------------------------------

if [[ $OSTYPE == "msys" || $OSTYPE == "cygwin" || -n "${WINDIR:-}" ]]; then
    echo "--- NWG replayer tests ---"
    if [ "$HAS_PYTHON" = true ] && [ -f .gui_replayer.py ]; then
        # The debug binary was already built in the Build step above.
        # Pass BIN env var so the replayer uses the freshly-built debug binary.
        # Use release binary (debug binary has eprintln! calls that fill the
        # stderr pipe buffer, causing the child process to block mid-commit).
        if BIN="target/release/corro.exe" $PYTHON .gui_replayer.py --test recrec5 2>&1; then
            pass "NWG replayer: recrec5"
        else
            fail "NWG replayer: recrec5"
        fi
        if BIN="target/release/corro.exe" $PYTHON .gui_replayer.py --test recrec6 2>&1; then
            pass "NWG replayer: recrec6"
        else
            fail "NWG replayer: recrec6"
        fi
    else
        skip "NWG replayer tests (no python or .gui_replayer.py)"
    fi
else
    skip "NWG replayer tests (not Windows)"
fi
echo ""

# ---------------------------------------------------------------------------
# GTK tests via WSL (Linux/WSL only)
# ---------------------------------------------------------------------------

if [ "$HAS_WSL" = true ] && [ "$HAS_WSL_BASH" = true ] && [ "$HAS_WSL_CARGO" = true ]; then
    echo "--- GTK3/GTK4 tests via WSL ---"

    # Build with GTK feature (single build, GTK_DLOPEN_PREFER_GTK3 runtime toggle)
    # DISPLAY=:0 is set inside WSL so that external test wrappers that add
    # 'DISPLAY=:0' after 'timeout' (wrong bash syntax) don't cause errors.
    # Inside the WSL bash -c string, DISPLAY=:0 is a regular env var prefix.
    WSL_PWD="/mnt/$(pwd | sed 's|^/\([a-z]\)/|\1/|')"
    if wsl bash -c "cd '$WSL_PWD' && DISPLAY=:0 cargo +nightly build --features gui" 2>&1; then
        pass "GTK build"
    else
        fail "GTK build"
    fi

    # Run Rust tests
    if wsl bash -c "cd '$WSL_PWD' && DISPLAY=:0 cargo +nightly test --test test_tiny5" 2>&1; then
        pass "GTK recrec5"
    else
        fail "GTK recrec5"
    fi
    if wsl bash -c "cd '$WSL_PWD' && DISPLAY=:0 cargo +nightly test --test test_tiny6" 2>&1; then
        pass "GTK recrec6"
    else
        fail "GTK recrec6"
    fi
elif [ "$HAS_WSL" = true ] && [ "$HAS_WSL_BASH" = true ]; then
    skip "GTK tests (cargo not installed in WSL)"
else
    skip "GTK tests (WSL/bash not available)"
fi

# ---------------------------------------------------------------------------
# WASM build (cross-compile check)
# ---------------------------------------------------------------------------

echo "--- WASM build ---"
if [ "$HAS_CARGO" = true ]; then
    if rustup target list --toolchain nightly 2>/dev/null | grep -q "wasm32-unknown-unknown (installed)"; then
        if cargo +nightly build --target wasm32-unknown-unknown --features wasm --no-default-features 2>&1; then
            pass "WASM build"
        else
            fail "WASM build"
        fi
    else
        skip "WASM build (wasm32-unknown-unknown target not installed)"
    fi
else
    skip "WASM build (cargo not available)"
fi
echo ""

echo ""
echo "=== Results: $PASS passed, $FAIL failed, $SKIP skipped ==="
exit $FAIL
