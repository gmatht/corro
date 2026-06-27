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
    if wsl.exe bash --version &>/dev/null; then
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

if command -v cargo &>/dev/null; then
    HAS_CARGO=true
fi

if [ "$HAS_WSL" = true ] && [ "$HAS_WSL_BASH" = true ]; then
    if wsl.exe bash -c "command -v cargo &>/dev/null" 2>/dev/null; then
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
    if cargo +nightly build --features gui 2>/dev/null; then
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
    git checkout -- docs/tests/subtotal-tiny.corro test_rec5.corro 2>/dev/null || true
    chmod +w test_rec5.corro 2>/dev/null || true
    cp -f docs/tests/subtotal-tiny.corro test_rec5.corro 2>/dev/null || true

    # ---------------------------------------------------------------------------
    # Rust integration tests (NWG on Windows, gtk on Linux)
    # ---------------------------------------------------------------------------

    echo "--- Running recording replay tests (test_tiny5 / test_tiny6) ---"
    if cargo +nightly test --features gui --test test_tiny5 2>/dev/null; then
        pass "recrec5 (test_tiny5)"
    else
        fail "recrec5 (test_tiny5)"
    fi

    if cargo +nightly test --features gui --test test_tiny6 2>/dev/null; then
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
        if cargo +nightly test --features gui --test "$t" 2>/dev/null; then
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
        if $PYTHON .gui_replayer.py --test recrec5 2>/dev/null; then
            pass "NWG replayer: recrec5"
        else
            fail "NWG replayer: recrec5"
        fi
        if $PYTHON .gui_replayer.py --test recrec6 2>/dev/null; then
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
    WSL_PWD="/mnt/$(pwd | sed 's|/|/|g')"
    if wsl.exe bash -c "cd '$WSL_PWD' && cargo +nightly build --features gui" 2>/dev/null; then
        pass "GTK build"
    else
        fail "GTK build"
    fi

    # Run Rust tests
    if wsl.exe bash -c "cd '$WSL_PWD' && cargo +nightly test --features gui --test test_tiny5" 2>/dev/null; then
        pass "GTK recrec5"
    else
        fail "GTK recrec5"
    fi
    if wsl.exe bash -c "cd '$WSL_PWD' && cargo +nightly test --features gui --test test_tiny6" 2>/dev/null; then
        pass "GTK recrec6"
    else
        fail "GTK recrec6"
    fi
elif [ "$HAS_WSL" = true ] && [ "$HAS_WSL_BASH" = true ]; then
    skip "GTK tests (cargo not installed in WSL)"
else
    skip "GTK tests (WSL/bash not available)"
fi

echo ""
echo "=== Results: $PASS passed, $FAIL failed, $SKIP skipped ==="
exit $FAIL
