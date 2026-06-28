"""
GUI Replayer — Windows-only test runner for the NWG backend.

Launches corro.exe with a .corro file, simulates keyboard input via
PostMessage (bypasses UIPI which blocks SendInput from background
processes), captures stdout/stderr, and reports pass/fail.

Usage:
    python .gui_replayer.py --test recrec5
    python .gui_replayer.py --test recrec6
    python .gui_replayer.py --list
"""

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time

try:
    import ctypes
    from ctypes import wintypes
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# Windows API constants
WM_KEYDOWN = 0x0100
WM_KEYUP = 0x0101
WM_SYSKEYDOWN = 0x0104
WM_SYSKEYUP = 0x0105
KF_ALTDOWN = 0x2000


def _post(hwnd, msg, wparam, lparam=0):
    """Post a message to the given window."""
    if not HAS_WIN32 or hwnd is None:
        return False
    ok = ctypes.windll.user32.PostMessageW(hwnd, msg, wparam, lparam)
    if not ok:
        print(f"  WARN: PostMessageW({hwnd}, 0x{msg:04x}, 0x{wparam:04x}) failed (returned {ok})")
    return ok


def key_down(hwnd, vk_code, syskey=False):
    """Post a key-down message to the window."""
    msg = WM_SYSKEYDOWN if syskey else WM_KEYDOWN
    _post(hwnd, msg, vk_code, 0)


def key_up(hwnd, vk_code, syskey=False):
    """Post a key-up message to the window."""
    msg = WM_SYSKEYUP if syskey else WM_KEYUP
    _post(hwnd, msg, vk_code, 0)


def send_key(hwnd, vk_code, syskey=False):
    """Post a single key press (down+up)."""
    key_down(hwnd, vk_code, syskey)
    time.sleep(0.05)
    key_up(hwnd, vk_code, syskey)
    time.sleep(0.05)


def send_text(hwnd, text):
    """Send a string of text via keystrokes (lowercase only, no Shift)."""
    for ch in text:
        vk = ord(ch.upper())
        send_key(hwnd, vk)


def send_enter(hwnd):
    send_key(hwnd, 0x0D)  # VK_RETURN


def send_escape(hwnd):
    send_key(hwnd, 0x1B)  # VK_ESCAPE


def send_tab(hwnd):
    send_key(hwnd, 0x09)  # VK_TAB


WM_CLOSE = 0x0010


def send_close(hwnd):
    """Post WM_CLOSE to the window to trigger the registered close handler."""
    _post(hwnd, WM_CLOSE, 0, 0)


def send_alt_f(hwnd):
    """Send Alt+F key sequence to set the NWG seq_alt_f flag.

    This activates the Rust-level Alt+F detector so that a subsequent
    Q keystroke (sent via the normal key path to the focused child)
    triggers save_before_quit.  We do NOT send Q here — the replayer
    sends all keystrokes including Q through the normal input path so
    that the Rust handle_key function processes them in order.

    Alt down uses WM_SYSKEYDOWN to set the Alt modifier.  F uses regular
    WM_KEYDOWN so TranslateMessage generates WM_CHAR('f') rather than
    WM_SYSCHAR('f'), avoiding the NWG menu accelerator.
    """
    key_down(hwnd, 0x12, syskey=True)   # VK_MENU, WM_SYSKEYDOWN
    time.sleep(0.1)
    key_down(hwnd, 0x46, syskey=False)  # VK_F, WM_KEYDOWN
    time.sleep(0.05)
    key_up(hwnd, 0x46, syskey=False)    # VK_F up
    time.sleep(0.05)
    key_up(hwnd, 0x12, syskey=True)     # VK_MENU up, WM_SYSKEYUP


def send_down(hwnd):
    send_key(hwnd, 0x28)  # VK_DOWN


def send_right(hwnd):
    send_key(hwnd, 0x27)  # VK_RIGHT


def find_corro_window():
    """Find the HWND of the corro window by enumerating windows."""
    if not HAS_WIN32:
        return None
    found_hwnd = [None]
    def enum_callback(hwnd, _lparam):
        buf = ctypes.create_unicode_buffer(256)
        length = ctypes.windll.user32.GetWindowTextW(hwnd, buf, 256)
        if length > 0 and "corro" in buf.value.lower():
            found_hwnd[0] = hwnd
            return False
        return True
    callback = ctypes.CFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    ctypes.windll.user32.EnumWindows(callback(enum_callback), 0)
    return found_hwnd[0]


def _launch_and_wait(binary, test_file):
    """Launch corro, wait for window, and return (proc, hwnd)."""
    proc = subprocess.Popen(
        [binary, "--gui", test_file],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    # Retry finding the window: initial 1.5s, then up to 5 more at 1s intervals
    hwnd = None
    for attempt in range(6):
        hwnd = find_corro_window()
        if hwnd:
            break
        if attempt == 0:
            time.sleep(1.5)
        else:
            time.sleep(1.0)
    if hwnd:
        # Bring to foreground for reliable focus (though PostMessage
        # bypasses foreground requirements).
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        ctypes.windll.user32.BringWindowToTop(hwnd)
        time.sleep(0.5)
    else:
        print("WARNING: corro window not found")
    return proc, hwnd


def _cleanup(proc, label=""):
    """Try to gracefully terminate corro, then force kill if needed.
    Returns True if the process exited cleanly (returncode == 0)."""
    elapsed = 0.0
    import time as _time
    t0 = _time.time()
    try:
        stdout, stderr = proc.communicate(timeout=12)
        elapsed = _time.time() - t0
        ok = proc.returncode == 0
        if not ok:
            print(f"  [{label}] exit code={proc.returncode} elapsed={elapsed:.1f}s")
        # Print captured output for diagnostics
        if stdout:
            sys.stdout.write(stdout.decode("utf-8", errors="replace"))
        if stderr:
            sys.stderr.write(stderr.decode("utf-8", errors="replace"))
        return ok
    except subprocess.TimeoutExpired:
        elapsed = _time.time() - t0
        print(f"  [{label}] communicate timed out after {elapsed:.1f}s, force-killing")
        proc.kill()
        try:
            stdout, stderr = proc.communicate(timeout=3)
            if stdout:
                sys.stdout.write(stdout.decode("utf-8", errors="replace"))
            if stderr:
                sys.stderr.write(stderr.decode("utf-8", errors="replace"))
        except subprocess.TimeoutExpired:
            print(f"  [{label}] force-kill also timed out")
            proc.kill()
        return False


def run_test_recrec5(binary="target/release/corro.exe"):
    """Test: open subtotal-tiny, enter data, quit."""
    print(f"Running recrec5 test with {binary}")
    test_file = "docs/tests/subtotal-tiny.corro"
    output_file = "docs/tests/subtotal-tiny.corro"  # corro writes back to same file
    if not os.path.exists(test_file):
        print(f"ERROR: test file not found: {test_file}")
        return False

    # Record initial output size (should be small, just test data)
    initial_size = os.path.getsize(output_file) if os.path.exists(output_file) else 0

    proc, hwnd = _launch_and_wait(binary, test_file)
    if not hwnd:
        print("WARNING: could not find corro window, keystrokes may go elsewhere")
    else:
        print(f"Found corro window (hwnd={hwnd})")

    # Type "42" into A1 and press Enter
    time.sleep(0.3)
    send_text(hwnd, "42")
    time.sleep(0.2)
    send_enter(hwnd)
    time.sleep(0.3)

    # Navigate to B1 and type "hello"
    time.sleep(0.2)
    send_text(hwnd, "hello")
    time.sleep(0.2)
    send_enter(hwnd)
    time.sleep(0.3)

    # Quit via WM_CLOSE.  The NWG backend's raw WM_CLOSE handler
    # calls quit_main_loop() which posts WM_QUIT, causing the message
    # loop to exit cleanly.  This avoids the race condition with
    # Alt+F+Q where PostMessageW(WM_KEYUP/Q) fails because the app
    # already started quitting after processing WM_KEYDOWN/Q.
    time.sleep(0.3)
    send_close(hwnd)
    time.sleep(0.5)

    result = _cleanup(proc, "recrec5")
    final_size = os.path.getsize(output_file) if os.path.exists(output_file) else 0
    output_exists = "yes" if os.path.exists(output_file) and final_size > initial_size else "no"
    print(f"  output_file={output_exists} initial={initial_size}b final={final_size}b")
    _restore_test_file()
    print("recrec5 done")
    return result


def run_test_recrec6(binary="target/release/corro.exe"):
    """Test: open subtotal-tiny, navigate cells, quit."""
    print(f"Running recrec6 test with {binary}")
    test_file = "docs/tests/subtotal-tiny.corro"
    if not os.path.exists(test_file):
        print(f"ERROR: test file not found: {test_file}")
        return False

    proc, hwnd = _launch_and_wait(binary, test_file)
    if not hwnd:
        print("WARNING: could not find corro window, keystrokes may go elsewhere")
    else:
        print(f"Found corro window (hwnd={hwnd})")

    # Navigate with arrow keys and Tab
    send_down(hwnd)
    time.sleep(0.2)
    send_right(hwnd)
    time.sleep(0.2)
    send_tab(hwnd)
    time.sleep(0.2)

    # Type data
    time.sleep(0.3)
    send_text(hwnd, "test")
    time.sleep(0.2)
    send_enter(hwnd)
    time.sleep(0.3)

    # Quit via WM_CLOSE (see recrec5 for rationale).
    time.sleep(0.3)
    send_close(hwnd)
    time.sleep(0.5)

    result = _cleanup(proc, "recrec6")
    _restore_test_file()
    print("recrec6 done")
    return result


def _restore_test_file():
    """Restore docs/tests/subtotal-tiny.corro to canonical 21-line state
    and ensure it is writable.

    The corro binary modifies this file in-place when launched with --gui,
    appending SET commands.  We restore it before each test run to avoid
    accumulating extra SET commands that would cause golden-file mismatches.

    The canonical file is the first 22 lines.
    We preserve them and discard any replayer artifacts that have been
    appended by previous runs.
    """
    canonical_lines = 22
    path = "docs/tests/subtotal-tiny.corro"
    if not os.path.exists(path):
        print(f"WARN: _restore_test_file: {path} not found, cannot restore")
        return
    try:
        # First clear the read-only attribute so we can write.
        # git checks out files read-only on Windows.
        if sys.platform == "win32":
            import subprocess as _sp
            # Use PowerShell for reliable attribute clearing (attrib.exe can be flaky)
            _sp.run(["powershell", "-Command",
                     f"Set-ItemProperty -LiteralPath '{path}' -Name IsReadOnly -Value $false"],
                    capture_output=True, text=True)
        import os as _os, stat as _stat
        _os.chmod(path, _stat.S_IWRITE | _stat.S_IREAD)
    except Exception as e:
        print(f"WARN: _restore_test_file clear read-only failed: {e}")
    try:
        with open(path, "r") as f:
            lines = f.readlines()
        if len(lines) != canonical_lines:
            with open(path, "w") as f:
                f.writelines(lines[:canonical_lines])
    except Exception as e:
        print(f"WARN: _restore_test_file error restoring {path}: {e}")


def main():
    _restore_test_file()
    parser = argparse.ArgumentParser(description="GUI Replayer for NWG tests")
    parser.add_argument("--test", choices=["recrec5", "recrec6", "all"], default="all")
    parser.add_argument("--binary", default="target/release/corro.exe")
    parser.add_argument("--list", action="store_true", help="List available tests")
    args = parser.parse_args()

    if args.list:
        print("Available tests:")
        print("  recrec5  - Open subtotal-tiny, enter data, quit")
        print("  recrec6  - Open subtotal-tiny, navigate cells, quit")
        return 0

    if sys.platform != "win32":
        print("ERROR: .gui_replayer.py is Windows-only (uses Win32 SendInput)")
        return 1

    if not HAS_WIN32:
        print("ERROR: ctypes not available (needed for Win32 API)")
        return 1

    binary = args.binary or os.environ.get("BIN", "")
    if not os.path.exists(binary):
        print(f"ERROR: binary not found at {binary}")
        print("  Try: cargo +nightly build --features gui")
        print("  Or set the BIN environment variable to the corro.exe path")
        return 1

    tests = []
    if args.test == "all":
        tests = [("recrec5", run_test_recrec5), ("recrec6", run_test_recrec6)]
    else:
        tests = [(args.test, {"recrec5": run_test_recrec5, "recrec6": run_test_recrec6}[args.test])]

    failed = 0
    for name, fn in tests:
        print(f"\n--- {name} ---")
        if fn(binary):
            print(f"  PASS: {name}")
        else:
            print(f"  FAIL: {name}")
            failed += 1

    return failed


if __name__ == "__main__":
    sys.exit(main())
