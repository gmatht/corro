#!/usr/bin/env python3
"""Small QMP helper for driving the Win95 QEMU VM.

Usage (as a CLI):
    qmp.py shot OUT.png                 screendump (saved as PNG)
    qmp.py key <qcode> [qcode ...]      tap keys one after another
    qmp.py combo <qcode>+<qcode>        press keys simultaneously (e.g. alt+f4)
    qmp.py type "<text>"                type text (UK layout: '\\' = less, ':' = shift+semicolon)
    qmp.py run-dialog "<command>"       open Start > Run and launch <command>
    qmp.py wait-stable [--min SECS] [--timeout SECS]
                                        poll screendumps until the screen has
                                        stopped changing (boot detection)
    qmp.py wait-change BASE.png [--timeout SECS]
                                        poll until the screen differs from BASE.png

QMP port can be overridden with the QMP_PORT env var (default 4445).
Every function can also be imported.
"""
import json
import os
import socket
import sys
import time

QMP_PORT = int(os.environ.get("QMP_PORT", "4445"))


def _qmp_iter():
    """Context-manager-free generator yielding (cmd -> response) plumbing."""
    raise NotImplementedError


class Qmp:
    def __init__(self, port=QMP_PORT):
        self.sock = socket.create_connection(("127.0.0.1", port), timeout=30)
        self.f = self.sock.makefile("rwb")
        self.f.readline()  # greeting
        self.cmd("qmp_capabilities")

    def cmd(self, name, args=None, quiet=True):
        req = {"execute": name}
        if args:
            req["arguments"] = args
        self.f.write(json.dumps(req).encode() + b"\n")
        self.f.flush()
        resp = json.loads(self.f.readline().decode())
        if not quiet and "error" in resp:
            print(f"QMP {name}: {resp['error']}", file=sys.stderr)
        return resp

    def close(self):
        try:
            self.sock.close()
        except OSError:
            pass


def screendump_png(q, path, tmp=None):
    """Screendump to `tmp` (ppm) and convert/save as PNG at `path`."""
    tmp = tmp or (path.rsplit(".", 1)[0] + ".ppm")
    q.cmd("screendump", {"filename": tmp, "format": "ppm"})
    time.sleep(0.5)
    from PIL import Image

    Image.open(tmp).save(path)


def _small(im, size=(320, 240)):
    return im.convert("L").resize(size)


def _diff(a, b):
    pa, pb = a.load(), b.load()
    total = 0
    for y in range(a.size[1]):
        for x in range(a.size[0]):
            total += abs(pa[x, y] - pb[x, y])
    return total / (255.0 * a.size[0] * a.size[1])


def wait_stable(q, min_elapsed=90, timeout=420, poll=5, verbose=False):
    """Wait until the screen stops changing (VM finished booting)."""
    from PIL import Image

    tmp = "/tmp/_qmp_wait.ppm"
    prev = None
    stable = 0
    start = time.time()
    while time.time() - start < timeout:
        time.sleep(poll)
        try:
            q.cmd("screendump", {"filename": tmp, "format": "ppm"}, quiet=True)
            time.sleep(0.5)
            cur = _small(Image.open(tmp))
        except Exception:
            continue
        if prev is not None:
            d = _diff(prev, cur)
            if verbose:
                print(f"  t={time.time()-start:5.0f}s diff={d:.4f}", flush=True)
            stable = stable + 1 if d < 0.002 else 0
        prev = cur
        if stable >= 2 and time.time() - start >= min_elapsed:
            return True
    return False


def _has_taskbar(im):
    """Win95 desktop detection via the two always-present taskbar widgets:
    the Start button (bottom-left) and the clock (bottom-right). Each is a
    gray box with dark text, so on a real desktop both boxes are mostly
    light-gray neutral pixels with a dash of near-black text pixels.
    Wallpapers fail (their bottom-left/right corners are dark or colored:
    the stock wallpaper's reflection measures light<0.3), black boot stages
    fail (light=0), and the clouds splash fails (colored)."""
    im = im.convert("RGB")
    w, h = im.size
    px = im.load()

    def box_stats(x0, x1, y0, y1):
        n = light = dark = 0
        for y in range(y0, y1):
            for x in range(x0, x1):
                r, g, b = px[x, y]
                n += 1
                if max(r, g, b) - min(r, g, b) <= 14:
                    v = (r + g + b) / 3
                    if 140 <= v <= 228:
                        light += 1
                    elif v < 60:
                        dark += 1
        return light / n, dark / n

    sl, sd = box_stats(2, 66, h - 24, h - 2)      # Start button
    cl, cd = box_stats(w - 96, w - 6, h - 24, h - 2)  # clock
    return sl > 0.40 and 0.02 < sd < 0.45 and cl > 0.60 and cd < 0.30


def wait_desktop(q, min_elapsed=140, timeout=540, poll=5, verbose=False,
                 esc_nudge=True):
    """Wait until the Windows 95 desktop (taskbar) is visible.

    After `min_elapsed`, a blocked boot is assumed and Esc is sent every
    ~30s to dismiss boot-time dialogs (e.g. the "Display Properties" error
    that Safe Mode boots show, which otherwise keeps the taskbar hidden
    forever)."""
    from PIL import Image

    tmp = "/tmp/_qmp_wait.ppm"
    start = time.time()
    hits = 0
    last_nudge = 0.0
    while time.time() - start < timeout:
        time.sleep(poll)
        now = time.time()
        try:
            q.cmd("screendump", {"filename": tmp, "format": "ppm"}, quiet=True)
            time.sleep(0.5)
            im = Image.open(tmp)
        except Exception:
            continue
        t = _has_taskbar(im)
        if verbose:
            print(f"  t={now-start:5.0f}s taskbar={t}", flush=True)
        if t:
            hits += 1
            if hits >= 2 and now - start >= min_elapsed:
                return True
        else:
            hits = 0
            if esc_nudge and now - start > min_elapsed and now - last_nudge > 25:
                q.cmd("send-key", {"keys": [{"type": "qcode", "data": "esc"}]})
                last_nudge = now
    return False


def wait_halt(q, timeout=150, poll=5):
    """Wait until the screen goes dark (Win95 'It is now safe to turn off
    your computer' screen) — i.e. the guest has shut down cleanly."""
    from PIL import Image

    tmp = "/tmp/_qmp_halt.ppm"
    start = time.time()
    while time.time() - start < timeout:
        time.sleep(poll)
        try:
            q.cmd("screendump", {"filename": tmp, "format": "ppm"}, quiet=True)
            time.sleep(0.5)
            im = Image.open(tmp).convert("L")
        except Exception:
            continue
        hist = im.histogram()
        total = sum(hist)
        dark = sum(hist[:32]) / total
        if dark > 0.98:
            return True
    return False


def wait_change(q, base_path, timeout=90, poll=3, threshold=0.008):
    """Wait until the screen differs from the PNG at base_path."""
    from PIL import Image

    base = _small(Image.open(base_path))
    tmp = "/tmp/_qmp_wait.ppm"
    start = time.time()
    while time.time() - start < timeout:
        time.sleep(poll)
        try:
            q.cmd("screendump", {"filename": tmp, "format": "ppm"}, quiet=True)
            time.sleep(0.5)
            cur = _small(Image.open(tmp))
        except Exception:
            continue
        if _diff(base, cur) > threshold:
            return True
    return False


# --- keyboard -------------------------------------------------------------

# US/UK keyboard: chars that need mapping to qcodes (letters/digits are 1:1).
CHAR_MAP = {
    " ": ["spc"],
    "\t": ["tab"],
    "\n": ["ret"],
    ".": ["dot"],
    ",": ["comma"],
    "-": ["minus"],
    "=": ["equal"],
    "/": ["slash"],
    ";": ["semicolon"],
    "'": ["apostrophe"],
    ":": ["shift", "semicolon"],
    "\\": ["less"],  # this layout maps backslash to the ISO key ('less' qcode)
    "": ["spc"],
}


def type_text(q, text, delay=0.2):
    for ch in text:
        keys = CHAR_MAP.get(ch)
        if keys is None:
            if ch.isupper():
                keys = ["shift", "shift-" + ch.lower()]
            else:
                keys = [ch]
        q.cmd("send-key", {"keys": [{"type": "qcode", "data": k} for k in keys]})
        time.sleep(delay)


def run_dialog(q, command, delay=0.2):
    """Open Start > Run, clear the edit field, type `command`, press Enter."""
    q.cmd("send-key", {"keys": [
        {"type": "qcode", "data": "ctrl"}, {"type": "qcode", "data": "esc"}]})
    time.sleep(2.5)
    q.cmd("send-key", {"keys": [{"type": "qcode", "data": "r"}]})
    time.sleep(3.5)
    for _ in range(40):
        q.cmd("send-key", {"keys": [{"type": "qcode", "data": "backspace"}]})
        time.sleep(0.05)
    type_text(q, command, delay)
    time.sleep(1.0)
    q.cmd("send-key", {"keys": [{"type": "qcode", "data": "ret"}]})


# --- CLI ------------------------------------------------------------------

def main(argv):
    if not argv:
        print(__doc__)
        return 2
    cmd, rest = argv[0], argv[1:]
    q = Qmp()
    try:
        if cmd == "shot":
            screendump_png(q, rest[0])
            print(rest[0])
        elif cmd == "key":
            for k in rest:
                q.cmd("send-key", {"keys": [{"type": "qcode", "data": k}]})
                time.sleep(0.5)
        elif cmd == "combo":
            keys = rest[0].split("+")
            q.cmd("send-key", {"keys": [{"type": "qcode", "data": k} for k in keys]})
        elif cmd == "type":
            type_text(q, rest[0])
        elif cmd == "run-dialog":
            run_dialog(q, rest[0])
        elif cmd == "wait-stable" or cmd == "wait-desktop":
            opts = {"min_elapsed": 90, "timeout": 420, "verbose": False}
            if cmd == "wait-desktop":
                opts = {"min_elapsed": 140, "timeout": 540, "verbose": False}
            fn = wait_stable if cmd == "wait-stable" else wait_desktop
            args = list(rest)
            while args:
                a = args.pop(0)
                if a == "--min":
                    opts["min_elapsed"] = float(args.pop(0))
                elif a == "--timeout":
                    opts["timeout"] = float(args.pop(0))
                elif a == "--verbose":
                    opts["verbose"] = True
            ok = fn(q, **opts)
            print("stable" if ok else "TIMEOUT")
            return 0 if ok else 1
        elif cmd == "wait-change":
            path = rest[0]
            timeout, threshold = 90, 0.008
            args = list(rest[1:])
            while args:
                a = args.pop(0)
                if a == "--timeout":
                    timeout = float(args.pop(0))
                elif a == "--threshold":
                    threshold = float(args.pop(0))
            ok = wait_change(q, path, timeout, threshold=threshold)
            print("changed" if ok else "TIMEOUT")
            return 0 if ok else 1
        elif cmd == "wait-halt":
            ok = wait_halt(q)
            print("halted" if ok else "TIMEOUT")
            return 0 if ok else 1
        else:
            print(f"unknown command: {cmd}", file=sys.stderr)
            return 2
    finally:
        q.close()
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
