#!/usr/bin/env python3
"""Record two corro windows editing one file, side by side.

corro's workbook is an append-only log, so two processes on the same file stay
in sync with no save step: each commit is appended, and the other window tails
the file (GUI: a 250ms poll; TUI: once per loop iteration).

This script runs two instances side by side under a throwaway X server and
screenshots the screen as they work, so the recording shows each value appearing
in *both* windows.

The edits are made by the applications themselves, not by synthetic keystrokes:
each process is given `CORRO_EDIT_SCRIPT` (e.g. `1000:A5=111,5000:A7=333`) and
applies those edits through the ordinary commit path — the same call a typed
value makes. Driving it through `xdotool` typing would mean calibrating pixel
coordinates against the window manager, which is brittle and silently writes to
the wrong cell when it drifts.

Usage:
    scripts/two_window_movie.py -o dist/corro-two-windows.mp4          # GUI + GUI
    scripts/two_window_movie.py --pair gui-tui -o dist/corro-gui-tui.mp4

Requires: Xvfb, a window manager (fluxbox/openbox/...), ffmpeg.
"""

from __future__ import annotations

import argparse
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent

WIN_W, WIN_H = 1200, 800
SCREEN_W, SCREEN_H = 2400, 820

# The display the demo runs on. xwd/convert must target it explicitly: the
# parent's DISPLAY is the operator's, not the throwaway server.
_DISPLAY = ":92"

# Edits each window performs (ms after session start : cell = value). They
# interleave so the recording alternates which window is writing.
# Edits alternate between the windows. `--tempo` scales these: 1.0 is the
# values below, 2.0 runs them at half speed (used for the shipped demo).
EDIT_TEMPO = 1.0
# Values long enough to read as *typed* at the capture rate. The recorder runs
# at 8 fps, so a three-character value is gone between two frames and the
# recording looks like the value simply appeared; these take long enough that
# the per-character animation is visible in the video.
_BASE_LEFT = "2000:A5=1111,10000:A7=33333,18000:A9=555555"
_BASE_RIGHT = "6000:A6=2222,14000:A8=44444,22000:A10=666666"


def scale_edits(script: str, tempo: float) -> str:
    """Multiply every timestamp in `MS:ADDR=VALUE,...` by `tempo`."""
    if tempo == 1.0:
        return script
    out = []
    for item in script.split(","):
        when, _, rest = item.partition(":")
        out.append(f"{int(int(when) * tempo)}:{rest}")
    return ",".join(out)


def last_edit_ms(scripts: list[str]) -> int:
    times = [int(i.split(":")[0]) for s in scripts for i in s.split(",")]
    return max(times, default=0)


def require(tool: str) -> str:
    path = shutil.which(tool)
    if path is None:
        sys.exit(f"error: {tool} not found on PATH")
    return path


def _x_env() -> dict:
    env = dict(os.environ)
    env["DISPLAY"] = _DISPLAY
    return env


def _geometry(wid: str) -> tuple[int, int, int, int] | None:
    geo = subprocess.run(
        ["xdotool", "getwindowgeometry", "--shell", wid],
        env=_x_env(), capture_output=True, text=True,
    ).stdout
    vals: dict[str, int] = {}
    for line in geo.splitlines():
        if "=" in line:
            k, _, v = line.partition("=")
            try:
                vals[k] = int(v)
            except ValueError:
                pass
    if "WIDTH" not in vals:
        return None
    return vals.get("X", 0), vals.get("Y", 0), vals["WIDTH"], vals["HEIGHT"]


def find_window(pid: int, deadline: float) -> str | None:
    """The corro toplevel owned by `pid`.

    GTK creates 10x10 helper windows alongside the toplevel, so a pid search
    returns several ids; the real window is the one with screen-sized geometry.
    """
    while time.monotonic() < deadline:
        for wid in subprocess.run(
            ["xdotool", "search", "--pid", str(pid)], env=_x_env(),
            capture_output=True, text=True,
        ).stdout.split():
            g = _geometry(wid)
            if g and g[2] > 200 and g[3] > 200:
                return wid
        time.sleep(0.3)
    return None


def place(wid: str, x: int, y: int) -> None:
    subprocess.run(["xdotool", "windowmove", "--sync", wid, str(x), str(y)],
                   env=_x_env(), capture_output=True)
    subprocess.run(["xdotool", "windowsize", "--sync", wid, str(WIN_W), str(WIN_H)],
                   env=_x_env(), capture_output=True)


class Recorder:
    """Records the whole display while a session runs.

    Uses ffmpeg's `x11grab`, which captures continuously at a real frame rate.
    The obvious approach — loop `xwd` + `convert` — is far too slow to be
    usable: at 2400x820 each of those takes ~500ms, so the loop delivers about
    **1 fps** however high the requested rate is. That silently truncated the
    recording: a 32-second session produced ~32 frames, which at the output
    frame rate played back as a 2.7-second clip.
    """

    def __init__(self, out_dir: Path, fps: float, width: int, height: int):
        self.pattern = out_dir / "frame-%05d.ppm"
        self.proc = subprocess.Popen(
            ["ffmpeg", "-y", "-loglevel", "error",
             "-f", "x11grab",
             "-framerate", str(fps),
             "-video_size", f"{width}x{height}",
             "-i", _DISPLAY,
             "-pix_fmt", "rgb24",
             str(self.pattern)],
            env=_x_env(),
            stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
        )

    def wait(self, seconds: float) -> None:
        time.sleep(seconds)

    def finish(self) -> None:
        # SIGINT lets ffmpeg finalize the last frame; terminate can truncate it.
        self.proc.send_signal(signal.SIGINT)
        try:
            self.proc.wait(timeout=30)
        except subprocess.TimeoutExpired:
            self.proc.kill()


def build_file(path: Path) -> None:
    """A small starting workbook, so both windows show content on open."""
    path.write_text(
        "CORRO_LOG 1\n"
        "SET $1:[A1 --- shared workbook ---\n"
        "SET $1:A1 10\n"
        "SET $1:A2 20\n"
        "SET $1:A3 30\n"
        "SET $1:B1 =A1*2\n"
        "SET $1:B2 =A2*2\n"
        "SET $1:B3 =A3*2\n"
    )


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("-o", "--out", type=Path, default=REPO_ROOT / "dist" / "corro-two-windows.mp4")
    ap.add_argument("--pair", choices=["gui-gui", "gui-tui"], default="gui-gui",
                    help="which two front-ends to show (default: gui-gui)")
    ap.add_argument("--fps", type=int, default=12, help="output frame rate")
    ap.add_argument("--capture-fps", type=float, default=8.0, help="screenshot rate")
    # Characters per second while a window types into a cell. The capture runs
    # at 8 fps, so much above ~8 cps the animation is invisible in the video.
    ap.add_argument("--typing-cps", type=float, default=6.0,
                    help="characters per second when a window types")
    ap.add_argument("--crf", type=int, default=20, help="x264 quality")
    ap.add_argument("--keep-frames", action="store_true")
    ap.add_argument("--frames-dir", type=Path, help="write frames here (implies --keep-frames)")
    ap.add_argument("--tempo", type=float, default=EDIT_TEMPO,
                    help="scales the edit timings (2.0 = half speed, default: %(default)s)")
    args = ap.parse_args()

    for tool in ("Xvfb", "xdotool", "ffmpeg"):
        require(tool)
    binary = REPO_ROOT / "target" / "debug" / "corro"
    if not binary.exists():
        sys.exit(f"error: {binary} not found — run: cargo build --features gui")
    if args.pair == "gui-tui" and not shutil.which("xterm"):
        sys.exit("error: xterm is needed for --pair gui-tui")

    work = Path(tempfile.mkdtemp(prefix="corro-twowin-"))
    if args.frames_dir:
        frames = args.frames_dir
        if frames.exists():
            shutil.rmtree(frames)
        frames.mkdir(parents=True)
        args.keep_frames = True
    else:
        frames = work / "frames"
        frames.mkdir()
    shared = work / "shared.corro"
    build_file(shared)

    # A stale server on this display would make our Xvfb fail to bind while the
    # old one keeps serving, so the windows we start are not the ones we think
    # we are looking at and `find_window` never finds them. Clear it first and
    # then wait for the new server to answer before starting anything.
    subprocess.run(["pkill", "-f", f"Xvfb {_DISPLAY}"], capture_output=True)
    time.sleep(0.5)
    xvfb = subprocess.Popen(
        ["Xvfb", _DISPLAY, "-screen", "0", f"{SCREEN_W}x{SCREEN_H}x24"],
        stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
    )
    deadline = time.monotonic() + 15
    while time.monotonic() < deadline:
        probe = subprocess.run(
            ["xdpyinfo", "-display", _DISPLAY],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        if probe.returncode == 0:
            break
        if xvfb.poll() is not None:
            sys.exit(f"error: Xvfb {_DISPLAY} exited immediately")
        time.sleep(0.25)
    else:
        sys.exit(f"error: Xvfb {_DISPLAY} never became ready")
    # A window manager is required, not cosmetic: without one, focus and
    # stacking between two toplevels are undefined and the second window can
    # end up hidden behind the first.
    wm = None
    for candidate in ("fluxbox", "openbox", "matchbox-window-manager", "twm"):
        if shutil.which(candidate):
            wm = subprocess.Popen([candidate], env=_x_env(),
                                  stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            break
    if wm is None:
        print("[two_window] warning: no window manager found; the windows may overlap")

    procs: list[subprocess.Popen] = []
    rec = Recorder(frames, args.capture_fps, SCREEN_W, SCREEN_H)
    try:
        time.sleep(2.0)

        left_script = scale_edits(_BASE_LEFT, args.tempo)
        right_script = scale_edits(_BASE_RIGHT, args.tempo)

        left_env = _x_env()
        left_env["CORRO_EDIT_SCRIPT"] = left_script
        left_env["CORRO_MOVIE_TYPING_CPS"] = str(args.typing_cps)
        right_env = _x_env()
        right_env["CORRO_EDIT_SCRIPT"] = right_script
        right_env["CORRO_MOVIE_TYPING_CPS"] = str(args.typing_cps)

        left = subprocess.Popen([str(binary), "--gui", str(shared)], env=left_env,
                                stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
        procs.append(left)
        if args.pair == "gui-gui":
            right = subprocess.Popen([str(binary), "--gui", str(shared)], env=right_env,
                                     stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
        else:
            right = subprocess.Popen(
                ["xterm", "-geometry", "150x46", "-title", "corro (terminal)",
                 "-e", str(binary), "--ratatui", str(shared)],
                env=right_env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
        procs.append(right)

        lw = find_window(left.pid, time.monotonic() + 25)
        rw = find_window(right.pid, time.monotonic() + 25)
        if not lw or not rw:
            sys.exit(f"error: expected two windows (left={lw}, right={rw})")
        place(lw, 0, 0)
        place(rw, WIN_W, 0)
        time.sleep(1.5)

        # Long enough for the last edit to land and be seen in both windows.
        hold = last_edit_ms([left_script, right_script]) / 1000.0 + 4.0
        rec.wait(hold)
        rec.finish()

        captured = sorted(frames.glob("frame-*.ppm"))
        print(f"[two_window] captured {len(captured)} frames")

        out = args.out
        out.parent.mkdir(parents=True, exist_ok=True)
        subprocess.run(
            ["ffmpeg", "-y", "-loglevel", "error",
             "-framerate", str(args.fps),
             "-i", str(frames / "frame-%05d.ppm"),
             "-vf", "scale=trunc(iw/2)*2:trunc(ih/2)*2",
             "-c:v", "libx264", "-pix_fmt", "yuv420p", "-crf", str(args.crf),
             "-movflags", "+faststart", str(out)],
            check=True,
        )
        print(f"[two_window] wrote {out} ({out.stat().st_size / 1024:.0f} KiB)")

        log = shared.read_text()
        print("--- shared log after the session ---")
        print(log)
        # The demo only means something if both windows actually wrote.
        wrote = [l for l in log.splitlines() if any(v in l for v in ("111", "222", "333", "444", "555", "666"))]
        print(f"edits committed by the two windows: {len(wrote)}")
        if len(wrote) < 6:
            print("warning: expected 6 scripted edits; the recording will be incomplete")
    finally:
        for p in procs:
            p.send_signal(signal.SIGTERM)
        for p in procs:
            try:
                p.wait(timeout=5)
            except subprocess.TimeoutExpired:
                p.kill()
        if wm:
            wm.terminate()
        xvfb.terminate()
        if args.keep_frames and args.frames_dir is None:
            print(f"[two_window] frames kept in {frames}")
        elif not args.keep_frames:
            shutil.rmtree(work, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
