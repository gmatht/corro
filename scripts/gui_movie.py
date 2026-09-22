#!/usr/bin/env python3
"""Record a GUI `--movie` replay to a video.

`corro --gui --movie FILE.corro` replays the workbook through the *live
window*: the same widget tree and draw callbacks as an interactive session,
with a timer applying one step at a time. This script runs that under a
throwaway X server, screenshots the window as it plays, and encodes the
screenshots with ffmpeg.

Recording the real window (rather than rendering frames through a parallel
code path) is deliberate: the demo then shows exactly what the app shows, and
cannot drift from it.

Usage:
    scripts/gui_movie.py docs/tests/subtotal.corro -o dist/corro-gui-movie.mp4
    scripts/gui_movie.py FILE.corro --video-only --cps 18

Requires: ffmpeg (with x11grab) and an X server binary (Xvfb).

To see which cells a recording will show as formula errors (#NAME, #PARSE, ...)
and which log line produced each one, use `scripts/movie_errors.py`.
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


def find_binary(explicit: str | None, release: bool) -> Path:
    if explicit:
        p = Path(explicit)
        if not p.exists():
            sys.exit(f"error: binary not found: {p}")
        return p
    profile = "release" if release else "debug"
    candidate = REPO_ROOT / "target" / profile / "corro"
    if candidate.exists():
        return candidate
    sys.exit(
        f"error: {candidate} not found — build it first:\n"
        f"  cargo build{' --release' if release else ''} --features gui"
    )


def require(tool: str) -> str:
    path = shutil.which(tool)
    if path is None:
        sys.exit(f"error: {tool} not found on PATH (needed to record the window)")
    return path


def run_movie(binary: Path, corro_file: Path, frames_dir: Path, args: argparse.Namespace) -> None:
    """Play the movie in a real window under Xvfb, screenshotting as it runs."""
    xvfb = require("Xvfb")
    require("ffmpeg")

    display = args.display
    env = dict(os.environ)
    env["DISPLAY"] = display

    # Clear a stale server on this display first: if one is still bound, our
    # Xvfb fails to start while the old one keeps serving, and the recording
    # then captures whatever that server holds.
    subprocess.run(["pkill", "-f", f"Xvfb {display}"], capture_output=True)
    time.sleep(0.5)
    xvfb_proc = subprocess.Popen(
        [xvfb, display, "-screen", "0", f"{args.width}x{args.height}x24"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
    )
    try:
        # Wait for the server to answer rather than guessing a sleep: a slow
        # machine can take longer than the old fixed 1.5s, and recording before
        # it is up yields empty or partial frames.
        deadline = time.monotonic() + 15
        while time.monotonic() < deadline:
            probe = subprocess.run(
                ["xdpyinfo", "-display", display],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            if probe.returncode == 0:
                break
            if xvfb_proc.poll() is not None:
                sys.exit(f"error: Xvfb {display} exited immediately")
            time.sleep(0.25)
        else:
            sys.exit(f"error: Xvfb {display} never became ready")
        cmd = [
            str(binary),
            "--gui",
            "--movie",
            "--movie-typing-cps",
            str(args.cps),
            "--movie-confirm-ms",
            str(args.confirm_ms),
            "--movie-menu-hold-ms",
            str(args.menu_hold_ms),
            str(corro_file),
        ]
        print(f"[gui_movie] {' '.join(cmd)}")
        # ffmpeg's x11grab records continuously at a real frame rate. The
        # obvious `xwd` + `convert` loop cannot: at this resolution each takes
        # ~500ms, so it delivers ~1 fps no matter what rate is requested, and
        # the recording silently comes out as a fraction of the session length.
        recorder = subprocess.Popen(
            ["ffmpeg", "-y", "-loglevel", "error",
             "-f", "x11grab",
             "-framerate", str(args.capture_fps),
             "-video_size", f"{args.width}x{args.height}",
             "-i", display,
             "-pix_fmt", "rgb24",
             "-frames:v", str(args.max_frames),
             str(frames_dir / "frame-%05d.ppm")],
            env=env, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
        )
        try:
            corro = subprocess.Popen(
                cmd, env=env, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True
            )
            corro.wait(timeout=args.timeout)
            # A short tail so the last replayed frame is on screen when the
            # recorder stops.
            time.sleep(0.5)
        finally:
            recorder.send_signal(signal.SIGINT)
            try:
                recorder.wait(timeout=30)
            except subprocess.TimeoutExpired:
                recorder.kill()
        n = len(list(frames_dir.glob("frame-*.ppm")))
        print(f"[gui_movie] captured {n} frames")
    finally:
        xvfb_proc.terminate()
        try:
            xvfb_proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            xvfb_proc.kill()


def encode(frames_dir: Path, out: Path, fps: int, crf: int, scale: str | None) -> None:
    """Encode the PPM sequence into an H.264 MP4 (yuv420p so it plays everywhere)."""
    out.parent.mkdir(parents=True, exist_ok=True)
    pattern = str(frames_dir / "frame-%05d.ppm")
    cmd = [
        "ffmpeg",
        "-y",
        "-loglevel",
        "error",
        "-framerate",
        str(fps),
        "-i",
        pattern,
        # H.264 wants even dimensions; the capture is 1200x800 already, but a
        # custom --size could be odd.
        "-vf",
        scale or "scale=trunc(iw/2)*2:trunc(ih/2)*2",
        "-c:v",
        "libx264",
        "-pix_fmt",
        "yuv420p",
        "-crf",
        str(crf),
        "-movflags",
        "+faststart",
        str(out),
    ]
    print(f"[gui_movie] encoding {len(list(frames_dir.glob('frame-*.ppm')))} frames -> {out}")
    subprocess.run(cmd, check=True)


def main() -> int:
    ap = argparse.ArgumentParser(description="Record a corro GUI --movie replay to a video")
    ap.add_argument("corro_file", type=Path, help="the .corro log to replay")
    ap.add_argument("-o", "--out", type=Path, default=Path("dist/corro-gui-movie.mp4"))
    ap.add_argument("--binary", help="path to the corro binary (default: target/{debug,release}/corro)")
    ap.add_argument("--release", action="store_true", help="prefer the release binary")
    ap.add_argument("--cps", type=float, default=22.0, help="typing speed, chars/sec (--movie-typing-cps)")
    ap.add_argument("--confirm-ms", type=int, default=140, help="per-step hold (--movie-confirm-ms)")
    ap.add_argument("--menu-hold-ms", type=int, default=900, help="menu flash hold (--movie-menu-hold-ms)")
    ap.add_argument("--fps", type=int, default=12, help="output frame rate")
    ap.add_argument("--crf", type=int, default=20, help="x264 quality (lower = better)")
    ap.add_argument("--scale", help="ffmpeg scale filter (default: even dimensions)")
    ap.add_argument("--keep-frames", action="store_true", help="keep the intermediate frames")
    ap.add_argument("--frames-dir", type=Path, help="directory for the frames (implies --keep-frames)")
    ap.add_argument("--display", default=":78", help="X display for the throwaway server")
    ap.add_argument("--width", type=int, default=1200, help="window width to record")
    ap.add_argument("--height", type=int, default=800, help="window height to record")
    ap.add_argument("--capture-fps", type=float, default=8.0, help="screenshot rate while recording")
    ap.add_argument("--max-frames", type=int, default=6000, help="safety cap on captured frames")
    ap.add_argument("--timeout", type=float, default=900, help="max seconds to let the movie play")
    args = ap.parse_args()

    if not args.corro_file.exists():
        sys.exit(f"error: input not found: {args.corro_file}")
    if shutil.which("ffmpeg") is None:
        sys.exit("error: ffmpeg not found on PATH")

    binary = find_binary(args.binary, args.release)

    keep = args.keep_frames or args.frames_dir is not None
    tmp: tempfile.TemporaryDirectory | None = None
    if args.frames_dir:
        frames_dir = args.frames_dir
        if frames_dir.exists():
            shutil.rmtree(frames_dir)
        frames_dir.mkdir(parents=True)
    elif keep:
        frames_dir = Path("dist") / f"{args.corro_file.stem}-frames"
        frames_dir.mkdir(parents=True, exist_ok=True)
    else:
        tmp = tempfile.TemporaryDirectory(prefix="corro-gui-movie-")
        frames_dir = Path(tmp.name)

    try:
        run_movie(binary, args.corro_file, frames_dir, args)
        frames = sorted(frames_dir.glob("frame-*.ppm"))
        if not frames:
            sys.exit("error: no frames were rendered (does the file contain any ops?)")
        encode(frames_dir, args.out, args.fps, args.crf, args.scale)
        size = args.out.stat().st_size
        print(f"[gui_movie] wrote {args.out} ({size / 1024:.0f} KiB, {len(frames)} frames @ {args.fps} fps)")
    finally:
        if tmp is not None:
            tmp.cleanup()
        elif keep:
            print(f"[gui_movie] frames kept in {frames_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
