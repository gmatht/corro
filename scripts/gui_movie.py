#!/usr/bin/env python3
"""Record a GUI `--movie` replay to a video.

`corro --gui --movie --movie-frames DIR FILE.corro` replays the workbook
line-by-line through the GUI renderer and writes one lossless PPM per frame.
This script drives that run and encodes the frames with ffmpeg, so the whole
demo video is reproducible from the repository with no display server, no
screen-capture tooling and no manual editing.

Usage:
    scripts/gui_movie.py docs/tests/subtotal.corro -o dist/corro-gui-movie.mp4
    scripts/gui_movie.py FILE.corro --keep-frames --cps 18

Options mirror the CLI's `--movie-*` pacing flags, so the video and an
interactive replay of the same file look the same.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
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


def run_movie(binary: Path, corro_file: Path, frames_dir: Path, args: argparse.Namespace) -> None:
    """Replay the movie, capturing one frame per painted step."""
    env = dict(os.environ)
    env["CORRO_MOVIE_FRAMES"] = str(frames_dir)
    # A movie render must never depend on a display: the capture path paints
    # into a raster surface directly.
    env.pop("DISPLAY", None)
    cmd = [
        str(binary),
        "--gui",
        "--movie",
        "--movie-frames",
        str(frames_dir),
        "--movie-typing-cps",
        str(args.cps),
        "--movie-confirm-ms",
        str(args.confirm_ms),
        "--movie-menu-hold-ms",
        str(args.menu_hold_ms),
        str(corro_file),
    ]
    print(f"[gui_movie] {' '.join(cmd)}")
    proc = subprocess.run(cmd, env=env, capture_output=True, text=True)
    if proc.stdout.strip():
        print(proc.stdout.strip())
    if proc.returncode != 0:
        sys.exit(f"error: corro exited {proc.returncode}\n{proc.stderr.strip()}")
    for line in proc.stderr.splitlines():
        if line.startswith("[corro]"):
            print(line)


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
