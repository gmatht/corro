#!/usr/bin/env python3
"""Record the corro GUI demo video.

Builds one video from a title card, a replay of each featured workbook, and a
closing card. The replays come from `scripts/gui_movie.py`, which plays the
movie in the real corro window (under a throwaway X server) and screenshots it —
so the video shows exactly what the application shows.

Usage:
    python3 scripts/demo_movie.py                 # write dist/corro-gui-movie.mp4
    python3 scripts/demo_movie.py --cps 7         # slower typing
    python3 scripts/demo_movie.py -o /tmp/x.mp4

Requires the same tools as `gui_movie.py` (Xvfb, xwd, ImageMagick, ffmpeg),
Pillow and a built GUI binary.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_FONT_DIRS = [
    Path("/usr/share/fonts/truetype/dejavu"),
    Path("/usr/share/fonts/TTF"),
    Path("/Library/Fonts"),
]

# Cards are rendered at the recording resolution so every frame in the video
# matches; the replays come back at the same size.
WIDTH, HEIGHT = 1200, 800

# Features replayed in the video: (workbook, card title, card description).
FEATURES = [
    (
        "docs/tests/subtotal.corro",
        "Subtotals and summary labels",
        "One marker in the left margin creates a subtotal or grand total column. "
        "corro never double counts: the grand total sums the subtotals, not the raw data.",
    ),
    (
        "docs/tests/main.corro",
        "Sheets, moves and formats",
        "A workbook is a log: sheets are created, copied and activated, rows are moved, "
        "and column formats are applied - each replayed as the interaction that produced it.",
    ),
]


def font(name: str, size: int) -> ImageFont.FreeTypeFont:
    for d in DEFAULT_FONT_DIRS:
        p = d / name
        if p.exists():
            return ImageFont.truetype(str(p), size)
    # Fall back to Pillow's bundled bitmap font rather than failing the run.
    return ImageFont.load_default()


def card(title: str, desc: str, out: Path) -> None:
    """Render a title/description card."""
    im = Image.new("RGB", (WIDTH, HEIGHT), (18, 22, 30))
    d = ImageDraw.Draw(im)
    d.text((60, 210), "corro", font=font("DejaVuSansMono-Bold.ttf", 96), fill=(120, 190, 255))
    d.text((60, 340), title, font=font("DejaVuSansMono-Bold.ttf", 40), fill=(255, 255, 255))
    body = font("DejaVuSansMono.ttf", 22)
    lines, cur = [], ""
    for word in desc.split():
        if len(cur) + len(word) + 1 > 72:
            lines.append(cur)
            cur = word
        else:
            cur = (cur + " " + word).strip()
    if cur:
        lines.append(cur)
    y = 410
    for ln in lines:
        d.text((60, y), ln, font=body, fill=(190, 200, 215))
        y += 32
    d.text(
        (60, 715),
        "corro --gui --movie  ·  recorded from the running window",
        font=body,
        fill=(110, 130, 160),
    )
    im.save(out)


def main() -> int:
    ap = argparse.ArgumentParser(description="Record the corro GUI demo video")
    ap.add_argument("-o", "--out", type=Path, default=ROOT / "dist" / "corro-gui-movie.mp4")
    ap.add_argument(
        "--cps",
        type=float,
        default=6.5,
        help="typing speed in chars/sec for the replayed workbooks (default: 6.5)",
    )
    ap.add_argument("--confirm-ms", type=int, default=400, help="per-step hold while replaying")
    ap.add_argument("--menu-hold-ms", type=int, default=1400, help="menu flash hold while replaying")
    ap.add_argument("--fps", type=int, default=12, help="output frame rate")
    ap.add_argument("--crf", type=int, default=20, help="x264 quality (lower is better)")
    ap.add_argument("--title-secs", type=float, default=3.0, help="how long each card holds")
    ap.add_argument("--capture-fps", type=float, default=8.0, help="screenshot rate while replaying")
    ap.add_argument("--keep-frames", action="store_true", help="keep the intermediate frames")
    args = ap.parse_args()

    for tool in ("ffmpeg", "Xvfb", "xwd", "convert"):
        if shutil.which(tool) is None:
            sys.exit(f"error: {tool} not found on PATH")

    work = Path(tempfile.mkdtemp(prefix="corro-demo-"))
    frames = work / "frames"
    frames.mkdir()
    n = 0

    def emit(im: Image.Image) -> None:
        nonlocal n
        im.save(frames / f"frame-{n:05d}.ppm")
        n += 1

    def hold(im: Image.Image, secs: float) -> None:
        # A card is a still; repeat it so the encoded video shows it for `secs`.
        for _ in range(max(1, int(args.fps * secs))):
            emit(im)

    try:
        c = work / "title.png"
        card(
            "GUI movie mode",
            "The --movie demo mode plays back a .corro workbook in the corro window itself, "
            "line by line, exactly as the app renders it - cards and replay in one video.",
            c,
        )
        hold(Image.open(c), args.title_secs)

        for path, title, desc in FEATURES:
            c = work / (Path(path).stem + ".png")
            card(title, desc, c)
            hold(Image.open(c), args.title_secs)

            # Replay through the real window and take the frames it produced.
            d = work / (Path(path).stem + "-frames")
            d.mkdir()
            subprocess.run(
                [
                    sys.executable, "scripts/gui_movie.py", path,
                    "--frames-dir", str(d),
                    "--keep-frames",
                    "--cps", str(args.cps),
                    "--confirm-ms", str(args.confirm_ms),
                    "--menu-hold-ms", str(args.menu_hold_ms),
                    "--capture-fps", str(args.capture_fps),
                    "-o", str(work / (Path(path).stem + ".mp4")),
                ],
                check=True,
                cwd=str(ROOT),
            )
            captured = sorted(d.glob("frame-*.ppm"))
            if not captured:
                sys.exit(f"error: no frames captured for {path}")
            for f in captured:
                emit(Image.open(f))
            shutil.rmtree(d, ignore_errors=True)

        c = work / "end.png"
        card(
            "Reproduce it",
            "cargo build --features gui  &&  scripts/demo_movie.py",
            c,
        )
        hold(Image.open(c), args.title_secs)

        args.out.parent.mkdir(parents=True, exist_ok=True)
        cmd = [
            "ffmpeg", "-y", "-loglevel", "error",
            "-framerate", str(args.fps),
            "-i", str(frames / "frame-%05d.ppm"),
            "-vf", "scale=trunc(iw/2)*2:trunc(ih/2)*2",
            "-c:v", "libx264", "-pix_fmt", "yuv420p", "-crf", str(args.crf),
            "-movflags", "+faststart", str(args.out),
        ]
        print(f"[demo_movie] encoding {n} frames -> {args.out}")
        subprocess.run(cmd, check=True)
        print(
            f"[demo_movie] wrote {args.out} "
            f"({args.out.stat().st_size / 1024:.0f} KiB, {n} frames, {n / args.fps:.1f}s "
            f"@ {args.cps} chars/sec)"
        )
    finally:
        if args.keep_frames:
            print(f"[demo_movie] frames kept in {frames}")
        else:
            shutil.rmtree(work, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
