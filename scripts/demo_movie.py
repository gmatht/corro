#!/usr/bin/env python3
"""Record the corro GUI demo video.

Builds one video from a title card, a replay of each featured workbook, and a
closing card. The replays come from `scripts/gui_movie.py`, which plays the
movie in the real corro window (under a throwaway X server) and screenshots it —
so the video shows exactly what the application shows.

Usage:
    python3 scripts/demo_movie.py                 # write dist/corro-gui-movie.mp4
    python3 scripts/demo_movie.py --cps 1.6       # slower typing
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

# Live-collaboration demos: (pair, card title, card description, extra note).
# These are recorded by `two_window_movie.py`, which runs two front-ends over
# one shared log.
TWO_WINDOW = [
    (
        "gui-gui",
        "Two windows, one file",
        "corro's workbook is an append-only log, so two windows on the same file "
        "stay in sync with no save step: every edit is appended and the other "
        "window tails it while you work.",
        None,
    ),
    (
        "gui-tui",
        "GUI and terminal together",
        "The same log is shared across front-ends: a native window and a terminal "
        "session edit one workbook, and each sees the other's edits as they land.",
        None,
    ),
]

# Features replayed in the video: (workbook, card title, card description).
FEATURES = [
    (
        "docs/tests/subtotal.corro",
        "Subtotals and summary labels",
        "One marker in the left margin creates a subtotal or grand total column. "
        "corro never double counts: the grand total sums the subtotals, not the raw data.",
        None,
    ),
    (
        "docs/tests/main.corro",
        "Sheets, moves and formats",
        "A workbook is a log: sheets are created, copied and activated, rows are moved, "
        "and column formats are applied - each replayed as the interaction that produced it.",
        # This fixture is a feature tour that deliberately includes broken
        # formulas, so cells showing #NAME/#PARSE/#CIRC are real evaluator
        # output. Say so on the card instead of letting a viewer read them as
        # rendering bugs. `scripts/movie_errors.py` lists them from the log.
        "Note: this test workbook intentionally contains broken formulas, so a few "
        "cells show #NAME / #PARSE / #CIRC.",
    ),
]


def font(name: str, size: int) -> ImageFont.FreeTypeFont:
    for d in DEFAULT_FONT_DIRS:
        p = d / name
        if p.exists():
            return ImageFont.truetype(str(p), size)
    # Fall back to Pillow's bundled bitmap font rather than failing the run.
    return ImageFont.load_default()


def card(title: str, desc: str, out: Path, note: str | None = None) -> None:
    """Render a title/description card, optionally with a highlighted note."""
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
    if note:
        y += 18
        nlines, cur = [], ""
        for word in note.split():
            if len(cur) + len(word) + 1 > 68:
                nlines.append(cur)
                cur = word
            else:
                cur = (cur + " " + word).strip()
        if cur:
            nlines.append(cur)
        for ln in nlines:
            d.text((60, y), ln, font=body, fill=(230, 180, 110))
            y += 30
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
        default=3.25,
        help="typing speed in chars/sec for the replayed workbooks (default: 3.25; "
             "lower is slower)",
    )
    ap.add_argument("--confirm-ms", type=int, default=800, help="per-step hold while replaying")
    ap.add_argument("--menu-hold-ms", type=int, default=2800, help="menu flash hold while replaying")
    ap.add_argument(
        "--tempo",
        type=float,
        default=2.0,
        help="scales the two-window edit timings (2.0 = half speed, matching --cps)",
    )
    ap.add_argument("--fps", type=int, default=12, help="output frame rate")
    ap.add_argument("--crf", type=int, default=20, help="x264 quality (lower is better)")
    ap.add_argument("--title-secs", type=float, default=6.0, help="how long each card holds")
    # Kept at 8/s while everything else halves: slower pacing means longer
    # segments, and a lower capture rate would drop the frames that show the
    # replay advancing.
    ap.add_argument("--capture-fps", type=float, default=8.0, help="screenshot rate while replaying")
    ap.add_argument("--keep-frames", action="store_true", help="keep the intermediate frames")
    args = ap.parse_args()

    for tool in ("ffmpeg", "Xvfb", "xwd", "convert"):
        if shutil.which(tool) is None:
            sys.exit(f"error: {tool} not found on PATH")

    work = Path(tempfile.mkdtemp(prefix="corro-demo-"))
    # Each segment is encoded to its own file as soon as it is recorded, then
    # the pieces are concatenated. Copying every captured frame into one
    # directory (the obvious approach) needs a second copy of 1-3 GB of raw
    # frames per replay, which ran the disk out of space mid-record.
    segments: list[Path] = []
    cards = work / "cards"
    cards.mkdir()

    def add_card(im: Image.Image, secs: float, tag: str) -> None:
        """Encode a still card as a `secs`-long segment."""
        png = cards / f"{tag}.png"
        im.save(png)
        out = work / f"seg-{len(segments):02d}.mp4"
        subprocess.run(
            ["ffmpeg", "-y", "-loglevel", "error",
             "-loop", "1", "-framerate", str(args.fps), "-i", str(png),
             "-t", str(secs),
             "-vf", "scale=trunc(iw/2)*2:trunc(ih/2)*2",
             "-c:v", "libx264", "-pix_fmt", "yuv420p", "-crf", str(args.crf),
             str(out)],
            check=True,
        )
        segments.append(out)

    def add_frames(pattern: Path, count: int, scale: str | None, fps: int) -> None:
        """Encode a captured frame sequence as a segment."""
        if count == 0:
            return
        out = work / f"seg-{len(segments):02d}.mp4"
        vf = scale or "scale=trunc(iw/2)*2:trunc(ih/2)*2"
        subprocess.run(
            ["ffmpeg", "-y", "-loglevel", "error",
             "-framerate", str(fps), "-i", str(pattern),
             "-vf", vf,
             "-c:v", "libx264", "-pix_fmt", "yuv420p", "-crf", str(args.crf),
             str(out)],
            check=True,
        )
        segments.append(out)

    try:
        c = work / "title.png"
        card(
            "GUI movie mode",
            "The --movie demo mode plays back a .corro workbook in the corro window itself, "
            "line by line, exactly as the app renders it - cards and replay in one video.",
            c,
        )
        add_card(Image.open(c), args.title_secs, f"card-{len(segments)}")

        for path, title, desc, note in FEATURES:
            c = work / (Path(path).stem + ".png")
            card(title, desc, c, note)
            add_card(Image.open(c), args.title_secs, f"card-{len(segments)}")

            # Replay through the real window and take the frames it produced.
            d = work / (Path(path).stem + "-frames")
            d.mkdir()
            subprocess.run(
                [
                    sys.executable, "scripts/gui_movie.py", path,
                    "--frames-dir", str(d),
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
            add_frames(d / "frame-%05d.ppm", len(captured), None, args.fps)
            # The capture is 1-3 GB of raw PPM; it is only needed until this
            # segment is encoded. Keeping every segment's frames until the end
            # of the run exhausts the disk.
            shutil.rmtree(d, ignore_errors=True)

        for pair, title, desc, note in TWO_WINDOW:
            c = work / f"two-{pair}.png"
            card(title, desc, c, note)
            add_card(Image.open(c), args.title_secs, f"card-{len(segments)}")

            d = work / f"two-{pair}-frames"
            d.mkdir()
            subprocess.run(
                [
                    sys.executable, "scripts/two_window_movie.py",
                    "--pair", pair,
                    # Match the replay pacing: the collaboration segments
                    # halve in speed along with everything else.
                    "--tempo", str(args.tempo),
                    "--frames-dir", str(d),
                    "-o", str(work / f"two-{pair}.mp4"),
                ],
                check=True,
                cwd=str(ROOT),
            )
            captured = sorted(d.glob("frame-*.ppm"))
            if not captured:
                sys.exit(f"error: no frames captured for {pair}")
            # The capture is two windows side by side on a 2400px screen; scale
            # to the video's width so the whole demo stays one shape.
            add_frames(
                d / "frame-%05d.ppm", len(captured),
                f"scale={WIDTH}:trunc(ih/2)*2", args.fps,
            )
            shutil.rmtree(d, ignore_errors=True)

        c = work / "end.png"
        card(
            "Reproduce it",
            "cargo build --features gui  &&  scripts/demo_movie.py",
            c,
        )
        add_card(Image.open(c), args.title_secs, f"card-{len(segments)}")

        if not segments:
            sys.exit("error: nothing was recorded")
        args.out.parent.mkdir(parents=True, exist_ok=True)
        listing = work / "segments.txt"
        listing.write_text("".join(f"file '{p.name}'" + chr(10) for p in segments))
        print(f"[demo_movie] concatenating {len(segments)} segments -> {args.out}")
        subprocess.run(
            ["ffmpeg", "-y", "-loglevel", "error",
             "-f", "concat", "-safe", "0", "-i", str(listing),
             "-c", "copy", "-movflags", "+faststart", str(args.out)],
            check=True,
        )
        print(
            f"[demo_movie] wrote {args.out} "
            f"({args.out.stat().st_size / 1024:.0f} KiB) @ {args.cps} chars/sec"
        )
    finally:
        if args.keep_frames:
            print(f"[demo_movie] work kept in {work}")
        else:
            shutil.rmtree(work, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
