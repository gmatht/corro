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
# Screen the two-window capture runs on (`two_window_movie.py` puts one 1200x800
# window at x=0 and another at x=1200 on a 2400x820 screen).
TWO_WINDOW_SCREEN = (2400, 820)

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
#
# `main.corro` used to be the second entry. It is a 241-step regression fixture
# that types and retypes values — accurate as a test, tedious to watch — so its
# slot is now the menu tour below. The fixture stays in the test suite.
FEATURES = [
    (
        "docs/tests/subtotal.corro",
        "Subtotals and summary labels",
        "One marker in the left margin creates a subtotal or grand total column. "
        "corro never double counts: the grand total sums the subtotals, not the raw data.",
        None,
    ),
]

# The menu tour that closes the video: `MS:Section>Item#index` stops, run by the
# movie driver after its replay finishes. Every stop opens the real menu and
# dispatches the item through the app's own action path. Items that open a file
# dialog, write a file, launch an editor, need typed input or quit the app are
# deliberately absent — the tour has to run unattended.
#
# `index` is the item's row within its popover, which is where the pointer is
# aimed; it does not affect which action runs.
MENU_TOUR = [
    ("File", "Default width", 4),
    ("File", "Column width", 4),
    ("File", "Sort view", 5),
    ("File", "Persist sort", 6),
    ("Edit", "Select all", 3),
    ("Edit", "Copy", 1),
    ("Edit", "Cut", 0),
    ("Insert", "Rows", 0),
    ("Insert", "Cols", 3),
    ("Insert", "Date", 6),
    ("Insert", "Time", 7),
    ("Insert", "Special Char", 4),
    ("Insert", "Aggregate", 5),
    ("Format", "All", 0),
    ("Format", "Currency ($)", 1),
    ("Format", "Right", 2),
    ("Format", "Reset", 3),
    ("Sheet", "New sheet", 2),
    ("Sheet", "Rename sheet", 3),
    ("Sheet", "Copy sheet", 4),
    ("Sheet", "Next sheet", 1),
    ("Sheet", "Prev sheet", 0),
    ("Sheet", "Balance books", 7),
    ("Help", "Row ops", 1),
    ("Help", "Col ops", 2),
]

# Milliseconds between tour stops. Long enough to watch the pointer travel, the
# menu open and the item take effect, short enough to stay watchable.
MENU_TOUR_STEP_MS = 1500

# Workbook the tour runs on: a small, readable sheet, so the effects of the
# Sheet and Format items are visible against it.
MENU_TOUR_WORKBOOK = """SET A1 1
SET A2 7
SET A3 4
SET B1 2
SET B2 3
SET B3 6
"""


def menu_tour_script() -> str:
    """The `CORRO_MENU_TOUR` value for [`MENU_TOUR`]."""
    return ",".join(
        f"{i * MENU_TOUR_STEP_MS}:{section}>{item}#{index}"
        for i, (section, item, index) in enumerate(MENU_TOUR)
    )


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
        default=6.5,
        help="typing speed in chars/sec for the replayed workbooks (default: 6.5; "
             "lower is slower)",
    )
    ap.add_argument("--confirm-ms", type=int, default=400, help="per-step hold while replaying")
    ap.add_argument("--menu-hold-ms", type=int, default=1400, help="menu flash hold while replaying")
    ap.add_argument(
        "--tempo",
        type=float,
        default=1.0,
        help="scales the two-window edit timings (1.0 = the scripted speed, "
             "matching --cps; 2.0 = half speed)",
    )
    # Kept at the capture rate: downsampling 16/s to 12/s would throw away a
    # quarter of the frames, and at 6.5 chars/sec a character only spans about
    # two frames to begin with.
    ap.add_argument("--fps", type=int, default=16, help="output frame rate")
    ap.add_argument("--crf", type=int, default=20, help="x264 quality (lower is better)")
    ap.add_argument("--title-secs", type=float, default=3.0, help="how long each card holds")
    # Must stay well above the typing rate or the per-character animation falls
    # between captured frames and a replay reads as values simply appearing.
    # At the default 6.5 chars/sec a character lands every ~154ms, so 16/s
    # captures each one about twice.
    ap.add_argument("--capture-fps", type=float, default=16.0, help="screenshot rate while replaying")
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
                    # The cap is a per-segment safety net, and the 6000 default
                    # is sized for the old 8/s rate. At 16/s a long workbook
                    # (`main.corro` is 241 steps) runs past it and the segment
                    # is silently truncated, so scale the allowance with the
                    # capture rate.
                    "--max-frames", str(int(args.capture_fps * 900)),
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

        # The menu tour: the pointer travels to each menu, the menu opens, and
        # the item fires — closing the video on the menus rather than on another
        # log replay. `gui_movie.py` hands the tour to the driver, which runs it
        # once the (trivial) replay finishes.
        c = work / "menu-tour.png"
        card(
            "The menus",
            "Every item in corro's menus, driven end to end: the pointer travels to the "
            "menu, it opens, and the item runs through the same action path a click uses.",
            c,
            "File, Edit, Insert, Format, Sheet and Help. Items that would open a dialog, "
            "write a file, launch an editor or quit are left out — this tour runs unattended.",
        )
        add_card(Image.open(c), args.title_secs, f"card-{len(segments)}")

        d = work / "menu-tour-frames"
        d.mkdir()
        tour_work = work / "menu-tour-base.corro"
        tour_work.write_text(MENU_TOUR_WORKBOOK)
        tour_seconds = len(MENU_TOUR) * MENU_TOUR_STEP_MS / 1000.0
        subprocess.run(
            [
                sys.executable, "scripts/gui_movie.py", str(tour_work),
                "--frames-dir", str(d),
                # The replay itself is three short lines; the tour is the
                # content, so type fast and hold briefly.
                "--cps", str(max(args.cps, 30)),
                "--confirm-ms", "60",
                "--menu-hold-ms", "200",
                "--capture-fps", str(args.capture_fps),
                "--max-frames", str(int(args.capture_fps * (tour_seconds + 30))),
                "--tour", menu_tour_script(),
                "-o", str(work / "menu-tour.mp4"),
            ],
            check=True,
            cwd=str(ROOT),
        )
        captured = sorted(d.glob("frame-*.ppm"))
        if not captured:
            sys.exit("error: no frames captured for the menu tour")
        add_frames(d / "frame-%05d.ppm", len(captured), None, args.fps)
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
                    # Match the replay pacing: both kinds of segment run at
                    # the same speed in the finished video.
                    "--tempo", str(args.tempo),
                    # Match the replay's typing rate so both kinds of segment
                    # read the same way in the finished video.
                    "--typing-cps", str(args.cps),
                    "--frames-dir", str(d),
                    "-o", str(work / f"two-{pair}.mp4"),
                ],
                check=True,
                cwd=str(ROOT),
            )
            captured = sorted(d.glob("frame-*.ppm"))
            if not captured:
                sys.exit(f"error: no frames captured for {pair}")
            # Two 1200x800 windows sit side by side on a 2400x820 screen. The
            # pair has to be scaled by the *same* factor on both axes or every
            # window comes out distorted: halving the width alone (2400 ->
            # 1200) left the height at 820, which stretched each window 2x too
            # wide; picking a smaller width factor than the height factor (the
            # 600x410 this used to do) squashed each one 2x too narrow. Scale
            # uniformly to the video width — 2400x820 -> 1200x410, so one window
            # is 600x410 against a true 1.50 aspect — then centre the pair
            # vertically.
            #
            # The target size is written out literally rather than derived from
            # `iw`/`ih`: inside a filter chain `ih` still refers to the
            # *original* input, so `pad=...:trunc(ih/2)*2` asks for a box
            # shorter than its input and ffmpeg rejects the whole graph.
            scaled_w = WIDTH
            scaled_h = round(TWO_WINDOW_SCREEN[1] * WIDTH / TWO_WINDOW_SCREEN[0])
            scaled_h -= scaled_h % 2  # yuv420p wants even dimensions
            add_frames(
                d / "frame-%05d.ppm", len(captured),
                f"scale={scaled_w}:{scaled_h},"
                f"pad={WIDTH}:{HEIGHT}:0:{(HEIGHT - scaled_h) // 2}:color=0x12161e",
                args.fps,
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
