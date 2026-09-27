"""Is this PNG a corro spreadsheet, or the simulator's home screen?

Used by .github/workflows/ios.yml, in two places:

* the screenshot step, as a pre-filter deciding whether to re-activate the app
  and try again;
* the verify step, as the hard gate.

WHY IT ASKS ABOUT *LIGHT SURFACE* AND NOT ABOUT GREYSCALE
-------------------------------------------------------
The first version of this check measured "how much of the image is
achromatic" (max(r,g,b) - min(r,g,b) <= 6), on the premise that "corro's
chrome is neutral throughout". Theming killed that premise: rswidgets'
`Role::Header` is (0.88, 0.90, 0.93) and `Role::Text` is (0.05, 0.05, 0.10)
-- both blue-tinged, a spread of 13 against a threshold of 6. So the check
rejected corro's own sheet while passing for reasons that had nothing to do
with the sheet.

A spreadsheet is a LIGHT surface carrying dark text. The springboard is a
saturated wallpaper, which has almost no near-white in it. Measured on the two
real images this repository has, rather than guessed:

    corro's sheet     light 31.0%   dark  1.3%
    iOS springboard   light 15.4%   dark 48.0%

Only `light` separates them, so only `light` is the gate. A first draft of the
replacement also required dark ink, reasoning that a grid has dark text and a
wallpaper does not -- which is backwards, because the wallpaper is the DARKER
image (48% vs 1.3%). That version rejected the known-good frame; measuring the
two real images is what caught it. `dark` is still reported, so that the next
reader can see why it is not the gate.

The gate is 0.22, which sits between the two measured values with margin on
both sides.

EXIT CODES
----------
0  looks like a corro sheet
1  not a PNG, or unreadable
2  measured, but does not look like a corro sheet

A blank white app passes this check: it is a light surface. Rejecting that is
the cell-text assertion's job, which lives in the workflow rather than here.
"""

import struct
import sys
import zlib

# Reference measurements, printed so a run's numbers can be compared to them.
LIGHT_GATE = 0.22


def decode_png(path):
    """(width, height, channels, rows) for a non-interlaced 8-bit PNG.

    The unfiltering loop is the standard five-filter PNG reconstruction; it is
    spelled out rather than pulled from a library because the workflow has no
    package install step and this has to run on a bare runner.
    """
    data = open(path, "rb").read()
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("not a PNG")
    pos, idat, w, h, ctype = 8, b"", 0, 0, 0
    while pos < len(data):
        ln = struct.unpack(">I", data[pos:pos + 4])[0]
        typ = data[pos + 4:pos + 8]
        body = data[pos + 8:pos + 8 + ln]
        if typ == b"IHDR":
            w, h, _bitdepth, ctype = struct.unpack(">IIBB", body[:10])
        elif typ == b"IDAT":
            idat += body
        elif typ == b"IEND":
            break
        pos += 12 + ln
    channels = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[ctype]
    raw = zlib.decompress(idat)
    stride = w * channels
    rows, prev = [], bytearray(stride)
    i = 0
    for _y in range(h):
        f = raw[i]
        i += 1
        line = bytearray(raw[i:i + stride])
        i += stride
        if f == 1:
            for x in range(channels, stride):
                line[x] = (line[x] + line[x - channels]) & 0xFF
        elif f == 2:
            for x in range(stride):
                line[x] = (line[x] + prev[x]) & 0xFF
        elif f == 3:
            for x in range(stride):
                a = line[x - channels] if x >= channels else 0
                line[x] = (line[x] + ((a + prev[x]) >> 1)) & 0xFF
        elif f == 4:
            for x in range(stride):
                a = line[x - channels] if x >= channels else 0
                b = prev[x]
                c = prev[x - channels] if x >= channels else 0
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                pr = a if (pa <= pb and pa <= pc) else (b if pb <= pc else c)
                line[x] = (line[x] + pr) & 0xFF
        rows.append(bytes(line))
        prev = line
    return w, h, channels, rows


def light_fractions(w, h, ch, rows):
    """(light_fraction, dark_fraction) over the lower two thirds."""
    tot = light = dark = 0
    for yy in range(h // 3, h, 2):
        row = rows[yy]
        for xx in range(0, w, 2):
            o = xx * ch
            lum = (row[o] + row[o + 1] + row[o + 2]) / 3.0
            tot += 1
            if lum >= 200:
                light += 1
            if lum < 110:
                dark += 1
    t = max(tot, 1)
    return light / t, dark / t


def measure(path):
    """(w, h, light_fraction, dark_fraction) over the lower two thirds."""
    w, h, ch, rows = decode_png(path)
    light, dark = light_fractions(w, h, ch, rows)
    return w, h, light, dark


# The two data columns the scripted edits write into, as x bands. These are
# the same bands the workflow's cell-text assertion uses, so "looks like a
# populated sheet" means the same thing in both places.
COL_A = (90, 190)
COL_B = (190, 300)
# A row band that skips the chrome: the formula bar and the column headers are
# always drawn, so ink up there proves nothing about the sheet.
BODY_ROWS = (240, None)


def column_ink(w, h, ch, rows, x0, x1):
    """Ink pixels in an x band, over the sheet body rather than the chrome."""
    lo = BODY_ROWS[0]
    hi = BODY_ROWS[1] if BODY_ROWS[1] is not None else h - 60
    return sum(1 for yy in range(lo, min(hi, len(rows)), 2)
               for xx in range(x0, x1, 2) if rows[yy][xx * ch] < 100)


def main(argv):
    if len(argv) < 2:
        print("usage: looks_like_corro.py <png> [--populated]", file=sys.stderr)
        return 1
    try:
        w, h, ch, rows = decode_png(argv[1])
        light, dark = light_fractions(w, h, ch, rows)
    except Exception as e:  # unreadable/corrupt: not a pass
        print(f"could not measure {argv[1]}: {e}", file=sys.stderr)
        return 1
    a = column_ink(w, h, ch, rows, *COL_A)
    b = column_ink(w, h, ch, rows, *COL_B)
    print(f"{w}x{h} light={light:.1%} dark={dark:.1%} "
          f"(gate: light >= {LIGHT_GATE:.0%})  ink colA={a} colB={b}")
    # `populated` additionally requires cell text, and is what the screenshot
    # step waits for. Without it the loop accepted the *chrome-only* frame on
    # its first attempt and the verify step then failed on a blank sheet -
    # two checks disagreeing about the same image, which is the confusion the
    # `--populated` flag removes.
    if len(argv) > 2 and argv[2] == "--populated":
        if a < 50 or b < 50:
            print("  ...a light surface, but no cell text in the data columns yet")
            return 2
        print("  ...and cell text in two data columns")
    return 0 if light >= LIGHT_GATE else 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
