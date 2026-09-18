#!/usr/bin/env python3
"""Build unsigned MSIX packages for corro from cross-compiled Windows exes.

An MSIX is an OPC zip: [Content_Types].xml + AppxManifest.xml +
AppxBlockMap.xml (sha256 over 64KiB blocks of every payload file) + payload.
Upload the .msix to Partner Center UNSIGNED -- the Store signs it after
certification, so no certificate is needed. (For local sideload testing you
DO need to sign it yourself on a Windows box; see README note at bottom.)

Usage:
    python3 scripts/package-msix.py --exe-x86 path/to/gcorro-x86.exe \
        --exe-x64 path/to/gcorro-x64.exe [--exe-arm64 path/to/gcorro-arm64.exe] \
        [--outdir dist] [--publisher "CN=..."] [--identity ...] [--version 0.7.0.0]

The --publisher MUST match the Publisher ID on your Partner Center account
(Account settings > ... or any existing submission's manifest); a mismatch
fails certification. Same for --identity (the reserved Store app name).
Defaults are placeholders that build a valid package but will NOT certify.
"""
import argparse
import base64
import hashlib
import os
import re
import struct
import sys
import zipfile

BLOCK = 65536
FIXED_DATE = (2020, 1, 1, 0, 0, 0)  # deterministic zips

CONTENT_TYPES = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\
<Default Extension="exe" ContentType="application/x-msdownload" />\
<Default Extension="png" ContentType="image/png" />\
<Default Extension="xml" ContentType="text/xml" />\
<Override PartName="/AppxManifest.xml" \
ContentType="application/vnd.ms-appx.manifest+xml" />\
<Override PartName="/AppxBlockMap.xml" \
ContentType="application/vnd.ms-appx.blockmap+xml" />\
</Types>
"""

MANIFEST = """\
<?xml version="1.0" encoding="utf-8"?>
<Package xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10" \
xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10" \
xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities" \
IgnorableNamespaces="uap rescap">
  <Identity Name="{identity}" Publisher="{publisher}" Version="{version}" ProcessorArchitecture="{arch}" />
  <Properties>
    <DisplayName>{display}</DisplayName>
    <PublisherDisplayName>{pubdisplay}</PublisherDisplayName>
    <Logo>Assets\\StoreLogo.png</Logo>
  </Properties>
  <Resources>
    <Resource Language="en-us" />
  </Resources>
  <Dependencies>
    <TargetDeviceFamily Name="Windows.Desktop" MinVersion="10.0.17763.0" MaxVersionTested="10.0.22621.0" />
  </Dependencies>
  <Capabilities>
    <rescap:Capability Name="runFullTrust" />
  </Capabilities>
  <Applications>
    <Application Id="App" Executable="gcorro.exe" EntryPoint="Windows.FullTrustApplication">
      <uap:VisualElements DisplayName="{display}" Description="{display} spreadsheet" BackgroundColor="#217346" \
Square150x150Logo="Assets\\Square150x150Logo.png" Square44x44Logo="Assets\\Square44x44Logo.png">
        <uap:DefaultTile Wide310x150Logo="Assets\\Wide310x150Logo.png" />
      </uap:VisualElements>
    </Application>
  </Applications>
</Package>
"""

# (filename, width, height) tile set; StoreLogo doubles as Properties/Logo.
ASSETS = [
    ("StoreLogo.png", 50, 50),
    ("Square44x44Logo.png", 44, 44),
    ("Square150x150Logo.png", 150, 150),
    ("Wide310x150Logo.png", 310, 150),
]


def make_tile(path, w, h):
    """Placeholder tile: green sheet with a white grid + highlighted cell."""
    from PIL import Image, ImageDraw
    img = Image.new("RGB", (w, h), (33, 115, 70))
    d = ImageDraw.Draw(img)
    step = max(8, min(w, h) // 8)
    for x in range(0, w, step):
        d.line([(x, 0), (x, h)], fill=(24, 84, 51))
    for y in range(0, h, step):
        d.line([(0, y), (w, y)], fill=(24, 84, 51))
    # highlighted A1 cell
    d.rectangle([step, step, 3 * step - 1, 2 * step - 1], fill=(255, 255, 255))
    d.rectangle([step, step, 2 * step - 1, 2 * step - 1], fill=(195, 230, 203))
    img.save(path, "PNG")


def block_hashes(data):
    return [hashlib.sha256(data[i:i + BLOCK]).digest()
            for i in range(0, max(len(data), 1), BLOCK)] if data else []


def b64(b):
    return base64.b64encode(b).decode("ascii")


def build_blockmap(files):
    """files: list of (arcname, bytes). Returns AppxBlockMap.xml bytes.
    File Names use BACKSLASH separators (MakeAppx convention, verified
    against a Microsoft-built package: 223/223 nested entries backslashed,
    0 forward-slashed). Forward slashes here break Windows' parser
    ("Error in parsing app package")."""
    out = ['<BlockMap xmlns="http://schemas.microsoft.com/appx/2010/blockmap" '
           'HashMethod="http://www.w3.org/2001/04/xmlenc#sha256">']
    for name, data in files:
        bname = name.replace("/", "\\")
        assert "/" not in bname, f"blockmap name not backslashed: {name}"
        # LfhSize = local file header size as WE write it: 30 + name length
        # (no extra fields; python zipfile seeks back and fills sizes, so no
        # data descriptor). The verifier recomputes hashes from these offsets.
        lfh = 30 + len(name.encode("utf-8"))
        out.append(f'  <File Name="{bname}" Size="{len(data)}" LfhSize="{lfh}">')
        for h in block_hashes(data):
            out.append(f'    <Block Hash="{b64(h)}" />')
        out.append("  </File>")
    out.append("</BlockMap>")
    return ("\n".join(out) + "\n").encode("utf-8")


def write_msix(path, payload_files):
    """payload_files: ORDERED list of (arcname, bytes).
    OPC readers (and Windows' parser in particular) expect
    [Content_Types].xml FIRST -- do NOT sort; callers pass MakeAppx order:
    Content_Types, BlockMap, Manifest, exe, Assets. Deterministic deflate."""
    names = [n for n, _ in payload_files]
    assert len(names) == len(set(names)), "duplicate arcnames"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for name, data in payload_files:
            zi = zipfile.ZipInfo(name, date_time=FIXED_DATE)
            zi.create_system = 0
            zi.external_attr = 0o600 << 16
            zi.extra = b""
            z.writestr(zi, data)


def verify_msix(path):
    """Re-read the package: check zip integrity, blockmap hashes, LfhSize."""
    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        assert bad is None, f"corrupt zip member: {bad}"
        names = z.namelist()
        assert "[Content_Types].xml" in names and "AppxManifest.xml" in names \
            and "AppxBlockMap.xml" in names, "missing OPC roots"
        import xml.etree.ElementTree as ET
        bm = ET.fromstring(z.read("AppxBlockMap.xml"))
        ns = {"b": "http://schemas.microsoft.com/appx/2010/blockmap"}
        for f in bm.findall("b:File", ns):
            name, size, lfh = f.get("Name"), int(f.get("Size")), int(f.get("LfhSize"))
            assert "/" not in name, f"forward slash in blockmap name: {name}"
            # blockmap names are backslashed; zip entries are forward-slashed
            zname = name.replace("\\", "/")
            info = z.getinfo(zname)
            # local header via raw offset (ZipExtFile yields the DATA stream,
            # not the header): flags@8, filename-len@26, extra-len@28.
            fp = open(path, "rb")
            fp.seek(info.header_offset)
            hdr = fp.read(30)
            fp.close()
            assert hdr[:4] == b"PK\x03\x04", f"bad local header for {name}"
            _flags, fn_len, extra_len = struct.unpack("<HHH", hdr[8:10] + hdr[26:30])
            assert 30 + fn_len + extra_len == lfh, f"LfhSize wrong for {name}"
            data = z.read(zname)
            assert len(data) == size, f"size wrong for {name}"
            blocks = f.findall("b:Block", ns)
            if not data:
                assert not blocks, f"empty file {name} should have no blocks"
                continue
            assert len(blocks) == (len(data) + BLOCK - 1) // BLOCK, \
                f"block count wrong for {name}"
            for i, blk in enumerate(blocks):
                h = hashlib.sha256(data[i * BLOCK:(i + 1) * BLOCK]).digest()
                assert b64(h) == blk.get("Hash"), f"block {i} hash wrong for {name}"
    return True


def cargo_version():
    with open("Cargo.toml") as f:
        m = re.search(r'^version\s*=\s*"([^"]+)"', f.read(), re.M)
    v = (m.group(1) if m else "0.0.0").split(".")
    return ".".join((v + ["0", "0", "0"])[:4])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--exe-x86", required=True)
    ap.add_argument("--exe-x64", required=True)
    ap.add_argument("--exe-arm64", default=None)
    ap.add_argument("--outdir", default="dist")
    ap.add_argument("--publisher", default="CN=8B3B25A0-23E7-4255-AAA7-58DC28AD6AA7")
    ap.add_argument("--identity", default="Dansted.org.corro")
    ap.add_argument("--pubdisplay", default="Dansted.org")
    ap.add_argument("--display", default="Corro")
    ap.add_argument("--version", default=None)
    args = ap.parse_args()
    version = args.version or cargo_version()

    os.makedirs("dist/msix-assets", exist_ok=True)
    for fname, w, h in ASSETS:
        p = os.path.join("dist/msix-assets", fname)
        if not os.path.exists(p):
            make_tile(p, w, h)

    arches = [("x86", args.exe_x86), ("x64", args.exe_x64)]
    if args.exe_arm64:
        arches.append(("arm64", args.exe_arm64))
    outputs = []
    for arch, exe in arches:
        with open(exe, "rb") as f:
            exebytes = f.read()
        manifest = MANIFEST.format(identity=args.identity, publisher=args.publisher,
                                   version=version, arch=arch, display=args.display,
                                   pubdisplay=args.pubdisplay).encode("utf-8")
        payload = [("[Content_Types].xml", CONTENT_TYPES.encode("utf-8"))]
        blockmap = build_blockmap([("AppxManifest.xml", manifest),
                                   ("gcorro.exe", exebytes)] +
                                  [(f"Assets/{fname}", open(os.path.join(
                                      "dist/msix-assets", fname), "rb").read())
                                   for fname, _, _ in ASSETS])
        payload += [("AppxBlockMap.xml", blockmap),
                    ("AppxManifest.xml", manifest),
                    ("gcorro.exe", exebytes)]
        for fname, _, _ in ASSETS:
            with open(os.path.join("dist/msix-assets", fname), "rb") as f:
                payload.append((f"Assets/{fname}", f.read()))
        out = os.path.join(args.outdir, f"corro-{version}-{arch}.msix")
        write_msix(out, payload)
        verify_msix(out)
        outputs.append((out, os.path.getsize(out)))
        print(f"wrote {out} ({os.path.getsize(out)} bytes) -- verified OK")
    if "PLACEHOLDER" in args.publisher + args.identity:
        print("NOTE: placeholder Identity/Publisher -- replace with your Partner",
              "Center values before submission (see script docstring).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
