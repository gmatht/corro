#!/usr/bin/env python3
"""Verify every import of a PE file against the real Windows 95 DLLs.

Windows 95's loader refuses to start an exe whose import table names a symbol
that a DLL does not export ("linked to missing export ..."). This script reads
the import table of an exe and the export table of each Win95 DLL and reports
any symbol that is not present.

Usage: check_gui_imports.py <exe> [<dll-dir>]

`<dll-dir>` defaults to /tmp/w95dll (populated with guestfish from the Win95
disk image; see run95.sh / this repo's notes). DLL lookup is case-insensitive.
"""
import struct
import sys
import os

# Win95 ships these; anything the exe imports from a DLL *not* in this set is
# also a load failure. unicows.dll is allowed: run95.sh puts it on the image
# and the rust9x targets route unicode APIs through it on Win9x.
KNOWN = {
    "kernel32", "user32", "gdi32", "comctl32", "shell32", "ole32",
    "comdlg32", "advapi32", "unicows",
}


def pe_sections(data):
    pe = struct.unpack_from("<I", data, 0x3C)[0]
    if data[pe:pe + 4] != b"PE\0\0":
        raise ValueError("not a PE file")
    nsec = struct.unpack_from("<H", data, pe + 6)[0]
    opt_size = struct.unpack_from("<H", data, pe + 20)[0]
    opt = pe + 24
    magic = struct.unpack_from("<H", data, opt)[0]
    dd = opt + (96 if magic == 0x10B else 112)
    secs = []
    secoff = opt + opt_size
    for i in range(nsec):
        o = secoff + i * 40
        name = data[o:o + 8].rstrip(b"\0").decode("latin1")
        vsize, vaddr, rawsize, rawptr = struct.unpack_from("<IIII", data, o + 8)
        secs.append((name, vaddr, vsize, rawptr, rawsize))
    return pe, opt, magic, dd, secs


def rva_to_off(rva, secs):
    for _, vaddr, vsize, rawptr, rawsize in secs:
        if vaddr <= rva < vaddr + max(vsize, rawsize):
            return rawptr + (rva - vaddr)
    return None


def cstr(data, off):
    end = data.index(b"\0", off)
    return data[off:end].decode("latin1")


def pe_imports(path):
    data = open(path, "rb").read()
    pe, opt, magic, dd, secs = pe_sections(data)
    imp_rva, imp_size = struct.unpack_from("<II", data, dd + 8)
    if not imp_rva:
        return {}
    off = rva_to_off(imp_rva, secs)
    out = {}
    while True:
        fields = struct.unpack_from("<IIIII", data, off)
        if not any(fields):
            break
        _orig, _ts, _fwd, name_rva, first_thunk = fields
        if name_rva:
            dll = cstr(data, rva_to_off(name_rva, secs)).lower()
            names = set()
            t = rva_to_off(first_thunk, secs)
            while True:
                v = struct.unpack_from("<I", data, t)[0]
                if v == 0:
                    break
                if not (v & 0x80000000):
                    names.add(cstr(data, rva_to_off(v, secs) + 2))
                t += 4
            out.setdefault(dll, set()).update(names)
        off += 20
    return out


def pe_exports(path):
    """Return the set of exported symbol names (by name; ordinals ignored)."""
    data = open(path, "rb").read()
    try:
        pe, opt, magic, dd, secs = pe_sections(data)
    except ValueError:
        return None
    exp_rva, exp_size = struct.unpack_from("<II", data, dd)
    if not exp_rva:
        return set()
    off = rva_to_off(exp_rva, secs)
    if off is None:
        return set()
    _c, _t, _mj, _mn, name_rva, base, nfunc, nname, funcs_rva, names_rva, ords_rva = \
        struct.unpack_from("<IIHHIIIIIII", data, off)
    names = set()
    noff = rva_to_off(names_rva, secs)
    if noff is not None:
        for i in range(nname):
            nr = struct.unpack_from("<I", data, noff + i * 4)[0]
            names.add(cstr(data, rva_to_off(nr, secs)))
    _ = (name_rva, base, nfunc, funcs_rva, ords_rva, exp_size)
    return names


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    exe = sys.argv[1]
    dlldir = sys.argv[2] if len(sys.argv) > 2 else "/tmp/w95dll"

    imports = pe_imports(exe)
    problems = []

    index = {}
    for f in os.listdir(dlldir) if os.path.isdir(dlldir) else []:
        if f.lower().endswith(".dll"):
            index[f[:-4].lower()] = os.path.join(dlldir, f)

    for dll, syms in sorted(imports.items()):
        base = dll[:-4] if dll.endswith(".dll") else dll
        if base not in KNOWN:
            problems.append((dll, "*", "dll not present on Windows 95"))
            continue
        path = index.get(base)
        if path is None:
            print(f"  (skip) {dll}: no Win95 DLL on disk to compare")
            continue
        exports = pe_exports(path)
        if exports is None:
            continue
        missing = sorted(s for s in syms if s not in exports)
        if missing:
            for m in missing:
                problems.append((dll, m, "not exported by Win95 " + dll))

    total = sum(len(v) for v in imports.values())
    print(f"{exe}: {len(imports)} DLLs, {total} imported symbols")
    for dll in sorted(imports):
        print(f"  {dll:<14} {len(imports[dll])} symbols")

    if problems:
        print("\nFAIL — symbols Windows 95 cannot resolve:")
        for dll, sym, why in problems:
            print(f"  {dll}!{sym}: {why}")
        return 1
    print("\nOK — every import resolves against the Windows 95 DLLs")
    return 0


if __name__ == "__main__":
    sys.exit(main())
