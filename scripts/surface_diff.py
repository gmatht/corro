#!/usr/bin/env python3
"""Compare the zork backend's public method surface against the other adapters.

Used to produce the "no open gaps" measurement in `docs/ZORK_MISSING.md`.
Kept in the repo rather than in a shell one-liner so the claim is
re-checkable: a number in a document that nobody can recompute is a number
nobody should trust.

    python3 scripts/surface_diff.py            # the whole picture
    python3 scripts/surface_diff.py --missing  # just what zork lacks

Method: for each adapter, walk top-level `impl <Type> { .. }` blocks and
collect the `pub fn` names. Then diff zork's set against the union of GTK,
NWG and pancurses. "Zork lacks" is a name no other adapter has either —
that is the honest definition of a parity gap, and it is deliberately not
"a name zork lacks", because a headless model legitimately has methods the
GUI backends do not (there is no compositor to hold the state).
"""

import re
import sys
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "rustxWidgets" / "rswidgets" / "src"

ADAPTERS = {
    "gtk": "backends_gtk_adapter_impl.rs",
    "nwg": "backends_nwg_adapter.rs",
    "panc": "backends_pancurses_adapter.rs",
    "zork": "backends_zork_adapter.rs",
}

# The type names whose method sets are compared. Deliberately explicit: a
# blanket "any impl" sweep picks up platform-internal helpers that are not
# part of the widget surface at all.
WIDGETS = [
    "Window", "Button", "Label", "BoxWidget", "Grid", "Entry", "Menu",
    "MenuBar", "SimpleAction", "Dialog", "DropDown", "CheckButton",
    "RadioButton", "TextView", "Canvas", "Overlay", "ScrolledWindow",
    "Fixed", "Spreadsheet", "Application",
]

IMPL_RE = re.compile(r"^(\s*)impl(?:\s+<[^>]*>)?\s+([A-Za-z_][A-Za-z0-9_]*)")
PUB_FN_RE = re.compile(r"\bpub fn ([A-Za-z_][A-Za-z0-9_]*)")


def methods(path: Path, types) -> set:
    """The `pub fn` names on the named types, in this file."""
    found = set()
    in_impl = False
    cur = None
    impl_indent = 0
    for line in path.read_text().split("\n"):
        m = IMPL_RE.match(line)
        if m and cur is None:
            name = m.group(2)
            if name in types:
                in_impl = True
                cur = name
                impl_indent = len(m.group(1))
        if in_impl:
            pm = PUB_FN_RE.search(line)
            if pm:
                found.add(f"{cur}.{pm.group(1)}")
            # The closing brace at the impl's own indent level ends it.
            if line.strip() == "}" and len(line) - len(line.lstrip()) == impl_indent:
                in_impl = False
                cur = None
    return found


def free_functions(path: Path) -> set:
    """Module-level `pub fn` names — the free-function API, not widget methods.

    Indentation-insensitive on purpose: the GTK, NWG and pancurses adapters
    wrap their whole body in a `mod`, so their `pub fn`s are indented while
    the zork adapter's are not. Matching only the column-0 form would report
    0 free functions for three of the four and quietly understate the diff.
    """
    names = set()
    for line in path.read_text().split("\n"):
        # Skip anything that is clearly inside a doc comment.
        m = PUB_FN_RE.search(line)
        if not m:
            continue
        before = line[: m.start()]
        # A `pub fn` preceded only by whitespace/`pub ` is a definition; one
        # preceded by other code is a call site in a body.
        if before.strip() in ("", "pub", "pub unsafe"):
            names.add(m.group(1))
    return names


def main() -> int:
    missing_only = "--missing" in sys.argv
    types = set(WIDGETS)
    methods_by = {k: methods(SRC / v, types) for k, v in ADAPTERS.items()}
    free_by = {k: free_functions(SRC / v) for k, v in ADAPTERS.items()}

    others = methods_by["gtk"] | methods_by["nwg"] | methods_by["panc"]
    zork = methods_by["zork"]

    if not missing_only:
        for k in ADAPTERS:
            print(f"{k}: {len(methods_by[k])} methods, {len(free_by[k])} free functions")
        print()

    gaps = sorted(others - zork)
    print(f"methods zork lacks ({len(gaps)}):")
    for g in gaps:
        where = [k for k in ("gtk", "nwg", "panc") if g in methods_by[k]]
        print(f"  {g}   ({', '.join(where)})")
    print()

    other_free = free_by["gtk"] | free_by["nwg"] | free_by["panc"]
    free_gaps = sorted(other_free - free_by["zork"])
    print(f"free functions zork lacks ({len(free_gaps)}):")
    for g in free_gaps:
        where = [k for k in ("gtk", "nwg", "panc") if g in free_by[k]]
        print(f"  {g}   ({', '.join(where)})")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
