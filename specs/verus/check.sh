#!/usr/bin/env bash
# Verify the Verus lemmas in this directory.
#
#     specs/verus/check.sh                     # both files
#     specs/verus/check.sh scan_order_lemmas   # one
#
# Opt-in, like `scripts/wasm-smoke.sh`: this needs a Verus install, and a
# `cargo test` that silently requires one is a test that fails for reasons that
# have nothing to do with the code under test. Nothing else in the repo depends
# on these files compiling -- they are lemmas, not a module of the crate.
#
# Verus is looked for in this order:
#   1. `$VERUS` (path to the `verus` binary),
#   2. `verus` on `PATH`,
#   3. the newest `/opt/verus-*` directory.
#
# Exit codes: 0 all verified, 1 a lemma failed to verify, 2 Verus not found.
set -euo pipefail

cd "$(dirname "$0")"

find_verus() {
    if [ -n "${VERUS:-}" ]; then printf '%s\n' "$VERUS"; return; fi
    if command -v verus >/dev/null 2>&1; then command -v verus; return; fi
    # `sort -V` puts 0.2026.10.18 after 0.2026.6.7, which plain sort does not.
    ls -d /opt/verus-* 2>/dev/null | sort -V | tail -1 | sed 's|$|/verus|'
}

VERUS_BIN=$(find_verus)
if [ -z "$VERUS_BIN" ] || [ ! -x "$VERUS_BIN" ]; then
    printf 'error: no verus binary found.\n' >&2
    printf '  Set $VERUS=/path/to/verus, put verus on PATH, or install it under /opt/verus-*.\n' >&2
    exit 2
fi

# `verus --version` prints a crate name and a blank line rather than a version
# string in 0.2026.10.x, so the version is read from the install's own
# `version.json` when it is there, and the path is reported otherwise.
VERUS_DIR=$(dirname "$VERUS_BIN")
if [ -f "$VERUS_DIR/version.json" ]; then
    printf 'using %s (%s)\n' "$VERUS_BIN" \
        "$(grep -o '"version": *"[^"]*"' "$VERUS_DIR/version.json" | head -1 | cut -d'"' -f4)"
else
    printf 'using %s\n' "$VERUS_BIN"
fi

# Run under a scratch directory: Verus writes `foo.rlib`/`foo.vir` next to its
# input, and a `*.vir` in `specs/` is not something to commit.
# `CARGO_TARGET_DIR`-style override, otherwise a fresh temp dir.
mkdir -p "${TMPDIR:-/tmp}"
WORK=$(mktemp -d "${TMPDIR:-/tmp}/verus-specs.XXXXXX")
trap 'rm -rf "$WORK"' EXIT
SOURCES=$(pwd)

files=("$@")
if [ ${#files[@]} -eq 0 ]; then
    files=()
    for f in ./*.rs; do files+=("$f"); done
fi

status=0
for f in "${files[@]}"; do
    [ -f "$f" ] || f="$f.rs"
    case "$f" in /*) ;; *) f="$SOURCES/$f" ;; esac
    printf '== %s\n' "${f#"$SOURCES"/}"
    # `--crate-type=lib`: these are libraries of lemmas, with no `main` to run.
    # Verification is what matters; the resulting `.rlib` is not used.
    if ! out=$(cd "$WORK" && "$VERUS_BIN" --crate-type=lib "$f" 2>&1); then
        printf '%s\n' "$out"
        status=1
        continue
    fi
    # The summary is the evidence, so it is printed rather than just an "ok".
    printf '%s\n' "$out" | grep -E 'verification results|error' || printf '%s\n' "$out"
    # A run that reports zero verified lemmas (e.g. an empty file) is a failure
    # that "0 errors" would otherwise hide.
    if printf '%s\n' "$out" | grep -q 'verification results:: 0 verified'; then
        printf 'error: %s verified nothing\n' "$f" >&2
        status=1
    fi
done

exit "$status"
