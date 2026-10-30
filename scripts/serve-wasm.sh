#!/usr/bin/env bash
# Build corro for wasm32 and serve `web/` so it can be opened in a browser.
#
#     scripts/serve-wasm.sh            # build, then serve on :8099
#     scripts/serve-wasm.sh 9000       # ... on another port
#     PORT=9000 scripts/serve-wasm.sh
#
# This is the only in-repo way to run the `wasm` backend; there is no
# `wasm-pack`/`trunk` dependency. Nothing here is run by `cargo test` or by CI's
# build job -- see docs/WASM_MISSING.md §8 for why the browser side is opt-in.
#
# Two versions have to agree or wasm-bindgen refuses the module it is given:
#
#   * the `wasm-bindgen` crate, pinned to one exact version in `Cargo.toml`;
#   * the `wasm-bindgen` CLI, which must be that same version.
#
# So the pin is read from the crate (not from a constant here, which would
# silently drift) and an installed CLI of any other version is rejected rather
# than run.
#
# `CARGO_TARGET_DIR` is honoured, so a working tree whose `target/` is a
# symlink to a volume that is not mounted still builds: set it to a directory
# that exists.
set -euo pipefail

cd "$(dirname "$0")/.."

PORT="${1:-${PORT:-8099}}"
OUT="$PWD/web"
TOOLCHAIN="${WASM_TOOLCHAIN:-stable}"
TARGET=wasm32-unknown-unknown

# `rust-toolchain.toml` pins nightly, which does not have the wasm32 standard
# library installed on most machines; `+stable` is what CI's wasm.yml uses and
# what this script builds with. The pin is deliberate, not a default to be
# overridden into a toolchain that cannot build the target.
say() { printf '%s\n' "$*"; }
die() { printf 'error: %s\n' "$*" >&2; exit 1; }

command -v cargo >/dev/null || die "cargo not on PATH"

# `grep -o -E --` rather than `grep -oE`: this repo's `grep` may be a ripgrep
# wrapper, and a bundled short cluster is read there as a flag (`-oE` ->
# "unrecognized flag --oE"), while the separated form means the same thing to
# both. The `--` ends option parsing before a pattern that starts with `^`.
PINNED=$(grep -o -E -- '^wasm-bindgen = "=([0-9.]+)"' Cargo.toml | grep -o -E -- '[0-9]+\.[0-9]+\.[0-9]+' || true)
[ -n "$PINNED" ] || die "no exact \`wasm-bindgen = \"=x.y.z\"\` pin in Cargo.toml"

if ! rustup target list --toolchain "$TOOLCHAIN" --installed 2>/dev/null | grep -q -x -- "$TARGET"; then
    say "== installing $TARGET for $TOOLCHAIN"
    rustup target add "$TARGET" --toolchain "$TOOLCHAIN"
fi

say "== building corro for $TARGET"
cargo "+$TOOLCHAIN" build --target "$TARGET" --features wasm --no-default-features

# Where cargo put the module, following a `target/` symlink or an explicit
# CARGO_TARGET_DIR.
TARGET_DIR=$(cargo "+$TOOLCHAIN" metadata --format-version 1 --no-deps 2>/dev/null \
    | grep -o -E -- '"target_directory":"[^"]+"' | head -1 | cut -d'"' -f4)
TARGET_DIR="${CARGO_TARGET_DIR:-$TARGET_DIR}"
WASM="$TARGET_DIR/$TARGET/debug/corro.wasm"
[ -s "$WASM" ] || die "no artifact at $WASM"

command -v wasm-bindgen >/dev/null || die "wasm-bindgen CLI not on PATH. Install the pinned version:
    cargo install wasm-bindgen-cli --version $PINNED --locked"

HAVE=$(wasm-bindgen --version | grep -o -E -- '[0-9]+\.[0-9]+\.[0-9]+' | head -1)
[ "$HAVE" = "$PINNED" ] || die "wasm-bindgen CLI is $HAVE but the crate is pinned to $PINNED:
    cargo install wasm-bindgen-cli --version $PINNED --locked --force"

# Remove the previous glue before regenerating it. `wasm-bindgen` writes the
# module under its content hash (`corro_bg.wasm` changes name when the code
# does), and an HTTP server that sends `Last-Modified` is entitled to answer a
# browser's revalidation with a 304 for the *old* file -- which a browser that
# is holding an old tab will then happily reuse. A directory that only ever
# holds the current pair cannot go stale.
rm -f "$OUT"/corro*.js "$OUT"/corro*_bg.wasm
say "== wasm-bindgen $PINNED -> web/"
wasm-bindgen --target web --no-typescript --out-dir "$OUT" "$WASM"

say ""
say "corro is at http://127.0.0.1:$PORT/  (Ctrl+C to stop)"
say ""
exec python3 -m http.server "$PORT" --directory "$OUT" --bind 127.0.0.1
