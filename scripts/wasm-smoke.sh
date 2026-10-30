#!/usr/bin/env bash
# Build the wasm app, serve it, and drive it in a headless browser.
#
#     scripts/wasm-smoke.sh                  # build + serve + test, then stop
#     PORT=8099 scripts/wasm-smoke.sh        # against a server already running
#
# Opt-in on purpose. This is the only test of the `wasm` backend that runs the
# app, and it is not part of `cargo test`: it needs a browser and a local HTTP
# server, and neither is something a unit-test run should silently acquire.
# CI's `wasm.yml` builds the module (`-D warnings`, asserts the wasm magic
# number) and stops there; see docs/WASM_MISSING.md §8.
#
# Exit codes: 0 pass, 1 a check failed, 2 the harness could not run (no
# browser, no module), 3 the build failed.
set -euo pipefail

cd "$(dirname "$0")/.."

PORT="${PORT:-8099}"
URL_="http://127.0.0.1:$PORT/"
SERVER_PID=""

cleanup() {
    [ -n "$SERVER_PID" ] && kill "$SERVER_PID" 2>/dev/null || true
}
trap cleanup EXIT

# A server the caller already started (or a page they are looking at) is
# reused: rebuilding behind a running server is not what "smoke test" means
# here, and the caller has said which URL they want.
if curl -fsS -o /dev/null --max-time 2 "$URL_"; then
    echo "== using the page already served at $URL_"
else
    echo "== building and serving"
    # `PORT=` is passed only as the port to listen on, not to skip the build:
    # `serve-wasm.sh` reads the same variable, and inheriting it here would
    # make the build skippable by setting an environment variable, which is
    # exactly what an opt-in harness must not allow.
    scripts/serve-wasm.sh "$PORT" >/tmp/corro-wasm-serve.log 2>&1 &
    SERVER_PID=$!
    for _ in $(seq 1 300); do
        if curl -fsS -o /dev/null --max-time 1 "$URL_"; then break; fi
        if ! kill -0 "$SERVER_PID" 2>/dev/null; then
            echo "== serve-wasm.sh exited; its log:" >&2
            cat /tmp/corro-wasm-serve.log >&2
            exit 3
        fi
        sleep 1
    done
    curl -fsS -o /dev/null --max-time 2 "$URL_" || { cat /tmp/corro-wasm-serve.log >&2; exit 3; }
fi

exec node scripts/wasm_smoke.mjs --url "$URL_" "$@"
