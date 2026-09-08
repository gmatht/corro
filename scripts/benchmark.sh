#!/usr/bin/env bash
# PGO/PGSO training workload for corro.
# Designed to be run as: optimize-project.sh --train-cmd '.../benchmark.sh'
set -euo pipefail
cd "$(dirname "$0")/.."

BENCH_DIR="target/bench-train"
mkdir -p "$BENCH_DIR"
BIN="target/release/corro"
[ -f "target/x86_64-unknown-linux-gnu/release/corro" ] && BIN="target/x86_64-unknown-linux-gnu/release/corro"

# 1. Live TSV test: watch generates changing data, corro displays it
echo "=== Benchmark 1: Live TSV ==="
DOCS_TSV="docs/tests"
WATCH_LOG="$BENCH_DIR/watch.log"
pushd "$DOCS_TSV" >/dev/null
watch -n0.1 "bash tsv.sh 2>/dev/null" >"$WATCH_LOG" 2>&1 &
WATCH_PID=$!
popd >/dev/null
sleep 0.3

timeout 5 "$BIN" tsvlive.corro 2>/dev/null || true
kill "$WATCH_PID" 2>/dev/null || true

# 2. Movie demos: automated TUI navigation
echo "=== Benchmark 2: Movie demos ==="
cd scripts
go run movie_runner.go movie_demo.txt 2>/dev/null || true
cd ..

# 3. Basic file operations
echo "=== Benchmark 3: File ops ==="
"$BIN" -e '5,10' junk.corro 2>/dev/null || true
"$BIN" -e 'SELECT * WHERE A > 10' junk.corro 2>/dev/null || true

# 4. Subtotal test
echo "=== Benchmark 4: Subtotals ==="
"$BIN" subtotal.corro 2>/dev/null || true

echo "=== Benchmark complete ==="
