#!/usr/bin/env bash
# Capture a corro GUI segfault with a full backtrace.
#
# Usage:
#   ./scripts/catch-gui-crash.sh            # no file argument
#   ./scripts/catch-gui-crash.sh foo.corro  # with a workbook
#
# Writes the backtrace to /tmp/corro-gui-crash.txt and also prints it.
# Tries three capture paths, in order: gdb, core dump, ulimit -c.
set -u

cd "$(dirname "$0")/.."
BIN=./target/debug/corro
OUT=/tmp/corro-gui-crash.txt
: > "$OUT"

echo "== corro GUI crash capture ==" | tee -a "$OUT"
echo "binary: $BIN"                     | tee -a "$OUT"
echo "DISPLAY=${DISPLAY:-<unset>}"      | tee -a "$OUT"
echo "args: $*"                         | tee -a "$OUT"

if [ ! -x "$BIN" ]; then
  echo "!! $BIN not built. Run: cargo build --features gui" | tee -a "$OUT"
  exit 2
fi

# Make any core dump land somewhere we can find it.
CORE_DIR=/tmp/corro-cores
mkdir -p "$CORE_DIR"
old_pattern=$(cat /proc/sys/kernel/core_pattern 2>/dev/null || echo "")
echo "core_pattern (current): $old_pattern" | tee -a "$OUT"

echo | tee -a "$OUT"
echo "== running under gdb ==" | tee -a "$OUT"
if command -v gdb >/dev/null 2>&1; then
  # Auto-answer the debuginfod prompt so this never blocks on stdin.
  gdb -batch \
      -ex 'set pagination off' \
      -ex 'set debuginfod enabled off' \
      -ex 'handle SIG33 nostop noprint pass' \
      -ex run \
      -ex 'echo \n===== BACKTRACE =====\n' \
      -ex 'bt 60' \
      -ex 'info registers rip rax rbx rcx rdx rsi rdi' \
      -ex 'echo \n===== THREADS =====\n' \
      -ex 'thread apply all bt 8' \
      --args "$BIN" --gui "$@" 2>&1 | tee -a "$OUT"
  rc=${PIPESTATUS[0]}
  echo "gdb exit: $rc" | tee -a "$OUT"
  if grep -qiE 'SIGSEGV|Segmentation fault' "$OUT"; then
    echo | tee -a "$OUT"
    echo ">>> Captured a segfault. Backtrace is in $OUT" | tee -a "$OUT"
    exit 0
  fi
  echo "gdb saw no segfault (the app may just run normally)." | tee -a "$OUT"
else
  echo "gdb not found; falling back to core dump." | tee -a "$OUT"
  ulimit -c unlimited
  "$BIN" --gui "$@" ; rc=$?
  echo "exit: $rc" | tee -a "$OUT"
  if [ "$rc" = 139 ]; then
    echo ">>> Segfault (139). Looking for a core..." | tee -a "$OUT"
    ls -la "$CORE_DIR" /tmp/core* ./*core* 2>/dev/null | tee -a "$OUT"
    core=$(ls -t "$CORE_DIR"/core* /tmp/core* ./*core* 2>/dev/null | head -1)
    if [ -n "${core:-}" ] && command -v gdb >/dev/null 2>&1; then
      gdb -batch -ex 'bt 60' "$BIN" "$core" 2>&1 | tee -a "$OUT"
    fi
  fi
  exit "$rc"
fi
