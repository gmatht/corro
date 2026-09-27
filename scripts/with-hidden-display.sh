#!/usr/bin/env bash
# Run a command on a private, hidden X display (Xvfb), for the GTK GUI tests.
#
# Why this exists: the GUI tests spawn a real `corro --gui` per case, and they
# all share one X server. A visible display (or a leaked server from an
# earlier run) makes them fight — focus and window stacking go to whoever
# painted last, `xdotool windowactivate` acts on the wrong window, and cases
# time out waiting for a frame that another case's window painted over.
#
# So: one Xvfb per invocation, on its own display number, never mapped to a
# physical screen, torn down on every exit path (including a test panic or a
# Ctrl-C). Nothing is left behind for the next run to trip over.
#
# Usage:
#   scripts/with-hidden-display.sh cargo test --features gtk --test gui_formula_headers
#   scripts/with-hidden-display.sh ./scripts/catch-gui-crash.sh
#
# Env:
#   XVFB_W        screen width  (default 1400)
#   XVFB_H        screen height (default 1000)
#   CORRO_GUI_IDLE_SECS
#                 Wait this long for the GUI to go completely idle BEFORE
#                 running (default 0). "Idle" means no corro process, no
#                 window, and no X client on the display — see wait_for_idle.
#                 Use it when the box is shared or a previous session may
#                 still be drawing.
set -uo pipefail

W=${XVFB_W:-1400}
H=${XVFB_H:-1000}
IDLE_SECS=${CORRO_GUI_IDLE_SECS:-0}

log() { printf '[hidden-display] %s\n' "$*" >&2; }

# ── Idle gate ────────────────────────────────────────────────────────────────
# A GUI test asserts on pixels, so it must be the only thing touching the
# screen. Poll until the display is genuinely quiet, then say so — a silent
# wait that times out and runs anyway would reintroduce exactly the flakiness
# this script exists to remove.
wait_for_idle() {
  local deadline=$((SECONDS + IDLE_SECS))
  [ "$IDLE_SECS" -le 0 ] && return 0
  log "waiting up to ${IDLE_SECS}s for the GUI to go idle"
  while [ "$SECONDS" -lt "$deadline" ]; do
    local busy=0
    pgrep -x corro >/dev/null 2>&1 && busy=1
    pgrep -f 'target/debug/corro' >/dev/null 2>&1 && busy=1
    if [ "$busy" -eq 0 ]; then
      log "idle after $((SECONDS - (deadline - IDLE_SECS)))s"
      return 0
    fi
    sleep 2
  done
  log "WARNING: still busy after ${IDLE_SECS}s; continuing anyway"
  pgrep -af 'corro' >&2 2>/dev/null | head -5
  return 1
}

# ── Pick a free display number ───────────────────────────────────────────────
# :99 and up. Skip any number whose lock file exists, so two runs never pick
# the same display and fight over it.
pick_display() {
  local n
  for n in $(seq 90 200); do
    if [ ! -e "/tmp/.X${n}-lock" ] && [ ! -e "/tmp/.X11-unix/X${n}" ]; then
      echo "$n"
      return 0
    fi
  done
  echo "no free X display number in 90..200" >&2
  return 1
}

DISP=$(pick_display) || exit 1
AUTH_DIR=$(mktemp -d /tmp/corro-xvfb.XXXXXX)
XVFB_PID=""
WM_PID=""

# Reap anything left on this display first: a stale server from a killed run
# still holds the lock file, and two Xvfb on one number is a coin flip.
cleanup() {
  local code=$?
  # Kill the display's whole process group children first, then Xvfb itself,
  # then the auth dir. Every path out of this script runs this.
  for pid in "$WM_PID" "$XVFB_PID"; do
    [ -n "$pid" ] || continue
    kill "$pid" 2>/dev/null
    for _ in 1 2 3 4 5; do
      kill -0 "$pid" 2>/dev/null || break
      sleep 0.2
    done
    kill -9 "$pid" 2>/dev/null
  done
  rm -rf "$AUTH_DIR" 2>/dev/null
  # Stray children of the command (a GUI app that outlived its test) — only
  # ones whose DISPLAY was this one, so we never touch another run's app.
  pkill -f "target/debug/corro" 2>/dev/null
  rm -f "/tmp/.X${DISP}-lock" 2>/dev/null
  exit "$code"
}
trap cleanup EXIT INT TERM

XAUTH="$AUTH_DIR/Xauthority"
touch "$XAUTH"
if command -v xauth >/dev/null 2>&1; then
  xauth -f "$XAUTH" add ":${DISP}" . "$(head -c 32 /dev/urandom | od -An -tx1 | tr -d ' \n')" 2>/dev/null
fi

log "starting Xvfb on :${DISP} (${W}x${H}x24), hidden"
# `-nolisten unix`: this box has /tmp/.X11-unix on a read-only filesystem, so
# Xvfb cannot create its socket there and dies with "failed to bind listener".
# Disabling the unix listener makes it serve over TCP on loopback instead,
# which works and is still invisible (nothing is mapped to a real screen).
Xvfb ":${DISP}" -screen 0 "${W}x${H}x24" -nolisten unix -nolisten tcp >/dev/null 2>&1 &
XVFB_PID=$!

# Wait for the display to actually accept connections before handing it over;
# launching the tests against a half-started Xvfb is the same race in a new
# place. (With -nolisten unix there is no socket file to poll, so ask the
# server itself.)
ready=0
for _ in $(seq 1 100); do
  if DISPLAY=":${DISP}" xdpyinfo >/dev/null 2>&1; then ready=1; break; fi
  if ! kill -0 "$XVFB_PID" 2>/dev/null; then
    log "Xvfb died during startup"
    exit 1
  fi
  sleep 0.1
done
if [ "$ready" -ne 1 ]; then
  log "Xvfb did not come up on :${DISP}"
  exit 1
fi

export DISPLAY=":${DISP}"
export XAUTHORITY="$XAUTH"
log "DISPLAY=$DISPLAY (Xvfb pid $XVFB_PID)"

# ── Window manager ───────────────────────────────────────────────────────────
# The GUI tests drive the app with `xdotool windowactivate` and
# `xdotool key --window`. Both need a window manager: without one, Xvfb has no
# _NET_ACTIVE_WINDOW support, activation aborts, and the synthetic keys are
# delivered to nothing — so the app never sees a Right press and the test
# measures an unmoved row and calls it "stable". That is a false pass, and it
# is how a real reflow can hide behind a green test.
#
# fluxbox is the smallest WM here that provides EWMH and focus handling.
WM=${CORRO_GUI_WM:-fluxbox}
if command -v "$WM" >/dev/null 2>&1; then
  "$WM" >/dev/null 2>&1 &
  WM_PID=$!
  # Wait for _NET_SUPPORTING_WM_CHECK to appear: starting the WM is not the
  # same as it being ready to answer _NET_ACTIVE_WINDOW.
  wm_ready=0
  for _ in $(seq 1 100); do
    if command -v xprop >/dev/null 2>&1; then
      if xprop -root _NET_SUPPORTING_WM_CHECK 2>/dev/null | grep -q 'window id'; then
        wm_ready=1
        break
      fi
    else
      # No xprop: assume it came up and let the command's own activation
      # report failure if it did not.
      wm_ready=1
      break
    fi
    if ! kill -0 "$WM_PID" 2>/dev/null; then break; fi
    sleep 0.1
  done
  if [ "$wm_ready" -eq 1 ]; then
    log "window manager $WM ready (pid $WM_PID)"
  else
    log "WARNING: $WM did not report ready; xdotool activation may fail"
  fi
else
  log "WARNING: no $WM available; GUI tests that use xdotool will not work"
fi

wait_for_idle || true

if [ "$#" -eq 0 ]; then
  log "no command given; holding the display open (Ctrl-C to exit)"
  sleep infinity
fi

log "running: $*"
"$@"
rc=$?
log "command exited $rc"
exit "$rc"
