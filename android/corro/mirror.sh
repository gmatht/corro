#!/usr/bin/env bash
# Live view of the Android emulator on X display :99, without scrcpy.
#
#   ./mirror.sh              # start mirroring (backgrounds nothing; runs the loop)
#   CORRO_FRAME=/tmp/f.png ./mirror.sh
#
# Why not scrcpy: the emulator's MediaCodec encoder dies after a while
# ("Encoding error: IllegalStateException"). The scrcpy *client* stays alive
# but stops receiving frames, so the VNC viewer silently shows a stale frame
# forever - it looks like the app froze when really only the mirror did.
# This loop uses `adb exec-out screencap` (no device encoder, no scrcpy) and
# writes each frame into an X window via feh, so the view stays live as long
# as the loop runs.
#
# The VNC server (:99, port 5999) mirrors this display, so a VNC viewer sees
# the same live window. Default to :99 unconditionally: an inherited stale
# DISPLAY (e.g. :0) would draw the window onto a display nobody is watching.
set -uo pipefail

ANDROID_SDK=${ANDROID_SDK:-$HOME/android-sdk}
PATH="$ANDROID_SDK/platform-tools:$PATH"
export PATH
export DISPLAY=${CORRO_DISPLAY:-:99}

FRAME=${CORRO_FRAME:-/tmp/live_frame.png}

: > /tmp/live_loop.log

# One long-lived viewer process fed by repeated file writes.
# Match exact process names only: a bare `pkill -f scrcpy` also matches this
# script's own command line (and the shell running it), killing the loop.
pkill -x scrcpy >/dev/null 2>&1
pkill -f "feh --reload" >/dev/null 2>&1
sleep 2

adb shell input keyevent 224 >/dev/null 2>&1
adb shell am start -n com.corro/.MainActivity >/dev/null 2>&1
sleep 4

# feh reloads the image whenever the file changes (-R), giving a live window.
setsid feh --reload 1 --geometry 480x1013 --title "corro on Android" "$FRAME" \
  </dev/null >>/tmp/live_loop.log 2>&1 &
sleep 3

n=0
while true; do
  # Write to a temp file, then rename: a partially-written frame would make
  # feh (and this loop) read a truncated PNG. Rename is atomic on the same fs.
  adb exec-out screencap -p > "$FRAME.tmp" 2>/dev/null
  [ -s "$FRAME.tmp" ] && mv -f "$FRAME.tmp" "$FRAME"
  n=$((n+1))
  echo "frame $n $(date +%T)" >> /tmp/live_loop.log
  sleep 1
done
