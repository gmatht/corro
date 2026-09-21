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
#
# The window must be at least as large as the frame, or feh renders NOTHING:
# given a window smaller than the image it does not scale down, it paints a
# flat grey rectangle. That is what made the first version of this script look
# broken over VNC — it used a hardcoded 480x1013 window against a 1080x2280
# screencap, so the viewer showed a solid grey box (3 distinct colours) while
# the frame file on disk was perfectly correct. Verified: the same image gives
# 3 distinct colours in a 400x800 window and 37 in a 1080x2280 one.
#
# So measure the frame instead of guessing, and let the VNC viewer do any
# scaling the user wants (it can zoom; feh cannot).
FRAME_W=$(python3 -c "
import struct
with open('$FRAME','rb') as f:
    d = f.read(33)
print(struct.unpack('>I', d[16:20])[0] if d[:8] == b'\x89PNG\r\n\x1a\n' else 0)
" 2>/dev/null)
FRAME_H=$(python3 -c "
import struct
with open('$FRAME','rb') as f:
    d = f.read(33)
print(struct.unpack('>I', d[20:24])[0] if d[:8] == b'\x89PNG\r\n\x1a\n' else 0)
" 2>/dev/null)
[ "${FRAME_W:-0}" -gt 0 ] 2>/dev/null || FRAME_W=1080
[ "${FRAME_H:-0}" -gt 0 ] 2>/dev/null || FRAME_H=2280
echo "feh window: ${FRAME_W}x${FRAME_H} (from the frame itself)" >>/tmp/live_loop.log
setsid feh --reload 1 --geometry "${FRAME_W}x${FRAME_H}" --title "corro on Android" "$FRAME" \
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
