#!/usr/bin/env bash
# Build corro for Windows 95 and run it inside Windows 95 under QEMU.
#
#   ./run95.sh                # headless (VNC :5 = host port 5905), full cycle
#   ./run95.sh --show         # open a visible QEMU window (needs a working X display)
#   ./run95.sh --release      # build optimized
#   ./run95.sh --no-build     # reuse the existing dist/corro-win95.exe
#   ./run95.sh --no-shutdown  # leave the VM running afterwards
#
# What it does:
#   1. ./build95.sh (always, unless --no-build)
#   2. stops any running Win95 QEMU (guestfish needs the disk exclusively)
#   3. uploads dist/corro-win95.exe to C:\CORRO.EXE, verifies md5,
#      and makes sure unicows.dll is on the image (needed on Win9x/ME)
#   4. boots the VM, waits until the desktop has settled
#   5. launches C:\CORRO.EXE via Start > Run (opens a DOS-box console window;
#      corro is a console TUI)
#   6. waits for the screen to change, then screenshots:
#        dist/qemu-boot.png     desktop before launch
#        dist/qemu-window.png   after launching corro
#
# While it runs you can watch the VM:
#   - --show mode: the QEMU window itself
#   - headless:    vncviewer localhost:5905
#
# Env overrides:
#   WIN95_DISK=/root/vm/win95-flat.qcow2   golden master (uploads target it;
#                                        NEVER booted directly, so it stays clean)
#   WIN95_SESSION=...                     per-run overlay booted by QEMU
#                                        (default: <DISK>-session.qcow2; recreated
#                                        every run, so kills/dirty shutdowns only
#                                        ever trash the overlay, never the golden)
#   WIN95_SOURCE=/root/vm/Win95.vmdk       pristine source image (never written)
#   UNICOWS_DLL=/opt/wine-stable/lib/wine/i386-windows/unicows.dll
#   QMP_PORT=4445
set -euo pipefail
cd "$(dirname "$0")"

DISK=${WIN95_DISK:-/root/vm/win95-flat.qcow2}
SESSION_DISK=${WIN95_SESSION:-${DISK%.qcow2}-session.qcow2}
QEMU_PIDFILE=/tmp/qemu-win95.pid
# Prints the pidfile PID, but ONLY if that process is actually our QEMU
# (matched by binary + our QMP port). A stale pidfile whose PID got
# recycled by an innocent process must never become a kill target.
our_qemu_pid() {
  [ -f "$QEMU_PIDFILE" ] || return 0
  local p; p=$(cat "$QEMU_PIDFILE")
  if [ -n "$p" ] && tr '\0' ' ' < "/proc/$p/cmdline" 2>/dev/null | \
      grep -q "qemu-system-i386.*$QMP_PORT"; then
    echo "$p"
  fi
}
SOURCE=${WIN95_SOURCE:-/root/vm/Win95.vmdk}
UNICOWS=${UNICOWS_DLL:-/opt/wine-stable/lib/wine/i386-windows/unicows.dll}
QMP_PORT=${QMP_PORT:-4445}
VNC_DISPLAY=:5                       # = TCP port 5905
MODE=headless
BUILD_ARGS=()
NO_BUILD=0
NO_SHUTDOWN=0
for a in "$@"; do
  case "$a" in
    --show) MODE=show ;;
    --release) BUILD_ARGS+=(--release) ;;
    --no-build) NO_BUILD=1 ;;
    --no-shutdown) NO_SHUTDOWN=1 ;;
    *) echo "unknown option: $a" >&2; exit 2 ;;
  esac
done

# ---------------------------------------------------------------- 1. build
if [ "$NO_BUILD" = 1 ]; then
  [ -f dist/corro-win95.exe ] || { echo "ERROR: --no-build but dist/corro-win95.exe missing" >&2; exit 1; }
  echo "== 1/6 skipping build (--no-build)"
else
  echo "== 1/6 building"
  ./build95.sh "${BUILD_ARGS[@]+"${BUILD_ARGS[@]}"}"
fi

# ------------------------------------------------------- 2. stop running VM
# Scoped to OUR VM via the pidfile: a blanket pkill would murder other
# operators' VMs sharing this host (their QMP/VNC ports differ, ours are
# fixed below, so only our own pid is ever a candidate).
echo "== 2/6 stopping our running VM (if any)"
pids=$(our_qemu_pid)
if [ -n "$pids" ]; then
  kill $pids 2>/dev/null || true
  sleep 3
  kill -9 $pids 2>/dev/null || true
  sleep 1
fi
rm -f "$QEMU_PIDFILE"

# ------------------------------------------------------------ 3. upload exe
echo "== 3/6 uploading to $DISK"
if [ ! -f "$DISK" ]; then
  echo "   scratch disk missing; converting from $SOURCE"
  qemu-img convert -O qcow2 "$SOURCE" "$DISK"
fi
export LIBGUESTFS_BACKEND=direct
guestfish -a "$DISK" -m /dev/sda1 upload dist/corro-win95.exe /CORRO.EXE
# verify the upload byte-for-byte (uploads have silently failed before)
guestfish -a "$DISK" -m /dev/sda1 download /CORRO.EXE /tmp/_corro_verify.exe
disk_md5=$(md5sum /tmp/_corro_verify.exe | cut -d' ' -f1)
want_md5=$(md5sum dist/corro-win95.exe | cut -d' ' -f1)
if [ "$disk_md5" != "$want_md5" ]; then
  echo "ERROR: upload md5 mismatch ($disk_md5 != $want_md5)" >&2
  exit 1
fi
rm -f /tmp/_corro_verify.exe
echo "   upload verified (md5 $want_md5)"

# make sure unicows.dll is on the image (required on Win9x/ME; the rust9x
# targets import the unicode layer through it). Try the case variants libguestfs
# is likely to see, upload only if none exists.
have_unicows=0
for p in /WINDOWS/SYSTEM/UNICOWS.DLL /WINDOWS/SYSTEM/unicows.dll \
         /Windows/System/unicows.dll /windows/system/unicows.dll; do
  if guestfish --ro -a "$DISK" -m /dev/sda1 exists "$p" >/dev/null 2>&1; then
    have_unicows=1
    break
  fi
done
if [ "$have_unicows" != 1 ]; then
  if [ -f "$UNICOWS" ]; then
    echo "   unicows.dll not found on image; uploading $UNICOWS"
    guestfish -a "$DISK" -m /dev/sda1 mkdir-p /WINDOWS/SYSTEM
    guestfish -a "$DISK" -m /dev/sda1 upload "$UNICOWS" /WINDOWS/SYSTEM/UNICOWS.DLL
  else
    echo "   WARNING: unicows.dll not on image and $UNICOWS missing; the exe" >&2
    echo "            may fail to load (needs unicows.dll on Win9x/ME)" >&2
  fi
else
  echo "   unicows.dll already on image"
fi

# ------------------------------------------------------------- 4. boot VM
# Boot a FRESH overlay backed by the golden master: every run starts from
# the same scandisked snapshot, and whatever happens inside the guest
# (kills, dirty shutdowns, test debris) lands in the overlay, which is
# discarded next run. The golden is only ever touched offline by step 3.
echo "== 4/6 booting VM (this takes a few minutes)"
rm -f "$SESSION_DISK"
qemu-img create -f qcow2 -F qcow2 -b "$DISK" "$SESSION_DISK" >/dev/null
echo "   session overlay: $SESSION_DISK (backed by $DISK)"
ACCEL=kvm
[ -w /dev/kvm ] || ACCEL=tcg
DISP_ARGS=(-display none -vnc "$VNC_DISPLAY")
if [ "$MODE" = show ]; then
  DISP_ARGS=(-display gtk)
fi
setsid qemu-system-i386 -machine pc -cpu pentium -m 256 -accel "$ACCEL" \
  -drive file="$SESSION_DISK",if=ide,index=0,media=disk -vga cirrus \
  "${DISP_ARGS[@]}" \
  -qmp tcp:127.0.0.1:$QMP_PORT,server,nowait -rtc base=localtime -net none \
  > /tmp/qemu-win95.log 2>&1 &
echo $! > "$QEMU_PIDFILE"
sleep 5
if ! kill -0 "$(cat "$QEMU_PIDFILE")" 2>/dev/null; then
  echo "ERROR: QEMU died at startup; see /tmp/qemu-win95.log" >&2
  echo "       (another VM may already hold VNC $VNC_DISPLAY / QMP $QMP_PORT)" >&2
  exit 1
fi

python3 scripts/qmp.py wait-stable --min 140 --timeout 720 --verbose || \
  echo "   WARNING: desktop detection timed out; continuing anyway"
# dismiss any boot-time dialogs (e.g. Display Properties in Safe Mode)
python3 scripts/qmp.py key esc esc
sleep 12                       # let Explorer finish drawing
python3 scripts/qmp.py shot dist/qemu-boot.png

# --------------------------------------------------------- 5. launch the app
echo "== 5/6 launching C:\\CORRO.EXE"
python3 scripts/qmp.py run-dialog 'c:\corro.exe'

# --------------------------------------------------------- 6. verify + shoot
echo "== 6/6 waiting for the app window"
if python3 scripts/qmp.py wait-change dist/qemu-boot.png --timeout 60; then
  sleep 8                      # let the console window / TUI fully paint
  python3 scripts/qmp.py shot dist/qemu-window.png
  echo
  echo "OK: corro was launched under Windows 95."
  echo "    screenshots: dist/qemu-boot.png (desktop), dist/qemu-window.png (after launch)"
else
  python3 scripts/qmp.py shot dist/qemu-window.png
  echo
  echo "WARNING: screen did not change after launch; eyeball dist/qemu-window.png"
fi

# ------------------------------------------------- optional: clean shutdown
# Courtesy only now (the overlay is discarded next run regardless): a halted
# guest leaves tidy logs and avoids a ScanDisk pass if anyone reboots THIS
# overlay by hand. Skipped with --no-shutdown.
if [ "$NO_SHUTDOWN" != 1 ]; then
  echo "== shutting the guest down cleanly (Start > Shut Down)"
  python3 scripts/qmp.py combo ctrl+esc
  sleep 4
  python3 scripts/qmp.py key u          # 'u' = Shut Down... hotkey
  sleep 4
  python3 scripts/qmp.py key ret        # confirm
  if python3 scripts/qmp.py wait-halt; then
    echo "   guest halted; stopping QEMU"
    pids=$(our_qemu_pid)
    [ -n "$pids" ] && kill $pids 2>/dev/null || true
    rm -f "$QEMU_PIDFILE"
  else
    echo "   WARNING: guest did not halt (corro may still be running and" >&2
    echo "            blocking shutdown); leaving QEMU up" >&2
  fi
else
  echo "VM keeps running headless; watch with: vncviewer localhost:5905"
  echo "QMP: python3 scripts/qmp.py shot out.png  (port $QMP_PORT)"
fi
