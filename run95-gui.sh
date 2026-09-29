#!/usr/bin/env bash
# Build corro's nwg GUI for Windows 95 and run it inside Windows 95.
#
#   ./run95-gui.sh                full cycle (headless, VNC :6 = host port 5906)
#   ./run95-gui.sh --show         visible QEMU window (needs a working X display)
#   ./run95-gui.sh --no-build     reuse dist/gcorro.exe
#   ./run95-gui.sh --no-shutdown  leave the VM running afterwards
#
# The sibling run95.sh does the same for the *console TUI* (corro-win95.exe,
# launched through Start > Run so it gets a DOS box). This one runs the GUI
# build, which is a different animal in three ways that matter:
#
#   * the entry point is WinMain through the VC6 GUI CRT, not main, so the
#     exe needs no console and Start > Run gives it none - which is what it
#     wants (see build95-gui.sh);
#   * the app takes its workbook from the command line, so the sheet is named
#     on the Run line rather than staged somewhere;
#   * the app is a GUI, so it is verified the way the ReactOS harness does it:
#     by driving it and looking at frames, not by spotting a console window.
#
# Both harnesses boot a FRESH per-run overlay backed by the golden master, so
# whatever happens in the guest (kills, dirty shutdowns, test debris) lands in
# the overlay and the golden master stays clean. The golden is only ever
# written offline, by the upload step.
set -euo pipefail
cd "$(dirname "$0")"

DISK=${WIN95_DISK:-/root/vm/win95-flat.qcow2}
SESSION_DISK=${WIN95_SESSION:-${DISK%.qcow2}-gui-session.qcow2}
QEMU_PIDFILE=/tmp/qemu-win95-gui.pid
SHEET=${WIN95_SHEET:-subtotal.corro}
UNICOWS=${UNICOWS_DLL:-/opt/wine-stable/lib/wine/i386-windows/unicows.dll}
QMP_PORT=${QMP_PORT:-4446}
VNC_DISPLAY=${VNC_DISPLAY:-:6}          # = TCP port 5906
MODE=headless
NO_BUILD=0
NO_SHUTDOWN=0
for a in "$@"; do
  case "$a" in
    --show) MODE=show ;;
    --no-build) NO_BUILD=1 ;;
    --no-shutdown) NO_SHUTDOWN=1 ;;
    *) echo "unknown option: $a" >&2; exit 2 ;;
  esac
done

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

# ---------------------------------------------------------------- 1. build
if [ "$NO_BUILD" = 1 ]; then
  [ -f dist/gcorro.exe ] || { echo "ERROR: --no-build but dist/gcorro.exe missing" >&2; exit 1; }
  echo "== 1/6 skipping build (--no-build)"
else
  echo "== 1/6 building (Win95 GUI)"
  ./build95-gui.sh
fi

# ------------------------------------------------------ 2. stop our own VM
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

# -------------------------------------------------------- 3. upload exe+sheet
echo "== 3/6 uploading to $DISK"
if [ ! -f "$DISK" ]; then
  echo "ERROR: $DISK missing (the golden master image)" >&2; exit 1
fi
export LIBGUESTFS_BACKEND=direct

# The workbook goes on the image too, so the Run line is just the exe name
# and the app opens with content. Launching an empty grid would prove only
# that it paints; the ReactOS harness got real verification by naming a
# sheet, and this one should too.
[ -f "$SHEET" ] || { echo "ERROR: sheet $SHEET not found" >&2; exit 1; }
guestfish -a "$DISK" -m /dev/sda1 upload dist/gcorro.exe /GCORRO.EXE
guestfish -a "$DISK" -m /dev/sda1 upload "$SHEET" /SHEET.CORRO

# Verify the upload byte-for-byte. Uploads have silently failed before, and
# a truncated exe on the guest is indistinguishable from a broken app.
guestfish -a "$DISK" -m /dev/sda1 download /GCORRO.EXE /tmp/_gcorro_verify.exe
disk_md5=$(md5sum /tmp/_gcorro_verify.exe | cut -d' ' -f1)
want_md5=$(md5sum dist/gcorro.exe | cut -d' ' -f1)
if [ "$disk_md5" != "$want_md5" ]; then
  echo "ERROR: upload md5 mismatch ($disk_md5 != $want_md5)" >&2
  exit 1
fi
rm -f /tmp/_gcorro_verify.exe
echo "   upload verified (md5 $want_md5), sheet $SHEET staged"

# unicows.dll is required on Win9x/ME: the rust9x targets import the Unicode
# layer through it, and without it the loader fails before any window opens
# (no window, no evidence). Wine's copy is the usual source.
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
    echo "   unicows.dll not on image; uploading $UNICOWS"
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
# Boot a FRESH overlay backed by the golden master: every run starts from the
# same snapshot, and whatever happens inside the guest lands in the overlay,
# which is discarded next run. The golden is only ever written by step 3.
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
  > /tmp/qemu-win95-gui.log 2>&1 &
echo $! > "$QEMU_PIDFILE"
sleep 5
if ! kill -0 "$(cat "$QEMU_PIDFILE")" 2>/dev/null; then
  echo "ERROR: QEMU died at startup; see /tmp/qemu-win95-gui.log" >&2
  echo "       (another VM may already hold VNC $VNC_DISPLAY / QMP $QMP_PORT)" >&2
  exit 1
fi

QMP_PORT=$QMP_PORT python3 scripts/qmp.py wait-stable --min 140 --timeout 720 --verbose || \
  echo "   WARNING: desktop detection timed out; continuing anyway"
# Dismiss any boot-time dialogs (e.g. Display Properties in Safe Mode)
QMP_PORT=$QMP_PORT python3 scripts/qmp.py key esc esc
sleep 12                       # let Explorer finish drawing
QMP_PORT=$QMP_PORT python3 scripts/qmp.py shot dist/qemu95-boot.png

# --------------------------------------------------------- 5. launch the app
# Start > Run, same route run95.sh uses. The GUI build wants no console, and
# Start > Run supplies none, so this also exercises the WinMain path.
echo "== 5/6 launching C:\\GCORRO.EXE C:\\SHEET.CORRO"
QMP_PORT=$QMP_PORT python3 scripts/qmp.py run-dialog 'c:\gcorro.exe c:\sheet.corro'

# --------------------------------------------------------- 6. verify + shoot
echo "== 6/6 waiting for the app window"
if QMP_PORT=$QMP_PORT python3 scripts/qmp.py wait-change dist/qemu95-boot.png --timeout 90; then
  sleep 10                     # let the GUI settle and paint its first frame
  QMP_PORT=$QMP_PORT python3 scripts/qmp.py shot dist/qemu95-gui.png
  echo
  echo "OK: corro's GUI was launched under Windows 95."
  echo "    screenshots: dist/qemu95-boot.png (desktop), dist/qemu95-gui.png (app)"
  echo
  echo "The app is still running. Drive it, then re-shoot:"
  echo "  QMP_PORT=$QMP_PORT python3 scripts/qmp.py type '42'"
  echo "  QMP_PORT=$QMP_PORT python3 scripts/qmp.py key ret right down"
  echo "  QMP_PORT=$QMP_PORT python3 scripts/qmp.py shot dist/qemu95-gui-2.png"
else
  QMP_PORT=$QMP_PORT python3 scripts/qmp.py shot dist/qemu95-gui.png
  echo
  echo "WARNING: screen did not change after launch; eyeball dist/qemu95-gui.png" >&2
  echo "         (a blank or unchanged screen usually means the loader failed -" >&2
  echo "          check unicows.dll is on the image and that the exe uploaded)" >&2
fi

# ------------------------------------------------- optional: clean shutdown
# Courtesy only: the overlay is discarded next run regardless, but a halted
# guest leaves tidy logs and avoids a ScanDisk pass if anyone reboots THIS
# overlay by hand. Skipped with --no-shutdown.
if [ "$NO_SHUTDOWN" != 1 ]; then
  echo "== shutting the guest down cleanly (Start > Shut Down)"
  QMP_PORT=$QMP_PORT python3 scripts/qmp.py combo ctrl+esc
  sleep 4
  QMP_PORT=$QMP_PORT python3 scripts/qmp.py key u          # 'u' = Shut Down...
  sleep 4
  QMP_PORT=$QMP_PORT python3 scripts/qmp.py key ret        # confirm
  if QMP_PORT=$QMP_PORT python3 scripts/qmp.py wait-halt; then
    echo "   guest halted; stopping QEMU"
    pids=$(our_qemu_pid)
    [ -n "$pids" ] && kill $pids 2>/dev/null || true
    rm -f "$QEMU_PIDFILE"
  else
    echo "   WARNING: guest did not halt (the app may still be running and" >&2
    echo "            blocking shutdown); leaving QEMU up" >&2
  fi
else
  echo "VM keeps running headless; watch with: vncviewer localhost:${VNC_DISPLAY#:}"
  echo "QMP: QMP_PORT=$QMP_PORT python3 scripts/qmp.py shot out.png"
fi
