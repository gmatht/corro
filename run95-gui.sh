#!/usr/bin/env bash
# Build corro's nwg GUI for Windows 95 and run it inside Windows 95.
#
#   ./run95-gui.sh                full cycle (headless, VNC :6 = host port 5906)
#   ./run95-gui.sh --show         visible QEMU window (needs a working X display)
#   ./run95-gui.sh --no-build     reuse dist/gcorro.exe
#   ./run95-gui.sh --no-shutdown  leave the VM running afterwards
#   ./run95-gui.sh --refresh-base rebuild the clean base from the golden
#                               master: stage the exe, run fsck, so the next
#                               boot goes straight to the desktop
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
# Always boots a CLEAN snapshot.
# -------------------------------
# Win95 decides whether to run ScanDisk from a dirty bit the OS sets in the
# FAT and clears on an orderly unmount. A VM that is killed, or whose Explorer
# is still running, leaves that bit set - so every later boot spends ten
# minutes in ScanDisk on an emulated 486 and then stops on a lost-cluster
# dialog that has to be answered by hand. That is pure overhead, and it made
# runs look broken when nothing was wrong.
#
# So the images are split in two:
#
#   BASE    a qcow2 that has been fsck'd and already has the exe, sheet and
#           unicows staged. Never booted; only written offline, by the staging
#           step and by `--refresh-base`.
#   SESSION a throwaway overlay on BASE, recreated every run. Everything the
#           guest does - including dirtying its own filesystem - lands here
#           and is discarded with it.
#
# Every boot therefore starts from a filesystem the OS believes is clean.
# `--refresh-base` rebuilds BASE from the golden master when the golden has
# itself drifted; that is the only step that ever writes it.
set -euo pipefail
cd "$(dirname "$0")"

GOLDEN=${WIN95_DISK:-/root/vm/win95-flat.qcow2}
# BASE is derived from the golden's name so the two cannot be mistaken for
# one another on disk.
BASE=${WIN95_BASE:-${GOLDEN%.qcow2}-gui-base.qcow2}
SESSION_DISK=${WIN95_SESSION:-${GOLDEN%.qcow2}-gui-session.qcow2}
QEMU_PIDFILE=/tmp/qemu-win95-gui.pid
SHEET=${WIN95_SHEET:-subtotal.corro}
UNICOWS=${UNICOWS_DLL:-/opt/wine-stable/lib/wine/i386-windows/unicows.dll}
QMP_PORT=${QMP_PORT:-4446}
VNC_DISPLAY=${VNC_DISPLAY:-:6}          # = TCP port 5906
MODE=headless
NO_BUILD=0
NO_SHUTDOWN=0
REFRESH_BASE=0
for a in "$@"; do
  case "$a" in
    --show) MODE=show ;;
    --no-build) NO_BUILD=1 ;;
    --no-shutdown) NO_SHUTDOWN=1 ;;
    --refresh-base) REFRESH_BASE=1 ;;
    *) echo "unknown option: $a" >&2; exit 2 ;;
  esac
done

# Clear a ScanDisk run, if one is on screen.
#
# The dialog that needs answering is "ScanDisk found N bytes of data ... To look
# at the data, choose Save". Its default is Save, which writes a .CHK file into
# the image root on every boot; Right moves focus to "Don't Fix It", which is
# what we want - the bytes are just lost clusters from however the previous VM
# was killed, and the filesystem is about to be fsck'd offline anyway. The
# follow-up "There was an error running SCANDISK.EXE ... Press any key"
# message just needs a keypress.
#
# Bounded and screen-driven, so it is a no-op when there is no ScanDisk to
# clear - which is the normal case now that BASE is fsck'd.
dismiss_scandisk() {
  local i
  for i in $(seq 1 24); do
    QMP_PORT=$QMP_PORT python3 scripts/qmp.py shot /tmp/_w95_scan.png >/dev/null 2>&1 || return 0
    # `command grep` with long options: the `grep` on PATH is an rg wrapper
    # here and it rejects bundled short flags like -qi.
    screen=$(tesseract /tmp/_w95_scan.png - --psm 6 2>/dev/null)
    if ! printf '%s' "$screen" | command grep --ignore-case --quiet scandisk; then
      return 0
    fi
    # Only answer the prompt once it is actually there; the scan text is on
    # screen long before the buttons are.
    if printf '%s' "$screen" | command grep --ignore-case --quiet "choose Save"; then
      echo "   ScanDisk prompt: choosing Don't Fix It"
      QMP_PORT=$QMP_PORT python3 scripts/qmp.py key right
      sleep 1
      QMP_PORT=$QMP_PORT python3 scripts/qmp.py key ret
      sleep 20
    fi
    # "Press any key to continue starting Windows..."
    QMP_PORT=$QMP_PORT python3 scripts/qmp.py key ret
    sleep 15
  done
}

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

# ------------------------------------------------------- 3. build the BASE
# Everything the guest needs is staged into BASE, never into the golden
# master. The golden is a shared, human-curated artifact (other operators run
# their own sessions off it); writing test binaries into it is how it drifts.
#
# BASE is rebuilt from the golden only when it is missing or when
# --refresh-base is given, so the common run does no image work at all - it
# just makes a throwaway overlay and boots.
echo "== 3/6 preparing the clean base"
if [ ! -f "$GOLDEN" ]; then
  echo "ERROR: $GOLDEN missing (the golden master image)" >&2; exit 1
fi
export LIBGUESTFS_BACKEND=direct
[ -f "$SHEET" ] || { echo "ERROR: sheet $SHEET not found" >&2; exit 1; }

if [ "$REFRESH_BASE" = 1 ] || [ ! -f "$BASE" ]; then
  if [ "$REFRESH_BASE" = 1 ]; then
    echo "   --refresh-base: rebuilding $BASE from $GOLDEN"
    rm -f "$BASE"
  else
    echo "   no base yet; creating $BASE from $GOLDEN"
  fi
  qemu-img convert -O qcow2 "$GOLDEN" "$BASE"

  # The workbook goes on the image too, so the Run line is just the exe name
  # and the app opens with content. Launching an empty grid would prove only
  # that it paints; the ReactOS harness got real verification by naming a
  # sheet, and this one should too.
  guestfish -a "$BASE" -m /dev/sda1 upload dist/gcorro.exe /GCORRO.EXE
  guestfish -a "$BASE" -m /dev/sda1 upload "$SHEET" /SHEET.CORRO

  # Verify the upload byte-for-byte. Uploads have silently failed before, and
  # a truncated exe on the guest is indistinguishable from a broken app.
  guestfish -a "$BASE" -m /dev/sda1 download /GCORRO.EXE /tmp/_gcorro_verify.exe
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
  # (no window, no evidence). Wine's copy is the usual source. This image had
  # no copy at all, which is a plausible reason an earlier attempt to launch
  # the GUI here produced nothing.
  have_unicows=0
  for p in /WINDOWS/SYSTEM/UNICOWS.DLL /WINDOWS/SYSTEM/unicows.dll \
           /Windows/System/unicows.dll /windows/system/unicows.dll; do
    if guestfish --ro -a "$BASE" -m /dev/sda1 exists "$p" >/dev/null 2>&1; then
      have_unicows=1
      break
    fi
  done
  if [ "$have_unicows" != 1 ]; then
    if [ -f "$UNICOWS" ]; then
      echo "   unicows.dll not on image; uploading $UNICOWS"
      guestfish -a "$BASE" -m /dev/sda1 mkdir-p /WINDOWS/SYSTEM
      guestfish -a "$BASE" -m /dev/sda1 upload "$UNICOWS" /WINDOWS/SYSTEM/UNICOWS.DLL
    else
      echo "   WARNING: unicows.dll not on image and $UNICOWS missing; the exe" >&2
      echo "            may fail to load (needs unicows.dll on Win9x/ME)" >&2
    fi
  else
    echo "   unicows.dll already on image"
  fi

  # Clear the filesystem's dirty bit BEFORE the first boot. The golden has
  # almost always been killed rather than shut down, so Windows would
  # otherwise run ScanDisk for ten minutes on an emulated 486 and then stop
  # on a lost-cluster prompt nobody is there to answer. fsck returns 0x1 when
  # it corrected something, which is the expected result here, not a failure.
  echo "   fsck to clear the dirty bit (skips ScanDisk on first boot)"
  # guestfish needs an explicit `run` before any appliance command, and the
  # here-doc form is the only one that takes it. `|| true` because fsck exits
  # 0x1 when it corrected something, which is the expected outcome here.
  guestfish -a "$BASE" <<'FSCK' || true
run
fsck fat /dev/sda1
FSCK
else
  echo "   reusing existing base $BASE (pass --refresh-base to rebuild)"
  # The exe may have been rebuilt since BASE was made, so refresh just the
  # binary when its md5 no longer matches. Cheap, and it keeps "no image work
  # on a normal run" from quietly becoming "always runs yesterday's binary".
  guestfish -a "$BASE" -m /dev/sda1 download /GCORRO.EXE /tmp/_gcorro_have.exe 2>/dev/null || true
  have_md5=$(md5sum /tmp/_gcorro_have.exe 2>/dev/null | cut -d' ' -f1)
  want_md5=$(md5sum dist/gcorro.exe | cut -d' ' -f1)
  rm -f /tmp/_gcorro_have.exe
  if [ "$have_md5" != "$want_md5" ]; then
    echo "   exe changed since the base was built; restaging"
    guestfish -a "$BASE" -m /dev/sda1 upload dist/gcorro.exe /GCORRO.EXE
    guestfish -a "$BASE" -m /dev/sda1 upload "$SHEET" /SHEET.CORRO
    guestfish -a "$BASE" <<'FSCK' || true
run
fsck fat /dev/sda1
FSCK
  fi
fi

# ------------------------------------------------------------- 4. boot VM
# Boot a FRESH throwaway overlay on BASE. Whatever happens inside the guest
# lands here and is discarded next run, so BASE stays fsck-clean no matter
# how violently this run ends.
echo "== 4/6 booting VM (this takes a few minutes)"
rm -f "$SESSION_DISK"
qemu-img create -f qcow2 -F qcow2 -b "$BASE" "$SESSION_DISK" >/dev/null
echo "   session overlay: $SESSION_DISK (backed by $BASE)"
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

# Win95 will still run ScanDisk if the guest dirties the volume, so handle
# it rather than wait it out: the scan itself is slow on an emulated 486, and
# it then stops on a lost-cluster prompt that has to be answered by hand.
# 'Right' moves focus off the default (Save) onto "Don't Fix It" - Save would
# write FILE0000.CHK into the image root on every run - and the dialogs that
# follow are cleared with a keypress each.
echo "   handling any ScanDisk / boot dialogs"
dismiss_scandisk
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
