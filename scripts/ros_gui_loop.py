#!/usr/bin/env python3
"""Boot the custom ReactOS corro LiveCD and launch the GUI, unattended.

This is the bring-up harness for the rust9x/Win95 GUI on ReactOS. It builds a
custom LiveCD from the stock ReactOS 0.4.16 image with:

  * gcorro.exe (and its Win95 unicode layer, unicows.dll) dropped into
    <live>/reactos/, so the app sits next to the rest of the live system;
  * a `zz_corro.bat` in the *All Users* and *Default User* profile Startup
    folders. Explorer runs those on login, so the app launches by itself with
    no operator input - the ReactOS LiveCD has no working autorun.inf
    (SHELL_CanInvokeAutoRunOnDrive is `#if 0`) and its Run/Startup handling
    works only for the Startup folders;
  * the batch launches the app through %SystemRoot% (the LiveCD mounts itself
    on a *variable* drive letter, `X:` here - not C:), and points
    CORRO_WIN95_LOG at a FAT12 floppy so the app's diagnostic log lands on a
    volume the host can read back.

It then boots the ISO under QEMU (headless, KVM when available), drives the
interactive LiveCD boot to the desktop (the FreeLDR menu, the CD "press any
key" gate and the "Select your language" splash each block until a key), and
snapshots the screen. Finally it copies the app log off the floppy.

Requires: qemu-system-i386, xorriso OR genisoimage (for -udf), hivexregedit,
tesseract, PIL. Run ros_harness.sh (a thin wrapper) for the common case.
"""
import glob
import os
import shutil
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
CORRO = os.path.abspath(os.path.join(HERE, ".."))
ROS_DIR = os.environ.get("ROS_DIR", "/root/vm/reactos")
STOCK_ISO = os.path.join(ROS_DIR, "ReactOS-0.4.16-i386.iso")
TREE = os.path.join(ROS_DIR, "build-tree")
# Per-run ISO, like the floppy: a shared path means the next run's
# genisoimage overwrites the image the running QEMU still has open, which
# shows up as an unexplained SIGKILL (rc=-9) rather than anything obvious.
OUT_ISO = os.path.join(ROS_DIR, f"corro-live-{os.getpid()}.iso")
# Per-run floppy: see make_floppy() for why it must not be shared.
FLOPPY = os.path.join(
    ROS_DIR, f"corro-log-{os.getpid()}.img")
UNICOWS = "/opt/wine-stable/lib/wine/i386-windows/unicows.dll"

# Fixed ports were a flakiness source: another session (or a leftover VM from a
# killed run) holding the port makes QEMU exit at startup, and the failure
# surfaces only as a QMP connection refusal with the real reason buried in
# qemu's own log. Ask the OS for a free port instead, and let QMP_PORT override.
def _free_port():
    import socket
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


QMP_PORT = int(os.environ.get("QMP_PORT") or _free_port())
VNC_DISPLAY = int(os.environ.get("VNC_DISPLAY") or (9000 + (QMP_PORT % 900)))
OUT = os.environ.get("ROS_OUT", "/tmp/rosrun")


def log(msg):
    print(msg, flush=True)


# --- ISO build -------------------------------------------------------------

def extract_stock_tree():
    """Extract the stock ISO preserving on-disc name CASE.

    A plain `mount -o loop` + cp is not good enough: the kernel ISO9660
    driver folds names to lower case, and the LiveCD's `Profiles` directory
    must keep its capital P or the live boot fails to find the profile and
    hangs at "Loading system hive". xorriso restores the real names.
    """
    if os.path.isdir(TREE):
        shutil.rmtree(TREE)
    os.makedirs(TREE)
    subprocess.run(
        ["xorriso", "-osirrox", "on", "-indev", STOCK_ISO,
         "-extract", "/", TREE],
        check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    subprocess.run(["chmod", "-R", "u+w", TREE], check=True)
    if not os.path.isdir(os.path.join(TREE, "Profiles")):
        raise SystemExit("extract did not preserve the 'Profiles' name case")


def make_genisoimage():
    for cand in ("/snap/lxd/40585/bin/genisoimage", "genisoimage", "mkisofs"):
        p = shutil.which(cand) or (cand if os.path.isfile(cand) else None)
        if p:
            return p
    raise SystemExit("need genisoimage (supports -udf); install it")


def write_startup_bat():
    """The batch Explorer runs from the profile Startup folders.

    %SystemRoot% is used on purpose: the LiveCD mounts itself on whatever
    drive letter it picks (X: in this image), not C:. The log goes to the
    floppy (A:) so the host can read it back.
    """
    bat = (
        "@echo off\r\n"
        "set CORRO_WIN95_LOG=A:\\gcorro.log\r\n"
        "cd /d \"%SystemRoot%\"\r\n"
        "\"%SystemRoot%\\gcorro.exe\"\r\n"
    )
    startup_dirs = [
        os.path.join(TREE, "Profiles", "All Users", "Start Menu", "Programs", "StartUp"),
        os.path.join(TREE, "Profiles", "Default User", "Start Menu", "Programs", "StartUp"),
    ]
    for d in startup_dirs:
        os.makedirs(d, exist_ok=True)
        with open(os.path.join(d, "zz_corro.bat"), "w", newline="") as f:
            f.write(bat)
    # a copy next to the exe, for reference / manual launch
    with open(os.path.join(TREE, "reactos", "run.bat"), "w", newline="") as f:
        f.write(bat)


def build_iso(exe):
    log("== building the custom LiveCD")
    extract_stock_tree()
    log("== installing gcorro.exe + unicows.dll into <live>/reactos")
    shutil.copy(exe, os.path.join(TREE, "reactos", "gcorro.exe"))
    # gcorro.exe statically imports the Win95 unicode layer; the loader looks
    # next to the exe first. ReactOS ships no unicows.dll, so without this the
    # process never starts (no window, no log).
    shutil.copy(UNICOWS, os.path.join(TREE, "reactos", "unicows.dll"))
    write_startup_bat()

    genisoimage = make_genisoimage()
    sortfile = "/tmp/bootfiles.sort"
    with open(sortfile, "w") as f:
        f.write("boot.catalog 4\n")
        f.write(os.path.join(TREE, "loader", "isoboot.bin") + " 3\n")
    # -udf matters: the LiveCD freeldr reads the payload through the UDF/ISO9660
    # tree the stock image ships, and a plain ISO9660-only rebuild stalls at
    # "Loading system hive". Match the stock recipe (loader/remaster.sh).
    subprocess.run([
        genisoimage, "-quiet", "-o", OUT_ISO, "-iso-level", "4", "-udf",
        "-allow-limited-size",
        "-publisher", "ReactOS Project", "-preparer", "ReactOS Project",
        "-volid", "ReactOS", "-volset", "ReactOS",
        "-eltorito-boot", "loader/isoboot.bin", "-no-emul-boot",
        "-boot-load-size", "4",
        "-hide", "boot.catalog", "-sort", sortfile, "-no-cache-inodes", TREE,
    ], check=True)
    log(f"   built {OUT_ISO} ({os.path.getsize(OUT_ISO)} bytes)")


def make_floppy():
    """Create the log floppy.

    Recreated from scratch every run, and named per-run (see FLOPPY's
    definition). A shared path is a flakiness source: QEMU takes a write lock
    on the image for as long as it runs, so two concurrent runs - or a run
    after one was killed - make the next QEMU exit at startup with "Failed to
    get \"write\" lock", which used to surface only as a QMP connection
    refusal.
    """
    with open(FLOPPY, "wb") as f:
        f.truncate(1474560)
    subprocess.run(["mkfs.vfat", "-F", "12", FLOPPY],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


# --- QEMU ------------------------------------------------------------------

def start_qemu():
    accel = "kvm" if os.access("/dev/kvm", os.W_OK) else "tcg"
    logp = open("/tmp/qemu-corro-ros.log", "w")
    return subprocess.Popen([
        "qemu-system-i386", "-machine", "pc", "-cpu", "qemu32", "-m", "2048",
        "-accel", accel,
        "-drive", f"file={OUT_ISO},format=raw,media=cdrom,if=ide,index=2",
        "-drive", f"file={FLOPPY},format=raw,if=floppy",
        "-usb", "-device", "usb-tablet", "-vga", "cirrus",
        # VNC display is a separate pick: the QMP port and the VNC port are
        # different numbers and neither implies the other, so a fixed ":6"
        # collided independently of QMP_PORT.
        "-display", "none", "-vnc", f":{VNC_DISPLAY}",
        "-qmp", f"tcp:127.0.0.1:{QMP_PORT},server,nowait",
        "-rtc", "base=localtime", "-net", "none", "-boot", "once=d",
    ], stdout=logp, stderr=subprocess.STDOUT, start_new_session=True)


def main():
    exe = sys.argv[1] if len(sys.argv) > 1 else os.path.join(CORRO, "dist", "gcorro.exe")
    os.makedirs(OUT, exist_ok=True)
    build_iso(exe)
    make_floppy()
    proc = start_qemu()
    log(f"== qemu pid {proc.pid}")
    sys.path.insert(0, os.path.join(CORRO, "scripts"))
    import qmp
    try:
        q = qmp.Qmp(QMP_PORT)
    except OSError as e:
        # A QMP refusal on its own is close to useless: the real reason is in
        # qemu's log, and the common causes (image locked by another QEMU, port
        # taken, bad ISO) never reach this exception. Report the actual cause
        # instead of making the reader go and find it.
        if proc.poll() is not None:
            log(f"== qemu exited immediately (rc={proc.returncode})")
        else:
            log(f"== qemu still running but QMP on {QMP_PORT} refused ({e})")
        logf = "/tmp/qemu-corro-ros.log"
        if os.path.exists(logf):
            log(f"== qemu said:\n" + open(logf).read()[-2000:])
        raise SystemExit(1)
    t0 = time.time()
    last_enter = 0.0
    at_desktop = False
    while time.time() - t0 < 360:
        try:
            q.cmd("screendump", {"filename": f"{OUT}/_boot.ppm", "format": "ppm"})
        except Exception:
            time.sleep(3)
            continue
        from PIL import Image, ImageOps
        try:
            im = Image.open(f"{OUT}/_boot.ppm")
        except Exception:
            time.sleep(2)
            continue
        ImageOps.autocontrast(im.convert("L")).save(f"{OUT}/_boot.png")
        txt = subprocess.run(["tesseract", f"{OUT}/_boot.png", "-", "--psm", "6"],
                             capture_output=True, text=True).stdout.lower()
        # Desktop detection. These must be things that CANNOT appear on a boot
        # screen: the FreeLDR menu itself offers "Run ReactOS Live CD" and
        # "Install ReactOS", so matching on those called the boot menu the
        # desktop and the run then sat on a menu for the rest of its budget
        # while the app never started. (The Recycle Bin / My Computer /
        # Start Menu strings only exist once Explorer is up.)
        # The app's own window is ALSO a success condition. Once the Startup
        # batch launches it, the window covers the desktop icons, so the
        # desktop markers stop matching and the loop kept re-sending Enter at
        # an already-working VM until the budget ran out. Match the app's title
        # so a run ends as soon as the thing under test is on screen.
        app_up = ("corro" in txt) and ("edit" in txt or "insert" in txt)
        if any(k in txt for k in ("recycle bin", "my computer",
                                  "command prompt", "start menu")) or app_up:
            if not at_desktop:
                what = "app" if app_up else "desktop"
                log(f"== {what} reached at t={time.time()-t0:.0f}s")
                at_desktop = True
                q.cmd("screendump", {"filename": f"{OUT}/desktop.png", "format": "ppm"})
            # let the Startup batch launch the app, then snapshot it
            for k in range(6):
                time.sleep(15)
                q.cmd("screendump", {"filename": f"{OUT}/app-{k}.png", "format": "ppm"})
            break
        # Boot gates that block until a key is pressed: the FreeLDR menu, the
        # CD driver "press any key" gate, and the first-boot language splash.
        # "run reactos" is deliberately NOT here even though it appears on the
        # boot menu: Enter on that menu has already been sent to get here, and
        # re-matching the menu text just sends more Enters at a screen that is
        # past the gate. Unreadable screens are handled by the stall re-send
        # below, which does not depend on OCR at all.
        gate = any(k in txt for k in ("language", "press any key",
                                      "select the operating"))
        if gate and time.time() - last_enter > 7:
            q.cmd("send-key", {"keys": [{"type": "qcode", "data": "ret"}]})
            last_enter = time.time()
            log(f"   [t={time.time()-t0:.0f}s] Enter at boot gate")
        # The gates do not appear at predictable wall-clock times (KVM vs TCG,
        # host load, disk cache state all move them) and a single Enter can be
        # dropped by the guest, so a long stall gets a re-send whether or not
        # OCR recognised the screen. Without this the run can sit at a gate
        # whose text was misread until the 360s budget runs out.
        elif time.time() - t0 > 45 and time.time() - last_enter > 12:
            q.cmd("send-key", {"keys": [{"type": "qcode", "data": "ret"}]})
            last_enter = time.time()
            log(f"   [t={time.time()-t0:.0f}s] Enter (stall re-send)")
        time.sleep(3)
    q.close()
    # fetch the app log off the floppy
    mount = "/tmp/_ros_floppy"
    os.makedirs(mount, exist_ok=True)
    subprocess.run(["umount", mount], stderr=subprocess.DEVNULL)
    if subprocess.run(["mount", "-o", "loop", FLOPPY, mount],
                      stderr=subprocess.DEVNULL).returncode == 0:
        for name in ("gcorro.log", "gcorro.log.panic", "gcorro.stdout.log",
                     "trace.txt"):
            src = os.path.join(mount, name)
            if os.path.exists(src):
                dst = os.path.join(OUT, name)
                shutil.copy(src, dst)
                log(f"== fetched {name} -> {dst}")
        subprocess.run(["umount", mount], stderr=subprocess.DEVNULL)
    log("== done; see /tmp/rosrun/ for screenshots and the app log")
    # The floppy is per-run (see FLOPPY), so remove it here or a long-lived
    # working directory slowly fills with 1.4 MB images. Sweep by pattern as
    # well: a run that is killed part-way never reaches this line, and those
    # are exactly the ones that leave debris.
    try:
        os.unlink(FLOPPY)
    except OSError:
        pass
    for pat in ("corro-log-*.img", "corro-live-*.iso"):
      for stale in glob.glob(os.path.join(ROS_DIR, pat)):
        if os.path.getmtime(stale) < time.time() - 3600:
          try:
            os.unlink(stale)
          except OSError:
            pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
