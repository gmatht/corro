corro on Android — live view
============================
A VNC server serves display :99, where scrcpy mirrors the Android emulator
showing corro running.

Connect:   <host-ip>:5999   (no password: -SecurityTypes None)
  host ip: 172.25.38.120  (container eth0; from the docker host, forward/map port 5999)
  display: :99  (1280x800)
  VNC listen: 0.0.0.0:5999

App:       com.corro/.MainActivity   (APK: /tmp/corro_apk_build/signed.apk)
Emulator:  AVD corro_avd (Android 13, x86_64, port 5554)
Logs:      adb logcat -s corro rswidgets
Screenshot: adb exec-out screencap -p > shot.png
Taps:      adb shell input tap X Y

Building
--------
    ./build_apk.sh [x86_64|arm64-v8a] [install]

The script builds the `corro_android` cdylib for the NDK target with
`--features generate-android-resources`, then packages the Java sources,
`app/src/main/res` and the `.so` with the raw toolchain (javac/d8/aapt2).

This crate is NOT in the corro workspace: it targets the Android NDK with
its own feature set, so build it from here (or via the script), not with a
workspace-wide `cargo build`. It only compiles for an Android target — the
JNI exports and `corro::gui::android_backend` are android-gated.

Android resources
-----------------
`build.rs` runs `rswidgets::android_generator` (the source is
`#[path]`-included from the rswidgets checkout, so there is no
`[build-dependencies]` entry) and writes the Material 3 theme, palette,
strings, adaptive launcher icon and manifest into this directory —
the root that contains `app/`. Generation is additive: every file this
project already ships is left alone, so it only fills gaps and never
fights hand edits. Delete a resource to have it recreated.

Off by default; enable it with the feature, or point the generator at
another project entirely:

    cargo ndk -t x86_64 build --features generate-android-resources
    RSWIDGETS_ANDROID_PROJECT=/path/to/other/project cargo ndk -t x86_64 build

See `rustxWidgets/docs/ANDROID_GUIDELINES.md` for what rswidgets generates
vs. what the app must supply (Material AAR, the Java shims, packaging).

Previewing the UI without an emulator
-------------------------------------
The widget tree this crate roots in the Activity is the same one the
desktop GUI builds. To look at it on a host (much faster than an APK
round-trip):

    cargo run --example android-ui --features gui      # from the repo root
