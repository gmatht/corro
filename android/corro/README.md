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
