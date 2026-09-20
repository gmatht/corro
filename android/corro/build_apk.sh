#!/bin/bash
# Build the corro Android APK from scratch (mirrors
# rustxWidgets/examples/android/build_apk.sh for the widget demo).
# Usage: ./build_apk.sh [ABI] [install]
#   ABI defaults to x86_64 (emulator); use arm64-v8a for real devices.
set -euo pipefail

ANDROID_HOME="${ANDROID_HOME:-$HOME/android-sdk}"
ANDROID_NDK_HOME="${ANDROID_NDK_HOME:-$ANDROID_HOME/ndk/25.2.9519653}"
PATH="$ANDROID_HOME/cmdline-tools/latest/bin:$ANDROID_HOME/platform-tools:$ANDROID_HOME/build-tools/33.0.2:$ANDROID_HOME/emulator:$PATH"

ABI="${1:-x86_64}"
# cargo-ndk target triple per ABI
case "$ABI" in
  x86_64)    TARGET="x86_64";    TRIPLE="x86_64-linux-android";    LIBDIR="x86_64" ;;
  arm64-v8a) TARGET="aarch64";   TRIPLE="aarch64-linux-android";   LIBDIR="arm64-v8a" ;;
  *) echo "unknown ABI: $ABI (want x86_64 or arm64-v8a)" >&2; exit 1 ;;
esac

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CORRO_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
BUILD_DIR="/tmp/corro_apk_build"
rm -rf "$BUILD_DIR" && mkdir -p "$BUILD_DIR"/{classes,dex,staging}

# 1. Build Rust .so (corro cdylib + spreadsheet pipeline, gui feature).
# `generate-android-resources` runs rswidgets::android_generator from build.rs
# (additive: existing resources are never overwritten) so the APK's icon,
# theme, palette, strings and manifest stay reproducible.
echo "==> Building Rust .so for $TRIPLE..."
cd "$CORRO_ROOT/android/corro"
cargo ndk -t "$TARGET" build --release --features generate-android-resources 2>&1 | tail -1

# 2. Compile Java sources
echo "==> Compiling Java..."
javac -d "$BUILD_DIR/classes" -source 8 -target 8 \
  -cp "$ANDROID_HOME/platforms/android-33/android.jar" \
  "$SCRIPT_DIR/app/src/main/java/com/corro/"*.java 2>&1 | grep -v warning || true

# 3. Convert to DEX
echo "==> Converting to DEX..."
d8 --release --lib "$ANDROID_HOME/platforms/android-33/android.jar" \
  --output "$BUILD_DIR/dex" "$BUILD_DIR/classes/com/corro/"*.class 2>&1

# 4. Package resources
echo "==> Packaging resources..."
aapt2 compile --dir "$SCRIPT_DIR/app/src/main/res" -o "$BUILD_DIR/resources.zip" 2>&1 | grep -v "^INFO\|^$" || true

# 5. Link base APK
echo "==> Linking APK..."
aapt2 link -o "$BUILD_DIR/unsigned.apk" \
  -I "$ANDROID_HOME/platforms/android-33/android.jar" \
  --manifest "$SCRIPT_DIR/app/src/main/AndroidManifest.xml" \
  "$BUILD_DIR/resources.zip" 2>&1

# 6. Add DEX and native lib
echo "==> Adding DEX and native lib..."
mkdir -p "$BUILD_DIR/staging/lib/$LIBDIR"
unzip -qo "$BUILD_DIR/unsigned.apk" -d "$BUILD_DIR/staging"
cp "$BUILD_DIR/dex/classes.dex" "$BUILD_DIR/staging/"
cp "$CORRO_ROOT/android/corro/target/$TRIPLE/release/libcorro_android.so" \
  "$BUILD_DIR/staging/lib/$LIBDIR/"
cd "$BUILD_DIR/staging" && zip -qr "$BUILD_DIR/unsigned_with_libs.apk" . && cd /tmp

# 7. Align and sign
echo "==> Aligning and signing..."
zipalign -v -p 4 "$BUILD_DIR/unsigned_with_libs.apk" "$BUILD_DIR/aligned.apk" 2>&1 | tail -1

KEYSTORE="$HOME/.android/corro-debug.keystore"
[ -f "$KEYSTORE" ] || keytool -genkey -v -keystore "$KEYSTORE" -alias androiddebugkey \
  -keyalg RSA -keysize 2048 -validity 10000 -storepass android -keypass android \
  -dname "CN=, OU=, O=, L=, S=, C=" 2>&1 >/dev/null

apksigner sign --ks "$KEYSTORE" --ks-pass pass:android --ks-key-alias androiddebugkey \
  --v1-signing-enabled true --v2-signing-enabled true --min-sdk-version 24 \
  --out "$BUILD_DIR/signed.apk" "$BUILD_DIR/aligned.apk" 2>&1

echo "==> APK built: $BUILD_DIR/signed.apk"

# Optionally install and run
if [ "${2:-}" = "install" ]; then
  echo "==> Installing..."
  adb install -r "$BUILD_DIR/signed.apk" 2>&1 | tail -1
  echo "==> Launching..."
  adb shell am start -n com.corro/.MainActivity
fi
