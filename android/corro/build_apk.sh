#!/bin/bash
# Build the corro Android APK from scratch (mirrors
# rustxWidgets/examples/android/build_apk.sh for the widget demo).
# Usage: ./build_apk.sh [ABI] [install]
#   ABI defaults to x86_64 (emulator); use arm64-v8a for real devices.
set -euo pipefail

ANDROID_HOME="${ANDROID_HOME:-$HOME/android-sdk}"
ANDROID_NDK_HOME="${ANDROID_NDK_HOME:-$ANDROID_HOME/ndk/25.2.9519653}"
# Build-tools 34+: d8 8.x and aapt2 that understands API 34 resources. The
# 33.0.2 d8 in this image crashes on AndroidX class files, and Material 1.12
# needs a compileSdk of 34.
BUILD_TOOLS="${BUILD_TOOLS:-34.0.0}"
PLATFORM="${PLATFORM:-android-34}"
PATH="$ANDROID_HOME/cmdline-tools/latest/bin:$ANDROID_HOME/platform-tools:$ANDROID_HOME/build-tools/$BUILD_TOOLS:$ANDROID_HOME/emulator:$PATH"

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

# 1b. Fetch the AndroidX/Material AARs the generated theme needs.
# The generated theme parents against Theme.Material3.*, which lives in the
# Material Components AAR (see rustxWidgets/docs/ANDROID_GUIDELINES.md §1:
# rswidgets does not fetch it). A Gradle build resolves these from Maven;
# this raw-toolchain script has to do it by hand. Cached in
# $BUILD_DIR/aars, so a rebuild is offline. Override AAR_CACHE to reuse a
# pre-populated tree (e.g. in CI).
AAR_CACHE="${AAR_CACHE:-$BUILD_DIR/aars}"
MAVEN=https://dl.google.com/dl/android/maven2
# name|group path|version
AAR_LIST="
material|com/google/android/material/material|1.12.0
appcompat|androidx/appcompat/appcompat|1.7.1
core|androidx/core/core|1.12.0
activity|androidx/activity/activity|1.7.2
fragment|androidx/fragment/fragment|1.6.2
lifecycle-runtime|androidx/lifecycle/lifecycle-runtime|2.6.2
cardview|androidx/cardview/cardview|1.0.0
coordinatorlayout|androidx/coordinatorlayout/coordinatorlayout|1.2.0
drawerlayout|androidx/drawerlayout/drawerlayout|1.2.0
recyclerview|androidx/recyclerview/recyclerview|1.3.2
viewpager2|androidx/viewpager2/viewpager2|1.1.0
transition|androidx/transition/transition|1.4.1
vectordrawable|androidx/vectordrawable/vectordrawable|1.1.0
vectordrawable-animated|androidx/vectordrawable/vectordrawable-animated|1.1.0
dynamicanimation|androidx/dynamicanimation/dynamicanimation|1.0.0
constraintlayout|androidx/constraintlayout/constraintlayout|2.1.4
"
mkdir -p "$AAR_CACHE"
echo "==> Fetching AndroidX/Material AARs into $AAR_CACHE..."
while IFS='|' read -r name path ver; do
  [ -n "$name" ] || continue
  dir="$AAR_CACHE/$name"
  if [ ! -d "$dir" ]; then
    mkdir -p "$dir"
    # </dev/null: unzip/curl must not eat the loop's here-string.
    curl -sfL "$MAVEN/$path/$ver/$name-$ver.aar" -o "$dir/$name.aar" </dev/null \
      || { echo "failed to fetch $name-$ver.aar" >&2; exit 1; }
    unzip -qo "$dir/$name.aar" 'res/*' 'classes.jar' -d "$dir" </dev/null 2>/dev/null || true
  fi
done <<< "$AAR_LIST"

# aapt2 cannot merge two libraries that both declare the same
# <declare-styleable> (Material and appcompat both define SearchView and
# Carousel), and it cannot see resources from separately-compiled zips.
# Compile every library's res/ tree in one pass, then drop the duplicate
# styleable blocks. Only the resource *table* is affected: the richer
# (Material) definition is kept.
echo "==> Merging library resources..."
rm -rf "$BUILD_DIR/libres" "$BUILD_DIR/libres.zip"
mkdir -p "$BUILD_DIR/libres"
# Copy every library's res/ tree, merging (not replacing) same-named XML files
# and preserving each file's xmlns declarations. A plain `cp -r` overwrites
# values-v31.xml etc., silently dropping the definitions that live only in
# the overwritten copy — which then fails link with "resource ... not found".
python3 - "$BUILD_DIR/libres" "$AAR_CACHE" <<'PY'
import os, re, shutil, sys
out, cache = sys.argv[1], sys.argv[2]

def merge_xml(dst, src):
    a = open(dst, encoding='utf-8').read()
    b = open(src, encoding='utf-8').read()
    def decls(x):
        m = re.search(r'<resources([^>]*)>', x)
        return dict(re.findall(r'(xmlns:[\w-]+)="([^"]*)"', m.group(1))) if m else {}
    d = decls(a); d.update(decls(b))
    a = re.sub(r'<resources[^>]*>', '<resources' + ''.join(f' {k}="{v}"' for k, v in sorted(d.items())) + '>', a, count=1)
    b = re.sub(r'^\s*<\?xml[^>]*\?>\s*', '', b)
    b = re.sub(r'^\s*<resources[^>]*>', '', b)
    b = re.sub(r'</resources>\s*$', '', b)
    open(dst, 'w', encoding='utf-8').write(re.sub(r'</resources>\s*$', b + '\n</resources>\n', a))

for name in sorted(os.listdir(cache)):
    base = os.path.join(cache, name, 'res')
    if not os.path.isdir(base):
        continue
    for root, _dirs, files in os.walk(base):
        rel = os.path.relpath(root, base)
        dst_dir = os.path.join(out, rel) if rel != '.' else out
        os.makedirs(dst_dir, exist_ok=True)
        for f in files:
            s = os.path.join(root, f); t = os.path.join(dst_dir, f)
            if os.path.exists(t) and f.endswith('.xml'):
                merge_xml(t, s)
            else:
                shutil.copy2(s, t)

# aapt2 rejects duplicate <declare-styleable> blocks (Material and appcompat
# both declare SearchView and Carousel). Keep the longest definition — the
# superset — and drop the others. Only the resource table is affected.
vals = os.path.join(out, 'values', 'values.xml')
if os.path.exists(vals):
    s = open(vals, encoding='utf-8').read()
    for name in ('SearchView', 'Carousel'):
        blocks = list(re.finditer(r'<declare-styleable name="%s">.*?</declare-styleable>' % name, s, re.S))
        if len(blocks) > 1:
            keep = max(blocks, key=lambda m: len(m.group(0)))
            for m in reversed(blocks):
                if m.start() != keep.start():
                    s = s[:m.start()] + s[m.end():]
    open(vals, 'w', encoding='utf-8').write(s)
PY
aapt2 compile --dir "$BUILD_DIR/libres" -o "$BUILD_DIR/libres.zip" 2>&1 | grep -iE 'error' && exit 1 || true

# 2. Compile Java sources
echo "==> Compiling Java..."
JAVAC_CP="$ANDROID_HOME/platforms/$PLATFORM/android.jar"
for d in "$AAR_CACHE"/*/; do [ -f "$d/classes.jar" ] && JAVAC_CP="$JAVAC_CP:$d/classes.jar"; done
javac -d "$BUILD_DIR/classes" -source 8 -target 8 \
  -bootclasspath "$ANDROID_HOME/platforms/$PLATFORM/android.jar" \
  -cp "$JAVAC_CP" \
  "$SCRIPT_DIR/app/src/main/java/com/corro/"*.java 2>&1 | grep -v warning || true

# 3. Convert to DEX. The app classes plus every AAR's classes.jar: without
# them the runtime would fail with NoClassDefFoundError on the first
# Material theme attribute lookup.
#
# `dx` (build-tools 29) is used when present because the `d8` shipped in
# build-tools 33/34 in this image crashes with an internal NPE on the
# AndroidX class files (anonymous inner classes without EnclosingMethod).
# dx needs --min-sdk-version >= 26 for their invokedynamic (string concat).
echo "==> Converting to DEX..."
DEPS=$(for d in "$AAR_CACHE"/*/; do [ -f "$d/classes.jar" ] && echo "$d/classes.jar"; done | tr '\n' ' ')
DX=""
for bt in 29.0.3 29.0.2 30.0.3; do
  [ -x "$ANDROID_HOME/build-tools/$bt/dx" ] && { DX="$ANDROID_HOME/build-tools/$bt/dx"; break; }
done
mkdir -p "$BUILD_DIR/dex"
if [ -n "$DX" ]; then
  (cd "$BUILD_DIR/classes" && "$DX" --dex --min-sdk-version=26 \
    --output "$BUILD_DIR/dex/" --core-library . $DEPS) 2>&1 | grep -iE '^.*error' && exit 1 || true
else
  echo "    (dx unavailable; falling back to d8)" >&2
  d8 --release --min-api 26 --lib "$ANDROID_HOME/platforms/$PLATFORM/android.jar" \
    --output "$BUILD_DIR/dex" "$BUILD_DIR/classes/com/corro/"*.class $DEPS 2>&1 | grep -iE '^.*error' && exit 1 || true
fi

# 4. Package resources
echo "==> Packaging resources..."
aapt2 compile --dir "$SCRIPT_DIR/app/src/main/res" -o "$BUILD_DIR/resources.zip" 2>&1 | grep -v "^INFO\|^$" || true

# 5. Link base APK (app resources + merged library resources)
echo "==> Linking APK..."
aapt2 link -o "$BUILD_DIR/unsigned.apk" \
  -I "$ANDROID_HOME/platforms/$PLATFORM/android.jar" \
  --manifest "$SCRIPT_DIR/app/src/main/AndroidManifest.xml" \
  --java "$BUILD_DIR/gen" \
  --min-sdk-version 24 --target-sdk-version 36 \
  "$BUILD_DIR/resources.zip" "$BUILD_DIR/libres.zip" 2>&1 | grep -iE 'error' && exit 1 || true

# 6. Add DEX and native lib. resources.arsc must be STORED (not deflated)
# and 4-byte aligned, or the API 30+ installer rejects the package
# ("Targeting R+ requires the resources.arsc ... uncompressed and aligned").
echo "==> Adding DEX and native lib..."
mkdir -p "$BUILD_DIR/staging/lib/$LIBDIR"
unzip -qo "$BUILD_DIR/unsigned.apk" -d "$BUILD_DIR/staging"
cp "$BUILD_DIR/dex/classes.dex" "$BUILD_DIR/staging/"
cp "$CORRO_ROOT/android/corro/target/$TRIPLE/release/libcorro_android.so" \
  "$BUILD_DIR/staging/lib/$LIBDIR/"
python3 - "$BUILD_DIR/staging" "$BUILD_DIR/unsigned_with_libs.apk" <<'PY'
import os, sys, zipfile
staging, out = sys.argv[1], sys.argv[2]
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for root, _dirs, files in os.walk(staging):
        for f in files:
            full = os.path.join(root, f)
            rel = os.path.relpath(full, staging)
            zi = zipfile.ZipInfo(rel)
            zi.compress_type = zipfile.ZIP_STORED if rel == 'resources.arsc' else zipfile.ZIP_DEFLATED
            zi.external_attr = 0o644 << 16
            with open(full, 'rb') as fh:
                z.writestr(zi, fh.read())
PY

# 7. Align and sign
echo "==> Aligning and signing..."
zipalign -f -p 4 "$BUILD_DIR/unsigned_with_libs.apk" "$BUILD_DIR/aligned.apk" 2>&1 | tail -1

KEYSTORE="$HOME/.android/corro-debug.keystore"
[ -f "$KEYSTORE" ] || keytool -genkey -v -keystore "$KEYSTORE" -alias androiddebugkey \
  -keyalg RSA -keysize 2048 -validity 10000 -storepass android -keypass android \
  -dname "CN=, OU=, O=, L=, S=, C=" 2>&1 >/dev/null

apksigner sign --ks "$KEYSTORE" --ks-pass pass:android --ks-key-alias androiddebugkey \
  --v1-signing-enabled true --v2-signing-enabled true --min-sdk-version 26 \
  --out "$BUILD_DIR/signed.apk" "$BUILD_DIR/aligned.apk" 2>&1

echo "==> APK built: $BUILD_DIR/signed.apk"

# Optionally install and run
if [ "${2:-}" = "install" ]; then
  echo "==> Installing..."
  adb install -r "$BUILD_DIR/signed.apk" 2>&1 | tail -1
  echo "==> Launching..."
  adb shell am start -n com.corro/.MainActivity
fi
