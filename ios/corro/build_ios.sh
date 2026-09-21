#!/bin/bash
# Build the corro iOS app: Rust static library + ObjC shims -> .app/.ipa.
#
# Mirrors android/corro/build_apk.sh in spirit (raw toolchain, reproducible,
# no IDE required) but the iOS toolchain is Xcode's, so this drives
# `xcodebuild` rather than hand-invoking clang and the packager.
#
# Usage: ./build_ios.sh [sim|device] [--ipa] [--run]
#   sim      build for the iOS simulator (no signing needed)
#   device   build for arm64 devices (needs a signing identity + profile)
#   --ipa    additionally export a signed .ipa (device only)
#   --run    install and launch on the booted simulator (sim only)
#
# Environment:
#   IOS_DEPLOYMENT_TARGET   deployment version (default 12.0; see README)
#   IOS_SIM_NAME            simulator to build for (default: first available iPhone)
#   DEVELOPMENT_TEAM        Apple team id, required for device builds
#   CARGO_TARGET_DIR        Rust target dir (default: this crate's target/)
#
# What it does, in order:
#   1. `cargo build --release` the corro_ios staticlib for the chosen target,
#      with -Zbuild-std (the iOS targets are tier-3 or need a matching std).
#   2. generate ios/corro/app/Corro.xcodeproj/project.pbxproj if absent (the
#      project is checked in, so this only fills a gap).
#   3. `xcodebuild` the app, pointing at the Rust staticlib.
set -euo pipefail

MODE="${1:-sim}"
shift || true
WANT_IPA=0
WANT_RUN=0
for arg in "$@"; do
  case "$arg" in
    --ipa) WANT_IPA=1 ;;
    --run) WANT_RUN=1 ;;
    *) echo "unknown option: $arg" >&2; exit 1 ;;
  esac
done

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CORRO_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
APP_DIR="$SCRIPT_DIR/app"
BUILD_DIR="${BUILD_DIR:-/tmp/corro_ios_build}"
IOS_DEPLOYMENT_TARGET="${IOS_DEPLOYMENT_TARGET:-12.0}"

case "$MODE" in
  sim)
    # Apple silicon host -> arm64 simulator; Intel host -> x86_64. `uname -m`
    # is the honest probe: the simulator runs the host's architecture.
    case "$(uname -m)" in
      arm64) RUST_TARGET="aarch64-apple-ios-sim"; XCODE_SDK="iphonesimulator"; XCODE_ARCH="arm64" ;;
      *)     RUST_TARGET="x86_64-apple-ios";     XCODE_SDK="iphonesimulator"; XCODE_ARCH="x86_64" ;;
    esac
    SIGN_ARGS=("CODE_SIGNING_ALLOWED=NO" "CODE_SIGNING_REQUIRED=NO" "CODE_SIGN_IDENTITY=")
    ;;
  device)
    RUST_TARGET="aarch64-apple-ios"; XCODE_SDK="iphoneos"; XCODE_ARCH="arm64"
    SIGN_ARGS=()
    if [ -z "${DEVELOPMENT_TEAM:-}" ]; then
      echo "warning: DEVELOPMENT_TEAM unset; a device build needs a signing identity" >&2
    fi
    ;;
  *) echo "unknown mode: $MODE (want sim or device)" >&2; exit 1 ;;
esac

echo "==> Building corro_ios for $RUST_TARGET (deployment target $IOS_DEPLOYMENT_TARGET)"
cd "$SCRIPT_DIR"

# -Zbuild-std needs the target's `core`/`std` SOURCES, not a prebuilt target:
# the iOS targets are either tier-3 (armv7s) or have no std installed. On a
# fresh machine — a hosted macOS runner, a new laptop — `rust-src` is missing
# and the build fails deep inside cargo with a confusing message, so check
# first and say what to run. (`rustup` may legitimately be absent when the
# toolchain came from a system package or a vendored copy; in that case the
# component is expected to be present already, and cargo will say so.)
if command -v rustup >/dev/null 2>&1; then
  if ! rustup component list --installed 2>/dev/null | grep -q '^rust-src'; then
    echo "    rust-src component missing; installing it (needed by -Zbuild-std)"
    rustup component add rust-src --toolchain nightly || {
      echo "could not install rust-src; run: rustup component add rust-src --toolchain nightly" >&2
      exit 1
    }
  fi
fi
if ! command -v cargo >/dev/null 2>&1; then
  echo "cargo not found in PATH" >&2
  exit 1
fi
# -Zbuild-std: the iOS targets have no prebuilt std in this toolchain. Metadata
# alone would not do — this has to produce a real staticlib for the link step —
# so an Xcode install is genuinely required from here on, unlike the host-side
# `cargo check` in /tmp/corro_ios_check.sh.
IPHONEOS_DEPLOYMENT_TARGET="$IOS_DEPLOYMENT_TARGET" \
  cargo +nightly build --release \
    --target "$RUST_TARGET" \
    -Zbuild-std=std,panic_abort \
    --config "build.rustflags=[\"-C\",\"link-arg=-mios-version-min=$IOS_DEPLOYMENT_TARGET\"]"

RLIB="$SCRIPT_DIR/target/$RUST_TARGET/release/libcorro_ios.a"
if [ ! -f "$RLIB" ]; then
  echo "Rust staticlib not found at $RLIB" >&2
  exit 1
fi
echo "    -> $(du -h "$RLIB" | cut -f1) $RLIB"

# The Xcode project is checked in; generating it only when missing keeps a
# hand-tuned project from being silently overwritten, the same additive rule
# rswidgets::android_generator follows for Android resources.
PROJECT="$APP_DIR/Corro.xcodeproj"
if [ ! -d "$PROJECT" ]; then
  echo "==> Generating $PROJECT"
  mkdir -p "$PROJECT"
  "$SCRIPT_DIR/gen_xcodeproj.sh" "$PROJECT"
fi

# Xcode's build dir is redirected into BUILD_DIR so nothing lands in the
# source tree (and so a CI cache can own it).
echo "==> xcodebuild ($XCODE_SDK, $XCODE_ARCH)"
xcodebuild \
  -project "$PROJECT" \
  -scheme corro \
  -configuration Release \
  -sdk "$XCODE_SDK" \
  -arch "$XCODE_ARCH" \
  -derivedDataPath "$BUILD_DIR/DerivedData" \
  CONFIGURATION_BUILD_DIR="$BUILD_DIR/$XCODE_SDK" \
  IPHONEOS_DEPLOYMENT_TARGET="$IOS_DEPLOYMENT_TARGET" \
  RUST_LIB_PATH="$RLIB" \
  ${DEVELOPMENT_TEAM:+DEVELOPMENT_TEAM="$DEVELOPMENT_TEAM"} \
  "${SIGN_ARGS[@]}" \
  build | tail -20

APP="$BUILD_DIR/$XCODE_SDK/corro.app"
if [ ! -d "$APP" ]; then
  echo "app bundle not found at $APP" >&2
  exit 1
fi
echo "==> Built $APP"

if [ "$MODE" = "device" ] && [ "$WANT_IPA" = 1 ]; then
  echo "==> Exporting .ipa"
  EXPORT_PLIST="$BUILD_DIR/ExportOptions.plist"
  cat > "$EXPORT_PLIST" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>method</key><string>development</string>
  <key>teamID</key><string>${DEVELOPMENT_TEAM:-}</string>
  <key>compileBitcode</key><false/>
</dict>
</plist>
PLIST
  xcodebuild -exportArchive \
    -archivePath "$BUILD_DIR/DerivedData/Build/Products/Release-iphoneos/corro.xcarchive" \
    -exportPath "$BUILD_DIR/ipa" \
    -exportOptionsPlist "$EXPORT_PLIST"
  echo "==> $(ls "$BUILD_DIR"/ipa/*.ipa)"
  echo "    Upload this to LambdaTest App Live (see README) or install via"
  echo "    Xcode Devices / Apple Configurator."
fi

if [ "$MODE" = "sim" ] && [ "$WANT_RUN" = 1 ]; then
  echo "==> Installing on the booted simulator"
  xcrun simctl install booted "$APP"
  xcrun simctl launch booted com.corro.corro
  echo "    Logs: xcrun simctl spawn booted log stream --predicate 'eventMessage CONTAINS \"rswidgets\"'"
fi
