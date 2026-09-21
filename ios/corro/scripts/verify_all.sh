#!/bin/bash
# Full verification sweep for the iOS work: every check that can run without
# macOS, plus the desktop regressions that must not have broken.
#
# What it deliberately does NOT cover: compiling the ObjC/Swift and running the
# app. That needs macOS + Xcode — `build_ios.sh sim --run` is the manual step
# (docs/IOS_GUIDELINES.md §9).
set -uo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
HOST="$(cd "$HERE/.." && pwd)"
CORRO_ROOT="$(cd "$HERE/../../.." && pwd)"
pass=0; fail=0
run() {
  local label="$1"; shift
  if "$@" >/tmp/ios_verify_out 2>&1; then
    echo "PASS  $label"; pass=$((pass+1))
  else
    echo "FAIL  $label"; tail -8 /tmp/ios_verify_out | sed 's/^/        /'; fail=$((fail+1))
  fi
}
run "rswidgets iOS check (sim + arm64 + armv7s)" "$HERE/check_rswidgets_ios.sh"
run "corro iOS check (sim + arm64 + armv7s)"     "$HERE/check_corro_ios.sh"
run "host cdylib check"                          "$HERE/check_host_ios.sh"
run "desktop rswidgets check" bash -c "cd '$CORRO_ROOT' && cargo check -p rswidgets"
run "desktop corro gui check" bash -c "cd '$CORRO_ROOT' && cargo check --features gui"
run "desktop TUI check"       bash -c "cd '$CORRO_ROOT' && cargo check"
run "corro tests"             bash -c "cd '$CORRO_ROOT' && cargo test --lib --quiet"
run "rswidgets tests"         bash -c "cd '$CORRO_ROOT' && cargo test -p rswidgets --quiet"
run "ios-ui example runs"     bash -c "cd '$CORRO_ROOT' && cargo run --quiet --example ios-ui --features gui-mobile-host"
run "xcodeproj generator"     bash -c "rm -rf /tmp/iv_xp && mkdir -p /tmp/iv_xp && '$HERE/../gen_xcodeproj.sh' /tmp/iv_xp/Corro.xcodeproj"
run "xcodeproj ids well-formed" bash -c "python3 -c '
import re
s = open(\"/tmp/iv_xp/Corro.xcodeproj/project.pbxproj\").read()
ids = set(re.findall(r\"C0DE[0-9A-F]+\", s))
assert ids, \"no identifiers\"
assert all(len(i) == 24 for i in ids), sorted(i for i in ids if len(i) != 24)
defined = set(re.findall(r\"^\t\t(C0DE[0-9A-F]+)\", s, re.M))
assert not (ids - defined), ids - defined
'"
echo
echo "passed: $pass   failed: $fail"
[ "$fail" -eq 0 ]
