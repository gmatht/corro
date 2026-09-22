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
run "ios workflow parses and steps are bash-clean" bash -c "python3 - <<'PY'
import subprocess, sys, tempfile, os
try:
    import yaml
except ImportError:
    print('pyyaml not installed; skipping'); sys.exit(0)
wf = '$CORRO_ROOT/.github/workflows/ios.yml'
d = yaml.safe_load(open(wf))
assert 'jobs' in d, 'no jobs'
bad = 0
for job, spec in d['jobs'].items():
    for step in spec.get('steps', []):
        if 'run' not in step:
            continue
        with tempfile.NamedTemporaryFile('w', suffix='.sh', delete=False) as f:
            f.write(step['run']); path = f.name
        r = subprocess.run(['bash', '-n', path], capture_output=True, text=True)
        os.unlink(path)
        if r.returncode:
            bad += 1
            print('bad step:', step.get('name'), r.stderr[:200])
assert bad == 0, f'{bad} steps with bash syntax errors'
txt = open(wf).read()
assert txt.index('SIM_UDID=') < txt.index('\$SIM_UDID'), 'SIM_UDID used before set'
print('workflow ok:', list(d['jobs']))
PY"
run "every selector the backend sends is implemented in the shim" "$HERE/check_selectors.sh"
run "xcodeproj definitions are unique" bash -c "python3 - <<'PY'
import re, sys
from collections import Counter
p = '/tmp/iv_xp/Corro.xcodeproj/project.pbxproj'
s = open(p).read()
defs = re.findall(r'^\\t\\t(C0DE[0-9A-F]+) /\\* [^*]+ \\*/ = \\{isa = ([A-Za-z]+);', s, re.M)
if not defs:
    print('no object definitions found'); sys.exit(1)
c = Counter(i for i, _ in defs)
dupes = {k: v for k, v in c.items() if v > 1}
if dupes:
    print('duplicate object ids:', dupes); sys.exit(1)
print('object definitions:', len(defs), 'unique')
PY"
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
