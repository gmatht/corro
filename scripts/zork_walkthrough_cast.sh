#!/usr/bin/env bash
# Record the zork backend's feature suite as an asciinema cast.
#
# Why this exists: `docs/ZORK_WALKTHROUGH.txt` records *what the backend built*,
# which is the artefact a diff reads. It says nothing about whether the tests
# that check it pass. This is the other half — the run itself, in the one form a
# human can watch rather than take a green tick on faith.
#
# Two parts, because a reviewer needs both and neither alone is enough:
#
#   1. the app, which prints its summary and writes the full walkthrough. This
#      is the record itself, and it is only reachable through the handoff (see
#      the note on `take_state` in `zork_handoff`), so a cast is the one way to
#      watch it being produced rather than only read the result.
#   2. the suite that asserts against it. Without this the cast is a demo; with
#      it, the demo is checked.
#
# It records the zork suite (`tests/zork_feature_parity.rs`) because that is the
# one that runs on a text backend: no X server, no OCR, no GTK.
#
# Usage:
#   scripts/zork_walkthrough_cast.sh            # docs/zork_feature_suite.cast
#   scripts/zork_walkthrough_cast.sh out.cast   # somewhere else
#
# The cast is committed beside `docs/corro.cast` and `docs/ZORK_WALKTHROUGH.txt`
# for the same reason those are: it is a record, and a reviewer should be able to
# see the run rather than trust a tick mark.
#
# Reproducibility: two runs from the same checkout and the same working
# directory produce identical content, and the test list is in a fixed order
# (see `--test-threads=1` below). What is *not* stable is the absolute path the
# app echoes when it writes the walkthrough, and cargo's own chatter -- a
# "Blocking waiting for file lock" line appears whenever another build is
# running. Both are environment, not script, and a regenerated cast differing
# only in those is the same recording.

set -euo pipefail

cd "$(dirname "$0")/.."

OUT="${1:-docs/zork_feature_suite.cast}"
SUITE="cargo +stable test --features zork --no-default-features --test zork_feature_parity"

command -v asciinema >/dev/null || {
    cat >&2 <<EOF
asciinema is not installed, so the recording is skipped.

The suite can still be run directly:
  $SUITE
EOF
    exit 0
}

# Build first, so the recording is not dominated by compile output and a compile
# error is reported as a compile error rather than as a truncated cast. `-q`
# because the point of the cast is the run, not the build.
echo "Building the zork binary (quiet)..." >&2
cargo +stable build -q --features zork --no-default-features

mkdir -p "$(dirname "$OUT")"

# The two parts as one script, so a single cast covers the whole story. The
# walkthrough goes to a scratch file: the committed copy in docs/ is a *record*
# of a past run, and having the recording rewrite it would make a checkout
# dirty on every playback.
# A fixed name, not `mktemp`: the app echoes the path it wrote to, so a random
# one makes two recordings of the same run differ in their content. Reproducible
# beats unique here -- the file is removed on exit either way, and nothing else
# runs concurrently in a scratch dir of our own making.
WALK="$PWD/target/zork_walkthrough_recording.txt"
TOUR="$(mktemp)"
trap 'rm -f "$WALK" "$TOUR"' EXIT

cat > "$TOUR" <<TOUR_SCRIPT
set -u
cd "$(pwd)"

rule() { printf '%s\n' "------------------------------------------------------------------------"; }
step() { printf '\n\n%s\n' "\$*"; rule; }

printf '\n'
rule
printf 'corro %s on the zork backend\n' "$(cargo +stable run -q --features zork --no-default-features -- --version 2>/dev/null | head -1 | sed 's/corro //')"
rule

step "[part 1 of 3]  The app builds its GUI tree, then records what it built"
printf 'The widget tree is corro'"'"'s real GUI construction path, built against the\n'
printf 'in-memory model instead of a display. The record is written to a file,\n'
printf 'because it has to be taken while the tree is still live: the REPL adopts\n'
printf 'the model on startup, after which all 98 nodes are gone.\n'
printf '\n'
printf '  \$ echo quit | CORRO_ZORK_WALKTHROUGH=<file> ./target/debug/corro\n'
printf '\n'
CORRO_ZORK_WALKTHROUGH="$WALK" ./target/debug/corro <<<'quit' || true

step "[part 2 of 3]  The recorded walkthrough, verbatim"
printf 'This is the whole record, verbatim -- every feature, every node, the\n'
printf 'full menu tree, and what the grid'"'"'s draw callback painted.\n'
printf '\n'
cat "$WALK"

step "[part 3 of 3]  The suite that asserts against it"
printf '32 tests, no X server, no OCR, no GTK. They re-ask the GUI feature\n'
printf 'questions against the widget tree, and check the record above.\n'
printf '\n'
# `--test-threads=1` for two reasons. The recording should be reproducible, and
# cargo schedules parallel tests in whatever order the machine picks -- two runs
# of the same suite differ in line order, which makes a cast impossible to
# diff. And the zork model is a thread-local singleton, so the adapter tests
# each reset it and are clearer read one at a time. The suite passes either
# way; this is about the recording being a record.
#
# The listing is left as cargo writes it, because that is the honest output.
# One consequence is that cargo's `Finished ...` status line lands on the same
# line as whichever test happens to be running when it prints, so one test name
# (`app_state_is_reported`, at the mercy of the scheduler) is not on a line of
# its own. It is in the recording and a player renders it; it is just sharing.
# Filtering cargo's chatter would tidy that at the cost of the recording no
# longer being the command's real output, which is the thing that makes it
# evidence.
$SUITE -- --test-threads=1
TOUR_SCRIPT

# 120x30 matches `docs/corro.cast`. Left to itself asciinema takes the
# dimensions of whatever terminal is attached, so a cast recorded on a wide
# screen is unwatchable in a narrow player -- and a record nobody can watch is
# not much of a record.
#
# 30 rows is also what keeps the walkthrough from being truncated mid-section
# when it scrolls, which is where the record's value is: a viewer can scroll
# back through the menu tree and the canvas numbers, or read the file instead.
#
# The idle limit is generous because there are multi-second steps (two full app
# runs for the determinism check); the default 2s would freeze the typing cursor
# through them. The cap keeps a hang from producing a cast that never ends.
#
# `-c` rather than piping stdin, so the script runs under a real pty and behaves
# as it does for a person at a keyboard.
asciinema rec \
    --overwrite \
    --idle-time-limit 10 \
    --cols 120 \
    --rows 30 \
    --title "corro on the zork backend — walkthrough and feature suite" \
    -c "bash $TOUR" \
    "$OUT"

echo "Wrote $OUT" >&2
echo "Play it with:  asciinema play $OUT" >&2
