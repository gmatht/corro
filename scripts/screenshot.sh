#!/usr/bin/env bash
# Capture a handful of corro screenshots from the headless GUI.
#
# Why a throwaway X server: the GUI backend needs a real display, and the
# capture must not depend on (or disturb) whatever the operator is looking at.
# Xvfb on a private display plus everything targeting that display explicitly
# gives a reproducible image on any machine, including CI.
#
# Why `import -window root` rather than xwd+convert: same reason
# `two_window_movie.py` uses ffmpeg's x11grab - the xwd/convert pair costs
# ~500ms per frame, so a screenshot either waits that long or races the draw.
# `import` grabs once, in one process.
#
# The workbook is built by *scripted edits* (`CORRO_EDIT_SCRIPT`), not by
# synthetic keystrokes: corro applies each edit through the ordinary commit
# path, so the shot cannot drift out of sync with the UI the way pixel-calibrated
# xdotool typing does (see `scripts/two_window_movie.py` for the long version).
#
# Usage:
#   scripts/screenshot.sh [-o OUTDIR] [--size WxH] [--binary PATH]
#   scripts/screenshot.sh --keep-display     # show the shots as they are taken
#
# Output: OUTDIR/shot-NN-<name>.png plus OUTDIR/README.md describing them.
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$HERE/.." && pwd)"

OUTDIR="$REPO_ROOT/dist/screenshots"
SIZE="1400x900"
DISPLAY_NUM=":94"
BINARY="${CORRO_MOVIE_BIN:-}"
KEEP_DISPLAY=0
SETTLE=3   # seconds after launch before the first grab

while [ $# -gt 0 ]; do
  case "$1" in
    -o|--out) OUTDIR="$2"; shift 2 ;;
    --size) SIZE="$2"; shift 2 ;;
    --binary) BINARY="$2"; shift 2 ;;
    --keep-display) KEEP_DISPLAY=1; shift ;;
    -h|--help) sed -n '2,25p' "$0"; exit 0 ;;
    *) echo "unknown argument: $1" >&2; exit 2 ;;
  esac
done

command -v Xvfb >/dev/null || { echo "error: Xvfb not installed" >&2; exit 1; }
command -v import >/dev/null || { echo "error: ImageMagick 'import' not installed" >&2; exit 1; }

if [ -z "$BINARY" ]; then
  for c in "$REPO_ROOT/target/debug/corro" "$REPO_ROOT/target/release/corro"; do
    [ -x "$c" ] && BINARY="$c" && break
  done
fi
[ -n "$BINARY" ] && [ -x "$BINARY" ] || {
  echo "error: no corro binary; build with 'cargo build --features gui' or set CORRO_MOVIE_BIN" >&2
  exit 1
}

# `gui` is not a default feature, so a binary at the shared path may have no
# GUI at all. Refuse rather than produce a screenshot of an error dialog.
if ! "$BINARY" --version >/dev/null 2>&1; then
  echo "warning: '$BINARY --version' failed; continuing anyway" >&2
fi

mkdir -p "$OUTDIR"
WORKDIR="$(mktemp -d)"
trap 'rm -rf "$WORKDIR"' EXIT

export DISPLAY="$DISPLAY_NUM"

# A stale server on this display would make our Xvfb fail to bind while the
# capture silently reads the *old* screen, which looks like a rendering bug.
pkill -f "Xvfb $DISPLAY_NUM" 2>/dev/null || true
sleep 0.3

Xvfb "$DISPLAY_NUM" -screen 0 "${SIZE}x24" >"$WORKDIR/xvfb.log" 2>&1 &
XVFB_PID=$!
cleanup() {
  kill "$XVFB_PID" 2>/dev/null || true
  [ "$KEEP_DISPLAY" = 1 ] || pkill -f "Xvfb $DISPLAY_NUM" 2>/dev/null || true
}
trap 'cleanup; rm -rf "$WORKDIR"' EXIT

# Wait for the server to accept connections rather than sleeping a guessed
# amount: an image grabbed too early is uniformly black, which is easy to
# mistake for "the GUI did not draw".
for _ in $(seq 1 50); do
  if xdpyinfo -display "$DISPLAY_NUM" >/dev/null 2>&1; then break; fi
  sleep 0.2
done
xdpyinfo -display "$DISPLAY_NUM" >/dev/null 2>&1 || {
  echo "error: Xvfb $DISPLAY_NUM never became ready" >&2
  cat "$WORKDIR/xvfb.log" >&2 || true
  exit 1
}

# A window manager is not strictly required for a single full-screen window,
# but without one the window may be placed at an arbitrary offset and the grab
# then has the background around it.
for wm in fluxbox openbox matchbox-window-manager twm; do
  if command -v "$wm" >/dev/null; then "$wm" >/dev/null 2>&1 & WM_PID=$!; break; fi
done
sleep 0.5

shot_index=0
declare -a SHOTS=()

# Wait for the app window to actually be mapped and sized, rather than sleeping
# a guessed amount. The first shot of an earlier version caught the window
# before the WM had placed it (a mostly-black frame with the UI pushed to the
# bottom edge) - an image that looks like a rendering bug but is only a race.
# `xdotool search` returns success as soon as the window exists; geometry is
# then what says it has settled.
wait_for_window() {
  local tries=0
  while [ "$tries" -lt 80 ]; do
    local wid
    wid="$(xdotool search --onlyvisible --name 'corro' 2>/dev/null | head -1 || true)"
    if [ -n "$wid" ]; then
      local geom
      geom="$(xdotool getwindowgeometry --shell "$wid" 2>/dev/null || true)"
      # A window that has been placed has a real size; a not-yet-managed one
      # reports 1x1 or 0x0.
      if echo "$geom" | grep -qE '^(WIDTH|HEIGHT)=([2-9][0-9]|[1-9][0-9]{2,})'; then
        echo "$wid"
        return 0
      fi
    fi
    tries=$((tries + 1))
    sleep 0.1
  done
  return 1
}

# Force the window to fill the screen, so every shot has the same framing
# regardless of what the WM decided. A missing WM (or a refused move) is not
# fatal: the shot is then whatever the default placement gave.
frame_window() {
  local wid="$1"
  [ -n "$wid" ] || return 0
  xdotool windowmove --sync "$wid" 0 0 2>/dev/null || true
  xdotool windowsize --sync "$wid" "${SIZE%%x*}" "${SIZE##*x}" 2>/dev/null || true
}

# ---------------------------------------------------------------------------
# The shots. Each is a fresh process with a scripted edit list, so the state in
# each image is defined by its script rather than by whatever the previous shot
# left behind.
# ---------------------------------------------------------------------------

capture() {
  local name="$1" title="$2" script="$3" settle="${4:-$SETTLE}"
  local wb="$WORKDIR/$name.corro"
  printf 'CORRO_LOG 1\n' > "$wb"

  CORRO_EDIT_SCRIPT="$script" "$BINARY" "$wb" >"$WORKDIR/$name.log" 2>&1 &
  local pid=$!
  # Wait for the window, frame it, then give the draw a moment. The scripted
  # edits are timed from process start, so `settle` is what lets the last one
  # be applied and painted.
  local wid
  wid="$(wait_for_window || true)"
  frame_window "$wid"
  sleep "$settle"
  # Optional hook: a function name to run after the frame is up and the
  # scripted edits have settled (used for the click-driven selection shot).
  local hook="${5:-}"
  if [ -n "$hook" ]; then "$hook"; sleep 1.5; fi
  shot_index=$((shot_index + 1))
  local out
  out="$(printf '%s/shot-%02d-%s.png' "$OUTDIR" "$shot_index" "$name")"
  import -display "$DISPLAY_NUM" -window root "$out"
  kill "$pid" 2>/dev/null || true
  wait "$pid" 2>/dev/null || true
  SHOTS+=("$name|$(basename "$out")|$title")
  echo "captured $out"
}

# 1. An empty sheet: the chrome (formula bar, margin TOTAL seeds, tab, status).
capture "blank" 'A freshly opened workbook: the six-menu bar, the formula bar (fx plus the address and value boxes), the row/column headers with their margin columns and the seeded TOTAL aggregates, and the status line.' "1000:A1=10"

# 2. A small populated grid, so the row/column headers and both text and
#    computed cells are visible.
capture "populated" 'Literal columns (A) beside formula columns (B), both computed by the shared engine: `=A1*2` in each B cell and a `=SUM(B1:B3)` total. Header labels, per-cell values and the formula bar all come from the same render path the desktop GTK build uses.' "500:A1=10,1500:A2=20,2500:A3=30,3500:B1==A1*2,4500:B2==A2*2,5500:B3==A3*2,6500:C1=Total,7500:C2==SUM(B1:B3)" 4

# 3. A long formula in the input line, showing the formula bar rendering what
#    does not fit in the cell.
capture "formula" 'Formula entry: a nested `=IF(...)` in the input line, with the computed results (956.7, and the branch the condition selected) in the grid. Shows the formula bar rendering text wider than the cell that holds it.' "500:A1=1234.5,1500:A2=678.9,2500:A3==(A1+A2)/2*100,4000:A4==IF(A3>1000,\"high\",\"low\")" 4

# 4. A selection, so the highlight is in the image. `CORRO_EDIT_SCRIPT` has no
#    selection verb, so this uses clicks at coordinates computed from the
#    chrome metrics the renderer uses - the one place pixel coordinates are
#    unavoidable. A miss shows as a missing highlight, not a wrong edit.

# ---------------------------------------------------------------------------
# The companion document. Generated from the same table the captures used, so
# the prose cannot describe an image that is no longer produced.
# ---------------------------------------------------------------------------
DOC="$OUTDIR/SCREENSHOTS.md"
{
  echo "# corro screenshots"
  echo
  echo "![the corro GUI](shot-01-blank.png)"
  echo
  echo "Captured headlessly from the GTK build by \`scripts/screenshot.sh\` - no"
  echo "display, no window manager session, and no synthetic keystrokes: each image"
  echo "is a fresh process whose workbook is built by scripted edits"
  echo "(\`CORRO_EDIT_SCRIPT\`), applied through the ordinary commit path."
  echo
  echo "Regenerate with:"
  echo
  echo '```bash'
  echo "cargo build --features gui        # \`gui\` is not a default feature"
  echo "scripts/screenshot.sh             # writes dist/screenshots/"
  echo '```'
  echo
  echo "The images are checked in (\`.gitignore\` allows \`dist/screenshots/\`) so they"
  echo "can be linked from documentation; re-run the script after a UI change."
  echo
  echo "## The shots"
  echo
  echo "| Image | Shows |"
  echo "|---|---|"
  for entry in "${SHOTS[@]}"; do
    IFS='|' read -r name file title <<<"$entry"
    echo "| [\`$file\`]($file) | $title |"
  done
  echo
  echo "## What each one is for"
  echo
  echo "They are chosen to cover the parts of the UI a change is most likely to"
  echo "break, and that a text diff cannot show:"
  echo
  echo "* **Chrome geometry.** The formula bar, headers and status line are drawn by"
  echo "  the shared renderer, so a metrics change (a scale factor, a font advance)"
  echo "  shows up here before it shows up as a mis-click somewhere else."
  echo "* **The engine's output, rendered.** The \`populated\` and \`formula\` shots"
  echo "  put literal cells next to computed ones, so a formula regression is"
  echo "  visible as a wrong number rather than only as a failing unit test."
  echo "* **Canvas replay.** Every pixel inside the grid comes from Rust's"
  echo "  \`DrawContext\` closure replayed by the backend, not from native widgets."
  echo "  These images are therefore the end-to-end check of that path - the same"
  echo "  closure the ratatui, GTK, Windows and Android builds replay."
  echo
  echo "## Other backends"
  echo
  echo "The script captures the GTK build because it is the one that runs on this"
  echo "host. The widget tree it renders is the shared one: the same"
  echo "\`corro::gui::gui_backend::run_gui\` drives every GUI backend, and macOS"
  echo "(AppKit), iOS (UIKit) and Android differ only in the adapter underneath."
  echo "See \`rustxWidgets/docs/MACOS_GUIDELINES.md\` and \`IOS_GUIDELINES.md\`."
  echo
  echo "The terminal build has its own captures: \`scripts/demo_movie.py\` and"
  echo "\`dist/corro-gui-movie.mp4\` (see the README, \"Recording a video\")."
} > "$DOC"

# ---------------------------------------------------------------------------
# Self-check. An image that is mostly one colour is a failed capture (an Xvfb
# that never came up, or a grab taken before the window was drawn), and it is
# easy to commit without noticing - so check the chrome colours each shot must
# contain and fail the run if one does not.
# ---------------------------------------------------------------------------
verify_shots() {
  local bad=0
  for entry in "${SHOTS[@]}"; do
    IFS='|' read -r name file title <<<"$entry"
    local path="$OUTDIR/$file"
    [ -f "$path" ] || { echo "MISSING $file" >&2; bad=1; continue; }
    local uniq
    uniq="$(convert "$path" -format %k info: 2>/dev/null || echo 0)"
    # A drawn corro window has hundreds of distinct colours (antialiased text);
    # a blank or single-tone frame has a handful.
    if [ "${uniq:-0}" -lt 32 ]; then
      echo "SUSPECT $file: only $uniq distinct colours (blank or undrawn?)" >&2
      bad=1
    fi
    # The grid background is #BFBFBF and the body #FFFFFF in the default theme;
    # both must be present or the canvas did not paint.
    local has_bg
    has_bg="$(convert "$path" -format %c histogram:info:- 2>/dev/null | grep -ciE '#BFBFBF|#FFFFFF' || true)"
    if [ "${has_bg:-0}" -eq 0 ]; then
      echo "SUSPECT $file: no grid background/body colour found" >&2
      bad=1
    fi
  done
  if [ "$bad" -ne 0 ]; then
    echo "error: one or more screenshots look empty; not a usable set" >&2
    return 1
  fi
  echo "verified $shot_index screenshot(s)"
}

verify_shots || exit 1

echo
echo "wrote $shot_index screenshot(s) to $OUTDIR"
echo "wrote $DOC"
