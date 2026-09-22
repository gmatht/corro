# iOS screenshots — what exists, and what does not

**A real iOS screenshot exists**, and this file was wrong for a while in saying
otherwise. The app *has* been run: `.github/workflows/ios.yml` builds it on a
hosted `macos-14` runner, boots an iPhone simulator, launches it, and
`screencapture`s the first frame. As of **run #97** (commit `7d64ee60`) the
whole job is green — build, boot, install, launch, **screenshot**, and the
health check that verifies the app drew.

Where to get it:

```
gh run download --repo gmatht/corro --name ios-simulator-screenshots
# or: Actions -> iOS -> a green run -> Artifacts -> ios-simulator-screenshots
```

That artifact (`ios-first-frame.png`, ~1.2 MB) is the image this directory used
to say did not exist. It needs an authenticated `gh`/browser session, which is
why it cannot be fetched and checked in from an unattended Linux checkout.

## What this directory contains

`desktop-tree-render.png` is the **fallback**: the same widget tree and draw
pipeline rendered by the desktop backend, so the sheet, the formula bar and the
chrome can be inspected offline (in a git clone, with no GitHub session). It is
useful for that, and it is *not* an iOS render — see the table below.

`../dist/screenshots/` holds a larger, regenerable set of the same kind,
produced by `../../../scripts/screenshot.sh` on a Linux host.

| File | What it is | What it proves | What it does NOT prove |
|---|---|---|---|
| `desktop-tree-render.png` | corro's real GUI (`cargo run --features gui -- --gui subtotal.corro`) captured from X on a Linux host | The sheet renders correctly through the shared `gui_backend` pipeline — grid lines, header bands, formula bar, the seeded sheet | Anything about iOS: not the iPhone metrics (`UIScreen.scale`), not the `UIStackView` layout, not Core Graphics text, not touch |
| `ios-simulator-screenshots` (CI artifact, not checked in) | `xcrun simctl io … screenshot` on a booted iPhone simulator, from `ios.yml` | The iOS build is real and the first frame draws: Rust startup + `SheetView.drawRect` at a live CGContext, observed at real sizes (375x616, 375x667) | Interaction: it is one frame, and the touch/keyboard paths still need the manual `simctl` taps below |

A larger, regenerable set of the same kind of image lives in
[`dist/screenshots/`](../../dist/screenshots/SCREENSHOTS.md) (blank workbook,
populated grid, formula entry), produced by `scripts/screenshot.sh`. None of
them is an iOS screenshot, for the reasons below.

## Why the iOS render cannot be produced *in this checkout*

1. **The ObjC does not compile here** — it needs `UIKit/UIKit.h` and
   `Foundation/Foundation.h`, which this Linux host has nowhere (§8b of the
   guidelines). The CI run compiles it on macOS, which is the whole point of
   that workflow.
2. **No simulator** without Xcode, so there is no `simctl` to screenshot. The
   simulator runs on the `macos-14` runner, not here.
3. The `ios-ui` example deliberately **exits before the event loop** (it is a CI
   check of the tree's *structure*, and a headless runner has no display), so it
   has no window to photograph.

So the split is: **the image is produced by CI** (which is real, and green), and
**this clone keeps a desktop-rendered stand-in** for offline inspection. Neither
one is a substitute for the other.

## How to get the real ones

**Without a Mac of your own:** this is already working.
`.github/workflows/ios.yml` builds on a hosted `macos-14` runner, boots a
simulator, launches the app, screenshots the first frame and uploads the PNG as
an artifact — **verified green on run #97** (`7d64ee60`), with every step
passing including the screenshot and the health check. Earlier runs are in the
Actions tab (97 of them at the time of writing); the failures along the way were
mostly the workflow's own probes, not the app (see the workflow header).

That runner is the *only* route to a real iOS screenshot from a non-Mac host,
and it is a cloud build rather than a cloud *capture*: the simulator must run
where the SDK is. Note also that a device farm is not an alternative here —
LambdaTest/BrowserStack want a **signed `.ipa`**, which needs a build first,
so they are a smoke test after this workflow succeeds, not a way around it. (A real-device farm
cannot do this — LambdaTest App Live wants a signed `.ipa`, so it needs a build
first. See the guidelines' table of the four kinds of "online simulator".)

**On a macOS host**, once `./build_ios.sh sim --run` works:

```sh
xcrun simctl io booted screenshot ios-12-sheet.png     # the sheet
xcrun simctl io booted tap X Y                          # then tap a cell
xcrun simctl io booted screenshot ios-12-tapped.png
```

Worth capturing, because each exercises a different part of the port: the first
frame (Core Graphics draw + the `CorroIosText` shim), a tap (touch → canvas
click → cursor move), and typing in the formula bar (soft-keyboard text →
`corro_ios_entry_changed` → commit). Those three are the paths the checklist in
`IOS_GUIDELINES.md` §9 names as still unverified.

## Reproducing the desktop render

`desktop-tree-render.png` was produced with the ad-hoc recipe below (a
throwaway display, a 45-second sleep, a window id looked up by hand):

```sh
DISPLAY=:99 cargo run --features gui -- --gui subtotal.corro &
sleep 45
DISPLAY=:99 import -window <the "corro 0.7.0" window id> shot.png
```

**Superseded**: use `scripts/screenshot.sh` instead. It does the same thing
reproducibly — Xvfb on a private display, waits for the window to be *mapped and
framed* rather than sleeping a fixed 45 seconds, builds each workbook with
`CORRO_EDIT_SCRIPT` (the ordinary commit path, so no synthetic keystrokes are
involved) and self-checks each capture so a blank frame fails the run. Its
output is checked in under `dist/screenshots/` with
[`SCREENSHOTS.md`](../../dist/screenshots/SCREENSHOTS.md) describing it.

The reason this image has any diagnostic value at all is that the `--gui` build
renders through the GTK backend while iOS runs the identical `gui_backend` code
path through the UIKit adapter. That shared pipeline is also the reason it is
still only a partial substitute — but it means the desktop shots above are
worth regenerating whenever the shared renderer changes, which the script makes
cheap.
