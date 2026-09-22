# iOS screenshots — what exists, and what does not

**There is no screenshot of the iOS 12+ UI.** The app has not been run: its
Objective-C has never been compiled, because that needs macOS and an Apple SDK
(`../rustxWidgets/docs/IOS_GUIDELINES.md` §3, §8b). Any image claiming to be
"the iOS UI" would be fabricated.

What this directory does contain is the closest honestly-obtainable artefact:
the **same widget tree and the same draw pipeline** rendered by the desktop
backend, so the sheet, the formula bar and the chrome can be inspected rather
than taken on trust.

| File | What it is | What it proves | What it does NOT prove |
|---|---|---|---|
| `desktop-tree-render.png` | corro's real GUI (`cargo run --features gui -- --gui subtotal.corro`) captured from X on a Linux host | The sheet renders correctly through the shared `gui_backend` pipeline — grid lines, header bands, formula bar, the seeded sheet | Anything about iOS: not the iPhone metrics (`UIScreen.scale`), not the `UIStackView` layout, not Core Graphics text, not touch |

A larger, regenerable set of the same kind of image lives in
[`dist/screenshots/`](../../dist/screenshots/SCREENSHOTS.md) (blank workbook,
populated grid, formula entry), produced by `scripts/screenshot.sh`. None of
them is an iOS screenshot, for the reasons below.

## Why the iOS render cannot be produced here

1. **The ObjC never compiles** without `UIKit/UIKit.h` and `Foundation/Foundation.h`
   — Zig does not ship them, and neither does this host (§8b of the guidelines).
2. **No simulator** without Xcode, so there is no `simctl` to screenshot.
3. The `ios-ui` example deliberately **exits before the event loop** (it is a CI
   check of the tree's *structure*, and a headless runner has no display), so it
   has no window to photograph.

## How to get the real ones

**Without a Mac of your own:** `.github/workflows/ios.yml` is written to do
this — builds on a hosted `macos-14` runner, boots a simulator, launches the
app, screenshots the first frame and uploads the PNG as a build artifact. Be
aware it is **untested**: it was authored without Xcode or GitHub access, so
only its syntax is verified. The first run may need fixes.

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
