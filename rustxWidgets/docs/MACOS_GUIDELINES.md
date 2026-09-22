# macOS guidelines (rswidgets)

What the AppKit backend requires from a host app, what the host must generate
or ship, and the rules the backend's own code follows. This is the sibling of
`IOS_GUIDELINES.md` — the two share most of their shape because they share
most of their *code*: both build on the common Apple runtime module
`src/backends/apple.rs`. Where a rule below differs from the iOS document, the
reason is AppKit vs UIKit, and it is called out.

The code is authoritative when this document drifts; `src/backends/macos.rs`,
`src/backends_macos_adapter.rs` and `src/backends/apple.rs` are the
implementation.

## 0. Targets, in one table

| Target | Rust target | Status |
|---|---|---|
| Intel Mac | `x86_64-apple-darwin` | tier-2; a plain `rustup target add` suffices |
| Apple silicon Mac | `aarch64-apple-darwin` | tier-2; the shipping configuration |

Both are 64-bit, so unlike iOS there is only one `CGFloat` shape (`f64`) to
worry about. There is no 32-bit macOS target in current Rust, and no
`-Zbuild-std` requirement — this is an ordinary tier-2 desktop target.

The check that proves the Rust half without a Mac is:

```text
rustxWidgets/scripts/check_rswidgets_macos.sh
```

It type-checks the AppKit paths for both macOS targets **and** re-checks an
iOS target, because both platforms compile the same `backends/apple.rs`: a
change to the shared module that breaks one platform must not pass on the
other.

## 1. The shared Apple module (what the two backends have in common)

`src/backends/apple.rs` is compiled for every `target_vendor = "apple"`
platform and names **no widget class**. It holds:

* the Objective-C runtime bindings (`objc_msgSend` cast per ABI shape,
  `objc_getClass`, `sel_registerName`, `objc_retain`);
* Foundation marshalling (`NSString` ↔ Rust `String` via
  `stringWithUTF8String:` / `UTF8String`);
* handle lifetime (`KEEP_ALIVE`), the `u64`-keyed callback registry, and the
  pointer-keyed `WidgetMeta` registry;
* `init_with_root_impl(root, vc, label)` — the shared bootstrap.

Both `backends::ios` and `backends::macos` re-export this module
(`pub use crate::backends::apple::*`), so `backends::ios::msg0` and
`backends::macos::msg0` are the same function. If you are adding a widget,
you are almost certainly editing the *adapter*, not this module. If you are
adding a runtime/Foundation helper, it goes here and both platforms get it.

## 2. What rswidgets does *not* do

As on iOS, there is **no resource generator**: a macOS app is defined by its
Xcode project (or `cargo-bundle`/`Info.plist`), its `Info.plist`, and its
icons, none of which the Rust side can sensibly synthesise. The host supplies:

* the **app bundle / Xcode project** — target, deployment version, bundle id,
  code signing;
* **`Info.plist`** — `NSPrincipalClass`, `NSMainNibFile` (if any), document
  types;
* the **ObjC shims** named in §3;
* the **icons** (an asset catalog).

What rswidgets owns is the whole Rust-side widget tree, the draw replay, and
the input dispatch on top of the shims.

## 3. The ObjC side the backend calls into

Each name is fixed: the Rust side resolves it with `objc_getClass` and sends
the listed selector. The `CorroMac*` names parallel the `CorroIos*` ones;
a host that already ships the iOS shims will find the shape familiar.

| Class | Contract |
|---|---|
| App delegate / window controller | On window load, calls the host's root-ready entry point, which ends in `rswidgets::backends::macos::init_with_root(content_view, window_controller)`. That is the whole bootstrap. |
| `CorroSheetView : NSView` | `corroSetCanvasId:` (called by `create_canvas`); **`isFlipped` must return `YES`** (see §5); `drawRect:` → `corro_macos_canvas_draw(canvas_id, ctx, w, h)`; `mouseDown:`/`mouseUp:` → `corro_macos_canvas_click(canvas_id, x, y)`; `keyDown:` → `corro_macos_canvas_key(canvas_id, keyval, mods)`. Without it, canvases are plain `NSView`s (the tree still builds, nothing draws). |
| `CorroMacTarget` | Class factory `targetWithCallbackId:` returning a target that, on its `corroFired:` action, calls `corro_macos_callback(callback_id)`. The `CorroIosTarget` equivalent: an `NSControl`'s `-target`/`-action` pair. |
| `CorroMacText` | `measure:font:size:slant:weight:` → a `malloc`ed `CGRect*` (or null) with the text width/height; `drawText:ctx:font:x:y:size:r:g:b:a:slant:weight:` draws into the live `CGContextRef`. See §5 for why text is delegated. |
| `CorroMacAlert` | `corroNewAlert` returns an alert object; `corroAddAction:` adds a button; `corroPresentDialog:` presents it (a standalone `runModal`, or `beginSheetModalForWindow:` when the host has a window). |
| Text-field delegate | `controlTextDidChange:` → `corro_macos_entry_changed(viewPtr)`; the Return path → `corro_macos_entry_activate(viewPtr)`; `controlTextDidBeginEditing:`/`controlTextDidEndEditing:` → `corro_macos_entry_focus(viewPtr, gained)`. These are the four entry signals. |
| Menu bar | Reads the model (a real `NSMenu` can be built from it — see §4) and dispatches a chosen item through `corro_macos_menu_action(name)`. |

**A missing shim degrades, never crashes.** Every resolution above returns
null when the class is absent, and the callers log once and become no-ops: the
widget tree still builds, the sheet just does not draw, or buttons are inert,
or dialogs never present.

## 4. The menu model (and why macOS is *not* iOS here)

iOS has no menubar, so its host must build a `UIMenu` from the model in §3.
macOS **does** have a real menubar, so a host can build a genuine `NSMenu`
tree — that is the natural macOS answer and is what a macOS port should do.

The adapter still exposes only the *model* (labels + `app.*` action names),
exactly like iOS and Android, because that is the one shape all three can
share; the host chooses how to render it. Selecting an item dispatches the
same `app.<name>` string the desktop menu registers, which routes through the
shared `dispatch_mobile_menu_action` table.

## 5. Drawing and text

* The draw model is **identical to iOS and Android**: the backend replays the
  registered `DrawContext` closure through `dispatch_draw`, which Core
  Graphics implements one primitive at a time (`fill_rect` →
  `CGContextFillRect`, `clip` → `CGContextClipToRect`, and so on). The closure
  is keyed by **canvas id**, not by view pointer.
* **Core Graphics is shared with iOS**, so this section's implementation is
  near-identical to the iOS one.
* **`isFlipped` is the one real trap.** AppKit's default view coordinate
  system has its origin at the *bottom left*; UIKit's is top-left. The shared
  drawing convention is top-left, so the host's `CorroSheetView` must override
  `isFlipped` to return `YES`. Without it the sheet renders upside down — a
  symptom that is obvious but easy to misdiagnose as a draw-context bug.
* **Text is delegated to the host shim.** `CorroMacText` resolves an `NSFont`
  from the family/size/weight and draws with `NSString`/`NSAttributedString`,
  so no font handling (or CoreText link) is duplicated in Rust. Any failure
  (missing shim, unrenderable string) falls back to the monospace estimate
  (`EST_CHAR_W = 7.2`, `EST_LINE_H = 1.2`), so layout never divides by zero.
* `drawText` takes a baseline on both Apple platforms: the shim offsets by the
  font's ascent so the shared top-left convention holds.

## 6. Layout, metrics and files

* **No expand flags.** `set_hexpand`/`set_vexpand`/`set_width_chars` are
  recorded on the handle (`WidgetMeta`) and honoured by `BoxWidget::append`:
  an expanding or canvas child is flagged flexible (`corroSetFlex:`), and a
  `set_width_chars(n)` request becomes a minimum width. This mirrors iOS's
  `UIStackView` distribution and Android's `LinearLayout` weight.
* **Density.** `NSScreen.backingScaleFactor` (1.0 / 2.0) is exposed as
  `display_density()`; corro multiplies its pixel metrics by it so chrome is
  legible on a Retina panel without a single desktop call site changing.
  **Unlike iOS there is no touch-target floor**: `MIN_TOUCH_PT = 44` is iOS
  only, and the AppKit adapter uses a compact desktop control height
  (`MIN_CTRL_PT = 24`) instead. macOS is a pointer platform.
* **Files.** Unlike iOS, macOS has real paths and a normal filesystem, so
  `notify` file watching works and open/save can use ordinary paths — the
  sandbox (if the app is sandboxed) is the host's concern, not the backend's.
  This is the one place where macOS is *closer to the desktop backends* than
  to iOS.

## 7. Threading and lifecycle

* **Everything runs on the main thread**: `drawRect:`, mouse/key events, the
  text-field delegate, and any Rust closure they call. The callback registries
  are `Mutex`-guarded and their `Send` impls are justified only by that plus
  single-threaded dispatch — never invoke them from another thread.
* Lock discipline is the shared rule: resolve the callback pointer under the
  lock, **drop the lock, then invoke**, because callbacks re-enter
  (`queue_redraw` from a draw closure) and holding the lock deadlocks.
* `MacosApp::run()` returns immediately. The host owns `NSApplication` and
  calls `-[NSApp run]`, which never returns, so the app *is* the event loop;
  corro leaks its `App` for the same reason it does on iOS and Android.
* `App::quit()` is a no-op on macOS too, by design: the host decides when to
  terminate (save, then `-[NSApp terminate:]`). It logs and returns.
* `applicationDidBecomeActive` / `…WillResignActive` / `…WillTerminate` are
  the host's to forward.

## 8. Handles and registries

Identical to iOS (`IOS_GUIDELINES.md` §8) because the code is the same module:
widgets are `#[repr(transparent)]` over a raw Objective-C object pointer,
retained by us (`objc_retain` on every handle we create) and listed in
`KEEP_ALIVE`. Never release a handle: widgets are shared across closures, and
a view removed from the hierarchy must not deallocate while a callback still
points at it.

Metadata (`WidgetMeta`: kind, text, expand flags, min chars, canvas id,
requested and laid-out size) is kept **beside** the objects, keyed by handle
pointer. Canvases are keyed by **canvas id** (`u64`, from 1); entries by view
pointer; button callbacks by `u64` id; menu actions by name. Two size maps
exist on purpose, exactly as on iOS and Android: `size_request` (what Rust
asked for, often a 1x1 placeholder) and `laid_out` (what the host reported
from `drawRect:`). The placeholder must never shrink a real laid-out size, or
the sheet viewport collapses to one row.

## 9. Why not Cacao / objc2?

The same reasoning as `IOS_GUIDELINES.md`, and worth stating explicitly since
it is the first question a macOS port raises:

* **Cacao is the wrong layer.** It is a standalone windowing/UI framework that
  owns its own app/run-loop/event model. Every rswidgets backend is a thin
  adapter under the `BackendApp` trait where *the host owns the loop*
  (iOS: `UIApplicationMain`; Android: the Activity; macOS: `NSApp`). Cacao
  would fight that, and it would reuse none of the existing `objc_msgSend`
  wrappers, registries, `WidgetMeta` or `KEEP_ALIVE` code.
* **Hand-rolled FFI is the house style.** The iOS backend already hand-rolls
  the ObjC runtime (deliberately, to keep the dependency list at zero on the
  old-SDK iOS 7 path); the Linux backend hand-rolls GTK through
  `gtk_dynamic_loader`; the Windows backend uses a *vendored* NWG.
* **`objc2` is the fallback, not the first choice.** It is more ergonomic, but
  it adds a dependency stack and, more importantly, would not be shared with
  the iOS backend (which is cfg-gated to keep its dependency-free shape). If
  the AppKit selector surface ever proves too large to hand-roll, `objc2` is
  the escape hatch — and it must be cfg-gated to `target_os = "macos"` only,
  so the iOS path keeps its current, old-SDK-capable shape.

## 10. Generating the shims (what is generated, and what is not)

The shims in §3 are not all hand-written. `rustxWidgets/rswidgets/src/apple_generator.rs`
holds the Rust↔ObjC **ABI contract as data** and emits
`app/CorroGeneratedShims.{h,m}`; the iOS host runs it from its own
`build.rs` (`--features generate-apple-shims`), exactly the pattern
`android/corro/build.rs` uses for `android_generator`.

**Why generate them at all.** The Rust side already names every selector and
its exact signature (in the adapters' `raw_send!` calls). Restating that by
hand is two independent declarations of one ABI, and a drift between them is
*not* a compile error — it is a mismatched `objc_msgSend`, the register-file
corruption §5 and `IOS_GUIDELINES.md` warn about. Generating the forwarding
shells removes that class of bug outright.

**Generated** (pure forwarding, no behaviour):

| Emitted | Was hand-written as |
|---|---|
| `Corro<Platform>Target` (+`targetWithCallbackId:`, -`corroFired:`) | the callback trampoline |
| `Corro<Platform>Alert` (+`corroNewAlert`, -`corroSetTitle:`, -`corroAddAction:`) over `UIAlertController`/`UIAlertView` (iOS) or `NSAlert` (macOS) | the dialog wrapper |
| `<base view> (CorroLayout)`: `corroSetSpacing:`, `corroSetFlex:`, `corroSetMinWidth:`, `corroSetCanvasId:`, `corroBoundsWidth/Height` | the layout category |
| the `corroPresentDialog:` category on `UIViewController`/`NSWindowController` | the presentation category |

**Also generated — but *configured***, because these have no single correct
body. Both are emitted by default and controlled by `ShimConfig`:

* **The canvas view** (`SheetView` / `CorroSheetView`), from
  `CanvasViewConfig`. The generated body covers the size report in both the
  layout and draw passes, the context handover, the click/key events, and — on
  AppKit only — the `isFlipped` override. The knobs exist because plausible
  hosts differ:

  | Field | Why it is a knob |
  |---|---|
  | `hardware_keys` | a headless or touch-only harness wants no first-responder plumbing |
  | `fill_superview` | the self-resizing fallback (and its re-entrancy guard) is only needed when the canvas is attached with `addSubview:`, not when a stack view sizes it |
  | `report_size_on_layout` | Rust replays the draw closure before the framework's first draw, so dropping the earlier report collapses the sheet to a 1x1 placeholder — kept as a knob so the tradeoff is explicit |
  | `trace_events` | the `fprintf` tracing the iOS app enabled while bringing itself up |
  | `deployment_target` | on iOS `UIKey` is 13.4+, so a lower target must guard the key path (on both the declaration and the definition — clang analyses each body separately) |

* **The text measurer** (`CorroIosText` / `CorroMacText`), from
  `TextShimConfig`: the font-fallback family, whether to emit the legacy API as
  a `respondsToSelector:` fallback, and `baseline_from_ascent` — the offset
  that converts the shared top-left convention to the baseline the text APIs
  want. That last one is a *named field* rather than a line in a `.m` body on
  purpose: without it every glyph in the sheet shifts up by the ascent, and a
  silent one-line deletion is exactly how that happens.

`emit_canvas` / `emit_text` turn either off, which is the **manual-override**
half: a host that already has a richer hand-written class keeps it, and the
generator still emits its **declarations** into the header. That host's `.m`
imports the generated header and *adopts* them, so a signature that drifts from
the Rust side is a compile error rather than a mismatched `objc_msgSend`.

`ios/corro` uses exactly that: `emit_canvas = false, emit_text = false` (its
canvas carries host-specific layout and its text class carries the SDK version
table), so its generated file is the 151-line forwarding set, while a fresh
host with no shims at all gets the whole 344-line canvas + text implementation
from the defaults.

A test in `apple_generator.rs` asserts that every `(class, selector)` pair the
adapters send — transcribed as `IOS_ADAPTER_SENDS` / `MACOS_ADAPTER_SENDS` —
is covered by the tables, so adding a `raw_send!` to an adapter without
registering it fails the suite rather than crashing at runtime.

That check runs from both ends:

* `apple_generator.rs`'s own tests assert the tables cover the transcribed
  `(class, selector)` lists.
* `ios/corro/scripts/check_selectors.sh` checks the **iOS** shim files really
  implement every selector the iOS adapter sends (it found a live crash:
  `create_box` sends `corroSetSpacing:`, which nothing implemented), and — since
  no macOS host exists yet — checks the **macOS** adapter's selectors are all
  *generatable*, which is what makes writing that host a "run the generator"
  step rather than a "reinvent the shims" step.
* `c7d73b01` is the case that justifies all of it: generating the declarations
  exposed `measure:` being declared with four arguments for a five-part
  selector, i.e. the ObjC side would have read an unset register. A hand-written
  pair of files cannot catch that without an SDK; a single signature table can.
