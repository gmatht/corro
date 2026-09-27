//
//  CorroMacBridge.h
//  The C surface a macOS app calls into. Every function here is exported by
//  `macos/corro` (a `#[no_mangle] extern "C"` of the same name) and declared
//  in `rustxWidgets/docs/MACOS_GUIDELINES.md` §3.
//
//  This file is a *hand-written* declaration of an ABI that the Rust side
//  already owns. That is the same duplication the shim generator exists to
//  remove, and it is worth being precise about why it survives here: the
//  generator emits `app/CorroGeneratedShims.m/.h` (the pure forwarding shims
//  the Rust adapter messages to), but *this* header is the app's view of the
//  Rust exports. `macos/corro/scripts/check_bridge.sh` diffs the declarations
//  below against the `#[no_mangle]` list in `macos/corro/src/lib.rs`, so a
//  drift fails CI rather than crashing at runtime inside `objc_msgSend`.
//
//  A mismatched signature here is not a compile error — it is a runtime crash
//  with an unrelated-looking backtrace. Treat this file as part of the Rust
//  crate's API.
//

#ifndef CorroMacBridge_h
#define CorroMacBridge_h

#include <stdbool.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

// ---------------------------------------------------------------------------
// Bootstrap
// ---------------------------------------------------------------------------

/// Boot corro in this window. `root` is the window's `contentView` and
/// `window_controller` its `NSWindowController`; both must outlive the call.
/// Returns immediately — AppKit owns the run loop, so the host calls
/// `-[NSApp run]` and that never returns.
void corro_macos_root_ready(void *root, void *window_controller);

// ---------------------------------------------------------------------------
// Canvas (CorroSheetView)
// ---------------------------------------------------------------------------

/// Replay the draw closure for this canvas at `w`x`h` points. `ctx` is the
/// live `CGContextRef` from `drawRect:` and is valid only for the call.
void corro_macos_canvas_draw(uint64_t canvas_id, void *ctx, int32_t w, int32_t h);

/// Report the size the host has laid this canvas out to. Separate from the
/// draw callback on purpose — see `corro_macos_canvas_draw` in the Rust crate.
void corro_macos_canvas_size(uint64_t canvas_id, int32_t w, int32_t h);

/// Route a press to the canvas that owns it. Coordinates are points with a
/// top-left origin, which is why the canvas view overrides `isFlipped`.
void corro_macos_canvas_click(uint64_t canvas_id, double x, double y);

/// Button- and modifier-aware press (`button` is `NSMouseButton`: 0 left,
/// 1 right, 2 centre). Used by the sheet-tab context menu, which needs to
/// know the press was a right click.
void corro_macos_canvas_click_button(uint64_t canvas_id, double x, double y,
                                     uint32_t button, uint32_t mods);

/// A drag, and the end of one.
void corro_macos_canvas_motion(uint64_t canvas_id, double x, double y, uint32_t mods);
void corro_macos_canvas_release(uint64_t canvas_id, double x, double y,
                                uint32_t button, uint32_t mods);

/// Hardware-keyboard input. `keyval` is a Unicode scalar for printable keys
/// and an `rswidgets::core::key` constant otherwise; `mods` is
/// 1 = Shift, 4 = Ctrl, 8 = Alt. Returns true when the key was consumed, so
/// AppKit skips its own handling.
bool corro_macos_canvas_key(uint64_t canvas_id, uint32_t keyval, uint32_t mods);

/// The row height and column advance in points, so a scroll wheel can convert
/// its pixel delta to whole cells with the metrics the renderer used. Writes
/// `[row_h, col_w]`.
void corro_macos_canvas_cell_size(double *out);

/// Scroll by whole cells; writes the *applied* `[d_rows, d_cols]`, which is
/// smaller at a sheet edge — that is how the host knows to stop accumulating.
void corro_macos_canvas_scroll_by(int32_t d_rows, int32_t d_cols, int32_t *out);

/// Trackpad pinch. Returns the scale actually applied (clamped to 0.4..4.0).
double corro_macos_canvas_zoom(double factor);
double corro_macos_canvas_zoom_reset(void);

// ---------------------------------------------------------------------------
// Entry (formula bar)
// ---------------------------------------------------------------------------

/// Typing, Return, and focus. See `MACOS_GUIDELINES.md` §3 — these are the four
/// entry signals an `NSControlTextEditingDelegate` produces.
void corro_macos_entry_changed(uint64_t view_ptr);
void corro_macos_entry_activate(uint64_t view_ptr);
/// Returns the handler's own 0/1 answer, so the delegate can decide whether
/// to continue into AppKit's default behaviour.
int32_t corro_macos_entry_focus(uint64_t view_ptr, bool gained);

// ---------------------------------------------------------------------------
// Controls
// ---------------------------------------------------------------------------

/// Called from `CorroMacTarget.corroFired:` with a registered callback id.
void corro_macos_callback(uint64_t callback_id);

// ---------------------------------------------------------------------------
// Menu
// ---------------------------------------------------------------------------

/// Dispatch a chosen `NSMenuItem`. `name` is the NUL-terminated
/// `app.<action>` the model published. Unknown names are ignored.
void corro_macos_menu_action(const char *name);

/// Walk the menu model to build a real `NSMenu`:
/// `corro_macos_menu_count()`, then per menu `corro_macos_menu_item_count()`,
/// then per item `corro_macos_menu_item()`.
size_t corro_macos_menu_count(void);
size_t corro_macos_menu_item_count(size_t menu_idx);
/// Both out-pointers are NUL-terminated strings valid until the next call.
/// Returns false for an out-of-range index.
bool corro_macos_menu_item(size_t menu_idx, size_t item_idx,
                           const char **label_out, const char **action_out);

#ifdef __cplusplus
}
#endif

#endif /* CorroMacBridge_h */
