//
//  CorroBridge.h
//  corro on iOS — the C surface between the app target and the Rust cdylib.
//
//  Declares the `extern "C"` entry points `ios/corro/src/lib.rs` exports.
//  Keeping them in one header means the shims cannot drift from the Rust side
//  silently: a missing or misspelled export shows up as a link error rather
//  than as a runtime no-op.
//
//  See rustxWidgets/docs/IOS_GUIDELINES.md §3 for the contract table.
//

#ifndef CORRO_BRIDGE_H
#define CORRO_BRIDGE_H

#include <stdbool.h>
#include <stddef.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

/// Boot: hand the backend the root view and its view controller, then run the
/// shared GUI pipeline. Returns immediately (UIKit owns the event loop).
void corro_ios_root_ready(void *root, void *view_controller);

/// SheetView draw: replay the Rust draw closure against the live CGContext.
void corro_ios_canvas_draw(uint64_t canvas_id, void *ctx, int32_t w, int32_t h);

/// SheetView tap: move the cursor to the tapped cell.
void corro_ios_canvas_click(uint64_t canvas_id, double x, double y);

/// SheetView hardware key: returns whether the key was consumed.
bool corro_ios_canvas_key(uint64_t canvas_id, uint32_t keyval, uint32_t mods);

/// Formula-field text changed (soft keyboard is the only source) — the signal
/// corro's edit state adopts.
void corro_ios_entry_changed(uint64_t view_ptr);

/// Formula-field IME Done/Return: commit and move down.
void corro_ios_entry_activate(uint64_t view_ptr);

/// Formula-field begin/end editing; returns whether the event was consumed.
int32_t corro_ios_entry_focus(uint64_t view_ptr, bool gained);

/// Generic control target action (buttons, switches, pickers).
void corro_ios_callback(uint64_t callback_id);

/// Menu item chosen in the host's UIMenu / action sheet (`app.<action>`).
void corro_ios_menu_action(const char *name);

/// Menu model walk: counts, then per-item label/action strings.
size_t corro_ios_menu_count(void);
size_t corro_ios_menu_item_count(size_t menu_idx);
bool corro_ios_menu_item(size_t menu_idx, size_t item_idx,
                         const char **label_out, const char **action_out);

#ifdef __cplusplus
}
#endif

#endif /* CORRO_BRIDGE_H */
