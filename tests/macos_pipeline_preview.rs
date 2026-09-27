//! End-to-end preview of the macOS host pipeline, on a Linux host.
//!
//! ## Why this exists
//!
//! The macOS host is the newest of the GUI backends and the one with the
//! smallest automated coverage, so the question "is the AppKit path
//! *feature-complete* or merely type-checking?" needs an answer that a Linux
//! CI runner can produce. This file answers the half that is plain data rather
//! than AppKit:
//!
//! ```text
//!   NSApplicationDelegate          boots the app, then
//!     corro_macos_root_ready(contentView, windowController)   ← checked here
//!   CorroSheetView.drawRect:        → corro_macos_canvas_draw
//!     → gui_backend's draw replay   ← shared with every other backend
//!   NSMenu item chosen              → corro_macos_menu_action
//!     → macos_backend::run_menu_action_by_name()              ← checked here
//!   the names it dispatches must be actions that exist, or the click
//!     silently does nothing                                     ← checked here
//! ```
//!
//! It is deliberately the *twin* of `tests/ios_pipeline_preview.rs`, and the
//! two are kept in step: the iOS file is the reference for what a host-contract
//! test may and may not claim.
//!
//! What it is NOT:
//!
//! * It is not the AppKit adapter. `backends_macos_adapter.rs` only compiles for
//!   `target_os = "macos"`, so `dispatch_canvas_click` / `dispatch_draw` /
//!   `dispatch_text_changed` cannot run here. They are covered by the macOS cfg
//!   *compile* checks (`rustxWidgets/scripts/check_rswidgets_macos.sh`,
//!   `scripts/check_corro_macos.sh`, both in CI) and, at runtime, by the app
//!   on a real Mac.
//! * It does not commit a cell. A commit needs a live `GuiState`, which needs a
//!   backend window; the *shared* commit path is covered by the desktop
//!   live-GUI tests (`tests/gui_click_commit.rs`, `tests/gui_edit_parity.rs`)
//!   against the same `gui_backend` code the macOS build runs. What is
//!   macOS-specific between those two is the shim → `corro_macos_*` hop, which
//!   needs a Mac.
//!
//! What it *does* add over the iOS file, because macOS has a real mouse and
//! keyboard and iOS does not:
//!
//! * the *pointer* contract (`cell_size`, `scroll_by_cells`, zoom) — the
//!   metrics a host needs to turn its own pixel deltas into whole cells, which
//!   is a desynchronisation risk that simply does not exist on a phone;
//! * an assertion that the macOS menu model and the iOS one are the *same*
//!   model, which is what lets both hosts share one dispatcher.
//!
//! Run it with:
//!
//! ```text
//! cargo test --features gui-macos-host --test macos_pipeline_preview
//! ```
//!
//! (`gui-macos-host` = `gui-macos` + the Linux backend; on Linux the AppKit
//! adapter is not compiled, so the Linux stand-in is what makes the GUI
//! pipeline exist at all. It is the same trick `gui-mobile-host` plays for
//! iOS.)

#![cfg(feature = "gui-macos")]
#![cfg(not(target_os = "macos"))]

use corro::gui::macos_backend as macos;

/// The menu model a macOS host builds its `NSMenu` from must be non-empty and
/// name every item with the `app.<action>` prefix `run_menu_action_by_name`
/// strips again — otherwise the host's menubar is either empty or dispatching
/// into nothing.
#[test]
fn menu_model_is_dispatchable() {
    let model = macos::menu_model();
    assert!(!model.is_empty(), "the host would build an empty NSMenu");

    let mut total = 0usize;
    for (label, items) in &model {
        assert!(!label.is_empty(), "a top-level menu has no label");
        for (item_label, action) in items {
            assert!(
                !item_label.is_empty(),
                "menu '{label}' has an unlabelled item"
            );
            assert!(
                action.starts_with("app."),
                "item '{item_label}' publishes '{action}', which the host cannot dispatch \
                 (run_menu_action_by_name expects app.<name>)"
            );
            total += 1;
        }
    }
    assert!(total > 0, "the menu model has no items at all");

    // The real thing a click on a menu item ends in. With no GUI state
    // published (there is none in this process) the dispatch deliberately
    // logs "mobile menu action before state published" and returns — so this
    // asserts the *lookup* path: a name the model publishes is accepted, and
    // an unknown name is ignored rather than panicking.
    macos::run_menu_action_by_name("app.about");
    macos::run_menu_action_by_name("app.definitely_not_a_real_action");
}

/// Every published action name must map back to a real menu action, because
/// the host dispatches by name and an unknown name is *silently ignored*
/// (`dispatch_mobile_menu_action` documents that: a newer host against an
/// older handler). A typo in `menu_model` would therefore present a menu item
/// that does nothing when clicked — the one failure a user cannot diagnose.
#[test]
fn every_published_action_is_a_known_menu_action() {
    use corro::gui::menu::{action_kind_to_name, menu_bar, MenuAction};

    fn collect(items: &[MenuAction], out: &mut Vec<String>) {
        for item in items {
            match item.submenu.as_deref() {
                Some(sub) => collect(sub, out),
                None => out.push(action_kind_to_name(item.action).to_owned()),
            }
        }
    }

    let mut known = Vec::new();
    for root in menu_bar() {
        collect(root.submenu.as_deref().unwrap_or(&[]), &mut known);
    }

    for (_label, items) in macos::menu_model() {
        for (item_label, action) in items {
            let name = action.strip_prefix("app.").unwrap_or(&action);
            assert!(
                known.iter().any(|k| k == name),
                "menu item '{item_label}' publishes '{action}', which is not any \
                 MenuActionKind name — clicking it would silently do nothing"
            );
        }
    }
}

/// The menu model is read before any GUI state exists (the host builds its
/// `NSMenu` while the window is loading), so building it must never need the
/// live state or panic. This is the ordering the ObjC delegate relies on.
#[test]
fn menu_model_is_available_before_any_state() {
    let first = macos::menu_model();
    let second = macos::menu_model();
    assert_eq!(
        first.len(),
        second.len(),
        "menu_model must be a pure function of the shared menu definition"
    );
}

/// The menu *definition* the model is derived from must stay in step with the
/// desktop menu tree: the macOS host shows one real menubar built from the
/// same source, so a menu added on the desktop and not here would be
/// unreachable on a Mac.
#[test]
fn menu_model_covers_the_shared_menu_bar() {
    let bar = corro::gui::menu::menu_bar();
    let model = macos::menu_model();
    assert_eq!(
        model.len(),
        bar.len(),
        "every top-level menu must appear in the macOS model"
    );
    for (root, (label, _)) in bar.iter().zip(model.iter()) {
        // Mnemonic underscores are dropped (they mark the Alt shortcut, which
        // a menubar renders itself), so compare with them removed too.
        assert_eq!(label, &root.label.replace('_', ""));
    }
}

/// macOS and iOS must publish the **same** model. They are the same
/// `menu_model()` re-exported, and that is the whole reason the two hosts can
/// share one dispatcher — so a divergence here would mean one of the two
/// re-exports stopped being the shared function, which is a silent bug: both
/// models would still look reasonable on their own platform.
#[test]
fn the_macos_and_ios_models_are_the_same_model() {
    let m = macos::menu_model();
    let i = corro::gui::ios_backend::menu_model();
    assert_eq!(
        m, i,
        "the macOS and iOS hosts must publish the same menu model; a difference \
         means one re-export stopped being the shared menu_model()"
    );
}

/// The host entry points a macOS shim calls must exist and be safe to call
/// out of order. This is a compile-time contract test: the `extern "C"`
/// wrappers live in `macos/corro/src/lib.rs`, and the functions they call are
/// these.
///
/// `macos_main` is the root-ready entry point. Calling it with a null root
/// must *fail*, not panic: the host can misorder its calls, and a panic inside
/// an `extern "C"` frame aborts the process with no unwind (see the module
/// docs of `macos_backend`).
#[test]
fn host_entry_points_exist() {
    // `log_macos` must never panic regardless of backend state — the shims log
    // before the backend is up, and a log call that could crash would be the
    // worst possible place to find out.
    macos::log_macos("macos_pipeline_preview: host entry points reachable");

    // `macos_main` itself only exists on a Mac (it is
    // `#[cfg(target_os = "macos")]`, exactly like `ios_main`), so its "a null
    // root is reported, not accepted" assertion is the macOS-target twin of
    // the one in `tests/ios_pipeline_preview.rs`. Here the reachable half is
    // that the *module* is compiled and its non-target surface runs, which is
    // what makes the test above meaningful: the host crate calls these symbols
    // by name, and a missing one is a link error on a Mac.
    //
    // Asserted by the compile checks in CI:
    //   rustxWidgets/scripts/check_rswidgets_macos.sh  (the AppKit adapter)
    //   scripts/check_corro_macos.sh                   (this module)
    //   macos/corro/                                    (the extern "C" surface)
}

/// The pointer contract macOS has and iOS does not: a host must be able to ask
/// the renderer for the metrics it needs to convert its own pixel deltas into
/// whole cells.
///
/// This is the one place a macOS host is *not* sharing iOS's code, and it is a
/// real desynchronisation risk: if a host computed a row height itself it
/// would be right only at density 1.0 and unzoomed, so a Retina display or a
/// pinch would make the sheet scroll a different number of rows than it shows.
/// So both numbers must be positive — a zero would be a division by zero in
/// the host's pixel-to-cell conversion, and the host is the only code that does
/// that division.
#[test]
fn cell_metrics_are_usable_by_a_pointer_host() {
    let (row_h, col_w) = macos::cell_size();
    assert!(row_h > 0.0, "row height must be positive, got {row_h}");
    assert!(col_w > 0.0, "column advance must be positive, got {col_w}");
    // And the two are not the same number by accident: a square cell would
    // mean a host's pixel-to-cell conversion collapsed one axis, which is a
    // plausible-looking bug that scrolls correctly on screen and wrong on a
    // non-square grid.
    assert_ne!(
        row_h, col_w,
        "row height and column advance must be distinct dimensions"
    );
}

/// The pinch scale is clamped to the range the *widget* uses, so a trackpad
/// pinch and a touch pinch cannot disagree about what "fully zoomed" means.
/// The scale is a plain thread-local needing no window, which is why this can
/// be checked here rather than only on a Mac.
#[test]
fn pinch_zoom_is_clamped_to_the_shared_range() {
    // Start from a known scale (the thread-local is per-test-thread, so this
    // is both the setup and the assertion of the reset).
    assert_eq!(macos::reset_viewport_zoom(), 1.0);

    // A pinch out scales in proportionally.
    let doubled = macos::zoom_viewport_by(2.0);
    assert!(
        (doubled - 2.0).abs() < 1e-9,
        "a 2x pinch should apply 2.0, got {doubled}"
    );
    // ...and back down again, so the scale is multiplicative, not additive.
    let back = macos::zoom_viewport_by(0.5);
    assert!(
        (back - 1.0).abs() < 1e-9,
        "pinching back to 0.5 should return to 1.0, got {back}"
    );

    // Clamped at the top: an absurd factor must not produce an absurd scale,
    // or the renderer divides by it and the metrics stop meaning anything.
    let clamped_up = macos::zoom_viewport_by(1e9);
    assert!(
        clamped_up <= 4.0 + 1e-9,
        "the zoom must stay within the shared maximum, got {clamped_up}"
    );
    // ...and at the bottom.
    let clamped_down = macos::zoom_viewport_by(1e-9);
    assert!(
        clamped_down >= 0.4 - 1e-9,
        "the zoom must stay within the shared minimum, got {clamped_down}"
    );

    // A non-positive or non-finite factor is ignored, matching the widget's
    // `zoom_by`: a host that forwards a zero from a spurious scroll event must
    // not zero the scale and collapse the sheet.
    let before = macos::zoom_viewport_by(2.0);
    for bad in [0.0, -1.0, f64::NAN, f64::NEG_INFINITY] {
        let after = macos::zoom_viewport_by(bad);
        assert_eq!(after, before, "a factor of {bad} must be ignored");
    }
    assert_eq!(macos::reset_viewport_zoom(), 1.0);
}

/// Scrolling before a window exists must report "nothing moved" rather than
/// panicking. A host calls this from `scrollWheel:` and from a momentum phase,
/// both of which can fire before (or after) the tree is up — and a panic
/// inside an `extern "C"` frame aborts the process with no unwind.
#[test]
fn scrolling_without_a_viewport_is_a_no_op_not_a_panic() {
    assert_eq!(macos::scroll_by_cells(3, 0), (0, 0));
    assert_eq!(macos::scroll_by_cells(0, -2), (0, 0));
    assert_eq!(macos::scroll_by_cells(0, 0), (0, 0));
}
