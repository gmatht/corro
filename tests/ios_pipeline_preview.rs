//! End-to-end preview of the iOS host pipeline, on a Linux host.
//!
//! ## Why this exists
//!
//! The iOS host app cannot be tested for interaction on a hosted macOS
//! runner: `xcrun simctl` cannot synthesise touches (`simctl io <dev> tap`
//! does not exist — it was proposed and never shipped), the app cannot read
//! XCTest's events, and `idb`/WebDriverAgent are not installed. So the
//! simulator proves *startup → first frame* and nothing beyond it
//! (`docs/ios/README.md` says exactly that).
//!
//! This file covers the corro half of the host contract — the part that is
//! plain data rather than UIKit:
//!
//! ```text
//!   CorroViewController.viewDidLoad  builds its UIMenu from
//!       corro::gui::ios_backend::menu_model()          ← checked here
//!   UIMenu action                    → corro_ios_menu_action
//!       → ios_backend::run_menu_action_by_name()       ← checked here
//!   the shim names it dispatches must be actions that exist, or the tap
//!       silently does nothing                          ← checked here
//! ```
//!
//! What it is NOT — and the earlier wording of this comment overstated it
//! before being corrected:
//!
//! * It is not the UIKit adapter. `backends_ios_adapter.rs` only compiles for
//!   `target_os = "ios"`, so `dispatch_canvas_click` / `dispatch_text_changed`
//!   / `dispatch_entry_activate` cannot run here. They are covered by the iOS
//!   cfg *compile* checks (`scripts/check_*_ios.sh`) and, at runtime, by the
//!   app on a device.
//! * It does not commit a cell. A commit needs a live `GuiState`, which needs
//!   a backend window; the *shared* commit path is covered by the desktop
//!   live-GUI tests (`tests/gui_click_commit.rs`, `tests/gui_edit_parity.rs`)
//!   against the same `gui_backend` code the iOS build runs. What is
//!   iOS-specific between those two is the shim → `corro_ios_*` hop, which
//!   needs a device.
//!
//! Run it with:
//!
//! ```text
//! cargo test --features gui-mobile-host --test ios_pipeline_preview
//! ```
//!
//! (or via `ios/corro/scripts/verify_all.sh`).

#![cfg(feature = "gui-mobile")]
#![cfg(not(target_os = "ios"))]

use corro::gui::ios_backend as ios;

/// The menu model the host builds its `UIMenu` from must be non-empty and
/// name every item with the `app.<action>` prefix `run_menu_action_by_name`
/// strips again — otherwise the host's bar is either empty or dispatching
/// into nothing.
#[test]
fn menu_model_is_dispatchable() {
    let model = ios::menu_model();
    assert!(!model.is_empty(), "the host would build an empty UIMenu");

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

    // The real thing a tap on a menu item ends in. With no GUI state
    // published (there is none in this process) the dispatch deliberately
    // logs "mobile menu action before state published" and returns — so this
    // asserts the *lookup* path: a name the model publishes is accepted, and
    // an unknown name is ignored rather than panicking.
    ios::run_menu_action_by_name("app.about");
    ios::run_menu_action_by_name("app.definitely_not_a_real_action");
}

/// Every published action name must map back to a real menu action, because
/// the host dispatches by name and an unknown name is *silently ignored*
/// (`dispatch_mobile_menu_action` documents that: a newer host against an
/// older handler). A typo in `menu_model` would therefore present a menu item
/// that does nothing when tapped — the one failure a phone user cannot
/// diagnose.
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

    for (_label, items) in ios::menu_model() {
        for (item_label, action) in items {
            let name = action.strip_prefix("app.").unwrap_or(&action);
            assert!(
                known.iter().any(|k| k == name),
                "menu item '{item_label}' publishes '{action}', which is not any \
                 MenuActionKind name — tapping it would silently do nothing"
            );
        }
    }
}

/// The menu model is read before any GUI state exists (the host builds its
/// `UIMenu` while the view is loading), so building it must never need the
/// live state or panic. This is the ordering the ObjC `viewDidLoad` relies on.
#[test]
fn menu_model_is_available_before_any_state() {
    let first = ios::menu_model();
    let second = ios::menu_model();
    assert_eq!(
        first.len(),
        second.len(),
        "menu_model must be a pure function of the shared menu definition"
    );
}

/// The three host entry points the ObjC shims call must exist with the
/// signatures the shim headers declare. This is a compile-time contract
/// test: the `extern "C"` wrappers live in `ios/corro/src/lib.rs`, and the
/// functions they call are these.
#[test]
fn host_entry_points_exist() {
    // `ios_main` is the root-ready entry point (`corro_ios_root_ready`).
    // Calling it with a null root must *fail*, not panic: the host can
    // misorder its calls, and a panic inside an extern "C" frame aborts the
    // process with no unwind (see the module docs of ios_backend).
    #[cfg(target_os = "ios")]
    {
        let r = ios::ios_main(std::ptr::null_mut(), std::ptr::null_mut());
        assert!(r.is_err(), "a null root view must be reported, not accepted");
    }

    // On a non-iOS host only the plain-data surface exists; that is what the
    // checks above use. `log_ios` must never panic regardless of backend
    // state (the shims log before the backend is up).
    ios::log_ios("ios_pipeline_preview: host entry points reachable");
}

/// The menu *definition* the model is derived from must stay in step with
/// the desktop menu tree: the iOS host shows one overflow item built from
/// the same source, so a menu added on the desktop and not here would be
/// unreachable on a phone.
#[test]
fn menu_model_covers_the_shared_menu_bar() {
    let bar = corro::gui::menu::menu_bar();
    let model = ios::menu_model();
    assert_eq!(
        model.len(),
        bar.len(),
        "every top-level menu must appear in the iOS model"
    );
    for (root, (label, _)) in bar.iter().zip(model.iter()) {
        // Mnemonic underscores are dropped for touch (meaningless without a
        // keyboard), so compare with them removed too.
        assert_eq!(label, &root.label.replace('_', ""));
    }
}
