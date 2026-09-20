//! Build corro's rswidgets widget tree the way the Android GUI path does.
//!
//! On Android the widget tree is rooted at the Activity's content view: the
//! `MainActivity` JNI export calls `corro::gui::android_backend::android_main`,
//! which runs the shared `gui_backend::run_gui` pipeline against that root
//! layout (`LinearLayout` weights stand in for GTK's expand flags; see
//! `rustxWidgets/docs/ANDROID_GUIDELINES.md` §6).
//!
//! This example is the same tree without Android: it stands in for the
//! Activity's root layout with the backend's root, builds the identical
//! children in the identical order, and exits before the event loop — so a
//! host build can verify the UI structure the APK will show, and so the
//! widget construction stays exercised on a desktop where it is easy to
//! debug.
//!
//! Run it with the GUI feature:
//!
//! ```text
//! cargo run --example android-ui --features gui
//! ```
//!
//! The Android *resources* (Material 3 theme, palette, adaptive launcher
//! icon, manifest) are generated separately, by the resource generator the
//! `corro_android` crate's `build.rs` calls:
//! `rswidgets::android_generator::run()` writing into
//! `android/corro` (`RSWIDGETS_ANDROID_PROJECT` overrides the root).

use rswidgets::common::{Canvas, Entry, Label, MenuBar, Orientation, WidgetBox, Window};
use rswidgets::App as RswidgetsApp;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // `App::init` is the backend-agnostic entry point: GTK on Linux/Windows,
    // the nwg adapter on Win9x, and `backends::android::init_with_layout`
    // on Android. The Android path is the same call, only the root layout
    // differs (the Activity's content view instead of a toplevel window).
    let rxapp = RswidgetsApp::init().map_err(|e| format!("rswidgets init failed: {e}"))?;

    // Stand-in for the Activity's root layout. On Android this is the
    // `LinearLayout` passed to `nativeInit`; `set_child_box` is how GTK
    // attaches the same subtree to a toplevel.
    let win: Window = rxapp.new_window()?;
    win.set_title("corro (Android UI preview)");

    // Vertical stack: menu bar, formula bar, sheet, tab strip, status line.
    let vbox: WidgetBox = rxapp.new_box(Orientation::Vertical, 0)?;

    // --- menu bar (same shared definition as GTK/ratatui/pancurses) --------
    let action_group = rxapp.ensure_action_group()?;
    let bar = corro::gui::menu::menu_bar();
    let mut menubar_model = rxapp.new_menu()?;
    for root in &bar {
        let sub = corro::gui::menu::build_common_menu(
            &rxapp,
            root.submenu.as_deref().unwrap_or(&[]),
            "app",
        )?;
        menubar_model.append_submenu(
            &corro::gui::menu::mnemonic_label(root.label, root.shortcut),
            &sub,
        );
    }
    let menubar: MenuBar = rxapp.new_menubar(&menubar_model, action_group)?;
    vbox.append(&menubar);

    // --- formula bar: address, "fx" hint, expanding entry, status ----------
    let formula_bar: WidgetBox = rxapp.new_box(Orientation::Horizontal, 2)?;
    let addr_label: Label = rxapp.new_label("A1")?;
    let f_label: Label = rxapp.new_label("  fx  ")?;
    let formula_entry: Entry = rxapp.new_entry()?;
    // Android has no expand flags: the backend records this and BoxWidget
    // turns it into a LinearLayout weight plus a minimum width (an empty
    // EditText measures zero), see ANDROID_GUIDELINES.md §6.
    formula_entry.set_hexpand(true);
    formula_bar.append(&addr_label);
    formula_bar.append(&f_label);
    formula_bar.append(&formula_entry);
    formula_bar.set_child_hexpand(&formula_entry, true);
    let formula_status: Label = rxapp.new_label("")?;
    formula_status.set_visible(false);
    formula_bar.append(&formula_status);

    // --- sheet canvas: the grid itself (SheetView on Android) -------------
    // Android renders this canvas through `SheetView.onDraw` ->
    // `nativeOnDraw` -> `backends_android_adapter::dispatch_draw`, i.e. the
    // same draw callback a desktop backend attaches here.
    let canvas: Canvas = rxapp.new_canvas()?;
    canvas.set_size_request(1, 1);
    canvas.set_can_focus(true);

    // --- sheet tab strip (below the grid) and status line ------------------
    let tabbar: Canvas = rxapp.new_canvas()?;
    tabbar.set_size_request(1, 22);
    tabbar.set_visible(false);
    let hints_label: Label = rxapp.new_label("Ready")?;

    vbox.append(&formula_bar);
    vbox.append(&canvas);
    // nwg looks the child up by handle after append; a no-op elsewhere.
    vbox.set_child_vexpand(&canvas, true);
    vbox.append(&tabbar);
    vbox.append(&hints_label);

    // Attach the tree and present it, then leave. Presenting here is what the
    // Android path does implicitly when the Activity sets its content view;
    // we do NOT enter the event loop — the point is the structure, and on a
    // headless machine (CI, container) there is no display to run it on.
    win.set_child_box(&vbox);
    win.present();

    let menu_roots: Vec<&str> = bar.iter().map(|m| m.label).collect();
    println!("corro Android UI tree");
    println!("  root        : vertical box (window on desktop, Activity layout on Android)");
    println!("  menu bar    : {} top-level menus {:?}", bar.len(), menu_roots);
    println!("  formula bar : label \"A1\" + label \"fx\" + expanding entry + status label");
    println!("  sheet       : canvas (SheetView / nativeOnDraw), expanding");
    println!("  tab strip   : canvas, hidden until the workbook has 2+ sheets");
    println!("  status line : label \"Ready\"");
    println!();
    println!("Android resources for this tree are generated by");
    println!("rswidgets::android_generator — see android/corro/build.rs.");
    Ok(())
}
