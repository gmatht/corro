//! Build corro's rswidgets widget tree the way the iOS GUI path does.
//!
//! On iOS the widget tree is rooted at the app's root `UIView`: the
//! `ViewController` calls the `corro_ios` cdylib's
//! `corro_ios_root_ready(root, view_controller)` entry point, which runs the
//! shared `gui_backend::run_gui` pipeline against that root (a `UIStackView`
//! distribution plus recorded expand flags stand in for GTK's expand
//! behaviour; see `rustxWidgets/docs/IOS_GUIDELINES.md` §6).
//!
//! This example is the same tree without iOS: it stands in for the app's root
//! view with the backend's root, builds the identical children in the
//! identical order, and exits before the event loop — so a host build can
//! verify the UI structure the `.app` will show, and so the widget
//! construction stays exercised on a desktop where it is easy to debug.
//!
//! Run it with the GUI feature:
//!
//! ```text
//! cargo run --example ios-ui --features gui
//! ```
//!
//! What is *not* here: the iOS-specific shims (`SheetView`, the text-field
//! delegate, the menu bar). Those are Swift/ObjC in `ios/corro/app`, because
//! they must exist inside the app bundle for UIKit to call them; this example
//! exercises the Rust half — the widget tree, the menu *model* and the
//! metrics — which is the part that regresses silently.

use rswidgets::common::{Canvas, Entry, Label, MenuBar, Orientation, WidgetBox, Window};
use rswidgets::App as RswidgetsApp;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // `App::init` is the backend-agnostic entry point: GTK on Linux, the nwg
    // adapter on Windows, `backends::android::init_with_layout` on Android
    // and `backends::ios::init_with_root` on iOS. The iOS path is the same
    // call, only the root view differs (the app's `UIView` instead of a
    // toplevel window).
    let rxapp = RswidgetsApp::init().map_err(|e| format!("rswidgets init failed: {e}"))?;

    // Stand-in for the app's root view. On iOS this is the view the host
    // passed to `corro_ios_root_ready`; `set_child_box` is how GTK attaches
    // the same subtree to a toplevel.
    let win: Window = rxapp.new_window()?;
    win.set_title("corro (iOS UI preview)");

    // Vertical stack: menu bar, formula bar, sheet, tab strip, status line.
    let vbox: WidgetBox = rxapp.new_box(Orientation::Vertical, 0)?;

    // --- menu bar (same shared definition as every other backend) ----------
    //
    // On iOS this model is what the host reads to build its `UIMenu`: the
    // adapter's `create_menubar` has no view behind it (a phone shows one
    // "⋯" item, not six text menus), so `corro::gui::ios_backend::menu_model`
    // flattens it and the Swift side renders it. Building it here proves the
    // model is well-formed.
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
    // iOS has no expand flags either: the backend records this and
    // `BoxWidget::append` turns it into a flexible `UIStackView` child plus a
    // minimum width (an empty `UITextField` measures zero), see
    // IOS_GUIDELINES.md §6.
    formula_entry.set_hexpand(true);
    formula_bar.append(&addr_label);
    formula_bar.append(&f_label);
    formula_bar.append(&formula_entry);
    formula_bar.set_child_hexpand(&formula_entry, true);
    let formula_status: Label = rxapp.new_label("")?;
    formula_status.set_visible(false);
    formula_bar.append(&formula_status);

    // --- sheet canvas: the grid itself (SheetView on iOS) ------------------
    // iOS renders this canvas through `SheetView.drawRect:` ->
    // `corro_ios_canvas_draw` -> `backends_ios_adapter::dispatch_draw`, i.e.
    // the same draw callback a desktop backend attaches here.
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
    // A vertical stack distributes along its axis; on the manual-layout shim
    // this is the recorded flex flag (no-op repeat elsewhere).
    vbox.set_child_vexpand(&canvas, true);
    vbox.append(&tabbar);
    vbox.append(&hints_label);

    // Attach the tree and present it, then leave. Presenting here is what the
    // iOS path does implicitly when the view controller's view loads; we do
    // NOT enter the event loop — the point is the structure, and on a
    // headless machine (CI, container) there is no display to run it on.
    win.set_child_box(&vbox);
    win.present();

    let menu_roots: Vec<&str> = bar.iter().map(|m| m.label).collect();
    println!("corro iOS UI tree");
    println!("  root        : vertical box (toplevel on desktop, app root UIView on iOS)");
    println!("  menu bar    : {} top-level menus {:?}", bar.len(), menu_roots);
    println!("  formula bar : label \"A1\" + label \"fx\" + expanding entry + status label");
    println!("  sheet       : canvas (SheetView / drawRect:), expanding");
    println!("  tab strip   : canvas, hidden until the workbook has 2+ sheets");
    println!("  status line : label \"Ready\"");
    println!();
    println!("The iOS menu model the host renders is:");
    for (menu, items) in corro::gui::ios_backend::menu_model() {
        println!("  {menu}: {}", items.len());
    }
    Ok(())
}
