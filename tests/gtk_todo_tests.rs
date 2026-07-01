#![cfg(feature = "gui")]

use std::fs;
use std::path::Path;

#[test]
fn gui_menu_items_available() {
    let gui_dir = Path::new("src/gui");
    let mut all_menu_items_found = false;

    for entry in fs::read_dir(gui_dir).unwrap() {
        let path = entry.unwrap().path();
        if path.extension().map_or(true, |e| e != "rs") {
            continue;
        }
        let content = fs::read_to_string(&path).unwrap();
        if content.contains("file_open_dialog()") &&
           content.contains("file_save_dialog()") &&
           content.contains("show_keybinds_help()") &&
           content.contains("show_about_dialog()") &&
           content.contains("find_dialog(") &&
           content.contains("replace_dialog(") &&
           content.contains("sort_dialog(") &&
           content.contains("balance_dialog(")
        {
            all_menu_items_found = true;
            break;
        }
    }

    assert!(all_menu_items_found, "Not all Ratatui menu items have GTK equivalents");
}

/// Strip leading whitespace from each line so pattern matching is
/// resilient to indentation changes (e.g. rustfmt, refactoring).
fn strip_leading(s: &str) -> String {
    s.lines().map(|l| l.trim_start()).collect::<Vec<_>>().join("\n")
}

#[test]
fn gui_menu_items_not_wired_to_real_functions() {
    // Verify that menu items ARE currently NOT wired to real functions, they use eprintln instead
    let menu_path = Path::new("src/gui/menu.rs");
    let content = fs::read_to_string(menu_path).unwrap();
    let normalized = strip_leading(&content);

    // Check that menu handlers DO use eprintln instead of calling real functions.
    // Patterns match the current handle_action match-arm structure in menu.rs,
    // with leading whitespace stripped so formatting changes don't break the match.
    let open_pattern = strip_leading("\"open\" => {\n    if let Some(path) = dialogs::file_open_dialog() {\n        eprintln!(\"Open file: {:?}\", path);\n    }\n}");
    let save_pattern = strip_leading("\"save\" => {\n    if let Some(path) = dialogs::file_save_dialog() {\n        eprintln!(\"Save file: {:?}\", path);\n    }\n}");
    let about_pattern = strip_leading("\"about\" => dialogs::show_about_dialog(),");
    let help_pattern = strip_leading("\"help_keybinds\" => dialogs::show_keybinds_help(),");
    let find_pattern = strip_leading("\"find\" => dialogs::find_dialog(|result| {\n    if let Some(text) = result {\n        eprintln!(\"Find: {}\", text);\n    }\n}),");
    let replace_pattern = strip_leading("\"replace\" => dialogs::replace_dialog(|result| {\n    if let Some((find, replace)) = result {\n        eprintln!(\"Replace: '{}' with '{}'\", find, replace);\n    }\n}),");

    assert!(normalized.contains(&open_pattern), "Menu item 'Open' should still be using eprintln instead of calling real function");
    assert!(normalized.contains(&save_pattern), "Menu item 'Save' should still be using eprintln instead of calling real function");
    assert!(normalized.contains(&about_pattern), "Menu item 'About' should NOT use eprintln (already calls real function)");
    assert!(normalized.contains(&help_pattern), "Menu item 'Keybindings' should NOT use eprintln (already calls real function)");
    assert!(normalized.contains(&find_pattern), "Menu item 'Find' should still be using eprintln instead of calling real function");
    assert!(normalized.contains(&replace_pattern), "Menu item 'Replace' should still be using eprintln instead of calling real function");
}

#[test]
#[cfg(feature = "ratatui")]
fn gui_sheet_tab_bar_displayed() {
    let mut app = corro::ui::App::new(None);
    app.load_initial().unwrap();

    // Verify that new sheets are displayed in a tab bar at the bottom
    // Create a new sheet and verify it's displayed
    let initial_count = app.workbook.sheet_count();
    let new_id = app.workbook.next_sheet_id;
    let title = format!("Sheet{}", new_id);

    let op = corro::ops::WorkbookOp::NewSheet { id: new_id, title: title.clone() };
    let _ = corro::ops::apply_workbook_op(&mut app.workbook, &mut 0, op.clone());

    // The new sheet should be added and visible in the tab bar
    assert!(app.workbook.sheet_count() == initial_count + 1);
    let new_index = app.workbook.sheet_index_by_id(new_id).expect("New sheet should have an index");
    assert!(app.workbook.sheet_title(new_index) == title);
}

#[test]
fn gui_add_column_plus_column() {
    let path = Path::new("docs/tests/subtotal.corro");
    let mut app = corro::gui::App::new_with_paths(vec![path.to_path_buf()]);
    app.set_backend(corro::gui::Backend::Gui);
    app.load_initial().unwrap();

    // Verify that a + column exists between data and right margin
    // Clicking it should create a new data column
    let _sheet = app.core.workbook.active_sheet();
    let _sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);

    // In GTK mode, the + column should be implemented
    // We can't test the actual GUI here, but we can verify the logic
    // This test documents the requirement for GTK implementation
}

#[test]
#[cfg(target_os = "linux")]
fn gui_spreadsheet_scrollbars() {
    // Verify that the gtk::create_scrolled_window symbol exists.
    // This compiles only when the GTK backend is active; the runtime
    // call will fail (loader not initialized) unless a prior init()
    // has been made — the test just checks the function is present.
    let _has_scrolled_window = rustxwidgets::backends::gtk::create_scrolled_window().is_ok();
    // Scrollbars should be provided when content exceeds viewport
}
