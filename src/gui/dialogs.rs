use crate::ops::WorkbookState;
#[cfg(feature = "gui")]
use crate::gui::app_alias::App;
use std::path::PathBuf;

#[allow(dead_code, unused_variables)] // only used by the gui feature
fn log_dialog_action(action: &str, detail: &str) {
    #[cfg(feature = "gui")]
    if let Ok(mut f) = std::fs::OpenOptions::new()
        .create(true).append(true)
        .open("/tmp/corro_keylog.txt")
    {
        use std::io::Write;
        let _ = writeln!(f, "KEY: MENU  ACTION: {action}  DETAIL: {detail}");
    }
}

pub fn file_open_dialog() -> Option<PathBuf> {
    #[cfg(feature = "gui")]
    return App::init().ok().and_then(|app| {
        app.open_file("Open Spreadsheet").ok().flatten().map(PathBuf::from)
    });
    #[allow(unreachable_code)]
    None
}

pub fn file_save_dialog() -> Option<PathBuf> {
    #[cfg(feature = "gui")]
    return App::init().ok().and_then(|app| {
        app.save_file("Save Spreadsheet").ok().flatten().map(PathBuf::from)
    });
    #[allow(unreachable_code)]
    None
}

pub fn show_about_dialog() {
    #[cfg(feature = "gui")]
    {
        log_dialog_action("about_dialog", "");
        if let Ok(rxapp) = rswidgets::App::init() {
            if let Ok(dialog) = rxapp.new_dialog() {
                if let Ok(label) = rxapp.new_label(&format!(
                    "corro {}\n\nAppend-only collaborative spreadsheet",
                    env!("CARGO_PKG_VERSION"),
                )) {
                    dialog.set_title("About corro");
                    dialog.set_default_size(300, 200);
                    dialog.append_content_area(&label);
                    dialog.add_button("Close", -7);
                    dialog.connect_response(move |_| {}).ok();
                    dialog.present();
                    return;
                }
            }
        }
    }
    #[cfg(not(feature = "gui"))]
    eprintln!("corro {} - append-only collaborative spreadsheet", env!("CARGO_PKG_VERSION"));
}

pub fn show_keybinds_help() {
    #[cfg(feature = "gui")]
    {
        log_dialog_action("keybinds_help", "");
        if let Ok(rxapp) = rswidgets::App::init() {
            if let Ok(dialog) = rxapp.new_dialog() {
                if let Ok(tv) = rxapp.create_textview() {
                    dialog.set_title("Keybindings");
                    tv.set_text(
                        "Navigation:    Arrow keys / Page Up/Down / Home / End\n\
                         Edit:          Enter (edit cell), F2 (edit cell)\n\
                         Cancel:        Escape\n\
                         Help:          F1\n\
                         Quit:          Ctrl+Q\n\
                         Menu:          Alt+underlined letter\n\
                         \n\
                         File menu:     Ctrl+O (open), Ctrl+S (save)\n\
                         Edit menu:     Ctrl+Z (undo), Ctrl+Y (redo)\n\
                                           Ctrl+X (cut), Ctrl+C (copy), Ctrl+V (paste)\n\
                                           Ctrl+F (find), Ctrl+H (replace)"
                    );
                    tv.set_wrap_mode(0);
                    tv.set_size_request(400, 300);
                    dialog.append_content_area(&tv);
                    dialog.add_button("Close", -7);
                    dialog.connect_response(move |_| {}).ok();
                    dialog.present();
                    let _ = Box::into_raw(Box::new(dialog));
                    return;
                }
            }
        }
    }
    #[cfg(not(feature = "gui"))]
    eprintln!("Keybindings: arrows=navigate, Enter=edit, Esc=cancel, F1=help, Ctrl+Q=quit");
}


pub fn find_dialog<F: FnOnce(Option<String>) + 'static>(on_result: F) {
    #[cfg(feature = "gui")]
    {
        use rswidgets::common::Entry as CommonEntry;
        if let Ok(rxapp) = rswidgets::App::init() {
            if let Ok(dialog) = rxapp.new_dialog() {
                if let Ok(entry) = rxapp.new_entry() {
                    dialog.set_title("Find");
                    dialog.append_content_area(&entry);
                    dialog.add_button("Cancel", 0);
                    dialog.add_button("Find", 1);
                    let entry_ptr = Box::into_raw(Box::new(entry)) as usize;
                    let mut on_result = Some(on_result);
                    let callback_called = std::cell::RefCell::new(false);
                    dialog.connect_response(move |response_id| {
                        let mut called = callback_called.borrow_mut();
                        if !*called {
                            *called = true;
                            if let Some(f) = on_result.take() {
                                let entry: &CommonEntry = unsafe { &*(entry_ptr as *const CommonEntry) };
                                if response_id == 1 {
                                    f(entry.get_text());
                                } else {
                                    f(None);
                                }
                            }
                        }
                    }).ok();
                    dialog.present();
                    let _ = Box::into_raw(Box::new(dialog));
                    return;
                }
            }
        }
    }
    on_result(None);
}

pub fn replace_dialog<F: FnOnce(Option<(String, String)>) + 'static>(on_result: F) {
    #[cfg(feature = "gui")]
    {
        use rswidgets::common::Entry as CommonEntry;
        use rswidgets::prelude::Orientation;
        if let Ok(rxapp) = rswidgets::App::init() {
            if let (Ok(dialog), Ok(find_entry), Ok(replace_entry), Ok(vbox)) =
                (rxapp.new_dialog(), rxapp.new_entry(), rxapp.new_entry(), rxapp.new_box(Orientation::Vertical, 4))
            {
                dialog.set_title("Replace");
                dialog.set_default_size(350, 150);
                if let Ok(find_label) = rxapp.new_label("Find:") {
                    vbox.append(&find_label);
                }
                vbox.append(&find_entry);
                if let Ok(replace_label) = rxapp.new_label("Replace with:") {
                    vbox.append(&replace_label);
                }
                vbox.append(&replace_entry);
                dialog.append_content_area(&vbox);
                dialog.add_button("Cancel", 0);
                dialog.add_button("Replace", 1);
                let find_ptr = Box::into_raw(Box::new(find_entry)) as usize;
                let replace_ptr = Box::into_raw(Box::new(replace_entry)) as usize;
                let mut on_result = Some(on_result);
                let callback_called = std::cell::RefCell::new(false);
                dialog.connect_response(move |response_id| {
                    let mut called = callback_called.borrow_mut();
                    if !*called {
                        *called = true;
                        if let Some(f) = on_result.take() {
                            let find_entry: &CommonEntry = unsafe { &*(find_ptr as *const CommonEntry) };
                            let replace_entry: &CommonEntry = unsafe { &*(replace_ptr as *const CommonEntry) };
                            if response_id == 1 {
                                f(Some((
                                    find_entry.get_text().unwrap_or_default(),
                                    replace_entry.get_text().unwrap_or_default(),
                                )));
                            } else {
                                f(None);
                            }
                        }
                    }
                }).ok();
                dialog.present();
                let _ = Box::into_raw(Box::new(dialog));
                return;
            }
        }
    }
    on_result(None);
}

pub fn sort_dialog<F: FnOnce(Option<(usize, bool)>) + 'static>(_workbook: &WorkbookState, on_result: F) {
    #[cfg(feature = "gui")]
    {
        use rswidgets::prelude::*;
        let cols: &[&str] = &["Column A", "Column B", "Column C", "Column D", "Column E"];
        if let Ok(rxapp) = rswidgets::App::init() {
            if let (Ok(dialog), Ok(sort_col), Ok(ascending), Ok(vbox)) =
                (rxapp.new_dialog(), rxapp.create_dropdown(cols), rxapp.create_checkbutton("Ascending"),
                 rxapp.new_box(Orientation::Vertical, 4))
            {
                dialog.set_title("Sort");
                dialog.set_default_size(300, 150);
                if let Ok(label) = rxapp.new_label("Sort column:") {
                    vbox.append(&label);
                }
                sort_col.set_hexpand(true);
                vbox.append(&sort_col);
                ascending.set_active(true);
                vbox.append(&ascending);
                dialog.append_content_area(&vbox);
                dialog.add_button("Cancel", 0);
                dialog.add_button("Sort", 1);
                let sort_col_ptr = Box::into_raw(Box::new(sort_col)) as usize;
                let ascending_ptr = Box::into_raw(Box::new(ascending)) as usize;
                let mut on_result = Some(on_result);
                let callback_called = std::cell::RefCell::new(false);
                dialog.connect_response(move |response_id| {
                    let mut called = callback_called.borrow_mut();
                    if !*called {
                        *called = true;
                        if let Some(f) = on_result.take() {
                            let sort_col: &DropDown = unsafe { &*(sort_col_ptr as *const DropDown) };
                            let ascending: &CheckButton = unsafe { &*(ascending_ptr as *const CheckButton) };
                            if response_id == 1 {
                                let col = sort_col.get_active().max(0) as usize;
                                let asc = ascending.is_active();
                                f(Some((col, asc)));
                            } else {
                                f(None);
                            }
                        }
                    }
                }).ok();
                dialog.present();
                let _ = Box::into_raw(Box::new(dialog));
                return;
            }
        }
    }
    on_result(None);
}

pub fn balance_dialog<F: FnOnce(Option<String>) + 'static>(on_result: F) {
    #[cfg(feature = "gui")]
    {
        use rswidgets::common::Entry as CommonEntry;
        if let Ok(rxapp) = rswidgets::App::init() {
            if let (Ok(dialog), Ok(entry)) = (rxapp.new_dialog(), rxapp.new_entry()) {
                dialog.set_title("Balance Books");
                if let Ok(label) = rxapp.new_label("Column to balance:") {
                    dialog.append_content_area(&label);
                }
                entry.set_text("A");
                entry.set_hexpand(true);
                dialog.append_content_area(&entry);
                dialog.add_button("Cancel", 0);
                dialog.add_button("Balance", 1);
                let entry_ptr = Box::into_raw(Box::new(entry)) as usize;
                let mut on_result = Some(on_result);
                dialog.connect_response(move |response_id| {
                    if let Some(f) = on_result.take() {
                        let entry: &CommonEntry = unsafe { &*(entry_ptr as *const CommonEntry) };
                        if response_id == 1 {
                            f(entry.get_text());
                        } else {
                            f(None);
                        }
                    }
                }).ok();
                dialog.present();
                let _ = Box::into_raw(Box::new(dialog));
                return;
            }
        }
    }
    on_result(None);
}
