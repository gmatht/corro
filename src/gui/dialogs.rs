use crate::ops::WorkbookState;
use std::path::PathBuf;
#[cfg(feature = "gui")]
use rustxwidgets::prelude::Orientation;

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
    return rustxwidgets::App::init().ok().and_then(|app| {
        app.open_file("Open Spreadsheet").ok().flatten().map(PathBuf::from)
    });
    #[allow(unreachable_code)]
    None
}

pub fn file_save_dialog() -> Option<PathBuf> {
    #[cfg(feature = "gui")]
    return rustxwidgets::App::init().ok().and_then(|app| {
        app.save_file("Save Spreadsheet").ok().flatten().map(PathBuf::from)
    });
    #[allow(unreachable_code)]
    None
}

pub fn show_about_dialog() {
    #[cfg(feature = "gui")]
    {
        use rustxwidgets::backends_gtk_adapter::{create_dialog, create_label};
        log_dialog_action("about_dialog", "");
        if let Ok(dialog) = create_dialog() {
            if let Ok(label) = create_label(&format!(
                "corro {}\n\nAppend-only collaborative spreadsheet",
                env!("CARGO_PKG_VERSION"),
            )) {
                dialog.set_title("About corro");
                dialog.set_default_size(300, 200);
                dialog.append_content_area(&label);
                dialog.add_button("Close", -7);
                // The default GtkDialog response handler closes on any
                // response, including GTK_RESPONSE_CLOSE (-7).
                dialog.connect_response(move |_| {}).ok();
                dialog.present();
                // Dialog::drop will call g_object_unref which finalizes it.
                // This is safe because gtk_widget_destroy on ROOT widgets
                // does NOT unref — only disposes children.  Dialog::drop
                // provides the final unref.
                return;
            }
        }
    }
    #[cfg(not(feature = "gui"))]
    eprintln!("corro {} - append-only collaborative spreadsheet", env!("CARGO_PKG_VERSION"));
}

pub fn show_keybinds_help() {
    #[cfg(feature = "gui")]
    {
        use rustxwidgets::backends_gtk_adapter::{create_dialog, create_textview};
        log_dialog_action("keybinds_help", "");
        if let Ok(dialog) = create_dialog() {
            if let Ok(tv) = create_textview() {
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
                    // The default GtkDialog handler closes on any response.
                    dialog.connect_response(move |_| {}).ok();
                    dialog.present();
                    // Keep alive: leak so GTK manages the lifecycle
                    let _ = Box::into_raw(Box::new(dialog));
                    return;
            }
        }
    }
    #[cfg(not(feature = "gui"))]
    eprintln!("Keybindings: arrows=navigate, Enter=edit, Esc=cancel, F1=help, Ctrl+Q=quit");
}


pub fn find_dialog<F: FnOnce(Option<String>) + 'static>(on_result: F) {
    #[cfg(feature = "gui")]
    {
        use rustxwidgets::backends_gtk_adapter::{create_dialog, create_entry};
        use rustxwidgets::Entry;
        if let Ok(dialog) = create_dialog() {
            if let Ok(entry) = create_entry() {
                    dialog.set_title("Find");
                    dialog.append_content_area(&entry);
                    dialog.add_button("Cancel", 0);
                    dialog.add_button("Find", 1);
                    let entry_ptr = Box::into_raw(Box::new(entry.clone())) as usize;
                    let mut on_result = Some(on_result);
                    let callback_called = std::cell::RefCell::new(false);
                    // The default GtkDialog handler closes the dialog on response.
                    dialog.connect_response(move |response_id| {
                        let mut called = callback_called.borrow_mut();
                        if !*called {
                            *called = true;
                            if let Some(f) = on_result.take() {
                                let entry: &Entry = unsafe { &*(entry_ptr as *const Entry) };
                                if response_id == 1 {
                                    f(entry.get_text());
                                } else {
                                    f(None);
                                }
                            }
                        }
                    }).ok();
                    dialog.present();
                    // Keep alive: leak so GTK manages the lifecycle
                    let _ = Box::into_raw(Box::new(dialog));
                    return;
                }
            }
        }
    on_result(None);
}

pub fn replace_dialog<F: FnOnce(Option<(String, String)>) + 'static>(on_result: F) {
    #[cfg(feature = "gui")]
    {
        use rustxwidgets::backends_gtk_adapter::{create_dialog, create_entry, create_label, create_box};
        use rustxwidgets::prelude::Orientation;
        use rustxwidgets::Entry;
        if let (Ok(dialog), Ok(find_entry), Ok(replace_entry), Ok(vbox)) =
            (create_dialog(), create_entry(), create_entry(), create_box(Orientation::Vertical, 4))
            {
                dialog.set_title("Replace");
                dialog.set_default_size(350, 150);
                if let Ok(find_label) = create_label("Find:") {
                    vbox.append(&find_label);
                }
                vbox.append(&find_entry);
                if let Ok(replace_label) = create_label("Replace with:") {
                    vbox.append(&replace_label);
                }
                vbox.append(&replace_entry);
                dialog.append_content_area(&vbox);
                dialog.add_button("Cancel", 0);
                dialog.add_button("Replace", 1);
                let find_ptr = Box::into_raw(Box::new(find_entry.clone())) as usize;
                let replace_ptr = Box::into_raw(Box::new(replace_entry.clone())) as usize;
                let mut on_result = Some(on_result);
                let callback_called = std::cell::RefCell::new(false);
                // The default GtkDialog handler closes the dialog on response.
                dialog.connect_response(move |response_id| {
                    let mut called = callback_called.borrow_mut();
                    if !*called {
                        *called = true;
                        if let Some(f) = on_result.take() {
                            let find_entry: &Entry = unsafe { &*(find_ptr as *const Entry) };
                            let replace_entry: &Entry = unsafe { &*(replace_ptr as *const Entry) };
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
                // Keep alive: leak so GTK manages the lifecycle
                let _ = Box::into_raw(Box::new(dialog));
                return;
            }
        }
    on_result(None);
}

pub fn sort_dialog<F: FnOnce(Option<(usize, bool)>) + 'static>(_workbook: &WorkbookState, on_result: F) {
    #[cfg(feature = "gui")]
{
        use rustxwidgets::backends_gtk_adapter::{create_dialog, create_dropdown, create_checkbutton, create_label};
        use rustxwidgets::prelude::Orientation;
        use rustxwidgets::{CheckButton, DropDown};
        let cols: &[&str] = &["Column A", "Column B", "Column C", "Column D", "Column E"];
        if let (Ok(dialog), Ok(sort_col), Ok(ascending), Ok(vbox)) =
            (create_dialog(), create_dropdown(cols), create_checkbutton("Ascending"),
             rustxwidgets::backends_gtk_adapter::create_box(Orientation::Vertical, 4))
        {
            dialog.set_title("Sort");
            dialog.set_default_size(300, 150);
            if let Ok(label) = create_label("Sort column:") {
                vbox.append(&label);
                let _ = Box::into_raw(Box::new(label));
            }
            sort_col.set_hexpand(true);
            vbox.append(&sort_col);
            ascending.set_active(true);
            vbox.append(&ascending);
            dialog.append_content_area(&vbox);
            dialog.add_button("Cancel", 0);
            dialog.add_button("Sort", 1);
            let sort_col_ptr = Box::into_raw(Box::new(sort_col.clone())) as usize;
            let ascending_ptr = Box::into_raw(Box::new(ascending.clone())) as usize;
            let mut on_result = Some(on_result);
            let callback_called = std::cell::RefCell::new(false);
            // The default GtkDialog handler closes the dialog on response.
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
            // Keep alive: leak so GTK manages the lifecycle
            let _ = Box::into_raw(Box::new(dialog));
            return;
        }
    }
    on_result(None);
}

pub fn balance_dialog<F: FnOnce(Option<String>) + 'static>(on_result: F) {
    #[cfg(feature = "gui")]
    {
        use rustxwidgets::backends_gtk_adapter::{create_dialog, create_entry, create_label};
        use rustxwidgets::Entry;
        if let (Ok(dialog), Ok(entry)) = (create_dialog(), create_entry()) {
                dialog.set_title("Balance Books");
                if let Ok(label) = create_label("Column to balance:") {
                    dialog.append_content_area(&label);
                    let _ = Box::into_raw(Box::new(label));
                }
                entry.set_text("A");
                entry.set_hexpand(true);
                dialog.append_content_area(&entry);
                dialog.add_button("Cancel", 0);
                dialog.add_button("Balance", 1);
                let entry_ptr = Box::into_raw(Box::new(entry.clone())) as usize;
                let mut on_result = Some(on_result);
                // The default GtkDialog handler closes the dialog on response.
                dialog.connect_response(move |response_id| {
                    if let Some(f) = on_result.take() {
                        let entry: &Entry = unsafe { &*(entry_ptr as *const Entry) };
                        if response_id == 1 {
                            f(entry.get_text());
                        } else {
                            f(None);
                        }
                    }
                }).ok();
                dialog.present();
                // Keep alive: leak so GTK manages the lifecycle
                let _ = Box::into_raw(Box::new(dialog));
                return;
            }
    }
    on_result(None);
}
