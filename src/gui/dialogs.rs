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

/// Spreadsheet file types the loaders accept (see `load_initial`): the
/// Open dialog filters to these by default, with an All-files fallback.
/// Extensions come from `ui_core::SPREADSHEET_EXTS` so the filter can't
/// drift from what the loaders actually open (unit-tested headlessly).
/// Only the `gui` Open-dialog body calls this.
#[allow(dead_code)]
pub(crate) fn spreadsheet_filter_patterns() -> Vec<String> {
    crate::ui_core::SPREADSHEET_EXTS
        .iter()
        .map(|e| format!("*.{e}"))
        .collect()
}

pub fn file_open_dialog() -> Option<PathBuf> {
    #[cfg(feature = "gui")]
    return App::init().ok().and_then(|app| {
        let patterns = spreadsheet_filter_patterns();
        let refs: Vec<&str> = patterns.iter().map(|s| s.as_str()).collect();
        let filters = [
            ("Spreadsheets", refs.as_slice()),
            ("All files", &["*.*"][..]),
        ];
        app.open_file_filtered("Open Spreadsheet", &filters)
            .ok()
            .flatten()
            .map(PathBuf::from)
    });
    #[allow(unreachable_code)]
    None
}

/// Save As dialog: filtered to `.corro`, suggesting a `.corro` name, and
/// forcing the extension on return (a workbook under a foreign extension
/// will not reopen as one). Mirrors the ratatui `to_corro_path` rule.
pub fn file_save_dialog() -> Option<PathBuf> {
    file_save_dialog_named("Sheet1.corro")
}

/// Save As dialog with a suggested filename (e.g. the current workbook
/// name); the `.corro` default still applies.
/// Params feed only the `gui` body; other builds take the None fallback.
#[allow(unused_variables)]
pub fn file_save_dialog_named(suggested: &str) -> Option<PathBuf> {
    #[cfg(feature = "gui")]
    return App::init().ok().and_then(|app| {
        app.save_file_filtered(
            "Save Spreadsheet",
            // Plain name only: join_dialog_filters appends " (*.corro)"
            // itself, so embedding it here shows "(*.corro) (*.corro)".
            &[(crate::ui_core::ext_filter_label("corro"), &["*.corro"].as_slice())],
            suggested,
        )
        .ok()
        .flatten()
        .map(|p| crate::ui_core::force_extension(&PathBuf::from(p), "corro"))
    });
    #[allow(unreachable_code)]
    None
}

/// Export dialog for `action` (`export_tsv/csv/ods/ascii`): matching type
/// filter, suggested filename with the format extension, and the extension
/// appended when the user types a bare name.
/// Params feed only the `gui` body; other builds take the None fallback.
#[allow(unused_variables)]
pub fn file_export_dialog(action: &str) -> Option<PathBuf> {
    let ext = crate::ui_core::export_ext_for_action(action);
    let title = match action {
        "export_csv" => "Export CSV",
        "export_ods" => "Export ODS",
        "export_ascii" => "Export ASCII",
        _ => "Export TSV",
    };
    let filter = crate::ui_core::ext_filter_label(ext);
    let suggested = format!("Sheet1.{ext}");
    let pat = format!("*.{ext}");
    let pats = [pat.as_str()];
    let filters = [(filter, pats.as_slice())];
    #[cfg(feature = "gui")]
    return App::init().ok().and_then(|app| {
        app.save_file_filtered(title, &filters, &suggested)
            .ok()
            .flatten()
            .map(|p| crate::ui_core::append_extension_if_missing(&PathBuf::from(p), ext))
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
                // A multi-line (read-only) text view rather than a Label: the
                // NWG STATIC used for labels vertically centers a single line
                // and collapses its client area to one line height, so a
                // two-line label silently drops its second line. EDIT boxes
                // keep their full client height and render the whole block on
                // every backend (same control the Keybindings dialog uses).
                if let Ok(tv) = rxapp.create_textview() {
                    tv.set_text(&format!(
                        "corro {}\n\nAppend-only collaborative spreadsheet",
                        env!("CARGO_PKG_VERSION"),
                    ));
                    dialog.set_title("About corro");
                    dialog.set_default_size(300, 200);
                    dialog.append_content_area(&tv);
                    dialog.add_button("Close", -7);
                    dialog.connect_response(move |_| {}).ok();
                    dialog.present();
                    // Leak the content like every other dialog does with its
                    // widgets (see prompt_dialog's entry_ptr): dropping the
                    // wrapper destroys the native control, leaving the dialog
                    // present but empty.
                    let _ = Box::into_raw(Box::new(tv));
                    let _ = Box::into_raw(Box::new(dialog));
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
                    // Leak the textview (see show_about_dialog): dropping the
                    // wrapper destroys the native EDIT and the dialog renders
                    // empty.
                    let _ = Box::into_raw(Box::new(tv));
                    let _ = Box::into_raw(Box::new(dialog));
                    return;
                }
            }
        }
    }
    #[cfg(not(feature = "gui"))]
    eprintln!("Keybindings: arrows=navigate, Enter=edit, Esc=cancel, F1=help, Ctrl+Q=quit");
}


/// Focus a dialog's text entry after present(): modeless (NWG) dialogs do
/// not take focus on their own the way modal GTK dialogs do — without this,
/// typing goes to whatever had focus before and Enter submits nothing.
#[cfg(feature = "gui")]
fn focus_dialog_entry(entry_ptr: usize) {
    use rswidgets::common::Entry as CommonEntry;
    let entry: &CommonEntry = unsafe { &*(entry_ptr as *const CommonEntry) };
    entry.grab_focus();
}

/// Wire OK/Cancel buttons (response path) and Enter-in-entry (activate path)
/// to one once-only confirm for a single-entry prompt dialog. Enter commits
/// exactly like the OK button (ratatui parity: type + Enter commits, Esc
/// cancels); Esc/cancel yields None. Rc-shared because FnOnce can only move
/// into one closure.
#[cfg(feature = "gui")]
fn wire_prompt_confirm<F: FnOnce(Option<String>) + 'static>(
    dialog: &rswidgets::common::Dialog,
    entry_ptr: usize,
    on_result: F,
) {
    use rswidgets::common::Entry as CommonEntry;
    let shared: std::rc::Rc<std::cell::RefCell<(Option<F>, bool)>> =
        std::rc::Rc::new(std::cell::RefCell::new((Some(on_result), false)));
    {
        let shared = shared.clone();
        dialog
            .connect_response(move |response_id| {
                let mut g = shared.borrow_mut();
                if g.1 {
                    return;
                }
                g.1 = true;
                if let Some(f) = g.0.take() {
                    let entry: &CommonEntry = unsafe { &*(entry_ptr as *const CommonEntry) };
                    if response_id == 1 {
                        f(entry.get_text());
                    } else {
                        f(None);
                    }
                }
            })
            .ok();
    }
    {
        // Enter confirms like the OK button (and closes: the response path's
        // auto-close does not run here, so close explicitly). On backends
        // whose entry-activate is a no-op stub, Tab/Space/click still work.
        let shared = shared.clone();
        let dlg = dialog.clone();
        let entry_ref: &CommonEntry = unsafe { &*(entry_ptr as *const CommonEntry) };
        entry_ref
            .connect_activate(move |_| {
                let mut g = shared.borrow_mut();
                if g.1 {
                    return;
                }
                g.1 = true;
                if let Some(f) = g.0.take() {
                    let entry: &CommonEntry = unsafe { &*(entry_ptr as *const CommonEntry) };
                    f(entry.get_text());
                    dlg.close();
                }
            })
            .ok();
    }
}

/// Generic single-entry modal prompt with caller-supplied title, OK button
/// label, and initial text. Prompt-gated menu actions (rename/copy/delete
/// sheet, go to cell, column widths, ...) each get correctly labeled chrome
/// through this — never a recycled "Find" dialog (which is what Rename
/// Sheet showed before this existed).
/// Params are consumed only by the `gui` body below; other backends take
/// the `on_result(None)` fallback (typed TUI prompts live elsewhere).
#[allow(unused_variables)]
pub fn prompt_dialog<F: FnOnce(Option<String>) + 'static>(
    title: &str,
    ok_label: &str,
    initial: &str,
    on_result: F,
) {
    #[cfg(feature = "gui")]
    {
        if let Ok(rxapp) = rswidgets::App::init() {
            if let Ok(dialog) = rxapp.new_dialog() {
                if let Ok(entry) = rxapp.new_entry() {
                    dialog.set_title(title);
                    entry.set_text(initial);
                    dialog.append_content_area(&entry);
                    dialog.add_button("Cancel", 0);
                    dialog.add_button(ok_label, 1);
                    let entry_ptr = Box::into_raw(Box::new(entry)) as usize;
                    wire_prompt_confirm(&dialog, entry_ptr, on_result);
                    dialog.present();
                    focus_dialog_entry(entry_ptr);
                    let _ = Box::into_raw(Box::new(dialog));
                    return;
                }
            }
        }
    }
    on_result(None);
}

pub fn find_dialog<F: FnOnce(Option<String>) + 'static>(on_result: F) {
    #[cfg(feature = "gui")]
    {
        if let Ok(rxapp) = rswidgets::App::init() {
            if let Ok(dialog) = rxapp.new_dialog() {
                if let Ok(entry) = rxapp.new_entry() {
                    dialog.set_title("Find");
                    dialog.append_content_area(&entry);
                    dialog.add_button("Cancel", 0);
                    dialog.add_button("Find", 1);
                    let entry_ptr = Box::into_raw(Box::new(entry)) as usize;
                    wire_prompt_confirm(&dialog, entry_ptr, on_result);
                    dialog.present();
                    focus_dialog_entry(entry_ptr);
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
        use rswidgets::common::Orientation;
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
                // Shared once-only confirm: OK button (response path) and
                // Enter in either entry (activate path, ratatui parity).
                let shared: std::rc::Rc<std::cell::RefCell<(Option<F>, bool)>> =
                    std::rc::Rc::new(std::cell::RefCell::new((Some(on_result), false)));
                {
                    let shared = shared.clone();
                    dialog.connect_response(move |response_id| {
                        let mut g = shared.borrow_mut();
                        if g.1 {
                            return;
                        }
                        g.1 = true;
                        if let Some(f) = g.0.take() {
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
                    }).ok();
                }
                {
                    // Enter in either field confirms like Replace (and
                    // closes: the response path's auto-close runs only there).
                    let shared = shared.clone();
                    let dlg = dialog.clone();
                    let confirm = move || {
                        let mut g = shared.borrow_mut();
                        if g.1 {
                            return;
                        }
                        g.1 = true;
                        if let Some(f) = g.0.take() {
                            let find_entry: &CommonEntry = unsafe { &*(find_ptr as *const CommonEntry) };
                            let replace_entry: &CommonEntry = unsafe { &*(replace_ptr as *const CommonEntry) };
                            f(Some((
                                find_entry.get_text().unwrap_or_default(),
                                replace_entry.get_text().unwrap_or_default(),
                            )));
                            dlg.close();
                        }
                    };
                    let confirm = std::rc::Rc::new(confirm);
                    let c1 = confirm.clone();
                    let find_ref: &CommonEntry = unsafe { &*(find_ptr as *const CommonEntry) };
                    find_ref.connect_activate(move |_| c1()).ok();
                    let c2 = confirm.clone();
                    let replace_ref: &CommonEntry = unsafe { &*(replace_ptr as *const CommonEntry) };
                    replace_ref.connect_activate(move |_| c2()).ok();
                }
                dialog.present();
                focus_dialog_entry(find_ptr);
                let _ = Box::into_raw(Box::new(dialog));
                return;
            }
        }
    }
    on_result(None);
}

/// Chrome for the Insert > Special Char picker dialog. Pinned by test so a
/// mislabeled dialog (recycled "Find") fails headlessly instead of in
/// screenshots.
pub const SPECIAL_CHAR_DIALOG_TITLE: &str = "Insert special char";
pub const SPECIAL_CHAR_DIALOG_OK: &str = "Insert";

/// Insert > Special Char picker dialog: the 10 labelled choices
/// (`"1: ∞"` … `"0: θ"`, same items/order as the ratatui picker) as a
/// native radio group with Cancel/Insert chrome. Arrows move selection
/// natively inside the group on every backend, Enter confirms via the
/// dialog default response, Esc cancels. Yields the selected INDEX (into
/// the shared choice table), not text: there is deliberately no free-text
/// entry — the formula bar is the arbitrary-input path.
///
/// Radios (not a dropdown) because a focused combo consumes Return on some
/// backends, which would swallow the confirm; radio groups navigate with
/// arrows and never eat Enter, so Down*n+Enter works identically everywhere.
/// Single column, deliberately: a multi-column grid would navigate
/// spatially (Down from row 1 lands 4 rows down the indices), breaking the
/// Down*n → nth-choice contract the ratatui reference defines (verified:
/// Down*2 in a 4-column trial landed on index 8, not 2). Linear layout
/// keeps arrows stepping ±1 through the indices on every backend.
/// Params are consumed only by the `gui` body below; other backends take
/// the `on_result(None)` fallback (the picker lives in shared state).
#[allow(unused_variables)]
pub fn special_char_dialog<F: FnOnce(Option<usize>) + 'static>(
    items: &[String],
    initial: usize,
    on_result: F,
) {
    #[cfg(feature = "gui")]
    {
        // Combined-gui flips the root prelude to pancurses-adapter types;
        // these dialog widgets are always native: prefer the platform
        // adapter, falling back to the prelude elsewhere (unchanged).
        #[cfg(target_os = "linux")]
        use rswidgets::backends_gtk_adapter::RadioButton;
        #[cfg(windows)]
        use rswidgets::backends_nwg_adapter::RadioButton;
        #[cfg(not(any(target_os = "linux", target_os = "windows")))]
        use rswidgets::prelude::RadioButton;
        use rswidgets::common::Orientation;
        if let Ok(rxapp) = rswidgets::App::init() {
            if let (Ok(dialog), Ok(vbox)) =
                (rxapp.new_dialog(), rxapp.new_box(Orientation::Vertical, 4))
            {
                dialog.set_title(SPECIAL_CHAR_DIALOG_TITLE);
                dialog.set_default_size(300, 340);
                // One radio group in choice order: the first stands alone,
                // each next joins the previous (native mutual exclusion +
                // arrow navigation in creation order, so Down*n lands on
                // the nth choice on every backend). Single column (see doc
                // above for why a grid would break the contract).
                let mut radios: Vec<RadioButton> = Vec::new();
                let mut built_ok = !items.is_empty();
                for item in items {
                    let group = radios.last();
                    match rxapp.create_radiobutton(group, item) {
                        Ok(rb) => {
                            rb.set_hexpand(true);
                            vbox.append(&rb);
                            radios.push(rb);
                        }
                        Err(_) => {
                            built_ok = false;
                            break;
                        }
                    }
                }
                if !built_ok {
                    on_result(None);
                    return;
                }
                let sel = initial.min(radios.len() - 1);
                radios[sel].set_active(true);
                dialog.append_content_area(&vbox);
                dialog.add_button("Cancel", 0);
                dialog.add_button(SPECIAL_CHAR_DIALOG_OK, 1);
                // Enter anywhere unhandled confirms (Insert): GTK natively,
                // NWG via dialog-level raw routing. Both inners expose the
                // same method, so no per-platform branching here.
                dialog.inner.set_default_response(1);
                let rb_ptr = Box::into_raw(Box::new(radios)) as usize;
                let mut on_result = Some(on_result);
                let callback_called = std::cell::RefCell::new(false);
                // Close FIRST, then report: teardown restores focus
                // synchronously, so the callee's entry grab (splice path)
                // lands instead of racing the destroy. The wrapper's own
                // post-response close becomes a harmless no-op.
                let dlg_close = dialog.clone();
                dialog.connect_response(move |response_id| {
                    let mut called = callback_called.borrow_mut();
                    if !*called {
                        *called = true;
                        if let Some(f) = on_result.take() {
                            let rbs: &Vec<RadioButton> =
                                unsafe { &*(rb_ptr as *const Vec<RadioButton>) };
                            let result = if response_id == 1 {
                                let idx = rbs.iter().position(|r| r.is_active()).unwrap_or(0);
                                Some(idx)
                            } else {
                                None
                            };
                            dlg_close.close();
                            f(result);
                        }
                    }
                }).ok();
                dialog.present();
                // Arrows navigate from the focused radio.
                let rbs: &Vec<RadioButton> = unsafe { &*(rb_ptr as *const Vec<RadioButton>) };
                rbs[sel].grab_focus();
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
        // (same native-widget shadowing as above)
        #[cfg(target_os = "linux")]
        use rswidgets::backends_gtk_adapter::{CheckButton, DropDown};
        #[cfg(windows)]
        use rswidgets::backends_nwg_adapter::{CheckButton, DropDown};
        #[cfg(not(any(target_os = "linux", target_os = "windows")))]
        use rswidgets::prelude::{CheckButton, DropDown};
        use rswidgets::common::Orientation;
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
                focus_dialog_entry(entry_ptr);
                let _ = Box::into_raw(Box::new(dialog));
                return;
            }
        }
    }
    on_result(None);
}
