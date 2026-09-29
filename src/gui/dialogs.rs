use crate::ops::WorkbookState;
#[cfg(any(feature = "gui", feature = "gui-core"))]
use crate::gui::app_alias::App;
use std::path::PathBuf;
use crate::balance::BalanceDirection;

#[allow(dead_code, unused_variables)] // only used by the gui feature
fn log_dialog_action(action: &str, detail: &str) {
    #[cfg(any(feature = "gui", feature = "gui-core"))]
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
    #[cfg(any(feature = "gui", feature = "gui-core"))]
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
            .as_deref()
            .and_then(resolve_picked_document)
    });
    #[allow(unreachable_code)]
    None
}

/// Turn a chosen document into something the loaders can read.
///
/// On a desktop the file dialog returns a **path**. On Android the Storage
/// Access Framework returns a `content://` **URI**, and there is no filesystem
/// path behind one — `std::fs` cannot open it and `Path::exists` is false for
/// it, so handing the URI on unchanged (which is what the code did before,
/// via `PathBuf::from`) made File &rarr; Open load nothing and say nothing.
///
/// So on Android the document is copied once into the app's private storage
/// and a real path is returned, which every existing loader works with
/// unchanged. The copy is what makes the *import* direction work at all: the
/// permission granted by the picker is per-document and does not survive a
/// reboot, while the copy does.
///
/// A path is passed straight through, so a host that substituted a plain file
/// chooser (or a test) still works and the check costs one `starts_with`.
#[cfg(any(feature = "gui", feature = "gui-core"))]
fn resolve_picked_document(picked: &str) -> Option<PathBuf> {
    if !picked.starts_with("content://") && !picked.starts_with("file://") {
        return Some(PathBuf::from(picked));
    }
    // Keep the extension in the copied name: the loaders dispatch on it
    // (`load_initial` matches "corro"/"ods"/"tsv"/"csv"), and a document
    // provider's display name is the only place the type survives.
    let name = picked
        .rsplit('/')
        .find(|s| !s.is_empty())
        .unwrap_or("document");
    #[cfg(target_os = "android")]
    {
        let local = rswidgets::backends::android::materialize_document(picked, name)?;
        Some(PathBuf::from(local))
    }
    // Off Android no dialog can hand back a URI, so this arm is unreachable.
    // It returns `None` rather than a fake path so a substituted file
    // chooser that *did* return one reports a failed open honestly.
    #[cfg(not(target_os = "android"))]
    {
        let _ = name;
        None
    }
}

/// Save As dialog: filtered to `.corro`, suggesting a `.corro` name, and
/// forcing the extension on return (a workbook under a foreign extension
/// will not reopen as one). Mirrors the ratatui `to_corro_path` rule.
pub fn file_save_dialog() -> Option<PathBuf> {
    file_save_dialog_named("Sheet1.corro")
}

/// Save As with a suggested filename, returning a [`SaveTarget`] rather than a
/// `PathBuf`.
///
/// The distinction is not cosmetic. A desktop returns a path and that is
/// where the bytes go; Android returns a `content://` URI, and
/// `write_workbook_log` writes through `std::fs` to a *temp sibling* and
/// renames it — neither of which means anything for a document provider. So
/// the URI is carried as [`SaveTarget::Document`] and the writer
/// ([`write_workbook`]) routes it through the content resolver instead.
///
/// [`file_save_dialog_named`] keeps the `PathBuf` signature for the desktop
/// callers that genuinely have a path, and delegates here.
#[cfg(any(feature = "gui", feature = "gui-core"))]
#[allow(unused_variables)]
pub(crate) fn file_save_target(suggested: &str) -> Option<SaveTarget> {
    // `gui-core` alongside `gui`, for the reason the rest of this file now
    // spells: a build that has the dialog *model* but not a GUI backend
    // still compiles it. The body is a no-op there, which is what the
    // `unreachable_code` arm below is for.
    #[cfg(any(feature = "gui", feature = "gui-core"))]
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
        .as_deref()
        .map(save_target)
    });
    #[allow(unreachable_code)]
    None
}

/// Save As dialog with a suggested filename (e.g. the current workbook
/// name); the `.corro` default still applies.
/// Params feed only the `gui` body; other builds take the None fallback.
#[allow(unused_variables)]
pub fn file_save_dialog_named(suggested: &str) -> Option<PathBuf> {
#[cfg(any(feature = "gui", feature = "gui-core"))]
    {
        // A path only on the backends that have one; a `SaveTarget` knows
        // whether it is a document, so the flattening here is the only place
        // that distinction is lost, and it is a caller that has already
        // decided it cannot use a URI.
        return file_save_target(suggested).and_then(|t| match t {
            SaveTarget::Path(p) => Some(p),
            SaveTarget::Document(_) => None,
        });
    }
    #[allow(unreachable_code)]
    None
}

/// Where a save should actually be written, given what the dialog returned.
///
/// A desktop returns a path and that is where the bytes go. Android returns a
/// `content://` URI, and `write_workbook_log` writes through `std::fs` to a
/// *temp sibling* and renames it — neither of which means anything for a
/// document provider. So the URI is returned unchanged, recognised by the
/// writer (see [`write_workbook`]), which routes it through the content
/// resolver instead.
///
/// The `.corro` extension is still forced, but on the *filename* rather than
/// the URI: the platform picker appends the extension itself when the name
/// lacks one, and a URI is not something an extension can be appended to.
#[cfg(any(feature = "gui", feature = "gui-core"))]
fn save_target(picked: &str) -> SaveTarget {
    if picked.starts_with("content://") {
        return SaveTarget::Document(picked.to_string());
    }
    SaveTarget::Path(crate::ui_core::force_extension(&PathBuf::from(picked), "corro"))
}

/// Where a save goes: a real file, or a Storage Access Framework document.
///
/// Plain data with no backend behind it, and `file_save_dialog_named` matches
/// on it unconditionally — so a `pancurses` build (which has dialogs but no GUI
/// backend) needs the type to exist even though it never constructs a
/// `Document`. The `dead_code` allow covers that build: nothing there reaches
/// it, but the match in the public entry point must still typecheck.
#[derive(Clone, Debug, PartialEq, Eq)]
#[allow(dead_code)]
pub(crate) enum SaveTarget {
    Path(PathBuf),
    /// A `content://` URI. Written through the content resolver.
    Document(String),
}

impl SaveTarget {
    /// The name to show in a status line. A URI is a long opaque string, so
    /// only its last segment is shown — which is the document's own name.
    #[allow(dead_code)]
    pub(crate) fn display(&self) -> String {
        match self {
            SaveTarget::Path(p) => p.display().to_string(),
            SaveTarget::Document(uri) => uri
                .rsplit('/')
                .find(|s| !s.is_empty())
                .unwrap_or(uri)
                .to_string(),
        }
    }
}

/// Write a workbook log to a [`SaveTarget`].
///
/// For a path this is exactly `write_workbook_log`, including the
/// write-temp-then-rename that makes a save atomic. For a document it is a
/// *stream* write through the content resolver, and the temp-and-rename step
/// is simply unavailable: a document provider exposes a single stream with no
/// sibling to rename over, so the log is written in one go. That is a real
/// difference in failure behaviour — a document save interrupted part-way
/// leaves a truncated file, where a path save leaves the old one intact — and
/// it is inherent to SAF, not a choice: there is no other handle to write to.
#[cfg(any(feature = "gui", feature = "gui-core"))]
pub(crate) fn write_workbook(target: &SaveTarget, workbook: &WorkbookState, sorts: &std::collections::HashMap<u32, Vec<crate::grid::SortSpec>>) -> Result<(), String> {
    match target {
        SaveTarget::Path(p) => crate::io::write_workbook_log(p, workbook, sorts).map_err(|e| e.to_string()),
        SaveTarget::Document(uri) => {
            // Only Android can produce a Document, so this is the one place
            // that needs the platform call. The cfg keeps the desktop build
            // free of an Android-only dependency rather than reaching for a
            // trait method that does nothing.
            #[cfg(target_os = "android")]
            {
                let text = crate::io::serialize_workbook_log(workbook, sorts);
                if rswidgets::backends::android::write_document(uri, text.as_bytes()) {
                    return Ok(());
                }
            }
            Err(format!(
                "could not write {}",
                SaveTarget::Document(uri.clone()).display()
            ))
        }
    }
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
    #[cfg(any(feature = "gui", feature = "gui-core"))]
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
    show_text_dialog(
        "about_dialog",
        "About corro",
        300,
        200,
        &format!(
            "corro {}\n\nAppend-only collaborative spreadsheet",
            env!("CARGO_PKG_VERSION"),
        ),
        None,
        &format!(
            "corro {} - append-only collaborative spreadsheet",
            env!("CARGO_PKG_VERSION")
        ),
    );
}

pub fn show_keybinds_help() {
    show_text_dialog(
        "keybinds_help",
        "Keybindings",
        400,
        300,
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
                           Ctrl+F (find), Ctrl+H (replace)",
        Some((0, 400, 300)),
        "Keybindings: arrows=navigate, Enter=edit, Esc=cancel, F1=help, Ctrl+Q=quit",
    );
}

/// A read-only, word/line-wrapped text dialog with a single Close button.
///
/// Both the About and Keybindings dialogs are exactly this — a title, a
/// block of text and a Close (dismiss) button — so they share one
/// implementation rather than two copies that drift in geometry, wrap mode
/// or lifetime handling.
///
/// A multi-line text view is used rather than a Label: the NWG STATIC used
/// for labels vertically centres a single line and collapses its client area
/// to one line height, so a two-line label silently drops its second line.
/// EDIT boxes keep their full client height and render the whole block on
/// every backend.
///
/// `wrap` is the optional `(mode, width, height)` size/wrap request (the
/// Keybindings text is wide and needs wrapping, the About text does not).
#[cfg_attr(not(any(feature = "gui", feature = "gui-core")), allow(unused_variables))]
fn show_text_dialog(
    action: &str,
    title: &str,
    w: i32,
    h: i32,
    text: &str,
    wrap: Option<(i32, i32, i32)>,
    fallback: &str,
) {
    log_dialog_action(action, "");
    #[cfg(any(feature = "gui", feature = "gui-core"))]
    {
        if let Ok(rxapp) = rswidgets::App::init() {
            if let Ok(dialog) = rxapp.new_dialog() {
                if let Ok(tv) = rxapp.create_textview() {
                    tv.set_text(text);
                    if let Some((mode, tw, th)) = wrap {
                        tv.set_wrap_mode(mode);
                        tv.set_size_request(tw, th);
                    }
                    dialog.set_title(title);
                    dialog.set_default_size(w, h);
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
    // No toolkit (or dialog creation failed): print the same content so the
    // information is not simply lost.
    let _ = (title, w, h, wrap);
    eprintln!("{fallback}");
}

/// Focus a dialog's text entry after present(): modeless (NWG) dialogs do
/// not take focus on their own the way modal GTK dialogs do — without this,
/// typing goes to whatever had focus before and Enter submits nothing.
#[cfg(any(feature = "gui", feature = "gui-core"))]
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
#[cfg(any(feature = "gui", feature = "gui-core"))]
fn wire_prompt_confirm<F: FnOnce(Option<String>) + 'static>(
    dialog: &rswidgets::common::Dialog,
    entry_ptr: usize,
    on_result: F,
) {
    use rswidgets::common::Entry as CommonEntry;
    // One guard shared by the two signals that can report a result (the
    // OK/Cancel response and Enter in the entry), so the caller's callback
    // runs exactly once whichever fires first.
    let once = super::once_callback::OnceCallback::new(on_result);
    {
        let once = once.clone();
        dialog
            .connect_response(move |response_id| {
                once.claim(|f| {
                    let entry: &CommonEntry = unsafe { &*(entry_ptr as *const CommonEntry) };
                    if response_id == 1 {
                        f(entry.get_text());
                    } else {
                        f(None);
                    }
                    true
                });
            })
            .ok();
    }
    {
        // Enter confirms like the OK button (and closes: the response path's
        // auto-close does not run here, so close explicitly). On backends
        // whose entry-activate is a no-op stub, Tab/Space/click still work.
        let once = once.clone();
        let dlg = dialog.clone();
        let entry_ref: &CommonEntry = unsafe { &*(entry_ptr as *const CommonEntry) };
        entry_ref
            .connect_activate(move |_| {
                once.claim(|f| {
                    let entry: &CommonEntry = unsafe { &*(entry_ptr as *const CommonEntry) };
                    f(entry.get_text());
                    dlg.close();
                    true
                });
            })
            .ok();
    }
}

/// Generic single-entry modal prompt with caller-supplied title, OK button
/// label, initial text, and an optional hint line describing the expected
/// input. Prompt-gated menu actions (rename/copy/delete sheet, go to cell,
/// column widths, ...) each get correctly labeled chrome through this — never
/// a recycled "Find" dialog (which is what Rename Sheet showed before this
/// existed).
///
/// `hint` is rendered as a label *above* the entry, so a prompt whose accepted
/// input is not self-evident (go to cell takes a cell ref, a row, a column, a
/// header/footer/margin ref) tells the user what to type instead of showing an
/// empty box. `None` renders no label, so the other prompts are unchanged.
///
/// The hint goes in a vertical box with the entry and the dialog gets an
/// explicit default size. Appending the label and the entry to the content
/// area directly left the box sized to the *entry alone* on a dialog with no
/// default size, so a taller hint row was clipped to half its height (the
/// dialog does not grow after `show_all`).
///
/// Params are consumed only by the `gui` body below; other backends take
/// the `on_result(None)` fallback (typed TUI prompts live elsewhere).
#[allow(unused_variables)]
pub fn prompt_dialog<F: FnOnce(Option<String>) + 'static>(
    title: &str,
    ok_label: &str,
    initial: &str,
    hint: Option<&str>,
    on_result: F,
) {
    #[cfg(any(feature = "gui", feature = "gui-core"))]
    {
        use rswidgets::common::Orientation;
        if let Ok(rxapp) = rswidgets::App::init() {
            if let Ok(dialog) = rxapp.new_dialog() {
                if let Ok(entry) = rxapp.new_entry() {
                    dialog.set_title(title);
                    entry.set_text(initial);
                    if hint.is_some() {
                        // Room for the hint line plus the entry, so neither is
                        // clipped. Without a hint the plain entry keeps the
                        // single-line dialog it has always had.
                        dialog.set_default_size(420, 150);
                        if let Ok(vbox) = rxapp.new_box(Orientation::Vertical, 4) {
                            if let Some(text) = hint {
                                if let Ok(label) = rxapp.new_label(text) {
                                    label.set_xalign(0.0);
                                    vbox.append(&label);
                                    let _ = Box::into_raw(Box::new(label));
                                }
                            }
                            vbox.append(&entry);
                            dialog.append_content_area(&vbox);
                            let _ = Box::into_raw(Box::new(vbox));
                        } else {
                            dialog.append_content_area(&entry);
                        }
                    } else {
                        dialog.append_content_area(&entry);
                    }
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
    #[cfg(any(feature = "gui", feature = "gui-core"))]
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
    #[cfg(any(feature = "gui", feature = "gui-core"))]
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
                let once = super::once_callback::OnceCallback::new(on_result);
                {
                    let once = once.clone();
                    dialog.connect_response(move |response_id| {
                        once.claim(|f| {
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
                            true
                        });
                    }).ok();
                }
                {
                    // Enter in either field confirms like Replace (and
                    // closes: the response path's auto-close runs only there).
                    let once = once.clone();
                    let dlg = dialog.clone();
                    let confirm = move || {
                        once.claim(|f| {
                            let find_entry: &CommonEntry = unsafe { &*(find_ptr as *const CommonEntry) };
                            let replace_entry: &CommonEntry = unsafe { &*(replace_ptr as *const CommonEntry) };
                            f(Some((
                                find_entry.get_text().unwrap_or_default(),
                                replace_entry.get_text().unwrap_or_default(),
                            )));
                            dlg.close();
                            true
                        });
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
    choice_dialog(SPECIAL_CHAR_DIALOG_TITLE, SPECIAL_CHAR_DIALOG_OK, items, initial, on_result)
}

/// Generic single-column radio-choice dialog. Shared by the special-char and
/// margin-aggregate pickers so their geometry, keyboard behaviour, and
/// close-then-report sequencing can never drift; only the title and the
/// confirm-button label differ.
pub fn choice_dialog<F: FnOnce(Option<usize>) + 'static>(
    title: &'static str,
    ok_label: &'static str,
    items: &[String],
    initial: usize,
    on_result: F,
) {
    choice_dialog_parented(title, ok_label, items, initial, std::ptr::null_mut(), on_result)
}

/// As [`choice_dialog`], but parented to `parent` so the window manager
/// places it as a child dialog (centred on the main window) rather than
/// wherever an unparented window happens to land (often a screen corner).
///
/// The parameters are only consumed with the `gui` feature: without it this
/// degrades to "no dialog, report cancel" (see the `#[cfg]` block), so they
/// are unused there by design.
#[cfg_attr(not(any(feature = "gui", feature = "gui-core")), allow(unused_variables))]
pub fn choice_dialog_parented<F: FnOnce(Option<usize>) + 'static>(
    title: &'static str,
    ok_label: &'static str,
    items: &[String],
    initial: usize,
    parent: *mut std::os::raw::c_void,
    on_result: F,
) {
    #[cfg(any(feature = "gui", feature = "gui-core"))]
    {
        use super::dialog_widgets::DropDown;
        if let Ok(rxapp) = rswidgets::App::init() {
            let refs: Vec<&str> = items.iter().map(String::as_str).collect();
            if let (Ok(dialog), Ok(dropdown)) =
                (rxapp.new_dialog(), rxapp.create_dropdown(&refs))
            {
                dialog.set_title(title);
                // A dropdown needs only one row: a compact box, not the
                // tall radio list (the "huge dialog" report). Width follows
                // the longest item so nothing is clipped.
                let widest = items.iter().map(|i| i.chars().count()).max().unwrap_or(12);
                let w = ((widest as i32 + 8) * 8).clamp(180, 420);
                dialog.set_default_size(w, 110);
                // Place as a child of the main window when we have it: an
                // unparented dialog is left at a screen corner by the WM.
                if !parent.is_null() {
                    dialog.set_transient_for(parent);
                }
                dropdown.set_hexpand(true);
                dropdown.set_active(Some(initial as u32));
                dialog.append_content_area(&dropdown);
                dialog.add_button("Cancel", 0);
                dialog.add_button(ok_label, 1);
                // Enter confirms (the dropdown's own popup handles Enter
                // while it is open, so this only fires once it is closed).
                dialog.inner.set_default_response(1);
                let dd_ptr = Box::into_raw(Box::new(dropdown)) as usize;
                let once = super::once_callback::OnceCallback::new(on_result);
                // Close FIRST, then report (same sequencing as the radio
                // version): teardown restores focus synchronously, so a
                // callee's focus grab lands instead of racing the destroy.
                let dlg_close = dialog.clone();
                dialog.connect_response(move |response_id| {
                    once.claim(|f| {
                        let dd: &DropDown = unsafe { &*(dd_ptr as *const DropDown) };
                        let result = if response_id == 1 {
                            let idx = dd.get_active() as i64;
                            if idx < 0 { Some(initial) } else { Some(idx as usize) }
                        } else {
                            None
                        };
                        dlg_close.close();
                        f(result);
                        true
                    });
                }).ok();
                dialog.present();
                let dd: &DropDown = unsafe { &*(dd_ptr as *const DropDown) };
                dd.grab_focus();
                let _ = Box::into_raw(Box::new(dialog));
                return;
            }
        }
    }
    on_result(None);
}

pub fn sort_dialog<F: FnOnce(Option<(usize, bool)>) + 'static>(_workbook: &WorkbookState, on_result: F) {
    #[cfg(any(feature = "gui", feature = "gui-core"))]
    {
        use super::dialog_widgets::{CheckButton, DropDown};
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
                let once = super::once_callback::OnceCallback::new(on_result);
                dialog.connect_response(move |response_id| {
                    once.claim(|f| {
                        let sort_col: &DropDown = unsafe { &*(sort_col_ptr as *const DropDown) };
                        let ascending: &CheckButton = unsafe { &*(ascending_ptr as *const CheckButton) };
                        if response_id == 1 {
                            let col = sort_col.get_active().max(0) as usize;
                            let asc = ascending.is_active();
                            f(Some((col, asc)));
                        } else {
                            f(None);
                        }
                        true
                    });
                }).ok();
                dialog.present();
                let _ = Box::into_raw(Box::new(dialog));
                return;
            }
        }
    }
    on_result(None);
}

/// Sort View dialog: a text entry for multi-column sort specs with a
/// column-picker dropdown populated from the workbook's non-empty columns.
///
/// The entry is pre-filled with an example showing two-column sort syntax.
/// Selecting a column from the dropdown and clicking "Add" appends it to the
/// entry (with `!` prefix when "Descending" is checked). The user can also
/// edit the entry directly.
///
/// Returns the raw sort-spec string (e.g. `"A,B"`, `"!C,D"`) on Sort, or
/// `None` on Cancel.
///
/// `workbook` is only consumed with the `gui` feature (it derives the column
/// list); without it this degrades to "no dialog, report cancel", so it is
/// unused there by design.
#[cfg_attr(not(any(feature = "gui", feature = "gui-core")), allow(unused_variables))]
pub fn sort_view_dialog<F: FnOnce(Option<String>) + 'static>(
    workbook: &WorkbookState,
    on_result: F,
) {
    #[cfg(any(feature = "gui", feature = "gui-core"))]
    {
        use crate::grid::CellAddr;
        use rswidgets::common::{Entry as CommonEntry, Orientation};

        use super::dialog_widgets::{CheckButton, DropDown};

        // Collect non-empty columns from the active sheet.
        let sheet = workbook.active_sheet();
        let mc = sheet.grid.main_cols();
        let mut col_names: Vec<String> = Vec::new();
        for c in 0..mc {
            let has_content = (0..sheet.grid.main_rows()).any(|r| {
                sheet
                    .grid
                    .get(&CellAddr::Main { row: r as u32, col: c as u32 })
                    .is_some_and(|v| !v.is_empty())
            });
            if has_content {
                col_names.push(crate::addr::excel_column_name(c));
            }
        }
        // Fall back to at least A–E so the dialog is never empty.
        if col_names.is_empty() {
            for c in 0..mc.min(5) {
                col_names.push(crate::addr::excel_column_name(c));
            }
        }
        let col_refs: Vec<&str> = col_names.iter().map(|s| s.as_str()).collect();

        if let Ok(rxapp) = rswidgets::App::init() {
            let built = (
                rxapp.new_dialog(),
                rxapp.new_entry(),
                rxapp.create_dropdown(&col_refs),
                rxapp.create_checkbutton("Descending"),
                rxapp.new_box(Orientation::Vertical, 6),
                rxapp.new_box(Orientation::Horizontal, 4),
            );
            if let (Ok(dialog), Ok(entry), Ok(dd), Ok(desc_chk), Ok(vbox), Ok(picker)) = built {
                dialog.set_title("Sort view");
                dialog.set_default_size(380, 220);

                // Hint with example syntax.
                if let Ok(l) = rxapp.new_label(
                    "Sort by columns (left = primary).  e.g. A,B sorts by A then B; \
                     ! prefix = descending (e.g. !A,B).",
                ) {
                    vbox.append(&l);
                }

                // Editable sort-spec entry.
                entry.set_hexpand(true);
                // Pre-fill with an example using the first two non-empty columns,
                // or the first column if only one exists.
                let example = if col_names.len() >= 2 {
                    format!("{},{}", col_names[0], col_names[1])
                } else if !col_names.is_empty() {
                    col_names[0].clone()
                } else {
                    "A".into()
                };
                entry.set_text(&example);
                vbox.append(&entry);

                // Column picker row: dropdown + descending checkbox + Add button.
                if let Ok(l) = rxapp.new_label("Add column:") {
                    vbox.append(&l);
                }
                dd.set_hexpand(true);
                picker.append(&dd);
                picker.append(&desc_chk);
                vbox.append(&picker);

                dialog.append_content_area(&vbox);
                dialog.add_button("Cancel", 0);
                dialog.add_button("Sort", 1);

                // Leak widgets so the closures can reference them.
                let entry_ptr = Box::into_raw(Box::new(entry)) as usize;
                let dd_ptr = Box::into_raw(Box::new(dd)) as usize;
                let desc_ptr = Box::into_raw(Box::new(desc_chk)) as usize;

                // Wire the dropdown change to auto-append the selected column
                // to the entry when the user picks one.
                {
                    let dd_ref: &DropDown = unsafe { &*(dd_ptr as *const DropDown) };
                    // Connect dropdown "changed" signal: when user selects a column.
                    let entry_ptr2 = entry_ptr;
                    let dd_ptr2 = dd_ptr;
                    let desc_ptr2 = desc_ptr;
                    dd_ref.connect_changed(move || {
                        let entry: &CommonEntry = unsafe { &*(entry_ptr2 as *const CommonEntry) };
                        let dd: &DropDown = unsafe { &*(dd_ptr2 as *const DropDown) };
                        let desc: &CheckButton = unsafe { &*(desc_ptr2 as *const CheckButton) };
                        let idx = dd.get_active() as i64;
                        if idx < 0 {
                            return;
                        }
                        let idx = idx as usize;
                        if idx >= col_names.len() {
                            return;
                        }
                        let name = &col_names[idx];
                        let mut current = entry.get_text().unwrap_or_default();
                        if !current.is_empty() && !current.ends_with(',') {
                            current.push(',');
                        }
                        if desc.is_active() {
                            current.push('!');
                        }
                        current.push_str(name);
                        entry.set_text(&current);
                    }).ok();
                }

                // Wire OK/Cancel.
                let entry_ptr2 = entry_ptr;
                let once = super::once_callback::OnceCallback::new(on_result);
                dialog.connect_response(move |response_id| {
                    once.claim(|f| {
                        if response_id == 1 {
                            let entry: &CommonEntry = unsafe { &*(entry_ptr2 as *const CommonEntry) };
                            let text = entry.get_text().unwrap_or_default();
                            if text.trim().is_empty() {
                                f(None);
                            } else {
                                f(Some(text));
                            }
                        } else {
                            f(None);
                        }
                        true
                    });
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

/// The Balance Books dialog: column, report type and direction.
///
/// Mirrors the TUI's `Mode::BalanceBooks` field for field
/// (`src/ui/mod.rs`), because a one-field prompt silently discarded two of the
/// three choices: `run_prompt_action` hardcoded `PosToNeg` and `persist = true`,
/// so a GUI user could not ask for the other direction or for a view-only
/// report. The TUI exposes all three, and this dialog matches it:
///
///   * `Column to Balance:` — which numeric column scores the rows, offered as
///     a drop-down of `column_choices` (the caller passes the non-blank,
///     numeric-bearing columns, so the list cannot point at a column that
///     would balance to nothing);
///   * `Report Type:` — View only (default, as in the TUI) or Persisted report;
///   * `Balance direction:` — PosToNeg or NegToPos;
///   * `Generate` / `Cancel`.
///
/// `column_choices` are the labels the drop-down shows; `initial_column` is the
/// index preselected (the caller's auto-detected column, or 0). An empty
/// `column_choices` still opens the dialog — the column then stays empty and
/// `run_balance_books` falls back to its own auto-detection.
///
/// Returns `None` on cancel.
pub fn balance_books_dialog<F: FnOnce(Option<BalanceChoice>) + 'static>(
    column_choices: &[String],
    initial_column: usize,
    on_result: F,
) {
    #[cfg(any(feature = "gui", feature = "gui-core"))]
    {
        use rswidgets::common::Orientation;

        // The widget types are resolved per backend by `dialog_widgets`
        // (gtk on Linux, nwg on Windows, the prelude elsewhere), so this
        // import is valid on every target — do not narrow it by target_os.
        use super::dialog_widgets::{CheckButton, DropDown, RadioButton};

        if let Ok(rxapp) = rswidgets::App::init() {
            // A drop-down needs at least one row to be usable; when no numeric
            // column exists fall back to a single placeholder so the widget is
            // well-formed, and the commit treats it as "no explicit column"
            // (run_balance_books then auto-detects and reports honestly).
            let labels: Vec<&str> = if column_choices.is_empty() {
                vec!["(no numeric column found)"]
            } else {
                column_choices.iter().map(String::as_str).collect()
            };
            let no_choices = column_choices.is_empty();
            let built = (
                rxapp.new_dialog(),
                rxapp.create_dropdown(&labels),
                rxapp.create_checkbutton("Persisted report"),
                rxapp.create_radiobutton(None, "Match +ve number with multiple -ve numbers"),
                rxapp.new_box(Orientation::Vertical, 6),
            );
            if let (Ok(dialog), Ok(column), Ok(persist), Ok(pos_neg), Ok(vbox)) = built {
                dialog.set_title("Balance books");
                dialog.set_default_size(420, 260);

                if let Ok(l) = rxapp.new_label(
                    "Balance rows into groups that sum to zero. The selected numeric column \
                     is used to score rows; all other columns are copied unchanged.",
                ) {
                    vbox.append(&l);
                }
                if let Ok(l) = rxapp.new_label("Column to Balance:") {
                    vbox.append(&l);
                }
                column.set_hexpand(true);
                // Preselect the caller's column, clamped into range.
                let initial = if no_choices {
                    0
                } else {
                    initial_column.min(column_choices.len() - 1)
                };
                column.set_active(Some(initial as u32));
                vbox.append(&column);

                if let Ok(l) = rxapp.new_label("Report Type:") {
                    vbox.append(&l);
                }
                // View only is the TUI default (`persist: false`), so the
                // checkbox starts clear.
                persist.set_active(false);
                vbox.append(&persist);

                if let Ok(l) = rxapp.new_label("Balance direction:") {
                    vbox.append(&l);
                }
                pos_neg.set_active(true);
                vbox.append(&pos_neg);
                // A second radio in the same group; the backend draws the
                // exclusivity.
                let neg_pos = rxapp.create_radiobutton(Some(&pos_neg), "Match -ve number with multiple +ve numbers");
                if let Ok(ref rb) = neg_pos {
                    vbox.append(rb);
                }

                dialog.append_content_area(&vbox);
                dialog.add_button("Cancel", 0);
                dialog.add_button("Generate", 1);

                // The widgets outlive the callback, so they are moved to the
                // heap and referred to by pointer (the same lifetime trick
                // `sort_dialog` uses).
                let column_ptr = Box::into_raw(Box::new(column)) as usize;
                let persist_ptr = Box::into_raw(Box::new(persist)) as usize;
                let neg_pos_ptr = neg_pos.ok().map(|r| Box::into_raw(Box::new(r)) as usize);
                // The response closure is `'static`, so it owns a copy of the
                // choices it maps the drop-down index back through.
                let column_choices_owned: Vec<String> = column_choices.to_vec();
                let once = super::once_callback::OnceCallback::new(on_result);
                dialog
                    .connect_response(move |response_id| {
                        once.claim(|f| {
                            if response_id != 1 {
                                f(None);
                                return true;
                            }
                            let column: &DropDown = unsafe { &*(column_ptr as *const DropDown) };
                            let persist: &CheckButton =
                                unsafe { &*(persist_ptr as *const CheckButton) };
                            let direction = match neg_pos_ptr {
                                Some(p) => {
                                    let rb: &RadioButton = unsafe { &*(p as *const RadioButton) };
                                    if rb.is_active() {
                                        BalanceDirection::NegToPos
                                    } else {
                                        BalanceDirection::PosToNeg
                                    }
                                }
                                None => BalanceDirection::PosToNeg,
                            };
                            // The drop-down reports an index; map it back to
                            // the column label. With no numeric columns the
                            // placeholder row must not masquerade as a column
                            // name, so the selection is left empty for
                            // run_balance_books to auto-detect (and report if
                            // it still finds nothing).
                            let idx = column.get_active() as i64;
                            let selected = if no_choices || idx < 0 {
                                String::new()
                            } else {
                                column_choices_owned
                                    .get(idx as usize)
                                    .cloned()
                                    .unwrap_or_default()
                            };
                            f(Some(BalanceChoice {
                                column: selected,
                                persist: persist.is_active(),
                                direction,
                            }));
                            true
                        });
                    })
                    .ok();
                dialog.present();

                // The TUI focuses the column field first, so the dialog can be
                // used without reaching for the mouse.
                let column: &DropDown = unsafe { &*(column_ptr as *const DropDown) };
                column.grab_focus();
                let _ = Box::into_raw(Box::new(dialog));
                return;
            }
        }
    }
    let _ = (column_choices, initial_column);
    on_result(None);
}

/// What the Balance Books dialog collected.
///
/// A named struct rather than a tuple: three fields where two are booleans/enums
/// read as an opaque `(String, bool, BalanceDirection)` at every call site.
#[derive(Clone, Debug)]
pub struct BalanceChoice {
    /// Column expression the user typed (`"A"`, `"B"`, ...); parsed by the
    /// caller, which can also fall back to the auto-chosen column when empty.
    pub column: String,
    /// Persisted report (TUI's `persist`).
    pub persist: bool,
    /// Which sign pairing to balance against.
    pub direction: BalanceDirection,
}
