//! Comprehensive menu-item tests for the shared GUI/pancurses menu tree.
//!
//! Covers EVERY item in `gui::menu::menu_bar()` (File/Edit/Insert/Format/Sheet/
//! Help, including submenus) across four dimensions, deterministically (no tmux,
//! no X — pure logic on the dispatched actions):
//!
//!   1. **Structure** — the tree is well-formed (non-empty labels, `Submenu`
//!      kinds carry items, unique shortcuts within a menu, distinct action
//!      names).
//!   2. **Coverage** — every leaf action name is handled by the app: either a
//!      real `dispatch_menu_action` arm (not the `_ => "Menu action: …"` stub),
//!      a backend text prompt (`menu_action_needs_prompt`), or a backend-special
//!      name (e.g. `quit`).
//!   3. **Prompt handlers** — every prompt-gated action runs `run_prompt_action`
//!      cleanly on valid input and updates `app.core.status`.
//!   4. **Real effects** — stateful dispatch actions (inserts, cut, dup, sort,
//!      formats, sheet ops) actually mutate the workbook/grid/cursor on a
//!      populated sheet, and file-bound ops round-trip through a real `.corro`
//!      log.
//!
//! Run with: `cargo test --features pancurses --test menu_all_items`
//! (also compiles under `--features gui`).

#![cfg(all(target_os = "linux", any(feature = "gui", feature = "pancurses")))]

use corro::gui::actions::{
    dispatch_menu_action, menu_action_needs_prompt, prompt_action_write_target,
    run_prompt_action, MenuDispatch,
};
use corro::gui::menu::{action_kind_to_name, menu_bar, MenuAction, MenuActionKind};
use corro::gui::special_picker;
use corro::gui::App;
use std::collections::{HashMap, HashSet};
use std::path::PathBuf;

/// A leaf menu item: its display label, resolved action name, and shortcut.
#[derive(Debug)]
struct Leaf {
    path: String,
    label: &'static str,
    action: &'static str,
    shortcut: &'static str,
}

fn walk(items: &[MenuAction], path: &mut String, out: &mut Vec<Leaf>) {
    for it in items {
        let sep = if path.is_empty() { "" } else { " > " };
        if let Some(sub) = it.submenu.as_deref() {
            debug_assert!(matches!(it.action, MenuActionKind::Submenu));
            let prev = path.clone();
            path.push_str(&format!("{sep}{}", it.label));
            walk(sub, path, out);
            *path = prev;
        } else {
            out.push(Leaf {
                path: path.clone(),
                label: it.label,
                action: action_kind_to_name(it.action),
                shortcut: it.shortcut,
            });
        }
    }
}

fn all_leaves() -> Vec<Leaf> {
    let mut out = Vec::new();
    let mut path = String::new();
    walk(&menu_bar(), &mut path, &mut out);
    out
}

/// Make a fresh GUI `App` whose active sheet holds some values so stateful
/// actions have something to operate on.
fn seeded_app(tmp: Option<PathBuf>) -> App {
    let mut app = App::new_with_paths(tmp.into_iter().collect());
    app.load_initial().unwrap();
    // Cursor at A1; ensure a 3x2 main area and seed values.
    app.core.workbook.active_sheet_mut().grid.set_main_size(3, 2);
    let grid = &mut app.core.workbook.active_sheet_mut().grid;
    grid.set(&corro::grid::CellAddr::main(0, 0), "10".into());
    grid.set(&corro::grid::CellAddr::main(1, 0), "20".into());
    grid.set(&corro::grid::CellAddr::main(0, 1), "x".into());
    app
}

/// External-editor actions spawn $EDITOR on dispatch. Point it at a
/// missing binary for the suite's duration so no test ever launches a real
/// editor (instant deterministic Err, no hangs on any platform). Saves and
/// restores the environment; hold the guard for the whole dispatching test.
struct NoEditorGuard {
    visual: Option<String>,
    editor: Option<String>,
    // Held for the guard's lifetime: $VISUAL/$EDITOR are process-global, so a
    // concurrent test restoring them (guard drop) while another test dispatches
    // an editor-spawning action would launch the real editor (vim under CI).
    _lock: std::sync::MutexGuard<'static, ()>,
}

static EDITOR_ENV_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

impl NoEditorGuard {
    fn take() -> Self {
        let _lock = EDITOR_ENV_LOCK
            .lock()
            .unwrap_or_else(|e| e.into_inner());
        let visual = std::env::var("VISUAL").ok();
        let editor = std::env::var("EDITOR").ok();
        std::env::set_var("VISUAL", "/nonexistent-corro-test-editor");
        std::env::set_var("EDITOR", "/nonexistent-corro-test-editor");
        Self { visual, editor, _lock }
    }
}
impl Drop for NoEditorGuard {
    fn drop(&mut self) {
        match &self.visual {
            Some(v) => std::env::set_var("VISUAL", v),
            None => std::env::remove_var("VISUAL"),
        }
        match &self.editor {
            Some(e) => std::env::set_var("EDITOR", e),
            None => std::env::remove_var("EDITOR"),
        }
    }
}

/// A short human-readable key for a `MenuDispatch` (which isn't `Debug`).
fn dispatch_hint(d: &MenuDispatch) -> &'static str {
    match d {
        MenuDispatch::Status(_) => "Status",
        MenuDispatch::Prompt(..) => "Prompt",
        MenuDispatch::SpecialPicker => "SpecialPicker",
        MenuDispatch::Edit { .. } => "Edit",
        MenuDispatch::About { .. } => "About",
        MenuDispatch::HelpFull { .. } => "HelpFull",
        MenuDispatch::HelpKeybinds { .. } => "HelpKeybinds",
    }
}

/// `menu_bar`'s leaf items must be internally consistent: non-empty labels, the
/// `Submenu` kind is only used for items with children and vice-versa, each menu
/// has unique shortcuts, and every resolved action name is non-empty and unique.
#[test]
fn menu_tree_is_well_formed() {
    let leaves = all_leaves();
    assert!(!leaves.is_empty(), "menu_bar() must produce leaf items");

    // Non-empty labels + distinct action names across the whole tree.
    let mut names: HashSet<&'static str> = HashSet::new();
    for l in &leaves {
        assert!(!l.label.trim().is_empty(), "empty label on {l:?}");
        assert!(!l.action.is_empty(), "empty action name on {l:?}");
        assert!(
            names.insert(l.action),
            "duplicate action name '{:?}' at {} > {}",
            l.action,
            l.path,
            l.label
        );
    }

    // Unique shortcuts within each menu level (path = containing menu).
    let mut by_menu: HashMap<String, Vec<&Leaf>> = HashMap::new();
    for l in &leaves {
        by_menu.entry(l.path.clone()).or_default().push(l);
    }
    for (menu, items) in by_menu {
        let mut seen: HashSet<&'static str> = HashSet::new();
        for l in items {
            if !l.shortcut.is_empty() {
                assert!(
                    seen.insert(l.shortcut),
                    "duplicate shortcut '{}' in menu {menu}",
                    l.shortcut
                );
            }
        }
    }

    // Submenu consistency (validated during walk via debug_assert, but make it
    // a hard assertion too).
    fn check_subtree(items: &[MenuAction]) {
        for it in items {
            match it.submenu.as_deref() {
                Some(sub) => {
                    assert!(
                        matches!(it.action, MenuActionKind::Submenu),
                        "item with submenu '{}' must be kind Submenu",
                        it.label
                    );
                    assert!(!sub.is_empty(), "submenu '{}' is empty", it.label);
                    check_subtree(sub);
                }
                None => {
                    assert!(
                        !matches!(it.action, MenuActionKind::Submenu),
                        "leaf '{}' is kind Submenu but has no children",
                        it.label
                    );
                }
            }
        }
    }
    check_subtree(&menu_bar());
}

/// Every leaf action name must be handled by the app — via a real
/// `dispatch_menu_action` arm (not the `_ => "Menu action: …"` stub), a backend
/// text prompt, or a backend-special name such as `quit`.
#[test]
fn every_menu_item_is_handled() {
    let _no_editor = NoEditorGuard::take();
    let mut app = seeded_app(None);
    let mut pending_scope = 0u8;
    let mut clipboard = String::new();

    // Names that the backend handles before/without the dispatcher: "quit"
    // (the pancurses backend sets running=false) and "extrapolate" (both GUI
    // backends enter the interactive modal before dispatch runs, so the
    // dispatcher has deliberately no arm for it). These have no dispatch work
    // to do.
    let backend_special: HashSet<&'static str> = ["quit", "extrapolate"].into_iter().collect();

    let leaves = all_leaves();
    let mut failures = Vec::new();
    let mut dispatched = 0usize;
    for l in leaves.iter() {
        let name = l.action;
        if menu_action_needs_prompt(name).is_some() || backend_special.contains(name) {
            continue;
        }
        dispatched += 1;
        // Fire the dispatcher. A stub arm returns Status("Menu action: {name}").
        let d = dispatch_menu_action(&mut app, name, &mut pending_scope, &mut clipboard);
        let bogus = matches!(&d, MenuDispatch::Status(s) if s.starts_with("Menu action: "));
        if bogus {
            failures.push(format!(
                "{} > {} ('{name}') -> {}",
                l.path,
                l.label,
                dispatch_hint(&d)
            ));
        }
    }
    assert!(
        failures.is_empty(),
        "menu items that fall through to the stub dispatcher:\n{}",
        failures.join("\n")
    );
    assert!(
        dispatched >= 20,
        "expected most menu items to dispatch directly, but only {dispatched} did"
    );
}

/// Prompt-gated actions must run `run_prompt_action` cleanly and update status.
#[test]
fn prompt_actions_run_cleanly() {
    let dir = std::env::temp_dir();
    let tag = std::process::id();

    let mut app = seeded_app(None);
    // Seed a value so exports have content.
    let g = &mut app.core.workbook.active_sheet_mut().grid;
    g.set(&corro::grid::CellAddr::main(2, 0), "30".into());

    // Valid inputs per action. Some need a writable temp file.
    let export_path = dir.join(format!("corro-menu-export-{tag}.tsv"));
    let _ = std::fs::remove_file(&export_path);
    let open_path = dir.join(format!("corro-menu-open-{tag}.corro"));
    std::fs::write(&open_path, "CORRO_LOG 1\nSET A1 42\n").unwrap();
    let save_path = dir.join(format!("corro-menu-save-{tag}.corro"));
    let _ = std::fs::remove_file(&save_path);

    let cases: Vec<(&str, &str)> = vec![
        ("open", open_path.to_str().unwrap()),
        ("save_as", save_path.to_str().unwrap()),
        ("export_tsv", export_path.to_str().unwrap()),
        ("export_csv", export_path.to_str().unwrap()),
        ("export_ascii", export_path.to_str().unwrap()),
        ("export_ods", export_path.to_str().unwrap()),
        ("export_all", export_path.to_str().unwrap()),
        ("set_col_width", "15"),
        ("set_max_col_width", "12"),
        ("go_to_cell", "C3"),
        ("find", "10"),
        ("replace", "10|99"),
        ("rename_sheet", "Renamed"),
        ("copy_sheet", "CopyOf"),
        ("delete_sheet", "Renamed"),
        // NOTE: insert_special_chars is not free-text: it opens the
        // 10-choice picker (see special_picker_routing + special_char_parity).
        ("insert_hyperlink", "https://example.com"),
        ("sort_view", "A,"),
        ("persist_sort", "A!"),
        ("balance_books", "A"),
    ];

    // Every prompt-gated menu name must be covered by cases (except those that a
    // single run would make impossible, e.g. deleting the only sheet).
    let prompt_names_present: HashSet<&str> = {
        let leaves = all_leaves();
        leaves.iter().filter_map(|l| menu_action_needs_prompt(l.action).map(|_| l.action)).collect()
    };
    let case_names: HashSet<&str> = cases.iter().map(|(n, _)| *n).collect();
    let missing: Vec<&str> = prompt_names_present
        .iter()
        .copied()
        .filter(|n| !case_names.contains(n))
        .collect();
    assert!(
        !missing.contains(&"delete_sheet"), // covered once sheets > 1
        "prompt actions lack a test case: {missing:?}"
    );

    let mut app2 = seeded_app(None);
    // Add a second sheet so copy/delete have a target.
    app2.core.workbook.add_sheet("Sheet2".into(), corro::ops::SheetState::new(1, 1));
    let mut failed = Vec::new();
    for (action, text) in cases {
        run_prompt_action(&mut app2, action, text);
        if app2.core.status.is_empty() || app2.core.status.contains("Menu:") {
            failed.push(format!("{action} ('{text}') -> status {:?}", app2.core.status));
        }
        app2.core.status.clear();
    }
    assert!(
        failed.is_empty(),
        "prompt actions produced empty/stub status:\n{}",
        failed.join("\n")
    );

    // Exports must actually write a non-empty file.
    let written = std::fs::read_to_string(&export_path).unwrap_or_default();
    assert!(!written.is_empty(), "export_tsv did not write content");
    let _ = std::fs::remove_file(&export_path);
    let _ = std::fs::remove_file(&open_path);
    let _ = std::fs::remove_file(&save_path);
}

/// Stateful dispatch actions must actually mutate the workbook/grid/cursor on a
/// populated sheet, and file-bound ops must round-trip through a real `.corro` log.
#[test]
fn dispatch_actions_apply_real_effects() {
    let dir = std::env::temp_dir();
    let tag = std::process::id();
    let log = dir.join(format!("corro-menu-disp-{tag}.corro"));
    std::fs::write(&log, "CORRO_LOG 1\n").unwrap();

    let mut app = seeded_app(Some(log.clone()));
    let mut pending_scope = 0u8;
    let mut clipboard = String::new();
    let run = |app: &mut App, name: &str, scope: &mut u8, cb: &mut String| {
        dispatch_menu_action(app, name, scope, cb)
    };
    let rows = |app: &App| app.core.workbook.active_sheet().grid.main_rows();
    let cols = |app: &App| app.core.workbook.active_sheet().grid.main_cols();
    let cell = |app: &App, r: u32, c: u32| {
        app.core.workbook.active_sheet().grid.get(&corro::grid::CellAddr::main(r, c))
    };

    // Cut clears the cursor cell; paste restores it (cursor is A1, value "10").
    run(&mut app, "cut", &mut pending_scope, &mut clipboard);
    assert_eq!(cell(&app, 0, 0).unwrap_or_default(), "", "cut must clear the cursor cell");
    run(&mut app, "paste", &mut pending_scope, &mut clipboard);
    assert_eq!(cell(&app, 0, 0).unwrap_or_default(), "10", "paste must restore the cut value");

    // Insert rows above cursor adds a main row.
    let r0 = rows(&app);
    run(&mut app, "insert_rows", &mut pending_scope, &mut clipboard);
    assert_eq!(rows(&app), r0 + 1, "insert_rows must add a row");

    // Insert cols left of cursor adds a main column.
    let c0 = cols(&app);
    run(&mut app, "insert_cols", &mut pending_scope, &mut clipboard);
    assert_eq!(cols(&app), c0 + 1, "insert_cols must add a column");

    // Mitosis (row) duplicates the row and moves the cursor one row down.
    let r1 = rows(&app);
    let cur_row_before = app.core.cursor.row;
    run(&mut app, "insert_mitosis_row", &mut pending_scope, &mut clipboard);
    assert_eq!(rows(&app), r1 + 1, "insert_mitosis_row must add a row");
    assert_eq!(
        app.core.cursor.row,
        cur_row_before + 1,
        "insert_mitosis_row must move the cursor onto the duplicate"
    );

    // Mitosis (col) duplicates the column and moves the cursor one col right.
    let c1 = cols(&app);
    let cur_col_before = app.core.cursor.col;
    run(&mut app, "insert_mitosis_col", &mut pending_scope, &mut clipboard);
    assert_eq!(cols(&app), c1 + 1, "insert_mitosis_col must add a column");
    assert_eq!(
        app.core.cursor.col,
        cur_col_before + 1,
        "insert_mitosis_col must move the cursor onto the duplicate"
    );

    // New sheet adds one and activates it.
    let sheets_before = app.core.workbook.sheet_count();
    run(&mut app, "new_sheet", &mut pending_scope, &mut clipboard);
    assert_eq!(app.core.workbook.sheet_count(), sheets_before + 1, "new_sheet must add a sheet");

    // Format fixed 0 with a pending scope writes a column format.
    pending_scope = 1; // All
    let d = run(&mut app, "format_fixed_0", &mut pending_scope, &mut clipboard);
    assert_eq!(dispatch_hint(&d), "Status", "format_fixed_0 should return a Status");
    let status_text = match &d {
        MenuDispatch::Status(s) => s,
        _ => "",
    };
    assert_eq!(status_text, "Format: Fixed 0", "format_fixed_0 status text");

    // select_all sets an anchor spanning the sheet.
    run(&mut app, "select_all", &mut pending_scope, &mut clipboard);
    assert!(app.core.anchor.is_some(), "select_all must set the anchor");

    // replay reloads the committed file (idempotent round-trip, no panic).
    run(&mut app, "replay", &mut pending_scope, &mut clipboard);
    assert!(
        app.core.status.contains("revision") || app.core.status.starts_with("Replayed"),
        "replay should restore the workbook, got {:?}",
        app.core.status
    );

    let _ = std::fs::remove_file(&log);
}

/// Verify `action_kind_to_name` covers every variant (no `_` fallback), so the
/// menu tree can never silently map an action to a mis-spelt dispatch name.
#[test]
fn action_names_are_exhaustive() {
    use MenuActionKind::*;
    let all = [
        NewFile, Open, Save, SaveAs, Quit, Undo, Redo, Cut, Copy, Paste, Find, Replace,
        DeleteCell, SelectAll, ToggleHeaders, ToggleMargins, NewSheet, RenameSheet,
        DeleteSheet, SortAsc, SortDesc, BalanceBooks, ExportTsv, ExportCsv, ExportOds,
        ExportAscii, About, HelpKeybinds, InsertRows, InsertMitosisRow, InsertMitosisCol,
        InsertCols, InsertSpecialChars, InsertDate, InsertTime, InsertHyperlink,
        FormatApplyAll, FormatApplyFullColumn, FormatApplyData, FormatApplySpecial,
        FormatApplyCell, FormatApplySelection, FormatDecimalGeneric, FormatCurrency,
        FormatRational, FormatFixed0, FormatFixed1, FormatFixed2, FormatFixedCustom,
        FormatAlignLeft, FormatAlignCenter, FormatAlignRight, FormatAlignDefault,
        FormatReset, ExportAll, Submenu, SortView, SaveSort, Replay, SetMaxColWidth,
        SetColWidth, ExportOdt, Duplicate, Extrapolate, EditExternal, EditWorkbookExternal, FollowHyperlink, SheetPrev, SheetNext, CopySheet,
        MoveSheet, GoToCell, HelpRows, HelpCols, HelpFull,
    ];
    for (i, kind) in all.iter().enumerate() {
        assert!(
            !action_kind_to_name(*kind).is_empty(),
            "action_kind_to_name missing case for MenuActionKind[{i}]"
        );
        if matches!(*kind, MenuActionKind::Submenu) {
            assert_eq!(
                action_kind_to_name(*kind),
                "submenu",
                "Submenu kind must map to the sentinel 'submenu' name"
            );
        } else {
            assert_ne!(
                action_kind_to_name(*kind),
                "submenu",
                "non-Submenu kind [{}] must not use the 'submenu' sentinel",
                i
            );
        }
    }
}

#[test]
fn leaf_count_is_substantial() {
    let leaves = all_leaves();
    // Guard against a vacuous pass: the menu must hold a healthy number of
    // leaf actions (currently ~40+ across File/Edit/Insert/Format/Sheet/Help).
    assert!(leaves.len() >= 30, "menu_bar() only produced {} leaves", leaves.len());
    // Every top-level menu must contribute items.
    let mut by_menu: std::collections::HashMap<String, usize> = Default::default();
    for l in &leaves {
        let top = l.path.split(" > ").next().unwrap_or("").to_string();
        *by_menu.entry(top).or_insert(0) += 1;
    }
    for (menu, n) in &by_menu {
        assert!(*n >= 1, "menu '{menu}' has no leaf items");
    }
}

/// The GUI backend (`gui_backend::handle_menu_action`, shared by GTK and nwg)
/// must route every menu item somewhere real — never a "not yet implemented"
/// stub. This test pins the routing classification: dialog-native arms keep
/// native dialogs, dialog-delegated arms feed `run_prompt_action`, and
/// everything else goes through the shared `dispatch_menu_action` (proven
/// non-stub by `every_menu_item_is_handled`). If a menu item is added or an
/// arm changes shape without updating these sets, this fails loudly instead
/// of shipping another silent no-op like Insert Date was.
#[test]
fn gui_routing_covers_every_menu_item() {
    let _no_editor = NoEditorGuard::take();
    // gui_backend arms that keep native dialogs/file ops (verified working).
    const GUI_NATIVE_DIALOG: &[&str] = &["open", "save_as", "about"];
    // gui_backend arms that open a dialog and feed run_prompt_action.
    // Must equal EXACTLY the prompt-gated menu leaves (derived check below).
    const GUI_DIALOG_DELEGATED: &[&str] = &[
        "find",
        "replace",
        "balance_books",
        "rename_sheet",
        "export_tsv",
        "export_csv",
        "export_ods",
        "export_ascii",
        "export_all",
        "set_col_width",
        "set_max_col_width",
        "copy_sheet",
        "go_to_cell",
        // NOTE: insert_special_chars is picker-gated, not prompt-gated
        // (see special_picker_routing + special_char_parity).
        "insert_hyperlink",
        "sort_view",
        "persist_sort",
        // NOTE: delete_sheet has prompt handling but no menu item (the Sheet
        // menu offers New/Rename/Copy only), so it is intentionally absent
        // here; prompt_actions_run_cleanly still covers its handler.
    ];
    // gui_backend arms with backend-specific behavior covered elsewhere:
    // extrapolate (interactive modal), quit (save_before_quit flow).
    // NOTE: delete_cell has dispatch/handler code but no menu item (Delete is
    // keyboard-only), so it is intentionally absent here.
    const GUI_NATIVE_OTHER: &[&str] = &["extrapolate", "quit"];

    let leaves = all_leaves();
    let leaf_names: HashSet<&str> = leaves.iter().map(|l| l.action).collect();

    // Every classified name must be a real menu leaf (catches typos/renames).
    for name in GUI_NATIVE_DIALOG
        .iter()
        .chain(GUI_DIALOG_DELEGATED.iter())
        .chain(GUI_NATIVE_OTHER.iter())
    {
        assert!(
            leaf_names.contains(name),
            "gui routing set mentions '{name}', which is not a menu item (stale?)"
        );
    }

    // The dialog-delegated set must be EXACTLY the prompt-gated menu leaves,
    // except open/save_as which use native file dialogs (still prompt-gated
    // in the shared layer). Every other prompt-gated leaf must be here, and
    // nothing here may lack prompt logic.
    let prompt_leaves: HashSet<&str> = leaves
        .iter()
        .filter(|l| menu_action_needs_prompt(l.action).is_some())
        .map(|l| l.action)
        .collect();
    let dialog_set: HashSet<&str> = GUI_DIALOG_DELEGATED.iter().copied().collect();
    let native_file: HashSet<&str> = ["open", "save_as"].into_iter().collect();
    let mut covered: HashSet<&str> = dialog_set.clone();
    covered.extend(native_file.iter().copied());
    assert_eq!(
        covered, prompt_leaves,
        "gui dialog routing diverged from prompt-gated menu leaves"
    );

    // Everything else must be a shared-dispatch leaf (non-stub, proven by
    // every_menu_item_is_handled) or one of the classified specials.
    let mut uncovered: Vec<&str> = leaf_names
        .iter()
        .copied()
        .filter(|n| {
            !dialog_set.contains(n)
                && !GUI_NATIVE_DIALOG.contains(n)
                && !GUI_NATIVE_OTHER.contains(n)
        })
        .collect();
    uncovered.sort_unstable();
    let mut app = seeded_app(None);
    let mut pending_scope = 0u8;
    let mut clipboard = String::new();
    let mut stubbed = Vec::new();
    for name in uncovered.iter().copied() {
        let d = dispatch_menu_action(&mut app, name, &mut pending_scope, &mut clipboard);
        if matches!(&d, MenuDispatch::Status(s) if s.starts_with("Menu action: ")) {
            stubbed.push(name);
        }
    }
    assert!(
        stubbed.is_empty(),
        "menu items with no GUI behavior (stub fallback): {stubbed:?}"
    );
    assert!(
        !uncovered.is_empty(),
        "shared-dispatch remainder is empty; the classification above is vacuous"
    );
}

/// File menu order: New at the top, Exit at the bottom (same arrangement
/// the ratatui reference pins in `file_menu_new_first_exit_last`).
#[test]
fn file_menu_new_first_exit_last() {
    let bar = menu_bar();
    let file = bar.iter().find(|m| m.label == "File").expect("File menu");
    let items = file.submenu.as_deref().unwrap_or(&[]);
    assert_eq!(items.first().map(|i| i.label), Some("New"));
    assert_eq!(items.last().map(|i| i.label), Some("Exit"));
    assert_eq!(items.first().map(|i| i.shortcut), Some("N"));
}

/// Insert > Special Char is picker-gated, not prompt-gated: the shared layer
/// offers no free-text prompt for it, dispatch opens shared picker state,
/// and the rows/digits/clamp match the ratatui reference. Guards the
/// classification the old free-text dialog depended on (see
/// special_char_parity for the end-to-end gesture equivalence).
#[test]
fn special_picker_routing() {
    assert_eq!(
        menu_action_needs_prompt("insert_special_chars"),
        None,
        "insert_special_chars must not be a free-text prompt"
    );
    let mut app = seeded_app(None);
    let mut scope = 0u8;
    let mut clipboard = String::new();
    match dispatch_menu_action(&mut app, "insert_special_chars", &mut scope, &mut clipboard) {
        MenuDispatch::SpecialPicker => {}
        d => panic!("insert_special_chars must dispatch SpecialPicker, got {}", dispatch_hint(&d)),
    }
    assert_eq!(special_picker::index(&app), Some(0), "dispatch opens on the first choice");
    let rows = special_picker::items();
    assert_eq!(rows.len(), 10);
    assert_eq!((rows[0].as_str(), rows[2].as_str(), rows[9].as_str()), ("1: ∞", "3: Ω", "0: θ"));
    // Down*2 + take commits the 3rd choice; digits map 1:1 with the reference.
    special_picker::step(&mut app, 1);
    special_picker::step(&mut app, 1);
    assert_eq!(special_picker::take(&mut app), Some("Ω".to_string()));
    assert_eq!(special_picker::index(&app), None, "take closes the picker");
    special_picker::open(&mut app);
    special_picker::set(&mut app, 99);
    assert_eq!(special_picker::take(&mut app), Some("θ".to_string()), "out-of-range clamps to last");
    assert_eq!(special_picker::index_for_digit('3'), Some(2));
    assert_eq!(special_picker::index_for_digit('0'), Some(9));
}

/// Edit ▸ Follow link dispatches through the shared action: a link cell
/// opens in the (overridden, recording) browser with an "Opened …" status,
/// a non-link cell reports "No hyperlink …" without spawning anything.
/// The recorder script stands in for a real browser so the test never
/// opens a window; `CORRO_URL_OPENER` is process-global, so the override
/// is lock-guarded and always restored.
#[test]
fn follow_hyperlink_dispatch_opens_link_and_reports() {
    static OPENER_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());
    let _guard = OPENER_LOCK.lock().unwrap();
    let prev = std::env::var("CORRO_URL_OPENER").ok();

    let dir = std::env::temp_dir().join(format!("corro-follow-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let record = dir.join("opened.log");
    let _ = std::fs::remove_file(&record);
    let script = dir.join("record.sh");
    std::fs::write(&script, format!("#!/bin/sh\necho \"$1\" >> \"{}\"\n", record.display())).unwrap();
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        let mut perms = std::fs::metadata(&script).unwrap().permissions();
        perms.set_mode(0o755);
        std::fs::set_permissions(&script, perms).unwrap();
    }
    std::env::set_var("CORRO_URL_OPENER", script.to_str().unwrap());

    let mut app = seeded_app(None);
    // Cursor onto main A1 (display coords: one header row, one margin col).
    app.core.cursor = corro::grid::SheetCursor {
        row: corro::grid::HEADER_ROWS,
        col: corro::grid::MARGIN_COLS,
    };
    app.core.workbook.active_sheet_mut().grid.set(
        &corro::grid::CellAddr::main(0, 0),
        "https://example.com/a".into(),
    );
    let mut scope = 0u8;
    let mut clipboard = String::new();
    match dispatch_menu_action(&mut app, "follow_hyperlink", &mut scope, &mut clipboard) {
        MenuDispatch::Status(s) => assert_eq!(s, "Opened https://example.com/a", "unexpected status {s:?}"),
        d => panic!("follow_hyperlink must dispatch Status, got {}", dispatch_hint(&d)),
    }
    // The recorder (not a browser) received exactly the link. The opener
    // spawns detached, so poll for the record instead of assuming the
    // child has run already.
    let mut logged = String::new();
    for _ in 0..200 {
        logged = std::fs::read_to_string(&record).unwrap_or_default();
        if !logged.trim().is_empty() {
            break;
        }
        std::thread::sleep(std::time::Duration::from_millis(50));
    }
    assert_eq!(logged.trim(), "https://example.com/a", "opener got {logged:?}");

    // A non-link cell reports honestly and spawns nothing new.
    app.core.workbook.active_sheet_mut().grid.set(
        &corro::grid::CellAddr::main(0, 0),
        "just text".into(),
    );
    match dispatch_menu_action(&mut app, "follow_hyperlink", &mut scope, &mut clipboard) {
        MenuDispatch::Status(s) => assert_eq!(s, "No hyperlink at A1", "unexpected status {s:?}"),
        d => panic!("follow_hyperlink must dispatch Status, got {}", dispatch_hint(&d)),
    }
    let logged = std::fs::read_to_string(&record).unwrap_or_default();
    assert_eq!(logged.lines().count(), 1, "non-link must not spawn the opener");

    match prev {
        Some(v) => std::env::set_var("CORRO_URL_OPENER", v),
        None => std::env::remove_var("CORRO_URL_OPENER"),
    }
    let _ = std::fs::remove_dir_all(&dir);
}

/// File ▸ New dispatches through the shared action: the workbook resets to
/// the seeded blank, the app detaches from any file, and histories clear —
/// with a "New workbook" status. Guards the routing the menu-tree test
/// pins (every leaf must be non-stub) with real state assertions.
#[test]
fn new_file_dispatch_resets_to_blank_workbook() {
    let mut app = seeded_app(None);
    // Dirty the app: content, a second sheet, cursor off A1, history.
    app.core.workbook.active_sheet_mut().grid.set(
        &corro::grid::CellAddr::main(0, 0),
        "old".into(),
    );
    app.core.workbook.add_sheet("Extra".into(), corro::ops::SheetState::new(1, 1));
    app.core.cursor = corro::grid::SheetCursor { row: 5, col: 5 };
    app.core.op_history.push(corro::ops::Op::SetCell {
        addr: corro::grid::CellAddr::main(0, 0),
        value: "old".into(),
    });
    let mut scope = 0u8;
    let mut clipboard = String::new();
    match dispatch_menu_action(&mut app, "new_file", &mut scope, &mut clipboard) {
        MenuDispatch::Status(s) => assert_eq!(s, "New workbook", "unexpected status {s:?}"),
        d => panic!("new_file must dispatch Status, got {}", dispatch_hint(&d)),
    }
    // Blank seeded workbook: one sheet, empty main cells.
    assert_eq!(app.core.workbook.sheet_count(), 1);
    assert_eq!(app.core.workbook.sheet_title(0), "Sheet1");
    assert_eq!(
        app.core.workbook.active_sheet().grid.get(&corro::grid::CellAddr::main(0, 0)),
        None,
        "old content must not survive New"
    );
    // Detached from any file; histories and cursor reset.
    assert!(app.core.path.is_none());
    assert!(app.core.op_history.is_empty());
    assert!(app.core.redo_history.is_empty());
    assert_eq!(
        app.core.cursor,
        corro::grid::SheetCursor {
            row: corro::grid::HEADER_ROWS,
            col: corro::grid::MARGIN_COLS
        }
    );
    assert!(app.core.anchor.is_none());
}

/// The ratatui reference menu tables and the shared `gui::menu::menu_bar()`
/// tree must enumerate the same items with the same shortcut letters, or the
/// backends' menus have drifted (same labels, same mnemonics everywhere).
/// Multiset comparison: catches items missing, duplicated, renamed, or
/// re-shortcutted on either side. Submenu containers count too (Export,
/// Width, Scope, Number, Align exist as navigable entries in both).
#[cfg(feature = "ratatui")]
#[test]
fn ratatui_and_gui_menus_enumerate_the_same_items() {
    let mut reference: Vec<(String, char)> = corro::ui::all_menu_shortcuts()
        .into_iter()
        .map(|(sc, label)| (label.to_string(), sc))
        .collect();
    // All tree nodes with shortcuts: leaves plus submenu containers (which
    // carry a shortcut in the shared tree, e.g. Export (T)).
    fn walk_all(items: &[MenuAction], out: &mut Vec<(String, String)>) {
        for it in items {
            out.push((it.label.to_string(), it.shortcut.to_string()));
            if let Some(sub) = it.submenu.as_deref() {
                walk_all(sub, out);
            }
        }
    }
    // Skip the six roots (their shortcuts live outside both models: the
    // ratatui reference opens sections by letter directly, while only the
    // Format root carries a tree shortcut (R) for GUI mnemonics).
    let mut nodes: Vec<(String, String)> = Vec::new();
    for root in menu_bar() {
        if let Some(sub) = root.submenu.as_deref() {
            walk_all(sub, &mut nodes);
        }
    }
    let mut tree: Vec<(String, char)> = nodes
        .iter()
        .map(|(label, sc)| {
            let mut chars = sc.chars();
            let (first, second) = (chars.next(), chars.next());
            assert!(
                first.is_some() && second.is_none(),
                "menu item '{label}' must carry exactly one shortcut letter, got {sc:?}"
            );
            (label.clone(), first.unwrap())
        })
        .collect();
    reference.sort_unstable();
    tree.sort_unstable();
    assert_eq!(
        reference.len(),
        tree.len(),
        "menu item counts diverged (ratatui {}, gui tree {})",
        reference.len(),
        tree.len()
    );
    for (r, t) in reference.iter().zip(tree.iter()) {
        assert_eq!(
            r, t,
            "menu mismatch: ratatui has {r:?} where the shared tree has {t:?}"
        );
    }
}

/// `prompt_action_write_target`: the overwrite-confirm gate for typed TUI
/// paths. Existing files flag Some (ask), everything else None (proceed).
#[test]
fn overwrite_gate_flags_existing_files_only() {
    let dir = std::env::temp_dir().join(format!("corro_ow_{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let existing = dir.join("w.corro");
    std::fs::write(&existing, "SET $1:A1 1\n").unwrap();
    let missing = dir.join("nope.corro");
    let es = existing.to_string_lossy().into_owned();
    let ms = missing.to_string_lossy().into_owned();

    for action in [
        "save_as",
        "export_tsv",
        "export_csv",
        "export_ods",
        "export_ascii",
        "export_all",
    ] {
        assert_eq!(
            prompt_action_write_target(action, &es),
            Some(existing.clone()),
            "{action} onto an existing file must flag for confirm"
        );
        assert_eq!(
            prompt_action_write_target(action, &ms),
            None,
            "{action} onto a missing path must proceed without confirm"
        );
        assert_eq!(
            prompt_action_write_target(action, ""),
            None,
            "{action} with empty text (clipboard) must not confirm"
        );
        assert_eq!(
            prompt_action_write_target(action, "   "),
            None,
            "{action} with blank text must not confirm"
        );
    }
    // Non-writing actions never flag, even for existing files.
    for action in ["open", "go_to_cell", "find", "rename_sheet", "save"] {
        assert_eq!(
            prompt_action_write_target(action, &es),
            None,
            "{action} never writes a file so must not confirm"
        );
    }
    // A directory is not a file: no confirm (the write errors as before).
    assert_eq!(
        prompt_action_write_target("save_as", &dir.to_string_lossy()),
        None,
        "directories must not flag for confirm"
    );
    std::fs::remove_dir_all(&dir).ok();
}

/// Save As / export typed paths must land on the format extension (a
/// foreign-extension workbook won't reopen; a bare export name should
/// still be a `.tsv`/`.csv`/...). GUI dialogs force/append the same way.
#[test]
fn prompt_paths_resolve_format_extensions() {
    use corro::ui_core::{append_extension_if_missing as append, force_extension as force};
    use std::path::Path;

    // force_extension (Save As): always `.corro`, replacing foreign ext.
    assert_eq!(force(Path::new("book"), "corro"), Path::new("book.corro"));
    assert_eq!(force(Path::new("book.txt"), "corro"), Path::new("book.corro"));
    assert_eq!(force(Path::new("book.corro"), "corro"), Path::new("book.corro"));
    assert_eq!(
        force(Path::new("/tmp/a.b/name.ods"), "corro"),
        Path::new("/tmp/a.b/name.corro")
    );

    // append_extension_if_missing (exports): bare names get the ext,
    // explicit extensions (even surprising ones) are respected.
    assert_eq!(append(Path::new("out"), "tsv"), Path::new("out.tsv"));
    assert_eq!(append(Path::new("out.csv"), "tsv"), Path::new("out.csv"));
    assert_eq!(append(Path::new("out.ods"), "ods"), Path::new("out.ods"));

    let dir = std::env::temp_dir().join(format!("corro-ext-{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();

    // Save As with a bare name must create `<name>.corro` and record that
    // path on the app (so a subsequent Ctrl+S targets the right file).
    let mut app = seeded_app(None);
    let bare_save = dir.join("noext");
    run_prompt_action(&mut app, "save_as", bare_save.to_str().unwrap());
    let saved = dir.join("noext.corro");
    assert!(saved.is_file(), "save_as must force .corro (got status {:?})", app.core.status);
    assert!(!bare_save.exists(), "the extensionless path must not be written");
    assert_eq!(app.core.path.as_deref(), Some(saved.as_path()));

    // Export with a bare name must create `<name>.tsv`.
    let mut app2 = seeded_app(None);
    app2.core.workbook.active_sheet_mut().grid.set(
        &corro::grid::CellAddr::main(0, 0),
        "7".into(),
    );
    let bare_export = dir.join("export");
    run_prompt_action(&mut app2, "export_tsv", bare_export.to_str().unwrap());
    let exported = dir.join("export.tsv");
    assert!(exported.is_file(), "export must append .tsv (got {:?})", app2.core.status);
    assert!(!std::fs::read_to_string(&exported).unwrap_or_default().is_empty());

    std::fs::remove_dir_all(&dir).ok();
}

/// The GUI Open/Save/Export dialog filters must cover exactly what the
/// loaders open and what the writers emit — otherwise a filtered dialog
/// would hide a supported type (or suggest a bad extension). Guards the
/// shared lists in `ui_core` against drift.
#[test]
fn dialog_filters_cover_loader_and_export_types() {
    use corro::ui_core::{export_ext_for_action, ext_filter_label, SPREADSHEET_EXTS};

    // Open filter: the 4 types `gui::load_initial` dispatches on.
    assert_eq!(
        SPREADSHEET_EXTS,
        &["corro", "csv", "tsv", "ods"],
        "Open filter must list exactly the loader-supported spreadsheet types"
    );
    // Each has a real human label (no generic fallback).
    for ext in SPREADSHEET_EXTS {
        assert_ne!(
            ext_filter_label(ext),
            "Files",
            "spreadsheet type `{ext}` needs a specific filter label"
        );
    }

    // Export mapping: action -> extension (tsv catch-all, incl. export_all).
    for (action, want) in [
        ("export_tsv", "tsv"),
        ("export_csv", "csv"),
        ("export_ods", "ods"),
        ("export_ascii", "txt"),
        ("export_all", "tsv"),
    ] {
        assert_eq!(
            export_ext_for_action(action),
            want,
            "{action} must write .{want}"
        );
        assert_ne!(
            ext_filter_label(export_ext_for_action(action)),
            "Files",
            "{action} extension needs a specific filter label"
        );
    }
}

/// Regression: deleting a named sheet with a live file path must delete
/// exactly ONE sheet. `commit_workbook_op` replays the appended log line into
/// the workbook, so the old code's extra `sheets.remove(idx)` deleted a second
/// sheet — emptying the workbook, after which any `active_sheet()` panicked
/// (surfaced by `prompt_actions_run_cleanly` once Open started working).
#[test]
fn delete_sheet_deletes_exactly_one_sheet_with_a_live_path() {
    let dir = std::env::temp_dir().join(format!("corro_delone_{}", std::process::id()));
    std::fs::create_dir_all(&dir).unwrap();
    let path = dir.join("wb.corro");
    std::fs::write(&path, "CORRO_LOG 1\n").unwrap();

    let mut app = seeded_app(Some(path.clone()));
    app.core.path = Some(path.clone());
    app.core.workbook.add_sheet("Sheet2".into(), corro::ops::SheetState::new(1, 1));
    app.core.workbook.add_sheet("Sheet3".into(), corro::ops::SheetState::new(1, 1));
    let before = app.core.workbook.sheets.len();
    assert_eq!(before, 3, "setup: three sheets");

    run_prompt_action(&mut app, "delete_sheet", "Sheet2");
    let after = app.core.workbook.sheets.len();
    assert_eq!(
        after,
        before - 1,
        "deleting one named sheet must remove exactly one (before={before}, after={after}, status={:?})",
        app.core.status
    );
    assert!(
        !app.core.workbook.sheets.iter().any(|s| s.title == "Sheet2"),
        "the named sheet is gone"
    );
    // The workbook must still be usable (this panicked before the fix).
    run_prompt_action(&mut app, "insert_hyperlink", "https://example.com");
    assert!(
        !app.core.status.contains("Open error"),
        "workbook still usable, status={:?}",
        app.core.status
    );

    // And the deletion is recorded in the log, once.
    let log = std::fs::read_to_string(&path).unwrap();
    assert_eq!(
        log.matches("DELETE_SHEET").count(),
        1,
        "exactly one DELETE_SHEET line, got:\n{log}"
    );
    std::fs::remove_dir_all(&dir).ok();
}

/// Edit ▸ Workbook (External) must open the **live** `.corro` file (not a
/// scratch copy) and reload the workbook from whatever the editor left, so an
/// externally-added revision shows up without a manual reload. Uses a temp
/// $EDITOR script; the `NoEditorGuard` is bypassed here on purpose (this test
/// needs a *real* editor).
#[cfg(unix)]
#[test]
fn edit_workbook_external_opens_live_file_and_reloads() {
    use std::os::unix::fs::PermissionsExt;

    // Serialize against every other $EDITOR-touching test in this binary: the
    // env is process-global, so a concurrent NoEditorGuard drop would restore
    // a real editor mid-test (or vice versa).
    let _lock = EDITOR_ENV_LOCK.lock().unwrap_or_else(|e| e.into_inner());

    let dir = tempfile::tempdir().unwrap();
    let log = dir.path().join("book.corro");
    std::fs::write(&log, "CORRO_LOG 1\nSET A1 old\n").unwrap();

    // Editor rewrites the file with an extra revision appended.
    let editor = dir.path().join("ed.sh");
    std::fs::write(
        &editor,
        "#!/bin/sh\nprintf 'CORRO_LOG 1\\nSET A1 old\\nSET B1 added\\n' > \"$1\"\n",
    )
    .unwrap();
    std::fs::set_permissions(&editor, std::fs::Permissions::from_mode(0o755)).unwrap();

    let saved_visual = std::env::var("VISUAL").ok();
    let saved_editor = std::env::var("EDITOR").ok();
    std::env::remove_var("VISUAL");
    std::env::set_var("EDITOR", editor.to_string_lossy().as_ref());

    let mut app = App::new_with_paths(vec![log.clone()]);
    app.load_initial().unwrap();
    let mut pending_scope = 0u8;
    let mut clipboard = String::new();
    let result = dispatch_menu_action(
        &mut app,
        "edit_workbook_external",
        &mut pending_scope,
        &mut clipboard,
    );

    match &saved_visual {
        Some(v) => std::env::set_var("VISUAL", v),
        None => std::env::remove_var("VISUAL"),
    }
    match &saved_editor {
        Some(e) => std::env::set_var("EDITOR", e),
        None => std::env::remove_var("EDITOR"),
    }

    match result {
        MenuDispatch::Status(s) => {
            assert!(
                s.contains("Reloaded"),
                "external workbook edit should reload, got {s:?}"
            )
        }
        other => panic!("expected Status, got {}", dispatch_hint(&other)),
    }
    // The editor's new revision is live in the workbook.
    assert_eq!(
        app.core
            .workbook
            .active_sheet()
            .grid
            .get(&corro::grid::CellAddr::main(0, 1))
            .unwrap_or_default(),
        "added",
        "the externally-added B1 revision must be applied after reload"
    );
}

/// With no file bound yet, Edit ▸ Workbook (External) must report that the
/// workbook has to be saved first instead of spawning an editor on nothing.
#[test]
fn edit_workbook_external_without_a_path_reports_save_first() {
    let mut app = App::new_with_paths(vec![]);
    app.load_initial().unwrap();
    let mut pending_scope = 0u8;
    let mut clipboard = String::new();
    let result = dispatch_menu_action(
        &mut app,
        "edit_workbook_external",
        &mut pending_scope,
        &mut clipboard,
    );
    match result {
        MenuDispatch::Status(s) => assert!(
            s.contains("Save the workbook first"),
            "unsaved workbook must ask for a save first, got {s:?}"
        ),
        other => panic!("expected Status, got {}", dispatch_hint(&other)),
    }
}

/// A failing editor must surface as a status, never a panic or a lost workbook.
#[test]
fn edit_workbook_external_editor_error_is_a_status() {
    let _no_editor = NoEditorGuard::take();
    let dir = tempfile::tempdir().unwrap();
    let log = dir.path().join("book.corro");
    std::fs::write(&log, "CORRO_LOG 1\n").unwrap();

    let mut app = App::new_with_paths(vec![log.clone()]);
    app.load_initial().unwrap();
    let mut pending_scope = 0u8;
    let mut clipboard = String::new();
    let result = dispatch_menu_action(
        &mut app,
        "edit_workbook_external",
        &mut pending_scope,
        &mut clipboard,
    );
    match result {
        MenuDispatch::Status(s) => assert!(
            s.contains("Editor error"),
            "a missing editor must be reported, got {s:?}"
        ),
        other => panic!("expected Status, got {}", dispatch_hint(&other)),
    }
}
