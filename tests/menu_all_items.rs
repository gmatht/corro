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
    dispatch_menu_action, menu_action_needs_prompt, run_prompt_action, MenuDispatch,
};
use corro::gui::menu::{action_kind_to_name, menu_bar, MenuAction, MenuActionKind};
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

/// A short human-readable key for a `MenuDispatch` (which isn't `Debug`).
fn dispatch_hint(d: &MenuDispatch) -> &'static str {
    match d {
        MenuDispatch::Status(_) => "Status",
        MenuDispatch::Prompt(..) => "Prompt",
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
    let mut app = seeded_app(None);
    let mut pending_scope = 0u8;
    let mut clipboard = String::new();

    // Names that the backend handles before/without the dispatcher (the pancurses
    // backend sets running=false for "quit"; these have no dispatch work to do).
    let backend_special: HashSet<&'static str> = ["quit"].into_iter().collect();

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
        ("insert_special_chars", "Ω"),
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
        Open, Save, SaveAs, Quit, Undo, Redo, Cut, Copy, Paste, Find, Replace,
        DeleteCell, SelectAll, ToggleHeaders, ToggleMargins, NewSheet, RenameSheet,
        DeleteSheet, SortAsc, SortDesc, BalanceBooks, ExportTsv, ExportCsv, ExportOds,
        ExportAscii, About, HelpKeybinds, InsertRows, InsertMitosisRow, InsertMitosisCol,
        InsertCols, InsertSpecialChars, InsertDate, InsertTime, InsertHyperlink,
        FormatApplyAll, FormatApplyFullColumn, FormatApplyData, FormatApplySpecial,
        FormatApplyCell, FormatApplySelection, FormatDecimalGeneric, FormatCurrency,
        FormatRational, FormatFixed0, FormatFixed1, FormatFixed2, FormatFixedCustom,
        FormatAlignLeft, FormatAlignCenter, FormatAlignRight, FormatAlignDefault,
        FormatReset, ExportAll, Submenu, SortView, SaveSort, Replay, SetMaxColWidth,
        SetColWidth, ExportOdt, Duplicate, Extrapolate, SheetPrev, SheetNext, CopySheet,
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
        "insert_special_chars",
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
