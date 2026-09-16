//! Shared corro app actions used by every GUI backend (pancurses, GTK, ...).
//!
//! These functions operate purely on [`crate::gui::App`]/[`crate::core::state::CoreApp`]
//! and contain no backend-specific widget code, so the same behavior is
//! guaranteed across backends (render-parity). They were extracted from
//! `pnc_backend.rs`, which previously duplicated this logic with pancurses
//! widget calls interleaved. Backends call these and then only have to perform
//! their own widget updates (e.g. formula-bar text, dialogs, clipboard).

use crate::grid::{CellAddr, CellFormat, NumberFormat, SheetCursor, TextAlign, HEADER_ROWS, MARGIN_COLS};
use crate::ops::{Op, SheetState, WorkbookOp};
use crate::gui::App;

/// Re-exported for the GUI/pancurses dispatch path (canonical home is
/// [`crate::ui_core`], which ratatui-only builds can also see).
pub use crate::ui_core::prompt_action_write_target;

/// Set a main cell value in the grid and log it to the live `.corro` file (if any).
pub fn commit_cell(app: &mut App, addr: CellAddr, value: String) {
    let sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
    app.core.workbook.active_sheet_mut().grid.set(&addr, value.clone());
    let op = Op::SetCell { addr, value };
    let wbo = WorkbookOp::SheetOp { sheet_id, op };
    if let Some(ref p) = app.core.path.clone() {
        let mut active_sheet = sheet_id;
        let _ = crate::io::commit_workbook_op(
            p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo,
        );
        app.core.ops_applied = app.core.ops_applied.saturating_add(1);
    }
}

/// Apply a single sheet op to the active sheet, committing to the live file
/// when one is open (matching ratatui's commit_workbook_op path).
pub fn apply_sheet_op(app: &mut App, op: Op) {
    let sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
    let wbo = WorkbookOp::SheetOp { sheet_id, op };
    if let Some(ref p) = app.core.path.clone() {
        let mut active_sheet = sheet_id;
        let _ = crate::io::commit_workbook_op(
            p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo,
        );
        app.core.ops_applied = app.core.ops_applied.saturating_add(1);
    } else {
        let mut active = sheet_id;
        let _ = crate::ops::apply_workbook_op(&mut app.core.workbook, &mut active, wbo);
    }
}

/// Simple A1-style label for a main cell (column letters + 1-indexed row).
pub fn main_addr_label(row: u32, col: u32) -> String {
    let mut name = String::new();
    let mut c = col;
    loop {
        name.insert(0, (b'A' + (c % 26) as u8) as char);
        if c < 26 { break; }
        c = c / 26 - 1;
    }
    format!("{}{}", name, row + 1)
}

/// Format-scope menu actions as data: (action name, scope id, status text).
/// One table instead of six near-identical match arms.
const FORMAT_SCOPES: &[(&str, u8, &str)] = &[
    ("format_apply_all", 1, "Format scope: All"),
    ("format_apply_full_column", 2, "Format scope: Full column"),
    ("format_apply_data", 3, "Format scope: Data"),
    ("format_apply_special", 4, "Format scope: Special"),
    ("format_apply_cell", 0, "Format scope: Cell"),
    ("format_apply_selection", 5, "Format scope: Selection"),
];

fn format_scope(name: &str) -> Option<(u8, &'static str)> {
    FORMAT_SCOPES.iter().find(|(n, _, _)| *n == name).map(|(_, s, t)| (*s, *t))
}

/// Format-value menu actions as data: (action name, format, status text).
/// One table instead of eleven near-identical match arms.
const FORMAT_VALUES: &[(&str, CellFormat, &str)] = &[
    ("format_decimal_generic", CellFormat { number: Some(NumberFormat::DecimalGeneric), align: None }, "Format: Decimal (generic)"),
    ("format_currency", CellFormat { number: Some(NumberFormat::Currency { decimals: 2 }), align: None }, "Format: Currency ($)"),
    ("format_rational", CellFormat { number: Some(NumberFormat::Rational), align: None }, "Format: Rational"),
    ("format_fixed_0", CellFormat { number: Some(NumberFormat::Fixed { decimals: 0 }), align: None }, "Format: Fixed 0"),
    ("format_fixed_1", CellFormat { number: Some(NumberFormat::Fixed { decimals: 1 }), align: None }, "Format: Fixed 1"),
    ("format_fixed_2", CellFormat { number: Some(NumberFormat::Fixed { decimals: 2 }), align: None }, "Format: Fixed 2"),
    ("format_fixed_custom", CellFormat { number: Some(NumberFormat::Fixed { decimals: 2 }), align: None }, "Format: Fixed n"),
    ("format_align_left", CellFormat { number: None, align: Some(TextAlign::Left) }, "Format: Align Left"),
    ("format_align_center", CellFormat { number: None, align: Some(TextAlign::Center) }, "Format: Align Center"),
    ("format_align_right", CellFormat { number: None, align: Some(TextAlign::Right) }, "Format: Align Right"),
    ("format_align_default", CellFormat { number: None, align: Some(TextAlign::Default) }, "Format: Align Default"),
    ("format_reset", CellFormat { number: None, align: None }, "Format reset"),
];

fn format_value(name: &str) -> Option<(CellFormat, &'static str)> {
    FORMAT_VALUES.iter().find(|(n, _, _)| *n == name).map(|(_, f, t)| (*f, *t))
}

/// Apply `fmt` to a format target scope (0=Cell,1=All,2=Full column,3=Data,4=Special,5=Selection).
pub fn apply_format(app: &mut App, scope: u8, main_row: u32, main_col: u32, fmt: CellFormat) {
    let (mr, mc) = {
        let sheet = app.core.workbook.active_sheet();
        (sheet.grid.main_rows(), sheet.grid.main_cols())
    };
    let (r0, r1, c0, c1) = match scope {
        1 | 3 | 4 => (0, mr, 0, mc),
        2 => (0, mr, main_col as usize, main_col as usize + 1),
        5 => {
            let (ar, ac) = app.core.anchor
                .map(|a| (a.row.saturating_sub(HEADER_ROWS), a.col.saturating_sub(MARGIN_COLS)))
                .unwrap_or((main_row as usize, main_col as usize));
            (
                ar.min(main_row as usize),
                ar.max(main_row as usize) + 1,
                ac.min(main_col as usize),
                ac.max(main_col as usize) + 1,
            )
        }
        _ => (main_row as usize, main_row as usize + 1, main_col as usize, main_col as usize + 1),
    };
    for r in r0.min(mr)..r1.min(mr) {
        for c in c0.min(mc)..c1.min(mc) {
            app.core.workbook.active_sheet_mut().grid.set_cell_format(
                CellAddr::Main { row: r as u32, col: c as u32 },
                fmt,
            );
        }
    }
}

/// Reorder the used main rows of the active sheet by `col` (ascending or descending).
/// Returns the number of rows sorted.
pub fn sort_sheet(app: &mut App, col: usize, asc: bool) -> usize {
    let g = app.core.workbook.active_sheet().grid.clone();
    let mr = g.main_rows();
    let mc = g.main_cols();
    if col >= mc || mr < 2 { return 0; }
    let used: Vec<usize> = (0..mr)
        .filter(|&r| (0..mc).any(|c| !g.get(&CellAddr::Main { row: r as u32, col: c as u32 }).unwrap_or_default().trim().is_empty()))
        .collect();
    if used.len() < 2 { return 0; }
    let mut pairs: Vec<(usize, String)> = used.iter()
        .map(|&r| (r, g.get(&CellAddr::Main { row: r as u32, col: col as u32 }).unwrap_or_default()))
        .collect();
    pairs.sort_by(|a, b| {
        let na: Option<f64> = a.1.trim().parse().ok();
        let nb: Option<f64> = b.1.trim().parse().ok();
        let ord = match (na, nb) {
            (Some(x), Some(y)) => x.partial_cmp(&y).unwrap_or(std::cmp::Ordering::Equal),
            _ => a.1.cmp(&b.1),
        };
        if asc { ord } else { ord.reverse() }
    });
    let mut cells: Vec<((usize, usize), String)> = Vec::new();
    for &r in &used {
        for c in 0..mc {
            let v = g.get(&CellAddr::Main { row: r as u32, col: c as u32 }).unwrap_or_default();
            if !v.trim().is_empty() {
                cells.push(((r, c), v));
            }
        }
    }
    {
        let g2 = app.core.workbook.active_sheet_mut();
        for &r in &used {
            for c in 0..mc {
                g2.grid.set(&CellAddr::Main { row: r as u32, col: c as u32 }, String::new());
            }
        }
    }
    for (i, &(old, _)) in pairs.iter().enumerate() {
        let new_row = used[i];
        for ((r, c), v) in cells.iter() {
            if *r == old {
                app.core.workbook.active_sheet_mut().grid.set(
                    &CellAddr::Main { row: new_row as u32, col: *c as u32 }, v.clone(),
                );
            }
        }
    }
    used.len()
}

/// Returns the prompt label for menu actions that require free-text input,
/// or `None` if the action runs immediately. Shared across backends so every
/// UI collects the same set of prompts (each backend supplies its own prompt
/// widget via the toolkit).
pub fn menu_action_needs_prompt(name: &str) -> Option<&'static str> {
    Some(match name {
        "open" => "Open file",
        "save_as" => "Save as",
        "export_tsv" => "Export TSV",
        "export_csv" => "Export CSV",
        "export_ods" => "Export ODS",
        "export_ascii" => "Export ASCII",
        "export_all" => "Export all to",
        "set_col_width" => "Column width",
        "set_max_col_width" => "Default width",
        "go_to_cell" => "Go to cell",
        "find" => "Find",
        "replace" => "Replace (find|replacement)",
        "rename_sheet" => "Rename sheet to",
        "copy_sheet" => "Copy sheet as",
        "delete_sheet" => "Delete sheet named",
        // NOTE: insert_special_chars is NOT free-text: it opens the
        // 10-choice picker (same items/order as the ratatui reference),
        // dispatched via MenuDispatch::SpecialPicker. The formula bar
        // remains the arbitrary-input path, so the picker box needs no
        // free-text entry.
        "insert_hyperlink" => "Insert hyperlink",
        "sort_view" => "sort cols [A,B,C]",
        "persist_sort" => "sort cols [A,B,C] (save)",
        "balance_books" => "Balance column",
        _ => return None,
    })
}

/// Actions that spawn a child process needing the real terminal (external
/// editor): terminal backends must suspend around dispatch and resume
/// after (see the pancurses suspend_terminal/resume_terminal pair). Native
/// GUI backends ignore this (no terminal state to suspend). Mirrors the
/// [`menu_action_needs_prompt`] pattern.
pub fn menu_action_needs_terminal_suspend(name: &str) -> bool {
    matches!(name, "edit_external" | "edit_workbook_external")
}

/// Result of [`dispatch_menu_action`]. The backend performs the backend-
/// specific part (text prompt, OSC 52 clipboard, dialog rendering, formula-bar
/// update); the corro operation itself is already applied to `app`.
pub enum MenuDispatch {
    /// Operation applied; show this status (backend updates its formula bar).
    Status(String),
    /// Backend should open a text prompt with (label, action name).
    Prompt(&'static str, &'static str),
    /// Backend should open the Insert > Special Char 10-choice picker over
    /// shared [`super::special_picker`] state (same items, order, arrows,
    /// digits, Enter, Esc as the ratatui reference).
    SpecialPicker,
    /// Backend should enter edit mode on the cursor cell with `value` as the
    /// in-progress buffer (matching ratatui's start_edit_mode actions such as
    /// Insert Date / Insert Time).
    Edit { value: String },
    /// Render the About dialog (backend-specific); show `status`.
    About { status: String },
    /// Render the Full-help dialog (backend-specific); show `status`.
    HelpFull { status: String },
    /// Render the Keybindings dialog (backend-specific); show `status`.
    HelpKeybinds { status: String },
}

/// Perform an immediate (non-prompt) menu action on `app`, returning how the
/// backend should present the result. Actions that need a text prompt are
/// handled by the backend via [`menu_action_needs_prompt`] (except `save`, which
/// degrades to a prompt when no file is loaded). Pure corro logic — no widget
/// code — so every backend gets identical behavior (render parity).
///
/// `pending_scope` is the format scope chosen by a prior Format ▸ Scope pick;
/// `clipboard` is the in-app (session) clipboard used by copy/cut/paste.
pub fn dispatch_menu_action(
    app: &mut App,
    name: &str,
    pending_scope: &mut u8,
    clipboard: &mut String,
) -> MenuDispatch {
    let hr = HEADER_ROWS;
    let lm = MARGIN_COLS;
    let main_row = app.core.cursor.row.saturating_sub(hr) as u32;
    let main_col = app.core.cursor.col.saturating_sub(lm) as u32;
    let addr = CellAddr::Main { row: main_row, col: main_col };
    match name {
        "save" => {
            if let Some(ref p) = app.core.path.clone() {
                // Canonical CORRO_LOG, same writer the TUI uses.
                match crate::io::write_workbook_log(
                    p,
                    &app.core.workbook,
                    &app.core.persisted_view_sort_cols,
                ) {
                    Ok(()) => MenuDispatch::Status("Saved".into()),
                    Err(e) => MenuDispatch::Status(format!("Save error: {e}")),
                }
            } else {
                MenuDispatch::Prompt("Save as", "save_as")
            }
        }
        "insert_rows" => {
            // Insert a blank row ABOVE the cursor (matching ratatui's
            // insert_rows_above_cursor): grow the grid, then move the rows
            // below the cursor down to make room. Status text matches.
            let original_main_rows = app.core.workbook.active_sheet().grid.main_rows() as u32;
            let row = main_row;
            if (row as usize) < original_main_rows as usize {
                let main_cols = app.core.workbook.active_sheet().grid.main_cols() as u32;
                apply_sheet_op(app, Op::SetMainSize { main_rows: original_main_rows + 1, main_cols });
                apply_sheet_op(app, Op::MoveRowRange { from: row, count: original_main_rows - row, to: original_main_rows + 1 });
                app.core.cursor = SheetCursor { row: hr + row as usize, col: app.core.cursor.col };
                MenuDispatch::Status(format!("Inserted 1 row above row {row}"))
            } else {
                // Cursor outside the main band: fall back to a plain grow.
                app.core.workbook.active_sheet_mut().grid.grow_main_row_at_bottom();
                MenuDispatch::Status("Inserted row".into())
            }
        }
        "insert_special_chars" => {
            // Picker, not free-text (ratatui parity): the backend opens
            // its 10-choice dialog/popup over shared picker state.
            super::special_picker::open(app);
            MenuDispatch::SpecialPicker
        }
        "insert_mitosis_row" => {
            // Mitosis COPIES the cursor's main row into a new row below it
            // (shifting lower rows down) and moves the cursor onto the
            // duplicate — matching the ratatui reference's
            // insert_mitosis_row_after_cursor (Op::DuplicateRow).  It used to
            // be aliased to a plain blank-row grow, copying nothing.
            let sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
            let main_rows = app.core.workbook.active_sheet().grid.main_rows();
            if (main_row as usize) < main_rows {
                let wbo = WorkbookOp::SheetOp { sheet_id, op: Op::DuplicateRow { row: main_row } };
                if let Some(ref p) = app.core.path.clone() {
                    let mut active_sheet = sheet_id;
                    let _ = crate::io::commit_workbook_op(
                        p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo,
                    );
                } else {
                    let mut active = sheet_id;
                    let _ = crate::ops::apply_workbook_op(&mut app.core.workbook, &mut active, wbo.clone());
                }
                app.core.cursor = SheetCursor {
                    row: hr + main_row as usize + 1,
                    col: app.core.cursor.col,
                };
                MenuDispatch::Status(format!("Inserted mitosis row after row {}", main_row + 1))
            } else {
                // Cursor outside the main band: fall back to a plain insert
                // (band-specific header/footer mitosis is not exposed here).
                app.core.workbook.active_sheet_mut().grid.grow_main_row_at_bottom();
                MenuDispatch::Status("Inserted row".into())
            }
        }
        "insert_cols" => {
            // Insert a blank column LEFT of the cursor (matching ratatui's
            // insert_cols_left_of_cursor): grow the grid, then move the
            // columns right of the cursor right to make room.
            let original_main_cols = app.core.workbook.active_sheet().grid.main_cols() as u32;
            let col = main_col;
            if (col as usize) < original_main_cols as usize {
                let main_rows = app.core.workbook.active_sheet().grid.main_rows() as u32;
                apply_sheet_op(app, Op::SetMainSize { main_rows, main_cols: original_main_cols + 1 });
                apply_sheet_op(app, Op::MoveColRange { from: col, count: original_main_cols - col, to: original_main_cols + 1 });
                app.core.cursor = SheetCursor { row: app.core.cursor.row, col: lm + col as usize };
                MenuDispatch::Status(format!("Inserted 1 column left of column {col}"))
            } else {
                // Cursor outside the main band: fall back to a plain grow.
                app.core.workbook.active_sheet_mut().grid.grow_main_col_at_right();
                MenuDispatch::Status("Inserted column".into())
            }
        }
        "insert_mitosis_col" => {
            // Mitosis COPIES the cursor's main column into a new column to its
            // right (shifting columns right) and moves the cursor onto the
            // duplicate — matching the ratatui reference's
            // insert_mitosis_col_after_cursor (Op::DuplicateCol).
            let sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
            let main_cols = app.core.workbook.active_sheet().grid.main_cols();
            if (main_col as usize) < main_cols {
                let wbo = WorkbookOp::SheetOp { sheet_id, op: Op::DuplicateCol { col: main_col } };
                if let Some(ref p) = app.core.path.clone() {
                    let mut active_sheet = sheet_id;
                    let _ = crate::io::commit_workbook_op(
                        p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo,
                    );
                } else {
                    let mut active = sheet_id;
                    let _ = crate::ops::apply_workbook_op(&mut app.core.workbook, &mut active, wbo.clone());
                }
                app.core.cursor = SheetCursor {
                    row: app.core.cursor.row,
                    col: lm + main_col as usize + 1,
                };
                MenuDispatch::Status(format!("Inserted mitosis col after col {}", main_col + 1))
            } else {
                app.core.workbook.active_sheet_mut().grid.grow_main_col_at_right();
                MenuDispatch::Status("Inserted column".into())
            }
        }
        "insert_date" => {
            // Enter edit mode with the date as the in-progress buffer,
            // matching ratatui's InsertDate (start_edit_mode). The user
            // presses Enter to commit.
            let d = crate::ui_core::today_string();
            MenuDispatch::Edit { value: d }
        }
        "insert_time" => {
            // Enter edit mode with the time as the in-progress buffer,
            // matching ratatui's InsertTime.
            let t = crate::ui_core::clock_string();
            MenuDispatch::Edit { value: t }
        }
        "delete_cell" | "delete" => {
            commit_cell(app, addr.clone(), String::new());
            MenuDispatch::Status(format!("Cleared {}", main_addr_label(main_row, main_col)))
        }
        "follow_hyperlink" => {
            // Shared follow logic (same status texts as ratatui's Ctrl+O /
            // Edit ▸ Follow link): opens the cursor cell's hyperlink in the
            // default browser. Pure corro logic — every GUI backend routes
            // here, so the behavior can never drift per backend.
            let grid = &app.core.workbook.active_sheet().grid;
            let status = crate::ui_core::follow_hyperlink(grid, app.core.cursor);
            MenuDispatch::Status(status)
        }
        "edit_external" => {
            // Blocking and modal-like: the backend waits while the editor
            // runs. Terminal backends suspend around dispatch (see
            // menu_action_needs_terminal_suspend); the native GUI inherits
            // the console (or lack thereof) as-is.
            let initial = app
                .core
                .workbook
                .active_sheet()
                .grid
                .get(&addr)
                .unwrap_or_default();
            match crate::editor::edit_text_externally(&initial) {
                Ok(Some(text)) => {
                    commit_cell(app, addr, text);
                    MenuDispatch::Status("External edit applied".into())
                }
                Ok(None) => MenuDispatch::Status("Unchanged".into()),
                Err(e) => MenuDispatch::Status(format!("Editor error: {e}")),
            }
        }
        "edit_workbook_external" => {
            // Edit ▸ Workbook (External): open the append-only log itself in
            // $EDITOR. Blocking/modal-like, like edit_external. The file is
            // the source of truth, so afterwards the workbook is reloaded
            // from it (anchored on the unchanged prefix: appends tail-apply,
            // a rewrite falls back to a full reload — same paths as another
            // window's Save).
            let Some(path) = app.core.path.clone() else {
                return MenuDispatch::Status(
                    "Save the workbook first (File ▸ Save as), then edit it externally".into(),
                );
            };
            match crate::editor::edit_workbook_externally(&path) {
                Ok(true) => match app.core.poll_log_tail() {
                    Ok(_) => MenuDispatch::Status(format!(
                        "Reloaded {} after external edit",
                        path.display()
                    )),
                    Err(e) => MenuDispatch::Status(format!("Reload error: {e}")),
                },
                Ok(false) => MenuDispatch::Status("Unchanged".into()),
                Err(e) => MenuDispatch::Status(format!("Editor error: {e}")),
            }
        }
        "select_all" => {
            let (mr, mc) = {
                let sheet = app.core.workbook.active_sheet();
                (sheet.grid.main_rows(), sheet.grid.main_cols())
            };
            if mr > 0 && mc > 0 {
                app.core.anchor = Some(SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS });
                app.core.cursor = SheetCursor {
                    row: HEADER_ROWS + mr.saturating_sub(1),
                    col: MARGIN_COLS + mc.saturating_sub(1),
                };
            }
            MenuDispatch::Status("Selected all".into())
        }
        "new_file" => {
            // Shared blank-document logic (same content and status texts as
            // ratatui's File ▸ New): template-or-seeded workbook, detached
            // from any file, histories/caches/watchers reset. Pure corro
            // logic — every GUI backend routes here, so the behavior can
            // never drift per backend. Backends refresh their viewport after
            // dispatch (workbook dims may change).
            let (workbook, note) = crate::ui_core::fresh_blank_workbook();
            app.core.workbook = workbook;
            app.core.workbook.ensure_active_sheet();
            app.core.view_sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
            if let Some(idx) = app.core.workbook.sheet_index_by_id(app.core.view_sheet_id) {
                app.core.workbook.active_sheet = idx;
                app.core.state = app.core.workbook.sheets[idx].state.clone();
            }
            app.core.path = None;
            app.core.source_path = None;
            app.core.import_source = None;
            app.core.offset = 0;
            app.core.ops_applied = 0;
            app.core.persisted_view_sort_cols.clear();
            app.core.revision_limit = None;
            app.core.revision_browse = false;
            app.core.revision_browse_limit = 0;
            app.core.watcher = None;
            app.core.op_history.clear();
            app.core.redo_history.clear();
            app.core.cursor = SheetCursor { row: hr, col: lm };
            app.core.anchor = None;
            app.core.unsaved_file = None;
            app.core.linked_source_mtimes.clear();
            app.core.edit_target_addr = None;
            app.core.edit_range_addrs = None;
            app.core.pending_lost_edit = None;
            app.core.pending_fit_to_content_on_commit = false;
            app.extrapolate = None;
            app.special_picker = None;
            MenuDispatch::Status(note.unwrap_or_else(|| "New workbook".into()))
        }
        "new_sheet" => {
            let id = app.core.workbook.next_sheet_id;
            let title = format!("Sheet{id}");
            let idx = app.core.workbook.add_sheet(title.clone(), SheetState::new_seeded());
            app.core.workbook.active_sheet = idx;
            app.core.view_sheet_id = id;
            app.core.cursor = SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS };
            let wbo = WorkbookOp::NewSheet { id, title: title.clone() };
            if let Some(ref p) = app.core.path.clone() {
                let mut active_sheet = id;
                let _ = crate::io::commit_workbook_op(p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo);
                app.core.ops_applied = app.core.ops_applied.saturating_add(1);
            }
            MenuDispatch::Status("New sheet created".into())
        }
        "copy" => {
            let val = app.core.workbook.active_sheet().grid.get(&addr).unwrap_or_default();
            *clipboard = val.clone();
            // ratatui's Copy sets no status (keeps the current one).
            MenuDispatch::Status(String::new())
        },
        "cut" => {
            let val = app.core.workbook.active_sheet().grid.get(&addr).unwrap_or_default();
            if val.is_empty() {
                MenuDispatch::Status("Nothing to cut".into())
            } else {
                *clipboard = val.clone();
                commit_cell(app, addr.clone(), String::new());
                MenuDispatch::Status("Selection cut".into())
            }
        },
        "paste" => {
            let val = clipboard.clone();
            if val.is_empty() {
                MenuDispatch::Status("Clipboard empty (use Copy/Cut first)".into())
            } else {
                commit_cell(app, addr.clone(), val);
                // ratatui's Paste sets no status (keeps the current one).
                MenuDispatch::Status(String::new())
            }
        }
        "sort_asc" | "sort_desc" => {
            let asc = name != "sort_desc";
            let n = sort_sheet(app, main_col as usize, asc);
            MenuDispatch::Status(if n > 0 {
                format!("Sorted {n} rows by column {}", main_addr_label(0, main_col))
            } else {
                "Nothing to sort".into()
            })
        }
        "replay" => {
            // Replay the current file (reload all revisions), matching
            // ratatui's Replay action. Status text matches.
            if let Some(ref p) = app.core.path.clone() {
                if p.exists() {
                    let mut workbook = crate::ops::WorkbookState::new();
                    let mut active_sheet = workbook.sheet_id(workbook.active_sheet);
                    match crate::io::load_workbook_revisions_partial(p, usize::MAX, &mut workbook, &mut active_sheet) {
                        Ok((off, replay)) => {
                            app.core.workbook = workbook;
                            app.core.offset = off;
                            app.core.ops_applied = replay.op_count;
                            app.core.cursor = SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS };
                            MenuDispatch::Status(format!("Replayed {} @ revision {}", p.display(), replay.op_count))
                        }
                        Err(e) => MenuDispatch::Status(format!("Replay error: {e}")),
                    }
                } else {
                    MenuDispatch::Status("Replay: file not found".into())
                }
            } else {
                MenuDispatch::Status("Replay: no file loaded".into())
            }
        }
        // NOTE: no "extrapolate" arm here on purpose. Both GUI backends
        // enter the interactive extrapolate modal before dispatch runs
        // (pancurses intercepts it in the menu callback, GTK has its own
        // modal arm), so a one-shot arm would be dead code — and a wrong one
        // (it would overwrite the cell below the cursor). The modal commits
        // through extrapolate_cells, shared with the ratatui reference.
        "duplicate" => {
            // Duplicate the cursor row (matching ratatui's Duplicate mode
            // Enter on a single cell, which inserts a mitosis row).
            let sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
            let main_rows = app.core.workbook.active_sheet().grid.main_rows();
            if (main_row as usize) < main_rows {
                let wbo = WorkbookOp::SheetOp { sheet_id, op: Op::DuplicateRow { row: main_row } };
                if let Some(ref p) = app.core.path.clone() {
                    let mut active_sheet = sheet_id;
                    let _ = crate::io::commit_workbook_op(p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo);
                    app.core.ops_applied = app.core.ops_applied.saturating_add(1);
                } else {
                    let mut active = sheet_id;
                    let _ = crate::ops::apply_workbook_op(&mut app.core.workbook, &mut active, wbo);
                }
                app.core.cursor = SheetCursor { row: hr + main_row as usize + 1, col: app.core.cursor.col };
                MenuDispatch::Status("Duplicated row".into())
            } else {
                MenuDispatch::Status("Nothing to duplicate".into())
            }
        }
        "sheet_prev" | "sheet_next" => {
            let n = app.core.workbook.sheet_count();
            if n > 1 {
                let delta = if name == "sheet_prev" { -1i32 } else { 1i32 };
                let cur = app.core.workbook.active_sheet as i32;
                let next = (cur + delta).rem_euclid(n as i32) as usize;
                app.core.workbook.active_sheet = next;
                app.core.view_sheet_id = app.core.workbook.sheet_id(next);
                app.core.cursor = SheetCursor { row: HEADER_ROWS, col: MARGIN_COLS };
                MenuDispatch::Status(format!("Sheet {} of {}", next + 1, n))
            } else {
                // Single sheet: no-op, keep the current status (matching
                // ratatui's switch_sheet early return).
                MenuDispatch::Status(String::new())
            }
        }
        "move_sheet" => {
            let n = app.core.workbook.sheet_count();
            if n > 1 {
                let cur = app.core.workbook.active_sheet;
                let rec = app.core.workbook.sheets.remove(cur);
                app.core.workbook.sheets.push(rec);
                app.core.workbook.active_sheet = n - 1;
                app.core.view_sheet_id = app.core.workbook.sheet_id(n - 1);
                MenuDispatch::Status(format!("Moved sheet to end ({})", app.core.workbook.sheet_title(n - 1)))
            } else {
                MenuDispatch::Status("Only one sheet".into())
            }
        }
        "help_rows" => MenuDispatch::Status("Row ops: v·select full rows, then r·move to target row".into()),
        "help_cols" => MenuDispatch::Status("Col ops: v·select full columns, then c·move to target column".into()),
        "help_full" => MenuDispatch::HelpFull { status: "Help: Full help".into() },
        _ if let Some((scope, status)) = format_scope(name) => {
            *pending_scope = scope;
            MenuDispatch::Status(status.into())
        }
        _ if let Some((fmt, status)) = format_value(name) => {
            apply_format(app, *pending_scope, main_row, main_col, fmt);
            MenuDispatch::Status(status.into())
        }
        "about" => MenuDispatch::About { status: "About".into() },
        "help_keybinds" => MenuDispatch::HelpKeybinds { status: "Help".into() },
        "toggle_headers" | "toggle_margins" => MenuDispatch::Status(format!("Toggle: {name} (fixed chrome in this build)")),
        "undo" | "redo" => MenuDispatch::Status(format!("{name}: no undo/redo history in the pancurses build yet")),
        _ => MenuDispatch::Status(format!("Menu action: {name}")),
    }
}

/// Perform the file/name operation submitted via a backend text prompt.
/// Mutates `app.core.status`; the backend then updates its formula bar.
/// Pure corro logic — no widget code.
pub fn run_prompt_action(app: &mut App, action: &str, text: &str) {
    let path = text.trim().to_string();
    match action {
        "open" => {
            if !path.is_empty() {
                match crate::io::load_workbook_file(std::path::Path::new(&path)) {
                    Ok(workbook) => {
                        app.core.workbook = workbook;
                        app.core.offset = 0;
                        app.core.ops_applied = 0;
                        app.core.path = Some(std::path::PathBuf::from(path.clone()));
                        app.core.status = format!("Opened {path}");
                    }
                    Err(e) => app.core.status = format!("Open error: {e}"),
                }
            }
        }
        "save_as" => {
            if !path.is_empty() {
                // Workbooks must keep `.corro` (a foreign extension will not
                // reopen as one); the GUI dialog already forces this, the
                // typed pancurses path resolves here (ratatui resolves in
                // its own Save arm via `to_corro_path`).
                let final_path =
                    crate::ui_core::force_extension(std::path::Path::new(&path), "corro");
                match crate::io::write_workbook_log(
                    &final_path,
                    &app.core.workbook,
                    &app.core.persisted_view_sort_cols,
                ) {
                    Ok(()) => {
                        app.core.path = Some(final_path.clone());
                        app.core.status = format!("Saved to {}", final_path.display());
                    }
                    Err(e) => app.core.status = format!("Save error: {e}"),
                }
            }
        }
        "export_tsv" | "export_csv" | "export_ods" | "export_ascii" | "export_all" => {
            if !path.is_empty() {
                // A bare typed name lands on the format extension (the GUI
                // dialog suggests/appends it already); an explicit
                // extension is always respected.
                let ext = crate::ui_core::export_ext_for_action(action);
                let final_path = crate::ui_core::append_extension_if_missing(
                    std::path::Path::new(&path),
                    ext,
                );
                let g = app.core.workbook.active_sheet().grid.clone();
                match std::fs::File::create(&final_path) {
                    Ok(mut f) => {
                        let r: Result<(), String> = match action {
                            "export_tsv" => { crate::export::export_tsv(&g, &mut f); Ok(()) }
                            "export_csv" => { crate::export::export_csv(&g, &mut f); Ok(()) }
                            "export_ascii" => { crate::export::export_ascii_table(&g, &mut f, true); Ok(()) }
                            "export_ods" => match crate::ods::export_ods_bytes(&g) {
                                Ok(bytes) => std::io::Write::write_all(&mut f, &bytes).map_err(|e| e.to_string()),
                                Err(e) => Err(e.to_string()),
                            },
                            "export_all" => { crate::export::export_tsv(&g, &mut f); Ok(()) }
                            _ => Ok(()),
                        };
                        match r {
                            Ok(()) => {
                                app.core.status =
                                    format!("Exported {action} to {}", final_path.display())
                            }
                            Err(e) => app.core.status = format!("Export error: {e}"),
                        }
                    }
                    Err(e) => app.core.status = format!("Export error: {e}"),
                }
            }
        }
        "set_col_width" | "set_max_col_width" => {
            if let Ok(w) = path.trim().parse::<u32>() {
                app.core.status = format!("Column width set to {w} (applied on next render)");
            } else {
                app.core.status = "Column width: enter a number".into();
            }
        }
        "sort_view" => {
            // Parse "A,B,C" (optional "!" prefix = descending) into view-sort
            // cols, matching the ratatui SortView prompt exactly.
            let cols: Vec<crate::grid::SortSpec> = path
                .split(',')
                .filter_map(|s| {
                    let s = s.trim();
                    if s.is_empty() {
                        None
                    } else {
                        let (desc, raw) = if let Some(rest) = s.strip_prefix('!') {
                            (true, rest)
                        } else {
                            (false, s)
                        };
                        crate::addr::parse_excel_column(raw).map(|c| crate::grid::SortSpec {
                            col: MARGIN_COLS + c as usize,
                            desc,
                        })
                    }
                })
                .collect();
            app.core.workbook.active_sheet_mut().grid.set_view_sort_cols(cols);
            app.core.status = "View sort updated".into();
        }
        "persist_sort" => {
            // Same parse as sort_view, but persist the sort (commit a
            // SetViewSortCols op to the live file when one is open) and
            // record it in the persisted-sort cache — matching ratatui's
            // SortView with persist=true ("View sort saved").
            let cols: Vec<crate::grid::SortSpec> = path
                .split(',')
                .filter_map(|s| {
                    let s = s.trim();
                    if s.is_empty() {
                        None
                    } else {
                        let (desc, raw) = if let Some(rest) = s.strip_prefix('!') {
                            (true, rest)
                        } else {
                            (false, s)
                        };
                        crate::addr::parse_excel_column(raw).map(|c| crate::grid::SortSpec {
                            col: MARGIN_COLS + c as usize,
                            desc,
                        })
                    }
                })
                .collect();
            let sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
            app.core.workbook.active_sheet_mut().grid.set_view_sort_cols(cols.clone());
            if !cols.is_empty() {
                app.core.persisted_view_sort_cols.insert(sheet_id, cols.clone());
            } else {
                app.core.persisted_view_sort_cols.remove(&sheet_id);
            }
            if let Some(ref p) = app.core.path.clone() {
                let mut active_sheet = sheet_id;
                let wbo = WorkbookOp::SheetOp {
                    sheet_id,
                    op: Op::SetViewSortCols { cols },
                };
                let _ = crate::io::commit_workbook_op(p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo);
                app.core.ops_applied = app.core.ops_applied.saturating_add(1);
            }
            app.core.status = "View sort saved".into();
        }
        "balance_books" => {
            // Generate a balance report sheet from the chosen amount column,
            // matching ratatui's BalanceBooks (persist=true).
            let col = if path.trim().is_empty() {
                crate::balance::choose_balance_column(&app.core.workbook.active_sheet().grid)
            } else {
                crate::addr::parse_excel_column(path.trim()).map(|c| c as usize)
            };
            let Some(col) = col else {
                app.core.status = "No balance column found".into();
                return;
            };
            let direction = crate::balance::BalanceDirection::PosToNeg;
            let report = crate::balance::build_balance_report(
                &app.core.workbook.active_sheet().grid, col, direction,
            );
            let source_sheet_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
            let source_title = app.core.workbook.sheet_title(app.core.workbook.active_sheet).to_string();
            let title = format!("Balance-{}", app.core.workbook.next_sheet_id);
            let id = app.core.workbook.next_sheet_id;
            let plan = crate::balance::balance_copy_plan(
                source_sheet_id, source_title.clone(), id, title.clone(), col,
                app.core.workbook.active_sheet().grid.main_rows(), &report, true,
            );
            let report_sheet = crate::balance::materialize_report_sheet(
                &app.core.workbook.active_sheet().clone(), &plan,
            );
            app.core.workbook.add_sheet(title.clone(), report_sheet);
            app.core.workbook.active_sheet = app.core.workbook.sheet_index_by_id(id).unwrap_or(app.core.workbook.active_sheet);
            app.core.view_sheet_id = id;
            if let Some(ref p) = app.core.path.clone() {
                let mut active_sheet = id;
                let wbo = WorkbookOp::BalanceReport {
                    id, title: title.clone(), source_sheet_id, amount_col: col, direction,
                    row_order: plan.row_order.clone(),
                    show_unmatched_heading: plan.show_unmatched_heading,
                    unmatched_start: plan.unmatched_start,
                    preserve_formulas: true,
                };
                let _ = crate::io::commit_workbook_op(p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo);
                app.core.ops_applied = app.core.ops_applied.saturating_add(1);
            }
            app.core.status = format!("Balance report saved as {title}");
        }
        "go_to_cell" => {
            if !path.is_empty() {
                if let Some((addr, _, _)) = crate::addr::parse_cell_ref_at(&path, 0) {
                    match addr {
                        crate::grid::CellAddr::Main { row, col } => {
                            app.core.cursor.row = HEADER_ROWS + row as usize;
                            app.core.cursor.col = MARGIN_COLS + col as usize;
                            app.core.status = format!("Go to {path}");
                        }
                        _ => app.core.status = format!("Unknown cell '{path}'"),
                    }
                } else {
                    app.core.status = format!("Unknown cell '{path}'");
                }
            }
        }
        "find" => {
            if !path.is_empty() {
                let g = app.core.workbook.active_sheet().grid.clone();
                let mut found = None;
                'find_loop: for r in 0..g.main_rows() {
                    for c in 0..g.main_cols() {
                        let v = g.get(&CellAddr::Main { row: r as u32, col: c as u32 }).unwrap_or_default();
                        if !v.trim().is_empty() && v.contains(&path) {
                            found = Some((r, c));
                            break 'find_loop;
                        }
                    }
                }
                if let Some((r, c)) = found {
                    app.core.cursor.row = HEADER_ROWS + r;
                    app.core.cursor.col = MARGIN_COLS + c;
                    app.core.status = format!("Found '{}' at {}", path, main_addr_label(r as u32, c as u32));
                } else {
                    app.core.status = format!("'{}' not found", path);
                }
            }
        }
        "replace" => {
            let (find, repl) = match path.split_once('|') {
                Some((f, r)) => (f.to_string(), r.to_string()),
                None => (path.clone(), String::new()),
            };
            if find.is_empty() {
                app.core.status = "Replace: enter find|replacement".into();
            } else {
                let g = app.core.workbook.active_sheet().grid.clone();
                let mr = g.main_rows();
                let mc = g.main_cols();
                let mut count = 0usize;
                for r in 0..mr {
                    for c in 0..mc {
                        let v = g.get(&CellAddr::Main { row: r as u32, col: c as u32 }).unwrap_or_default();
                        if v.contains(&find) {
                            let nv = v.replace(&find, &repl);
                            commit_cell(app, CellAddr::Main { row: r as u32, col: c as u32 }, nv);
                            count += 1;
                        }
                    }
                }
                app.core.status = format!("Replaced {count} occurrence(s)");
            }
        }
        "rename_sheet" => {
            if !path.is_empty() {
                let id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
                if let Some(sheet) = app.core.workbook.sheets.iter_mut().find(|s| s.id == id) {
                    sheet.title = path.clone();
                }
                let wbo = WorkbookOp::RenameSheet { id, title: path.clone() };
                if let Some(ref p) = app.core.path.clone() {
                    let mut active_sheet = id;
                    let _ = crate::io::commit_workbook_op(p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo);
                    app.core.ops_applied = app.core.ops_applied.saturating_add(1);
                }
                app.core.status = format!("Renamed sheet to {path}");
            }
        }
        "copy_sheet" => {
            if !path.is_empty() {
                let source_id = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
                let id = app.core.workbook.next_sheet_id;
                if let Some(source) = app.core.workbook.sheets.iter().find(|s| s.id == source_id) {
                    let nidx = app.core.workbook.add_sheet_record(crate::ops::SheetRecord {
                        id,
                        title: path.clone(),
                        state: source.state.clone(),
                        linked_source: source.linked_source.clone(),
                    });
                    app.core.workbook.active_sheet = nidx;
                }
                let wbo = WorkbookOp::CopySheet { source_id, id, title: path.clone() };
                if let Some(ref p) = app.core.path.clone() {
                    let mut active_sheet = source_id;
                    let _ = crate::io::commit_workbook_op(p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo);
                    app.core.ops_applied = app.core.ops_applied.saturating_add(1);
                }
                app.core.status = format!("Copied sheet as {path}");
            }
        }
        "delete_sheet" => {
            if app.core.workbook.sheets.len() > 1 {
                let idx = if path.is_empty() {
                    app.core.workbook.active_sheet
                } else {
                    app.core.workbook.sheets.iter().position(|s| s.title == path).unwrap_or(app.core.workbook.active_sheet)
                };
                let id = app.core.workbook.sheets[idx].id;
                let wbo = WorkbookOp::DeleteSheet { id };
                let mut active_sheet = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
                // Apply the deletion exactly once. `commit_workbook_op` ends
                // with `tail_apply_workbook`, which replays the appended line
                // into the workbook — so the sheet is already gone. The old
                // code *also* did `sheets.remove(idx)` here, which deleted a
                // second sheet and could empty the workbook entirely (then
                // any later `active_sheet()` panicked).
                match app.core.path.clone() {
                    Some(p) => {
                        let _ = crate::io::commit_workbook_op(
                            p.as_path(),
                            &mut app.core.offset,
                            &mut app.core.workbook,
                            &mut active_sheet,
                            &wbo,
                        );
                        app.core.ops_applied = app.core.ops_applied.saturating_add(1);
                    }
                    // No file: apply the op directly (same arm the commit
                    // path replays: removes, fixes active_sheet, and keeps at
                    // least one sheet via ensure_active_sheet).
                    None => {
                        let _ = crate::ops::apply_workbook_op(
                            &mut app.core.workbook,
                            &mut active_sheet,
                            wbo,
                        );
                    }
                }
                app.core.status = if path.is_empty() { "Deleted active sheet".into() } else { format!("Deleted sheet {path}") };
            } else {
                app.core.status = "Cannot delete the last sheet".into();
            }
        }
        "insert_hyperlink" => {
            if !path.is_empty() {
                let hr = HEADER_ROWS;
                let lm = MARGIN_COLS;
                let addr = CellAddr::Main {
                    row: app.core.cursor.row.saturating_sub(hr) as u32,
                    col: app.core.cursor.col.saturating_sub(lm) as u32,
                };
                commit_cell(app, addr.clone(), path.clone());
                app.core.status = format!("Inserted hyperlink {path}");
            }
        }
        _ => { app.core.status = format!("Menu: {action}"); }
    }
}
