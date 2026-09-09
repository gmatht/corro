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
use chrono;

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
        "insert_special_chars" => "Insert special char",
        "insert_hyperlink" => "Insert hyperlink",
        "sort_view" => "sort cols [A,B,C]",
        "persist_sort" => "sort cols [A,B,C] (save)",
        _ => return None,
    })
}

/// Result of [`dispatch_menu_action`]. The backend performs the backend-
/// specific part (text prompt, OSC 52 clipboard, dialog rendering, formula-bar
/// update); the corro operation itself is already applied to `app`.
pub enum MenuDispatch {
    /// Operation applied; show this status (backend updates its formula bar).
    Status(String),
    /// Backend should open a text prompt with (label, action name).
    Prompt(&'static str, &'static str),
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
                let snap = crate::ops::WorkbookSnapshot::from_workbook(&app.core.workbook);
                match crate::io::save_workbook(p, &snap) {
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
            let d = chrono::Local::now().format("%Y-%m-%d").to_string();
            MenuDispatch::Edit { value: d }
        }
        "insert_time" => {
            // Enter edit mode with the time as the in-progress buffer,
            // matching ratatui's InsertTime.
            let t = chrono::Local::now().format("%H:%M:%S").to_string();
            MenuDispatch::Edit { value: t }
        }
        "delete_cell" | "delete" => {
            commit_cell(app, addr.clone(), String::new());
            MenuDispatch::Status(format!("Cleared {}", main_addr_label(main_row, main_col)))
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
        "new_sheet" => {
            let id = app.core.workbook.next_sheet_id;
            let title = format!("Sheet{id}");
            let idx = app.core.workbook.add_sheet(title.clone(), SheetState::new(1, 1));
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
            MenuDispatch::Status(format!("Copied {}", main_addr_label(main_row, main_col)))
        },
        "cut" => {
            let val = app.core.workbook.active_sheet().grid.get(&addr).unwrap_or_default();
            *clipboard = val.clone();
            commit_cell(app, addr.clone(), String::new());
            MenuDispatch::Status(format!("Cut {}", main_addr_label(main_row, main_col)))
        },
        "paste" => {
            let val = clipboard.clone();
            if val.is_empty() {
                MenuDispatch::Status("Clipboard empty (use Copy/Cut first)".into())
            } else {
                commit_cell(app, addr.clone(), val);
                MenuDispatch::Status(format!("Pasted at {}", main_addr_label(main_row, main_col)))
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
        "replay" => MenuDispatch::Status("Replay: not available in the pancurses build".into()),
        "extrapolate" => MenuDispatch::Status("Extrapolate: not available in the pancurses build".into()),
        "duplicate" => {
            let val = app.core.workbook.active_sheet().grid.get(&addr).unwrap_or_default();
            if val.is_empty() {
                MenuDispatch::Status("Nothing to duplicate".into())
            } else {
                let below = CellAddr::Main { row: main_row + 1, col: main_col };
                commit_cell(app, below.clone(), val);
                MenuDispatch::Status(format!(
                    "Duplicated {} to {}",
                    main_addr_label(main_row, main_col),
                    main_addr_label(main_row + 1, main_col),
                ))
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
        "format_apply_all" => { *pending_scope = 1; MenuDispatch::Status("Format scope: All".into()) }
        "format_apply_full_column" => { *pending_scope = 2; MenuDispatch::Status("Format scope: Full column".into()) }
        "format_apply_data" => { *pending_scope = 3; MenuDispatch::Status("Format scope: Data".into()) }
        "format_apply_special" => { *pending_scope = 4; MenuDispatch::Status("Format scope: Special".into()) }
        "format_apply_cell" => { *pending_scope = 0; MenuDispatch::Status("Format scope: Cell".into()) }
        "format_apply_selection" => { *pending_scope = 5; MenuDispatch::Status("Format scope: Selection".into()) }
        "format_decimal_generic" => { apply_format(app, *pending_scope, main_row, main_col, CellFormat { number: Some(NumberFormat::DecimalGeneric), align: None }); MenuDispatch::Status("Format: Decimal (generic)".into()) }
        "format_currency" => { apply_format(app, *pending_scope, main_row, main_col, CellFormat { number: Some(NumberFormat::Currency { decimals: 2 }), align: None }); MenuDispatch::Status("Format: Currency ($)".into()) }
        "format_rational" => { apply_format(app, *pending_scope, main_row, main_col, CellFormat { number: Some(NumberFormat::Rational), align: None }); MenuDispatch::Status("Format: Rational".into()) }
        "format_fixed_0" => { apply_format(app, *pending_scope, main_row, main_col, CellFormat { number: Some(NumberFormat::Fixed { decimals: 0 }), align: None }); MenuDispatch::Status("Format: Fixed 0".into()) }
        "format_fixed_1" => { apply_format(app, *pending_scope, main_row, main_col, CellFormat { number: Some(NumberFormat::Fixed { decimals: 1 }), align: None }); MenuDispatch::Status("Format: Fixed 1".into()) }
        "format_fixed_2" => { apply_format(app, *pending_scope, main_row, main_col, CellFormat { number: Some(NumberFormat::Fixed { decimals: 2 }), align: None }); MenuDispatch::Status("Format: Fixed 2".into()) }
        "format_fixed_custom" => { apply_format(app, *pending_scope, main_row, main_col, CellFormat { number: Some(NumberFormat::Fixed { decimals: 2 }), align: None }); MenuDispatch::Status("Format: Fixed n".into()) }
        "format_align_left" => { apply_format(app, *pending_scope, main_row, main_col, CellFormat { number: None, align: Some(TextAlign::Left) }); MenuDispatch::Status("Format: Align Left".into()) }
        "format_align_center" => { apply_format(app, *pending_scope, main_row, main_col, CellFormat { number: None, align: Some(TextAlign::Center) }); MenuDispatch::Status("Format: Align Center".into()) }
        "format_align_right" => { apply_format(app, *pending_scope, main_row, main_col, CellFormat { number: None, align: Some(TextAlign::Right) }); MenuDispatch::Status("Format: Align Right".into()) }
        "format_align_default" => { apply_format(app, *pending_scope, main_row, main_col, CellFormat { number: None, align: Some(TextAlign::Default) }); MenuDispatch::Status("Format: Align Default".into()) }
        "format_reset" => { apply_format(app, *pending_scope, main_row, main_col, CellFormat { number: None, align: None }); MenuDispatch::Status("Format reset".into()) }
        "about" => MenuDispatch::About { status: "About".into() },
        "help_keybinds" => MenuDispatch::HelpKeybinds { status: "Help".into() },
        "toggle_headers" | "toggle_margins" => MenuDispatch::Status(format!("Toggle: {name} (fixed chrome in this build)")),
        "balance_books" => MenuDispatch::Status("Balance books: create a report from a data sheet — not available in the pancurses build".into()),
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
                match crate::io::load_workbook_snapshot(std::path::Path::new(&path)) {
                    Ok(snap) => {
                        app.core.workbook = crate::ops::WorkbookState::from_snapshot(&snap);
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
                let p = std::path::Path::new(&path);
                let snap = crate::ops::WorkbookSnapshot::from_workbook(&app.core.workbook);
                match crate::io::save_workbook(p, &snap) {
                    Ok(()) => {
                        app.core.path = Some(std::path::PathBuf::from(path.clone()));
                        app.core.status = format!("Saved to {path}");
                    }
                    Err(e) => app.core.status = format!("Save error: {e}"),
                }
            }
        }
        "export_tsv" | "export_csv" | "export_ods" | "export_ascii" | "export_all" => {
            if !path.is_empty() {
                let g = app.core.workbook.active_sheet().grid.clone();
                match std::fs::File::create(&path) {
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
                            Ok(()) => app.core.status = format!("Exported {action} to {path}"),
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
                if let Some(ref p) = app.core.path.clone() {
                    let mut active_sheet = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
                    let _ = crate::io::commit_workbook_op(p, &mut app.core.offset, &mut app.core.workbook, &mut active_sheet, &wbo);
                    app.core.ops_applied = app.core.ops_applied.saturating_add(1);
                }
                app.core.workbook.sheets.remove(idx);
                if app.core.workbook.active_sheet >= idx {
                    app.core.workbook.active_sheet = app.core.workbook.active_sheet.saturating_sub(1);
                }
                app.core.status = if path.is_empty() { "Deleted active sheet".into() } else { format!("Deleted sheet {path}") };
            } else {
                app.core.status = "Cannot delete the last sheet".into();
            }
        }
        "insert_special_chars" | "insert_hyperlink" => {
            if !path.is_empty() {
                let hr = HEADER_ROWS;
                let lm = MARGIN_COLS;
                let addr = CellAddr::Main {
                    row: app.core.cursor.row.saturating_sub(hr) as u32,
                    col: app.core.cursor.col.saturating_sub(lm) as u32,
                };
                commit_cell(app, addr.clone(), path.clone());
                app.core.status = if action == "insert_special_chars" {
                    format!("Inserted '{}'", path)
                } else {
                    format!("Inserted hyperlink {path}")
                };
            }
        }
        _ => { app.core.status = format!("Menu: {action}"); }
    }
}
