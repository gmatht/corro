use crate::gui::dialogs;
use rustxwidgets::common::{SimpleAction, MenuItemDef, SubmenuDef};

pub struct MenuAction {
    pub label: &'static str,
    pub shortcut: &'static str,
    pub action: MenuActionKind,
}

#[derive(Clone, Copy)]
pub enum MenuActionKind {
    Open,
    Save,
    SaveAs,
    Quit,
    Undo,
    Redo,
    Cut,
    Copy,
    Paste,
    Find,
    Replace,
    DeleteCell,
    SelectAll,
    ToggleHeaders,
    ToggleMargins,
    NewSheet,
    RenameSheet,
    DeleteSheet,
    SortAsc,
    SortDesc,
    BalanceBooks,
    ExportTsv,
    ExportCsv,
    ExportOds,
    ExportAscii,
    About,
    HelpKeybinds,
    // Insert menu
    InsertRows,
    InsertMitosisRow,
    InsertMitosisCol,
    InsertCols,
    InsertSpecialChars,
    InsertDate,
    InsertTime,
    InsertHyperlink,
    // Format menu
    FormatApplyAll,
    FormatApplyFullColumn,
    FormatApplyData,
    FormatApplySpecial,
    FormatApplyCell,
    FormatApplySelection,
    FormatDecimalGeneric,
    FormatCurrency,
    FormatRational,
    FormatFixed0,
    FormatFixed1,
    FormatFixed2,
    FormatFixedCustom,
    FormatAlignLeft,
    FormatAlignCenter,
    FormatAlignRight,
    FormatAlignDefault,
    FormatReset,
    ExportAll,
}

pub fn action_kind_to_name(kind: MenuActionKind) -> &'static str {
    match kind {
        MenuActionKind::Open => "open",
        MenuActionKind::Save => "save",
        MenuActionKind::SaveAs => "save_as",
        MenuActionKind::Quit => "corro_quit",
        MenuActionKind::Undo => "undo",
        MenuActionKind::Redo => "redo",
        MenuActionKind::Cut => "cut",
        MenuActionKind::Copy => "copy",
        MenuActionKind::Paste => "paste",
        MenuActionKind::Find => "find",
        MenuActionKind::Replace => "replace",
        MenuActionKind::DeleteCell => "delete_cell",
        MenuActionKind::SelectAll => "select_all",
        MenuActionKind::ToggleHeaders => "toggle_headers",
        MenuActionKind::ToggleMargins => "toggle_margins",
        MenuActionKind::NewSheet => "new_sheet",
        MenuActionKind::RenameSheet => "rename_sheet",
        MenuActionKind::DeleteSheet => "delete_sheet",
        MenuActionKind::SortAsc => "sort_asc",
        MenuActionKind::SortDesc => "sort_desc",
        MenuActionKind::BalanceBooks => "balance_books",
        MenuActionKind::ExportTsv => "export_tsv",
        MenuActionKind::ExportCsv => "export_csv",
        MenuActionKind::ExportOds => "export_ods",
        MenuActionKind::ExportAscii => "export_ascii",
        MenuActionKind::About => "about",
        MenuActionKind::HelpKeybinds => "help_keybinds",
        MenuActionKind::InsertRows => "insert_rows",
        MenuActionKind::InsertMitosisRow => "insert_mitosis_row",
        MenuActionKind::InsertMitosisCol => "insert_mitosis_col",
        MenuActionKind::InsertCols => "insert_cols",
        MenuActionKind::InsertSpecialChars => "insert_special_chars",
        MenuActionKind::InsertDate => "insert_date",
        MenuActionKind::InsertTime => "insert_time",
        MenuActionKind::InsertHyperlink => "insert_hyperlink",
        MenuActionKind::FormatApplyAll => "format_apply_all",
        MenuActionKind::FormatApplyFullColumn => "format_apply_full_column",
        MenuActionKind::FormatApplyData => "format_apply_data",
        MenuActionKind::FormatApplySpecial => "format_apply_special",
        MenuActionKind::FormatApplyCell => "format_apply_cell",
        MenuActionKind::FormatApplySelection => "format_apply_selection",
        MenuActionKind::FormatDecimalGeneric => "format_decimal_generic",
        MenuActionKind::FormatCurrency => "format_currency",
        MenuActionKind::FormatRational => "format_rational",
        MenuActionKind::FormatFixed0 => "format_fixed_0",
        MenuActionKind::FormatFixed1 => "format_fixed_1",
        MenuActionKind::FormatFixed2 => "format_fixed_2",
        MenuActionKind::FormatFixedCustom => "format_fixed_custom",
        MenuActionKind::FormatAlignLeft => "format_align_left",
        MenuActionKind::FormatAlignCenter => "format_align_center",
        MenuActionKind::FormatAlignRight => "format_align_right",
        MenuActionKind::FormatAlignDefault => "format_align_default",
        MenuActionKind::FormatReset => "format_reset",
        MenuActionKind::ExportAll => "export_all",
    }
}

/// Create a SimpleAction, connect its callback, and register it.
pub fn register_action<F: FnMut() + 'static>(
    rxapp: &rustxwidgets::App,
    name: &str,
    mut f: F,
) -> Result<SimpleAction, Box<dyn std::error::Error>> {
    let action = rxapp.new_simple_action(name)?;
    action.connect_activate(move |_| f())?;
    rxapp.register_action(&action)?;
    Ok(action)
}

/// Execute a menu action by name, wiring it to the appropriate dialog or stub.
/// This is a shared handler intended for non-GUI backends (pancurses/ratatui).
/// Currently unused — the GUI backend uses its own `handle_menu_action` in
/// `gui_backend.rs`. Keep the `eprintln!` stubs until real logic is wired.
#[allow(dead_code)]
pub fn handle_action(name: &str, rxapp: &rustxwidgets::App) {
    match name {
        "open" => {
            if let Some(path) = dialogs::file_open_dialog() {
                eprintln!("Open file: {:?}", path);
            }
        }
        "save" => {
            if let Some(path) = dialogs::file_save_dialog() {
                eprintln!("Save file: {:?}", path);
            }
        }
        "save_as" => {
            if let Some(path) = dialogs::file_save_dialog() {
                eprintln!("Save file as: {:?}", path);
            }
        }
        "quit" => {
            rxapp.quit();
        }
        "find" => dialogs::find_dialog(|result| {
            if let Some(text) = result {
                eprintln!("Find: {}", text);
            }
        }),
        "replace" => dialogs::replace_dialog(|result| {
            if let Some((find, replace)) = result {
                eprintln!("Replace: '{}' with '{}'", find, replace);
            }
        }),
        "sort_asc" => {
            let wb = crate::ops::WorkbookState::default();
            dialogs::sort_dialog(&wb, |result| {
                if let Some((col, asc)) = result {
                    eprintln!("Sort col {} asc: {}", col, asc);
                }
            });
        }
        "sort_desc" => {
            let wb = crate::ops::WorkbookState::default();
            dialogs::sort_dialog(&wb, |result| {
                if let Some((col, asc)) = result {
                    eprintln!("Sort col {} desc: {}", col, !asc);
                }
            });
        }
        "balance_books" => dialogs::balance_dialog(|result| {
            if let Some(col) = result {
                eprintln!("Balance col: {}", col);
            }
        }),
        "about" => dialogs::show_about_dialog(),
        "help_keybinds" => dialogs::show_keybinds_help(),
        "rename_sheet" => dialogs::find_dialog(|result| {
            if let Some(name) = result {
                eprintln!("Rename sheet to: {}", name);
            }
        }),
        _ => eprintln!("Menu action: {name}"),
    }
}

pub const FILE_MENU: &[MenuAction] = &[
    MenuAction { label: "_Open",        shortcut: "Ctrl+O",       action: MenuActionKind::Open },
    MenuAction { label: "_Save",        shortcut: "Ctrl+S",       action: MenuActionKind::Save },
    MenuAction { label: "Save _As",     shortcut: "Ctrl+Shift+S", action: MenuActionKind::SaveAs },
    MenuAction { label: "_Quit",        shortcut: "Ctrl+Q",       action: MenuActionKind::Quit },
];

pub const TOOLS_MENU: &[MenuAction] = &[
    MenuAction { label: "Export T_SV",  shortcut: "",             action: MenuActionKind::ExportTsv },
    MenuAction { label: "Export _Csv",  shortcut: "",             action: MenuActionKind::ExportCsv },
    MenuAction { label: "Export O_DS",  shortcut: "",             action: MenuActionKind::ExportOds },
    MenuAction { label: "Export ASC_II",shortcut: "",             action: MenuActionKind::ExportAscii },
    MenuAction { label: "Export Al_l",  shortcut: "",             action: MenuActionKind::ExportAll },
];

pub const EDIT_MENU: &[MenuAction] = &[
    MenuAction { label: "_Undo",         shortcut: "Ctrl+Z", action: MenuActionKind::Undo },
    MenuAction { label: "_Redo",         shortcut: "Ctrl+Y", action: MenuActionKind::Redo },
    MenuAction { label: "_Cut",          shortcut: "Ctrl+X", action: MenuActionKind::Cut },
    MenuAction { label: "C_opy",         shortcut: "Ctrl+C", action: MenuActionKind::Copy },
    MenuAction { label: "_Paste",        shortcut: "Ctrl+V", action: MenuActionKind::Paste },
    MenuAction { label: "_Delete",       shortcut: "Del",    action: MenuActionKind::DeleteCell },
    MenuAction { label: "_Select All",   shortcut: "Ctrl+A", action: MenuActionKind::SelectAll },
    MenuAction { label: "_Find",         shortcut: "Ctrl+F", action: MenuActionKind::Find },
    MenuAction { label: "R_eplace",      shortcut: "Ctrl+H", action: MenuActionKind::Replace },
];

pub const VIEW_MENU: &[MenuAction] = &[
    MenuAction { label: "_Toggle Headers", shortcut: "", action: MenuActionKind::ToggleHeaders },
    MenuAction { label: "Toggle _Margins", shortcut: "", action: MenuActionKind::ToggleMargins },
];

pub const SHEET_MENU: &[MenuAction] = &[
    MenuAction { label: "_New Sheet",   shortcut: "", action: MenuActionKind::NewSheet },
    MenuAction { label: "_Rename Sheet",shortcut: "", action: MenuActionKind::RenameSheet },
    MenuAction { label: "_Delete Sheet",shortcut: "", action: MenuActionKind::DeleteSheet },
];

pub const INSERT_MENU: &[MenuAction] = &[
    MenuAction { label: "_Rows",          shortcut: "", action: MenuActionKind::InsertRows },
    MenuAction { label: "_Mitosis (Row)", shortcut: "", action: MenuActionKind::InsertMitosisRow },
    MenuAction { label: "Mi_tosis (Col)", shortcut: "", action: MenuActionKind::InsertMitosisCol },
    MenuAction { label: "_Cols",          shortcut: "", action: MenuActionKind::InsertCols },
    MenuAction { label: "_Special Char",  shortcut: "", action: MenuActionKind::InsertSpecialChars },
    MenuAction { label: "_Date",          shortcut: "", action: MenuActionKind::InsertDate },
    MenuAction { label: "_Time",          shortcut: "", action: MenuActionKind::InsertTime },
    MenuAction { label: "_Hyperlink",     shortcut: "", action: MenuActionKind::InsertHyperlink },
];

pub const FORMAT_MENU: &[MenuAction] = &[
    MenuAction { label: "Scope: _All",        shortcut: "", action: MenuActionKind::FormatApplyAll },
    MenuAction { label: "Scope: _Full Col",   shortcut: "", action: MenuActionKind::FormatApplyFullColumn },
    MenuAction { label: "Scope: _Data",       shortcut: "", action: MenuActionKind::FormatApplyData },
    MenuAction { label: "Scope: _Special",    shortcut: "", action: MenuActionKind::FormatApplySpecial },
    MenuAction { label: "Scope: Ce_ll",       shortcut: "", action: MenuActionKind::FormatApplyCell },
    MenuAction { label: "Scope: _Selection",  shortcut: "", action: MenuActionKind::FormatApplySelection },
    MenuAction { label: "Decimal (_generic)", shortcut: "", action: MenuActionKind::FormatDecimalGeneric },
    MenuAction { label: "C_urrency ($)",      shortcut: "", action: MenuActionKind::FormatCurrency },
    MenuAction { label: "Rat_ional",          shortcut: "", action: MenuActionKind::FormatRational },
    MenuAction { label: "F_ixed 0",           shortcut: "", action: MenuActionKind::FormatFixed0 },
    MenuAction { label: "Fi_xed 1",           shortcut: "", action: MenuActionKind::FormatFixed1 },
    MenuAction { label: "Fixed _2",           shortcut: "", action: MenuActionKind::FormatFixed2 },
    MenuAction { label: "Fixed _n",           shortcut: "", action: MenuActionKind::FormatFixedCustom },
    MenuAction { label: "Align _Left",        shortcut: "", action: MenuActionKind::FormatAlignLeft },
    MenuAction { label: "Align _Center",      shortcut: "", action: MenuActionKind::FormatAlignCenter },
    MenuAction { label: "Align _Right",       shortcut: "", action: MenuActionKind::FormatAlignRight },
    MenuAction { label: "Align D_efault",     shortcut: "", action: MenuActionKind::FormatAlignDefault },
    MenuAction { label: "Rese_t",             shortcut: "", action: MenuActionKind::FormatReset },
];

pub const DATA_MENU: &[MenuAction] = &[
    MenuAction { label: "Sort _Ascending",  shortcut: "", action: MenuActionKind::SortAsc },
    MenuAction { label: "Sort _Descending", shortcut: "", action: MenuActionKind::SortDesc },
    MenuAction { label: "_Balance Books",   shortcut: "", action: MenuActionKind::BalanceBooks },
];

pub const HELP_MENU: &[MenuAction] = &[
    MenuAction { label: "_Keybindings", shortcut: "F1", action: MenuActionKind::HelpKeybinds },
    MenuAction { label: "_About",       shortcut: "",   action: MenuActionKind::About },
];

fn actions_to_defs(items: &[MenuAction]) -> Vec<MenuItemDef> {
    items.iter().map(|a| MenuItemDef {
        label: a.label,
        action: action_kind_to_name(a.action),
        submenu: None,
    }).collect()
}

/// All submenus as SubmenuDef slices — used by rustxwidgets::App::build_menu_model().
pub fn all_submenus() -> Vec<SubmenuDef> {
    // Deliberately leak Vec backing buffers so the returned slices live forever.
    vec![
        SubmenuDef { label: "File",   prefix: "app", items: actions_to_defs(FILE_MENU).leak() },
        SubmenuDef { label: "Edit",   prefix: "app", items: actions_to_defs(EDIT_MENU).leak() },
        SubmenuDef { label: "View",   prefix: "app", items: actions_to_defs(VIEW_MENU).leak() },
        SubmenuDef { label: "Insert", prefix: "app", items: actions_to_defs(INSERT_MENU).leak() },
        SubmenuDef { label: "Format", prefix: "app", items: actions_to_defs(FORMAT_MENU).leak() },
        SubmenuDef { label: "Sheet",  prefix: "app", items: actions_to_defs(SHEET_MENU).leak() },
        SubmenuDef { label: "Data",   prefix: "app", items: actions_to_defs(DATA_MENU).leak() },
        SubmenuDef { label: "Tools",  prefix: "app", items: actions_to_defs(TOOLS_MENU).leak() },
        SubmenuDef { label: "Help",   prefix: "app", items: actions_to_defs(HELP_MENU).leak() },
    ]
}

/// Menu bar display text for backends that render a static text bar (pancurses).
pub fn menu_bar_text() -> String {
    let defs = all_submenus();
    let mut s = String::new();
    for (i, sm) in defs.iter().enumerate() {
        if i > 0 { s.push_str("   "); }
        if i == 0 {
            s.push('[');
            s.push_str(sm.label);
            s.push(']');
        } else {
            s.push_str(sm.label);
        }
    }
    s
}
