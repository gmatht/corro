use crate::gui::dialogs;
use crate::gui::app_alias::App;
use rustxwidgets::{Menu, SimpleAction};

pub struct MenuAction {
    pub label: &'static str,
    pub shortcut: &'static str,
    pub action: MenuActionKind,
    /// When set, this item opens a submenu instead of dispatching an action.
    pub submenu: Option<&'static [MenuAction]>,
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
    // Submenu placeholder (item opens a submenu; never dispatched)
    Submenu,
    // File menu (ratatui parity)
    SortView,
    SaveSort,
    Replay,
    SetMaxColWidth,
    SetColWidth,
    ExportAll,
    ExportOdt,
    // Edit menu (ratatui parity)
    Duplicate,
    Extrapolate,
    // Sheet menu (ratatui parity)
    SheetPrev,
    SheetNext,
    CopySheet,
    MoveSheet,
    GoToCell,
    // Help menu (ratatui parity)
    HelpRows,
    HelpCols,
    HelpFull,
}

pub fn action_kind_to_name(kind: MenuActionKind) -> &'static str {
    match kind {
        MenuActionKind::Open => "open",
        MenuActionKind::Save => "save",
        MenuActionKind::SaveAs => "save_as",
        MenuActionKind::Quit => "quit",
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
        MenuActionKind::Submenu => "submenu",
        MenuActionKind::SortView => "sort_view",
        MenuActionKind::SaveSort => "persist_sort",
        MenuActionKind::Replay => "replay",
        MenuActionKind::SetMaxColWidth => "set_max_col_width",
        MenuActionKind::SetColWidth => "set_col_width",
        MenuActionKind::ExportAll => "export_all",
        MenuActionKind::ExportOdt => "export_ods",
        MenuActionKind::Duplicate => "duplicate",
        MenuActionKind::Extrapolate => "extrapolate",
        MenuActionKind::SheetPrev => "sheet_prev",
        MenuActionKind::SheetNext => "sheet_next",
        MenuActionKind::CopySheet => "copy_sheet",
        MenuActionKind::MoveSheet => "move_sheet",
        MenuActionKind::GoToCell => "go_to_cell",
        MenuActionKind::HelpRows => "help_rows",
        MenuActionKind::HelpCols => "help_cols",
        MenuActionKind::HelpFull => "help_full",
    }
}

/// Build a submenu model from action descriptors.
pub fn build_submenu(rxapp: &App, items: &[MenuAction], prefix: &str) -> Result<Menu, Box<dyn std::error::Error>> {
    let menu = rxapp.create_menu()?;
    for item in items {
        if let Some(sub) = item.submenu {
            let sub_menu = build_submenu(rxapp, sub, prefix)?;
            menu.append_submenu(item.label, &sub_menu);
        } else {
            let name = action_kind_to_name(item.action);
            menu.append(item.label, &format!("{}.{}", prefix, name));
        }
    }
    Ok(menu)
}

/// Create a SimpleAction, connect its callback, and register it.
pub fn register_action<F: FnMut() + 'static>(
    rxapp: &App,
    name: &str,
    mut f: F,
) -> Result<SimpleAction, Box<dyn std::error::Error>> {
    let action = rxapp.create_simple_action(name)?;
    #[cfg(feature = "pancurses")]
    action.on_activate(move || f());
    #[cfg(not(feature = "pancurses"))]
    action.connect_activate(move |_| f())?;
    #[cfg(not(feature = "pancurses"))]
    rxapp.register_action(&action)?;
    Ok(action)
}

/// Execute a menu action by name, wiring it to the appropriate dialog or stub.
/// This is called when a menu item is activated.
pub fn handle_action(name: &str) {
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
            #[cfg(all(unix, not(feature = "pancurses")))]
            let _ = rustxwidgets::backends_gtk_adapter::quit_main_loop();
            #[cfg(all(windows, not(feature = "pancurses")))]
            rustxwidgets::backends_nwg_adapter::quit_main_loop();
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
    MenuAction { label: "Open file",    shortcut: "Ctrl+O", action: MenuActionKind::Open, submenu: None },
    MenuAction { label: "Save as",      shortcut: "Ctrl+S", action: MenuActionKind::SaveAs, submenu: None },
    MenuAction { label: "Export",       shortcut: "", action: MenuActionKind::Submenu, submenu: Some(EXPORT_MENU) },
    MenuAction { label: "Width",        shortcut: "", action: MenuActionKind::Submenu, submenu: Some(WIDTH_MENU) },
    MenuAction { label: "Sort view",    shortcut: "", action: MenuActionKind::SortView, submenu: None },
    MenuAction { label: "Persist sort", shortcut: "", action: MenuActionKind::SaveSort, submenu: None },
    MenuAction { label: "Exit",         shortcut: "Ctrl+Q", action: MenuActionKind::Quit, submenu: None },
    MenuAction { label: "Replay",       shortcut: "", action: MenuActionKind::Replay, submenu: None },
];

pub const EXPORT_MENU: &[MenuAction] = &[
    MenuAction { label: "TSV",         shortcut: "", action: MenuActionKind::ExportTsv, submenu: None },
    MenuAction { label: "CSV",         shortcut: "", action: MenuActionKind::ExportCsv, submenu: None },
    MenuAction { label: "ASCII table", shortcut: "", action: MenuActionKind::ExportAscii, submenu: None },
    MenuAction { label: "Export all",  shortcut: "", action: MenuActionKind::ExportAll, submenu: None },
    MenuAction { label: "ODS",         shortcut: "", action: MenuActionKind::ExportOdt, submenu: None },
];

pub const WIDTH_MENU: &[MenuAction] = &[
    MenuAction { label: "Default width", shortcut: "", action: MenuActionKind::SetMaxColWidth, submenu: None },
    MenuAction { label: "Column width",  shortcut: "", action: MenuActionKind::SetColWidth, submenu: None },
];

pub const EDIT_MENU: &[MenuAction] = &[
    MenuAction { label: "Cut",         shortcut: "Ctrl+X", action: MenuActionKind::Cut, submenu: None },
    MenuAction { label: "Copy",        shortcut: "Ctrl+C", action: MenuActionKind::Copy, submenu: None },
    MenuAction { label: "Paste",       shortcut: "Ctrl+V", action: MenuActionKind::Paste, submenu: None },
    MenuAction { label: "Find",        shortcut: "Ctrl+F", action: MenuActionKind::Find, submenu: None },
    MenuAction { label: "Replace",     shortcut: "Ctrl+H", action: MenuActionKind::Replace, submenu: None },
    MenuAction { label: "Duplicate",   shortcut: "", action: MenuActionKind::Duplicate, submenu: None },
    MenuAction { label: "Extrapolate", shortcut: "", action: MenuActionKind::Extrapolate, submenu: None },
];

pub const VIEW_MENU: &[MenuAction] = &[
    MenuAction { label: "Toggle Headers", shortcut: "", action: MenuActionKind::ToggleHeaders, submenu: None },
    MenuAction { label: "Toggle Margins", shortcut: "", action: MenuActionKind::ToggleMargins, submenu: None },
];

pub const SHEET_MENU: &[MenuAction] = &[
    MenuAction { label: "Prev sheet",    shortcut: "", action: MenuActionKind::SheetPrev, submenu: None },
    MenuAction { label: "Next sheet",    shortcut: "", action: MenuActionKind::SheetNext, submenu: None },
    MenuAction { label: "New sheet",     shortcut: "", action: MenuActionKind::NewSheet, submenu: None },
    MenuAction { label: "Rename sheet",  shortcut: "", action: MenuActionKind::RenameSheet, submenu: None },
    MenuAction { label: "Copy sheet",    shortcut: "", action: MenuActionKind::CopySheet, submenu: None },
    MenuAction { label: "Move sheet",    shortcut: "", action: MenuActionKind::MoveSheet, submenu: None },
    MenuAction { label: "Go",            shortcut: "", action: MenuActionKind::GoToCell, submenu: None },
    MenuAction { label: "Balance books", shortcut: "", action: MenuActionKind::BalanceBooks, submenu: None },
];

pub const INSERT_MENU: &[MenuAction] = &[
    MenuAction { label: "Rows",          shortcut: "", action: MenuActionKind::InsertRows, submenu: None },
    MenuAction { label: "Mitosis (Row)", shortcut: "", action: MenuActionKind::InsertMitosisRow, submenu: None },
    MenuAction { label: "Mitosis (Col)", shortcut: "", action: MenuActionKind::InsertMitosisCol, submenu: None },
    MenuAction { label: "Cols",          shortcut: "", action: MenuActionKind::InsertCols, submenu: None },
    MenuAction { label: "Special Char",  shortcut: "", action: MenuActionKind::InsertSpecialChars, submenu: None },
    MenuAction { label: "Date",          shortcut: "", action: MenuActionKind::InsertDate, submenu: None },
    MenuAction { label: "Time",          shortcut: "", action: MenuActionKind::InsertTime, submenu: None },
    MenuAction { label: "Hyperlink",     shortcut: "", action: MenuActionKind::InsertHyperlink, submenu: None },
];

pub const FORMAT_MENU: &[MenuAction] = &[
    MenuAction { label: "Scope",  shortcut: "", action: MenuActionKind::Submenu, submenu: Some(FORMAT_SCOPE_MENU) },
    MenuAction { label: "Number", shortcut: "", action: MenuActionKind::Submenu, submenu: Some(FORMAT_NUMBER_MENU) },
    MenuAction { label: "Align",  shortcut: "", action: MenuActionKind::Submenu, submenu: Some(FORMAT_ALIGN_MENU) },
    MenuAction { label: "Reset",  shortcut: "", action: MenuActionKind::FormatReset, submenu: None },
];

pub const FORMAT_SCOPE_MENU: &[MenuAction] = &[
    MenuAction { label: "All",        shortcut: "", action: MenuActionKind::FormatApplyAll, submenu: None },
    MenuAction { label: "Full col",   shortcut: "", action: MenuActionKind::FormatApplyFullColumn, submenu: None },
    MenuAction { label: "Data",       shortcut: "", action: MenuActionKind::FormatApplyData, submenu: None },
    MenuAction { label: "Special",    shortcut: "", action: MenuActionKind::FormatApplySpecial, submenu: None },
    MenuAction { label: "Cell",       shortcut: "", action: MenuActionKind::FormatApplyCell, submenu: None },
    MenuAction { label: "Selection",  shortcut: "", action: MenuActionKind::FormatApplySelection, submenu: None },
];

pub const FORMAT_NUMBER_MENU: &[MenuAction] = &[
    MenuAction { label: "Decimal (generic)", shortcut: "", action: MenuActionKind::FormatDecimalGeneric, submenu: None },
    MenuAction { label: "Currency ($)",      shortcut: "", action: MenuActionKind::FormatCurrency, submenu: None },
    MenuAction { label: "Rational",          shortcut: "", action: MenuActionKind::FormatRational, submenu: None },
    MenuAction { label: "Fixed 0",           shortcut: "", action: MenuActionKind::FormatFixed0, submenu: None },
    MenuAction { label: "Fixed 1",           shortcut: "", action: MenuActionKind::FormatFixed1, submenu: None },
    MenuAction { label: "Fixed 2",           shortcut: "", action: MenuActionKind::FormatFixed2, submenu: None },
    MenuAction { label: "Fixed n",           shortcut: "", action: MenuActionKind::FormatFixedCustom, submenu: None },
];

pub const FORMAT_ALIGN_MENU: &[MenuAction] = &[
    MenuAction { label: "Left",    shortcut: "", action: MenuActionKind::FormatAlignLeft, submenu: None },
    MenuAction { label: "Center",  shortcut: "", action: MenuActionKind::FormatAlignCenter, submenu: None },
    MenuAction { label: "Right",   shortcut: "", action: MenuActionKind::FormatAlignRight, submenu: None },
    MenuAction { label: "Default", shortcut: "", action: MenuActionKind::FormatAlignDefault, submenu: None },
];

pub const DATA_MENU: &[MenuAction] = &[
    MenuAction { label: "Sort Ascending",  shortcut: "", action: MenuActionKind::SortAsc, submenu: None },
    MenuAction { label: "Sort Descending", shortcut: "", action: MenuActionKind::SortDesc, submenu: None },
    MenuAction { label: "Balance Books",   shortcut: "", action: MenuActionKind::BalanceBooks, submenu: None },
];

pub const HELP_MENU: &[MenuAction] = &[
    MenuAction { label: "About",     shortcut: "", action: MenuActionKind::About, submenu: None },
    MenuAction { label: "Row ops",   shortcut: "", action: MenuActionKind::HelpRows, submenu: None },
    MenuAction { label: "Col ops",   shortcut: "", action: MenuActionKind::HelpCols, submenu: None },
    MenuAction { label: "Full help", shortcut: "", action: MenuActionKind::HelpFull, submenu: None },
];
