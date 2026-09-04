use crate::gui::dialogs;
use crate::gui::app_alias::App;
use rustxwidgets::{Menu, SimpleAction};

pub struct MenuAction {
    pub label: &'static str,
    pub shortcut: &'static str,
    pub action: MenuActionKind,
    /// When set, this item opens a submenu instead of dispatching an action.
    pub submenu: Option<Vec<MenuAction>>,
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
        if let Some(sub) = item.submenu.as_deref() {
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

/// Build a `&'static [MenuAction]` from a nested tree of items.  An item is
/// `"Label" => ActionKind`, or `"Label" => [ items ]` for a submenu.  The tree
/// *is* the menu — no separate submenu constants.  (Accelerators are a later
/// migration phase; the `shortcut` field is currently always empty.)
macro_rules! menu_items {
    () => { Vec::new() };
    ($label:literal => $action:ident) => {
        vec![MenuAction { label: $label, shortcut: "", action: MenuActionKind::$action, submenu: None }]
    };
    ($label:literal => $action:ident ($shortcut:literal)) => {
        vec![MenuAction { label: $label, shortcut: $shortcut, action: MenuActionKind::$action, submenu: None }]
    };
    ($label:literal => $action:ident, $($rest:tt)*) => {
        {
            let mut v = vec![MenuAction { label: $label, shortcut: "", action: MenuActionKind::$action, submenu: None }];
            v.extend(menu_items!($($rest)*));
            v
        }
    };
    ($label:literal => $action:ident ($shortcut:literal), $($rest:tt)*) => {
        {
            let mut v = vec![MenuAction { label: $label, shortcut: $shortcut, action: MenuActionKind::$action, submenu: None }];
            v.extend(menu_items!($($rest)*));
            v
        }
    };
    ($label:literal => [ $($items:tt)* ]) => {
        vec![MenuAction { label: $label, shortcut: "", action: MenuActionKind::Submenu, submenu: Some(menu_items!($($items)*)) }]
    };
    ($label:literal => ($shortcut:literal) [ $($items:tt)* ]) => {
        vec![MenuAction { label: $label, shortcut: $shortcut, action: MenuActionKind::Submenu, submenu: Some(menu_items!($($items)*)) }]
    };
    ($label:literal => [ $($items:tt)* ], $($rest:tt)*) => {
        {
            let mut v = vec![MenuAction { label: $label, shortcut: "", action: MenuActionKind::Submenu, submenu: Some(menu_items!($($items)*)) }];
            v.extend(menu_items!($($rest)*));
            v
        }
    };
    ($label:literal => ($shortcut:literal) [ $($items:tt)* ], $($rest:tt)*) => {
        {
            let mut v = vec![MenuAction { label: $label, shortcut: $shortcut, action: MenuActionKind::Submenu, submenu: Some(menu_items!($($items)*)) }];
            v.extend(menu_items!($($rest)*));
            v
        }
    };
}

/// The application menu bar as a nested tree (matches the ratatui reference:
/// File, Edit, Insert, Format, Sheet, Help).  Every backend builds its native
/// menu from this single definition.
pub fn menu_bar() -> Vec<MenuAction> {
    menu_items! {
    "File" => [
        "Open file"    => Open ("O"),
        "Save as"      => SaveAs ("A"),
        "Export"       => ("T") [
            "TSV"         => ExportTsv ("T"),
            "CSV"         => ExportCsv ("C"),
            "ASCII table" => ExportAscii ("A"),
            "Export all"  => ExportAll ("L"),
            "ODS"         => ExportOds ("D"),
        ],
        "Width"        => ("C") [
            "Default width" => SetMaxColWidth ("D"),
            "Column width"  => SetColWidth ("C"),
        ],
        "Sort view"    => SortView ("S"),
        "Persist sort" => SaveSort ("P"),
        "Exit"         => Quit ("X"),
        "Replay"       => Replay ("R"),
    ],
    "Edit" => [
        "Cut"         => Cut ("X"),
        "Copy"        => Copy ("C"),
        "Paste"       => Paste ("P"),
        "Find"        => Find ("F"),
        "Replace"     => Replace ("R"),
        "Duplicate"   => Duplicate ("D"),
        "Extrapolate" => Extrapolate ("E"),
    ],
    "Insert" => [
        "Rows"          => InsertRows ("R"),
        "Mitosis (Row)" => InsertMitosisRow ("M"),
        "Mitosis (Col)" => InsertMitosisCol ("O"),
        "Cols"          => InsertCols ("C"),
        "Special Char"  => InsertSpecialChars ("S"),
        "Date"          => InsertDate (";"),
        "Time"          => InsertTime (":"),
        "Hyperlink"     => InsertHyperlink ("H"),
    ],
    "Format" => [
        "Scope"  => ("S") [
            "All"        => FormatApplyAll ("A"),
            "Full col"   => FormatApplyFullColumn ("F"),
            "Data"       => FormatApplyData ("D"),
            "Special"    => FormatApplySpecial ("S"),
            "Cell"       => FormatApplyCell ("C"),
            "Selection"  => FormatApplySelection ("L"),
        ],
        "Number" => ("N") [
            "Decimal (generic)" => FormatDecimalGeneric ("D"),
            "Currency ($)"      => FormatCurrency ("$"),
            "Rational"          => FormatRational ("R"),
            "Fixed 0"           => FormatFixed0 ("0"),
            "Fixed 1"           => FormatFixed1 ("1"),
            "Fixed 2"           => FormatFixed2 ("2"),
            "Fixed n"           => FormatFixedCustom ("N"),
        ],
        "Align"  => ("A") [
            "Left"    => FormatAlignLeft ("L"),
            "Center"  => FormatAlignCenter ("C"),
            "Right"   => FormatAlignRight ("R"),
            "Default" => FormatAlignDefault ("D"),
        ],
        "Reset"  => FormatReset ("R"),
    ],
    "Sheet" => [
        "Prev sheet"    => SheetPrev ("["),
        "Next sheet"    => SheetNext ("]"),
        "New sheet"     => NewSheet ("N"),
        "Rename sheet"  => RenameSheet ("R"),
        "Copy sheet"    => CopySheet ("C"),
        "Move sheet"    => MoveSheet ("M"),
        "Go"            => GoToCell ("G"),
        "Balance books" => BalanceBooks ("B"),
    ],
    "Help" => [
        "About"     => About ("A"),
        "Row ops"   => HelpRows ("R"),
        "Col ops"   => HelpCols ("C"),
        "Full help" => HelpFull ("H"),
    ],
} }

/// Build a pancurses `Menu` model from a `&[MenuAction]` tree (the pancurses
/// backend's converter for the shared `MENU_BAR` definition).
#[cfg(feature = "pancurses")]
pub fn build_menu_model(menu: &rustxwidgets::backends_pancurses_adapter::Menu, items: &[MenuAction]) {
    for item in items {
        if let Some(sub) = item.submenu.as_deref() {
            let sub_menu = rustxwidgets::backends_pancurses_adapter::create_menu().expect("create submenu");
            build_menu_model(&sub_menu, sub);
            if item.shortcut.is_empty() {
                menu.append_submenu(item.label, &sub_menu);
            } else {
                menu.append_submenu_with_shortcut(item.label, item.shortcut, &sub_menu);
            }
        } else {
            if item.shortcut.is_empty() {
                menu.append(item.label, action_kind_to_name(item.action));
            } else {
                menu.append_with_shortcut(item.label, action_kind_to_name(item.action), item.shortcut);
            }
        }
    }
}
