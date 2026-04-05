# KEYBINDINGS

This is the proposed canonical keybinding reference for Corro's current TUI.
It favors spreadsheet conventions where they fit, and documents the bindings that are implemented today.

## Modes

| Mode | Purpose |
| --- | --- |
| Normal | Navigate, select, edit, move, export, undo |
| Edit | Enter or replace a cell's raw value |
| Open path | Choose a `.corro`, `.csv`, or `.tsv` file |
| Export | Write to a file or copy to clipboard |
| Width / sort prompts | Set column widths or view sort order |
| Help | Show a compact shortcut summary |
| Quit prompt | Confirm or cancel quitting |

## Normal Mode

| Key | Action | Notes |
| --- | --- | --- |
| `Arrows`, `hjkl` | Move cursor | Moves through header, main, footer, and side margins |
| `Shift+Arrows` | Extend selection | Starts or grows a rectangular selection |
| `Ctrl+Shift++` | Insert rows | Inserts rows above the current row or selected rows |
| `Alt+I` | Open Insert menu | Menu includes rows, cols, special chars, and hyperlink actions |
| `v` | Toggle cell selection | Selects the current rectangle anchor, or clears selection |
| `e`, `Enter` | Edit current cell | Starts with the current displayed value |
| Any printable key | Start editing | The first typed character seeds the edit buffer |
| `o` | Open path prompt | Set or change the active file path |
| `Alt+F` | Open path prompt | Same as `o`, via the accelerator layer |
| `N` | New sheet | Adds a new blank sheet to the workbook |
| `Ctrl+PageUp`, `Ctrl+PageDown` | Switch sheets | Moves to the previous / next sheet in the tab bar |
| `t` | Export TSV prompt | Selection-aware when a selection exists |
| `c` | Export CSV or move cols | Exports CSV when no selection exists; otherwise moves selected columns |
| `r` | Move rows or select rows | Expands to full rows first, then moves selected rows on the next `r` |
| `Delete`, `Backspace` | Clear cell or selection | Deletes the current cell if nothing is selected |
| `Ctrl+Z` | Undo | Appends and applies the inverse op |
| `Ctrl+Q` | Quit immediately | Bypasses the quit prompt |
| `?`, `h` | Help | Opens the help view |
| `Esc` | Cancel selection or quit prompt | Clears the active selection; otherwise opens the quit prompt |
| `Alt+Arrow` | Move selected rows/cols by one | Works only when a full-row or full-column selection exists |

## Menu Mode

Menu mode is an internal accelerator layer. It is mostly reached through `Alt` shortcuts and can be closed with `Esc`.

| Key | Action | Notes |
| --- | --- | --- |
| `F` | Open path prompt | Same as `o` / `Alt+F` |
| `R` | Explain row ops | Returns to Normal mode with a hint in the status line |
| `C` | Explain col ops | Returns to Normal mode with a hint in the status line |
| `I` | Open Insert menu | Same as `Alt+I` |
| `T` | Export TSV prompt | Same as `t` / `Alt+T` |
| `Alt+A` / `A` | Export ASCII prompt | Opens the ASCII export prompt |
| `E` | Export full prompt | Same as `Alt+E` |
| `W` | Set default width | Same as `Alt+W` |
| `S` | Sort view prompt | Same as `Alt+S` |
| `X` | Exit | Opens the quit prompt |
| `?` | Help | Same as `?` from Normal mode |
| `A` | About | Opens the about page |
| `Esc` | Close menu | Returns to Normal mode |

## Selection Flow

| Key | Action | Notes |
| --- | --- | --- |
| `v` | Start or clear a selection | Selection is cell-based by default |
| `r` | Expand to full rows | If rows are already selected, the current row becomes the move target |
| `c` | Expand to full columns | If columns are already selected, the current column becomes the move target |
| `Esc` | Cancel selection | Leaves the cursor in place |

Row and column moves use the current cursor as the destination target.
The target must be inside the main grid, not in headers or margins.

## Edit Mode

| Key | Action | Notes |
| --- | --- | --- |
| `Enter` | Commit edit | Saves the buffer to the current cell |
| `Esc` | Discard edit | Returns to Normal mode |
| `Backspace` | Delete character | Standard text editing |
| Printable characters | Insert text | Raw text is stored as-is |
| `=` + `Arrows` | Formula ref builder | When the buffer is just `=`, arrow keys move the referenced cell and rewrite the formula reference |

Edit mode accepts plain values or formulas beginning with `=`.
Cross-sheet references use numeric sheet IDs in formulas, like `#2!A1`.
It also accepts shorthand cell-targeted edits like `A1: value`, `^A:A1`, `_B:A1`, `<0: value`, and `>0: value`.

## Open Path Prompt

| Key | Action | Notes |
| --- | --- | --- |
| `Enter` | Open path | `.tsv` and `.csv` are imported, `link <file> <revision>` opens a log snapshot, other files are treated as append-only log files |
| `Esc` | Cancel | Returns to Normal mode |
| Printable characters | Edit path | Standard text entry |
| `Backspace` | Delete character | Standard text editing |

If the chosen file does not exist, Corro treats it as a new file and creates the log on first write.
When more than one sheet exists, Corro shows a bottom tab bar with the active sheet highlighted.

## Export Prompts

| Key | Action | Notes |
| --- | --- | --- |
| `Enter` | Export / copy | Blank filename means clipboard |
| `Esc` | Cancel | Returns to Normal mode |
| Printable characters | Edit filename | Standard text entry |
| `Backspace` | Delete character | Standard text editing |

Export actions:

| Trigger | Output |
| --- | --- |
| `t` | TSV export of the main sheet, or the selected rectangle if one exists |
| `c` | CSV export of the main sheet when no selection exists |
| `Alt+T` / `T` | TSV export prompt |
| `Alt+A` / `A` | ASCII table export prompt |
| `Alt+E` / `E` | Full-sheet export prompt, including headers and margins |

## Width and Sort

| Key | Action | Notes |
| --- | --- | --- |
| `Alt+W` / `W` | Set default column width | Prompts for a number |
| `Alt+X` / `X` | Set or clear a column override | Prompt accepts `col=width` or just `col` to clear the override |
| `Alt+S` / `S` | Sort view | Prompt accepts main-column letters like `A,B,C`; uppercase `S` saves |

Column width prompts use zero-based main-column indexes.
Sort order uses Excel-style main-column letters.

## Help And Quit

| Key | Action | Notes |
| --- | --- | --- |
| `?`, `h` | Open help | From Normal mode; help is a full-page scrollable screen |
| `A` | Open about | From Help mode or menu |
| `Esc` | Close help/about | Also closes the menu and most prompts |
| `q` | Open quit prompt | Normal mode only |
| `Q` | Confirm quit | Quit prompt only |
| `B`, `Esc` | Cancel quit prompt | Quit prompt only |

## Common Spreadsheet Shortcuts

These are common spreadsheet keys from Excel and Google Sheets.
They are documented here as the preferred long-term direction, but they are not all implemented yet.

| Shortcut | Status | Notes |
| --- | --- | --- |
| `Ctrl+C`, `Ctrl+X`, `Ctrl+V` | Reserved | Copy, cut, paste |
| `Ctrl+Y`, `Ctrl+Shift+Z` | Reserved | Redo |
| `Ctrl+Arrow` | Reserved | Jump to the edge of a data region |
| `Ctrl+Home`, `Ctrl+End` | Reserved | Jump to sheet corners |
| `F2` | Reserved | Enter cell edit mode |
| `Ctrl+PageUp`, `Ctrl+PageDown` | Reserved | Sheet navigation |
| `F4` | Reserved | Toggle absolute / relative references |
| `Alt+=` | Reserved | Auto-sum |
| `Ctrl+Shift+L` | Reserved | Toggle filters |
| `Ctrl+Alt+M` | Reserved | Comment / note insertion |
| `Ctrl+~` | Reserved | Show formulas |

## Notes

| Topic | Behavior |
| --- | --- |
| Empty edit buffer | Commits an empty string, which clears the cell |
| Selection delete | Deletes every selected cell in the rectangle |
| Undo | Replays the inverse op as a new log entry when a file is open |
| Clipboard fallback | Blank export filename copies to clipboard when a clipboard tool is available |
| Menu mode | Internal command-menu state; Alt accelerators jump straight to actions |
