# Standard action vocabulary

A documented **convention** of generic action strings that any rustxwidgets
app may use, so backends can give them consistent, platform-correct behavior
without knowing the app.  This is the *generic subset* — app-specific actions
(`insert_date`, `balance_books`, …) are **not** part of it.

## The vocabulary

| Action | Meaning | Backend behavior |
|---|---|---|
| `about` | Show the app's About dialog | GTK: native about dialog; pancurses: info dialog |
| `exit` | Quit the app | GTK: close window; pancurses: `running = false` |
| `open` | Open a file | GTK: native file dialog; pancurses: path prompt |
| `save` | Save the current file | GTK: native save dialog; pancurses: path prompt |
| `save_as` | Save to a new path | GTK: native save-as dialog; pancurses: path prompt |
| `new` | Create a new document | backend default or app handler |
| `close` | Close the current document | backend default or app handler |
| `help` | Show help | backend default or app handler |
| `cut` / `copy` / `paste` | Clipboard operations | backend default or app handler |
| `find` / `replace` | Search | backend default or app handler |
| `undo` / `redo` | History | backend default or app handler |

## Rules

1. **It is a convention, not a hardcoded structure.**  Apps *may* use these
   strings; the toolkit recognizes them for default behavior.  An app is free
   to use any other string for its own actions.
2. **App-specific actions are not in the vocabulary.**  `insert_date`,
   `balance_books`, `sort_view` are app-specific and stay out — the toolkit
   must never special-case them (see AGENTS.md "Toolkit purity").
3. **A backend may special-case a standard action** (e.g. macOS moves
   `about`/`exit` to the system menu, mirroring `wxID_ABOUT`/`wxID_EXIT`); it
   must never special-case an app-specific action.
4. **The vocabulary is the generic-vs-specific line.**  The toolkit knows
   about `exit` (generic); it does not know about corro's File menu
   (specific).

## Why it exists

- **Platform semantics**: a backend must recognize `about`/`exit` by name to
  give them platform-correct behavior (macOS system menu).
- **Backend default behavior**: `exit` → close window, `open`/`save` → native
  file dialogs, for free.
- **Consistency + parity testing**: `exit` quits the same way on every
  backend; tests can assert it.
