# Settings: TUI/GUI options & format proposal

Status: **proposal only — nothing here is implemented.** This document
records the agreed direction so a future implementation has a target.
See also: `Principles.txt` (project invariants), `ALGORITHMS.md`
(grid internals).

## 1. Problem

User-facing behavior is currently controlled by three disjoint mechanisms:

- **CLI flags** (`--gui`, `--ratatui`, `--export`, `--movie*`, …).
- **`CORRO_*` environment variables** (`CORRO_TEMPLATE`, `CORRO_LOG`,
  `CORRO_IDLE_MARKER`, `CORRO_DEBUG_LOG`, `CORRO_DEBUG_OVERLAY`,
  `CORRO_AUTO_UNSAVED`, `CORRO_TERM_COLS/ROWS`, `CORRO_KEY_LOG`, …).
  Half of these are test/diagnostic harnesses, half are real options —
  nothing distinguishes them.
- **Hardcoded constants** (`FONT_SIZE = 12.0`, `"monospace"` everywhere in
  `gui_backend.rs`, 120×40-ish layout assumptions, English-only `TOTAL`
  display seed).

There is no config file, no settings UI, and no documented precedence.
This proposal adds the first two layers (file + documented precedence)
while explicitly deferring a settings dialog.

## 2. Non-goals (reaffirmed)

- **No settings dialog/UI in this round.** Editing stays in the user's
  text editor; the app never writes the config file itself.
- **No per-sheet persisted UI state.** Pins, cursor, viewport, selection
  stay session-only and are never written anywhere (existing decision).
- **No live reload (initially).** Settings apply at startup. Hot-reload
  is future work with its own test plan.
- **No keybinding remapping yet.** Keymaps stay compiled in; remappable
  bindings would be a separate proposal (input pipeline differs per
  backend more than any other surface).

## 3. Precedence (highest wins)

1. CLI flags (explicit per-invocation intent).
2. Environment variables (`CORRO_*` — scripting/CI overrides).
3. Config file (user persistent preferences).
4. Built-in defaults (current hardcoded behavior, unchanged).

Every existing `CORRO_*` var keeps working with identical meaning; the
config file only *adds* persistence for things that currently require
the environment on every launch. Anything settable in more than one
layer resolves by the chain above, and `--help` documents the chain.

## 4. Config file location & format (proposed)

- **Location:** `$XDG_CONFIG_HOME/corro/config.toml`, falling back to
  `~/.corro/config.toml` (mirrors the existing XDG-state convention
  used for debug logs). Missing file = defaults (never an error).
- **Format: TOML.** Rationale: typed values, comments, sections map
  cleanly onto per-backend option groups. Open question: no TOML parser
  is currently vendored — either add `toml` (small, pure Rust) or start
  with a strict documented subset. Do NOT invent a new format.
- **`[meta] version = 1` header required.** Unknown top-level keys and
  unknown per-section keys are **ignored with a one-line stderr warning**
  (forward compatibility: old binaries keep working on newer files).
  A wrong-typed value for a known key is also warn-and-default, never
  abort — a typo must not brick startup.
- Validation happens once at startup; the resolved table is logged at
  debug level (`CORRO_DEBUG_LOG`) so "why did my setting not apply" is
  answerable.

## 5. Option inventory (proposed keys)

Only keys with a concrete consumer are listed. Anything speculative
stays out until it has one.

```toml
[meta]
version = 1

[document]
# Template workbook for new documents (extends CORRO_TEMPLATE today).
# Unset/empty/missing file => built-in seeded blank, with a status note
# on failure (same fallback rule as the env var).
template = "~/sheets/Report.corro"

[startup]
# Backend when no --gui/--ratatui flag is given. Values: "ratatui" | "gui".
# CLI flags still win; gcorro* binaries keep defaulting to GUI.
backend = "ratatui"

[display]
# Base font size in points (today: hardcoded 12.0 in gui_backend.rs).
font_size = 12.0
# Font family preference list, first available wins per backend.
font_family = ["monospace"]

[display.gtk]
# GTK font description suffix, e.g. "Monospace 12" mapping; empty = default.
# font = ""

[display.nwg]
# Win32 LOGFONT face name; empty = system default.
# font = ""

[display.tui]
# Reserved for terminal-specific display (unicode width handling, mouse
# protocol). No keys proposed yet — section exists so future keys land here.

[behavior]
# Autosave unsaved buffers on edit (extends CORRO_AUTO_UNSAVED today).
# autosave = true
```

Deliberately absent for now: keybindings (§2), theme/color overrides
(no theming surface exists in any backend yet — proposing keys without
a consumer invites dead config), per-sheet anything.

## 6. Backend-parity rules

- Shared options (`[document]`, `[startup]`, `[behavior]`) are read
  once in shared startup code; every backend (ratatui, pancurses, GTK,
  nwg, wasm) sees identical values. Backend-specific behavior from a
  shared misread is a bug.
- Backend sections (`[display.gtk]`, `[display.nwg]`, …) map to native
  concepts per backend (GTK font description vs Win32 face name); the
  *preference* (`font_family` list) stays shared, the *spelling* is
  per-backend. A backend ignores sections that don't apply to it.
- Anything that changes rendering needs a signed-off reference render
  per backend before the key ships (same bar as any other display
  change in this repo).

## 7. Diagnostics vs settings

Existing `CORRO_DEBUG_*`, `CORRO_IDLE_MARKER`, `CORRO_KEY_LOG`,
`CORRO_TERM_*`, `CORRO_WASM_OUTPUT` are **test/diagnostic harness**,
not user settings: they stay env-only, undocumented in `--help`, and
are never promoted into the config file. The day a diagnostic needs
persistence, it graduates explicitly with a documented reason — not by
drift. `CORRO_TEMPLATE` and `CORRO_AUTO_UNSAVED` are the two that
already behave as settings; the config file gives them persistent
spellings (`[document].template`, `[behavior].autosave`) while the env
vars keep working per the precedence chain.

## 8. Test plan sketch (for the implementation)

- Precedence unit tests: each layer overridepute (env beats file beats
  defaults; CLI beats all), missing/invalid file falls back silently,
  unknown keys warn without aborting.
- Per-option tests beside the consumer (e.g. template resolution next
  to the existing template loader tests; font plumbing per backend).
- No GUI pixel tests for preferences (combinatorial explosion) — one
  test per backend proving its section *applies* (e.g. family list
  resolves to a native handle), values covered at the shared layer.

## 9. Open questions (need owner answers before implementing)

1. TOML parser vendored vs hand-rolled subset?
2. Should `corro` ever *write* the file (e.g. `corro --save-defaults`)?
   Proposal says no; confirm.
3. Autosave semantics (`CORRO_AUTO_UNSAVED` exists but was built for
   tests): promote as-is or redefine?
4. Per-workbook overrides (`.corro-settings` beside a file)? Out of
   scope here, but it will be asked — record a no (or yes) now.
