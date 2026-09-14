# Recording & Replay Format

## Directory Structure

Each recording directory (e.g. `rec_About/`, `rec_CellEdit/`) contains:

| File | Content |
|------|---------|
| `timing.log` | Delay/byte-count pairs, one per line |
| `keystrokes.log` | `Script started on...\n` + raw binary keystrokes + `Script done on...\n` |
| `command.txt` | (optional) The command to run before replaying keystrokes |

## timing.log — Two Format Variants

### Advanced format (from `script --log-timing`, util-linux >= 2.38)

First character of first line is `H`, `I`, or `O`:
```
H 0.000000 START_TIME 2026-01-01 12:00:00+00:00
H 0.000000 TERM xterm-256color
I 0.500000 3
I 0.300000 5
I 0.300000 1
H 0.000000 DURATION 1.100
H 0.000000 EXIT_CODE 0
```

- `H` = header/metadata (skip)
- `I` = input event: `<delay_seconds> <byte_count>` — read `byte_count` bytes from keystrokes.log
- `O` = output event (skip)

### Classic format (from `gen_menu_recordings.sh`, older `script`)

Just delay/byte-count pairs, no prefix:
```
0.500 3
0.300 5
0.300 1
```

## keystrokes.log

```
Script started on 2026-01-01 12:00:00+00:00\n
<binary bytes: raw terminal escape sequences + typed characters>
Script done on 2026-01-01 12:00:05+00:00\n
```

- **First line**: header — skip this (`find('\n') + 1` = start of data)
- **Body**: raw keystroke bytes, read in chunks per `timing.log`
- **Last line**: trailer — script writes this on exit

## Escape Sequence → Key Mapping

The Python replayer (`split_cmds()` in `.gui_replayer.py`) maps these escape sequences:

| Escape | Key |
|--------|-----|
| `\x1b[A` | Up |
| `\x1b[B` | Down |
| `\x1b[C` | Right |
| `\x1b[D` | Left |
| `\x1b[H` | Home |
| `\x1b[F` | End |
| `\x1b[5~` | Page_Up |
| `\x1b[6~` | Page_Down |
| `\x1b[2~` | Insert |
| `\x1b[3~` | Delete |
| `\x1b[1~` | Home |
| `\x1b[4~` | End |
| `\x1bOP` | F1 |
| `\x1bOQ` | F2 |
| `\x1bOR` | F3 |
| `\x1bOS` | F4 |
| `\x1bOH` | Home |
| `\x1bOF` | End |
| `\x1b<alpha>` | `alt+<letter>` (e.g. `\x1bh` = `alt+h`) |
| `\x1b` alone | Escape |
| printable chars | typed as-is |
| `\n` / `\r` | Return |
| `\t` | Tab |
| `\x7f` | BackSpace |

## Replay Flow

1. Detects format: first char `H`/`I`/`O` → advanced, else classic
2. Skips keystrokes.log header line
3. For each timing line:
   a. Sleep `delay` seconds (capped at 2.0)
   b. Read `byte_count` bytes from keystrokes.log
   c. Parse bytes through `split_cmds()` → list of `(type, arg)` commands
   d. Dispatch via `xdotool` (Linux) or Win32 `PostMessage` (Windows)
4. After all lines: take screenshot, OCR, quit sequence

## xdotool Dispatch

```
Escape      → xdotool key --window <id> Escape
Return/Tab  → xdotool key --window <id> Return / Tab
Up/Down/... → xdotool key --window <id> Up / Down / ...
F1-F4       → xdotool key --window <id> F1 / F2 / F3 / F4
alt+<char>  → xdotool key --window <id> alt+<char>
text        → xdotool type --window <id> <text>
```

## Dialog Detection

After replay, the replayer runs `xdotool search --name` for `About`, `Key Bindings`, etc. If found, prints `DIALOG_FOUND: search="About" wid=...`. The test runner asserts dialogs appeared for `rec_HelpAbout`, `rec_HelpKeybindings`, `rec_HelpFull`, `rec_HelpRowOps`, `rec_HelpColOps`.
