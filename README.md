# corro (WIP)

`corro` is an early-stage Rust TUI spreadsheet experiment built around an **append-only JSONL operation log** and a small **sparse** sheet model.

There is now a working prototype release: [`0.0.1`](https://github.com/gmatht/corro/releases/tag/0.0.1).

## Prototype demo

`corro.cast` is included in this repository as a quick terminal walkthrough.

> Note: GitHub README rendering blocks the script-powered player, but the snippet below is kept for local/docs viewers that allow it.
>
> Fallback: run `asciinema play -s 4 corro.cast` locally.

<link rel="stylesheet" type="text/css" href="https://cdn.jsdelivr.net/npm/asciinema-player@3.8.0/dist/bundle/asciinema-player.css" />
<script src="https://cdn.jsdelivr.net/npm/asciinema-player@3.8.0/dist/bundle/asciinema-player.min.js"></script>
<div id="corro-demo"></div>
<script>
  AsciinemaPlayer.create("corro.cast", document.getElementById("corro-demo"), { speed: 4 });
</script>

## Current limitations

This is still a rough prototype and has important limitations:

- Creating a new file currently depends on shell workflows (for example `touch`) rather than an in-app flow.
- Instead of formulas, the current model uses special columns/rows, which are cleaner and harder to misuse.
- Formula support is still worth adding in the future for broader spreadsheet compatibility and expressiveness.

See `DESIGN.md` for the current architecture and decisions.

