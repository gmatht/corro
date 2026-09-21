#!/usr/bin/env python3
"""List the formula errors a `--movie` replay shows, and where they come from.

A demo video of `docs/tests/*.corro` contains cells displaying `#NAME`,
`#PARSE` and friends — these are real evaluator output for test fixtures that
deliberately contain broken formulas, not rendering artefacts. This walks the
same replay the video records and reports each error cell, the value it holds
and the log line that produced it.

Usage:
    scripts/movie_errors.py docs/tests/main.corro
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent

PROBE = r'''
fn main() {
    let path = std::env::args().nth(1).expect("path");
    let mut movie = corro::gui::movie::GuiMovie::new(std::path::Path::new(&path)).expect("parse");
    let mut app = corro::gui::App::new_with_paths(vec![std::path::PathBuf::from(&path)]);
    let mut active = app.core.workbook.sheet_id(app.core.workbook.active_sheet);
    app.core.view_sheet_id = active;

    // cell label -> (value, display, first log line that produced it)
    let mut seen: std::collections::BTreeMap<String, (String, String, String)> = Default::default();
    for i in 0..movie.len() {
        movie.apply_step(&mut app.core.workbook, &mut active, i).expect("step");
        app.core.workbook.ensure_active_sheet();
        let grid = &app.core.workbook.active_sheet().grid;
        let mc = grid.main_cols();
        for (addr, raw) in grid.iter_nonempty() {
            let disp = corro::formula::cell_effective_display(grid, &addr);
            if disp.starts_with('#') {
                let label = corro::addr::cell_ref_text(&addr, mc);
                seen.entry(label).or_insert_with(|| {
                    (raw.clone(), disp.clone(), movie.steps[i].line.clone())
                });
            }
        }
    }
    for (label, (raw, disp, line)) in &seen {
        println!("{label}\t{disp}\t{raw}\t{line}");
    }
}
'''


def main() -> int:
    if len(sys.argv) != 2:
        sys.exit(__doc__.strip().splitlines()[-1])
    target = Path(sys.argv[1])
    if not target.exists():
        sys.exit(f"error: {target} not found")

    probe_dir = REPO_ROOT / "examples"
    probe_dir.mkdir(exist_ok=True)
    probe = probe_dir / "movie_errors_probe.rs"
    probe.write_text(PROBE)
    try:
        out = subprocess.run(
            ["cargo", "run", "-q", "--features", "gui", "--example", "movie_errors_probe",
             "--", str(target)],
            cwd=str(REPO_ROOT),
            capture_output=True,
            text=True,
        )
    finally:
        probe.unlink(missing_ok=True)
    if out.returncode != 0:
        sys.exit(f"error: probe failed\n{out.stderr.strip()}")

    rows = [l.split("\t") for l in out.stdout.splitlines() if "\t" in l and not l.startswith("DEBUG")]
    if not rows:
        print(f"{target}: no formula errors are shown by the replay")
        return 0
    print(f"{target}: {len(rows)} cell(s) show a formula error")
    width = max(len(r[0]) for r in rows)
    for label, disp, raw, line in rows:
        print(f"  {label:<{width}}  {disp:<8} holds {raw:<12} from: {line}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
