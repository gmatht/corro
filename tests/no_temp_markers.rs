//! Meta guard: temporary experiment markers must never reach the tree.
//!
//! A temporary negative-control line left uncommitted in
//! `src/gui/gui_backend.rs` once shipped into a local release build and
//! produced the nwg fx-bar cram (entry fixed-small, labels covered) with
//! zero diagnostics — the committed code was correct the whole time, so
//! no test caught it. This test fails on any such marker anywhere in the
//! shipped sources, on every host (plain file scan, no backend needed).
//!
//! Ordinary TODO/FIXME/XXX/HACK comments are fine and NOT matched — only
//! the experiment prefixes listed in FORBIDDEN below (and explicit
//! do-not-commit banners).
use std::path::{Path, PathBuf};

// Built via concat so this file itself contains no literal marker text.
const FORBIDDEN: &[&str] = &[
    concat!("TEMP-", "NEGCONTROL"),
    concat!("TEMP-", "DIAG"),
    concat!("TEMP-", "EXPERIMENT"),
    concat!("TEMP-", "HACK"),
    concat!("REMOVE-", "BEFORE-COMMIT"),
    concat!("DO-", "NOT-COMMIT"),
];

fn roots() -> Vec<PathBuf> {
    let manifest = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    ["src", "tests"]
        .iter()
        .map(|d| manifest.join(d))
        .chain([
            manifest.join("rustxWidgets/rswidgets/src"),
            manifest.join("rustxWidgets/gtk_dynamic_loader/src"),
        ])
        .collect()
}

fn scan(dir: &Path, hits: &mut Vec<String>) {
    let entries = match std::fs::read_dir(dir) {
        Ok(e) => e,
        Err(_) => return,
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            if path.file_name().map(|n| n == "target").unwrap_or(false) {
                continue;
            }
            scan(&path, hits);
            continue;
        }
        let is_code = path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| matches!(e, "rs" | "toml" | "sh" | "py"))
            .unwrap_or(false);
        if !is_code {
            continue;
        }
        let Ok(content) = std::fs::read_to_string(&path) else {
            continue;
        };
        for (n, line) in content.lines().enumerate() {
            for marker in FORBIDDEN {
                if line.contains(marker) {
                    hits.push(format!("{}:{}: {marker}", path.display(), n + 1));
                }
            }
        }
    }
}

#[test]
fn no_temporary_experiment_markers_in_tree() {
    let mut hits = Vec::new();
    for root in roots() {
        scan(&root, &mut hits);
    }
    assert!(
        hits.is_empty(),
        "temporary experiment markers must not ship (revert or finish them):\n{}",
        hits.join("\n")
    );
}
