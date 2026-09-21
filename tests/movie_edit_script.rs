//! `CORRO_EDIT_SCRIPT`: a window can be scripted to *make* edits.
//!
//! The two-window demo (concurrent editing of one file) needs both windows to
//! produce edits through the ordinary commit path, so that each commit is
//! appended to the shared log and the other window's tailer picks it up.
//! Driving that with synthetic keystrokes means calibrating pixel coordinates
//! against the window manager — brittle, and it silently writes to the wrong
//! cell when the calibration drifts (which is exactly what happened while
//! building the demo). Parsing the script is pure, so it is tested here for
//! every front-end, and the TUI's use of it is exercised in `ui` tests.
use corro::ui_core::{edit_script_from_env, EditStep};

/// The parser reads the process environment, so scripts must not overlap.
static ENV_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

fn parse(script: &str) -> Vec<EditStep> {
    let _guard = ENV_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    std::env::set_var("CORRO_EDIT_SCRIPT", script);
    let out = edit_script_from_env();
    std::env::remove_var("CORRO_EDIT_SCRIPT");
    out
}

#[test]
fn parses_cells_values_and_order() {
    let steps = parse("1000:A5=111,5000:B2=hello");
    assert_eq!(steps.len(), 2);
    assert_eq!(steps[0].row, 4); // A5 -> main row 4
    assert_eq!(steps[0].col, 0); // column A
    assert_eq!(steps[0].value, "111");
    assert_eq!(steps[0].at_ms, 1000);
    assert_eq!(steps[1].row, 1); // B2
    assert_eq!(steps[1].col, 1); // column B
    assert_eq!(steps[1].value, "hello");
}

#[test]
fn steps_are_returned_in_time_order() {
    let steps = parse("9000:C1=last,100:A1=first,5000:B1=middle");
    let times: Vec<u64> = steps.iter().map(|s| s.at_ms).collect();
    assert_eq!(times, vec![100, 5000, 9000]);
}

#[test]
fn multi_letter_columns_and_newlines_are_accepted() {
    let steps = parse("10:AA1=x\n20:AB3=y");
    assert_eq!(steps.len(), 2);
    assert_eq!(steps[0].col, 26); // AA
    assert_eq!(steps[1].col, 27); // AB
    assert_eq!(steps[1].row, 2);
}

#[test]
fn malformed_items_are_skipped_not_fatal() {
    // A typo must degrade to fewer edits, never a panic: the demo runs this.
    let steps = parse("nonsense,100:A1=ok,200:=missing,300:B2,400:C3=");
    assert_eq!(steps.len(), 2, "kept: {steps:?}");
    assert_eq!(steps[0].value, "ok");
    assert_eq!(steps[1].value, ""); // an explicit empty value is still an edit
}

#[test]
fn no_env_var_means_no_edits() {
    let _guard = ENV_LOCK.lock().unwrap_or_else(|e| e.into_inner());
    std::env::remove_var("CORRO_EDIT_SCRIPT");
    assert!(edit_script_from_env().is_empty());
}
