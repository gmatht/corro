//! Windows-only reproduction for the fx-bar cram: the address label, fx
//! label and formula entry pile up at the left instead of the entry
//! filling the bar.
//!
//! Runs only on Windows (needs the native NWG backend + a message pump;
//! Linux CI skips it via cfg). Build the same widget tree the app builds
//! (window -> vbox -> hbox formula bar [addr, fx, expanding entry]),
//! present it, then poll the native entry rect until it fills the bar or
//! a deadline passes. A persistently narrow entry reproduces the report
//! (fixed-small widgets crammed left, empty bar right); resizing or late
//! layouts heal transient states, so only a stable narrow rect fails.
//!
//! Run on Windows with: cargo test --test fx_bar_nwg_cram
#![cfg(all(windows, not(feature = "pancurses"), not(feature = "zork")))]

use std::time::{Duration, Instant};

fn rect_of(hwnd: *mut std::os::raw::c_void) -> Option<(i32, i32, i32, i32)> {
    if hwnd.is_null() {
        return None;
    }
    unsafe {
        let mut rect: winapi::shared::windef::RECT = std::mem::zeroed();
        if winapi::um::winuser::GetWindowRect(hwnd as _, &mut rect) == 0 {
            return None;
        }
        Some((rect.left, rect.top, rect.right, rect.bottom))
    }
}

#[test]
fn fx_bar_entry_fills_remainder_no_overlap() {
    use rswidgets::prelude::*;
    use std::os::raw::c_void;

    let app = App::init().expect("App::init");
    let win = app.create_window().expect("window");
    win.set_title("fx cram repro");
    let vbox = app
        .create_box(Orientation::Vertical, 2)
        .expect("root vbox");
    let bar = app
        .create_box(Orientation::Horizontal, 2)
        .expect("formula bar");
    let addr = app.create_label("A1").expect("addr label");
    let fx = app.create_label("  fx  ").expect("fx label");
    let entry = app.create_entry().expect("formula entry");
    entry.set_hexpand(true);
    bar.append(&addr);
    bar.append(&fx);
    bar.append(&entry);
    bar.set_child_hexpand(&entry, true);
    vbox.append(&bar);
    win.set_child_box(&vbox);
    win.present();

    let entry_hwnd = *AsRef::<*mut c_void>::as_ref(&entry);
    let bar_hwnd = *AsRef::<*mut c_void>::as_ref(&bar);
    let addr_hwnd = *AsRef::<*mut c_void>::as_ref(&addr);
    assert!(!entry_hwnd.is_null(), "entry must have a native HWND");
    assert!(!bar_hwnd.is_null(), "formula bar must have a native HWND");

    // Settle-poll: a late layout pass heals transient states; only a
    // stably narrow entry (the reported cram) fails the deadline.
    let deadline = Instant::now() + Duration::from_secs(10);
    let (ex0, ex1, bw) = loop {
        nwg::dispatch_thread_events();
        let Some((bl, _bt, br, _bb)) = rect_of(bar_hwnd) else {
            panic!("formula bar has no rect");
        };
        let bar_w = br - bl;
        let entry_rel = rect_of(entry_hwnd).map(|(l, _t, r, _b)| (l - bl, r - bl));
        if let Some((x0, x1)) = entry_rel {
            if x1 - x0 > bar_w / 2 && x0 >= 0 {
                break (x0, x1, bar_w);
            }
        }
        if Instant::now() > deadline {
            let addr_rel =
                rect_of(addr_hwnd).map(|(l, _t, r, _b)| (l - bl, r - bl));
            panic!(
                "fx entry never filled the bar (cram reproduced): bar_w={bar_w} entry={entry_rel:?} addr={addr_rel:?}"
            );
        }
        std::thread::sleep(Duration::from_millis(100));
    };
    // Entry starts right of the labels and reaches near the right edge.
    assert!(
        ex0 >= 0 && ex0 < 200,
        "entry must start right after the labels (x0: {ex0})"
    );
    assert!(
        ex1 >= bw - 20,
        "entry must reach near the bar edge, not cram left (x1: {ex1}, bar: {bw})"
    );
}
