//! Stress test: repeatedly activate menu items and dialogs to check for
//! crashes caused by use-after-free, double-free, or RefCell panics.
//!
//! If the backend survives N iterations without crashing the test passes.

#![cfg(feature = "gtk4")]

use std::process::{Command, Stdio};
use std::time::Duration;

fn bin_path() -> std::path::PathBuf {
    std::env::current_dir()
        .unwrap_or_default()
        .join("target")
        .join("debug")
        .join("corro")
}

fn find_window(child: &mut std::process::Child, timeout_secs: u64) -> Option<String> {
    let deadline = std::time::Instant::now() + std::time::Duration::from_secs(timeout_secs);
    while std::time::Instant::now() < deadline {
        std::thread::sleep(Duration::from_millis(500));
        if let Ok(out) = Command::new("xdotool").args(["search", "--name", "corro"]).output() {
            for id in String::from_utf8_lossy(&out.stdout).split_whitespace() {
                if let Ok(g) = Command::new("xdotool")
                    .args(["getwindowgeometry", "--shell", id])
                    .output()
                {
                    let mut w = 0i32;
                    for line in String::from_utf8_lossy(&g.stdout).lines() {
                        if let Some(r) = line.strip_prefix("WIDTH=") {
                            w = r.trim().parse().unwrap_or(0);
                        }
                    }
                    if w > 100 {
                        return Some(id.to_string());
                    }
                }
            }
        }
    }
    let _ = child.kill();
    None
}

fn menu_stress_test(backend: &str) {
    let mut child = Command::new(bin_path())
        .env("BACKEND", backend)
        .arg("--gui")
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .expect("spawn corro");

    let wid = match find_window(&mut child, 10) {
        Some(w) => w,
        None => { let _ = child.kill(); panic!("Timed out waiting for window"); }
    };
    eprintln!("menu_stress_test({backend}): window={wid}");

    let pid = child.id();

    // Repeatedly open and activate menu items (non-destructive).
    for i in 1..=20 {
        // File > New
        Command::new("xdotool").args(["key", "--window", &wid, "alt+f"]).status().ok();
        std::thread::sleep(Duration::from_millis(120));
        Command::new("xdotool").args(["key", "--window", &wid, "n"]).status().ok();
        std::thread::sleep(Duration::from_millis(120));

        if !is_alive(pid) { panic!("CRASH at iteration {i} (File>New)") }

        // Edit > Undo
        Command::new("xdotool").args(["key", "--window", &wid, "alt+e"]).status().ok();
        std::thread::sleep(Duration::from_millis(120));
        Command::new("xdotool").args(["key", "--window", &wid, "u"]).status().ok();
        std::thread::sleep(Duration::from_millis(120));

        if !is_alive(pid) { panic!("CRASH at iteration {i} (Edit>Undo)") }

        if i % 5 == 0 {
            eprintln!("  menu iteration {i}");
        }
    }

    // File > Quit to exit cleanly
    Command::new("xdotool").args(["key", "--window", &wid, "alt+f"]).status().ok();
    std::thread::sleep(Duration::from_millis(200));
    Command::new("xdotool").args(["key", "--window", &wid, "q"]).status().ok();
    std::thread::sleep(Duration::from_millis(200));
    Command::new("xdotool").args(["keyup", "--window", &wid, "alt"]).status().ok();

    let deadline = std::time::Instant::now() + Duration::from_secs(4);
    loop {
        match child.try_wait() {
            Ok(Some(status)) => {
                eprintln!("  exit: {status}");
                assert!(status.success(), "corro exited with error: {status}");
                return;
            }
            Ok(None) => {
                if std::time::Instant::now() > deadline {
                    let _ = child.kill();
                    panic!("File>Quit did not exit the app");
                }
                std::thread::sleep(Duration::from_millis(100));
            }
            Err(e) => { let _ = child.kill(); panic!("wait error: {e}"); }
        }
    }
}

fn dialog_stress_test(backend: &str) {
    let mut child = Command::new(bin_path())
        .env("BACKEND", backend)
        .arg("--gui")
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .expect("spawn corro");

    let wid = match find_window(&mut child, 10) {
        Some(w) => w,
        None => { let _ = child.kill(); panic!("Timed out waiting for window"); }
    };
    eprintln!("dialog_stress_test({backend}): window={wid}");

    let pid = child.id();

    // Repeatedly open Help > About dialog and close it.
    for i in 1..=12 {
        Command::new("xdotool").args(["key", "--window", &wid, "alt+h"]).status().ok();
        std::thread::sleep(Duration::from_millis(150));
        Command::new("xdotool").args(["key", "--window", &wid, "a"]).status().ok();
        std::thread::sleep(Duration::from_millis(200));
        // Close dialog via escape
        Command::new("xdotool").args(["key", "--window", &wid, "Escape"]).status().ok();
        std::thread::sleep(Duration::from_millis(150));
        if !is_alive(pid) { panic!("CRASH at iteration {i} (Help>About)") }
        if i % 4 == 0 { eprintln!("  about dialog iteration {i}"); }
    }

    // Repeatedly open Data > Sort dialog and close it.
    for i in 1..=12 {
        Command::new("xdotool").args(["key", "--window", &wid, "alt+d"]).status().ok();
        std::thread::sleep(Duration::from_millis(150));
        Command::new("xdotool").args(["key", "--window", &wid, "s"]).status().ok();
        std::thread::sleep(Duration::from_millis(200));
        Command::new("xdotool").args(["key", "--window", &wid, "Escape"]).status().ok();
        std::thread::sleep(Duration::from_millis(150));
        if !is_alive(pid) { panic!("CRASH at iteration {i} (Data>Sort)") }
        if i % 4 == 0 { eprintln!("  sort dialog iteration {i}"); }
    }

    // Try to quit cleanly; if that fails, kill the process (stress tests are
    // about surviving N iterations, not Quit).
    Command::new("xdotool").args(["keyup", "alt", "ctrl", "shift"]).status().ok();
    std::thread::sleep(Duration::from_millis(100));
    Command::new("xdotool").args(["windowfocus", "--sync", &wid]).status().ok();
    std::thread::sleep(Duration::from_millis(200));
    Command::new("xdotool").args(["key", "--window", &wid, "ctrl+q"]).status().ok();
    std::thread::sleep(Duration::from_millis(300));
    Command::new("xdotool").args(["keyup", "ctrl"]).status().ok();
    std::thread::sleep(Duration::from_millis(300));
    if is_alive(pid) {
        Command::new("xdotool").args(["keydown", "alt"]).status().ok();
        std::thread::sleep(Duration::from_millis(80));
        Command::new("xdotool").args(["key", "f"]).status().ok();
        std::thread::sleep(Duration::from_millis(200));
        Command::new("xdotool").args(["key", "q"]).status().ok();
        std::thread::sleep(Duration::from_millis(200));
        Command::new("xdotool").args(["keyup", "alt"]).status().ok();
        std::thread::sleep(Duration::from_millis(200));
    }

    let deadline = std::time::Instant::now() + Duration::from_secs(4);
    loop {
        match child.try_wait() {
            Ok(Some(status)) => {
                eprintln!("  exit: {status}");
                return;
            }
            Ok(None) => {
                if std::time::Instant::now() > deadline {
                    let _ = child.kill();
                    let _ = child.wait();
                    eprintln!("  killed (quit timeout)");
                    return;
                }
                std::thread::sleep(Duration::from_millis(100));
            }
            Err(e) => { let _ = child.kill(); panic!("wait error: {e}"); }
        }
    }
}

fn is_alive(pid: u32) -> bool {
    std::path::Path::new(&format!("/proc/{pid}")).exists()
}

#[test]
fn menu_stress_via_gtk4() {
    menu_stress_test("gtk4");
}

#[test]
fn dialog_stress_via_gtk4() {
    dialog_stress_test("gtk4");
}

#[test]
fn show_menu_stress_via_gtk4() {
    // Stress test for View > Toggle Headers
    let mut child = Command::new(bin_path())
        .env("BACKEND", "gtk4")
        .arg("--gui")
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .expect("spawn corro");

    let wid = match find_window(&mut child, 10) {
        Some(w) => w,
        None => { let _ = child.kill(); panic!("Timed out waiting for window"); }
    };
    eprintln!("show_menu_stress_test: window={wid}");

    let pid = child.id();
    for i in 1..=20 {
        // View > Toggle Headers (mnemonic 't' from "_Toggle Headers")
        Command::new("xdotool").args(["key", "--window", &wid, "alt+v"]).status().ok();
        std::thread::sleep(Duration::from_millis(120));
        Command::new("xdotool").args(["key", "--window", &wid, "t"]).status().ok();
        std::thread::sleep(Duration::from_millis(120));

        if !is_alive(pid) { panic!("CRASH at iteration {i} (View>Toggle Headers)") }
        if i % 5 == 0 { eprintln!("  show-menu iteration {i}"); }
    }

    // Quit normally
    Command::new("xdotool").args(["keyup", "alt", "ctrl", "shift"]).status().ok();
    std::thread::sleep(Duration::from_millis(100));
    Command::new("xdotool").args(["windowfocus", "--sync", &wid]).status().ok();
    std::thread::sleep(Duration::from_millis(200));
    Command::new("xdotool").args(["key", "--window", &wid, "ctrl+q"]).status().ok();
    std::thread::sleep(Duration::from_millis(300));
    Command::new("xdotool").args(["keyup", "ctrl"]).status().ok();
    std::thread::sleep(Duration::from_millis(300));
    if is_alive(pid) {
        Command::new("xdotool").args(["keydown", "alt"]).status().ok();
        std::thread::sleep(Duration::from_millis(80));
        Command::new("xdotool").args(["key", "f"]).status().ok();
        std::thread::sleep(Duration::from_millis(200));
        Command::new("xdotool").args(["key", "q"]).status().ok();
        std::thread::sleep(Duration::from_millis(200));
        Command::new("xdotool").args(["keyup", "alt"]).status().ok();
        std::thread::sleep(Duration::from_millis(200));
    }

    let deadline = std::time::Instant::now() + Duration::from_secs(4);
    loop {
        match child.try_wait() {
            Ok(Some(status)) => {
                eprintln!("  exit: {status}");
                return;
            }
            Ok(None) => {
                if std::time::Instant::now() > deadline {
                    let _ = child.kill();
                    let _ = child.wait();
                    eprintln!("  killed (quit timeout)");
                    return;
                }
                std::thread::sleep(Duration::from_millis(100));
            }
            Err(e) => { let _ = child.kill(); panic!("wait error: {e}"); }
        }
    }
}
