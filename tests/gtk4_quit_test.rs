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

fn find_windows(child: &mut std::process::Child, timeout_secs: u64) -> Vec<String> {
    let deadline = std::time::Instant::now() + std::time::Duration::from_secs(timeout_secs);
    while std::time::Instant::now() < deadline {
        std::thread::sleep(Duration::from_millis(500));
        if let Ok(out) = Command::new("xdotool").args(["search", "--name", "corro"]).output() {
            let ids: Vec<String> = String::from_utf8_lossy(&out.stdout)
                .split_whitespace()
                .filter(|id| {
                    Command::new("xdotool")
                        .args(["getwindowgeometry", "--shell", id])
                        .output()
                        .ok()
                        .map(|g| {
                            let mut w = 0i32;
                            for line in String::from_utf8_lossy(&g.stdout).lines() {
                                if let Some(r) = line.strip_prefix("WIDTH=") {
                                    w = r.trim().parse().unwrap_or(0);
                                }
                            }
                            w > 100
                        })
                        .unwrap_or(false)
                })
                .map(|s| s.to_string())
                .collect();
            if !ids.is_empty() {
                return ids;
            }
        }
    }
    let _ = child.kill();
    panic!("Timed out waiting for window");
}

fn send_alt_f_q(wid: &str) {
    Command::new("xdotool").args(["windowfocus", "--sync", wid]).status().ok();
    std::thread::sleep(Duration::from_millis(200));
    Command::new("xdotool").args(["key", "--window", wid, "alt+f"]).status().ok();
    std::thread::sleep(Duration::from_millis(300));
    Command::new("xdotool").args(["key", "--window", wid, "q"]).status().ok();
    std::thread::sleep(Duration::from_millis(300));
    Command::new("xdotool").args(["keyup", "--window", wid, "alt"]).status().ok();
}

fn test_alt_f_q_quits(backend: &str) {
    let mut child = if backend == "gtk4" {
        Command::new(bin_path())
            .env("BACKEND", "gtk4")
            .arg("--gui")
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .spawn()
            .expect("spawn corro")
    } else {
        Command::new(bin_path())
            .arg("--gui")
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .spawn()
            .expect("spawn corro")
    };

    let windows = find_windows(&mut child, 10);
    eprintln!("Found windows: {:?}", windows);

    // Try the full Alt+F+Q sequence on each window in series.
    // One of them is the real corro window that will receive the keystrokes.
    for wid in &windows {
        for _attempt in 0..3 {
            Command::new("xdotool").args(["windowfocus", "--sync", wid]).status().ok();
            std::thread::sleep(Duration::from_millis(200));
            Command::new("xdotool").args(["key", "--window", wid, "alt+f"]).status().ok();
            std::thread::sleep(Duration::from_millis(300));
            Command::new("xdotool").args(["key", "--window", wid, "q"]).status().ok();
            std::thread::sleep(Duration::from_millis(300));
            Command::new("xdotool").args(["keyup", "--window", wid, "alt"]).status().ok();
            // Check if we exited after each attempt
            if let Ok(Some(_)) = child.try_wait() {
                return;
            }
        }
        if let Ok(Some(_)) = child.try_wait() {
            return;
        }
    }

    // Wait for exit
    let deadline = std::time::Instant::now() + std::time::Duration::from_secs(4);
    loop {
        match child.try_wait() {
            Ok(Some(status)) => {
                eprintln!("Process exited with: {status}");
                assert!(status.success(), "corro exited with error: {status}");
                return; // PASS
            }
            Ok(None) => {
                if std::time::Instant::now() > deadline {
                    let _ = child.kill();
                    let _ = child.wait();
                    panic!("Alt+F+Q did not quit the app (backend={backend})");
                }
                std::thread::sleep(Duration::from_millis(100));
            }
            Err(e) => {
                let _ = child.kill();
                panic!("Error waiting for corro: {e}");
            }
        }
    }
}

#[test]
fn alt_f_q_quits_via_gtk3() {
    test_alt_f_q_quits("gtk3");
}

#[test]
fn alt_f_q_quits_via_gtk4() {
    test_alt_f_q_quits("gtk4");
}
