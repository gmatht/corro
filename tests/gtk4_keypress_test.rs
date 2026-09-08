#![cfg(feature = "gtk4")]

/// Test: Alt+F should open the File menu popover, changing the window pixels.
#[test]
fn alt_f_opens_file_menu() {
    use std::process::{Command, Stdio};
    use std::time::Duration;

    let bin = std::env::current_dir()
        .unwrap_or_default()
        .join("target")
        .join("debug")
        .join("corro");

    let mut child = Command::new(&bin)
        .env("BACKEND", "gtk4")
        .arg("--gui")
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .spawn()
        .expect("spawn corro");

    // Wait for a corro window > 100px
    let all_windows = wait_for_windows(10);
    if all_windows.is_empty() {
        let _ = child.kill();
        panic!("Timed out waiting for windows");
    }

    // Focus all windows
    for wid in &all_windows {
        Command::new("xdotool").args(["windowfocus", "--sync", wid]).status().ok();
    }
    std::thread::sleep(Duration::from_millis(500));

    // Take before screenshots
    for wid in &all_windows {
        let b = format!("/tmp/gtk4_b_{wid}.xwd");
        Command::new("xwd").args(["-id", wid, "-out", &b]).status().ok();
    }

    // Send Alt+F to ALL windows (one of them is the real corro window)
    for wid in &all_windows {
        Command::new("xdotool").args(["windowfocus", "--sync", wid]).status().ok();
    }
    std::thread::sleep(Duration::from_millis(200));
    for wid in &all_windows {
        Command::new("xdotool").args(["key", "--window", wid, "alt+f"]).status().ok();
    }
    std::thread::sleep(Duration::from_millis(500));
    for wid in &all_windows {
        Command::new("xdotool").args(["keyup", "--window", wid, "alt"]).status().ok();
    }

    // Take after screenshots and compare each for changes
    for wid in &all_windows {
        let a = format!("/tmp/gtk4_a_{wid}.xwd");
        Command::new("xwd").args(["-id", wid, "-out", &a]).status().ok();
        let b = format!("/tmp/gtk4_b_{wid}.xwd");

        let bpng = format!("/tmp/gtk4_b_{wid}.png");
        let apng = format!("/tmp/gtk4_a_{wid}.png");
        Command::new("convert").args([&b, &bpng]).status().ok();
        Command::new("convert").args([&a, &apng]).status().ok();

        // Check for pixel difference
        let script = format!("/tmp/gtk4_diff_{wid}.py");
        let code = format!(
            "import sys\nfrom PIL import Image\nb4 = Image.open('{}').convert('RGBA')\naf = Image.open('{}').convert('RGBA')\n\
             pix_b4 = list(b4.getdata())\npix_af = list(af.getdata())\ndiffs = sum(1 for i in range(len(pix_b4)) if pix_b4[i] != pix_af[i])\n\
             total = len(pix_b4)\npct = diffs / total * 100\nprint('Differing pixels: %d / %d (%.1f%%)' % (diffs, total, pct))\n\
             if diffs < 200:\n    print('FAIL window {wid}: no change')\n    sys.exit(0)\nelse:\n    print('PASS window {wid}: Alt+F caused visible change')\n    sys.exit(2)\n",
            bpng, apng
        );
        std::fs::write(&script, &code).ok();
        let py_out = Command::new("python3").args([&script]).output().expect("python3");
        eprintln!("Window {wid}: {}", String::from_utf8_lossy(&py_out.stdout));
        if py_out.status.code() == Some(2) {
            // Found the right window!
            let _ = child.kill();
            let _ = child.wait();
            return; // PASS
        }
    }

    // None of the windows showed the menu
    let _ = child.kill();
    let _ = child.wait();
    panic!("Alt+F did not cause visible change on any window");
}

fn wait_for_windows(timeout_secs: u64) -> Vec<String> {
    let deadline = std::time::Instant::now() + std::time::Duration::from_secs(timeout_secs);
    while std::time::Instant::now() < deadline {
        std::thread::sleep(std::time::Duration::from_millis(500));
        if let Ok(out) = std::process::Command::new("xdotool")
            .args(["search", "--name", "corro"]).output()
        {
            let txt = String::from_utf8_lossy(&out.stdout);
            let windows: Vec<String> = txt.split_whitespace()
                .map(|s| s.to_string())
                .filter(|id| {
                    if let Ok(geo) = std::process::Command::new("xdotool")
                        .args(["getwindowgeometry", "--shell", id]).output()
                    {
                        let gtxt = String::from_utf8_lossy(&geo.stdout);
                        let mut w = 0i32;
                        for line in gtxt.lines() {
                            if let Some(r) = line.strip_prefix("WIDTH=") {
                                w = r.trim().parse().unwrap_or(0);
                            }
                        }
                        w > 100
                    } else { false }
                })
                .collect();
            if !windows.is_empty() {
                return windows;
            }
        }
    }
    Vec::new()
}
