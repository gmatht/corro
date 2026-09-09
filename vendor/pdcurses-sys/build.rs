extern crate cc;

// Win95/9x compatibility (vendored patch): the win32a GDI flavor draws via
// ExtTextOutW/TextOutW, which are no-op stubs on Windows 95 — a win32a build
// shows a blank window there. The classic win32 (console) flavor uses only
// ANSI console APIs (WriteConsoleOutputA, ReadConsoleInputA, ...), so for
// rust9x targets we always compile the win32 flavor and drop PDC_WIDE (with
// PDC_WIDE undefined, the W-gated calls in the win32 flavor disappear).
// The 64-bit chtype ABI pdcurses-sys expects comes from CHTYPE_LONG, which
// independently defaults to 2 (64-bit) — unchanged.
fn is_rust9x_target() -> bool {
    // CARGO_CFG_TARGET_FAMILY is the comma-separated cfg(target_family)
    // list (e.g. "rust9x,windows" for the rust9x msvc targets).
    std::env::var("CARGO_CFG_TARGET_FAMILY")
        .map(|v| v.split(|c| c == ',' || c == ' ').any(|f| f == "rust9x"))
        .unwrap_or(false)
}

fn main() {
    // Non-Windows targets never link pdcurses (the pancurses crate uses
    // ncurses there; nothing in this workspace references pdcurses_sys
    // outside `cfg(windows)`). Emit no artifacts so `--features pancurses`
    // stays buildable on Linux, where the tmux render-parity tests run.
    // (The `links = "pdcurses"` key still reserves the native lib name.)
    if std::env::var("CARGO_CFG_TARGET_OS").map(|os| os != "windows").unwrap_or(false) {
        return;
    }
    let mut build = cc::Build::new();
    build
        .file("src/PDCurses/pdcurses/addch.c") //Common PDCurses files
        .file("src/PDCurses/pdcurses/addchstr.c")
        .file("src/PDCurses/pdcurses/addstr.c")
        .file("src/PDCurses/pdcurses/attr.c")
        .file("src/PDCurses/pdcurses/beep.c")
        .file("src/PDCurses/pdcurses/bkgd.c")
        .file("src/PDCurses/pdcurses/border.c")
        .file("src/PDCurses/pdcurses/clear.c")
        .file("src/PDCurses/pdcurses/color.c")
        .file("src/PDCurses/pdcurses/debug.c")
        .file("src/PDCurses/pdcurses/delch.c")
        .file("src/PDCurses/pdcurses/deleteln.c")
        .file("src/PDCurses/pdcurses/deprec.c")
        .file("src/PDCurses/pdcurses/getch.c")
        .file("src/PDCurses/pdcurses/getstr.c")
        .file("src/PDCurses/pdcurses/getyx.c")
        .file("src/PDCurses/pdcurses/inch.c")
        .file("src/PDCurses/pdcurses/inchstr.c")
        .file("src/PDCurses/pdcurses/initscr.c")
        .file("src/PDCurses/pdcurses/inopts.c")
        .file("src/PDCurses/pdcurses/insch.c")
        .file("src/PDCurses/pdcurses/insstr.c")
        .file("src/PDCurses/pdcurses/instr.c")
        .file("src/PDCurses/pdcurses/kernel.c")
        .file("src/PDCurses/pdcurses/keyname.c")
        .file("src/PDCurses/pdcurses/mouse.c")
        .file("src/PDCurses/pdcurses/move.c")
        .file("src/PDCurses/pdcurses/outopts.c")
        .file("src/PDCurses/pdcurses/overlay.c")
        .file("src/PDCurses/pdcurses/pad.c")
        .file("src/PDCurses/pdcurses/panel.c")
        .file("src/PDCurses/pdcurses/printw.c")
        .file("src/PDCurses/pdcurses/refresh.c")
        .file("src/PDCurses/pdcurses/scanw.c")
        .file("src/PDCurses/pdcurses/scr_dump.c")
        .file("src/PDCurses/pdcurses/scroll.c")
        .file("src/PDCurses/pdcurses/slk.c")
        .file("src/PDCurses/pdcurses/termattr.c")
        .file("src/PDCurses/pdcurses/terminfo.c")
        .file("src/PDCurses/pdcurses/touch.c")
        .file("src/PDCurses/pdcurses/util.c")
        .file("src/PDCurses/pdcurses/window.c")
        .include("src/PDCurses")
        .define("PDC_FORCE_UTF8", Some("Y")) // Makes PDCurses ignore the system locale, and treat all narrow-character strings as UTF-8
        .define("PDC_RGB", Some("Y")); // Use RGB colors, it's what most people expect them to be

    if is_rust9x_target() {
        // No PDC_WIDE on rust9x: PDCurses' W-function usage must be compiled
        // out for Windows 95 (see comment above). pancurses on Windows only
        // uses the narrow API (addstr/getch/...), so this is transparent.
        // Also shim the UCRT-only CRT symbols zig's mingw headers reference
        // (the vintage link uses the VC6 CRT).
        build.file("src/win95_vc6_crt_shim.c");
        // Use the REAL console (CONOUT$/CONIN$) instead of the (possibly
        // redirected) standard handles.
        build.define("_WIN32_W9X_CONSOLE", Some("1"));
        // TEMPORARY diagnostic trace of the wgetch input loop to c:\pdc.trace.
        build.define("PDC_TRACE", Some("1"));
        build.define("PDC_TRACE_REV", Some("A"));
    } else {
        build.define("PDC_WIDE", Some("Y")); // Build with wide-character (Unicode) support
    }

    flavor_specifics(&mut build);

    build.compile("libpdcurses.a");
}

// Use win32a if it's chosen, or if no flavor is chosen.
// For rust9x targets, always use the win32 (console) flavor instead —
// see is_rust9x_target() above.
#[cfg(any(feature = "win32a", all(not(feature="win32"), not(feature="win32a"))))]
fn flavor_specifics(build: &mut cc::Build) {
    if is_rust9x_target() {
        return win32_flavor(build);
    }
    win32a_flavor(build);
}

#[cfg(any(feature = "win32a", all(not(feature="win32"), not(feature="win32a"))))]
fn win32a_flavor(build: &mut cc::Build) {
    println!("cargo:rustc-link-lib=dylib=user32");
    println!("cargo:rustc-link-lib=dylib=gdi32");
    println!("cargo:rustc-link-lib=dylib=comdlg32");
    println!("cargo:rustc-link-lib=dylib=shell32");

    build
        .file("src/PDCurses/win32a/pdcclip.c")
        .file("src/PDCurses/win32a/pdcdisp.c")
        .file("src/PDCurses/win32a/pdcgetsc.c")
        .file("src/PDCurses/win32a/pdckbd.c")
        .file("src/PDCurses/win32a/pdcscrn.c")
        .file("src/PDCurses/win32a/pdcsetsc.c")
        .file("src/PDCurses/win32a/pdcutil.c");
}

#[cfg(feature = "win32")]
fn flavor_specifics(build: &mut cc::Build) {
    win32_flavor(build);
}

#[cfg(any(feature = "win32", any(feature = "win32a", all(not(feature="win32"), not(feature="win32a")))))]
fn win32_flavor(build: &mut cc::Build) {
    println!("cargo:rustc-link-lib=dylib=user32");

    build
        .file("src/PDCurses/win32/pdcclip.c")
        .file("src/PDCurses/win32/pdcdisp.c")
        .file("src/PDCurses/win32/pdcgetsc.c")
        .file("src/PDCurses/win32/pdckbd.c")
        .file("src/PDCurses/win32/pdcscrn.c")
        .file("src/PDCurses/win32/pdcsetsc.c")
        .file("src/PDCurses/win32/pdcutil.c");
}