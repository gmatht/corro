//! macOS backend core: AppKit layer over the shared Apple runtime.
//!
//! The macOS sibling of [`crate::backends::ios`]. Both delegate the
//! Apple-generic half — Objective-C runtime, `NSString` marshalling, handle
//! lifetime, callback registry, widget registry — to
//! [`crate::backends::apple`]; what lives here is what is genuinely AppKit:
//!
//! * [`display_scale`] — `NSScreen.backingScaleFactor` (1.0 / 2.0). iOS uses
//!   `UIScreen.scale`, so the selector is different even though the idea is
//!   the same.
//! * [`init_with_root`] — the typed `NSView*`/`NSWindowController*` entry
//!   point (a thin wrapper over [`apple::init_with_root_impl`]).
//! * [`MacosApp`] / [`init_backend`] — the `BackendApp`.
//!
//! ## Who owns the run loop
//!
//! On iOS `UIApplicationMain` never returns and the host app is the loop.
//! On macOS a plain `NSApplication` app has *two* valid shapes, and this
//! backend supports both:
//!
//! * **Host-driven** (the iOS/Android shape): a host `NSApplicationDelegate`
//!   calls [`init_with_root`] from `applicationDidFinishLaunching:` and then
//!   `-[NSApp run]`. [`MacosApp::run`] returns immediately, exactly like
//!   `IosApp::run`, because `NSApp` owns the loop; the app is leaked.
//! * **Self-driven**: [`init_backend`] is still available for the
//!   `backends::init()` contract, but note that a macOS GUI app *must* run
//!   `NSApplicationMain`/`-[NSApp run]` on the main thread — there is no
//!   useful "backend owns a blocking loop" variant that also gets AppKit
//!   event delivery without an `NSApplication`. So `run()` stays a no-op and
//!   the host is responsible for `[NSApp run]`, mirroring iOS.
//!
//! See `rustxWidgets/docs/MACOS_GUIDELINES.md` and `rustxWidgets/docs/IOS_GUIDELINES.md`.

#![allow(dead_code)] // several helpers are used only by some widget kinds

// Re-export the shared Apple runtime surface, so a macOS host can reach the
// same names the iOS host does (`backends::macos::msg0`, `WidgetMeta`, ...).
#[cfg(target_os = "macos")]
pub use crate::backends::apple::*;

// `log_macos` mirrors the iOS `log_ios` alias; the shared module calls it
// `log_apple` (NSLog works on both platforms).
#[cfg(target_os = "macos")]
pub use crate::backends::apple::log_apple as log_macos;

use std::error::Error as StdError;

// ------------------------------------------------------------------
// macOS-specific: root view + display scale + backend entry point
// ------------------------------------------------------------------

/// Hand the (AppKit) backend the host's root view and window controller.
///
/// Mirrors Android's `init_with_layout(env, activity, layout)` and iOS's
/// `init_with_root(root, vc)`: after this, `create_*` builds real AppKit
/// objects as subviews of `root` (typically the window's `contentView`).
///
/// # Safety
/// `root` and `vc` must be live Objective-C objects (`NSView*` and
/// `NSWindowController*`) owned by the caller for the process lifetime.
/// Called from the host's root-ready entry point.
#[cfg(target_os = "macos")]
pub unsafe fn init_with_root(
    root: *mut std::os::raw::c_void,
    vc: *mut std::os::raw::c_void,
) -> Result<(), Box<dyn StdError + Send + Sync>> {
    // SAFETY: forwarded unchanged; the shared layer only calls `-retain`.
    unsafe { crate::backends::apple::init_with_root_impl(root, vc, "macos") }
}

/// Display scale (`NSScreen.backingScaleFactor`: 1.0 on a non-Retina display,
/// 2.0 on Retina). The macOS analogue of iOS's `UIScreen.scale` and Android's
/// `display_density`: corro multiplies its pixel metrics by this so chrome is
/// legible on a Retina panel without changing a desktop call site.
///
/// Returns `None` before init or when `NSScreen` is unavailable, so callers
/// fall back to 1.0 and still render.
#[cfg(target_os = "macos")]
pub fn display_scale() -> Option<f64> {
    use crate::backends::apple::{cls, is_initialized, msg0, msg0c};
    if !is_initialized() {
        return None;
    }
    let screen_cls = cls("NSScreen");
    if screen_cls.is_null() {
        return None;
    }
    unsafe {
        // `[NSScreen mainScreen]` can be nil when no display is attached (a
        // headless CI box); fall back to `[NSScreen screens][0]` and finally
        // give up, so the caller uses 1.0.
        let mut screen = msg0(screen_cls, "mainScreen");
        if screen.is_null() {
            let screens = msg0(screen_cls, "screens");
            if screens.is_null() {
                return None;
            }
            screen = msg0(screens, "firstObject");
            if screen.is_null() {
                return None;
            }
        }
        let scale = msg0c(screen, "backingScaleFactor");
        if scale.is_finite() && scale > 0.0 {
            Some(scale)
        } else {
            None
        }
    }
}

/// The macOS `BackendApp`. Like iOS's `IosApp`, `run()` returns immediately:
/// the host app owns `NSApplication` and calls `-[NSApp run]`, which never
/// returns; AppKit then drives every callback on the main thread. There is no
/// loop for rswidgets to own — the app *is* the loop.
#[cfg(target_os = "macos")]
pub struct MacosApp;

#[cfg(target_os = "macos")]
impl MacosApp {
    pub fn new() -> Result<Box<dyn crate::backends::BackendApp>, Box<dyn StdError + Send + Sync>> {
        Ok(Box::new(MacosApp))
    }
}

#[cfg(target_os = "macos")]
impl crate::backends::BackendApp for MacosApp {
    fn run(self: Box<Self>) -> Result<(), Box<dyn StdError + Send + Sync>> {
        Ok(())
    }
}

/// The macOS backend entry point, matching the shape `backends::init()`
/// expects (`Result<Box<dyn BackendApp>, Box<dyn Error + Send + Sync>>`).
#[cfg(target_os = "macos")]
pub fn init_backend(
) -> Result<Box<dyn crate::backends::BackendApp>, Box<dyn StdError + Send + Sync>> {
    Ok(Box::new(MacosApp))
}

// ------------------------------------------------------------------
// Tests
// ------------------------------------------------------------------

#[cfg(test)]
#[cfg(target_os = "macos")]
mod tests {
    use super::*;

    #[test]
    fn test_not_initialized_before_init() {
        assert!(!is_initialized());
        assert!(root_view().is_none());
        assert!(display_scale().is_none());
    }
}
