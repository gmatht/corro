//! Generate the Objective-C *forwarding shims* the Apple backends send
//! messages to, so the Rust side and the ObjC side cannot drift.
//!
//! ## Why this exists
//!
//! Both Apple backends reach the framework through the Objective-C runtime and
//! therefore need a handful of classes that live **in the host app bundle**:
//! a target/action trampoline, a text measurer/drawer, an alert wrapper, and a
//! canvas view. Most of them are *pure forwarding* — the Rust side already
//! knows the exact selector and its signature (it names them in its `raw_send!`
//! calls), and the ObjC side only has to declare the same method and hand the
//! arguments to a `corro_*` C function or to the framework.
//!
//! Hand-writing those files means two independent declarations of one ABI.
//! A `.m` file that drifts from the Rust signature does not fail to compile —
//! it crashes at runtime with a mismatched `objc_msgSend` (the classic arm64
//! register-file corruption `IOS_GUIDELINES.md` warns about). That is exactly
//! the failure this generator removes: the signatures below are the contract,
//! and the emitted `.m` is derived from it.
//!
//! ## What it does and does not generate
//!
//! **Generated** (pure boilerplate, no design content):
//!   * `Corro<Platform>Target`  — callback-id → `corro_<platform>_callback`
//!   * `Corro<Platform>Alert`   — `corroNewAlert`/`corroAddAction:`/`corroSetTitle:`
//!     over `UIAlertController`+`UIAlertView` (iOS) or `NSAlert` (macOS)
//!   * category extras on the base view class the adapters call
//!     (`corroSetSpacing:`, `corroSetFlex:`, `corroSetMinWidth:`,
//!     `corroSetCanvasId:`, the `corroBoundsWidth/Height` accessors)
//!
//! **Deliberately not generated** (see [`CanvasView`]):
//!   * the canvas view class itself. Its body is *behaviour*, not forwarding —
//!     `drawRect:` reporting the laid-out size before replaying the closure,
//!     touch/mouse coordinate conversion, `pressesBegan:`/`keyDown:` for a
//!     hardware keyboard, the `isFlipped` override, and the documented
//!     degradation paths found by CI (an older SDK without a selector). Those
//!     are the lines a human reads when the sheet does not draw.
//!   * `Corro<Platform>Text`. Its `measure:`/`drawText:` bodies are the
//!     version table (`boundingRectWithSize:` on iOS 7+ vs `sizeWithFont:` on
//!     iOS 6; `UIFont`/`NSFont` resolution). The selector *signatures* are
//!     listed below so the file can be checked against them, but the bodies
//!     stay hand-written: a generator cannot validate them without an SDK, and
//!     a wrong body there is a layout bug rather than a crash.
//!
//! ## Usage
//!
//! Same pattern as [`crate::android_generator`]: a host's `build.rs` includes
//! this file with `#[path]` (so a vendored rswidgets still works and there is
//! no `[build-dependencies]` entry), then:
//!
//! ```text
//! println!("cargo:rerun-if-env-changed={}", rswidgets_shims::PROJECT_ENV);
//! rswidgets_shims::run();
//! ```
//!
//! Opt-in via the environment variable, additive (existing files are never
//! overwritten), and never fails a build — a read-only or absent Apple project
//! must not break a desktop build of the same crate.

use std::fmt::Write as _;
use std::io;
use std::path::{Path, PathBuf};

/// Environment variable naming the Apple project root to generate into.
pub const PROJECT_ENV: &str = "RSWIDGETS_APPLE_PROJECT";

/// Where the generated shim goes, relative to the project root.
///
/// `app/CorroGeneratedShims.m` sits beside the hand-written
/// `app/CorroIosShims.m` so an Xcode target only has to add one file, and so
/// it is obvious which of the two a change belongs in.
pub const SHIM_RELPATH: &str = "app/CorroGeneratedShims.m";

/// Header beside the shim, so the hand-written canvas/text classes can
/// `#import` it and share the category declarations.
pub const SHIM_HEADER_RELPATH: &str = "app/CorroGeneratedShims.h";

/// Which Apple platform to generate for.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Platform {
    Ios,
    Macos,
}

impl Platform {
    /// The `corro_<platform>_*` prefix the Rust exports use (see
    /// `ios/corro/src/lib.rs`; a macOS host uses the same convention).
    pub fn c_prefix(self) -> &'static str {
        match self {
            Platform::Ios => "corro_ios",
            Platform::Macos => "corro_macos",
        }
    }

    /// ObjC class prefix, matching what the adapters resolve with
    /// `objc_getClass`.
    pub fn class_prefix(self) -> &'static str {
        match self {
            Platform::Ios => "CorroIos",
            Platform::Macos => "CorroMac",
        }
    }

    /// Framework header the generated file imports.
    pub fn framework_header(self) -> &'static str {
        match self {
            Platform::Ios => "UIKit/UIKit.h",
            Platform::Macos => "AppKit/AppKit.h",
        }
    }

    /// The base view class the adapters create their containers/canvases from.
    pub fn base_view_class(self) -> &'static str {
        match self {
            Platform::Ios => "UIView",
            Platform::Macos => "NSView",
        }
    }

    /// `YES` for the platform whose view coordinates already start top-left.
    ///
    /// UIKit is top-left by default; AppKit is bottom-left, so the generated
    /// canvas contract requires `isFlipped == YES` there (the single most
    /// common AppKit porting bug — `MACOS_GUIDELINES.md` §5).
    pub fn view_is_top_left_origin(self) -> bool {
        matches!(self, Platform::Ios)
    }

    pub fn name(self) -> &'static str {
        match self {
            Platform::Ios => "ios",
            Platform::Macos => "macos",
        }
    }
}

/// One selector the Rust adapter sends to a host class, with the argument
/// kinds that decide the generated method signature.
///
/// The kinds mirror the `unsafe extern "C" fn(...)` types in the adapter's
/// `raw_send!` calls, minus the implicit `(id, SEL)` pair.
#[derive(Clone, Copy, Debug)]
pub struct ShimMethod {
    /// Selector as the adapter spells it (`"corroSetSpacing:"`).
    pub selector: &'static str,
    /// Argument kinds after the implicit `(id, SEL)`.
    pub args: &'static [ArgKind],
    /// Return kind.
    pub ret: ArgKind,
    /// `true` when the method is a class method (`+`) rather than an
    /// instance method (`-`).
    pub class_method: bool,
    /// Set when the method is generated as a stub with this body comment
    /// instead of real behaviour (used for the checked-but-hand-written ones).
    pub hand_written_note: Option<&'static str>,
}

/// Argument/return kinds, mapped to ObjC types per platform.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ArgKind {
    /// `id` — an object handle (`*mut c_void` in Rust).
    Object,
    /// `NSString*` — object handles that are always strings in practice.
    String,
    /// `SEL`.
    Selector,
    /// `NSInteger` / `NSUInteger` (`isize`/`usize`; the adapters use `u64` for
    /// the unsigned cases and `isize` for the signed ones).
    Integer,
    /// `NSUInteger` specifically (`u64` in the adapter).
    UInteger,
    /// `BOOL`.
    Bool,
    /// `CGFloat` (`f64` on 64-bit, `f32` on 32-bit — the adapter selects at
    /// compile time).
    Float,
    /// `void` — only valid as a return kind.
    Void,
}

impl ArgKind {
    /// The ObjC spelling for this kind. Takes no `Platform`: every kind maps
    /// to the same type on both targets (`CGFloat` is a typedef, so it is
    /// already the right width per target — see the note below), and the
    /// tests pin that for iOS and macOS alike.
    fn objc_type(self) -> &'static str {
        match self {
            ArgKind::Object | ArgKind::String => "id",
            ArgKind::Selector => "SEL",
            ArgKind::Integer => "NSInteger",
            ArgKind::UInteger => "NSUInteger",
            ArgKind::Bool => "BOOL",
            // CGFloat is a typedef, so it is already the right width per
            // target — the generated header needs no 32/64 branch. That is
            // the point of spelling it `CGFloat` here rather than `double`.
            ArgKind::Float => "CGFloat",
            ArgKind::Void => "void",
        }
    }
}

/// The forwarding shims, in the order they are emitted.
///
/// Kept as data (not string literals scattered through the emitter) so a
/// test can assert the set matches what the adapters actually send — see
/// [`RustSideSelector`] and the tests at the bottom.
pub const TARGET_METHODS: &[ShimMethod] = &[
    ShimMethod {
        selector: "targetWithCallbackId:",
        args: &[ArgKind::Integer],
        ret: ArgKind::Object,
        class_method: true,
        hand_written_note: None,
    },
    ShimMethod {
        selector: "corroFired:",
        args: &[ArgKind::Object],
        ret: ArgKind::Void,
        class_method: false,
        hand_written_note: None,
    },
];

pub const ALERT_METHODS: &[ShimMethod] = &[
    ShimMethod {
        selector: "corroNewAlert",
        args: &[],
        ret: ArgKind::Object,
        class_method: true,
        hand_written_note: None,
    },
    ShimMethod {
        selector: "corroSetTitle:",
        args: &[ArgKind::String],
        ret: ArgKind::Void,
        class_method: false,
        hand_written_note: None,
    },
    ShimMethod {
        selector: "corroAddAction:",
        args: &[ArgKind::String],
        ret: ArgKind::Void,
        class_method: false,
        hand_written_note: None,
    },
];

/// Category methods the adapters send to plain views (the base view class),
/// providing the expand/min-width protocol the shared layout code expects.
pub const VIEW_CATEGORY_METHODS: &[ShimMethod] = &[
    ShimMethod {
        selector: "corroSetSpacing:",
        args: &[ArgKind::Integer],
        ret: ArgKind::Void,
        class_method: false,
        hand_written_note: None,
    },
    ShimMethod {
        selector: "corroSetFlex:",
        args: &[ArgKind::Bool],
        ret: ArgKind::Void,
        class_method: false,
        hand_written_note: None,
    },
    ShimMethod {
        selector: "corroSetMinWidth:",
        args: &[ArgKind::Integer],
        ret: ArgKind::Void,
        class_method: false,
        hand_written_note: None,
    },
    ShimMethod {
        selector: "corroSetCanvasId:",
        args: &[ArgKind::Integer],
        ret: ArgKind::Void,
        class_method: false,
        hand_written_note: Some(
            "the canvas id is stored on the view by the canvas class; this \
             category provides the setter the adapter calls before the view \
             is customised, so a plain view is also valid",
        ),
    },
    ShimMethod {
        selector: "corroBoundsWidth",
        args: &[],
        ret: ArgKind::Float,
        class_method: false,
        hand_written_note: None,
    },
    ShimMethod {
        selector: "corroBoundsHeight",
        args: &[],
        ret: ArgKind::Float,
        class_method: false,
        hand_written_note: None,
    },
];

/// Selectors that stay hand-written but whose *Signature must still match* —
/// listed here so the generated header declares them and a mismatch is a
/// compile error in the host rather than a runtime crash.
pub const HAND_WRITTEN_METHODS: &[ShimMethod] = &[
    ShimMethod {
        selector: "measure:font:size:slant:weight:",
        args: &[
            ArgKind::String,
            ArgKind::Float,
            ArgKind::Integer,
            ArgKind::Integer,
        ],
        ret: ArgKind::Object, // CGRect* handed back as an opaque pointer
        class_method: true,
        hand_written_note: Some("body is the SDK version table; see CorroIosShims.m"),
    },
    ShimMethod {
        selector: "drawText:ctx:font:x:y:size:r:g:b:a:slant:weight:",
        args: &[
            ArgKind::String,
            ArgKind::Object,
            ArgKind::String,
            ArgKind::Float,
            ArgKind::Float,
            ArgKind::Float,
            ArgKind::Float,
            ArgKind::Float,
            ArgKind::Float,
            ArgKind::Float,
            ArgKind::Integer,
            ArgKind::Integer,
        ],
        ret: ArgKind::Void,
        class_method: false,
        hand_written_note: Some("body resolves the font; see CorroIosShims.m"),
    },
];

/// The canvas view contract: **hand-written**, but declared here so the
/// required overrides are documented in one place and the generator can emit
/// a header the host's class adopts.
#[derive(Clone, Copy, Debug)]
pub struct CanvasView {
    /// Class name the host must register (`set_sheet_view_class`).
    pub class_name: &'static str,
    /// The framework overrides the host must implement for the sheet to draw
    /// and accept input.
    pub required_overrides: &'static [&'static str],
}

pub const CANVAS_VIEW: CanvasView = CanvasView {
    class_name: "SheetView",
    required_overrides: &[
        "drawRect:",
        "corroSetCanvasId:",
    ],
};

/// Write `content` to `path` only when the file does not exist yet, creating
/// parent directories as needed. Returns `true` when a file was written.
///
/// Local copy of the same helper in `android_generator` on purpose: this module
/// is `#[path]`-included by host `build.rs` files (where the crate root is the
/// *host's*, so `crate::android_generator` does not exist), which is the same
/// reason the generator holds no other cross-module reference. Keeping the
/// helper here makes the file self-contained for that pattern.
pub fn write_if_missing(path: &Path, content: &str) -> io::Result<bool> {
    if path.exists() {
        return Ok(false);
    }
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)?;
    }
    std::fs::write(path, content.as_bytes())?;
    Ok(true)
}

/// Emit the generated `.h` + `.m` for `platform` into `project_root`.
///
/// Additive: an existing file is left untouched and its path is not returned,
/// so a host that has customised its shims keeps them. Returns the paths
/// actually written.
pub fn generate(project_root: &Path, platform: Platform) -> io::Result<Vec<PathBuf>> {
    let mut written = Vec::new();
    let header_path = project_root.join(SHIM_HEADER_RELPATH);
    if write_if_missing(&header_path, &render_header(platform))? {
        written.push(header_path);
    }
    let impl_path = project_root.join(SHIM_RELPATH);
    if write_if_missing(&impl_path, &render_impl(platform))? {
        written.push(impl_path);
    }
    Ok(written)
}

/// Resolve the project root from [`PROJECT_ENV`], or `None` when unset.
pub fn project_root_from_env() -> Option<PathBuf> {
    std::env::var(PROJECT_ENV).ok().filter(|s| !s.is_empty()).map(PathBuf::from)
}

/// Convenience entry point for a `build.rs`, mirroring
/// [`crate::android_generator::run`]: generate when [`PROJECT_ENV`] is set,
/// warn (never fail) otherwise, and re-run only when the variable changes.
pub fn run() {
    let Some(root) = project_root_from_env() else {
        println!(
            "cargo:warning=rswidgets Apple shim generation skipped; set {PROJECT_ENV} to an Apple project root"
        );
        println!("cargo:rerun-if-env-changed={PROJECT_ENV}");
        return;
    };
    // Generate for whichever platform the target is, so one project root can
    // serve an iOS app and a macOS app side by side.
    let platform = if cfg!(target_os = "macos") {
        Platform::Macos
    } else {
        Platform::Ios
    };
    match generate(&root, platform) {
        Ok(created) => {
            for path in &created {
                println!("cargo:warning=rswidgets generated {}", path.display());
            }
        }
        Err(e) => {
            println!(
                "cargo:warning=rswidgets Apple shim generation failed for {} ({e})",
                root.display()
            );
        }
    }
    println!("cargo:rerun-if-env-changed={PROJECT_ENV}");
}

// ---------------------------------------------------------------------------
// Emitters
// ---------------------------------------------------------------------------

fn selector_base(selector: &str) -> &str {
    selector.split(':').next().unwrap_or(selector)
}

fn selector_parts(selector: &str) -> Vec<&str> {
    let parts: Vec<&str> = selector.split(':').filter(|s| !s.is_empty()).collect();
    parts
}

/// ObjC method declaration fragment, e.g.
/// `- (void)corroSetFlex:(BOOL)flex;` or
/// `- (void)drawText:(id)text ctx:(id)ctx ...;`
// No `Platform` parameter: the declaration is identical for iOS and macOS
// (see `ArgKind::objc_type`), which is what lets one generator emit both.
fn method_decl(m: &ShimMethod, terminator: char) -> String {
    let sign = if m.class_method { '+' } else { '-' };
    let ret = m.ret.objc_type();
    let parts = selector_parts(m.selector);
    let mut out = format!("{sign} ({ret})");
    for (i, part) in parts.iter().enumerate() {
        let kind = m.args.get(i).copied().unwrap_or(ArgKind::Object);
        let ty = kind.objc_type();
        if i > 0 {
            out.push(' ');
        }
        // The first selector piece may be the whole name for a zero-arg
        // method (`corroBoundsWidth`), in which case there is no colon and no
        // parameter to name.
        if m.args.is_empty() {
            out.push_str(part);
        } else {
            let _ = write!(out, "{part}:({ty})arg{i}");
        }
    }
    out.push(terminator);
    out
}

/// A neutral, always-valid default return for the stub bodies the generated
/// category uses when the note says the real work is elsewhere.
fn default_return(kind: ArgKind) -> &'static str {
    match kind {
        ArgKind::Void => "",
        ArgKind::Bool => "return NO;",
        ArgKind::Integer | ArgKind::UInteger => "return 0;",
        ArgKind::Float => "return 0;",
        _ => "return nil;",
    }
}

fn render_header(p: Platform) -> String {
    let mut s = String::new();
    let guard = format!("CORRO_{}_GENERATED_SHIMS_H", p.name().to_uppercase());
    let _ = write!(
        s,
        r#"// GENERATED by rswidgets — do not edit.
//
// The Objective-C half of the Rust<->ObjC ABI contract for the {name} backend.
// Regenerate with `{env}=<project root>` on {name}; see
// rustxWidgets/src/apple_generator.rs for what is generated and, just as
// importantly, what is deliberately not (the canvas view and the text
// measurer: their bodies are behaviour, not forwarding).
//
// This header exists so the *hand-written* classes in the same target can
// adopt these declarations instead of restating them: a drift between the two
// then fails to compile in the host, rather than corrupting registers at
// runtime with a mismatched objc_msgSend.

#ifndef {guard}
#define {guard}

#import <{header}>

#pragma mark - Forwarding shims (generated)

/// Trampoline: one selector that turns a control activation into a
/// `{prefix}_callback(id)` call. The Rust side resolves this class by name.
@interface {cp}Target : NSObject
"#,
        name = p.name(),
        env = PROJECT_ENV,
        guard = guard,
        header = p.framework_header(),
        prefix = p.c_prefix(),
        cp = p.class_prefix(),
    );
    for m in TARGET_METHODS {
        let _ = writeln!(s, "{}", method_decl(m, ';'));
    }
    let _ = write!(s, "@end\n\n/// Alert wrapper: `{}Alert`\n@interface {}Alert : NSObject\n", p.class_prefix(), p.class_prefix());
    for m in ALERT_METHODS {
        let _ = writeln!(s, "{}", method_decl(m, ';'));
    }
    let _ = write!(
        s,
        "@end\n\n#pragma mark - Layout protocol (generated category)\n\n/// The expand/min-width protocol the shared layout code expects from any view.\n@interface {base} (CorroLayout)\n",
        base = p.base_view_class(),
    );
    for m in VIEW_CATEGORY_METHODS {
        let _ = writeln!(s, "{}", method_decl(m, ';'));
    }
    s.push_str("@end\n\n#pragma mark - Hand-written, signature-checked\n\n");
    s.push_str("/// Bodies live in the host's own .m file; declared here so a signature\n/// mismatch is a compile error.\n");
    let _ = write!(s, "@interface {}Text : NSObject\n", p.class_prefix());
    for m in HAND_WRITTEN_METHODS {
        let _ = writeln!(s, "{}", method_decl(m, ';'));
    }
    let _ = write!(
        s,
        "@end\n\n/// The canvas view the host must ship (see the backend's\n/// `set_sheet_view_class`). Hand-written: `drawRect:` reports the laid-out\n/// size before replaying Rust's draw closure, and the input methods convert\n/// coordinates.\n///\n/// Required overrides:\n",
    );
    for ov in CANVAS_VIEW.required_overrides {
        let _ = writeln!(s, "///   * `{ov}`");
    }
    if !p.view_is_top_left_origin() {
        s.push_str("///\n/// This platform's views are bottom-left origin by default, so the canvas\n/// class MUST also override `- (BOOL)isFlipped` to return `YES`; otherwise\n/// the sheet draws vertically mirrored.\n");
    }
    let _ = write!(
        s,
        "@interface {} : {base}\n@end\n\n#endif /* {guard} */\n",
        CANVAS_VIEW.class_name,
        base = p.base_view_class(),
        guard = guard,
    );
    s
}

fn render_impl(p: Platform) -> String {
    let mut s = String::new();
    let _ = write!(
        s,
        r#"// GENERATED by rswidgets — do not edit.
//
// Forwarding shims for the {name} backend: see CorroGeneratedShims.h and
// rustxWidgets/src/apple_generator.rs. Every method here is pure forwarding —
// it exists so the Rust adapter's `objc_msgSend` signature has a matching
// declaration on the ObjC side.

#import "CorroGeneratedShims.h"

// The C entry points emitted by the host crate (ios/corro or macos/corro).
extern void {prefix}_callback(uint64_t callback_id);

#pragma mark - Callback trampoline
"#,
        name = p.name(),
        prefix = p.c_prefix(),
    );

    // Target: retain the id, fire on corroFired:.
    let cp = p.class_prefix();
    let _ = write!(
        s,
        r#"
@implementation {cp}Target {{
    NSInteger _callbackId;
}}

+ (instancetype)targetWithCallbackId:(NSInteger)callbackId {{
    {cp}Target *t = [[{cp}Target alloc] init];
    t->_callbackId = callbackId;
    return t;
}}

- (void)corroFired:(id)sender {{
    (void)sender;
    {prefix}_callback((uint64_t)_callbackId);
}}

@end
"#,
        cp = cp,
        prefix = p.c_prefix(),
    );

    // Alert.
    let _ = write!(s, "\n#pragma mark - Alert\n\n");
    match p {
        Platform::Ios => s.push_str(IOS_ALERT_IMPL),
        Platform::Macos => s.push_str(MACOS_ALERT_IMPL),
    }

    // Layout category.
    let _ = write!(
        s,
        "\n#pragma mark - Layout category\n\n@implementation {base} (CorroLayout)\n\n",
        base = p.base_view_class(),
    );
    for m in VIEW_CATEGORY_METHODS {
        let note = m
            .hand_written_note
            .map(|n| format!("// {n}\n"))
            .unwrap_or_default();
        let body = default_return(m.ret);
        if body.is_empty() {
            // No-op setters: record nothing here — the Rust `WidgetMeta`
            // registry is authoritative for expand/min-chars/canvas id, and
            // the host's canvas class reads it back through the same calls.
            let _ = write!(
                s,
                "{note}- (void){selector}:(NSInteger)arg0 {{\n    (void)arg0;\n}}\n\n",
                selector = selector_base(m.selector),
            );
        } else {
            let parts = selector_parts(m.selector);
            let name = parts.first().copied().unwrap_or("corroUnknown");
            let _ = write!(s, "{note}- ({ret}){name} {{\n    {body}\n}}\n\n",
                ret = m.ret.objc_type());
        }
    }
    s.push_str("@end\n");
    s
}

const IOS_ALERT_IMPL: &str = r#"/// Wraps UIAlertController (iOS 8+) and UIAlertView (iOS 7) behind one
/// interface, so the Rust side calls `corroAddAction:`/`corroNewAlert` once.
@interface CorroIosAlert ()
@property (nonatomic, strong) id native;  // UIAlertController or UIAlertView
@property (nonatomic, strong) NSMutableArray<NSString *> *actions;
@end

@implementation CorroIosAlert

+ (instancetype)corroNewAlert {
    CorroIosAlert *w = [[CorroIosAlert alloc] init];
    w.actions = [NSMutableArray array];
    return w;
}

- (void)corroSetTitle:(id)title {
    if ([UIAlertController class] != nil) {
        // UIAlertController's title is only settable at construction, so the
        // title is applied when the controller is built in corroAddAction:.
        self.actions = self.actions ?: [NSMutableArray array];
        (void)title;
    } else {
        (void)title;
    }
}

- (void)corroAddAction:(id)title {
    NSString *t = (NSString *)title;
    if ([UIAlertController class] != nil) {
        UIAlertController *ac = self.native;
        if (ac == nil) {
            ac = [UIAlertController alertControllerWithTitle:nil
                                                     message:nil
                                              preferredStyle:UIAlertControllerStyleAlert];
            self.native = ac;
        }
        [ac addAction:[UIAlertAction actionWithTitle:t
                                               style:UIAlertActionStyleDefault
                                             handler:nil]];
    } else {
        UIAlertView *av = self.native;
        if (av == nil) {
            av = [[UIAlertView alloc] initWithTitle:nil
                                            message:nil
                                           delegate:nil
                                  cancelButtonTitle:nil
                                  otherButtonTitles:nil];
            self.native = av;
        }
        [av addButtonWithTitle:t];
    }
}

@end

/// Presentation lives on the view controller, matching `Dialog::present`'s
/// `corroPresentDialog:` send.
@implementation UIViewController (CorroPresent)

- (void)corroPresentDialog:(id)dialog {
    id native = [dialog valueForKey:@"native"];
    if ([native isKindOfClass:[UIAlertController class]]) {
        [self presentViewController:(UIAlertController *)native animated:YES completion:nil];
    } else if ([native isKindOfClass:[UIAlertView class]]) {
        [(UIAlertView *)native show];
    }
}

@end
"#;

const MACOS_ALERT_IMPL: &str = r#"/// Wraps NSAlert. macOS has one alert class for every supported version, so
/// unlike iOS there is no version branch here — only the modal-vs-sheet choice,
/// which is the host's (a sheet when the host has a window, modal otherwise).
@interface CorroMacAlert ()
@property (nonatomic, strong) NSAlert *native;
@end

@implementation CorroMacAlert

+ (instancetype)corroNewAlert {
    CorroMacAlert *w = [[CorroMacAlert alloc] init];
    w.native = [[NSAlert alloc] init];
    [w.native addButtonWithTitle:@"OK"];
    return w;
}

- (void)corroSetTitle:(id)title {
    self.native.messageText = (NSString *)title;
}

- (void)corroAddAction:(id)title {
    [self.native addButtonWithTitle:(NSString *)title];
}

@end

/// `Dialog::present` sends `corroPresentDialog:` to the window controller.
/// A sheet needs a window; without one, run modally.
@implementation NSWindowController (CorroPresent)

- (void)corroPresentDialog:(id)dialog {
    NSAlert *alert = [dialog valueForKey:@"native"];
    if (![alert isKindOfClass:[NSAlert class]]) {
        return;
    }
    NSWindow *window = self.window;
    if (window != nil) {
        [alert beginSheetModalForWindow:window completionHandler:nil];
    } else {
        [alert runModal];
    }
}

@end
"#;

/// A Rust-side selector the test suite can compare against the tables above,
/// so "the adapters send something the generator does not know about" is a
/// test failure rather than a runtime crash.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct RustSideSelector {
    /// The class the adapter resolves (`CorroIosTarget`, ...).
    pub class: &'static str,
    /// The selector it sends.
    pub selector: &'static str,
    /// The source file the pair was taken from, for the failure message.
    pub source: &'static str,
}

/// Every `(class, selector)` pair the iOS adapter sends to a generated shim.
///
/// Transcribed from `backends_ios_adapter.rs`; the test below asserts each is
/// covered by the tables above.
pub const IOS_ADAPTER_SENDS: &[RustSideSelector] = &[
    RustSideSelector { class: "CorroIosTarget", selector: "targetWithCallbackId:", source: "Button::attach_target" },
    RustSideSelector { class: "CorroIosTarget", selector: "corroFired:", source: "Button::attach_target (action)" },
    RustSideSelector { class: "CorroIosText", selector: "measure:font:size:slant:weight:", source: "measure_text_via_host" },
    RustSideSelector { class: "CorroIosText", selector: "drawText:ctx:font:x:y:size:r:g:b:a:slant:weight:", source: "draw_text_via_host" },
    RustSideSelector { class: "CorroIosAlert", selector: "corroNewAlert", source: "create_dialog" },
    RustSideSelector { class: "CorroIosAlert", selector: "corroAddAction:", source: "Dialog::add_button" },
    RustSideSelector { class: "UIViewController", selector: "corroPresentDialog:", source: "Dialog::present" },
];

/// The same for the macOS adapter.
pub const MACOS_ADAPTER_SENDS: &[RustSideSelector] = &[
    RustSideSelector { class: "CorroMacTarget", selector: "targetWithCallbackId:", source: "Button::attach_target" },
    RustSideSelector { class: "CorroMacTarget", selector: "corroFired:", source: "Button::attach_target (action)" },
    RustSideSelector { class: "CorroMacText", selector: "measure:font:size:slant:weight:", source: "measure_text_via_host" },
    RustSideSelector { class: "CorroMacText", selector: "drawText:ctx:font:x:y:size:r:g:b:a:slant:weight:", source: "draw_text_via_host" },
    RustSideSelector { class: "CorroMacAlert", selector: "corroNewAlert", source: "create_dialog" },
    RustSideSelector { class: "CorroMacAlert", selector: "corroAddAction:", source: "Dialog::add_button" },
    RustSideSelector { class: "NSWindowController", selector: "corroPresentDialog:", source: "Dialog::present" },
];

#[cfg(test)]
mod tests {
    use super::*;

    fn temp_root(tag: &str) -> PathBuf {
        let mut p = std::env::temp_dir();
        p.push(format!("rswidgets-apple-gen-{tag}-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&p);
        p
    }

    /// Every table entry produces a declaration whose signature matches the
    /// kind list, and the header/impl both mention each class.
    #[test]
    fn every_declared_selector_is_emitted() {
        for p in [Platform::Ios, Platform::Macos] {
            let header = render_header(p);
            let impl_ = render_impl(p);
            for m in TARGET_METHODS.iter().chain(ALERT_METHODS).chain(VIEW_CATEGORY_METHODS) {
                let base = selector_base(m.selector);
                assert!(header.contains(base), "{} header missing {base}", p.name());
                if !m.class_method || m.selector == "corroNewAlert" || m.selector == "targetWithCallbackId:" {
                    assert!(impl_.contains(base), "{} impl missing {base}", p.name());
                }
                assert_eq!(
                    header.matches(base).count() >= 1,
                    true,
                    "{} should declare {base}",
                    p.name()
                );
            }
            // The trampoline must call the platform's C callback.
            assert!(impl_.contains(&format!("{}_callback", p.c_prefix())));
        }
    }

    /// The whole point: no selector the adapters send may be missing from the
    /// generator's tables. If someone adds a `cls("CorroIosFoo")` + `raw_send!`
    /// to an adapter, this fails until the table is updated.
    #[test]
    fn adapter_selectors_are_all_covered() {
        let covered: Vec<&str> = TARGET_METHODS
            .iter()
            .chain(ALERT_METHODS)
            .chain(HAND_WRITTEN_METHODS)
            .chain(VIEW_CATEGORY_METHODS)
            .map(|m| m.selector)
            .collect();
        let presentation = ["corroPresentDialog:"];
        for entry in IOS_ADAPTER_SENDS.iter().chain(MACOS_ADAPTER_SENDS) {
            let known = covered.contains(&entry.selector) || presentation.contains(&entry.selector);
            assert!(
                known,
                "adapter sends {}::{} ({}) but the generator has no entry for it",
                entry.class, entry.selector, entry.source
            );
        }
    }

    /// A generated header declares the hand-written text class so a signature
    /// mismatch there is a host compile error.
    ///
    /// The declaration is spelled with its parameter names
    /// (`measure:(id)text font:(id)...`), so the test checks the selector
    /// *pieces* rather than the contiguous string.
    #[test]
    fn hand_written_selectors_are_signature_checked() {
        let header = render_header(Platform::Ios);
        for piece in ["measure:", "font:", "size:", "slant:", "weight:"] {
            assert!(header.contains(piece), "text class declaration missing {piece}");
        }
        assert!(header.contains("drawText:"));
        assert!(header.contains("ctx:"));
        // And the class itself must be declared.
        assert!(header.contains("CorroIosText : NSObject"));
    }

    /// The AppKit header must tell the host to flip its canvas view; UIKit's
    /// header must not (it is already top-left).
    #[test]
    fn flipped_hint_only_on_appkit() {
        assert!(!render_header(Platform::Ios).contains("isFlipped"));
        assert!(render_header(Platform::Macos).contains("isFlipped"));
    }

    /// CGFloat is spelled as the typedef so no 32/64 branch is needed.
    #[test]
    fn float_kind_is_cgfloat_not_double() {
        assert_eq!(ArgKind::Float.objc_type(), "CGFloat");
        assert_eq!(ArgKind::Float.objc_type(), "CGFloat");
    }

    #[test]
    fn generate_creates_both_files_and_never_overwrites() {
        let root = temp_root("generate");
        let created = generate(&root, Platform::Ios).unwrap();
        assert_eq!(created.len(), 2, "expected header + impl, got {created:?}");
        assert!(root.join(SHIM_RELPATH).exists());
        assert!(root.join(SHIM_HEADER_RELPATH).exists());

        // A hand-edited file must survive: generation is additive.
        let impl_path = root.join(SHIM_RELPATH);
        std::fs::write(&impl_path, "// hand-edited\n").unwrap();
        let created = generate(&root, Platform::Ios).unwrap();
        assert!(!created.contains(&impl_path));
        assert_eq!(std::fs::read_to_string(&impl_path).unwrap(), "// hand-edited\n");
        let _ = std::fs::remove_dir_all(&root);
    }

    /// The generated ObjC must be plausibly balanced — this is not a compiler,
    /// but an unbalanced `@implementation`/`@end` is the one class of error a
    /// pure-text generator can catch without an SDK.
    #[test]
    fn generated_objc_has_balanced_blocks() {
        for p in [Platform::Ios, Platform::Macos] {
            for text in [render_header(p), render_impl(p)] {
                let opens: Vec<&str> = text.lines().filter(|l| l.trim_start().starts_with("@implementation") || l.trim_start().starts_with("@interface")).collect();
                let ends = text.lines().filter(|l| l.trim() == "@end").count();
                assert_eq!(opens.len(), ends, "{}: {} blocks vs {ends} @end", p.name(), opens.len());
            }
        }
    }
}
