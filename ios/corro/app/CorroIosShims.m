//
//  CorroIosShims.m
//  corro on iOS — the ObjC classes the Rust backend resolves by name.
//
//  Every class here exists because the Rust side cannot (or should not) know
//  the version-specific UIKit API for it. rustxWidgets/docs/IOS_GUIDELINES.md
//  §3 is the contract table; this file is one worked implementation of it.
//
//  Nothing here calls back into Rust on its own initiative: each method is a
//  funnel from a UIKit callback into one `corro_ios_*` C function.
//

#import <UIKit/UIKit.h>
#import <CoreGraphics/CoreGraphics.h>
#import <objc/runtime.h>

#import "CorroBridge.h"
// The generated half of the Rust<->ObjC ABI contract: the forwarding shims and
// the layout category live in CorroGeneratedShims.{h,m} (regenerate with
// `cargo build --features generate-apple-shims` in this crate). Importing it
// here means the classes below *adopt* those declarations rather than restate
// them, so a signature that drifts from the Rust side is a compile error
// instead of a mismatched objc_msgSend at runtime.
#import "CorroGeneratedShims.h"
#import <stdio.h>

// ---------------------------------------------------------------------------
// SheetView — the grid canvas
// ---------------------------------------------------------------------------

// The class the Rust backend instantiates for `create_canvas` (registered
// with `set_sheet_view_class`, see corro_ios_root_ready). It remembers the
// Rust canvas id so draw and touch callbacks can name it.
// `SheetView : UIView` is declared by the GENERATED header; this is the
// extension carrying the state the hand-written body needs. (Re-declaring the
// class itself is what collided with the generated header.)
@interface SheetView ()
/// Set by the backend right after construction (`corroSetCanvasId:`).
@property (nonatomic, assign) uint64_t corroCanvasId;
/// Re-entrancy guard for the fill-the-superview adjustment in layoutSubviews.
@property (nonatomic, assign) BOOL corroFillingSuperview;
/// Declared with its availability so clang accepts the UIKey uses inside.
- (void)corroHandlePresses:(NSSet<UIPress *> *)presses API_AVAILABLE(ios(13.4));
@end

// Declared ahead of use so the availability annotation is visible at the call
// site as well as at the definition (clang checks both).
@implementation SheetView

- (void)corroSetCanvasId:(int64_t)canvasId {
    self.corroCanvasId = (uint64_t)canvasId;
}

// Report the laid-out size as soon as UIKit knows it.
//
// Rust replays the draw closure BEFORE UIKit's first `drawRect:` (registration
// time and every `queue_redraw`), so if the size only arrived with the first
// frame the early replays would use the 1x1 placeholder from
// `set_size_request` and lay the sheet out as a single row.
- (void)layoutSubviews {
    [super layoutSubviews];
    CGRect b = self.bounds;
    // A view that is part of a stack view is sized by it; this only matters for
    // the ones attached with addSubview: (the canvas inside its container),
    // where nothing else gives it a height. Filling the superview is the
    // intended behaviour for a sheet, and it stops the 0-height case that made
    // drawRect: never fire.
    if (self.superview != nil && (CGRectGetHeight(b) < 1.0 || CGRectGetWidth(b) < 1.0)) {
        // Guarded: changing the frame inside layoutSubviews schedules another
        // pass, and without the flag that recurses until the stack runs out.
        if (!self.corroFillingSuperview) {
            self.corroFillingSuperview = YES;
            self.frame = self.superview.bounds;
            b = self.bounds;
            self.corroFillingSuperview = NO;
        }
    }
    fprintf(stderr, "[corro] SheetView.layoutSubviews canvas=%llu %dx%d\n",
            self.corroCanvasId, (int)CGRectGetWidth(b), (int)CGRectGetHeight(b));
    fflush(stderr);
    if (CGRectGetWidth(b) > 0 && CGRectGetHeight(b) > 0) {
        corro_ios_canvas_size(self.corroCanvasId,
                              (int32_t)CGRectGetWidth(b),
                              (int32_t)CGRectGetHeight(b));
    }
}

// All rendering is in Rust: hand the live CGContext over and let the
// registered draw closure replay its primitives.
- (void)drawRect:(CGRect)rect {
    CGContextRef ctx = UIGraphicsGetCurrentContext();
    CGRect b = self.bounds;
    fprintf(stderr, "[corro] SheetView.drawRect canvas=%llu %dx%d ctx=%s\n",
            self.corroCanvasId,
            (int)CGRectGetWidth(b), (int)CGRectGetHeight(b),
            ctx == NULL ? "NULL" : "ok");
    fflush(stderr);
    // Report the size here as well as in layoutSubviews: at draw time the
    // bounds are final, which is exactly what Rust needs before it replays the
    // draw closure.
    if (CGRectGetWidth(b) > 0 && CGRectGetHeight(b) > 0) {
        corro_ios_canvas_size(self.corroCanvasId,
                              (int32_t)CGRectGetWidth(b),
                              (int32_t)CGRectGetHeight(b));
    }
    if (ctx == NULL) {
        return;
    }
    corro_ios_canvas_draw(self.corroCanvasId, (void *)ctx,
                          (int32_t)CGRectGetWidth(self.bounds),
                          (int32_t)CGRectGetHeight(self.bounds));
}

// A tap moves the cursor. `touchesEnded:` rather than `touchesBegan:` so a
// drag that starts on the grid does not move the selection mid-gesture.
- (void)touchesEnded:(NSSet<UITouch *> *)touches withEvent:(UIEvent *)event {
    UITouch *touch = touches.anyObject;
    if (touch != nil) {
        CGPoint p = [touch locationInView:self];
        corro_ios_canvas_click(self.corroCanvasId, (double)p.x, (double)p.y);
    }
    [super touchesEnded:touches withEvent:event];
}

// The view must be able to take focus for hardware keyboards (iPad, simctl).
- (BOOL)canBecomeFirstResponder {
    return YES;
}

// Hardware keys are funnelled the same way touches are. `keyCommands` would be
// the modern route but only handles a fixed set; `pressesBegan:` covers a
// physical keyboard and the iPad keyboard accessory bar.
//
// `UIPress.key` and the whole `UIKey` class only exist on iOS 13.4+, while this
// app deploys to 12.0 — so the whole body is behind an availability check and
// the method degrades to UIKit's default handling on older systems. (Found by
// CI: clang rejected an unguarded use with -Wunguarded-availability-new.)
- (void)pressesBegan:(NSSet<UIPress *> *)presses withEvent:(UIPressesEvent *)event {
    if (@available(iOS 13.4, *)) {
        [self corroHandlePresses:presses];
    }
    [super pressesBegan:presses withEvent:event];
}

// Annotated API_AVAILABLE so clang accepts the UIKey uses inside; the caller
// checks the same availability before calling. (Guarding only at the call site
// is not enough - clang analyses each method body on its own, which is why the
// warning persisted after the previous attempt.)
- (void)corroHandlePresses:(NSSet<UIPress *> *)presses API_AVAILABLE(ios(13.4)) {
    for (UIPress *press in presses) {
        UIKey *key = press.key;
        if (key == nil) {
            continue;
        }
        uint32_t keyval = 0;
        if (key.characters.length > 0) {
            keyval = (uint32_t)[key.characters characterAtIndex:0];
        } else {
            // Non-printable: map the escape-ish keys the shared key handling
            // knows (see rswidgets::core::key) onto their constants.
            switch (key.keyCode) {
                case UIKeyboardHIDUsageKeyboardReturnOrEnter:
                case UIKeyboardHIDUsageKeypadEnter: keyval = 0xFF0D; break;
                case UIKeyboardHIDUsageKeyboardEscape: keyval = 0xFF1B; break;
                case UIKeyboardHIDUsageKeyboardTab: keyval = 0xFF09; break;
                // NB: the SDK name is DeleteOrBackspace, not "Backspace" (the
                // wrong spelling is a compile error, not a fallback - found by
                // probing UIKeyConstants.h on a real SDK).
                case UIKeyboardHIDUsageKeyboardDeleteOrBackspace: keyval = 0xFF08; break;
                case UIKeyboardHIDUsageKeyboardDeleteForward: keyval = 0xFFFF; break;
                case UIKeyboardHIDUsageKeyboardLeftArrow: keyval = 0xFF51; break;
                case UIKeyboardHIDUsageKeyboardUpArrow: keyval = 0xFF52; break;
                case UIKeyboardHIDUsageKeyboardRightArrow: keyval = 0xFF53; break;
                case UIKeyboardHIDUsageKeyboardDownArrow: keyval = 0xFF54; break;
                case UIKeyboardHIDUsageKeyboardHome: keyval = 0xFF50; break;
                case UIKeyboardHIDUsageKeyboardEnd: keyval = 0xFF57; break;
                default: break;
            }
        }
        if (keyval == 0) {
            continue;
        }
        // Modifier bitmask matches the shared convention: 1 = Shift,
        // 4 = Ctrl, 8 = Alt.
        uint32_t mods = 0;
        if ((key.modifierFlags & UIKeyModifierShift) != 0) mods |= 1;
        if ((key.modifierFlags & UIKeyModifierControl) != 0) mods |= 4;
        if ((key.modifierFlags & UIKeyModifierAlternate) != 0) mods |= 8;
        if (corro_ios_canvas_key(self.corroCanvasId, keyval, mods)) {
            // Consumed: handled. (The caller has already forwarded to super,
            // so returning here simply stops inspecting further presses.)
            return;
        }
    }
}

@end

// ---------------------------------------------------------------------------
// UIView (CorroLayout) — the layout hints the Rust backend sends
// ---------------------------------------------------------------------------

// These are NOT optional. The backend is written against a selector contract,
// and Objective-C raises `unrecognized selector sent to instance` for a missing
// implementation - it does not silently ignore the message. The app was dying
// at exactly this point (`run_gui: new_box` printed, nothing after), because
// create_box sends `corroSetSpacing:`, which nothing implemented.
//
// A category on UIView covers every widget, since they are all UIViews; the
// methods only record what the backend asked for, and `layoutSubviews` in the
// concrete views uses it where it matters.
// The method *signatures* come from CorroGeneratedShims.h (generated from the
// Rust side's `raw_send!` calls); only the storage the implementation needs is
// declared here. Adding a method to the generated set therefore does not
// require touching this file, and changing one of its signatures cannot
// silently diverge.

// ---------------------------------------------------------------------------
// CorroIosTarget — the callback trampoline
// ---------------------------------------------------------------------------

// UIControl targets are (id, SEL, id), so the callback id has to travel as an
// object or as the target itself. The target is the simpler carrier: the
// backend calls `targetWithCallbackId:` and we hold the id.
// `targetWithCallbackId:` and `corroFired:` are declared in the generated
// header; this extension adds the storage the implementation needs.
@interface CorroIosTarget ()
@property (nonatomic, assign) uint64_t callbackId;
@end


// ---------------------------------------------------------------------------
// CorroIosPicker — the drop-down contract
// ---------------------------------------------------------------------------

// The Rust DropDown sends corroAddPickerItem:, corroSetSelectedIndex: and
// corroSelectedIndex. Same lesson as the layout hints above: a missing
// implementation is a crash, not a no-op, so all three exist even though the
// picker UI itself is a plain UIButton here.
@interface CorroIosPicker : UIButton
@property (nonatomic, strong) NSMutableArray<NSString *> *corroItems;
@property (nonatomic, assign) NSInteger corroIndex;
@end

// The generated header declares `corroBoundsWidth`/`corroBoundsHeight` but
// emits no body for them (the generator's stub would return 0, and the adapter
// sizes children to their parent with these - a 0 return is the 0x0 canvas
// that made the sheet never draw). The bodies live here.
@implementation UIView (CorroBounds)
- (CGFloat)corroBoundsWidth {
    return CGRectGetWidth(self.bounds);
}
- (CGFloat)corroBoundsHeight {
    return CGRectGetHeight(self.bounds);
}
@end

@implementation CorroIosPicker

- (void)corroAddPickerItem:(NSString *)title {
    if (self.corroItems == nil) {
        self.corroItems = [NSMutableArray array];
    }
    [self.corroItems addObject:title];
    // Show the first item until something is selected.
    if (self.corroItems.count == 1) {
        [self setTitle:title forState:UIControlStateNormal];
    }
}

- (void)corroSetSelectedIndex:(NSInteger)index {
    self.corroIndex = index;
    if (index >= 0 && (NSUInteger)index < self.corroItems.count) {
        [self setTitle:self.corroItems[(NSUInteger)index] forState:UIControlStateNormal];
    }
}

- (NSInteger)corroSelectedIndex {
    return self.corroIndex;
}

@end

// ---------------------------------------------------------------------------
// CorroIosText — font resolution, measurement and drawing
// ---------------------------------------------------------------------------

// Version-specific font/text API lives here so the Rust side does not have to
// branch on the deployment target:
//   iOS 7+  boundingRectWithSize:options:attributes: / drawAtPoint:withAttributes:
//   iOS 6- sizeWithFont: / drawAtPoint:withFont:
// `measure:font:size:slant:weight:` and `drawText:ctx:font:x:y:size:r:g:b:a:
// slant:weight:` are declared in the generated header. The font helper below is
// this file's own business, so it stays in an extension.
@interface CorroIosText ()
// `measure:` and `drawText:` are declared by the GENERATED header
// (CorroGeneratedShims.h) from the same signature table the Rust adapter
// sends, so they are deliberately not restated here - a second declaration is
// what produced the conflicting-types errors, and one source of truth is the
// point of generating them.
//
// ⚠️ Both are CLASS methods (`+`). The Rust adapter resolves this class with
// `objc_getClass` and sends the selectors to the CLASS, because the shim is
// stateless and never instantiated. Implementing them as instance methods made
// `[CorroIosText drawText:...]` raise "unrecognized selector sent to instance",
// which the Rust side reported as
//   fatal runtime error: Rust cannot catch foreign exceptions
// on the first real frame - an ObjC exception crossing into Rust, and the
// offending selector appeared nowhere in the log. Keep them `+`.
+ (UIFont *)fontForFamily:(NSString *)family size:(CGFloat)size weight:(NSInteger)weight;
@end

@implementation CorroIosText

+ (UIFont *)fontForFamily:(NSString *)family size:(CGFloat)size weight:(NSInteger)weight {
    UIFont *font = nil;
    if (family.length > 0) {
        font = [UIFont fontWithName:family size:size];
    }
    if (font == nil) {
        // The grid asks for a monospace family by preference; the system font
        // is the honest fallback, and the Rust estimate assumes monospace, so
        // a measured layout stays close either way.
        if (@available(iOS 13.0, *)) {
            font = [UIFont monospacedSystemFontOfSize:size
                                               weight:(weight != 0 ? UIFontWeightBold
                                                                   : UIFontWeightRegular)];
        } else {
            font = [UIFont fontWithName:@"Courier" size:size];
            if (font == nil) {
                font = [UIFont systemFontOfSize:size];
            }
        }
    }
    if (font == nil) {
        font = [UIFont systemFontOfSize:size];
    }
    return font;
}

// Returns a malloc'd CGRect* (or NULL) so the ABI is one pointer in / one
// pointer out — no struct-return register rules to get wrong on 32-bit armv7.
// The caller (Rust) frees it with `free`.
+ (CGRect *)measure:(NSString *)text
               font:(NSString *)family
               size:(CGFloat)size
              slant:(NSInteger)slant
             weight:(NSInteger)weight {
    (void)slant; // no synthetic oblique on iOS; weight is what the grid uses
    if (text.length == 0) {
        CGRect *r = (CGRect *)calloc(1, sizeof(CGRect));
        return r;
    }
    UIFont *font = [CorroIosText fontForFamily:family size:size weight:weight];
    CGSize measured;
    if ([text respondsToSelector:@selector(boundingRectWithSize:options:attributes:context:)]) {
        CGRect box = [text boundingRectWithSize:CGSizeMake(CGFLOAT_MAX, CGFLOAT_MAX)
                                        options:NSStringDrawingUsesLineFragmentOrigin
                                     attributes:@{ NSFontAttributeName: font }
                                        context:nil];
        measured = box.size;
    } else {
#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wdeprecated-declarations"
        measured = [text sizeWithFont:font];
#pragma clang diagnostic pop
    }
    CGRect *r = (CGRect *)calloc(1, sizeof(CGRect));
    r->origin = CGPointZero;
    r->size = measured;
    return r;
}

// Draws at the shared top-left convention: the caller gives a top edge, the
// text API wants a baseline, so offset by the ascent.
+ (void)drawText:(NSString *)text
             ctx:(CGContextRef)ctx
            font:(NSString *)family
               x:(CGFloat)x
               y:(CGFloat)y
            size:(CGFloat)size
               r:(CGFloat)red
               g:(CGFloat)green
               b:(CGFloat)blue
               a:(CGFloat)alpha
           slant:(NSInteger)slant
          weight:(NSInteger)weight {
    (void)slant;
    if (text.length == 0 || ctx == NULL) {
        return;
    }
    UIFont *font = [CorroIosText fontForFamily:family size:size weight:weight];
    CGContextSaveGState(ctx);
    // UIKit's coordinate system is top-left origin, y down — matching the
    // shared draw convention, so no flip is needed inside drawRect:'s context.
    [[UIColor colorWithRed:red green:green blue:blue alpha:alpha] setFill];
    CGFloat baseline = y + font.ascender;
    if ([text respondsToSelector:@selector(drawAtPoint:withAttributes:)]) {
        [text drawAtPoint:CGPointMake(x, baseline)
           withAttributes:@{ NSFontAttributeName: font,
                             NSForegroundColorAttributeName:
                                 [UIColor colorWithRed:red green:green blue:blue alpha:alpha] }];
    } else {
#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wdeprecated-declarations"
        [[UIColor colorWithRed:red green:green blue:blue alpha:alpha] set];
        [text drawAtPoint:CGPointMake(x, baseline) withFont:font];
#pragma clang diagnostic pop
    }
    CGContextRestoreGState(ctx);
}

@end

// ---------------------------------------------------------------------------
// CorroIosAlert — dialogs across iOS versions
// ---------------------------------------------------------------------------

// `corroNewAlert` returns a UIAlertController on iOS 8+ and a UIAlertView on
// iOS 7, wrapped so the Rust side sees one API. The wrapper is an NSObject
// holding whichever concrete object the running OS supports.
// `corroNewAlert`, `corroAddAction:` and `corroSetTitle:` are declared in the
// generated header; only the private storage is declared here.
@interface CorroIosAlert ()
@property (nonatomic, strong) id native;       // UIAlertController or UIAlertView
@property (nonatomic, strong) NSMutableArray<NSString *> *titles;
@end


// `corroPresentDialog:` on a view controller: present whatever the alert
// wrapper holds. Declared as a category so the backend's message send resolves.
@interface UIViewController (CorroIosPresent)
- (void)corroPresentDialog:(id)dialog;
@end

// `corroPresentDialog:` is implemented by the GENERATED category
// `UIViewController (CorroPresent)` in CorroGeneratedShims.m. It was also
// implemented here, which is two implementations of one selector across two
// categories of the same class - the later-loaded one silently wins.
