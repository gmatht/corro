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

// ---------------------------------------------------------------------------
// SheetView — the grid canvas
// ---------------------------------------------------------------------------

// The class the Rust backend instantiates for `create_canvas` (registered
// with `set_sheet_view_class`, see corro_ios_root_ready). It remembers the
// Rust canvas id so draw and touch callbacks can name it.
@interface SheetView : UIView
/// Set by the backend right after construction (`corroSetCanvasId:`).
@property (nonatomic, assign) uint64_t corroCanvasId;
@end

// Declared ahead of use so the availability annotation is visible at the call
// site as well as at the definition (clang checks both).
@interface SheetView ()
- (void)corroHandlePresses:(NSSet<UIPress *> *)presses API_AVAILABLE(ios(13.4));
@end

@implementation SheetView

- (void)corroSetCanvasId:(int64_t)canvasId {
    self.corroCanvasId = (uint64_t)canvasId;
}

// All rendering is in Rust: hand the live CGContext over and let the
// registered draw closure replay its primitives.
- (void)drawRect:(CGRect)rect {
    CGContextRef ctx = UIGraphicsGetCurrentContext();
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
// CorroIosTarget — the callback trampoline
// ---------------------------------------------------------------------------

// UIControl targets are (id, SEL, id), so the callback id has to travel as an
// object or as the target itself. The target is the simpler carrier: the
// backend calls `targetWithCallbackId:` and we hold the id.
@interface CorroIosTarget : NSObject
@property (nonatomic, assign) uint64_t callbackId;
+ (instancetype)targetWithCallbackId:(NSInteger)callbackId;
- (void)corroFired:(id)sender;
@end

@implementation CorroIosTarget

+ (instancetype)targetWithCallbackId:(NSInteger)callbackId {
    CorroIosTarget *t = [[CorroIosTarget alloc] init];
    t.callbackId = (uint64_t)callbackId;
    return t;
}

- (void)corroFired:(id)sender {
    (void)sender;
    corro_ios_callback(self.callbackId);
}

@end

// ---------------------------------------------------------------------------
// CorroIosText — font resolution, measurement and drawing
// ---------------------------------------------------------------------------

// Version-specific font/text API lives here so the Rust side does not have to
// branch on the deployment target:
//   iOS 7+  boundingRectWithSize:options:attributes: / drawAtPoint:withAttributes:
//   iOS 6- sizeWithFont: / drawAtPoint:withFont:
@interface CorroIosText : NSObject
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
- (CGRect *)measure:(NSString *)text
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
- (void)drawText:(NSString *)text
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
@interface CorroIosAlert : NSObject
+ (instancetype)corroNewAlert;
- (void)corroAddAction:(NSString *)title;
- (void)corroSetTitle:(NSString *)title;
@end

@interface CorroIosAlert ()
@property (nonatomic, strong) id native;       // UIAlertController or UIAlertView
@property (nonatomic, strong) NSMutableArray<NSString *> *titles;
@end

@implementation CorroIosAlert

+ (instancetype)corroNewAlert {
    CorroIosAlert *wrapper = [[CorroIosAlert alloc] init];
    wrapper.titles = [NSMutableArray array];
    if ([UIAlertController class] != nil) {
        wrapper.native = [UIAlertController alertControllerWithTitle:nil
                                                            message:nil
                                                     preferredStyle:UIAlertControllerStyleAlert];
    } else {
#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wdeprecated-declarations"
        wrapper.native = [[UIAlertView alloc] initWithTitle:nil
                                                   message:nil
                                                  delegate:nil
                                         cancelButtonTitle:nil
                                         otherButtonTitles:nil];
#pragma clang diagnostic pop
    }
    return wrapper;
}

- (void)corroSetTitle:(NSString *)title {
    [self.native setValue:title forKey:@"title"];
}

- (void)corroAddAction:(NSString *)title {
    [self.titles addObject:title];
    if ([self.native isKindOfClass:[UIAlertController class]]) {
        UIAlertController *ac = (UIAlertController *)self.native;
        [ac addAction:[UIAlertAction actionWithTitle:title
                                              style:UIAlertActionStyleDefault
                                            handler:nil]];
    } else {
#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wdeprecated-declarations"
        [(UIAlertView *)self.native addButtonWithTitle:title];
#pragma clang diagnostic pop
    }
}

@end

// `corroPresentDialog:` on a view controller: present whatever the alert
// wrapper holds. Declared as a category so the backend's message send resolves.
@interface UIViewController (CorroIosPresent)
- (void)corroPresentDialog:(id)dialog;
@end

@implementation UIViewController (CorroIosPresent)

- (void)corroPresentDialog:(id)dialog {
    CorroIosAlert *wrapper = (CorroIosAlert *)dialog;
    if (wrapper == nil) {
        return;
    }
    if ([wrapper.native isKindOfClass:[UIAlertController class]]) {
        [self presentViewController:(UIAlertController *)wrapper.native animated:YES completion:nil];
    } else {
#pragma clang diagnostic push
#pragma clang diagnostic ignored "-Wdeprecated-declarations"
        [(UIAlertView *)wrapper.native show];
#pragma clang diagnostic pop
    }
}

@end
