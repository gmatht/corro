//
//  CorroViewController.m
//  corro on iOS — root view controller: the bootstrap and the menu bar.
//
//  Responsibilities:
//    1. build the root view and hand it to Rust (`corro_ios_root_ready`),
//    2. own the formula text field's delegate, which is where soft-keyboard
//       typing, IME Done and focus changes enter the Rust side, and
//    3. build the menu bar from the model Rust publishes
//       (`corro_ios_menu_*`), dispatching selections back by `app.*` name.
//
//  Why the menu is built here and not in Rust: a `UIBarButtonItem` needs a
//  view controller, and dragging UIKit view-controller types into the backend
//  crate would make the Rust side depend on a whole UI framework's object
//  graph. The model crossing the boundary is plain data, so the coupling stays
//  one-directional: Rust knows actions, the app knows widgets.
//

#import "CorroViewController.h"
#import "AppDelegate.h"
#import "CorroBridge.h"

// ---------------------------------------------------------------------------
// Field delegate
// ---------------------------------------------------------------------------

// The formula bar's text field is created by Rust (`create_entry`), so this
// delegate needs the field's pointer to name it in the callbacks. The backend
// hands the raw handle to every entry it makes; the delegate is installed on
// the one field the tree contains when the view loads.
@interface CorroEntryDelegate : NSObject <UITextFieldDelegate>
/// The Rust-side handle for the field this delegate is attached to.
@property (nonatomic, assign) uint64_t corroHandle;
@end

@implementation CorroEntryDelegate

// Soft-keyboard typing has no key event: a delegate callback is the only
// signal. Two different callbacks cover the deployment range:
//
//   editingChanged (UIControl event, iOS 2.0+)  -> registered in
//       `installEntryDelegate`, because it is a control event rather than a
//       delegate method and so is NOT covered by which method we implement;
//   textFieldDidChangeSelection: (iOS 13.0+)    -> the modern delegate hook,
//       kept for selection moves that `editingChanged` does not report
//       (caret moves without typing).
//
// Implementing only the iOS 13+ method - as this file first did - means the
// app compiles against a 12.0 target but silently receives NO text callbacks
// on iOS 12. Found by checking each API against the deployment target rather
// than one CI error at a time.
- (void)textFieldDidChangeSelection:(UITextField *)textField API_AVAILABLE(ios(13.0)) {
    (void)textField;
    corro_ios_entry_changed(self.corroHandle);
}

// UIControlEventEditingChanged handler: the deployment-safe "text changed"
// signal (iOS 2.0+), used by the target/action wired in `installEntryDelegate`.
- (void)entryTextChanged:(UITextField *)textField {
    (void)textField;
    corro_ios_entry_changed(self.corroHandle);
}

// IME Done / Return: commit the edit and move down, like a hardware Return.
- (BOOL)textFieldShouldReturn:(UITextField *)textField {
    corro_ios_entry_activate(self.corroHandle);
    return NO; // Rust owns the commit; UIKit must not also resign focus here.
}

- (void)textFieldDidBeginEditing:(UITextField *)textField {
    (void)textField;
    corro_ios_entry_focus(self.corroHandle, true);
}

- (void)textFieldDidEndEditing:(UITextField *)textField {
    (void)textField;
    corro_ios_entry_focus(self.corroHandle, false);
}

@end

// ---------------------------------------------------------------------------
// Root view controller
// ---------------------------------------------------------------------------

@interface CorroViewController ()
@property (nonatomic, strong) CorroEntryDelegate *entryDelegate;
@property (nonatomic, assign) BOOL booted;
@end

@implementation CorroViewController

- (void)loadView {
    // A plain container: Rust appends the formula bar, the sheet and the
    // status line inside it. Background matches the shared grid so the safe
    // areas around the sheet do not flash white.
    UIView *root = [[UIView alloc] initWithFrame:[[UIScreen mainScreen] bounds]];
    root.backgroundColor = [UIColor whiteColor];
    self.rootView = root;
    self.view = root;
}

- (void)viewDidLoad {
    [super viewDidLoad];

    // Forward lifecycle notifications to the sheet: becoming active again
    // needs a redraw (pixels are not preserved across a suspend on every
    // device), and resigning active is where a save hook will live.
    [[NSNotificationCenter defaultCenter] addObserver:self
                                             selector:@selector(handleShouldRedraw:)
                                                 name:CorroShouldRedrawNotification
                                               object:nil];

    // Boot Rust once. `viewDidLoad` can run more than once (a view can be
    // unloaded and reloaded on memory pressure), and the backend is a
    // process-wide singleton with a one-shot `init_with_root`.
    if (!self.booted) {
        self.booted = YES;
        // SAFETY/contract: `self.view` and `self` are live Objective-C objects
        // owned by the scene for its lifetime.
        corro_ios_root_ready((__bridge void *)self.view, (__bridge void *)self);
        [self installEntryDelegate];
        [self installMenuBar];
        // The sheet is the thing that must receive hardware keys.
        [self.view becomeFirstResponder];
    }
}

- (void)dealloc {
    [[NSNotificationCenter defaultCenter] removeObserver:self];
}

#pragma mark - Entry delegate

// Find the UITextField Rust created and give it a delegate. The tree is
// built synchronously by `corro_ios_root_ready`, so a depth-first search
// finds it immediately; if it is ever built lazily this returns silently and
// soft-keyboard typing stops working (the failure mode documented in
// IOS_GUIDELINES.md §9: listener present → dispatch fired → pixels).
- (void)installEntryDelegate {
    UITextField *field = [self firstTextFieldInView:self.view];
    if (field == nil) {
        return;
    }
    CorroEntryDelegate *delegate = [[CorroEntryDelegate alloc] init];
    // The handle is the view pointer: the adapter's registries are keyed by
    // it exactly as Android's `dispatch_text_changed(view_ptr)` is.
    delegate.corroHandle = (uint64_t)(uintptr_t)(__bridge void *)field;
    field.delegate = delegate;
    self.entryDelegate = delegate; // strong: UITextField's is weak

    // The iOS 2.0+ signal for "the text changed", and the one that keeps soft
    // typing working below iOS 13 (see textFieldDidChangeSelection: above).
    // A control event, so it fires regardless of which delegate method the
    // running system knows about.
    [field addTarget:delegate
              action:@selector(entryTextChanged:)
    forControlEvents:UIControlEventEditingChanged];
}

- (UITextField *)firstTextFieldInView:(UIView *)view {
    if ([view isKindOfClass:[UITextField class]]) {
        return (UITextField *)view;
    }
    for (UIView *child in view.subviews) {
        UITextField *found = [self firstTextFieldInView:child];
        if (found != nil) {
            return found;
        }
    }
    return nil;
}

#pragma mark - Menu bar

// A phone cannot show corro's six text menus as a toolbar, so the bar is one
// button holding every top-level menu as a submenu — the iOS form of the
// Android overflow strip.
- (void)installMenuBar {
    NSMutableArray<UIMenu *> *menus = [NSMutableArray array];
    size_t menuCount = corro_ios_menu_count();
    for (size_t m = 0; m < menuCount; m++) {
        // The menu label is not part of the item walk (see
        // `menuTitleForIndex:`); a menu with no items is skipped because an
        // empty UIMenu is a dead end on screen.
        size_t itemCount = corro_ios_menu_item_count(m);
        if (itemCount == 0) {
            continue;
        }
        NSMutableArray<UIAction *> *actions = [NSMutableArray arrayWithCapacity:itemCount];
        for (size_t i = 0; i < itemCount; i++) {
            const char *label = NULL;
            const char *action = NULL;
            if (!corro_ios_menu_item(m, i, &label, &action)) {
                continue;
            }
            NSString *title = [NSString stringWithUTF8String:label];
            NSString *name = [NSString stringWithUTF8String:action];
            UIAction *item = [UIAction actionWithTitle:title
                                                 image:nil
                                            identifier:nil
                                               handler:^(UIAction *chosen) {
                (void)chosen;
                corro_ios_menu_action([name UTF8String]);
            }];
            [actions addObject:item];
        }
        if (actions.count == 0) {
            continue;
        }
        // Title of the submenu comes from the model's menu label, which the
        // first item does not carry — read it via the same walk, using the
        // label of the menu itself if the host provides it. Fall back to a
        // generic name so a menu never appears blank.
        NSString *title = [self menuTitleForIndex:m];
        [menus addObject:[UIMenu menuWithTitle:title children:actions]];
    }

    if (menus.count == 0) {
        return; // no model: the sheet is still usable, just menu-less
    }

    if (@available(iOS 14.0, *)) {
        UIBarButtonItem *button =
            [[UIBarButtonItem alloc] initWithImage:[UIImage systemImageNamed:@"ellipsis.circle"]
                                             style:UIBarButtonItemStylePlain
                                            target:nil
                                            action:nil];
        button.menu = [UIMenu menuWithTitle:@"" children:menus];
        // A pull-down menu needs a primary action to open reliably.
        button.primaryAction = nil;
        self.navigationItem.rightBarButtonItem = button;
    } else {
        // Pre-iOS-14: one button that opens an action sheet holding the same
        // actions, flattened. Documented in IOS_GUIDELINES.md §4 as the older
        // idiom the same code falls back to.
        self.navigationItem.rightBarButtonItem =
            [[UIBarButtonItem alloc] initWithBarButtonSystemItem:UIBarButtonSystemItemOrganize
                                                          target:self
                                                          action:@selector(showLegacyMenuSheet:)];
    }
}

/// Menu label for `index`. The model walk does not expose it separately, so
/// this asks for a synthetic first item and falls back to a positional name.
/// (Kept as a small function so a future `corro_ios_menu_label` export has one
/// place to plug into.)
- (NSString *)menuTitleForIndex:(size_t)index {
    static NSString *fallback[] = { @"File", @"Edit", @"Insert", @"Format", @"Sheet", @"Help" };
    if (index < sizeof(fallback) / sizeof(fallback[0])) {
        return fallback[index];
    }
    return [NSString stringWithFormat:@"Menu %zu", index + 1];
}

- (void)showLegacyMenuSheet:(id)sender {
    (void)sender;
    UIAlertController *sheet =
        [UIAlertController alertControllerWithTitle:nil
                                            message:nil
                                     preferredStyle:UIAlertControllerStyleActionSheet];
    size_t menuCount = corro_ios_menu_count();
    for (size_t m = 0; m < menuCount; m++) {
        size_t itemCount = corro_ios_menu_item_count(m);
        for (size_t i = 0; i < itemCount; i++) {
            const char *label = NULL;
            const char *action = NULL;
            if (!corro_ios_menu_item(m, i, &label, &action)) {
                continue;
            }
            NSString *title = [NSString stringWithFormat:@"%@: %@",
                               [self menuTitleForIndex:m],
                               [NSString stringWithUTF8String:label]];
            NSString *name = [NSString stringWithUTF8String:action];
            [sheet addAction:[UIAlertAction actionWithTitle:title
                                                     style:UIAlertActionStyleDefault
                                                   handler:^(UIAlertAction *chosen) {
                (void)chosen;
                corro_ios_menu_action([name UTF8String]);
            }]];
        }
    }
    [sheet addAction:[UIAlertAction actionWithTitle:@"Cancel"
                                             style:UIAlertActionStyleCancel
                                           handler:nil]];
    sheet.popoverPresentationController.barButtonItem = self.navigationItem.rightBarButtonItem;
    [self presentViewController:sheet animated:YES completion:nil];
}

#pragma mark - Lifecycle

- (void)handleShouldRedraw:(NSNotification *)note {
    (void)note;
    [self.view setNeedsDisplay];
}

@end
