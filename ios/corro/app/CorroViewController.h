//
//  CorroViewController.h
//  corro on iOS — the root view controller.
//
//  Its view *is* the widget tree's root: the Rust backend is handed this view
//  and appends every widget to it (see rustxWidgets/docs/IOS_GUIDELINES.md §3).
//

#import <UIKit/UIKit.h>

@interface CorroViewController : UIViewController

/// The view Rust built into. Kept as a typed property so the shims can add
/// the menu button and reach the canvas without searching the hierarchy.
@property (nonatomic, strong) UIView *rootView;

@end
