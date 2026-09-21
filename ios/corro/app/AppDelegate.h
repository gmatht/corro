//
//  AppDelegate.h
//  corro on iOS
//
//  The scene entry point. iOS 13+ goes through a scene delegate; the app
//  delegate is kept for the pre-scene path (iOS 12 and earlier) and for
//  application-level lifecycle.
//

#import <UIKit/UIKit.h>

/// Posted when the app becomes active again: the sheet must redraw, because
/// the framebuffer is not preserved across a suspend on every device.
extern NSNotificationName const CorroShouldRedrawNotification;

/// Posted just before the app is suspended.
extern NSNotificationName const CorroWillSuspendNotification;

@interface AppDelegate : UIResponder <UIApplicationDelegate>

/// Set on the pre-scene path (iOS 12 and earlier); the scene delegate owns the
/// window on iOS 13+.
@property (nonatomic, strong) UIWindow *window;

@end
