//
//  AppDelegate.m
//  corro on iOS — application + scene lifecycle.
//
//  Two responsibilities:
//    1. own the window and root view controller (scene path on iOS 13+, the
//       legacy path before that), and
//    2. forward lifecycle transitions to the Rust side, which must not keep
//       running while the app is suspended.
//
//  The root view controller is the bootstrap point: when its view loads it
//  calls `corro_ios_root_ready`, which boots the whole Rust GUI inside the
//  app's own view (see rustxWidgets/docs/IOS_GUIDELINES.md §3).
//

#import "AppDelegate.h"
#import "CorroViewController.h"
#import <stdio.h>

/// Posted when the app becomes active again. The root view controller listens
/// and asks the sheet to redraw, because the framebuffer is not preserved
/// across a suspend on every device.
NSNotificationName const CorroShouldRedrawNotification = @"CorroShouldRedraw";

/// Posted just before the app is suspended: a place for a future save hook.
NSNotificationName const CorroWillSuspendNotification = @"CorroWillSuspend";

@implementation AppDelegate

#pragma mark - Legacy (pre-scene, iOS 12 and earlier) window setup

// On iOS 13+ the scene delegate supplies the window and this never runs; the
// delegate still has to keep this method for the older deployment targets.
- (BOOL)application:(UIApplication *)application
didFinishLaunchingWithOptions:(NSDictionary<UIApplicationLaunchOptionsKey, id> *)launchOptions {
    fprintf(stderr, "[corro] application:didFinishLaunchingWithOptions\n");
    fflush(stderr);
    (void)application;
    (void)launchOptions;
    if (@available(iOS 13.0, *)) {
        // Scene-based: the scene delegate builds the window in
        // `scene:willConnectToSession:options:`.
        return YES;
    }
    self.window = [[UIWindow alloc] initWithFrame:[[UIScreen mainScreen] bounds]];
    CorroViewController *root = [[CorroViewController alloc] init];
    self.window.rootViewController = root;
    [self.window makeKeyAndVisible];
    return YES;
}

#pragma mark - Lifecycle forwarding

// The backend runs no background work of its own, but a suspended app must not
// have a timer or a draw loop armed. Forwarding these makes that explicit and
// gives the shims a place to pause, and the way back a reason to redraw.
- (void)applicationDidBecomeActive:(UIApplication *)application {
    (void)application;
    [[NSNotificationCenter defaultCenter] postNotificationName:CorroShouldRedrawNotification
                                                        object:nil];
}

- (void)applicationWillResignActive:(UIApplication *)application {
    (void)application;
    [[NSNotificationCenter defaultCenter] postNotificationName:CorroWillSuspendNotification
                                                        object:nil];
}

- (void)applicationWillTerminate:(UIApplication *)application {
    (void)application;
    // iOS gives no time for real work here: this exists so a future save hook
    // has a place to live, not to run one now. The log is diagnostic: if the
    // app disappears after a clean startup, this line says whether UIKit was
    // tearing it down or the process was reaped without any callback.
    fprintf(stderr, "[corro] applicationWillTerminate\n");
    fflush(stderr);
}

@end
