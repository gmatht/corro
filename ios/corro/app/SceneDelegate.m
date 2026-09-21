//
//  SceneDelegate.m
//  corro on iOS — scene lifecycle (iOS 13+).
//
//  Owns the window and wraps the root view controller in a navigation
//  controller, which is where the menu bar button lives (see
//  CorroViewController.m).
//

#import "SceneDelegate.h"
#import "CorroViewController.h"

@implementation SceneDelegate

- (void)scene:(UIScene *)scene
willConnectToSession:(UISceneSession *)session
      options:(UISceneConnectionOptions *)connectionOptions {
    (void)session;
    (void)connectionOptions;
    if (![scene isKindOfClass:[UIWindowScene class]]) {
        return;
    }
    UIWindowScene *windowScene = (UIWindowScene *)scene;
    self.window = [[UIWindow alloc] initWithFrame:windowScene.coordinateSpace.bounds];
    self.window.windowScene = windowScene;

    CorroViewController *root = [[CorroViewController alloc] init];
    UINavigationController *nav = [[UINavigationController alloc] initWithRootViewController:root];
    // The sheet is a full-screen grid, so the bar stays out of the way: it is
    // small and the tree below it is what matters.
    nav.navigationBarHidden = NO;
    self.window.rootViewController = nav;
    [self.window makeKeyAndVisible];
}

@end
