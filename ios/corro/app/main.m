//
//  main.m
//  corro on iOS — the process entry point.
//
//  A UIKit app needs a `main` (without one the link fails with
//  `"_main", referenced from ...`). `UIApplicationMain` never returns: it
//  creates the application object, loads the Info.plist scene manifest and
//  hands control to the app/scene delegates, which is where the Rust side is
//  bootstrapped.
//
//  The delegate class name is passed explicitly rather than inferred from
//  the plist, so the pre-scene path (iOS 12) and the scene path (13+) both
//  work from this one entry point.
//

#import <UIKit/UIKit.h>
#import <stdio.h>
#import "AppDelegate.h"

int main(int argc, char *argv[]) {
    int rc = 0;
    @autoreleasepool {
        fprintf(stderr, "[corro] main: entering UIApplicationMain\n");
        fflush(stderr);
        // UIApplicationMain is documented never to return for a normal app. If
        // it DOES return, the process has been told to exit and the app will
        // disappear right after a clean startup - which is exactly what the
        // simulator showed. Logging both sides makes that unambiguous.
        rc = UIApplicationMain(argc, argv, nil, NSStringFromClass([AppDelegate class]));
        fprintf(stderr, "[corro] main: UIApplicationMain RETURNED %d - the app is exiting\n", rc);
        fflush(stderr);
    }
    return rc;
}
