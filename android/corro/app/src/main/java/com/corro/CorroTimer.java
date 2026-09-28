package com.corro;

/**
 * {@link Runnable} bridge for {@code rswidgets}'s timers.
 *
 * <p>There is no event loop on Android to run a source on, so
 * {@code backends::android::schedule_timeout} posts one of these through a
 * {@link android.os.Handler} and hands it the timer id. A repeating timer
 * re-posts itself on the Rust side, so this class stays a plain one-shot
 * {@code run()} that forwards and returns.
 *
 * <p>The native call is the same dispatch for every timer; the id is what
 * tells Rust which one fired.
 */
public class CorroTimer implements Runnable {
    private final long timerId;

    public CorroTimer(long timerId) {
        this.timerId = timerId;
    }

    @Override
    public void run() {
        nativeFire(timerId);
    }

    private static native void nativeFire(long timerId);
}
