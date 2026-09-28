package com.corro;

import android.view.View;
import android.view.ViewTreeObserver;

/**
 * {@code ViewTreeObserver.OnGlobalLayoutListener} bridge.
 *
 * <p>Android's "the bounds settled" hook, the platform's counterpart of
 * GTK's {@code size-allocate}. It matters on Android because the layout runs
 * <em>after</em> the widget tree is built: {@code nativeInit} returns before
 * a single pixel exists, so a caller that needs a real width or height has
 * nowhere else to ask. GTK's {@code Window::set_layout_cb} is the same
 * request, and this is the same answer.
 *
 * <p>The listener holds the <em>view</em> pointer rather than a callback id,
 * because a view may register a layout callback exactly once and the Rust
 * registry is keyed by the view.
 */
public class CorroLayout implements ViewTreeObserver.OnGlobalLayoutListener {
    private final long viewPtr;

    public CorroLayout(long viewPtr) {
        this.viewPtr = viewPtr;
    }

    @Override
    public void onGlobalLayout() {
        nativeLayout(viewPtr);
    }

    /**
     * Register on the view's observer. The Rust side calls
     * {@code getViewTreeObserver} and then {@code addOnGlobalLayoutListener}
     * on the result, because {@code View} has no direct
     * {@code setOnGlobalLayoutListener}.
     */
    public static ViewTreeObserver observerFor(View view) {
        return view.getViewTreeObserver();
    }

    private static native void nativeLayout(long viewPtr);
}
