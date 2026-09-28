package com.corro;

import android.view.View;

/**
 * {@code View.OnScrollChangeListener} bridge for a scrolled window.
 *
 * <p>Reports the new offset on the axis that actually moved, which is the
 * adapter's {@code on_scroll(bool vertical, f64)} contract. The choice is
 * made here rather than in Rust because the platform hands back both axes:
 * a horizontal fling also produces a non-zero vertical delta (overscroll,
 * sub-pixel rounding), so "report whichever is non-zero" would make a
 * horizontal scroll look like a vertical one.
 */
public class CorroScrolled implements View.OnScrollChangeListener {
    private final long windowPtr;

    public CorroScrolled(long windowPtr) {
        this.windowPtr = windowPtr;
    }

    @Override
    public void onScroll(View v, int scrollX, int scrollY, View oldView, int oldScrollX,
                         int oldScrollY, int dx, int dy) {
        boolean vertical = Math.abs(dy) >= Math.abs(dx);
        nativeScrolled(windowPtr, vertical ? 1 : 0, vertical ? scrollY : scrollX);
    }

    private static native void nativeScrolled(long windowPtr, int vertical, int value);
}
