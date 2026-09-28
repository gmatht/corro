package com.corro;

import android.view.View;

/**
 * {@code View.OnScrollChangeListener} bridge for a scrolled window.
 *
 * <p>Reports the new offset on the axis that actually moved, which is the
 * adapter's {@code on_scroll(bool vertical, f64)} contract. The choice is
 * made here rather than in Rust because the platform hands back both axes: a
 * horizontal fling also produces a non-zero vertical delta (overscroll,
 * sub-pixel rounding), so "report whichever is non-zero" would make a
 * horizontal scroll look like a vertical one.
 */
public class CorroScrolled implements View.OnScrollChangeListener {
    private final long windowPtr;

    public CorroScrolled(long windowPtr) {
        this.windowPtr = windowPtr;
    }

    @Override
    public void onScrollChange(View v, int scrollX, int scrollY, int oldScrollX, int oldScrollY) {
        // Without the deltas the moved axis has to be inferred: whichever
        // offset differs from the previous one is the one that moved. Ties
        // (a diagonal fling) go to the vertical axis, which is the one a
        // spreadsheet scrolls and the one a caller watching a row count
        // cares about.
        int dx = scrollX - oldScrollX;
        int dy = scrollY - oldScrollY;
        boolean vertical = Math.abs(dy) >= Math.abs(dx);
        nativeScrolled(windowPtr, vertical ? 1 : 0, vertical ? scrollY : scrollX);
    }

    private static native void nativeScrolled(long windowPtr, int vertical, int value);
}
