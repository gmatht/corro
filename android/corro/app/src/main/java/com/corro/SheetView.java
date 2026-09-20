package com.corro;

import android.content.Context;
import android.graphics.Canvas;
import android.view.MotionEvent;
import android.view.View;

/**
 * Spreadsheet grid view. All rendering happens in Rust: onDraw funnels the
 * live android.graphics.Canvas back into the registered Rust draw closure
 * (rswidgets draw callback), which replays DrawContext primitives through
 * JNI. Taps funnel back as canvas clicks (cursor moves + redraw).
 */
public class SheetView extends View {
    private final long canvasId;

    public SheetView(Context context, long canvasId) {
        super(context);
        this.canvasId = canvasId;
        setFocusable(true);
        setFocusableInTouchMode(true);
    }

    @Override
    protected void onDraw(Canvas canvas) {
        super.onDraw(canvas);
        nativeOnDraw(canvasId, canvas, getWidth(), getHeight());
    }

    @Override
    public boolean onTouchEvent(MotionEvent event) {
        if (event.getAction() == MotionEvent.ACTION_DOWN) {
            requestFocus();
            nativeOnTouch(canvasId, event.getX(), event.getY());
            return true;
        }
        return super.onTouchEvent(event);
    }

    private static native void nativeOnDraw(long canvasId, Canvas canvas, int w, int h);
    private static native void nativeOnTouch(long canvasId, float x, float y);
}
