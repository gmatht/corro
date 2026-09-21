package com.corro;

import android.content.Context;
import android.graphics.Canvas;
import android.view.MotionEvent;
import android.view.View;
import android.view.ViewConfiguration;

/**
 * Spreadsheet grid view. All rendering happens in Rust: onDraw funnels the
 * live android.graphics.Canvas back into the registered Rust draw closure
 * (rswidgets draw callback), which replays DrawContext primitives through
 * JNI.
 *
 * Touch handling, likewise, only decides *what* gesture happened and hands
 * whole-cell counts to Rust:
 * <ul>
 *   <li>a tap moves the cursor to the tapped cell;</li>
 *   <li>a drag scrolls the sheet, converted from pixels to rows/columns
 *       using the metrics Rust renders with ({@link #nativeCellSize}).</li>
 * </ul>
 * Android has no native scrolling here (the sheet is drawn by Rust into a
 * plain Canvas, not a ScrollView), so the drag has to be turned into viewport
 * movement explicitly.
 */
public class SheetView extends View {
    private final long canvasId;

    /** Row height / default column width in device pixels, from Rust. */
    private float cellW = 8f;
    private float cellH = 20f;

    /** Sub-cell pixel remainder carried between drag events, so a slow drag
     *  still scrolls smoothly instead of rounding every event to zero. */
    private float pendingX = 0f;
    private float pendingY = 0f;

    private float lastX = 0f;
    private float lastY = 0f;

    /** Where the gesture started (ACTION_DOWN), so drag distance is measured
     *  from the true origin rather than from wherever slop was exceeded. */
    private float downX = 0f;
    private float downY = 0f;

    /** True once the current gesture has moved far enough to be a drag (as
     *  opposed to a tap). Until then the touch may still turn out to be a
     *  tap, so the cursor must not move yet. */
    private boolean dragging = false;

    private final int touchSlop;

    public SheetView(Context context, long canvasId) {
        super(context);
        this.canvasId = canvasId;
        setFocusable(true);
        setFocusableInTouchMode(true);
        touchSlop = ViewConfiguration.get(context).getScaledTouchSlop();
        refreshCellSize();
    }

    /** Ask Rust for the grid metrics so drags convert pixels to cells with
     *  the same numbers the renderer used. Cheap; called once per layout in
     *  case the density or font metrics differ from the defaults. */
    private void refreshCellSize() {
        float[] size = nativeCellSize();
        if (size != null && size.length >= 2 && size[0] > 0f && size[1] > 0f) {
            cellH = size[0];
            cellW = size[1];
        }
    }

    @Override
    protected void onSizeChanged(int w, int h, int oldw, int oldh) {
        super.onSizeChanged(w, h, oldw, oldh);
        refreshCellSize();
    }

    @Override
    protected void onDraw(Canvas canvas) {
        super.onDraw(canvas);
        nativeOnDraw(canvasId, canvas, getWidth(), getHeight());
    }

    @Override
    public boolean onTouchEvent(MotionEvent event) {
        switch (event.getActionMasked()) {
            case MotionEvent.ACTION_DOWN:
                requestFocus();
                downX = event.getX();
                downY = event.getY();
                lastX = downX;
                lastY = downY;
                pendingX = 0f;
                pendingY = 0f;
                dragging = false;
                return true;

            case MotionEvent.ACTION_MOVE: {
                if (!dragging) {
                    // Not yet a drag: only promote once the finger has
                    // travelled past touch slop, so a slightly shaky tap still
                    // selects a cell instead of scrolling by a row.
                    if (Math.hypot(event.getX() - downX, event.getY() - downY) > touchSlop) {
                        dragging = true;
                        // Scrolling starts from where the *gesture* began, not
                        // from where slop was exceeded: the travel already
                        // spent getting past slop is real drag distance, and
                        // discarding it loses the first row of every gesture.
                        lastX = downX;
                        lastY = downY;
                    } else {
                        return true;
                    }
                }
                scrollByDrag(event.getX(), event.getY());
                return true;
            }

            case MotionEvent.ACTION_UP:
                if (!dragging) {
                    nativeOnTouch(canvasId, event.getX(), event.getY());
                }
                dragging = false;
                pendingX = 0f;
                pendingY = 0f;
                return true;

            case MotionEvent.ACTION_CANCEL:
                dragging = false;
                pendingX = 0f;
                pendingY = 0f;
                return true;

            default:
                return super.onTouchEvent(event);
        }
    }

    /**
     * Convert a drag from the last point to (x, y) into whole rows/columns
     * and hand them to Rust.
     *
     * Dragging the content upwards (finger moves up, dy negative) must reveal
     * later rows, so the row delta is the negated pixel delta: dragging up
     * scrolls the sheet down. The sub-cell remainder stays pending in
     * this view, so the mapping never loses motion to rounding.
     */
    private void scrollByDrag(float x, float y) {
        float dx = x - lastX;
        float dy = y - lastY;
        lastX = x;
        lastY = y;

        pendingX += -dx;
        pendingY += -dy;

        int dCols = 0;
        int dRows = 0;
        if (cellW > 0f) {
            dCols = (int) (pendingX / cellW);
            pendingX -= dCols * cellW;
        }
        if (cellH > 0f) {
            dRows = (int) (pendingY / cellH);
            pendingY -= dRows * cellH;
        }
        if (dRows != 0 || dCols != 0) {
            nativeScrollBy(dRows, dCols);
        }
    }

    private static native void nativeOnDraw(long canvasId, Canvas canvas, int w, int h);
    private static native void nativeOnTouch(long canvasId, float x, float y);
    private static native void nativeScrollBy(int dRows, int dCols);
    private static native float[] nativeCellSize();
}
