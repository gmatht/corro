package com.corro;

import android.content.Context;
import android.graphics.Canvas;
import android.view.GestureDetector;
import android.view.MotionEvent;
import android.view.ScaleGestureDetector;
import android.view.View;

/**
 * Spreadsheet grid view. All rendering happens in Rust: onDraw funnels the
 * live android.graphics.Canvas back into the registered Rust draw closure
 * (rswidgets draw callback), which replays DrawContext primitives through
 * JNI.
 *
 * <h3>Gestures</h3>
 * This class only classifies the raw touch stream and reports the platform
 * facts it alone can see (is this a finger? is a second pointer down?). What a
 * gesture <em>means</em> is Rust's decision, so the phone and desktop rules
 * live in one place ({@code rswidgets::gridview}, mirrored for corro's live
 * sheet in {@code gui_backend::mobile_gesture}). One gesture means different
 * things on different pointers, which is the point:
 * <ul>
 *   <li><b>Pinch</b> — the sheet's view scale, exactly like a photo. It scales
 *       the grid, the gutters and the hit targets together, so a tap after a
 *       pinch still lands on the cell under the finger.</li>
 *   <li><b>Finger drag</b> — pans the sheet. A phone has no room for a
 *       scrollbar, and a drag is the only natural way to move a grid.</li>
 *   <li><b>Long press, then drag</b> — selects a range. This is the phone's
 *       substitute for shift-click: Rust arms selection on the long press, and
 *       the drag that follows extends it instead of panning.</li>
 *   <li><b>Tap</b> — moves the cursor to the tapped cell.</li>
 *   <li><b>Mouse / stylus drag</b> — selects, like a desktop spreadsheet
 *       (detected via {@link MotionEvent#getToolType}).</li>
 * </ul>
 */
public class SheetView extends View {
    private final long canvasId;

    /** Row height / default column width in device pixels, from Rust. */
    private float cellW = 8f;
    private float cellH = 20f;

    /** Sub-cell pixel remainder carried between scroll events, so a slow drag
     *  still scrolls smoothly instead of rounding every event to zero. */
    private float pendingX = 0f;
    private float pendingY = 0f;

    /** Where the previous scroll event was, for the incremental delta. */
    private float lastX = 0f;
    private float lastY = 0f;

    /** True between ACTION_DOWN and ACTION_UP/CANCEL, and false once a pinch
     *  takes over — a two-finger gesture must not also pan or tap. */
    private boolean gestureActive = false;

    private final ScaleGestureDetector scaleDetector;
    private final GestureDetector gestureDetector;

    public SheetView(Context context, long canvasId) {
        super(context);
        this.canvasId = canvasId;
        setFocusable(true);
        setFocusableInTouchMode(true);
        refreshCellSize();

        // Pinch. `onScale` fires repeatedly during a gesture, each time with
        // the span ratio since the previous call, so the accumulated product
        // is the gesture's total zoom — which is exactly what the shared zoom
        // API takes (an incremental factor, not an absolute scale).
        scaleDetector = new ScaleGestureDetector(
                context, new ScaleGestureDetector.SimpleOnScaleGestureListener() {
            @Override
            public boolean onScale(ScaleGestureDetector detector) {
                float factor = detector.getScaleFactor();
                if (!(factor > 0f) || Float.isNaN(factor)) {
                    return true;
                }
                nativeZoom(factor);
                // The metrics include the zoom, so the pixel→cell conversion
                // the scroll path uses has just changed: re-read them, or a
                // drag after a pinch would pan by the old cell size.
                refreshCellSize();
                return true;
            }
        });

        // Long press: the signal that turns a finger drag into a selection.
        // Taps are handled from the motion stream below (they need the same
        // press bookkeeping), not from onSingleTapUp, which would double-report
        // the gesture. A double tap resets the zoom, the phone's equivalent of
        // "reset view".
        gestureDetector = new GestureDetector(
                context, new GestureDetector.SimpleOnGestureListener() {
            /**
             * A long press is ambiguous on a touch screen: it can mean
             * "start a range selection" or "open this cell's menu". Rust owns
             * the gesture state and decides which; this only reports the press
             * and acts on the code that comes back.
             *
             * <p>corro's sheet uses {@code MENU_MODE = true}, so a long press
             * is its right-click. That leaves range selection to a
             * mouse/stylus drag, and on a finger to a long press when
             * {@code MENU_MODE} is false — the two cannot share one gesture,
             * and the menu is the better default for a phone: it is reachable
             * from anywhere, and a selection is still reachable by dragging.
             */
            @Override
            public void onLongPress(MotionEvent e) {
                if (!gestureActive) {
                    return;
                }
                int outcome = nativeGestureLongPress(
                        canvasId, e.getX(), e.getY(), MENU_MODE);
                if (outcome == OUTCOME_LONG_PRESS_MENU) {
                    // The gesture is consumed by the menu: a following release
                    // must not also fire a tap, or the cursor would jump to
                    // the pressed cell *after* the menu already acted.
                    gestureActive = false;
                    nativeCellContextMenu(e.getX(), e.getY());
                }
            }

            @Override
            public boolean onDoubleTap(MotionEvent e) {
                // The sheet is already zoomed or not: resetting is idempotent,
                // so a spurious double tap (two quick taps on cells) is
                // harmless. Refreshing the cached metrics afterwards is not
                // optional — the drag conversion depends on them.
                nativeResetZoom();
                refreshCellSize();
                return true;
            }
        });
        gestureDetector.setIsLongpressEnabled(true);
    }

    /** Ask Rust for the grid metrics so drags convert pixels to cells with
     *  the same numbers the renderer used. Cheap; called on the first layout
     *  and after every zoom (the metrics include the pinch scale). */
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

    /** A finger pans (and long-presses to select); a mouse/stylus drags to
     *  select. Anything not clearly a mouse is treated as a finger, which is
     *  the safe default on a phone. */
    private static boolean isTouchPointer(MotionEvent event) {
        int tool = event.getToolType(0);
        return tool != MotionEvent.TOOL_TYPE_MOUSE;
    }

    /** Outcome codes shared with Rust ({@code rswidgets::gridview::DragOutcome}). */
    private static final int OUTCOME_SCROLL = 2;
    private static final int OUTCOME_TAP = 3;
    /** The long press opened a context menu rather than arming a selection. */
    private static final int OUTCOME_LONG_PRESS_MENU = 5;

    /**
     * Whether a long press opens the cell's context menu.
     *
     * <p>True for corro's sheet: on a phone a long press is the platform's
     * own gesture for "this thing has a menu", and it is the only route to
     * the cell actions a desktop gets from a right-click.
     */
    private static final boolean MENU_MODE = true;

    @Override
    public boolean onTouchEvent(MotionEvent event) {
        // The detectors see every event first: the scale detector owns
        // multi-finger gestures, the gesture detector supplies the long press.
        scaleDetector.onTouchEvent(event);
        gestureDetector.onTouchEvent(event);

        // A pinch is never a drag: while it is in progress the pan/selection
        // path stays out of the way, or a two-finger gesture would also scroll
        // the sheet (the classic two-finger bug).
        if (scaleDetector.isInProgress()) {
            gestureActive = false;
            return true;
        }

        switch (event.getActionMasked()) {
            case MotionEvent.ACTION_DOWN:
                requestFocus();
                lastX = event.getX();
                lastY = event.getY();
                pendingX = 0f;
                pendingY = 0f;
                gestureActive = true;
                // The press origin is the gesture's anchor: the first move is
                // measured from here, so the travel spent getting past the
                // drag slop is not lost.
                nativeGestureDown(canvasId, event.getX(), event.getY(), isTouchPointer(event));
                return true;

            case MotionEvent.ACTION_MOVE: {
                if (!gestureActive) {
                    return true;
                }
                int outcome = nativeGestureMove(canvasId, event.getX(), event.getY());
                if (outcome == OUTCOME_SCROLL) {
                    scrollByDrag(event.getX(), event.getY());
                }
                return true;
            }

            case MotionEvent.ACTION_POINTER_DOWN:
                // A second finger arrived: this is a pinch now. Drop the
                // in-flight one-finger gesture so its release cannot fire a tap
                // (or its move pan the sheet underneath the pinch).
                if (gestureActive) {
                    gestureActive = false;
                    nativeGestureCancel(canvasId);
                }
                return true;

            case MotionEvent.ACTION_UP:
                if (gestureActive) {
                    gestureActive = false;
                    // Rust decides whether this was a tap or the end of a
                    // drag, and *reports* it. A tap is then routed to THIS
                    // canvas's click handler, so the tab strip's taps do not
                    // move the sheet's cursor.
                    if (nativeGestureUp(canvasId, event.getX(), event.getY()) == OUTCOME_TAP) {
                        nativeOnTouch(canvasId, event.getX(), event.getY());
                    }
                }
                pendingX = 0f;
                pendingY = 0f;
                return true;

            case MotionEvent.ACTION_CANCEL:
                if (gestureActive) {
                    gestureActive = false;
                    nativeGestureCancel(canvasId);
                }
                pendingX = 0f;
                pendingY = 0f;
                return true;

            default:
                return super.onTouchEvent(event);
        }
    }

    /**
     * Convert a scroll drag from the last point to (x, y) into whole
     * rows/columns.
     *
     * The conversion happens in Rust ({@link #nativeDragBy}) because only Rust
     * knows the live metrics — a pinch changes them, and a cell size cached
     * here would pan by the stale amount. The sub-cell remainder stays pending
     * in this view, so the mapping never loses motion to rounding.
     *
     * Dragging the content upwards (finger moves up, dy negative) must reveal
     * later rows, so Rust applies the sign; this only accumulates pixels.
     */
    private void scrollByDrag(float x, float y) {
        pendingX += x - lastX;
        pendingY += y - lastY;
        lastX = x;
        lastY = y;

        int[] applied = nativeDragBy(canvasId, pendingX, pendingY);
        if (applied == null || applied.length < 2) {
            return;
        }
        // Keep whatever did not make a whole cell: a slow drag still scrolls.
        pendingX -= applied[1] * cellW;
        pendingY -= applied[0] * cellH;
    }

    /**
     * Hardware keyboard, {@code adb shell input keyevent}, and the soft
     * keyboard's own arrow/delete keys.
     *
     * <p>This is what makes corro's shared {@code handle_key} reachable on a
     * phone. Android sends a key event to the focused view, and the sheet
     * takes focus on {@code ACTION_DOWN} and on the first tap, so it is the
     * view that sees them.
     *
     * <p>Returning true only when corro consumed the key matters: an
     * unconsumed key must reach the platform, or the system Back button and
     * the volume keys would stop working. That is why the *decision* lives in
     * Rust — only it knows whether a key was a cell navigation or something
     * the platform should have.
     */
    @Override
    public boolean onKeyDown(int keyCode, KeyEvent event) {
        if (CorroKeyBridge.dispatch(this, event, canvasId)) {
            return true;
        }
        return super.onKeyDown(keyCode, event);
    }

    /**
     * A secondary-button click, which on Android means a <em>mouse</em> right
     * click: touch has no buttons, but the emulator's pointer and a stylus do
     * report one.
     *
     * <p>Handled here rather than in {@code onTouchEvent} because a
     * {@code MotionEvent} only exposes the button as a modifier on a
     * secondary click, which arrives as {@code ACTION_DOWN} with
     * {@code BUTTON_SECONDARY} in the meta state. The gesture recogniser
     * above would otherwise treat it as a tap and move the cursor, so a
     * right-click opened the cell under the pointer's *previous* position
     * instead of the one clicked.
     */
    @Override
    public boolean onGenericMotionEvent(MotionEvent event) {
        if (event.getAction() == MotionEvent.ACTION_BUTTON_PRESS
                && (event.getButtonState() & MotionEvent.BUTTON_SECONDARY) != 0) {
            nativeClickButton(canvasId, event.getX(), event.getY(), 3 /* right */, 0);
            return true;
        }
        return super.onGenericMotionEvent(event);
    }

    private static native void nativeOnDraw(long canvasId, Canvas canvas, int w, int h);
    private static native void nativeClickButton(long canvasId, float x, float y,
                                                 int button, int mods);
    private static native void nativeOnTouch(long canvasId, float x, float y);
    private static native float[] nativeCellSize();
    private static native float nativeZoom(float factor);
    private static native float nativeResetZoom();
    private static native void nativeGestureDown(long canvasId, float x, float y, boolean isTouch);
    private static native int nativeGestureLongPress(long canvasId, float x, float y,
                                                     boolean menuMode);
    /** Open the cell's context menu, the actions for the cell at (x, y). */
    private static native void nativeCellContextMenu(float x, float y);
    private static native int nativeGestureMove(long canvasId, float x, float y);
    private static native int nativeGestureUp(long canvasId, float x, float y);
    private static native void nativeGestureCancel(long canvasId);
    private static native int[] nativeDragBy(long canvasId, float dx, float dy);
}
