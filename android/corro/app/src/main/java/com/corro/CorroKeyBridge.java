package com.corro;

import android.view.KeyEvent;
import android.view.View;

/**
 * The single funnel every key event takes on its way to corro.
 *
 * <h3>Why this class exists</h3>
 * Android delivers a key to the <em>focused</em> view, and when nothing has
 * focus it walks up to the Activity. corro's key handling lives in one
 * place — {@code gui_backend::handle_key} — shared with the GTK and NWG
 * builds, and it was completely unreachable here: the adapter's
 * {@code dispatch_canvas_key} had a live registry that nothing ever called.
 * The only key that reached corro was Enter, via the IME editor-action
 * listener. So a spreadsheet on a phone could be driven only by tapping
 * cells and using the menu strip: no arrow keys, no Shift+arrows, no Escape,
 * no PageUp/PageDown, no Ctrl+S.
 *
 * <h3>Sources</h3>
 * Three, in the order Android prefers them:
 * <ol>
 *   <li>{@link SheetView#onKeyDown} — hardware keyboard, {@code adb shell
 *       input keyevent}, and the soft keyboard's own arrow/delete keys,
 *       which Android delivers as key events to the focused view when the IME
 *       does not consume them.</li>
 *   <li>{@link MainActivity#onKeyDown} — the Activity-level fallback, so a
 *       key no view claims still reaches corro. This is where Ctrl+Q lands
 *       with the soft keyboard up.</li>
 *   <li>{@link CorroKeyListener#onKey} — a focused formula entry, which must
 *       see the key before the grid does.</li>
 * </ol>
 *
 * <h3>What is <em>not</em> here</h3>
 * Translation. The Android keycode is converted to the GDK keysym and the
 * modifier state to a GdkModifierType bitmask in Rust
 * ({@code backends::android::android_keycode_to_gdk}), so the shared handler
 * needs no Android arm of its own and there is exactly one table to keep
 * correct rather than one per platform.
 */
public final class CorroKeyBridge {
    private CorroKeyBridge() {
    }

    /**
     * Offer a key to corro.
     *
     * @param view     the view the event came from, used only for the
     *                 long-press/right-click fallback; may be null
     * @param canvasId the canvas that should see the key, or 0 for none
     * @return true when corro consumed the key
     */
    public static boolean dispatch(View view, KeyEvent event, long canvasId) {
        if (event == null) {
            return false;
        }
        // Only the press acts. Android delivers a press and a release for
        // every key, and feeding both would move the cursor twice and repeat
        // every undo.
        if (event.getAction() != KeyEvent.ACTION_DOWN) {
            return false;
        }
        // A focused entry is routed to corro's entry handler rather than the
        // grid's, which is what makes arrow keys move a caret inside a cell
        // edit instead of moving the cell cursor out from under it.
        long viewPtr = view == null ? 0L : viewPointer(view);
        return nativeKey(event.getKeyCode(), event.getMetaState(), canvasId, viewPtr);
    }

    /**
     * The Rust-side handle of a view, tagged onto the view itself.
     *
     * <p>JNI cannot read a Rust field off a Java object, so the mapping is
     * stored where both sides can see it: a tag id carrying the pointer.
     * Views that are not corro widgets (the menu strip, a dialog button)
     * return 0, which the Rust side reads as "no entry".
     */
    private static long viewPointer(View view) {
        Object tag = view.getTag();
        if (tag instanceof Long) {
            return (Long) tag;
        }
        return 0L;
    }

    /**
     * Tag a view with the Rust handle the adapter was constructed from.
     *
     * <p>Called from Rust when an entry gets a key handler, so only entries
     * that actually registered one pay the tag.
     */
    public static void tagViewPointer(View view, long ptr) {
        if (view != null) {
            view.setTag(ptr);
        }
    }

    private static native boolean nativeKey(int keyCode, int metaMask, long canvasId,
                                           long viewPtr);
}
