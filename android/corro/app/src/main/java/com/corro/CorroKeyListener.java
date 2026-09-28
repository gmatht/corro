package com.corro;

import android.view.KeyEvent;
import android.view.View;

/**
 * Key bridge for a focused {@code EditText} — corro's formula entry.
 *
 * <h3>What it used to do</h3>
 * One thing: forward Enter, because {@code adb shell input keyevent} and
 * physical keyboards deliver a raw key event that never becomes an
 * {@code OnEditorActionListener} callback, so a hardware Enter had to be
 * turned into the same native activate call as the IME's.
 *
 * <h3>What it does now</h3>
 * Offers <em>every</em> key to corro first, then keeps the Enter fallback.
 * That ordering is the point: an arrow key, a Backspace or a Delete inside a
 * cell edit has to be applied by corro's edit model — it owns the buffer, the
 * formula-aware delete-at-caret rules and the caret position — rather than by
 * the platform's own {@code Editable}. Returning false for keys corro
 * declines leaves ordinary typing to the IME, which is where it should be:
 * soft typing arrives as text through the {@code TextWatcher}, and routing
 * character keys here as well would apply each keystroke twice.
 */
public class CorroKeyListener implements View.OnKeyListener {
    private final long viewPtr;

    public CorroKeyListener(long viewPtr) {
        this.viewPtr = viewPtr;
    }

    @Override
    public boolean onKey(View v, int keyCode, KeyEvent event) {
        // Let corro decide. It knows whether a key is a cell-navigation
        // action, an edit action, or something the platform should have.
        if (CorroKeyBridge.dispatch(v, event, 0L)) {
            return true;
        }
        if (keyCode == KeyEvent.KEYCODE_ENTER
                && event.getAction() == KeyEvent.ACTION_DOWN) {
            nativeEntryActivate(viewPtr);
            return true;
        }
        return false;
    }

    private static native void nativeEntryActivate(long viewPtr);
}
