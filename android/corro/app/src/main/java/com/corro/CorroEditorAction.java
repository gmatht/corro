package com.corro;

import android.view.KeyEvent;
import android.view.inputmethod.EditorInfo;
import android.widget.TextView;

/**
 * Bridges the IME "Done"/Enter key to Rust: fires the registered Rust
 * `connect_activate` callback, which corro routes into `handle_key(RETURN)`
 * so committing from the soft keyboard behaves exactly like a hardware
 * Enter (commit cell, move down).
 */
public class CorroEditorAction implements TextView.OnEditorActionListener {
    private final long viewPtr;

    public CorroEditorAction(long viewPtr) {
        this.viewPtr = viewPtr;
    }

    @Override
    public boolean onEditorAction(TextView v, int actionId, KeyEvent event) {
        boolean isEnter = actionId == EditorInfo.IME_ACTION_DONE
                || actionId == EditorInfo.IME_ACTION_GO
                || actionId == EditorInfo.IME_ACTION_SEND
                || actionId == EditorInfo.IME_NULL
                || (event != null
                    && event.getKeyCode() == KeyEvent.KEYCODE_ENTER
                    && event.getAction() == KeyEvent.ACTION_DOWN);
        if (isEnter) {
            nativeEntryActivate(viewPtr);
            return true;
        }
        return false;
    }

    private static native void nativeEntryActivate(long viewPtr);
}
