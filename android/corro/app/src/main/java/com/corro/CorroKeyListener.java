package com.corro;

import android.view.KeyEvent;
import android.view.View;

/**
 * Fallback Enter bridge for hardware/adb-injected keys: `adb shell input
 * keyevent` and physical keyboards deliver raw key events to the focused
 * view, which never reach `OnEditorActionListener`. Firing the same native
 * activate callback keeps commit behavior identical across soft keyboard,
 * hardware keys, and automation.
 */
public class CorroKeyListener implements View.OnKeyListener {
    private final long viewPtr;

    public CorroKeyListener(long viewPtr) {
        this.viewPtr = viewPtr;
    }

    @Override
    public boolean onKey(View v, int keyCode, KeyEvent event) {
        if (keyCode == KeyEvent.KEYCODE_ENTER
                && event.getAction() == KeyEvent.ACTION_DOWN) {
            nativeEntryActivate(viewPtr);
            return true;
        }
        return false;
    }

    private static native void nativeEntryActivate(long viewPtr);
}
