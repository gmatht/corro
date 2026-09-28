package com.corro;

import android.widget.AdapterView;

/**
 * {@code AdapterView.OnItemSelectedListener} bridge.
 *
 * <p>JNI cannot create a proxy for a Java interface, so every platform
 * listener needs a concrete shim class. This is the one for a
 * {@code Spinner}: without it a dropdown on Android could be *set* and read
 * back with {@code getSelectedItemPosition}, but a user pick never reached
 * the caller — every filter combo in a dialog was inert.
 *
 * <p>The position is handed straight to Rust, which owns the callback
 * registry and the `connect_changed` closure.
 */
public class CorroItemSelected implements AdapterView.OnItemSelectedListener {
    private final long callbackId;

    public CorroItemSelected(long callbackId) {
        this.callbackId = callbackId;
    }

    @Override
    public void onItemSelected(AdapterView<?> parent, android.view.View view,
                               int position, long id) {
        // Nothing selected yet (the spinner is still binding its adapter).
        if (position < 0) {
            return;
        }
        nativeItemSelected(callbackId, position);
    }

    @Override
    public void onNothingSelected(AdapterView<?> parent) {
        // Deliberately silent: a transient "nothing selected" fires while the
        // adapter is being rebuilt, and reporting it as a change would make a
        // dialog reset the caller's selection on every open.
    }

    private static native void nativeItemSelected(long callbackId, int position);
}
