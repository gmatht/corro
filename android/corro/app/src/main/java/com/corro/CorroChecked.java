package com.corro;

import android.widget.CompoundButton;

/**
 * {@code CompoundButton.OnCheckedChangeListener} bridge, shared by
 * {@code CheckBox} and {@code RadioButton}.
 *
 * <p>JNI cannot create a proxy for a Java interface, so every platform
 * listener needs a concrete shim class. Without one, a checkbox on Android
 * could be set programmatically and read back with {@code isChecked}, but a
 * user tap never reached the caller.
 *
 * <p>The two uses are the same shape — a toggle in a dialog, a radio option —
 * so one class serves both, keyed only by the callback id.
 */
public class CorroChecked implements CompoundButton.OnCheckedChangeListener {
    private final long callbackId;

    public CorroChecked(long callbackId) {
        this.callbackId = callbackId;
    }

    @Override
    public void onCheckedChanged(CompoundButton button, boolean isChecked) {
        nativeChecked(callbackId, isChecked);
    }

    private static native void nativeChecked(long callbackId, boolean checked);
}
