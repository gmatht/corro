package com.corro;

import android.text.Editable;
import android.text.TextWatcher;

/**
 * Bridges EditText text changes to Rust: every change calls
 * nativeEntryChanged(viewPtr), which runs the registered Rust
 * `connect_changed` callback (corro's formula-entry sync).
 */
public class CorroTextWatcher implements TextWatcher {
    private final long viewPtr;

    public CorroTextWatcher(long viewPtr) {
        this.viewPtr = viewPtr;
    }

    @Override
    public void beforeTextChanged(CharSequence s, int start, int count, int after) {}

    @Override
    public void onTextChanged(CharSequence s, int start, int before, int count) {}

    @Override
    public void afterTextChanged(Editable s) {
        nativeEntryChanged(viewPtr);
    }

    private static native void nativeEntryChanged(long viewPtr);
}
