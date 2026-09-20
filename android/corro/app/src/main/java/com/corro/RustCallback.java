package com.corro;

import android.view.View;

/**
 * Bridge class that Android calls when a user taps a View.
 * The native method dispatchCallback is served by Rust's callback
 * registry (rswidgets::backends::android).
 */
public class RustCallback implements View.OnClickListener {
    private final long callbackId;

    public RustCallback(long id) {
        this.callbackId = id;
    }

    @Override
    public void onClick(View v) {
        nativeDispatchCallback(callbackId);
    }

    private static native void nativeDispatchCallback(long id);
}
