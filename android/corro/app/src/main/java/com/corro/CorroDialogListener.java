package com.corro;

import android.content.DialogInterface;

/**
 * {@code DialogInterface.OnClickListener} bridge for a dialog's buttons.
 *
 * <p>Without this, {@code Dialog::add_button} labelled the buttons and
 * {@code Dialog::present} showed them, but nothing reported which one the
 * user pressed — so every {@code connect_response} on Android was a no-op
 * returning 0, and a caller waiting for "OK" or "Cancel" waited forever.
 *
 * <p><b>One listener object for the whole dialog,</b> not one per button.
 * Android's {@code setPositiveButton} and friends take the listener
 * themselves, so a per-button object would be the only way to tell them
 * apart — and then the Rust side would need the object identity, which a
 * {@code DialogInterface} click does not carry. Sharing one object and
 * telling the buttons apart by which slot they occupy is what the response-id
 * table in the adapter is for.
 */
public class CorroDialogListener implements DialogInterface.OnClickListener {
    /** Which button this instance stands for. */
    private final String role;
    private final long builderPtr;

    public CorroDialogListener(long builderPtr, String role) {
        this.builderPtr = builderPtr;
        this.role = role;
    }

    @Override
    public void onClick(DialogInterface dialog, int which) {
        nativeDialogButton(builderPtr, role);
    }

    private static native void nativeDialogButton(long builderPtr, String role);
}
