package com.corro;

import android.app.Activity;
import android.os.Bundle;
import android.content.Intent;
import android.view.KeyEvent;
import android.widget.LinearLayout;

public class MainActivity extends Activity {
    static {
        System.loadLibrary("corro_android");
    }

    // Pass the activity AND its content view to Rust
    private static native void nativeInit(Activity activity, LinearLayout rootLayout);

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        // Root layout that Rust populates (formula bar, sheet, hints)
        LinearLayout root = new LinearLayout(this);
        root.setOrientation(LinearLayout.VERTICAL);
        setContentView(root);

        // Bind the Storage Access Framework picker to this Activity before
        // Rust boots, so a File menu action taken the moment the sheet
        // appears already has somewhere to start an intent from.
        CorroFile.attach(this);

        // Boot corro: backend init + spreadsheet UI
        nativeInit(this, root);
    }

    /**
     * Forward the document picker's result.
     *
     * <p>SAF is asynchronous: {@code startActivityForResult} returns
     * immediately and the chosen {@code content://} URI arrives here, long
     * after Rust's file dialog has blocked waiting for it. Without this
     * forward the result is dropped and every Open and Save As reports
     * "cancelled".
     *
     * <p>Forwarded unconditionally rather than filtered on the request code
     * here: {@link CorroFile#onActivityResult} checks the code, and keeping
     * the check in one place means a second picker with a different code
     * needs no second hook here.
     */
    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        CorroFile.onActivityResult(requestCode, resultCode, data);
    }

    /**
     * Open a document the user handed us from outside (a VIEW/SEND intent).
     *
     * <p>What a file manager's "Open with corro" produces. Without this the
     * intent is dropped, which on Android looks identical to the app not
     * being installed as a handler for spreadsheets at all.
     */
    @Override
    protected void onNewIntent(Intent intent) {
        super.onNewIntent(intent);
        setIntent(intent);
        CorroFile.deliverExternalIntent(intent);
    }

    /**
     * The last resort in the key chain.
     *
     * <p>Android delivers a key to the focused view and then, if no view
     * claims it, to the Activity. corro already handles its keys in
     * {@link SheetView#onKeyDown} and {@link CorroKeyListener#onKey}, but
     * there is a gap this closes: when the soft keyboard is up and the
     * formula entry holds focus, Ctrl+Q / Ctrl+S / Ctrl+Z reach <em>neither</em>
     * — the entry deliberately lets control chords through, and the sheet does
     * not have focus. The window-level accelerators live in the same handler,
     * so a key that arrives here is offered to it before the platform gets
     * it.
     *
     * <p>Returning false when corro declines keeps Back, Home and the volume
     * keys working: they are simply keys it does not want.
     */
    @Override
    public boolean onKeyDown(int keyCode, KeyEvent event) {
        if (CorroKeyBridge.dispatch(getCurrentFocus(), event, 0L)) {
            return true;
        }
        return super.onKeyDown(keyCode, event);
    }
}
