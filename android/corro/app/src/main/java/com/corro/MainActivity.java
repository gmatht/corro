package com.corro;

import android.app.Activity;
import android.os.Bundle;
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

        // Boot corro: backend init + spreadsheet UI
        nativeInit(this, root);
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
