package com.corro;

import android.app.Activity;
import android.os.Bundle;
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
}
