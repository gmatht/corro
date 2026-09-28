package com.corro;

/**
 * The Android system clipboard ({@code ClipboardManager}).
 *
 * <h3>Why this is not already covered</h3>
 * corro's GTK and NWG <em>GUI</em> backends keep copy/cut/paste in in-process
 * Rust state and never touch a platform clipboard. The two backends that do
 * reach a system clipboard are the TUI (which shells out to {@code xclip} /
 * {@code pbcopy} / {@code clip}) and the pancurses backend (OSC 52). Neither
 * can apply on a phone — there is no {@code xclip} and no terminal — so the
 * supported Android equivalent is {@code ClipboardManager}.
 *
 * <h3>Why it is plain text and not a MIME type</h3>
 * A spreadsheet cell is text, and plain text is the one representation every
 * other app on the device can read. Android's clipboard is a
 * {@code ClipData}, so the choice is between {@code text/plain} and a
 * private type only corro understands — and a private type would make a copy
 * unusable in exactly the case it is for: pasting into a text field in
 * another app.
 *
 * <h3>Read and write both</h3>
 * The write side is the obvious half. The read side matters just as much: it
 * is what lets a value copied in another app land in a cell, which is the
 * other direction of the same feature. A clipboard bridge that only writes is
 * a "send to other apps" button.
 */
public final class CorroClipboard {
    private CorroClipboard() {
    }

    /**
     * Put text on the system clipboard.
     *
     * <p>A {@link ClipData} with a single {@code text/plain} item. The
     * explicit MIME type is what makes the paste offer appear in a text
     * field; a clip with no type is shown by nothing.
     *
     * @return false if the clipboard is unavailable, so the caller can report
     *         the copy as in-app-only rather than pretending it went out
     */
    public static boolean setText(String text) {
        if (text == null) {
            return false;
        }
        try {
            android.content.ClipboardManager cm = clipboardOf();
            if (cm == null) {
                return false;
            }
            android.content.ClipData clip =
                    android.content.ClipData.newPlainText("corro cell", text);
            cm.setPrimaryClip(clip);
            return true;
        } catch (Throwable t) {
            android.util.Log.w("corro", "clipboard write failed: " + t);
            return false;
        }
    }

    /**
     * The clipboard's plain text, or null.
     *
     * <p>Returns null rather than an empty string for "nothing there", so a
     * caller can tell an empty clipboard from a cell holding "". Pasting
     * nothing is a no-op; pasting an empty string into a cell is a real edit.
     *
     * <p>Coerced from {@code coerceToText} rather than read as a string
     * directly: a clip from another app may be a URI, an intent or an HTML
     * span, and {@code coerceToText} is the platform's own answer to "what
     * text would a pasteable clip give", which is exactly the question.
     */
    public static String getText() {
        try {
            android.content.ClipboardManager cm = clipboardOf();
            if (cm == null) {
                return null;
            }
            android.content.ClipData clip = cm.getPrimaryClip();
            if (clip == null || clip.getItemCount() == 0) {
                return null;
            }
            CharSequence text = clip.getItemAt(0).coerceToText(contextOf());
            return text == null ? null : text.toString();
        } catch (Throwable t) {
            android.util.Log.w("corro", "clipboard read failed: " + t);
            return null;
        }
    }

    /** The system clipboard service, or null where there is none. */
    private static android.content.ClipboardManager clipboardOf() {
        android.content.Context ctx = contextOf();
        if (ctx == null) {
            return null;
        }
        return (android.content.ClipboardManager)
                ctx.getSystemService(android.content.Context.CLIPBOARD_SERVICE);
    }

    /**
     * A context to reach the service with.
     *
     * <p>The Activity, which is the app's only context and is always present
     * once the sheet is on screen. {@code getSystemService} is available from
     * any {@code Context}, and the Activity is the one corro has, so there is
     * no need for the (non-public) application-object lookup.
     */
    private static android.content.Context contextOf() {
        return CorroFile.activityContext();
    }
}
