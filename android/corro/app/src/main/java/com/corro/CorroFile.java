package com.corro;

import android.app.Activity;
import android.content.ContentResolver;
import android.content.Intent;
import android.net.Uri;
import android.util.Log;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * The Storage Access Framework, bridged to corro's synchronous file-dialog
 * API.
 *
 * <h3>Why this class exists</h3>
 * Android has no filesystem a normal app may read or write, and no
 * {@code FILE_CHOOSER}. The supported answer is the Storage Access Framework:
 * {@code ACTION_OPEN_DOCUMENT} and {@code ACTION_CREATE_DOCUMENT} hand the
 * user the platform's document picker and return a {@code content://} URI.
 *
 * <p>Before this, corro's File &rarr; Open silently opened nothing and
 * File &rarr; Save As silently discarded the save, because
 * {@code open_file_filtered} returned {@code Err} on Android. The whole
 * File and Export menus were dead.
 *
 * <h3>The asynchrony problem</h3>
 * SAF is inherently asynchronous: {@code startActivityForResult} returns
 * immediately and the URI arrives later in {@code onActivityResult}, on the
 * Activity's looper. corro's file dialogs are synchronous — {@code
 * open_file_filtered} returns {@code Option<PathBuf>} — and their callers run
 * <em>on</em> the UI thread, inside a native call. So the answer has to be
 * pumped out by hand.
 *
 * <p>{@link #awaitResult} does that with {@link MessageQueue#next()} in a
 * loop. That call <em>blocks</em> until a message is due rather than spinning,
 * so the wait costs nothing while the user is in the picker — which is the
 * alternative to this approach: a busy-wait would burn battery for the
 * seconds or minutes a document picker is open.
 *
 * <h3>URIs, not paths</h3>
 * A returned {@code content://} URI is not a filesystem path and cannot be
 * opened with {@code std::fs}. Rust therefore keeps the URI string wherever a
 * path would go and reads/writes it through
 * {@link #readBytes(String)} / {@link #writeBytes(String, byte[])}, which call
 * {@code ContentResolver.openInputStream} / {@code openOutputStream}. The
 * permission granted by the picker is <em>persistable</em> (requested below),
 * so a saved document can be rewritten on a later launch without asking again.
 */
public final class CorroFile {
    private static final String TAG = "corro";

    /** In-flight request; the Activity's result lands in {@link #result}. */
    private static Activity activity;
    private static String result;
    private static int requestCode = 0x0c07;
    private static boolean answered;
    private static final List<String> pendingMimes = new ArrayList<>();
    private static String pendingName;
    private static boolean pendingSave;

    private CorroFile() {
    }

    /**
     * Bind to the Activity. The picker can only be started from one, and
     * {@code onActivityResult} must be forwarded here by the host.
     */
    public static void attach(Activity act) {
        activity = act;
    }

    /**
     * The bound Activity, or null before {@link #attach}.
     *
     * <p>Also the app's <em>only</em> {@code Context}, which is why the
     * clipboard bridge asks here rather than carrying its own handle: there
     * is one Activity, it is registered once, and a second path to it would
     * be a second thing that can be null.
     */
    public static android.content.Context activityContext() {
        return activity;
    }

    /**
     * Show the platform's "open document" picker.
     *
     * @param mimes MIME types to accept; {@code *&#47;*} is the any-file case
     * @return the chosen {@code content://} URI, or null if cancelled or
     *         unavailable
     */
    public static String open(String[] mimes) {
        return run(false, mimes, null);
    }

    /**
     * Show the platform's "create document" picker, seeded with a filename.
     *
     * @return the chosen {@code content://} URI, or null if cancelled
     */
    public static String create(String mime, String defaultName) {
        return run(true, new String[]{mime}, defaultName);
    }

    private static String run(boolean save, String[] mimes, String defaultName) {
        if (activity == null) {
            Log.w(TAG, "CorroFile: no activity attached");
            return null;
        }
        // A document handed to the app from outside wins over showing a
        // picker: the user already chose it, and asking again would make
        // "Open with corro" open a *file chooser*.
        if (!save) {
            String pending = takeExternal();
            if (pending != null) {
                return pending;
            }
        }
        Activity act = activity;
        Intent intent;
        if (save) {
            // CREATE, not OPEN: the user picks where to write. Without
            // CREATE the platform will not let them name a new file, and
            // Save As would then silently overwrite an existing one.
            intent = new Intent(Intent.ACTION_CREATE_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
            intent.setType(mimes != null && mimes.length > 0 ? mimes[0] : "*/*");
            if (defaultName != null) {
                intent.putExtra(Intent.EXTRA_TITLE, defaultName);
            }
        } else {
            intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
            intent.addCategory(Intent.CATEGORY_OPENABLE);
            // Ask for *both* a concrete type and the wildcard list: a
            // provider that honours neither returns nothing, and an empty
            // filter shows the user a picker with nothing in it.
            if (mimes != null && mimes.length == 1) {
                intent.setType(mimes[0]);
            } else {
                intent.setType("*/*");
                if (mimes != null && mimes.length > 0) {
                    intent.putExtra(Intent.EXTRA_MIME_TYPES, mimes);
                }
            }
        }
        // Persistable, so a document opened now can still be written after a
        // restart. Without this flag the grant lasts only for the life of the
        // Activity and the next save would fail with a SecurityException that
        // is very hard to attribute later.
        intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION
                | Intent.FLAG_GRANT_WRITE_URI_PERMISSION
                | Intent.FLAG_GRANT_PERSISTABLE_URI_PERMISSION);

        answered = false;
        result = null;
        pendingSave = save;
        pendingMimes.clear();
        if (mimes != null) {
            for (String m : mimes) {
                if (m != null && !m.isEmpty() && !pendingMimes.contains(m)) {
                    pendingMimes.add(m);
                }
            }
        }
        pendingName = defaultName;
        try {
            act.startActivityForResult(intent, requestCode);
        } catch (Exception e) {
            Log.w(TAG, "CorroFile: no document picker: " + e);
            return null;
        }
        return awaitResult();
    }

    /**
     * Block until the picker answers.
     *
     * <p>SAF is asynchronous but corro's file dialogs are synchronous: the
     * Rust caller is inside a native call, on the UI thread, holding no
     * opportunity to return and be called back. So the UI thread has to wait
     * here, and the wait must not be a spin — the user may be in a document
     * picker for minutes, and a busy-wait would burn a core for every one of
     * them.
     *
     * <p>{@link #awaited} is signalled by {@link #onActivityResult} from the
     * platform's own delivery of {@code startActivityForResult}, which the
     * Activity posts to the main looper while this thread is parked. That is
     * the only reason this works: parking the main looper's thread does
     * <em>not</em> stop the looper from dispatching a message that is already
     * queued, which is exactly what {@code onActivityResult} is.
     *
     * <p>The deadline is generous but bounded: the user is in another app for
     * as long as they want, yet a picker killed by the system mid-selection
     * would otherwise never answer and the process would hang forever.
     */
    private static String awaitResult() {
        synchronized (awaited) {
            long deadline = android.os.SystemClock.uptimeMillis() + 10 * 60 * 1000L;
            while (!answered) {
                long remaining = deadline - android.os.SystemClock.uptimeMillis();
                if (remaining <= 0) {
                    Log.w(TAG, "CorroFile: picker timed out");
                    break;
                }
                try {
                    // wait, not waitUntil: `onActivityResult` is already
                    // queued, so it cannot arrive early enough to be missed
                    // by a bounded wait, and waking early would just re-park.
                    awaited.wait(remaining);
                } catch (InterruptedException e) {
                    Thread.currentThread().interrupt();
                    Log.w(TAG, "CorroFile: interrupted while waiting");
                    break;
                }
            }
            // Re-arm for the next pick. Done here rather than in `run` so a
            // timed-out pick cannot leave `answered` set and make the *next*
            // dialog return instantly with the previous answer.
            answered = false;
        }
        return result;
    }

    /** The waiter {@link #awaitResult} parks on. */
    private static final Object awaited = new Object();

    /**
     * The host Activity's {@code onActivityResult}, forwarded here.
     *
     * <p>Forwarded rather than registered because the platform delivers the
     * result to the Activity, and the Activity is the host's — rswidgets
     * cannot subclass it. The host keeps its own result plumbing and calls
     * this; see MainActivity.
     */
    public static void onActivityResult(int request, int resultCode, Intent data) {
        if (request != requestCode) {
            return;
        }
        if (resultCode == Activity.RESULT_OK && data != null && data.getData() != null) {
            result = data.getData().toString();
            if (activity != null) {
                try {
                    activity.getContentResolver().takePersistableUriPermission(
                            data.getData(),
                            Intent.FLAG_GRANT_READ_URI_PERMISSION
                                    | Intent.FLAG_GRANT_WRITE_URI_PERMISSION);
                } catch (SecurityException e) {
                    // Not every provider offers a persistable grant. The URI
                    // still works for this session, so this is a warning, not
                    // a failure: failing here would turn a "cannot re-save
                    // after a restart" into a "cannot open at all".
                    Log.w(TAG, "CorroFile: persistable grant refused: " + e);
                }
            }
        }
        synchronized (awaited) {
            answered = true;
            // Wake the parked UI thread. This is posted to the main looper
            // while it is parked in awaitResult, which is why the wait works
            // at all: parking the thread does not stop the looper dispatching
            // an already-queued message.
            awaited.notifyAll();
        }
    }

    /**
     * Read a document's whole contents.
     *
     * <p>Streams through a {@link ByteArrayOutputStream} rather than trusting
     * {@code ContentResolver.openFileDescriptor} with a length: a document
     * provider may report a size it does not honour (a compressed or
     * generated document), and a short read would silently truncate a
     * workbook.
     */
    public static byte[] readBytes(String uriString) {
        if (activity == null || uriString == null) {
            return new byte[0];
        }
        ContentResolver cr = activity.getContentResolver();
        InputStream in = null;
        try {
            Uri uri = Uri.parse(uriString);
            in = cr.openInputStream(uri);
            if (in == null) {
                return new byte[0];
            }
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            byte[] buf = new byte[16 * 1024];
            int n;
            while ((n = in.read(buf)) > 0) {
                out.write(buf, 0, n);
            }
            return out.toByteArray();
        } catch (Exception e) {
            Log.w(TAG, "CorroFile: read failed: " + e);
            return new byte[0];
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (Exception ignored) {
                    // Closing a stream that already failed is not itself an
                    // error worth reporting.
                }
            }
        }
    }

    /**
     * Overwrite a document's contents.
     *
     * <p>Truncating matters: a document that already exists and is being
     * saved over with a <em>shorter</em> workbook would otherwise keep its
     * tail, producing a file that parses as corrupt rather than as the older,
     * longer revision. {@code "wt"} is truncate-and-write;
     * {@code "wa"} (append) would be exactly the wrong mode here.
     */
    public static boolean writeBytes(String uriString, byte[] data) {
        if (activity == null || uriString == null) {
            return false;
        }
        ContentResolver cr = activity.getContentResolver();
        OutputStream out = null;
        try {
            Uri uri = Uri.parse(uriString);
            out = cr.openOutputStream(uri, "wt");
            if (out == null) {
                return false;
            }
            out.write(data);
            out.flush();
            return true;
        } catch (Exception e) {
            Log.w(TAG, "CorroFile: write failed: " + e);
            return false;
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (Exception ignored) {
                    // The bytes are already written; a close failure is not
                    // worth turning a successful save into a failed one.
                }
            }
        }
    }

    /**
     * A document handed to the app from outside — a file manager's "Open with
     * corro" produces a {@code VIEW} or {@code SEND} intent.
     *
     * <p>Delivered through the same {@code answer} path an in-app pick uses,
     * so the Rust side has one place to read a picked document and no need to
     * know whether the user tapped a dialog or tapped the app in a chooser.
     * The intent arrives while the sheet is already running, so the answer is
     * queued and picked up by whichever file dialog is next — there is no
     * dialog to block, and nothing should be loaded twice.
     */
    public static void deliverExternalIntent(Intent intent) {
        if (intent == null) {
            return;
        }
        // A VIEW intent carries the document in getData(); a SEND ("share to")
        // carries it in EXTRA_STREAM, because a share carries a payload rather
        // than a target. Both are "a document the user handed us", and both
        // must land in the same place or one of them silently does nothing.
        Uri uri = intent.getData();
        if (uri == null && Intent.ACTION_SEND.equals(intent.getAction())) {
            uri = sendStream(intent);
        }
        if (uri == null || !ContentResolver.SCHEME_CONTENT.equals(uri.getScheme())) {
            return;
        }
        if (activity != null) {
            try {
                activity.getContentResolver().takePersistableUriPermission(uri,
                        Intent.FLAG_GRANT_READ_URI_PERMISSION);
            } catch (SecurityException ignored) {
                // A VIEW intent's grant is not persistable on every provider.
                // The document is still readable now, which is all an
                // externally-delivered one needs to be.
            }
        }
        external = uri.toString();
        Log.i(TAG, "CorroFile: external document " + external);
    }

    /**
     * Take a document delivered by {@link #deliverExternalIntent}, if one is
     * waiting. Consumed on read, so a second Open does not reopen the same
     * document by accident.
     */
    public static String takeExternal() {
        String uri = external;
        external = null;
        return uri;
    }

    /** A document handed to the app from outside, awaiting an in-app Open. */
    private static String external;

    /**
     * The {@code EXTRA_STREAM} of a SEND intent.
     *
     * <p>The untyped {@code getParcelableExtra(String)} is deprecated from
     * API 33, but the typed {@code getParcelableExtra(String, Class)} that
     * replaced it does not exist before it — and this app's {@code minSdk} is
     * 24. So the version is checked rather than either call being assumed.
     * The stream is a {@link Uri} for every document-sending app in practice;
     * a non-URI payload is ignored rather than crashing, because a share
     * intent is untrusted input from another app.
     */
    private static Uri sendStream(Intent intent) {
        try {
            Object stream;
            if (android.os.Build.VERSION.SDK_INT >= android.os.Build.VERSION_CODES.TIRAMISU) {
                stream = intent.getParcelableExtra(Intent.EXTRA_STREAM, Uri.class);
            } else {
                // Deprecated from API 33, but the replacement does not exist
                // before it and this app supports API 24. The @Suppress is
                // scoped to this statement because the deprecation is
                // unavoidable, not overlooked.
                @SuppressWarnings("deprecation")
                Object legacy = intent.getParcelableExtra(Intent.EXTRA_STREAM);
                stream = legacy;
            }
            return stream instanceof Uri ? (Uri) stream : null;
        } catch (Exception e) {
            Log.w(TAG, "CorroFile: bad SEND payload: " + e);
            return null;
        }
    }

    /**
     * Copy a document into the app's private storage and return the path.
     *
     * <p>For the *import* direction this sidesteps the whole URI problem: an
     * opened document is read once and copied to a real file, so every
     * existing loader (which all take a {@code Path}) works unchanged, and the
     * permission can lapse afterwards.
     */
    public static String materialize(String uriString, String fileName) {
        byte[] data = readBytes(uriString);
        if (data.length == 0) {
            return null;
        }
        try {
            java.io.File dir = new java.io.File(
                    activity.getFilesDir(), "corro");
            if (!dir.exists() && !dir.mkdirs()) {
                return null;
            }
            java.io.File dest = new java.io.File(dir, fileName);
            java.io.FileOutputStream out = new java.io.FileOutputStream(dest);
            try {
                out.write(data);
                out.flush();
            } finally {
                out.close();
            }
            return dest.getAbsolutePath();
        } catch (Exception e) {
            Log.w(TAG, "CorroFile: materialize failed: " + e);
            return null;
        }
    }
}
