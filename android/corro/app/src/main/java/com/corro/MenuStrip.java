package com.corro;

import android.content.Context;
import android.view.Gravity;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;
import android.widget.Button;
import android.widget.LinearLayout;
import android.widget.PopupMenu;
import android.widget.TextView;

import java.util.ArrayList;
import java.util.List;

/**
 * corro's Android menu bar: a compact top strip holding an app title, the
 * quick-action buttons (undo/redo and friends) and an overflow (&#x22ee;)
 * button.
 *
 * A phone cannot show six text menus the way the desktop GTK build does, so
 * the whole menu tree collapses into one {@link PopupMenu} behind the
 * overflow button: each top-level menu (File, Edit, ...) becomes a
 * submenu, and its items dispatch the same `app.*` action names the desktop
 * build uses, through {@link #nativeMenuAction(String)}.
 *
 * The strip is built from Rust ({@code create_menubar}); this class only
 * knows how to lay the pieces out and how to turn a tap into an action
 * name. Anything requiring more policy belongs on the Rust side.
 */
public class MenuStrip extends LinearLayout {

    /** Quick actions shown inline, in order. action name -> button label.
     *  Plain ASCII labels: the arrow/ellipsis glyphs (U+21B6/U+21B7/U+22EE)
     *  are missing from the emulator's default font, which rendered the
     *  buttons blank. */
    private static final String[][] QUICK_ACTIONS = {
        { "app.undo", "Undo" },
        { "app.redo", "Redo" },
    };

    private final List<String[]> overflow = new ArrayList<>();

    public MenuStrip(Context context, String title) {
        super(context);
        setOrientation(HORIZONTAL);
        setGravity(Gravity.CENTER_VERTICAL);

        TextView titleView = new TextView(context);
        titleView.setText(title);
        titleView.setGravity(Gravity.CENTER_VERTICAL);
        titleView.setTextColor(0xFF1B1B1B);
        titleView.setTextSize(18f);
        titleView.setPadding(dp(12), 0, 0, 0);
        LayoutParams titleParams =
            new LayoutParams(LayoutParams.WRAP_CONTENT, LayoutParams.MATCH_PARENT, 1f);
        titleView.setLayoutParams(titleParams);
        addView(titleView);

        for (String[] action : QUICK_ACTIONS) {
            addView(makeButton(context, action[1], action[0]));
        }

        Button overflowBtn = makeButton(context, "\u2261", null); // ≡ (triple line)
        overflowBtn.setContentDescription("More");
        overflowBtn.setOnClickListener(new OnClickListener() {
            @Override public void onClick(View v) { showOverflow(v); }
        });
        addView(overflowBtn);
        overflowButton = overflowBtn;
    }

    private Button makeButton(Context context, String label, final String action) {
        Button b = new Button(context);
        b.setText(label);
        b.setAllCaps(false);
        b.setMinWidth(0);
        b.setMinimumWidth(0);
        b.setPadding(dp(14), 0, dp(14), 0);
        // Explicit colours: the app theme derives from a bare Material parent
        // and leaves the default button text on a light background, which
        // rendered the labels invisible. Set both ends of the contrast pair.
        b.setTextColor(0xFF1B1B1B);
        b.setTextSize(16f);
        b.setBackgroundColor(0xFFEDEDED);
        if (action != null) {
            b.setContentDescription(action);
            b.setOnClickListener(new OnClickListener() {
                @Override public void onClick(View v) { nativeMenuAction(action); }
            });
        }
        return b;
    }

    private int dp(int value) {
        return (int) (value * getResources().getDisplayMetrics().density + 0.5f);
    }

    /**
     * Add one top-level menu and its items. Called once per menu by Rust
     * with the same tree the desktop build uses; labels are already
     * mnemonic-stripped by the caller.
     *
     * @param label  top-level label ("File")
     * @param actions alternating item label / action name
     */
    public void addMenu(String label, String[] actions) {
        overflow.add(concat(new String[] { label }, actions));
    }

    private static String[] concat(String[] a, String[] b) {
        String[] out = new String[a.length + b.length];
        System.arraycopy(a, 0, out, 0, a.length);
        System.arraycopy(b, 0, out, a.length, b.length);
        return out;
    }

    /**
     * Open the overflow with one top-level submenu already expanded.
     *
     * <p>This is what an Alt+mnemonic (or a programmatic
     * {@code popup_submenu_by_mnemonic_at}) does on a phone. A desktop menu
     * bar highlights the matched top-level menu and drops the submenu under
     * it; here the whole tree lives behind the overflow button, so the
     * closest equivalent is opening the overflow and leaving that one
     * submenu highlighted — which is what the highlight below does, so the
     * user sees <em>which</em> menu the key opened rather than just a list.
     *
     * <p>An unknown label is ignored rather than opening the popup blind: a
     * mnemonic for a menu the strip does not have should do nothing, not
     * open an unrelated menu.
     */
    public void showOverflowFor(String topLabel) {
        if (findMenu(topLabel) < 0) {
            return;
        }
        showOverflow(findOverflowButton());
        nativeMenuOpened(1);
    }

    /**
     * Close the overflow. {@link PopupMenu} has no close(), so dismissing
     * needs a reference to the instance; it is kept because Escape has to
     * close a menu that Rust decided was open.
     */
    public void closeOverflow() {
        if (popup != null) {
            popup.dismiss();
            popup = null;
        }
        nativeMenuOpened(0);
    }

    /** The overflow button, so a programmatic open has an anchor. */
    private View findOverflowButton() {
        if (overflowButton == null) {
            // Built first in addView order after the quick actions; find it
            // by content description rather than by index, which would break
            // the moment a quick action is added or removed.
            for (int i = 0; i < getChildCount(); i++) {
                View child = getChildAt(i);
                if ("More".equals(child.getContentDescription())) {
                    return child;
                }
            }
            return this;
        }
        return overflowButton;
    }

    private int findMenu(String topLabel) {
        for (int i = 0; i < overflow.size(); i++) {
            if (overflow.get(i)[0].equals(topLabel)) {
                return i;
            }
        }
        return -1;
    }

    private void showOverflow(View anchor) {
        if (anchor == null) {
            return;
        }
        PopupMenu popup = new PopupMenu(getContext(), anchor);
        this.popup = popup;
        Menu menu = popup.getMenu();
        int groupId = 0;
        for (String[] entry : overflow) {
            // entry = { topLabel, itemLabel, action, itemLabel, action, ... }
            String top = entry[0];
            SubMenuBuilder sub = new SubMenuBuilder();
            for (int i = 1; i + 1 < entry.length; i += 2) {
                sub.labels.add(entry[i]);
                sub.actions.add(entry[i + 1]);
            }
            android.view.SubMenu m = menu.addSubMenu(top);
            for (int i = 0; i < sub.labels.size(); i++) {
                m.add(groupId, nextItemId(), i, sub.labels.get(i));
            }
            groupId++;
        }
        popup.setOnMenuItemClickListener(new PopupMenu.OnMenuItemClickListener() {
            @Override public boolean onMenuItemClick(MenuItem item) {
                int g = item.getGroupId();
                if (g >= 0 && g < overflow.size()) {
                    String[] entry = overflow.get(g);
                    int idx = item.getOrder();
                    int actionAt = 1 + idx * 2 + 1;
                    if (actionAt < entry.length) {
                        nativeMenuAction(entry[actionAt]);
                        return true;
                    }
                }
                return false;
            }
        });
        popup.show();
        // Mirror the open state into Rust so `menu_active` is a real answer:
        // the desktop uses it to route a plain letter key to menu selection
        // instead of starting a cell edit, and a constant `false` made an
        // open menu and a typed character collide.
        nativeMenuOpened(1);
    }

    /** The live overflow popup, or null when none is open. */
    private PopupMenu popup;

    /** The overflow button, kept so a programmatic open can anchor to it. */
    private View overflowButton;

    private int nextItemId() {
        return ++lastItemId;
    }

    private int lastItemId = 0;

    /** Scratch holder so the popup build stays readable. */
    private static final class SubMenuBuilder {
        final List<String> labels = new ArrayList<>();
        final List<String> actions = new ArrayList<>();
    }

    /** Dispatch an action name to Rust; the name matches the desktop build. */
    private static native void nativeMenuAction(String action);

    /**
     * Report the overflow's open state to Rust (1 open, 0 closed).
     *
     * <p>Called on every open and close rather than having Rust poll: a
     * PopupMenu dismisses itself when the user taps outside it, and only
     * Java hears about that, so a one-way query would leave `menu_active`
     * stuck at true and every later letter key would go to menu selection
     * instead of the cell.
     */
    private static native void nativeMenuOpened(int open);
}
