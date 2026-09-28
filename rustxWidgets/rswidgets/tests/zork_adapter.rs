//! Adapter-level tests for the zork backend.
//!
//! These drive the *public adapter surface* (`create_*` + widget methods) over
//! the thread-local model, which is what application code actually calls. The
//! model unit tests in `backends::zork::model` cover the operations; these
//! cover the wiring, so a method that forgets to pass its arguments through
//! fails here.
//!
//! The thread-local singleton is shared within a test binary, so every test
//! starts from a fresh model via the `serial` helper rather than assuming an
//! empty one.

#![cfg(feature = "zork")]

use std::cell::RefCell;
use std::rc::Rc;

use rswidgets::backends_zork_adapter as za;

/// Run `f` against a fresh thread-local model.
///
/// The adapter's free functions operate on a thread-local `ZorkState`, so tests
/// that assert absolute node ids would otherwise interfere with each other.
/// These tests only assert *relative* facts, but starting from a clean model
/// keeps them independent regardless.
fn serial<R>(f: impl FnOnce() -> R) -> R {
    rswidgets::backends::zork::model::reset_for_test();
    f()
}

mod geometry {
    use super::*;

    #[test]
    fn window_default_size_is_recorded() {
        serial(|| {
            let w = za::create_window().unwrap();
            w.set_default_size(640, 480);
            assert_eq!(rswidgets::backends::zork::get_size_request(w.id()).map(|(a, b)| (a, b)), Some((Some(640), Some(480))));
        });
    }

    #[test]
    fn window_resize_overrides_default_size() {
        serial(|| {
            let w = za::create_window().unwrap();
            w.set_default_size(640, 480);
            w.resize(800, 600);
            assert_eq!(
                rswidgets::backends::zork::get_size_request(w.id()),
                Some((Some(800), Some(600)))
            );
        });
    }

    #[test]
    fn box_layout_positions_children() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let bx = za::create_box(za::Orientation::Horizontal, 4).unwrap();
            let a = za::create_button("a").unwrap();
            let b = za::create_button("b").unwrap();
            bx.append(&a);
            bx.append(&b);
            a.set_size_request(10, 6);
            b.set_size_request(20, 6);
            bx.layout(0, 0, 200, 100);
            assert_eq!(rswidgets::backends::zork::get_offset(a.id()), Some((Some(0), Some(0))));
            assert_eq!(rswidgets::backends::zork::get_offset(b.id()), Some((Some(14), Some(0))));
            assert_eq!(bx.measure(), Some((34, 6)));
        });
    }

    #[test]
    fn grid_attach_records_cell_and_grows_grid() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let g = za::create_grid().unwrap();
            let b = za::create_button("b").unwrap();
            g.attach(&b, 2, 3, 1, 1);
            assert_eq!(g.dimensions(), (3, 4));
            assert_eq!(rswidgets::backends::zork::get_offset(b.id()), Some((Some(2), Some(3))));
        });
    }

    #[test]
    fn append_keeps_both_children() {
        // Regression: `append_child` used to delegate to `set_child`, which
        // replaced the child list, so the second append silently dropped the
        // first child.
        serial(|| {
            let _w = za::create_window().unwrap();
            let bx = za::create_box(za::Orientation::Vertical, 0).unwrap();
            let a = za::create_button("a").unwrap();
            let b = za::create_button("b").unwrap();
            bx.append(&a);
            bx.append(&b);
            assert_eq!(rswidgets::backends::zork::first_child(bx.id()), Some(a.id()));
            let layers = rswidgets::backends::zork::measure_box(bx.id());
            assert!(layers.is_some());
            // Both children are present.
            let snap = rswidgets::backends::zork::model::with_state_for_test(|s| s.snapshot());
            let node = snap.nodes.iter().find(|n| n.id == bx.id()).unwrap();
            assert_eq!(node.children.len(), 2, "both children must be attached");
        });
    }
}

mod visibility_and_classes {
    use super::*;

    #[test]
    fn set_visible_and_classes_round_trip() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let l = za::create_label("x").unwrap();
            assert!(rswidgets::backends::zork::get_visible(l.id()));
            l.set_visible(false);
            assert!(!rswidgets::backends::zork::get_visible(l.id()));
            l.add_class("primary");
            l.add_class("primary");
            assert_eq!(rswidgets::backends::zork::classes(l.id()), vec!["primary".to_string()]);
            l.remove_class("primary");
            assert!(!rswidgets::backends::zork::has_class(l.id(), "primary"));
        });
    }

    #[test]
    fn label_fixed_width_pins_size_request() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let l = za::create_label("x").unwrap();
            l.set_fixed_width(Some(120));
            assert_eq!(rswidgets::backends::zork::get_size_request(l.id()), Some((Some(120), None)));
            l.set_fixed_width(None);
            assert_eq!(rswidgets::backends::zork::get_size_request(l.id()), Some((None, None)));
        });
    }

    #[test]
    fn expansion_and_margins_are_recorded() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let bx = za::create_box(za::Orientation::Vertical, 0).unwrap();
            let c = za::create_label("c").unwrap();
            bx.append(&c);
            bx.set_hexpand(true);
            bx.set_child_hexpand(&c, true);
            bx.set_child_vexpand(&c, true);
            c.set_margin_start(4);
            c.set_margin_top(8);
            let props = rswidgets::backends::zork::model::with_state_for_test(|s| s.node(c.id()).unwrap().props.clone());
            assert!(props.hexpand);
            assert!(props.vexpand);
            assert!(props.hexpand, "child expansion is recorded on the child");
            assert_eq!(props.margin_start, 4);
            assert_eq!(props.margin_top, 8);
        });
    }
}

mod focus_and_caret {
    use super::*;

    #[test]
    fn entry_reports_a_real_caret() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let e = za::create_entry().unwrap();
            e.set_text("hello");
            assert_eq!(e.get_position(), Some(5));
            e.set_position(2);
            assert_eq!(e.get_position(), Some(2));
            e.set_position(99);
            assert_eq!(e.get_position(), Some(5), "clamped to the buffer length");
        });
    }

    #[test]
    fn grab_focus_and_has_focus_agree() {
        // The old stub hard-coded `has_focus() == false`, so a caller could
        // never tell whether focus moved.
        serial(|| {
            let _w = za::create_window().unwrap();
            let a = za::create_entry().unwrap();
            let b = za::create_entry().unwrap();
            assert!(!a.has_focus());
            a.grab_focus();
            assert!(a.has_focus());
            b.grab_focus();
            assert!(b.has_focus());
            assert!(!a.has_focus(), "focus moves, it does not accumulate");
        });
    }
}

mod radio_and_check {
    use super::*;

    #[test]
    fn radio_group_is_exclusive_through_the_adapter() {
        // Passing a seed joins its group. An ungrouped seed is promoted to be
        // its own group, so the seed and every member stay exclusive.
        serial(|| {
            let _w = za::create_window().unwrap();
            let seed = za::create_radiobutton(None, "seed").unwrap();
            let a = za::create_radiobutton(Some(&seed), "a").unwrap();
            let b = za::create_radiobutton(Some(&seed), "b").unwrap();
            assert_eq!(seed.group(), Some(seed.id()));
            assert_eq!(a.group(), seed.group());
            assert_eq!(b.group(), seed.group());

            a.set_active(true);
            assert!(a.is_active());
            b.set_active(true);
            assert!(b.is_active());
            assert!(!a.is_active(), "setting one radio clears the group");
        });
    }

    #[test]
    fn ungrouped_radios_are_independent() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let a = za::create_radiobutton(None, "a").unwrap();
            let b = za::create_radiobutton(None, "b").unwrap();
            a.set_active(true);
            b.set_active(true);
            assert!(a.is_active(), "no seed means no group, so no clearing");
            assert!(b.is_active());
        });
    }

    #[test]
    fn radio_can_be_grouped_after_creation() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let a = za::create_radiobutton(None, "a").unwrap();
            let b = za::create_radiobutton(None, "b").unwrap();
            let group = 42;
            a.set_group(group);
            b.set_group(group);
            a.set_active(true);
            b.set_active(true);
            assert!(b.is_active());
            assert!(!a.is_active());
        });
    }

    #[test]
    fn check_toggle_fires_callbacks() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let cb = za::create_checkbutton("enable").unwrap();
            let hits = Rc::new(RefCell::new(0));
            {
                let h = hits.clone();
                cb.on_toggle(move || *h.borrow_mut() += 1).unwrap();
            }
            assert!(!cb.is_active());
            cb.toggle();
            assert!(cb.is_active());
            assert_eq!(*hits.borrow(), 1);
            cb.toggle();
            assert!(!cb.is_active());
            assert_eq!(*hits.borrow(), 2);
        });
    }

    #[test]
    fn labels_are_gettable_and_settable() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let cb = za::create_checkbutton("before").unwrap();
            assert_eq!(cb.get_label(), Some("before".to_string()));
            cb.set_label("after");
            assert_eq!(cb.get_label(), Some("after".to_string()));
        });
    }
}

mod buttons_and_canvas {
    use super::*;

    #[test]
    fn emit_clicked_fires_the_handler() {
        // The old `emit_clicked` returned `Ok(0)` without firing anything.
        serial(|| {
            let _w = za::create_window().unwrap();
            let b = za::create_button("go").unwrap();
            let hits = Rc::new(RefCell::new(0));
            {
                let h = hits.clone();
                b.on_click(move || *h.borrow_mut() += 1).unwrap();
            }
            b.emit_clicked().unwrap();
            assert_eq!(*hits.borrow(), 1);
        });
    }

    #[test]
    fn canvas_keeps_and_runs_its_draw_callback() {
        use rswidgets::backends::headless::DrawOp;
        serial(|| {
            let _w = za::create_window().unwrap();
            let c = za::create_canvas().unwrap();
            c.set_content_size(100, 50);
            c.set_draw_callback(Box::new(|ctx, w, h| {
                assert_eq!((w, h), (100, 50));
                ctx.fill_rect(0.0, 0.0, 10.0, 10.0, 1.0, 0.0, 0.0, 1.0);
                ctx.draw_text(0.0, 0.0, "hello", "sans", 12.0, 0.0, 0.0, 0.0, 1.0);
            }));
            let ops = c.draw();
            assert!(matches!(ops[0], DrawOp::FillRect { .. }));
            assert!(ops.iter().any(|o| matches!(o, DrawOp::Text { text, .. } if text == "hello")));
            // Still drawable afterwards.
            assert_eq!(c.draw().len(), ops.len());
            c.clear_draw_callback();
            assert!(c.draw().is_empty());
        });
    }

    #[test]
    fn canvas_pointer_hooks_fire_in_order() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let c = za::create_canvas().unwrap();
            let order = Rc::new(RefCell::new(Vec::new()));
            {
                let o = order.clone();
                c.on_click_button(Box::new(move |_, _, b, _| o.borrow_mut().push(format!("btn{}", b))));
            }
            {
                let o = order.clone();
                c.on_click(Box::new(move |x, y| o.borrow_mut().push(format!("click{},{}", x as i32, y as i32))));
            }
            rswidgets::backends::zork::pointer_click_button(c.id(), 3.0, 4.0, 1, 0);
            assert_eq!(*order.borrow(), vec!["btn1".to_string(), "click3,4".to_string()]);
        });
    }

    #[test]
    fn queue_redraw_is_counted() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let c = za::create_canvas().unwrap();
            assert_eq!(c.redraw_count(), 0);
            c.queue_redraw();
            c.queue_redraw();
            assert_eq!(c.redraw_count(), 2);
        });
    }

    #[test]
    fn canvas_force_draw_uses_the_fallback_size() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let c = za::create_canvas().unwrap();
            let drew = Rc::new(RefCell::new((0, 0)));
            {
                let d = drew.clone();
                c.set_draw_callback(Box::new(move |_, w, h| *d.borrow_mut() = (w, h)));
            }
            c.force_draw(std::ptr::null_mut(), 320, 240);
            assert_eq!(*drew.borrow(), (320, 240));
            // An explicit size wins over the fallback.
            c.set_content_size(10, 20);
            c.force_draw(std::ptr::null_mut(), 320, 240);
            assert_eq!(*drew.borrow(), (10, 20));
        });
    }

    #[test]
    fn canvas_key_hook_can_consume() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let c = za::create_canvas().unwrap();
            let raw = Rc::new(RefCell::new(0));
            c.on_key(Box::new(|_| true));
            {
                let r = raw.clone();
                c.on_key_raw(Box::new(move |_, _| { *r.borrow_mut() += 1; true }));
            }
            assert!(rswidgets::backends::zork::key(c.id(), 42));
            assert_eq!(*raw.borrow(), 0, "a consumed key skips key_raw");
        });
    }
}

mod overlay {
    use super::*;

    #[test]
    fn base_child_and_layers_are_distinct() {
        // The old body routed `add_overlay` and `set_child` through the same
        // call, so a layer was indistinguishable from the base.
        serial(|| {
            let _w = za::create_window().unwrap();
            let ov = za::create_overlay().unwrap();
            let base = za::create_label("base").unwrap();
            let layer = za::create_button("menu").unwrap();
            ov.set_child(&base);
            ov.add_overlay(&layer);
            assert_eq!(rswidgets::backends::zork::first_child(ov.id()), Some(base.id()));
            assert_eq!(ov.layers(), vec![layer.id()]);
        });
    }

    #[test]
    fn pass_through_and_show_all_are_recorded() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let ov = za::create_overlay().unwrap();
            let base = za::create_label("base").unwrap();
            let layer = za::create_button("menu").unwrap();
            ov.set_child(&base);
            ov.add_overlay(&layer);
            ov.set_overlay_pass_through(&layer, true);
            let snap = rswidgets::backends::zork::model::with_state_for_test(|s| s.snapshot());
            assert_eq!(snap.overlay_pass_through.get(&layer.id()), Some(&true));

            base.set_visible(false);
            ov.show_all();
            assert!(rswidgets::backends::zork::get_visible(base.id()));
        });
    }

    #[test]
    fn remove_detaches_the_layer() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let ov = za::create_overlay().unwrap();
            let base = za::create_label("base").unwrap();
            let layer = za::create_button("menu").unwrap();
            ov.set_child(&base);
            ov.add_overlay(&layer);
            ov.remove(&layer);
            assert!(ov.layers().is_empty());
        });
    }
}

mod scrolled_window {
    use super::*;

    #[test]
    fn scroll_is_clamped_and_reported() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let sw = za::create_scrolled_window().unwrap();
            sw.scroll_to(0.0, 100.0, 20.0, 999.0, 100.0, 20.0);
            assert_eq!(sw.get_scroll(), (0.0, 80.0));
        });
    }

    #[test]
    fn on_scroll_fires_with_axis_and_position() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let sw = za::create_scrolled_window().unwrap();
            let seen = Rc::new(RefCell::new(Vec::new()));
            {
                let s = seen.clone();
                sw.on_scroll(Box::new(move |v, p| s.borrow_mut().push((v, p))));
            }
            rswidgets::backends::zork::scroll(sw.id(), true, 12.0);
            assert_eq!(*seen.borrow(), vec![(true, 12.0)]);
            assert_eq!(sw.get_scroll().1, 12.0);
        });
    }
}

mod menus {
    use super::*;

    #[test]
    fn select_dispatches_the_named_action() {
        // The old `Harness::select` fired the menu node's own callbacks, which
        // made the action name on each item decorative.
        serial(|| {
            let _w = za::create_window().unwrap();
            let action = za::create_simple_action("app.open").unwrap();
            let menu = za::create_menu().unwrap();
            menu.append("Open", "app.open");
            let hits = Rc::new(RefCell::new(0));
            {
                let h = hits.clone();
                action.on_activate(move || *h.borrow_mut() += 1).unwrap();
            }
            assert_eq!(menu.select(0), Some("Open".to_string()));
            assert_eq!(*hits.borrow(), 1);
        });
    }

    #[test]
    fn append_item_binds_to_the_action_name() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let action = za::create_simple_action("app.save").unwrap();
            let menu = za::create_menu().unwrap();
            menu.append_item("Save", &action);
            let hits = Rc::new(RefCell::new(0));
            {
                let h = hits.clone();
                action.on_activate(move || *h.borrow_mut() += 1).unwrap();
            }
            assert_eq!(menu.select(0), Some("Save".to_string()));
            assert_eq!(*hits.borrow(), 1);
        });
    }

    #[test]
    fn separator_is_not_selectable() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let menu = za::create_menu().unwrap();
            menu.append_separator("");
            menu.append("Open", "app.open");
            assert_eq!(menu.select(0), None, "a separator must not select");
            assert_eq!(menu.select(1), Some("Open".to_string()));
            assert_eq!(menu.select(9), None, "out of range");
        });
    }

    #[test]
    fn check_item_toggles_when_selected() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let menu = za::create_menu().unwrap();
            let idx = menu.append_check("Wrap", "app.wrap", false);
            assert!(!menu.items()[idx].checked);
            menu.select(idx);
            assert!(menu.items()[idx].checked);
            menu.select(idx);
            assert!(!menu.items()[idx].checked);
        });
    }

    #[test]
    fn radio_items_stay_exclusive() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let menu = za::create_menu().unwrap();
            let a = menu.append_radio("a", "", 1, false);
            let b = menu.append_radio("b", "", 1, false);
            menu.set_item_checked(b, true);
            assert!(menu.items()[b].checked);
            assert!(!menu.items()[a].checked);
        });
    }

    #[test]
    fn section_and_accelerator_are_recorded() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let menu = za::create_menu().unwrap();
            menu.append_section("Recent");
            menu.append_with_shortcut("Save", "app.save", "Ctrl+S");
            let items = menu.items();
            assert_eq!(items[0].kind, rswidgets::backends::zork::MenuItemKind::Section);
            assert!(!items[0].enabled);
            assert_eq!(items[1].accelerator, "Ctrl+S");
        });
    }

    #[test]
    fn append_key_strips_the_mnemonic_marker() {
        // GTK marks a mnemonic with a trailing-ish `&` before the character
        // (`&Save`); `&_` escapes a literal underscore.
        serial(|| {
            let _w = za::create_window().unwrap();
            let menu = za::create_menu().unwrap();
            menu.append_key("&Save", "app.save");
            menu.append_key("&Print", "app.print");
            menu.append_key("A&B", "app.amp");
            let items = menu.items();
            assert_eq!(items[0].label, "Save");
            assert_eq!(items[1].label, "Print");
            assert_eq!(items[2].label, "AB");
        });
    }

    #[test]
    fn menubar_exposes_its_items() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let menu = za::create_menu().unwrap();
            menu.append("Open", "app.open");
            let bar = za::create_menubar(&menu, std::ptr::null_mut()).unwrap();
            assert_eq!(bar.items().len(), 1);
            // No popups in a headless model, so mnemonics report "not handled".
            assert!(!bar.activate_submenu_by_mnemonic(b'f' as u32));
            assert!(!bar.menu_active());
        });
    }
}

mod dialogs {
    use super::*;

    #[test]
    fn response_requires_a_registered_button_and_carries_the_id() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let d = za::create_dialog().unwrap();
            let seen = Rc::new(RefCell::new(Vec::new()));
            {
                let s = seen.clone();
                d.connect_response(move |r| s.borrow_mut().push(r)).unwrap();
            }
            d.respond(1);
            assert!(seen.borrow().is_empty(), "unregistered response must not fire");
            d.add_button("OK", 1);
            d.add_button("Cancel", 2);
            d.respond(2);
            assert_eq!(*seen.borrow(), vec![2]);
            assert_eq!(d.response_ids(), vec![1, 2]);
        });
    }

    #[test]
    fn close_marks_destroyed() {
        // The old `close` was a comment-only no-op, so a closed dialog was
        // indistinguishable from a live one.
        serial(|| {
            let _w = za::create_window().unwrap();
            let d = za::create_dialog().unwrap();
            assert!(!d.is_destroyed());
            d.close();
            assert!(d.is_destroyed());
            assert!(!rswidgets::backends::zork::get_visible(d.id()));
        });
    }

    #[test]
    fn transient_for_is_recorded() {
        serial(|| {
            let w = za::create_window().unwrap();
            let d = za::create_dialog().unwrap();
            d.set_transient_for(*w.as_ref());
            assert_eq!(d.get_transient_for(), Some(w.id()));
        });
    }

    #[test]
    fn content_area_is_the_first_child() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let d = za::create_dialog().unwrap();
            let c = za::create_canvas().unwrap();
            d.append_content_area(&c);
            assert_eq!(d.get_content_area(), c.id() as *mut std::os::raw::c_void);
        });
    }
}

mod textview {
    use super::*;

    #[test]
    fn append_is_incremental_and_readable() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let tv = za::create_textview().unwrap();
            tv.append_text("line 1\n");
            tv.append_text("line 2\n");
            assert_eq!(tv.get_text(), Some("line 1\nline 2\n".to_string()));
            assert_eq!(tv.get_buffer(), tv.get_text());
        });
    }

    #[test]
    fn editable_round_trips() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let tv = za::create_textview().unwrap();
            assert!(tv.is_editable());
            tv.set_editable(false);
            assert!(!tv.is_editable());
        });
    }
}

mod tree_integrity {
    use super::*;

    #[test]
    fn reparenting_leaves_exactly_one_parent_link() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let bx1 = za::create_box(za::Orientation::Vertical, 0).unwrap();
            let bx2 = za::create_box(za::Orientation::Vertical, 0).unwrap();
            let c = za::create_label("c").unwrap();
            bx1.append(&c);
            rswidgets::backends::zork::set_child(bx2.id(), c.id());
            let snap = rswidgets::backends::zork::model::with_state_for_test(|s| s.snapshot());
            let listing: Vec<usize> = snap
                .nodes
                .iter()
                .filter(|n| n.children.contains(&c.id()))
                .map(|n| n.id)
                .collect();
            assert_eq!(listing, vec![bx2.id()], "exactly one parent");
            let node = snap.nodes.iter().find(|n| n.id == c.id()).unwrap();
            assert_eq!(node.parent, Some(bx2.id()));
        });
    }

    #[test]
    fn box_cannot_be_its_own_child() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let bx = za::create_box(za::Orientation::Vertical, 0).unwrap();
            bx.append(&bx);
            let snap = rswidgets::backends::zork::model::with_state_for_test(|s| s.snapshot());
            let node = snap.nodes.iter().find(|n| n.id == bx.id()).unwrap();
            assert!(node.children.is_empty());
        });
    }
}

mod spreadsheet {
    use super::*;

    #[test]
    fn cells_round_trip_through_the_adapter() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let ss = za::create_spreadsheet().unwrap();
            ss.set_cell(1, 1, "42");
            ss.set_raw_cell(1, 2, "=1+1");
            ss.set_cell_style(1, 1, 3);
            ss.set_border_title("Sheet1");
            assert_eq!(ss.get_cell(1, 1), Some("42".to_string()));
            assert_eq!(ss.get_cell(1, 2), Some("=1+1".to_string()));
            assert_eq!(ss.get_border_title(), Some("Sheet1".to_string()));
            let snap = rswidgets::backends::zork::model::with_state_for_test(|s| s.snapshot());
            let entry = snap.nodes.iter().find(|n| n.id == ss.id()).unwrap();
            let styled = entry.cells.iter().find(|c| c.row == 1 && c.col == 1).unwrap();
            assert_eq!(styled.style, 3);
            let raw = entry.cells.iter().find(|c| c.col == 2).unwrap();
            assert!(raw.raw);
        });
    }

    #[test]
    fn spreadsheet_click_hook_fires() {
        serial(|| {
            let _w = za::create_window().unwrap();
            let ss = za::create_spreadsheet().unwrap();
            let seen = Rc::new(RefCell::new((0.0, 0.0)));
            {
                let s = seen.clone();
                ss.on_click(Box::new(move |x, y| *s.borrow_mut() = (x, y)));
            }
            rswidgets::backends::zork::pointer_click(ss.id(), 12.0, 34.0);
            assert_eq!(*seen.borrow(), (12.0, 34.0));
        });
    }
}
