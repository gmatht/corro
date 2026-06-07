use rustxwidgets::prelude::*;
use std::cell::RefCell;
use std::rc::Rc;

const ROWS: usize = 100;
const COLS: usize = 26;
const CELL_W: i32 = 150;
const CELL_H: i32 = 28;

type CellFormat = (bool, bool, u8, String, String);

fn col_label(n: usize) -> String {
    if n < 26 {
        format!("{}", (b'A' + (n as u8)) as char)
    } else {
        let mut s = String::new();
        let mut v = n;
        loop {
            s.insert(0, (b'A' + (v % 26) as u8) as char);
            v /= 26;
            if v == 0 { break; }
            v -= 1;
        }
        s
    }
}

fn compute_row_y(heights: &[i32], row: usize) -> i32 {
    let mut y = 0;
    for i in 0..row {
        if i < heights.len() { y += heights[i]; }
    }
    y
}

fn parse_hex(hex: &str) -> (f64, f64, f64) {
    if hex.len() >= 7 && hex.as_bytes()[0] == b'#' {
        let r = u8::from_str_radix(&hex[1..3], 16).unwrap_or(0) as f64 / 255.0;
        let g = u8::from_str_radix(&hex[3..5], 16).unwrap_or(0) as f64 / 255.0;
        let b = u8::from_str_radix(&hex[5..7], 16).unwrap_or(0) as f64 / 255.0;
        (r, g, b)
    } else {
        (0.0, 0.0, 0.0)
    }
}

fn main() {
    let chw = 46f64;
    let cw = CELL_W as f64;
    let ch = CELL_H as f64;
    let total_w = chw + COLS as f64 * cw;
    let total_h = ch + ROWS as f64 * ch;

    let app = App::init().expect("init");
    let win = app.create_window().expect("window");
    win.set_title("Spreadsheet");
    win.set_default_size(1600, 1000);

    let texts: Rc<RefCell<Vec<Vec<String>>>> = Rc::new(RefCell::new(
        (0..ROWS).map(|_| (0..COLS).map(|_| String::new()).collect()).collect()
    ));
    let fmts: Rc<RefCell<Vec<Vec<CellFormat>>>> = Rc::new(RefCell::new(
        (0..ROWS).map(|_| (0..COLS).map(|_| (false, false, 0u8, "#000000".into(), "#ffffff".into())).collect()).collect()
    ));
    let row_heights: Rc<RefCell<Vec<i32>>> = Rc::new(RefCell::new(vec![CELL_H; ROWS]));
    let sel: Rc<RefCell<Option<(usize, usize)>>> = Rc::new(RefCell::new(Some((0, 0))));
    let editing_entry: Rc<RefCell<Option<Entry>>> = Rc::new(RefCell::new(None));
    let text_input_active: Rc<RefCell<bool>> = Rc::new(RefCell::new(false));

    {
        let mut t = texts.borrow_mut();
        t[0][0] = "Short".into();
        t[1][0] = "VeryLongHeaderThatOverflows".into();
        t[2][0] = "CellWithAVeryLongWordThatWillSpanMultipleCells".into();
        t[3][0] = "NoOverflowHereBcNextIsUsed".into();
        t[3][1] = "X".into();
        t[4][0] = "42".into();
        t[4][1] = "=SUM".into();
        t[5][0] = "TRUE".into();
        t[5][1] = "FALSE".into();
        t[6][0] = "-123.45".into();
        t[7][0] = "Hello".into();
        t[7][1] = "World".into();
        t[8][0] = "Revenue".into();
        t[8][1] = "1000".into();
        t[8][2] = "2000".into();
        t[8][3] = "=SUM(B9:C9)".into();
    }

    let vbox = app.create_box(Orientation::Vertical, 0).expect("vbox");

    let toolbar = app.create_box(Orientation::Horizontal, 2).expect("toolbar");
    let bold_btn = app.create_button("B").expect("bold");
    let italic_btn = app.create_button("I").expect("italic");
    let al_l = app.create_button("AL").expect("al_l");
    let al_c = app.create_button("AC").expect("al_c");
    let al_r = app.create_button("AR").expect("al_r");
    let hl_btn = app.create_button("HL").expect("hl");
    let fg_btn = app.create_button("FG").expect("fg");

    for b in [&bold_btn, &italic_btn, &al_l, &al_c, &al_r, &hl_btn, &fg_btn] { b.set_size_request(28, 24); }
    toolbar.append(&bold_btn);
    toolbar.append(&italic_btn);
    toolbar.append(&al_l);
    toolbar.append(&al_c);
    toolbar.append(&al_r);
    toolbar.append(&hl_btn);
    toolbar.append(&fg_btn);
    let toolbar_box = app.create_box(Orientation::Vertical, 0).expect("toolbar_box");
    toolbar_box.append(&toolbar);

    let formula_bar = app.create_box(Orientation::Horizontal, 4).expect("formula_bar");
    let fx_label = app.create_label("  fx  ").expect("fx_label");
    formula_bar.append(&fx_label);
    let formula_entry = app.create_entry().expect("formula_entry");
    formula_entry.set_width_chars(40);
    formula_entry.set_size_request(400, 26);
    formula_bar.append(&formula_entry);

    let overlay = app.create_overlay().expect("overlay");
    let canvas = app.create_canvas().expect("canvas");
    let total_w_i = total_w as i32;
    let total_h_i = total_h as i32;
    canvas.set_size_request(total_w_i, total_h_i);
    canvas.set_content_size(total_w_i, total_h_i);
    overlay.set_child(&canvas);
    overlay.set_size_request(total_w_i, total_h_i);

    let scrolled = {
        let s = app.create_scrolled_window().expect("scrolled");
        s.set_policy(0, 0);
        s.set_child(&overlay);
        s
    };

    let text_input_active2 = text_input_active.clone();
    let _ = formula_entry.connect_focus_in_event(move |_| { *text_input_active2.borrow_mut() = true; 0 });
    let text_input_active3 = text_input_active.clone();
    let _ = formula_entry.connect_focus_out_event(move |_| { *text_input_active3.borrow_mut() = false; 0 });

    let commit_edit = {
        let text_commit = texts.clone();
        let edit_entry = editing_entry.clone();
        let text_active = text_input_active.clone();
        let formula = formula_entry.clone();
        let cv = canvas.clone();
        let overlay_for_commit = overlay.clone();
        let sel_c = sel.clone();
        move || {
            if let Some(e) = edit_entry.borrow_mut().take() {
                *text_active.borrow_mut() = false;
                let new_text = e.get_text().unwrap_or_default();
                if let Some((r, c)) = *sel_c.borrow() {
                    if r < text_commit.borrow().len() && c < text_commit.borrow()[r].len() {
                        text_commit.borrow_mut()[r][c] = new_text.clone();
                    }
                }
                formula.set_text(&new_text);
                overlay_for_commit.remove(&e);
                cv.queue_redraw();
            }
        }
    };

    let refresh_selection = {
        let texts_ref = texts.clone();
        let formula_ref = formula_entry.clone();
        let sel_ref = sel.clone();
        let cv = canvas.clone();
        move || {
            if let Some((r, c)) = *sel_ref.borrow() {
                if r < texts_ref.borrow().len() && c < texts_ref.borrow()[r].len() {
                    formula_ref.set_text(&texts_ref.borrow()[r][c]);
                }
            }
            cv.queue_redraw();
        }
    };

    let start_edit: Rc<dyn Fn(usize, usize)> = {
        let texts_edit = texts.clone();
        let fmts_edit = fmts.clone();
        let edit_entry = editing_entry.clone();
        let text_active = text_input_active.clone();
        let overlay_edit = overlay.clone();
        let rh = row_heights.clone();
        let commit = commit_edit.clone();
        let sel_edit = sel.clone();
        let refresh = refresh_selection.clone();
        let app_for_edit = app.clone();
        let self_ref: Rc<RefCell<Option<Rc<dyn Fn(usize, usize)>>>> = Rc::new(RefCell::new(None));
        let self_ref2 = self_ref.clone();
        let start: Rc<dyn Fn(usize, usize)> = Rc::new(move |r: usize, c: usize| {
            if edit_entry.borrow().is_some() { return; }
            if let Ok(entry) = app_for_edit.create_entry() {
                *text_active.borrow_mut() = true;
                entry.set_text(&texts_edit.borrow()[r][c]);
                let fmt = &fmts_edit.borrow()[r][c];
                if fmt.0 && fmt.1 { entry.add_class("cell-both"); }
                else if fmt.0 { entry.add_class("cell-bold"); }
                else if fmt.1 { entry.add_class("cell-italic"); }
                match fmt.3.as_str() {
                    "#cc0000" => { entry.add_class("cell-fg-red"); }
                    "#0000cc" => { entry.add_class("cell-fg-blue"); }
                    "#006600" => { entry.add_class("cell-fg-green"); }
                    _ => {}
                }
                let rh_b = rh.borrow();
                let left = chw as i32 + c as i32 * CELL_W;
                let top = CELL_H + compute_row_y(&rh_b, r);
                drop(rh_b);
                overlay_edit.add_overlay(&entry);
                overlay_edit.set_overlay_pass_through(&entry, false);
                entry.set_margin_start(left);
                entry.set_margin_top(top);
                entry.set_halign(1);
                entry.set_valign(1);
                entry.set_visible(true);
                overlay_edit.show_all();
                entry.set_size_request(CELL_W, rh.borrow()[r]);

                let commit2 = commit.clone();
                let sel2 = sel_edit.clone();
                let refresh2 = refresh.clone();
                let start_again = self_ref2.clone();
                let _ = entry.connect_activate(move |_param| {
                    commit2();
                    let current = *sel2.borrow();
                    if let Some((_, _)) = current {
                        let next = current.map(|(row, col)| (row + 1, col)).filter(|(row, _)| *row < ROWS);
                        if let Some((nr, nc)) = next {
                            *sel2.borrow_mut() = Some((nr, nc));
                            refresh2();
                            if let Some(sa) = start_again.borrow().as_ref().cloned() {
                                sa(nr, nc);
                            }
                        }
                    }
                });

                entry.grab_focus();
                *edit_entry.borrow_mut() = Some(entry);
            }
        });
        *self_ref.borrow_mut() = Some(start.clone());
        start
    };

    let t_draw = texts.clone();
    let f_draw = fmts.clone();
    let sel_draw = sel.clone();
    let edit_draw = editing_entry.clone();
    canvas.set_draw_callback(Box::new(move |ctx: &mut dyn DrawContext, _w: i32, _h: i32| {
        let t = t_draw.borrow();
        let f = f_draw.borrow();
        let mut overflow_end_col = vec![vec![0usize; COLS]; ROWS];
        for r in 0..ROWS {
            for c in 0..COLS {
                let text = &t[r][c];
                if text.is_empty() { continue; }
                let cb = f[r][c].0;
                let ci = f[r][c].1;
                let (_, _, tw, _) = ctx.text_extents_styled(text, "monospace", 13.0, if ci { 1 } else { 0 }, if cb { 1 } else { 0 });
                let cx = chw + c as f64 * cw;
                if cx + tw > cx + cw {
                    let mut lo = c;
                    for oc in (c + 1)..COLS {
                        if !t[r][oc].is_empty() { break; }
                        lo = oc;
                    }
                    overflow_end_col[r][c] = lo;
                }
            }
        }

        ctx.clear(1.0, 1.0, 1.0, 1.0);

        ctx.fill_rect(0.0, 0.0, chw, ch, 0.91, 0.91, 0.91, 1.0);
        ctx.fill_rect(chw, 0.0, COLS as f64 * cw, ch, 0.8, 0.8, 0.8, 1.0);
        ctx.fill_rect(0.0, ch, chw, ROWS as f64 * ch, 0.8, 0.8, 0.8, 1.0);

        for c in 0..COLS {
            let lbl = col_label(c);
            let (xb, _, w, _) = ctx.text_extents_styled(&lbl, "monospace", 12.0, 0, 0);
            let x = chw + c as f64 * cw + cw / 2.0 - xb - w / 2.0;
            ctx.draw_text(x, ch / 2.0, &lbl, "monospace", 12.0, 0.0, 0.0, 0.0, 1.0);
        }
        for r in 0..ROWS {
            let lbl = format!("{}", r + 1);
            let (xb, _, w, _) = ctx.text_extents_styled(&lbl, "monospace", 12.0, 0, 0);
            let x = chw / 2.0 - xb - w / 2.0;
            ctx.draw_text(x, ch + r as f64 * ch + ch / 2.0, &lbl, "monospace", 12.0, 0.0, 0.0, 0.0, 1.0);
        }

        for r in 0..=ROWS {
            let y = ch + r as f64 * ch;
            ctx.stroke_rect(0.0, y, total_w, 0.0, 0.7, 0.7, 0.7, 1.0, 0.5);
        }
        for bc in 0..=COLS {
            let x = chw + bc as f64 * cw;
            ctx.stroke_rect(x, 0.0, 0.0, ch, 0.7, 0.7, 0.7, 1.0, 0.5);
            for r in 0..ROWS {
                let skip = bc > 0 && overflow_end_col[r][bc - 1] >= bc;
                if !skip {
                    let y1 = ch + r as f64 * ch;
                    let y2 = ch + (r + 1) as f64 * ch;
                    ctx.stroke_rect(x, y1, 0.0, y2 - y1, 0.7, 0.7, 0.7, 1.0, 0.5);
                }
            }
        }

        let editing = edit_draw.borrow().is_some();
        let edit_target = if editing { *sel_draw.borrow() } else { None };
        for r in 0..ROWS {
            for c in 0..COLS {
                let text = &t[r][c];
                if text.is_empty() { continue; }
                if edit_target == Some((r, c)) { continue; }
                let (bold, italic, align, fg_hex, bg_hex) = &f[r][c];
                let cx = chw + c as f64 * cw;
                let cy = ch + r as f64 * ch;

                if bg_hex != "#ffffff" {
                    let (br, bg, bb) = parse_hex(bg_hex);
                    ctx.fill_rect(cx, cy + 1.0, cw, ch - 2.0, br, bg, bb, 1.0);
                }

                let (fr, fgg, fb) = if fg_hex != "#000000" { parse_hex(fg_hex) } else { (0.0, 0.0, 0.0) };
                let slant = if *italic { 1 } else { 0 };
                let weight = if *bold { 1 } else { 0 };
                let (xb, _, tw, _) = ctx.text_extents(text, "monospace", 13.0);
                let pad = 4.0;
                let tx = match *align {
                    0 => cx + pad - xb,
                    1 => cx + cw / 2.0 - xb - tw / 2.0,
                    _ => cx + cw - pad - xb - tw,
                };
                let ty = cy + ch / 2.0;

                let last_oc = overflow_end_col[r][c];
                let text_right = tx + xb + tw + 2.0;
                let clip_right = if last_oc > c {
                    let mut limit = text_right;
                    for oc in (c + 1)..COLS {
                        if !t[r][oc].is_empty() {
                            limit = limit.min(chw + oc as f64 * cw);
                            break;
                        }
                    }
                    limit.min(total_w)
                } else { cx + cw };

                ctx.save();
                ctx.clip(cx, cy, clip_right - cx, ch);
                ctx.draw_text_styled(tx, ty, text, "monospace", 13.0, fr, fgg, fb, 1.0, slant, weight);
                ctx.restore();
            }
        }

        if let Some((sr, sc)) = *sel_draw.borrow() {
            let sx = chw + sc as f64 * cw;
            let sy = ch + sr as f64 * ch;
            ctx.fill_rect(sx, sy, cw, ch, 0.83, 0.91, 1.0, 0.3);
            ctx.stroke_rect(sx + 0.5, sy + 0.5, cw - 1.0, ch - 1.0, 0.1, 0.45, 0.91, 1.0, 2.0);
        }
    }));

    let sel_click = sel.clone();
    let fe_click = formula_entry.clone();
    let txt_click = texts.clone();
    let edit_click = editing_entry.clone();
    let commit_click = commit_edit.clone();
    let start_click = start_edit.clone();
    let fmts_click = fmts.clone();
    let rh_click = row_heights.clone();
    let cv_click = canvas.clone();
    canvas.on_click(Box::new(move |x: f64, y: f64| {
        let col = {
            let cx = x - chw;
            if cx <= 0.0 { 0 } else { ((cx / cw) as usize).min(COLS - 1) }
        };
        let rh_b = rh_click.borrow();
        let grid_row = if y < ch {
            drop(rh_b);
            0
        } else {
            let data_y = y - ch;
            let mut acc = 0.0f64;
            let mut ri = 0;
            for hi in rh_b.iter() {
                if acc + *hi as f64 > data_y { break; }
                acc += *hi as f64;
                ri += 1;
            }
            drop(rh_b);
            (ri + 1).min(ROWS)
        };

        if grid_row == 0 && col < COLS { return; }
        if x < chw && grid_row >= 1 && grid_row <= ROWS { return; }

        let data_row = grid_row.saturating_sub(1);
        if col < COLS && data_row < ROWS {
            commit_click();
            *sel_click.borrow_mut() = Some((data_row, col));
            let t = txt_click.borrow();
            let fmt = fmts_click.borrow()[data_row][col].clone();
            if data_row < t.len() && col < t[data_row].len() {
                let text = &t[data_row][col];
                if fmt.0 { fe_click.set_text(&format!("*{}*", text)); }
                else if fmt.1 { fe_click.set_text(&format!("/{}/", text)); }
                else { fe_click.set_text(text); }
            }
            drop(t);
            cv_click.queue_redraw();
            if edit_click.borrow().is_none() {
                start_click(data_row, col);
                if let Some(e) = edit_click.borrow().as_ref() {
                    e.set_text(&txt_click.borrow()[data_row][col]);
                }
            }
        }
    }));

    let sel_kb = sel.clone();
    let edit_kb = editing_entry.clone();
    let text_active_kb = text_input_active.clone();
    let commit_kb = commit_edit.clone();
    let start_kb = start_edit.clone();
    let refresh_kb = refresh_selection.clone();
    canvas.on_key(Box::new(move |keyval: u32| -> bool {
        if *text_active_kb.borrow() { return false; }
        if edit_kb.borrow().is_some() {
            if keyval == 27 { // Escape
                let _ = edit_kb.borrow_mut().take();
                return true;
            } else if keyval == 13 || keyval == 10 { // Enter
                commit_kb();
                if let Some((r, c)) = sel_kb.borrow().map(|(r, c)| (r + 1, c)).filter(|(r, _)| *r < ROWS) {
                    *sel_kb.borrow_mut() = Some((r, c));
                    refresh_kb();
                    start_kb(r, c);
                }
                return true;
            }
            return false;
        }
        let mut coord = sel_kb.borrow_mut();
        if let Some((r, c)) = *coord {
            match keyval {
                38 | 119 => { if r > 0 { *coord = Some((r - 1, c)); } }
                40 | 115 => { if r + 1 < ROWS { *coord = Some((r + 1, c)); } }
                37 | 97 => { if c > 0 { *coord = Some((r, c - 1)); } }
                39 | 100 => { if c + 1 < COLS { *coord = Some((r, c + 1)); } }
                13 | 10 => {
                    drop(coord);
                    commit_kb();
                    start_kb(r, c);
                    return true;
                }
                _ => {
                    if keyval >= 32 && keyval <= 126 {
                        drop(coord);
                        commit_kb();
                        start_kb(r, c);
                        if let Some(entry) = edit_kb.borrow().as_ref() {
                            if let Some(ch) = std::char::from_u32(keyval) {
                                entry.set_text(&ch.to_string());
                            }
                        }
                        return true;
                    }
                }
            }
        }
        drop(coord);
        refresh_kb();
        false
    }));

    fn apply_fmt(entry: &Entry, r: usize, c: usize, fmts: &RefCell<Vec<Vec<CellFormat>>>) {
        let b = fmts.borrow();
        if r >= b.len() || c >= b[r].len() { return; }
        let fmt = &b[r][c];
        entry.remove_class("cell-bold");
        entry.remove_class("cell-italic");
        entry.remove_class("cell-both");
        entry.remove_class("cell-fg-red");
        entry.remove_class("cell-fg-blue");
        entry.remove_class("cell-fg-green");
        if fmt.0 && fmt.1 { entry.add_class("cell-both"); }
        else if fmt.0 { entry.add_class("cell-bold"); }
        else if fmt.1 { entry.add_class("cell-italic"); }
        match fmt.3.as_str() {
            "#cc0000" => entry.add_class("cell-fg-red"),
            "#0000cc" => entry.add_class("cell-fg-blue"),
            "#006600" => entry.add_class("cell-fg-green"),
            _ => {}
        }
    }

    let edit_entry_fmt = editing_entry.clone();
    let sel_fmt = sel.clone();
    let fmts_fmt = fmts.clone();
    let cv_fmt = canvas.clone();
    let _ = bold_btn.on_click(move || {
        let coord = *sel_fmt.borrow();
        if let Some((r, c)) = coord {
            let mut f = fmts_fmt.borrow_mut();
            if r < f.len() && c < f[r].len() { f[r][c].0 = !f[r][c].0; cv_fmt.queue_redraw(); }
        }
        if let Some((r, c)) = coord {
            if let Some(ee) = edit_entry_fmt.borrow().as_ref() {
                apply_fmt(ee, r, c, &fmts_fmt);
            }
        }
    });

    let edit_entry_ital = editing_entry.clone();
    let sel_ital = sel.clone();
    let fmts_ital = fmts.clone();
    let cv_ital = canvas.clone();
    let _ = italic_btn.on_click(move || {
        let coord = *sel_ital.borrow();
        if let Some((r, c)) = coord {
            let mut f = fmts_ital.borrow_mut();
            if r < f.len() && c < f[r].len() { f[r][c].1 = !f[r][c].1; cv_ital.queue_redraw(); }
        }
        if let Some((r, c)) = coord {
            if let Some(ee) = edit_entry_ital.borrow().as_ref() {
                apply_fmt(ee, r, c, &fmts_ital);
            }
        }
    });

    let _ = al_l.on_click({
        let sel_al = sel.clone(); let fmts_al = fmts.clone(); let cv_al = canvas.clone();
        let ed_al = editing_entry.clone();
        move || {
            let coord = *sel_al.borrow();
            if let Some((r, c)) = coord { fmts_al.borrow_mut()[r][c].2 = 0; cv_al.queue_redraw(); }
            if let Some((r, c)) = coord {
                if let Some(ee) = ed_al.borrow().as_ref() { apply_fmt(ee, r, c, &fmts_al); }
            }
        }
    });
    let _ = al_c.on_click({
        let sel_al = sel.clone(); let fmts_al = fmts.clone(); let cv_al = canvas.clone();
        let ed_ac = editing_entry.clone();
        move || {
            let coord = *sel_al.borrow();
            if let Some((r, c)) = coord { fmts_al.borrow_mut()[r][c].2 = 1; cv_al.queue_redraw(); }
            if let Some((r, c)) = coord {
                if let Some(ee) = ed_ac.borrow().as_ref() { apply_fmt(ee, r, c, &fmts_al); }
            }
        }
    });
    let _ = al_r.on_click({
        let sel_al = sel.clone(); let fmts_al = fmts.clone(); let cv_al = canvas.clone();
        let ed_ar = editing_entry.clone();
        move || {
            let coord = *sel_al.borrow();
            if let Some((r, c)) = coord { fmts_al.borrow_mut()[r][c].2 = 2; cv_al.queue_redraw(); }
            if let Some((r, c)) = coord {
                if let Some(ee) = ed_ar.borrow().as_ref() { apply_fmt(ee, r, c, &fmts_al); }
            }
        }
    });

    let edit_entry_hl = editing_entry.clone();
    let sel_hl = sel.clone();
    let fmts_hl = fmts.clone();
    let cv_hl = canvas.clone();
    let _ = hl_btn.on_click(move || {
        let coord = *sel_hl.borrow();
        if let Some((r, c)) = coord {
            let mut f = fmts_hl.borrow_mut();
            if r < f.len() && c < f[r].len() {
                let bg = f[r][c].4.clone();
                f[r][c].4 = if bg == "#ffff00" { "#88ff88".into() } else if bg == "#88ff88" { "#ffffff".into() } else { "#ffff00".into() };
                cv_hl.queue_redraw();
            }
        }
        if let Some((r, c)) = coord {
            if let Some(ee) = edit_entry_hl.borrow().as_ref() {
                apply_fmt(ee, r, c, &fmts_hl);
            }
        }
    });
    let edit_entry_fg = editing_entry.clone();
    let sel_fg = sel.clone();
    let fmts_fg = fmts.clone();
    let cv_fg = canvas.clone();
    let _ = fg_btn.on_click(move || {
        let coord = *sel_fg.borrow();
        if let Some((r, c)) = coord {
            let mut f = fmts_fg.borrow_mut();
            if r < f.len() && c < f[r].len() {
                let fg = f[r][c].3.clone();
                f[r][c].3 = if fg == "#000000" { "#cc0000".into() } else if fg == "#cc0000" { "#0000cc".into() } else if fg == "#0000cc" { "#006600".into() } else { "#000000".into() };
                cv_fg.queue_redraw();
            }
        }
        if let Some((r, c)) = coord {
            if let Some(ee) = edit_entry_fg.borrow().as_ref() {
                apply_fmt(ee, r, c, &fmts_fg);
            }
        }
    });

    vbox.append(&toolbar_box);
    vbox.append(&formula_bar);
    vbox.append(&scrolled);
    vbox.set_child_vexpand(&scrolled, true);
    vbox.set_child_hexpand(&scrolled, true);
    win.set_child(&vbox);
    win.present();

    web_sys::console::log_1(&"WASM x-sheet running".into());
    app.run().expect("run");
}
