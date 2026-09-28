//! Interactive text REPL driver over the [`crate::backends::zork::model`].
//!
//! This is deliberately *one* of possibly many drivers of the pure model. It is
//! kept for manual exploration and demos. For automated tests prefer
//! [`crate::backends::zork::harness::Harness`].

use std::io::{self, BufRead, Write};

use crate::backends::BackendApp;
use crate::backends::zork::model::{ZorkKind, ZorkNode, ZorkState};

pub struct ZorkApp {
    state: ZorkState,
}

impl ZorkApp {
    pub fn new() -> Self {
        ZorkApp { state: ZorkState::new() }
    }
}

impl Default for ZorkApp {
    fn default() -> Self {
        Self::new()
    }
}

impl BackendApp for ZorkApp {
    /// A REPL owns its input loop: nothing else dispatches for us.
    fn owns_event_loop(&self) -> bool {
        true
    }

    fn run(mut self: Box<Self>) -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
        if self.state.nodes.is_empty() {
            return Ok(());
        }
        self.state.running = true;

        let mut rl = rustyline::DefaultEditor::new()?;

        while self.state.running {
            if let Some(node) = self.state.node(self.state.current_id) {
                describe_room(&self.state, node);
            }

            let prompt = "> ";
            let readline = rl.readline(prompt);
            match readline {
                Ok(line) => {
                    let trimmed = line.trim().to_string();
                    if !trimmed.is_empty() {
                        rl.add_history_entry(&trimmed)?;
                        execute_command(&mut self.state, &trimmed);
                    }
                }
                Err(rustyline::error::ReadlineError::Interrupted)
                | Err(rustyline::error::ReadlineError::Eof) => {
                    self.state.running = false;
                }
                Err(e) => {
                    eprintln!("Error: {}", e);
                    break;
                }
            }
        }

        Ok(())
    }
}

fn short_desc(_state: &ZorkState, node: &ZorkNode) -> String {
    match &node.kind {
        ZorkKind::Window { title } => format!("Window \"{}\"", title),
        ZorkKind::Button { label } => format!("Button \"{}\"", label),
        ZorkKind::Label { text } => {
            let preview: String = text.chars().take(30).collect();
            let preview = if text.chars().count() > 30 { format!("{}...", preview) } else { text.clone() };
            format!("Label: \"{}\"", preview)
        }
        ZorkKind::Entry { .. } => "Entry".into(),
        ZorkKind::CheckButton { label, .. } => format!("CheckButton \"{}\"", label),
        ZorkKind::RadioButton { label, .. } => format!("RadioButton \"{}\"", label),
        ZorkKind::BoxWidget { horizontal, .. } => format!("{} Box", if *horizontal { "Horizontal" } else { "Vertical" }),
        ZorkKind::Grid { .. } => "Grid".into(),
        ZorkKind::Dialog { title } => format!("Dialog \"{}\"", title),
        ZorkKind::DropDown { items, selected } => {
            let current = selected.and_then(|s| items.get(s)).map(|s| s.as_str()).unwrap_or("(none)");
            format!("DropDown [{}]", current)
        }
        ZorkKind::TextView { .. } => "TextView".into(),
        ZorkKind::Canvas { .. } => "Canvas".into(),
        ZorkKind::Overlay => "Overlay".into(),
        ZorkKind::ScrolledWindow => "ScrolledWindow".into(),
        ZorkKind::Fixed => "Fixed".into(),
        ZorkKind::Application => "Application".into(),
        ZorkKind::Spreadsheet { .. } => "Spreadsheet".into(),
        ZorkKind::Menu => "Menu".into(),
        ZorkKind::MenuBar => "MenuBar".into(),
        ZorkKind::SimpleAction => "SimpleAction".into(),
    }
}

/// A one-line property summary, printed by `props` and `look -v`.
///
/// The model records geometry, visibility, expansion, margins, classes, focus
/// and scroll for every node; before this the REPL could not show any of them,
/// so a `set` had no visible effect from the player's side.
fn prop_summary(state: &ZorkState, node: &ZorkNode) -> String {
    let p = &node.props;
    let mut out: Vec<String> = Vec::new();
    if let (Some(x), Some(y)) = (p.offset_x, p.offset_y) {
        out.push(format!("at ({}, {})", x, y));
    }
    if let (Some(w), Some(h)) = (p.width, p.height) {
        out.push(format!("size {}x{}", w, h));
    }
    if !p.visible {
        out.push("hidden".to_string());
    }
    if p.hexpand {
        out.push("hexpand".to_string());
    }
    if p.vexpand {
        out.push("vexpand".to_string());
    }
    if p.margin_start != 0 {
        out.push(format!("margin-start {}", p.margin_start));
    }
    if p.margin_top != 0 {
        out.push(format!("margin-top {}", p.margin_top));
    }
    if !p.classes.is_empty() {
        out.push(format!("classes [{}]", p.classes.join(", ")));
    }
    if p.fixed_width.is_some() {
        out.push(format!("fixed-width {:?}", p.fixed_width));
    }
    if let Some(f) = &p.font {
        out.push(format!("font {:?} @ {}", f, p.font_size));
    }
    if p.hscroll != 0.0 || p.vscroll != 0.0 {
        out.push(format!("scroll ({}, {})", p.hscroll, p.vscroll));
    }
    if state.has_focus(node.id) {
        out.push("focused".to_string());
    }
    if node.destroyed {
        out.push("destroyed".to_string());
    }
    if out.is_empty() {
        "no properties set".to_string()
    } else {
        out.join(", ")
    }
}

fn dir_name(dirs: &[(&str, &ZorkNode)], id: usize) -> String {
    dirs.iter()
        .find(|(_, t)| t.id == id)
        .map(|(d, _)| (*d).to_string())
        .unwrap_or_else(|| "?".to_string())
}

fn describe_room(state: &ZorkState, node: &ZorkNode) {
    println!();
    match &node.kind {
        ZorkKind::Window { title } => {
            println!("You are in a Window{}", if title.is_empty() { ".".to_string() } else { format!(" titled \"{}\".", title) });
        }
        ZorkKind::Dialog { title } => {
            println!("You are in a Dialog{}", if title.is_empty() { ".".to_string() } else { format!(" titled \"{}\".", title) });
        }
        ZorkKind::Button { label } => {
            println!("You are standing on a Button labeled \"{}\".", label);
        }
        ZorkKind::Label { text } => {
            let preview: String = text.chars().take(40).collect();
            let preview = if text.chars().count() > 40 { format!("{}...", preview) } else { text.clone() };
            println!("You are looking at a Label: \"{}\".", preview);
        }
        ZorkKind::Entry { buffer, .. } => {
            let preview: String = buffer.chars().take(40).collect();
            let preview = if buffer.is_empty() {
                "empty".into()
            } else if buffer.chars().count() > 40 {
                format!("{}...", preview)
            } else {
                buffer.clone()
            };
            println!("You are at an Entry containing \"{}\".", preview);
        }
        ZorkKind::CheckButton { label, checked } => {
            println!("You are at a CheckButton labeled \"{}\" ({}).", label, if *checked { "checked" } else { "unchecked" });
        }
        ZorkKind::RadioButton { label, checked, .. } => {
            println!("You are at a RadioButton labeled \"{}\" ({}).", label, if *checked { "selected" } else { "not selected" });
        }
        ZorkKind::DropDown { items, selected } => {
            let current = selected.and_then(|s| items.get(s)).map(|s| s.as_str()).unwrap_or("(none)");
            println!("You are at a DropDown. Current selection: \"{}\".", current);
        }
        ZorkKind::TextView { text } => {
            let preview: String = text.chars().take(40).collect();
            let preview = if text.chars().count() > 40 { format!("{}...", preview) } else { text.clone() };
            println!("You are reading a TextView: \"{}\".", preview);
        }
        ZorkKind::BoxWidget { horizontal, .. } => {
            println!("You are in a {} Box.", if *horizontal { "horizontal" } else { "vertical" });
        }
        ZorkKind::Grid { .. } => {
            println!("You are in a Grid layout.");
        }
        ZorkKind::MenuBar => {
            println!("You are at a MenuBar.");
            let items = state.menu_items.get(&node.id).cloned().unwrap_or_default();
            if !items.is_empty() {
                println!("Menu items:");
                for (i, item) in items.iter().enumerate() {
                    println!("  {}. {}", i + 1, item.label);
                }
            }
        }
        ZorkKind::Menu => {
            println!("You are at a Menu.");
            let items = state.menu_items.get(&node.id).cloned().unwrap_or_default();
            if !items.is_empty() {
                println!("Items:");
                for (i, item) in items.iter().enumerate() {
                    println!("  {}. {}", i + 1, item.label);
                }
            }
        }
        _ => {
            println!("You are in an unknown widget.");
        }
    }

    let parent_dir = node.parent.and_then(|pid| state.node(pid));
    let children: Vec<&ZorkNode> = node.children.iter().filter_map(|cid| state.node(*cid)).collect();
    let siblings: Vec<&ZorkNode> = if let Some(pid) = node.parent {
        state.node(pid).map(|p| p.children.iter().filter_map(|cid| state.node(*cid)).collect()).unwrap_or_default()
    } else {
        Vec::new()
    };

    let mut my_idx = None;
    for (i, sib) in siblings.iter().enumerate() {
        if sib.id == node.id {
            my_idx = Some(i);
            break;
        }
    }

    let mut dirs: Vec<(&str, &ZorkNode)> = Vec::new();
    if let Some(idx) = my_idx {
        if idx > 0 {
            if let Some(prev) = siblings.get(idx - 1) {
                dirs.push(("west", prev));
            }
        }
        if idx + 1 < siblings.len() {
            if let Some(next) = siblings.get(idx + 1) {
                dirs.push(("east", next));
            }
        }
    }
    if !children.is_empty() {
        if let Some(first) = children.first() {
            dirs.push(("north", first));
        }
    }
    if parent_dir.is_some() {
        dirs.push(("south", parent_dir.unwrap()));
    }

    if !dirs.is_empty() {
        println!();
        for (dir, target) in &dirs {
            let desc = short_desc(state, target);
            println!("To the {}: {}", dir, desc);
        }
    }

    let summary = prop_summary(state, node);
    if summary != "no properties set" {
        println!();
        println!("({})", summary);
    }

    println!();
    println!("You see:");
    println!("  0. (yourself) {}", short_desc(state, node));
    let mut idx = 1;
    for (_, target) in &dirs {
        println!("  {}. {} ({})", idx, short_desc(state, target), dir_name(&dirs, target.id));
        idx += 1;
    }
    if let Some(pid) = node.parent {
        if let Some(p) = state.node(pid) {
            for cid in &p.children {
                if *cid == node.id {
                    continue;
                }
                if dirs.iter().any(|(_, t)| t.id == *cid) {
                    continue;
                }
                if let Some(other) = state.node(*cid) {
                    println!("  {}. {}", idx, short_desc(state, other));
                    idx += 1;
                }
            }
        }
    }
    println!();
}

fn execute_command(state: &mut ZorkState, line: &str) {
    let parts: Vec<&str> = line.split_whitespace().collect();
    if parts.is_empty() {
        return;
    }

    let cmd = parts[0].to_lowercase();
    let args: Vec<&str> = parts[1..].to_vec();

    match cmd.as_str() {
        "look" | "l" => {
            if let Some(node) = state.node(state.current_id) {
                describe_room(state, node);
            }
        }
        "go" | "g" => {
            if args.is_empty() {
                println!("Go where? (north/south/east/west or a number)");
                return;
            }
            let dir = args[0].to_lowercase();
            navigate_to(state, &dir);
        }
        "north" | "n" => navigate_to(state, "north"),
        "south" | "s" => navigate_to(state, "south"),
        "east" | "e" => navigate_to(state, "east"),
        "west" | "w" => navigate_to(state, "west"),
        "click" | "press" => {
            let target_id = if !args.is_empty() {
                resolve_arg_to_id(state, args[0])
            } else {
                Some(state.current_id)
            };
            if let Some(id) = target_id {
                let kind_is_menu = state.node(id).map(|n| matches!(n.kind, ZorkKind::MenuBar | ZorkKind::Menu)).unwrap_or(false);
                if kind_is_menu {
                    {
                        let items = state.menu_items.get(&id).cloned().unwrap_or_default();
                        if items.is_empty() {
                            println!("This menu has no items.");
                        } else {
                            println!("Select an item:");
                            for (i, item) in items.iter().enumerate() {
                                let mark = match item.kind {
                                    crate::backends::zork::model::MenuItemKind::Separator => " (separator)",
                                    crate::backends::zork::model::MenuItemKind::Section => " (section)",
                                    crate::backends::zork::model::MenuItemKind::Check => {
                                        if item.checked { " [x]" } else { " [ ]" }
                                    }
                                    crate::backends::zork::model::MenuItemKind::Radio => {
                                        if item.checked { " (o)" } else { " ( )" }
                                    }
                                    crate::backends::zork::model::MenuItemKind::Normal => "",
                                };
                                println!("  {}. {}{}{}", i + 1, item.label, mark, if item.accelerator.is_empty() { String::new() } else { format!("  ({})", item.accelerator) });
                            }
                        }
                    }
                    return;
                }
                // Everything else that can be activated goes through
                // `pointer_click`, so its click hooks run before its
                // callbacks. The old match only handled Button and said
                // "You can't click that" for an Entry or CheckButton, which do
                // have click semantics.
                let clickable = state.node(id).is_some_and(|n| {
                    matches!(
                        n.kind,
                        ZorkKind::Button { .. }
                            | ZorkKind::Entry { .. }
                            | ZorkKind::CheckButton { .. }
                            | ZorkKind::RadioButton { .. }
                            | ZorkKind::DropDown { .. }
                            | ZorkKind::TextView { .. }
                            | ZorkKind::Canvas { .. }
                    )
                });
                if !clickable {
                    println!("You can't click that.");
                    return;
                }
                let desc = short_desc(state, state.node(id).unwrap());
                // A toggleable button flips; anything else just activates.
                if state.node(id).is_some_and(|n| {
                    matches!(n.kind, ZorkKind::CheckButton { .. } | ZorkKind::RadioButton { .. })
                }) {
                    state.toggle(id);
                    println!("You press the {}.", desc);
                } else {
                    state.pointer_click(id, 0.0, 0.0);
                    println!("You press the {}. It clicks!", desc);
                }
            }
        }
        "select" | "choose" => {
            if args.is_empty() {
                println!("Select what? Use 'select <number>'.");
                return;
            }
            let num: usize = match args[0].parse() {
                Ok(n) => n,
                Err(_) => {
                    println!("Usage: select <number>");
                    return;
                }
            };
            // Route through `menu_select`, which fires the SimpleAction the
            // item names. The old code only printed the action name, so nothing
            // ever happened when you picked an item.
            let idx = num.saturating_sub(1);
            match state.menu_select(state.current_id, idx) {
                Some(label) => println!("You selected \"{}\".", label),
                None => println!("Nothing selectable there."),
            }
        }
        "examine" | "exam" | "x" => {
            let target_id = if !args.is_empty() {
                resolve_arg_to_id(state, args[0])
            } else {
                Some(state.current_id)
            };
            if let Some(id) = target_id {
                examine(state, id);
            }
        }
        "type" | "write" => {
            if !matches!(state.node(state.current_id).map(|n| &n.kind), Some(ZorkKind::Entry { .. })) {
                println!("You are not at an Entry. Navigate to an Entry first.");
                return;
            }
            println!("(Enter text, press Enter when done)");
            print!("> ");
            io::stdout().flush().ok();
            let mut input = String::new();
            io::stdin().lock().read_line(&mut input).ok();
            let text = input.trim().to_string();
            state.set_entry_text(state.current_id, &text);
            println!("You inscribed \"{}\" into the Entry.", text);
            state.click(state.current_id);
        }
        "read" => {
            examine(state, state.current_id);
        }
        "toggle" => {
            // `ZorkState::toggle` is group-aware, so using it here means a
            // radio toggle clears its siblings (the old inline toggle just
            // flipped this button and left the group inconsistent).
            let id = state.current_id;
            let is_toggleable = state.node(id).is_some_and(|n| {
                matches!(n.kind, ZorkKind::CheckButton { .. } | ZorkKind::RadioButton { .. })
            });
            if !is_toggleable {
                println!("You can't toggle that.");
                return;
            }
            let was = match state.node(id) {
                Some(n) => match &n.kind {
                    ZorkKind::CheckButton { checked, .. } | ZorkKind::RadioButton { checked, .. } => *checked,
                    _ => false,
                },
                None => false,
            };
            state.toggle(id);
            let now = match state.node(id) {
                Some(n) => match &n.kind {
                    ZorkKind::CheckButton { checked, .. } | ZorkKind::RadioButton { checked, .. } => *checked,
                    _ => false,
                },
                None => false,
            };
            let noun = match state.node(id).map(|n| &n.kind) {
                Some(ZorkKind::CheckButton { .. }) => "CheckButton",
                Some(ZorkKind::RadioButton { .. }) => "RadioButton",
                _ => "widget",
            };
            println!("{} is now {}.", noun, if now { "on" } else { "off" });
            // A radio group only ever has one member on; if this toggle turned
            // something on, say which sibling it cleared.
            if now && !was {
                if let Some(ZorkKind::RadioButton { group_id, .. }) = state.node(id).map(|n| &n.kind) {
                    if *group_id != 0 {
                        let others_off = state
                            .nodes
                            .iter()
                            .filter(|n| matches!(n.kind, ZorkKind::RadioButton { group_id: g, .. } if g == *group_id && n.id != id))
                            .all(|n| !matches!(n.kind, ZorkKind::RadioButton { checked: true, .. }));
                        if others_off {
                            println!("The rest of the group is now unselected.");
                        }
                    }
                }
            }
        }
        "props" | "attributes" => {
            if let Some(node) = state.node(state.current_id) {
                println!("{}: {}", short_desc(state, node), prop_summary(state, node));
                if let ZorkKind::Entry { buffer, cursor } = &node.kind {
                    println!("  caret at {} of {} chars", cursor, buffer.chars().count());
                }
                if let ZorkKind::Canvas { .. } = &node.kind {
                    println!("  redraw requests: {}", node.pointer.redraws);
                }
                if !node.dialog_buttons.is_empty() {
                    let btns: Vec<String> = node
                        .dialog_buttons
                        .iter()
                        .map(|(l, r)| format!("{}={}", l, r))
                        .collect();
                    println!("  buttons: {}", btns.join(", "));
                }
                if !node.cells.is_empty() {
                    println!("  {} cell(s) set", node.cells.len());
                }
            }
        }
        "set" => {
            if args.len() < 2 {
                println!("Usage: set <property> <value>");
                println!("  properties: size, offset, margin-start, margin-top, class, xalign, fixed-width, font, scroll");
                return;
            }
            let prop = args[0].to_lowercase();
            let id = state.current_id;
            // The value is the rest of the line, so a class name or font can
            // contain spaces.
            let value = args[1..].join(" ");
            match prop.as_str() {
                "size" => {
                    if let Some((w, h)) = parse_pair(&value) {
                        state.set_size_request(id, w, h);
                        println!("Size set to {}x{}.", w, h);
                    } else {
                        println!("Usage: set size <w> <h>");
                    }
                }
                "offset" => {
                    if let Some((x, y)) = parse_pair(&value) {
                        state.set_offset(id, x, y);
                        println!("Offset set to ({}, {}).", x, y);
                    } else {
                        println!("Usage: set offset <x> <y>");
                    }
                }
                "margin-start" => {
                    if let Some(px) = value.parse::<i32>().ok() {
                        state.set_margin_start(id, px);
                        println!("Margin-start set to {}.", px);
                    } else {
                        println!("Usage: set margin-start <px>");
                    }
                }
                "margin-top" => {
                    if let Some(px) = value.parse::<i32>().ok() {
                        state.set_margin_top(id, px);
                        println!("Margin-top set to {}.", px);
                    } else {
                        println!("Usage: set margin-top <px>");
                    }
                }
                "class" => {
                    if let Some(cls) = value.strip_suffix('+') {
                        state.add_class(id, cls);
                        println!("Added class \"{}\".", cls);
                    } else if let Some(cls) = value.strip_suffix('-') {
                        state.remove_class(id, cls);
                        println!("Removed class \"{}\".", cls);
                    } else {
                        state.add_class(id, &value);
                        println!("Added class \"{}\" (use 'set class <name>-' to remove).", value);
                    }
                }
                "xalign" => {
                    if let Some(x) = value.parse::<f32>().ok() {
                        state.set_xalign(id, x);
                        println!("x-align set to {}.", x);
                    } else {
                        println!("Usage: set xalign <0.0-1.0>");
                    }
                }
                "fixed-width" => {
                    match value.parse::<i32>() {
                        Ok(w) => {
                            state.set_fixed_width(id, Some(w));
                            println!("Width pinned to {}.", w);
                        }
                        // "none" / "off" release the pin.
                        Err(_) if matches!(value.as_str(), "none" | "off") => {
                            state.set_fixed_width(id, None);
                            println!("Width pin released.");
                        }
                        _ => println!("Usage: set fixed-width <px> | none"),
                    }
                }
                "font" => {
                    // "font <size>" or "font <family> <size>".
                    let mut parts = value.rsplitn(2, ' ');
                    let size = parts.next().and_then(|s| s.parse::<f64>().ok());
                    let family = parts.next().unwrap_or("sans");
                    match size {
                        Some(sz) => {
                            state.set_font_style(id, Some(family), sz);
                            println!("Font set to {} @ {}.", family, sz);
                        }
                        None => println!("Usage: set font [family] <size>"),
                    }
                }
                "scroll" => {
                    if let Some((v, upper)) = parse_pair_f64(&value) {
                        state.set_scroll(id, 0.0, upper, 0.0, v, upper, 0.0);
                        println!("Scrolled to {}.", v);
                    } else {
                        println!("Usage: set scroll <value> <upper>");
                    }
                }
                "title" => {
                    state.set_window_title(id, &value);
                    println!("Title set to \"{}\".", value);
                }
                "text" => {
                    match &state.node(id).map(|n| &n.kind) {
                        Some(ZorkKind::Label { .. }) => {
                            state.set_label_text(id, &value);
                            println!("Label text set.");
                        }
                        Some(ZorkKind::TextView { .. }) => {
                            state.set_textview_text(id, &value);
                            println!("TextView text set.");
                        }
                        Some(ZorkKind::Entry { .. }) => {
                            state.set_entry_text(id, &value);
                            println!("Entry text set.");
                        }
                        _ => println!("That has no text."),
                    }
                }
                "check" => {
                    match value.as_str() {
                        "on" | "true" | "yes" | "1" => {
                            state.set_checkbutton_checked(id, true);
                            state.set_radiobutton_checked(id, true);
                            println!("Checked.");
                        }
                        "off" | "false" | "no" | "0" => {
                            state.set_checkbutton_checked(id, false);
                            state.set_radiobutton_checked(id, false);
                            println!("Unchecked.");
                        }
                        _ => println!("Usage: set check on|off"),
                    }
                }
                _ => println!("Unknown property \"{}\". Try 'help'.", prop),
            }
        }
        "hide" => {
            state.set_visible(state.current_id, false);
            println!("Hidden.");
        }
        "show" => {
            state.set_visible(state.current_id, true);
            println!("Shown.");
        }
        "focus" => {
            let id = state.current_id;
            state.set_focus(id);
            if state.has_focus(id) {
                println!("You focus it.");
            } else {
                println!("You can't focus that.");
            }
        }
        "layout" => {
            // Lay the current box's children out, and report what it needs.
            let id = state.current_id;
            if let Some((w, h)) = state.measure_box(id) {
                state.layout_box(id, 0, 0, w, h);
                println!("Box laid out: {}x{}.", w, h);
            } else {
                println!("You can only lay out a Box.");
            }
        }
        "inventory" | "i" => {
            println!("Current path:");
            let mut path: Vec<String> = Vec::new();
            let mut cur = Some(state.current_id);
            while let Some(id) = cur {
                if let Some(node) = state.node(id) {
                    path.push(short_desc(state, node));
                    cur = node.parent;
                } else {
                    break;
                }
            }
            path.reverse();
            for (i, desc) in path.iter().enumerate() {
                println!("  {}{}", "  ".repeat(i), desc);
            }
        }
        "back" => {
            if let Some(pid) = state.node(state.current_id).and_then(|n| n.parent) {
                state.current_id = pid;
                println!("You go back south.");
            } else {
                println!("You can't go back from here.");
            }
        }
        "quit" | "q" | "exit" => {
            println!("Goodbye!");
            state.running = false;
        }
        "help" | "?" => {
            println!("Commands:");
            println!("  look / l              - describe your surroundings");
            println!("  go north/south/east/west - move in a direction");
            println!("  go <number>           - go to numbered item");
            println!("  north/n, south/s, east/e, west/w - quick move");
            println!("  click / press [n]     - press a button or open menu (default: current)");
            println!("  select / choose <n>   - select a menu item from a MenuBar/Menu");
            println!("  examine / x [n]       - examine something in detail");
            println!("  type / write          - enter text into an Entry (sub-prompt)");
            println!("  read                  - read text at current location");
            println!("  toggle                - toggle a CheckButton/RadioButton");
            println!("  props                 - show this widget's recorded properties");
            println!("  set <prop> <value>    - set size/offset/margin/class/xalign/");
            println!("                          fixed-width/font/scroll/title/text/check");
            println!("  hide / show           - toggle visibility");
            println!("  focus                 - take keyboard focus");
            println!("  layout                - lay out a Box's children");
            println!("  inventory / i         - show your path");
            println!("  back                  - go back the way you came");
            println!("  quit / q / exit       - exit the game");
            println!("  help / ?              - show this help");
        }
        _ => {
            if let Ok(num) = cmd.parse::<usize>() {
                let target = resolve_number_target(state, num);
                if let Some(id) = target {
                    let desc = state.node(id).map(|n| short_desc(state, n)).unwrap_or_default();
                    state.prev_location = Some(state.current_id);
                    state.current_id = id;
                    println!("You move to {}.", desc);
                } else {
                    println!("Invalid number.");
                }
            } else {
                println!("I don't understand \"{}\". Type 'help' for commands.", cmd);
            }
        }
    }
}

/// Parse `"<a> <b>"` as a pair of `i32`, for `set size` / `set offset`.
fn parse_pair(value: &str) -> Option<(i32, i32)> {
    let mut it = value.split_whitespace();
    let a = it.next()?.parse().ok()?;
    let b = it.next()?.parse().ok()?;
    Some((a, b))
}

/// Parse `"<a> <b>"` as a pair of `f64`, for `set scroll`.
fn parse_pair_f64(value: &str) -> Option<(f64, f64)> {
    let mut it = value.split_whitespace();
    let a = it.next()?.parse().ok()?;
    let b = it.next()?.parse().ok()?;
    Some((a, b))
}

fn resolve_arg_to_id(state: &ZorkState, arg: &str) -> Option<usize> {    if let Ok(num) = arg.parse::<usize>() {
        resolve_number_target(state, num)
    } else {
        let dir = arg.to_lowercase();
        navigate_find(state, &dir)
    }
}

fn resolve_number_target(state: &ZorkState, num: usize) -> Option<usize> {
    if num == 0 {
        return Some(state.current_id);
    }
    let node = state.node(state.current_id)?;
    let mut items: Vec<usize> = Vec::new();

    let siblings: Vec<usize> = if let Some(pid) = node.parent {
        state.node(pid).map(|p| p.children.clone()).unwrap_or_default()
    } else {
        Vec::new()
    };
    let children: Vec<usize> = node.children.clone();

    for id in &siblings {
        if *id == node.id {
            continue;
        }
        items.push(*id);
    }
    for id in &children {
        if !items.contains(id) {
            items.push(*id);
        }
    }

    items.get(num - 1).copied()
}

fn navigate_find(state: &ZorkState, dir: &str) -> Option<usize> {
    let node = state.node(state.current_id)?;
    let siblings: Vec<&ZorkNode> = if let Some(pid) = node.parent {
        state.node(pid).map(|p| p.children.iter().filter_map(|cid| state.node(*cid)).collect()).unwrap_or_default()
    } else {
        Vec::new()
    };
    let children: Vec<&ZorkNode> = node.children.iter().filter_map(|cid| state.node(*cid)).collect();

    let mut my_idx = None;
    for (i, sib) in siblings.iter().enumerate() {
        if sib.id == node.id {
            my_idx = Some(i);
            break;
        }
    }

    match dir {
        "north" | "n" => children.first().map(|n| n.id),
        "south" | "s" => node.parent,
        "east" | "e" => {
            if let Some(idx) = my_idx {
                siblings.get(idx + 1).map(|n| n.id)
            } else {
                None
            }
        }
        "west" | "w" => {
            if let Some(idx) = my_idx {
                if idx > 0 {
                    siblings.get(idx - 1).map(|n| n.id)
                } else {
                    node.parent
                }
            } else {
                node.parent
            }
        }
        _ => None,
    }
}

fn navigate_to(state: &mut ZorkState, dir: &str) {
    if let Some(target_id) = navigate_find(state, dir) {
        state.prev_location = Some(state.current_id);
        state.current_id = target_id;
        println!("You go {}.", dir);
    } else if let Ok(num) = dir.parse::<usize>() {
        if let Some(id) = resolve_number_target(state, num) {
            let desc = state.node(id).map(|n| short_desc(state, n)).unwrap_or_default();
            state.prev_location = Some(state.current_id);
            state.current_id = id;
            println!("You move to {}.", desc);
        } else {
            println!("You can't go that way.");
        }
    } else {
        println!("You can't go that way.");
    }
}

fn examine(state: &ZorkState, id: usize) {
    if let Some(node) = state.node(id) {
        match &node.kind {
            ZorkKind::Label { text } => {
                println!("The label reads:");
                println!("\"{}\"", text);
            }
            ZorkKind::Entry { buffer, .. } => {
                if buffer.is_empty() {
                    println!("The Entry is blank.");
                } else {
                    println!("The Entry contains:");
                    println!("\"{}\"", buffer);
                }
            }
            ZorkKind::TextView { text } => {
                println!("The TextView contains:");
                for line in text.lines() {
                    println!("  {}", line);
                }
            }
            ZorkKind::Button { label } => {
                println!("A button labeled \"{}\". It looks clickable.", label);
            }
            ZorkKind::CheckButton { label, checked } => {
                println!("CheckButton \"{}\": currently {}.", label, if *checked { "CHECKED" } else { "UNCHECKED" });
            }
            ZorkKind::RadioButton { label, checked, .. } => {
                println!("RadioButton \"{}\": currently {}.", label, if *checked { "SELECTED" } else { "NOT SELECTED" });
            }
            ZorkKind::DropDown { items, selected } => {
                println!("DropDown with {} items:", items.len());
                for (i, item) in items.iter().enumerate() {
                    let marker = if Some(i) == *selected { " <--" } else { "" };
                    println!("  {}. {}{}", i, item, marker);
                }
            }
            ZorkKind::Window { title } => {
                println!("A window titled \"{}\".", title);
            }
            ZorkKind::Dialog { title } => {
                println!("A dialog titled \"{}\".", title);
            }
            ZorkKind::BoxWidget { horizontal, spacing } => {
                println!("A {} box with spacing {}.", if *horizontal { "horizontal" } else { "vertical" }, spacing);
            }
            ZorkKind::Grid { cols, rows } => {
                println!("A grid with {} cols, {} rows.", cols, rows);
            }
            ZorkKind::MenuBar => {
                println!("A MenuBar.");
                let items = state.menu_items.get(&node.id).cloned().unwrap_or_default();
                if items.is_empty() {
                    println!("It has no items.");
                } else {
                    println!("It contains:");
                    for (i, item) in items.iter().enumerate() {
                        println!("  {}. {} ({})", i + 1, item.label, item.action);
                    }
                }
            }
            ZorkKind::Menu => {
                println!("A Menu.");
                let items = state.menu_items.get(&node.id).cloned().unwrap_or_default();
                if items.is_empty() {
                    println!("It has no items.");
                } else {
                    println!("It contains:");
                    for (i, item) in items.iter().enumerate() {
                        println!("  {}. {} ({})", i + 1, item.label, item.action);
                    }
                }
            }
            ZorkKind::Canvas { .. } => {
                let (w, h) = state.canvas_size(node.id);
                println!("A Canvas, {}x{}, with {} redraw request(s).", w, h, node.pointer.redraws);
                if node.draw.is_some() {
                    println!("It has a draw callback.");
                } else {
                    println!("It has no draw callback.");
                }
            }
            ZorkKind::Overlay => {
                println!("An Overlay with {} layer(s) over its base child.", node.overlays.len());
            }
            ZorkKind::ScrolledWindow => {
                println!("A ScrolledWindow, scrolled to ({}, {}).", node.props.hscroll, node.props.vscroll);
            }
            ZorkKind::Fixed => {
                println!("A Fixed container with {} child(ren).", node.children.len());
            }
            ZorkKind::Application => {
                println!("The Application.");
            }
            ZorkKind::Spreadsheet { .. } => {
                println!("A Spreadsheet with {} cell(s) set.", node.cells.len());
                let mut keys: Vec<&(u32, u32)> = node.cells.keys().collect();
                keys.sort();
                for k in keys {
                    if let Some(cell) = node.cells.get(k) {
                        println!("  R{}C{}: {}{}", k.0, k.1, cell.text, if cell.raw { " (raw)" } else { "" });
                    }
                }
            }
            _ => {
                println!("There's nothing special about this.");
            }
        }
        // Every node carries the same property bag, so show it for all of them.
        let summary = prop_summary(state, node);
        if summary != "no properties set" {
            println!();
            println!("Properties: {}", summary);
        }
    } else {
        println!("Nothing to examine.");
    }
}
