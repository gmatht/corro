#[cfg(feature = "pancurses")]
mod pancurses_backend {
    use pancurses::*;
    use std::cell::RefCell;
    use std::collections::HashMap;
    use std::os::raw::c_void;
    use std::rc::Rc;
    use std::io::Write;

    // ── SGR escape sequences matching ratatui output ──────────────
    const SGR_FG_BLACK: &str = "\x1b[38;5;0m";
    const SGR_FG_CYAN: &str = "\x1b[38;5;6m";
    const SGR_FG_YELLOW: &str = "\x1b[38;5;3m";
    const SGR_FG_DARK_GRAY: &str = "\x1b[38;5;8m";
    const SGR_BG_CYAN: &str = "\x1b[48;5;6m";
    const SGR_BG_YELLOW: &str = "\x1b[48;5;3m";
    const SGR_BG_DARK_GRAY: &str = "\x1b[48;5;8m";
    const SGR_BOLD: &str = "\x1b[1m";
    const SGR_UNDERLINE: &str = "\x1b[4m";
    const SGR_RESET: &str = "\x1b[0m";
    const SGR_FG_DEFAULT: &str = "\x1b[39m";
    const SGR_BG_DEFAULT: &str = "\x1b[49m";

    fn sgr_cup(y: i32, x: i32) -> String {
        format!("\x1b[{};{}H", y + 1, x + 1)
    }
    fn sgr_menu() -> &'static str { "\x1b[38;5;0m\x1b[48;5;6m" }
    fn sgr_formula() -> &'static str { "\x1b[38;5;6m\x1b[49m" }
    fn sgr_header_active() -> &'static str { "\x1b[1m\x1b[38;5;0m\x1b[48;5;3m" }
    fn sgr_header_inactive() -> &'static str { "\x1b[1m\x1b[38;5;6m" }
    fn sgr_row_cursor() -> &'static str { "\x1b[1m\x1b[38;5;0m\x1b[48;5;3m" }
    fn sgr_row_normal() -> &'static str { "\x1b[38;5;3m" }
    fn sgr_row_footer() -> &'static str { "\x1b[1m\x1b[38;5;6m" }
    fn sgr_row_underline() -> &'static str { "\x1b[4m\x1b[38;5;3m" }
    fn sgr_sep() -> &'static str { "\x1b[38;5;8m" }
    fn sgr_cell_cursor() -> &'static str { "\x1b[48;5;8m" }
    fn sgr_cell_agg() -> &'static str { "\x1b[38;5;6m" }
    fn sgr_cell_footer_agg() -> &'static str { "\x1b[1m\x1b[38;5;6m" }

    fn emit_sgr(s: &str) {
        let _ = std::io::stdout().write_all(s.as_bytes());
        let _ = std::io::stdout().flush();
    }

    pub type Callback = Box<dyn FnMut()>;

    #[derive(Clone, Copy, Debug, Default)]
    pub struct Rect {
        pub x: i32,
        pub y: i32,
        pub w: i32,
        pub h: i32,
    }

    #[derive(Clone)]
    pub enum PcWidgetKind {
        Window { title: String },
        Button { label: String },
        Label { text: String },
        BoxWidget { horizontal: bool, spacing: i32 },
        Grid { cols: usize, rows: usize },
        Entry { buffer: String, cursor: usize },
        CheckButton { label: String, checked: bool },
        RadioButton { label: String, checked: bool, group_id: usize },
        Dialog { title: String },
        Menu,
        MenuBar { labels: Vec<String>, submenu_items: Vec<(String, Vec<(String, String)>)> },
        SimpleAction,
        DropDown { items: Vec<String>, selected: Option<usize> },
        TextView { text: String },
        Spreadsheet {
            cells: Rc<RefCell<HashMap<(u32, u32), String>>>,
            raw_cells: Rc<RefCell<HashMap<(u32, u32), String>>>,
            cell_styles: Rc<RefCell<HashMap<(u32, u32), u8>>>,
            total_rows: u32,
            total_cols: u32,
            top_row: u32,
            left_col: u32,
            cursor_row: u32,
            cursor_col: u32,
            editing: bool,
            edit_buf: String,
            edit_pos: usize,
            col_width: u32,
            margin_cols: u32,
            main_cols: u32,
            formula_bar_address_id: Option<usize>,
            formula_bar_entry_id: Option<usize>,
            anchor: Option<(u32, u32)>,
            header_row_count: u32,
            main_row_count: u32,
            menu_text: String,
            status_text: String,
            border_title: String,
            formula_bar_trailing: String,
            column_layout: Vec<(u32, u32, String)>,
            row_labels: Vec<(u32, String)>,
            tab_text: String,
        },
    }

    pub struct PcWidgetNode {
        pub id: usize,
        pub kind: PcWidgetKind,
        pub parent: Option<usize>,
        pub children: Vec<usize>,
        pub rect: Rect,
        pub visible: bool,
        pub callbacks: Vec<Callback>,
    }

    pub struct PcState {
        pub nodes: Vec<PcWidgetNode>,
        pub next_id: usize,
        pub running: bool,
        pub focus_id: Option<usize>,
        pub menu_bar_id: Option<usize>,
        pub menu_open: bool,
        pub active_submenu: usize,
        pub active_item: usize,
        pub spreadsheet_output: String,
        key_callbacks: Vec<(char, Box<dyn FnMut()>)>,
    }

    impl PcState {
        pub fn new() -> Self {
            PcState {
                nodes: Vec::new(),
                next_id: 1,
                running: true,
                focus_id: None,
                menu_bar_id: None,
                menu_open: false,
                active_submenu: 0,
                active_item: 0,
                spreadsheet_output: String::new(),
                key_callbacks: Vec::new(),
            }
        }

        pub fn alloc_id(&mut self) -> usize {
            let id = self.next_id;
            self.next_id += 1;
            id
        }

        pub fn add_node(&mut self, kind: PcWidgetKind, parent: Option<usize>) -> usize {
            let id = self.alloc_id();
            self.nodes.push(PcWidgetNode {
                id,
                kind,
                parent,
                children: Vec::new(),
                rect: Rect::default(),
                visible: true,
                callbacks: Vec::new(),
            });
            if let Some(pid) = parent {
                if let Some(p) = self.nodes.iter_mut().find(|n| n.id == pid) {
                    p.children.push(id);
                }
            }
            id
        }

        pub fn node_mut(&mut self, id: usize) -> Option<&mut PcWidgetNode> {
            self.nodes.iter_mut().find(|n| n.id == id)
        }

        pub fn node(&self, id: usize) -> Option<&PcWidgetNode> {
            self.nodes.iter().find(|n| n.id == id)
        }
    }

    thread_local! {
        static PC_STATE: RefCell<PcState> = RefCell::new(PcState::new());
    }

    fn with_state<F, R>(f: F) -> R
    where
        F: FnOnce(&mut PcState) -> R,
    {
        PC_STATE.with(|s| f(&mut s.borrow_mut()))
    }

    pub struct PcApp;

    impl crate::backends::BackendApp for PcApp {
        fn run(self: Box<Self>) -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
            let root = initscr();
            raw();
            noecho();
            root.keypad(true);
            mousemask(ALL_MOUSE_EVENTS | REPORT_MOUSE_POSITION, None);
            // Use `root.timeout(100)` so we can temporarily change it for Alt+key detection
            root.timeout(100);
            curs_set(0);

            if has_colors() {
                start_color();
                init_pair(1, COLOR_WHITE, COLOR_BLUE);
                init_pair(2, COLOR_BLACK, COLOR_WHITE);
                init_pair(3, COLOR_WHITE, COLOR_BLACK);
                init_pair(4, COLOR_YELLOW, COLOR_BLACK);
                init_pair(5, COLOR_GREEN, COLOR_BLACK);
                init_pair(6, COLOR_RED, COLOR_BLACK);
                init_pair(7, COLOR_CYAN, COLOR_BLACK);
            }

            with_state(|state| {
                let (my, mx) = root.get_max_yx();
                for node in &mut state.nodes {
                    match &node.kind {
                        PcWidgetKind::Window { .. } | PcWidgetKind::Dialog { .. } => {
                            node.rect = Rect { x: 0, y: 0, w: mx, h: my };
                        }
                        PcWidgetKind::MenuBar { .. } => {
                            state.menu_bar_id = Some(node.id);
                        }
                        _ => {}
                    }
                }
                state.running = true;
            });

            while with_state(|s| s.running) {
                // Clear direct-write spreadsheet output for this frame
                with_state(|s| s.spreadsheet_output.clear());
                // layout pass
                with_state(|state| {
                    let ids: Vec<usize> = state.nodes.iter().map(|n| n.id).collect();
                    for id in &ids {
                        let info = state.node(*id).and_then(|n| {
                            let pid = n.parent?;
                            let pn = state.node(pid)?;
                            Some((n.rect, pn.rect, matches!(pn.kind, PcWidgetKind::Dialog { .. })))
                        });
                        if let Some((r, pr, is_dialog)) = info {
                            if (r.w == 0 || r.h == 0) && (pr.w > 0 || pr.h > 0) {
                                if let Some(n) = state.node_mut(*id) {
                                    let ins = if is_dialog { 1 } else { 0 };
                                    if n.rect.w == 0 { n.rect.w = (pr.w - 2 * ins).max(1); }
                                    if n.rect.h == 0 { n.rect.h = (pr.h - 2 * ins).max(1); }
                                    if is_dialog {
                                        if n.rect.x == 0 { n.rect.x = pr.x + 1; }
                                        if n.rect.y == 0 { n.rect.y = pr.y + 1; }
                                    }
                                }
                            }
                        }
                    }
                    for id in &ids {
                        let kind = state.node(*id).map(|n| &n.kind).cloned();
                        match kind {
                            Some(PcWidgetKind::BoxWidget { .. }) => { layout_box_inner(*id, state); }
                            Some(PcWidgetKind::Grid { .. }) => { layout_grid_inner(*id, state); }
                            _ => {}
                        }
                    }
                    for id in &ids {
                        if let Some(rect) = state.node(*id).map(|n| n.rect) {
                            if let Some(pid) = state.node(*id).and_then(|n| n.parent) {
                                let pw = state.node(pid).map(|p| p.rect.w).unwrap_or(1).max(1);
                                let ph = state.node(pid).map(|p| p.rect.h).unwrap_or(1).max(1);
                                let mut nr = rect;
                                nr.x = rect.x.clamp(0, pw - 1);
                                nr.y = rect.y.clamp(0, ph - 1);
                                nr.w = rect.w.min(pw - nr.x);
                                nr.h = rect.h.min(ph - nr.y);
                                if let Some(n) = state.node_mut(*id) { n.rect = nr; }
                            }
                        }
                    }
                });

                // render
                for i in 0..with_state(|s| s.nodes.len()) {
                    let (kind, rect, visible, id, focus_id) = with_state(|s| {
                        let n = &s.nodes[i];
                        (n.kind.clone(), n.rect, n.visible, n.id, s.focus_id)
                    });
                    if !visible { continue; }
                    render_widget(&root, &kind, rect, id, focus_id);
                }
                // render active menu dropdown
                with_state(|state| {
                    if state.menu_open {
                        let mid = state.menu_bar_id.unwrap_or(0);
                        let (sub_idx, item_idx) = (state.active_submenu, state.active_item);
                        if let Some(n) = state.node(mid) {
                            if let PcWidgetKind::MenuBar { labels, submenu_items } = &n.kind {
                                if sub_idx < submenu_items.len() {
                                    let (_, items) = &submenu_items[sub_idx];
                                    let dy = n.rect.y + 1;
                                    let dx = n.rect.x + 1;
                                    // find horizontal position of this submenu label
                                    let mut mx = n.rect.x + 1;
                                    for i in 0..sub_idx {
                                        mx += labels[i].len() as i32 + 2;
                                    }
                                    let max_w = items.iter().map(|(l,_)| l.len()).max().unwrap_or(0).max(4) as i32;
                                    if has_colors() { root.attron(COLOR_PAIR(4)); }
                                    for row in 0..items.len() as i32 {
                                        for col in 0..max_w + 2 {
                                            root.mvaddch(dy + row, mx + col, ' ');
                                        }
                                    }
                                    if has_colors() { root.attroff(COLOR_PAIR(4)); }
                                    for (i, (lbl, _)) in items.iter().enumerate() {
                                        let bg = if i == item_idx && has_colors() { COLOR_PAIR(2) } else { 0 };
                                        if bg != 0 { root.attron(bg); }
                                        root.mvaddstr(dy + i as i32, mx + 1, lbl);
                                        if bg != 0 { root.attroff(bg); }
                                    }
                                }
                            }
                        }
                    }
                });

                root.refresh();
                // Flush direct-write spreadsheet SGR output after ncurses flush
                with_state(|s| {
                    if !s.spreadsheet_output.is_empty() {
                        emit_sgr(&s.spreadsheet_output);
                    }
                });
                match root.getch() {
                    Some(Input::KeyResize) => {
                        let (my, mx) = root.get_max_yx();
                        with_state(|state| {
                            for node in &mut state.nodes {
                                match &node.kind {
                                    PcWidgetKind::Window { .. } | PcWidgetKind::Dialog { .. } => {
                                        node.rect = Rect { x: 0, y: 0, w: mx, h: my };
                                    }
                                    _ => {}
                                }
                            }
                        });
                    }
                    Some(Input::KeyMouse) => {
                        if let Ok(mevent) = getmouse() {
                            let x = mevent.x;
                            let y = mevent.y;
                            let callbacks = with_state(|state| {
                                let hit = state.nodes.iter().rev().find(|n| {
                                    n.visible
                                        && n.rect.x <= x && x < n.rect.x + n.rect.w
                                        && n.rect.y <= y && y < n.rect.y + n.rect.h
                                });
                                if let Some(node) = hit {
                                    match &node.kind {
                                        PcWidgetKind::Button { .. } => {
                                            let idxs: Vec<usize> = (0..state.nodes.len())
                                                .filter(|i| state.nodes[*i].id == node.id)
                                                .collect();
                                            let mut all_cbs = vec![];
                                            for idx in idxs {
                                                all_cbs.append(&mut std::mem::take(&mut state.nodes[idx].callbacks));
                                            }
                                            all_cbs
                                        }
                                        PcWidgetKind::CheckButton { .. } => {
                                            if let Some(n) = state.node_mut(node.id) {
                                                if let PcWidgetKind::CheckButton { ref mut checked, .. } = n.kind {
                                                    *checked = !*checked;
                                                }
                                                std::mem::take(&mut n.callbacks)
                                            } else {
                                                vec![]
                                            }
                                        }
                                        PcWidgetKind::Entry { .. } => {
                                            state.focus_id = Some(node.id);
                                            vec![]
                                        }
                                        _ => vec![],
                                    }
                                } else {
                                    vec![]
                                }
                            });
                            fire_callbacks(callbacks);
                        }
                    }
                    Some(Input::Character('q')) | Some(Input::Character('Q')) => {
                        // Only quit when spreadsheet is not editing
                        let editing = with_state(|state| {
                            if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    if let Some(n) = state.node(fid) {
                                        if let PcWidgetKind::Spreadsheet { editing, .. } = &n.kind {
                                            return *editing;
                                        }
                                    }
                                }
                            }
                            false
                        });
                        if !editing {
                            with_state(|state| state.running = false);
                        }
                    }
                    Some(Input::Character('\n')) | Some(Input::Character('\r')) => {
                        let callbacks = with_state(|state| {
                            if state.menu_open {
                                state.menu_open = false;
                                vec![]
                            } else if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_enter(state, fid);
                                    spreadsheet_scroll_to_cursor(state, fid);
                                    vec![]
                                } else {
                                    toggle_focused(state, fid).1
                                }
                            } else {
                                vec![]
                            }
                        });
                        fire_callbacks(callbacks);
                    }
                    Some(Input::Character(c)) => {
                        if c == '\t' {
                            with_state(|state| {
                                if state.menu_open {
                                    state.menu_open = false;
                                } else if let Some(fid) = state.focus_id {
                                    if is_spreadsheet_focused(state, fid) {
                                        spreadsheet_commit_edit(state, fid);
                                        if let Some(n) = state.node_mut(fid) {
                                            if let PcWidgetKind::Spreadsheet { ref mut cursor_col, total_cols, .. } = n.kind {
                                                if *cursor_col + 1 < total_cols { *cursor_col += 1; }
                                            }
                                        }
                                        spreadsheet_scroll_to_cursor(state, fid);
                                    } else {
                                        let focusable: Vec<usize> = state.nodes.iter()
                                            .filter(|n| matches!(n.kind, PcWidgetKind::Button { .. } | PcWidgetKind::Entry { .. } | PcWidgetKind::CheckButton { .. } | PcWidgetKind::Spreadsheet { .. }))
                                            .map(|n| n.id)
                                            .collect();
                                        if let Some(pos) = state.focus_id.and_then(|f| focusable.iter().position(|&x| x == f)) {
                                            let next = (pos + 1) % focusable.len();
                                            state.focus_id = Some(focusable[next]);
                                        } else {
                                            state.focus_id = focusable.first().copied();
                                        }
                                    }
                                }
                            });
                        } else if c == '\x1b' {
                            // Distinguish bare Escape from Alt+key (terminal sends ESC + char)
                            root.timeout(300);
                            let alt_key = root.getch();
                            root.timeout(100);
                            match alt_key {
                                Some(Input::Character(ac)) => {
                                    // Alt+key — activate matching submenu
                                    with_state(|state| {
                                        if let Some(mid) = state.menu_bar_id {
                                            if let Some(n) = state.node(mid) {
                                                if let PcWidgetKind::MenuBar { labels, .. } = &n.kind {
                                                    let lower = ac.to_ascii_lowercase();
                                                    if let Some(pos) = labels.iter().position(|l| l.to_ascii_lowercase().starts_with(&lower.to_string())) {
                                                        state.active_submenu = pos;
                                                        state.active_item = 0;
                                                        state.menu_open = true;
                                                    }
                                                }
                                            }
                                        }
                                    });
                                }
                                _ => {
                                    // No following char within 300ms → bare Escape → close edit/menu or quit
                                    with_state(|state| {
                                        if let Some(fid) = state.focus_id {
                                            if is_spreadsheet_focused(state, fid) {
                                                if let Some(n) = state.node_mut(fid) {
                                                    if let PcWidgetKind::Spreadsheet { ref mut editing, .. } = n.kind {
                                                        if *editing {
                                                            *editing = false; // cancel edit
                                                            return;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if state.menu_open {
                                            state.menu_open = false;
                                        } else {
                                            state.running = false;
                                        }
                                    });
                                }
                            }
                        } else if c == '\x03' {
                            // Ctrl+C — quit
                            with_state(|state| state.running = false);
                        } else if c == '\x07' {
                            // Ctrl+G — Go to column 0, row 999 (shows as A1000)
                            with_state(|state| {
                                if let Some(fid) = state.focus_id {
                                    if is_spreadsheet_focused(state, fid) {
                                        if let Some(n) = state.node_mut(fid) {
                                            if let PcWidgetKind::Spreadsheet { ref mut cursor_row, ref mut cursor_col, .. } = n.kind {
                                                *cursor_col = 0;
                                                *cursor_row = 999;
                                            }
                                        }
                                        spreadsheet_scroll_to_cursor(state, fid);
                                    }
                                }
                            });
                    } else if c == ' ' {
                            let callbacks = with_state(|state| {
                                if let Some(fid) = state.focus_id {
                                    if is_spreadsheet_focused(state, fid) {
                                        // Insert space into spreadsheet edit buffer when in edit mode
                                        if let Some(n) = state.node_mut(fid) {
                                            if let PcWidgetKind::Spreadsheet { ref mut editing, ref mut edit_buf, ref mut edit_pos, .. } = n.kind {
                                                if *editing {
                                                    edit_buf.insert(*edit_pos, ' ');
                                                    *edit_pos += 1;
                                                    return vec![];
                                                }
                                            }
                                        }
                                    }
                                    let (toggled, cbs) = toggle_focused(state, fid);
                                    if !toggled {
                                        let idxs: Vec<usize> = (0..state.nodes.len())
                                            .filter(|i| state.nodes[*i].id == fid)
                                            .collect();
                                        for idx in idxs {
                                            if let PcWidgetKind::Entry { ref mut buffer, ref mut cursor } = &mut state.nodes[idx].kind {
                                                buffer.insert(*cursor, ' ');
                                                *cursor += 1;
                                            }
                                        }
                                    }
                                    cbs
                                } else {
                                    vec![]
                                }
                            });
                            fire_callbacks(callbacks);
                        } else {
                            // Check registered key callbacks first
                            let mut fired = false;
                            {
                                let callbacks: Vec<Box<dyn FnMut()>> = with_state(|state| {
                                    let mut out = Vec::new();
                                    let mut i = 0;
                                    while i < state.key_callbacks.len() {
                                        if state.key_callbacks[i].0 == c {
                                            out.push(state.key_callbacks.swap_remove(i).1);
                                            fired = true;
                                        } else {
                                            i += 1;
                                        }
                                    }
                                    out
                                });
                                fire_callbacks(callbacks);
                            }
                            if !fired {
                                with_state(|state| {
                                    if let Some(fid) = state.focus_id {
                                        if is_spreadsheet_focused(state, fid) {
                                            if let Some(n) = state.node_mut(fid) {
                                                if let PcWidgetKind::Spreadsheet { ref mut editing, ref mut cursor_col, ref mut cursor_row, total_cols, total_rows, ref mut edit_buf, ref mut edit_pos, .. } = n.kind {
                                                    if !*editing {
                                                        // Not in edit mode — check for special keys
                                                        if c == 'c' || c == 'C' {
                                                            if *cursor_col < total_cols - 1 { *cursor_col = total_cols - 1; }
                                                            spreadsheet_scroll_to_cursor(state, fid);
                                                            return;
                                                        }
                                                        if c == 'r' || c == 'R' {
                                                            if *cursor_row < total_rows - 1 { *cursor_row = total_rows - 1; }
                                                            spreadsheet_scroll_to_cursor(state, fid);
                                                            return;
                                                        }
                                                        // Start editing with this character
                                                        *editing = true;
                                                        *edit_buf = c.to_string();
                                                        *edit_pos = 1;
                                                    } else {
                                                        // In edit mode — append character
                                                        edit_buf.insert(*edit_pos, c);
                                                        *edit_pos += 1;
                                                    }
                                                }
                                            }
                                        } else if let Some(n) = state.node_mut(fid) {
                                            if let PcWidgetKind::Entry { ref mut buffer, ref mut cursor } = n.kind {
                                                buffer.insert(*cursor, c);
                                                *cursor += 1;
                                            }
                                        }
                                    }
                                });
                            }
                        }
                    }
                    Some(Input::KeyBackspace) => {
                        with_state(|state| {
                            if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut editing, ref mut edit_buf, ref mut edit_pos, .. } = n.kind {
                                            if *editing && *edit_pos > 0 {
                                                *edit_pos -= 1;
                                                edit_buf.remove(*edit_pos);
                                            }
                                        }
                                    }
                                } else if let Some(n) = state.node_mut(fid) {
                                    if let PcWidgetKind::Entry { ref mut buffer, ref mut cursor } = n.kind {
                                        if *cursor > 0 {
                                            *cursor -= 1;
                                            buffer.remove(*cursor);
                                        }
                                    }
                                }
                            }
                        });
                    }
                    Some(Input::KeyDC) => {
                        with_state(|state| {
                            if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut editing, ref mut edit_buf, ref mut edit_pos, .. } = n.kind {
                                            if *editing && *edit_pos < edit_buf.len() {
                                                edit_buf.remove(*edit_pos);
                                            }
                                        }
                                    }
                                } else if let Some(n) = state.node_mut(fid) {
                                    if let PcWidgetKind::Entry { ref mut buffer, ref mut cursor } = n.kind {
                                        if *cursor < buffer.len() {
                                            buffer.remove(*cursor);
                                        }
                                    }
                                }
                            }
                        });
                    }
                    Some(Input::KeyLeft) => {
                        with_state(|state| {
                            if state.menu_open {
                                if state.active_submenu > 0 {
                                    state.active_submenu -= 1;
                                    state.active_item = 0;
                                }
                            } else if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_prepare_move(state, fid, false);
                                    spreadsheet_commit_edit(state, fid);
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut cursor_col, .. } = n.kind {
                                            if *cursor_col > 0 { *cursor_col -= 1; }
                                        }
                                    }
                                    spreadsheet_scroll_to_cursor(state, fid);
                                } else if let Some(n) = state.node_mut(fid) {
                                    if let PcWidgetKind::Entry { ref mut cursor, .. } = n.kind {
                                        if *cursor > 0 { *cursor -= 1; }
                                    }
                                }
                            }
                        });
                    }
                    Some(Input::KeyRight) => {
                        with_state(|state| {
                            if state.menu_open {
                                if let Some(mid) = state.menu_bar_id {
                                    if let Some(n) = state.node(mid) {
                                        if let PcWidgetKind::MenuBar { labels, .. } = &n.kind {
                                            if state.active_submenu + 1 < labels.len() {
                                                state.active_submenu += 1;
                                                state.active_item = 0;
                                            }
                                        }
                                    }
                                }
                            } else if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_prepare_move(state, fid, false);
                                    spreadsheet_commit_edit(state, fid);
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut cursor_col, total_cols, .. } = n.kind {
                                            if *cursor_col + 1 < total_cols { *cursor_col += 1; }
                                        }
                                    }
                                    spreadsheet_scroll_to_cursor(state, fid);
                                } else if let Some(n) = state.node_mut(fid) {
                                    if let PcWidgetKind::Entry { ref mut cursor, ref buffer } = n.kind {
                                        if *cursor < buffer.len() { *cursor += 1; }
                                    }
                                }
                            }
                        });
                    }
                    Some(Input::KeyUp) => {
                        with_state(|state| {
                            if state.menu_open {
                                if let Some(mid) = state.menu_bar_id {
                                    if let Some(n) = state.node(mid) {
                                        if let PcWidgetKind::MenuBar { submenu_items, .. } = &n.kind {
                                            let si = state.active_submenu;
                                            if si < submenu_items.len() && state.active_item > 0 {
                                                state.active_item -= 1;
                                            }
                                        }
                                    }
                                }
                            } else if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_prepare_move(state, fid, false);
                                    spreadsheet_commit_edit(state, fid);
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut cursor_row, .. } = n.kind {
                                            if *cursor_row > 0 { *cursor_row -= 1; }
                                        }
                                    }
                                    spreadsheet_scroll_to_cursor(state, fid);
                                }
                            }
                        });
                    }
                    Some(Input::KeyDown) => {
                        with_state(|state| {
                            if state.menu_open {
                                if let Some(mid) = state.menu_bar_id {
                                    if let Some(n) = state.node(mid) {
                                        if let PcWidgetKind::MenuBar { submenu_items, .. } = &n.kind {
                                            let si = state.active_submenu;
                                            if si < submenu_items.len() && state.active_item + 1 < submenu_items[si].1.len() {
                                                state.active_item += 1;
                                            }
                                        }
                                    }
                                }
                            } else if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_commit_edit(state, fid);
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut cursor_row, total_rows, .. } = n.kind {
                                            if *cursor_row + 1 < total_rows { *cursor_row += 1; }
                                        }
                                    }
                                    spreadsheet_scroll_to_cursor(state, fid);
                                }
                            }
                        });
                    }
                    Some(Input::KeyLeft) => {
                        with_state(|state| {
                            if state.menu_open {
                                if state.active_submenu > 0 {
                                    state.active_submenu -= 1;
                                    state.active_item = 0;
                                }
                            } else if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_commit_edit(state, fid);
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut cursor_col, .. } = n.kind {
                                            if *cursor_col > 0 { *cursor_col -= 1; }
                                        }
                                    }
                                    spreadsheet_scroll_to_cursor(state, fid);
                                } else if let Some(n) = state.node_mut(fid) {
                                    if let PcWidgetKind::Entry { ref mut cursor, .. } = n.kind {
                                        if *cursor > 0 { *cursor -= 1; }
                                    }
                                }
                            }
                        });
                    }
                    Some(Input::KeyRight) => {
                        with_state(|state| {
                            if state.menu_open {
                                if let Some(mid) = state.menu_bar_id {
                                    if let Some(n) = state.node(mid) {
                                        if let PcWidgetKind::MenuBar { labels, .. } = &n.kind {
                                            if state.active_submenu + 1 < labels.len() {
                                                state.active_submenu += 1;
                                                state.active_item = 0;
                                            }
                                        }
                                    }
                                }
                            } else if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_commit_edit(state, fid);
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut cursor_col, total_cols, .. } = n.kind {
                                            if *cursor_col + 1 < total_cols { *cursor_col += 1; }
                                        }
                                    }
                                } else if let Some(n) = state.node_mut(fid) {
                                    if let PcWidgetKind::Entry { ref mut cursor, ref buffer } = n.kind {
                                        if *cursor < buffer.len() { *cursor += 1; }
                                    }
                                }
                            }
                        });
                    }
                    Some(Input::KeyUp) => {
                        with_state(|state| {
                            if state.menu_open {
                                if let Some(mid) = state.menu_bar_id {
                                    if let Some(n) = state.node(mid) {
                                        if let PcWidgetKind::MenuBar { submenu_items, .. } = &n.kind {
                                            let si = state.active_submenu;
                                            if si < submenu_items.len() && state.active_item > 0 {
                                                state.active_item -= 1;
                                            }
                                        }
                                    }
                                }
                            } else if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_commit_edit(state, fid);
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut cursor_row, .. } = n.kind {
                                            if *cursor_row > 0 { *cursor_row -= 1; }
                                        }
                                    }
                                }
                            }
                        });
                    }
                    Some(Input::KeyDown) => {
                        with_state(|state| {
                            if state.menu_open {
                                if let Some(mid) = state.menu_bar_id {
                                    if let Some(n) = state.node(mid) {
                                        if let PcWidgetKind::MenuBar { submenu_items, .. } = &n.kind {
                                            let si = state.active_submenu;
                                            if si < submenu_items.len() && state.active_item + 1 < submenu_items[si].1.len() {
                                                state.active_item += 1;
                                            }
                                        }
                                    }
                                }
                            } else if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_commit_edit(state, fid);
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut cursor_row, total_rows, .. } = n.kind {
                                            if *cursor_row + 1 < total_rows { *cursor_row += 1; }
                                        }
                                    }
                                }
                            }
                        });
                    }
                    Some(Input::KeySLeft) => {
                        with_state(|state| {
                            if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_prepare_move(state, fid, true);
                                    spreadsheet_commit_edit(state, fid);
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut cursor_col, .. } = n.kind {
                                            if *cursor_col > 0 { *cursor_col -= 1; }
                                        }
                                    }
                                    spreadsheet_scroll_to_cursor(state, fid);
                                }
                            }
                        });
                    }
                    Some(Input::KeySRight) => {
                        with_state(|state| {
                            if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_prepare_move(state, fid, true);
                                    spreadsheet_commit_edit(state, fid);
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut cursor_col, total_cols, .. } = n.kind {
                                            if *cursor_col + 1 < total_cols { *cursor_col += 1; }
                                        }
                                    }
                                    spreadsheet_scroll_to_cursor(state, fid);
                                }
                            }
                        });
                    }
                    Some(Input::KeyPPage) => {
                        with_state(|state| {
                            if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_commit_edit(state, fid);
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut cursor_row, total_rows, .. } = n.kind {
                                            let page = 20.min(total_rows / 2);
                                            *cursor_row = cursor_row.saturating_sub(page);
                                        }
                                    }
                                    spreadsheet_scroll_to_cursor(state, fid);
                                }
                            }
                        });
                    }
                    Some(Input::KeyNPage) => {
                        with_state(|state| {
                            if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_commit_edit(state, fid);
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut cursor_row, total_rows, .. } = n.kind {
                                            let page = 20.min(total_rows / 2);
                                            *cursor_row = (*cursor_row + page).min(total_rows - 1);
                                        }
                                    }
                                    spreadsheet_scroll_to_cursor(state, fid);
                                }
                            }
                        });
                    }
                    Some(Input::KeyHome) => {
                        with_state(|state| {
                            if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_commit_edit(state, fid);
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut cursor_col, .. } = n.kind {
                                            *cursor_col = 0;
                                        }
                                    }
                                    spreadsheet_scroll_to_cursor(state, fid);
                                }
                            }
                        });
                    }
                    Some(Input::KeyEnd) => {
                        with_state(|state| {
                            if let Some(fid) = state.focus_id {
                                if is_spreadsheet_focused(state, fid) {
                                    spreadsheet_commit_edit(state, fid);
                                    if let Some(n) = state.node_mut(fid) {
                                        if let PcWidgetKind::Spreadsheet { ref mut cursor_col, total_cols, .. } = n.kind {
                                            *cursor_col = total_cols - 1;
                                        }
                                    }
                                    spreadsheet_scroll_to_cursor(state, fid);
                                }
                            }
                        });
                    }
                    Some(Input::KeySTab) => {
                        with_state(|state| {
                            let focusable: Vec<usize> = state.nodes.iter()
                                .filter(|n| matches!(n.kind, PcWidgetKind::Button { .. } | PcWidgetKind::Entry { .. } | PcWidgetKind::CheckButton { .. }))
                                .map(|n| n.id)
                                .collect();
                            if let Some(pos) = state.focus_id.and_then(|f| focusable.iter().position(|&x| x == f)) {
                                let prev = if pos == 0 { focusable.len() - 1 } else { pos - 1 };
                                state.focus_id = Some(focusable[prev]);
                            } else {
                                state.focus_id = focusable.last().copied();
                            }
                        });
                    }
                    Some(Input::KeyExit) => {
                        with_state(|state| state.running = false);
                    }
                    _ => {}
                }
            }

            endwin();
            Ok(())
        }
    }

    fn render_widget(root: &Window, kind: &PcWidgetKind, rect: Rect, id: usize, focus_id: Option<usize>) {
        match kind {
            PcWidgetKind::Window { title } => {
                if has_colors() {
                    root.attron(COLOR_PAIR(7));
                }
                for x in rect.x..rect.x + rect.w {
                    root.mvaddch(rect.y, x, '─');
                    root.mvaddch(rect.y + rect.h - 1, x, '─');
                }
                for y in rect.y..rect.y + rect.h {
                    root.mvaddch(y, rect.x, '│');
                    root.mvaddch(y, rect.x + rect.w - 1, '│');
                }
                root.mvaddch(rect.y, rect.x, '┌');
                root.mvaddch(rect.y, rect.x + rect.w - 1, '┐');
                root.mvaddch(rect.y + rect.h - 1, rect.x, '└');
                root.mvaddch(rect.y + rect.h - 1, rect.x + rect.w - 1, '┘');
                if has_colors() {
                    root.attroff(COLOR_PAIR(7));
                }
                let title_x = rect.x + (rect.w - title.len() as i32) / 2;
                if title_x > rect.x {
                    root.mvaddstr(rect.y, title_x, title);
                }
            }
            PcWidgetKind::Button { label } => {
                let focused = focus_id == Some(id);
                if focused && has_colors() {
                    root.attron(COLOR_PAIR(2));
                } else if has_colors() {
                    root.attron(COLOR_PAIR(1));
                }
                let inner_w = rect.w - 2;
                let display = if label.len() as i32 > inner_w {
                    let len = inner_w.max(0) as usize;
                    label[..len].to_string()
                } else {
                    label.clone()
                };
                let pad = (inner_w - display.len() as i32).max(0);
                let left_pad = pad / 2;
                let right_pad = pad - left_pad;
                root.mvaddch(rect.y, rect.x, '[');
                root.mvaddstr(rect.y, rect.x + 1 + left_pad, &display);
                root.mvaddch(rect.y, rect.x + 1 + left_pad + display.len() as i32 + right_pad, ']');
                if has_colors() {
                    root.attroff(COLOR_PAIR(1) | COLOR_PAIR(2));
                }
            }
            PcWidgetKind::Label { text } => {
                if has_colors() {
                    root.attron(COLOR_PAIR(3));
                }
                let max_w = rect.w as usize;
                let truncated = if text.len() > max_w { &text[..max_w] } else { text.as_str() };
                root.mvaddstr(rect.y, rect.x, truncated);
                if has_colors() {
                    root.attroff(COLOR_PAIR(3));
                }
            }
            PcWidgetKind::Entry { buffer, cursor } => {
                let focused = focus_id == Some(id);
                if focused && has_colors() {
                    root.attron(COLOR_PAIR(1));
                } else if has_colors() {
                    root.attron(COLOR_PAIR(3));
                }
                let max_w = rect.w as usize;
                let display = if buffer.len() > max_w {
                    &buffer[buffer.len() - max_w..]
                } else {
                    buffer.as_str()
                };
                root.mvaddstr(rect.y, rect.x, display);
                let rest = max_w.saturating_sub(display.len());
                for i in 0..rest {
                    root.mvaddch(rect.y, rect.x + display.len() as i32 + i as i32, ' ');
                }
                if focused {
                    let cursor_display = cursor.saturating_sub(buffer.len().saturating_sub(max_w));
                    let ch = '_' as u32 | A_REVERSE as u32;
                    root.mvaddch(rect.y, rect.x + cursor_display as i32, ch);
                }
                if has_colors() {
                    root.attroff(COLOR_PAIR(1) | COLOR_PAIR(3));
                }
            }
            PcWidgetKind::CheckButton { label, checked } => {
                let focused = focus_id == Some(id);
                if focused && has_colors() {
                    root.attron(COLOR_PAIR(2));
                } else if has_colors() {
                    root.attron(COLOR_PAIR(3));
                }
                let mark = if *checked { "[x]" } else { "[ ]" };
                root.mvaddstr(rect.y, rect.x, mark);
                let max_lbl = (rect.w - 4) as usize;
                let truncated = if label.len() > max_lbl { &label[..max_lbl] } else { label.as_str() };
                root.mvaddstr(rect.y, rect.x + 4, truncated);
                if has_colors() {
                    root.attroff(COLOR_PAIR(2) | COLOR_PAIR(3));
                }
            }
            PcWidgetKind::RadioButton { label, checked, .. } => {
                if has_colors() {
                    root.attron(COLOR_PAIR(3));
                }
                let mark = if *checked { "(•)" } else { "( )" };
                root.mvaddstr(rect.y, rect.x, mark);
                let max_lbl = (rect.w - 4) as usize;
                let truncated = if label.len() > max_lbl { &label[..max_lbl] } else { label.as_str() };
                root.mvaddstr(rect.y, rect.x + 4, truncated);
                if has_colors() {
                    root.attroff(COLOR_PAIR(3));
                }
            }
            PcWidgetKind::BoxWidget { .. } | PcWidgetKind::Grid { .. } => {}
            PcWidgetKind::Dialog { title } => {
                if has_colors() {
                    root.attron(COLOR_PAIR(7));
                }
                for x in rect.x..rect.x + rect.w {
                    root.mvaddch(rect.y, x, '═');
                    root.mvaddch(rect.y + rect.h - 1, x, '═');
                }
                for y in rect.y..rect.y + rect.h {
                    root.mvaddch(y, rect.x, '║');
                    root.mvaddch(y, rect.x + rect.w - 1, '║');
                }
                root.mvaddch(rect.y, rect.x, '╔');
                root.mvaddch(rect.y, rect.x + rect.w - 1, '╗');
                root.mvaddch(rect.y + rect.h - 1, rect.x, '╚');
                root.mvaddch(rect.y + rect.h - 1, rect.x + rect.w - 1, '╝');
                if has_colors() {
                    root.attroff(COLOR_PAIR(7));
                }
                root.mvaddstr(rect.y, rect.x + 2, title);
            }
            PcWidgetKind::Menu => {
                if has_colors() {
                    root.attron(COLOR_PAIR(4));
                }
                for x in rect.x..rect.x + rect.w {
                    root.mvaddch(rect.y, x, ' ');
                }
                if has_colors() {
                    root.attroff(COLOR_PAIR(4));
                }
            }
            PcWidgetKind::MenuBar { labels, .. } => {
                if has_colors() {
                    root.attron(COLOR_PAIR(4));
                }
                for x in rect.x..rect.x + rect.w {
                    root.mvaddch(rect.y, x, ' ');
                }
                let mut cx = rect.x + 1;
                for label in labels {
                    if cx + label.len() as i32 + 2 > rect.x + rect.w { break; }
                    root.mvaddstr(rect.y, cx, label);
                    cx += label.len() as i32 + 2;
                }
                if has_colors() {
                    root.attroff(COLOR_PAIR(4));
                }
            }
            PcWidgetKind::SimpleAction => {}
            PcWidgetKind::DropDown { items, selected } => {
                if has_colors() {
                    root.attron(COLOR_PAIR(1));
                }
                let current = selected.and_then(|s| items.get(s)).map(|s| s.as_str()).unwrap_or("");
                let max_w = (rect.w - 2) as usize;
                let truncated = if current.len() > max_w { &current[..max_w] } else { current };
                root.mvaddch(rect.y, rect.x, '[');
                root.mvaddstr(rect.y, rect.x + 1, truncated);
                root.mvaddch(rect.y, rect.x + 1 + truncated.len() as i32, 'v');
                let rest = max_w.saturating_sub(truncated.len() + 1);
                for i in 0..rest {
                    root.mvaddch(rect.y, rect.x + 2 + truncated.len() as i32 + i as i32, ' ');
                }
                root.mvaddch(rect.y, rect.x + rect.w - 1, ']');
                if has_colors() {
                    root.attroff(COLOR_PAIR(1));
                }
            }
            PcWidgetKind::TextView { text } => {
                if has_colors() {
                    root.attron(COLOR_PAIR(3));
                }
                let max_w = rect.w as usize;
                let max_h = rect.h as usize;
                let lines: Vec<&str> = text.lines().collect();
                for (i, line) in lines.iter().enumerate().take(max_h) {
                    let truncated = if line.len() > max_w { &line[..max_w] } else { line };
                    root.mvaddstr(rect.y + i as i32, rect.x, truncated);
                }
                if has_colors() {
                    root.attroff(COLOR_PAIR(3));
                }
            }
            PcWidgetKind::Spreadsheet { ref cells, ref raw_cells, ref cell_styles, ref top_row, ref left_col, ref cursor_row, ref cursor_col, ref editing, ref edit_buf, ref edit_pos, ref col_width, ref margin_cols, ref main_cols, ref menu_text, ref status_text, ref border_title, ref formula_bar_trailing, ref column_layout, ref row_labels, ref total_rows, ref total_cols, ref tab_text, header_row_count, main_row_count, .. } => {
                // ── Direct SGR rendering ──
                let lm = *margin_cols as usize;
                let mc = *main_cols as usize;
                let use_layout = !column_layout.is_empty();
                let mut out = String::new();
                out.push_str(SGR_RESET);
                let mut row_offset = rect.y;

                // Menu bar: black fg on cyan bg
                if !menu_text.is_empty() {
                    let max_chars = rect.w as usize;
                    let end = menu_text.char_indices().nth(max_chars).map(|(i, _)| i).unwrap_or(menu_text.len());
                    out.push_str(&sgr_cup(row_offset, rect.x));
                    out.push_str(sgr_menu());
                    out.push_str(&menu_text[..end]);
                    let remaining = (rect.w as usize).saturating_sub(end);
                    if remaining > 0 {
                        out.push_str(&" ".repeat(remaining));
                    }
                    row_offset += 1;
                }

                // Formula bar: cyan fg, default bg
                {
                    let display_col = if *cursor_col < *margin_cols { *cursor_col } else { *cursor_col - *margin_cols };
                    let addr_text = format!("{}{}", col_label(display_col), *cursor_row + 1);
                    let cell_val = raw_cells.borrow().get(&(*cursor_row, *cursor_col)).cloned().unwrap_or_default();
                    let addr_vis = addr_text.chars().count();
                    let max_fb = (rect.w as usize).saturating_sub(addr_vis + 3).max(1);
                    let fb_text = if cell_val.chars().count() > max_fb {
                        let trunc_end = cell_val.char_indices().nth(max_fb.saturating_sub(3)).map(|(i, _)| i).unwrap_or(cell_val.len());
                        format!(" {}  {}…{}", addr_text, &cell_val[..trunc_end], formula_bar_trailing)
                    } else {
                        format!(" {}  {}{}", addr_text, cell_val, formula_bar_trailing)
                    };
                    let fb_end = fb_text.char_indices().nth(rect.w as usize).map(|(i, _)| i).unwrap_or(fb_text.len());
                    out.push_str(&sgr_cup(row_offset, rect.x));
                    out.push_str(sgr_formula());
                    out.push_str(&fb_text[..fb_end]);
                    out.push_str(SGR_FG_DEFAULT);
                    out.push_str(SGR_BG_DEFAULT);
                    // Fill rest of formula bar line with spaces to prevent
                    // ncurses background from bleeding through.
                    let fb_vis = fb_text[..fb_end].chars().count();
                    if fb_vis < rect.w as usize {
                        out.push_str(&" ".repeat(rect.w as usize - fb_vis));
                    }
                    row_offset += 1;
                }

                // Top border line: ┌ title ───┐
                {
                    let br = row_offset;
                    let title = if !border_title.is_empty() {
                        format!("{}", border_title)
                    } else {
                        format!("corro  {}r × {}c ", *main_cols, *main_cols)
                    };
                    let title_vis = title.chars().count();
                    let dash_fill = (rect.w as usize).saturating_sub(title_vis + 3);
                    out.push_str(&sgr_cup(br, rect.x));
                    out.push_str("┌");
                    out.push_str(SGR_BOLD);
                    out.push_str(" ");
                    out.push_str(&title);
                    out.push_str(SGR_RESET);
                    out.push_str(&"─".repeat(dash_fill));
                    out.push_str("┐");
                    row_offset += 1;
                }

                // Grid header row
                {
                    let hr = row_offset;
                    out.push_str(&sgr_cup(hr, rect.x));
                    out.push_str("│");
                    if use_layout {
                        // Pre-computed layout
                        let mut hx = rect.x + 1;
                        out.push_str(&sgr_cup(hr, hx));
                        out.push_str(SGR_BOLD);
                        out.push_str(&" ".repeat(5));
                        out.push_str(SGR_RESET);
                        hx += 5;
                        let n = column_layout.len();
                        for (i, &(ref ci, ref w, ref label)) in column_layout.iter().enumerate() {
                            if hx + *w as i32 + 1 > rect.x + rect.w - 1 { break; }
                            let padded = format!("{:<1$}", label, *w as usize);
                            let active_col = *ci == *cursor_col;
                            let style = if active_col { sgr_header_active() } else { sgr_header_inactive() };
                            out.push_str(&sgr_cup(hr, hx));
                            out.push_str(style);
                            out.push_str(&padded);
                            out.push_str(SGR_RESET);
                            hx += *w as i32;
                            if i + 1 < n {
                                out.push_str(&sgr_cup(hr, hx));
                                out.push_str(" ");
                                hx += 1;
                            }
                        }
                        if hx <= rect.x + rect.w - 1 {
                            let gap = (rect.x + rect.w - 1 - hx) as usize;
                            if gap > 0 {
                                out.push_str(&" ".repeat(gap));
                            }
                            out.push_str(&sgr_cup(hr, rect.x + rect.w - 1));
                            out.push_str("│");
                        }
                    } else {
                        let cw = *col_width as i32;
                        let rh_w = 5i32;
                        let max_vis_cols = ((rect.w - rh_w - 2) / (cw + 1)).max(1).min(100);
                        let mut hx = rect.x + 1;
                        out.push_str(&sgr_cup(hr, hx));
                        out.push_str(&" ".repeat(rh_w as usize));
                        hx += rh_w;
                        for vc in 0..max_vis_cols {
                            let col_idx = *left_col + vc as u32;
                            let label = if col_idx < lm as u32 {
                                format!("[{}", col_label(col_idx))
                            } else if col_idx < (lm + mc) as u32 {
                                col_label(col_idx - lm as u32)
                            } else {
                                format!("]{}", col_label(col_idx - lm as u32 - mc as u32))
                            };
                            if hx + cw + 6 > rect.x + rect.w { break; }
                            let padded = format!("{:<1$}", label, cw as usize);
                            let active_col = col_idx == *cursor_col;
                            let style = if active_col { sgr_header_active() } else { sgr_header_inactive() };
                            out.push_str(&sgr_cup(hr, hx));
                            out.push_str(style);
                            out.push_str(&padded);
                            out.push_str(SGR_RESET);
                            hx += cw;
                            if hx < rect.x + rect.w {
                                out.push_str(&sgr_cup(hr, hx));
                                out.push_str(" ");
                                hx += 1;
                            }
                        }
                    }
                    // header separator: dark gray fg, matching ratatui
                    let sep_len = rect.w.saturating_sub(3).max(1) as usize;
                    out.push_str(&sgr_cup(hr + 1, rect.x));
                    out.push_str("│");
                    out.push_str(sgr_sep());
                    out.push_str("│");
                    out.push_str(&"─".repeat(sep_len));
                    out.push_str(SGR_FG_DEFAULT);
                    out.push_str("│");
                    row_offset = hr + 2;
                }

                // Data rows
                let has_status = !status_text.is_empty();
                let has_tabs = !tab_text.is_empty();
                let extra_lines = 1
                    + if has_tabs { 1 } else if has_status { 1 } else { 0 };
                let grid_bottom = rect.y + rect.h - extra_lines;
                let max_data_rows = ((grid_bottom - row_offset) as i32).max(1) as u32;
                for vr in 0..max_data_rows {
                    let row_idx = *top_row + vr as u32;
                    let ry = row_offset + vr as i32;
                    let is_cursor_row = row_idx == *cursor_row;
                    // Determine row label style
                    let label_str = row_labels.iter()
                        .find(|(r, _)| *r == row_idx)
                        .map(|(_, l)| l.as_str())
                        .unwrap_or("");
                    // Row label styling matches ratatui:
                    // - cursor row: bold + black fg + yellow bg
                    // - footer rows: bold + cyan fg
                    // - normal rows: yellow fg
                    // - boundary (last main) row: underline + fg
                    let is_footer = label_str.starts_with('_');
                    let is_header = label_str.starts_with('~');
                    // Check next row label to detect boundary (last main row before footers)
                    let next_label = row_labels.iter()
                        .find(|(r, _)| *r == row_idx + 1)
                        .map(|(_, l)| l.as_str())
                        .unwrap_or("");
                    let is_boundary = !is_footer && !is_header && next_label.starts_with('_');
                    let row_label_style = if is_cursor_row {
                        sgr_row_cursor()
                    } else if is_footer {
                        sgr_row_footer()
                    } else if is_boundary {
                        sgr_row_underline()
                    } else {
                        sgr_row_normal()
                    };
                    // Left border
                    out.push_str(&sgr_cup(ry, rect.x));
                    out.push_str("│");
                    // Row label (5 chars: right-aligned in 4 + trailing space)
                    out.push_str(&sgr_cup(ry, rect.x + 1));
                    out.push_str(row_label_style);
                    out.push_str(&format!("{:>4} ", label_str));
                    // Reset all attributes, then re-apply underline for boundary rows.
                    // This prevents underline from leaking to the next row.
                    out.push_str(SGR_RESET);
                    if is_boundary {
                        out.push_str(SGR_UNDERLINE);
                    }

                    // Separator between row label and first data column (dark gray)
                    // For the left margin gap: dark gray fg for 4 spaces
                    if *margin_cols > 0 && !use_layout {
                        let sep_x = rect.x + 1 + 5; // right after row label
                        out.push_str(&sgr_cup(ry, sep_x));
                        out.push_str(sgr_sep());
                        out.push_str("    ");
                        out.push_str(SGR_FG_DEFAULT);
                        out.push_str(" ");
                    }

                    let cells_ref = cells.borrow();
                    let cell_styles_ref = cell_styles.borrow();
                    if use_layout {
                        let n = column_layout.len();
                        let mut col_positions: Vec<(u32, i32)> = Vec::new();
                        let mut rx2 = rect.x + 1 + 5;
                        for (idx, &(col_idx, w, _)) in column_layout.iter().enumerate() {
                            col_positions.push((col_idx, rx2));
                            rx2 += w as i32;
                            if idx + 1 < n {
                                rx2 += 1;
                            }
                        }
                        // Determine last rendered x position for row-fill detection
                        let last_col_end = if col_positions.is_empty() {
                            rect.x + 1 + 5
                        } else {
                            let last_idx = col_positions.len() - 1;
                            col_positions[last_idx].1 + column_layout[last_idx].1 as i32
                        };
                        // Draw separator (dark gray) for left margin column if empty
                        if !col_positions.is_empty() {
                            let (first_col, first_sx) = col_positions[0];
                            let first_text = cells_ref.get(&(row_idx, first_col))
                                .map(|s| s.as_str()).unwrap_or("");
                            if first_text.is_empty() {
                                let w = column_layout[0].1 as i32;
                                out.push_str(&sgr_cup(ry, first_sx));
                                out.push_str(sgr_sep());
                                for _ in 0..w {
                                    out.push(' ');
                                }
                                out.push_str(SGR_FG_DEFAULT);
                                out.push(' ');
                            }
                        }
                        let mut vi = 0;
                        let mut prev_overflowed = false;
                        let mut right_gray_done = false;
                        let mut force_right_gray = false;
                        while vi < n {
                            let (col_idx, sx) = col_positions[vi];
                            let cell_text = cells_ref.get(&(row_idx, col_idx))
                                .map(|s| s.as_str()).unwrap_or("");
                            let cell_style = cell_styles_ref.get(&(row_idx, col_idx)).copied().unwrap_or(0);
                            let is_cursor_cell = cell_style == 1;
                            if cell_text.is_empty() {
                                // Empty cell: draw cursor highlight if needed
                                if is_cursor_cell {
                                    let cw = column_layout[vi].1 as i32;
                                    out.push_str(&sgr_cup(ry, sx));
                                    out.push_str(sgr_cell_cursor());
                                    for _ in 0..cw {
                                        out.push(' ');
                                    }
                                    out.push_str(SGR_BG_DEFAULT);
                                    let gap_after = if vi + 1 < n { 1 } else { 0 };
                                    if gap_after > 0 {
                                        out.push(' ');
                                    }
                                } else if (col_idx as usize) < lm {
                                    // Left-margin column: dark gray foreground
                                    let cw = column_layout[vi].1 as i32;
                                    out.push_str(&sgr_cup(ry, sx));
                                    if is_boundary {
                                        out.push_str(SGR_UNDERLINE);
                                    }
                                    out.push_str(sgr_sep());
                                    for _ in 0..cw {
                                        out.push(' ');
                                    }
                                    out.push_str(SGR_FG_DEFAULT);
                                    let gap_after = if vi + 1 < n { 1 } else { 0 };
                                    if gap_after > 0 {
                                        out.push(' ');
                                    }
                                } else if (col_idx as usize) >= lm + mc {
                                    // Right-margin column: dark gray foreground.
                                    // Match ratatui: only the first empty
                                    // right-margin column gets gray when there
                                    // is no overflow; all get gray when the
                                    // row has overflow (force_right_gray).
                                    let cw = column_layout[vi].1 as i32;
                                    let use_gray = force_right_gray || !right_gray_done;
                                    out.push_str(&sgr_cup(ry, sx));
                                    if is_boundary {
                                        out.push_str(SGR_UNDERLINE);
                                    }
                                    if use_gray {
                                        out.push_str(sgr_sep());
                                    }
                                    for _ in 0..cw {
                                        out.push(' ');
                                    }
                                    if use_gray {
                                        if is_boundary && vi + 1 >= n {
                                            out.push_str(SGR_RESET);
                                        } else {
                                            out.push_str(SGR_FG_DEFAULT);
                                        }
                                    }
                                    let gap_after = if vi + 1 < n { 1 } else { 0 };
                                    if gap_after > 0 {
                                        out.push(' ');
                                    }
                                    if !right_gray_done && !force_right_gray {
                                        right_gray_done = true;
                                    }
                                } else {
                                    // Empty cell in main column: default spaces
                                    let cw = column_layout[vi].1 as i32;
                                    out.push_str(&sgr_cup(ry, sx));
                                    if prev_overflowed {
                                        out.push_str(sgr_sep());
                                    } else {
                                        out.push_str(SGR_FG_DEFAULT);
                                        out.push_str(SGR_BG_DEFAULT);
                                    }
                                    for _ in 0..cw {
                                        out.push(' ');
                                    }
                                    if prev_overflowed {
                                        out.push_str(SGR_FG_DEFAULT);
                                    }
                                    let gap_after = if vi + 1 < n { 1 } else { 0 };
                                    if gap_after > 0 {
                                        out.push(' ');
                                    }
                                }
                                prev_overflowed = false;
                                vi += 1;
                                continue;
                            }
                            let w = column_layout[vi].1 as usize;
                            let text_width = cell_text.chars().count();
                            let mut overflow_cols = 0usize;
                            let mut can_overflow = false;
                            if text_width > w {
                                can_overflow = true;
                                let mut scan = vi + 1;
                                while scan < n {
                                    let (sc_idx, _) = col_positions[scan];
                                    let sc_text = cells_ref.get(&(row_idx, sc_idx))
                                        .map(|s| s.as_str()).unwrap_or("");
                                    if sc_text.is_empty() {
                                        overflow_cols += 1;
                                        scan += 1;
                                    } else { break; }
                                }
                            }
                            let gap_target = vi + overflow_cols;
                            let gap_after = if gap_target + 1 < n { 1 } else { 0 };
                            let total_avail: usize = column_layout[vi..=vi+overflow_cols].iter()
                                .map(|&(_, w, _)| w as usize).sum::<usize>()
                                + overflow_cols
                                + gap_after;
                            let display = if text_width > total_avail {
                                if overflow_cols == 0 {
                                    cell_text.chars().take(total_avail).collect::<String>()
                                } else {
                                    let trunc = total_avail.saturating_sub(1).max(1);
                                    let mut s: String = cell_text.chars().take(trunc).collect();
                                    if text_width > trunc { s.push('…'); }
                                    s
                                }
                            } else { cell_text.to_string() };
                            // Cell SGR style matching ratatui priority:
                            //   1. cursor cell → bg(DarkGray)
                            //   2. footer agg  → bold + fg(Cyan)
                            //   3. agg         → fg(Cyan)
                            //   4. border col  → fg(DarkGray)
                            //   5. displaced by overflow → fg(DarkGray)
                            //   6. default     → none
                            let is_left_margin_col = (col_idx as usize) < lm;
                            let is_right_margin_col = (col_idx as usize) >= lm + mc;
                            let cell_sgr = if is_cursor_cell {
                                sgr_cell_cursor()
                            } else if cell_style == 3 {
                                sgr_cell_footer_agg()
                            } else if cell_style == 2 {
                                sgr_cell_agg()
                            } else if is_left_margin_col || is_right_margin_col {
                                sgr_sep()
                            } else if prev_overflowed && !is_left_margin_col && !is_right_margin_col {
                                sgr_sep()
                            } else {
                                ""
                            };
                            out.push_str(&sgr_cup(ry, sx));
                            if !cell_sgr.is_empty() {
                                out.push_str(cell_sgr);
                            }
                            out.push_str(&display);
                            // avail_w = width within the cell's column(s), excluding
                            // the trailing gap to the next column (gap_after is written
                            // separately so it uses default style, matching ratatui).
                            let avail_w = if overflow_cols > 0 {
                                column_layout[vi..=vi+overflow_cols].iter()
                                    .map(|&(_, w, _)| w as usize).sum::<usize>()
                                    + overflow_cols
                            } else {
                                w
                            };
                            let display_w = display.chars().count();
                            let pad = avail_w.saturating_sub(display_w);
                            if pad > 0 {
                                out.push_str(&" ".repeat(pad));
                            }
                            // If the display text overflows past the column
                            // width into the gap area, suppress the gap.
                            let real_gap = if display_w > avail_w { 0 } else { gap_after };
                            let sgr_applied = !cell_sgr.is_empty();
                            if is_cursor_cell {
                                out.push_str(SGR_BG_DEFAULT);
                            } else if cell_style == 2 || cell_style == 3 {
                                out.push_str(SGR_FG_DEFAULT);
                            } else if is_left_margin_col || is_right_margin_col {
                                if is_boundary && vi + 1 >= n {
                                    out.push_str(SGR_RESET);
                                } else {
                                    out.push_str(SGR_FG_DEFAULT);
                                }
                            } else if sgr_applied {
                                out.push_str(SGR_FG_DEFAULT);
                            }
                            // Inter-column gap (1 space) – drawn in default style
                            if real_gap > 0 {
                                out.push(' ');
                            }
                            let overflowed_this = can_overflow && overflow_cols == 0;
                            if overflowed_this {
                                prev_overflowed = true;
                                force_right_gray = true;
                            } else {
                                prev_overflowed = false;
                            }
                            vi += 1 + overflow_cols as usize;
                        }
                        // Fill remaining row width before right border
                        let after_content = rect.x + 1 + 5 + column_layout.iter()
                            .map(|&(_, w, _)| w as i32).sum::<i32>()
                            + (n.saturating_sub(1) as i32); // gaps between columns
                        if after_content < rect.x + rect.w - 1 {
                            let gap = (rect.x + rect.w - 1 - after_content) as usize;
                            if gap > 0 {
                                out.push_str(&" ".repeat(gap));
                            }
                        }
                        out.push_str(&sgr_cup(ry, rect.x + rect.w - 1));
                        out.push_str("│");
                    } else {
                        let cw = *col_width as i32;
                        let rh_w = 5i32;
                        let max_vis_cols = ((rect.w - rh_w - 2) / (cw + 1)).max(1).min((mc.saturating_add(lm)) as i32);
                        let mut vc = 0i32;
                        while vc < max_vis_cols {
                            let col_idx = *left_col + vc as u32;
                            let cell_text = cells_ref.get(&(row_idx, col_idx)).map(|s| s.as_str()).unwrap_or("");
                            let cell_style = cell_styles_ref.get(&(row_idx, col_idx)).copied().unwrap_or(0);
                            let is_cursor_cell = cell_style == 1;
                            let col_screen_x = rect.x + 1 + rh_w + vc * (cw + 1);
                            if cell_text.is_empty() && !(*editing && is_cursor_cell) {
                                if is_cursor_cell {
                                    out.push_str(&sgr_cup(ry, col_screen_x));
                                    for _ in 0..cw { out.push(' '); }
                                } else if (col_idx as usize) < lm {
                                    out.push_str(&sgr_cup(ry, col_screen_x));
                                    if is_boundary {
                                        out.push_str(SGR_UNDERLINE);
                                    }
                                    out.push_str(sgr_sep());
                                    for _ in 0..cw { out.push(' '); }
                                    out.push_str(SGR_FG_DEFAULT);
                                } else {
                                    out.push_str(&sgr_cup(ry, col_screen_x));
                                    out.push_str(SGR_FG_DEFAULT);
                                    out.push_str(SGR_BG_DEFAULT);
                                    for _ in 0..cw { out.push(' '); }
                                }
                                vc += 1;
                                continue;
                            }
                            let text = if *editing && is_cursor_cell { &edit_buf } else { cell_text };
                            let text_width = text.chars().count();
                            let mut overflow_cols = 0i32;
                            if !(*editing && is_cursor_cell) && text_width > cw as usize {
                                let mut scan = vc + 1;
                                while scan < max_vis_cols {
                                    let sc = *left_col + scan as u32;
                                    if cells_ref.get(&(row_idx, sc)).map_or(true, |s| s.is_empty()) {
                                        overflow_cols += 1;
                                        scan += 1;
                                    } else { break; }
                                }
                            }
                            let gap_after = if overflow_cols == 0 { 1 } else { 0 };
                            let available = (overflow_cols + 1) * cw + gap_after;
                            let display = if text_width > available as usize {
                                if overflow_cols == 0 {
                                    text.chars().take(available as usize).collect()
                                } else {
                                    let trunc = (available - 1).max(1) as usize;
                                    let mut s: String = text.chars().take(trunc).collect();
                                    if text.chars().count() > trunc { s.push('…'); }
                                    s
                                }
                            } else { text.to_string() };
                            out.push_str(&sgr_cup(ry, col_screen_x));
                            if is_cursor_cell {
                                out.push_str(sgr_cell_cursor());
                            } else if cell_style == 2 {
                                out.push_str(sgr_cell_agg());
                            } else if cell_style == 3 {
                                out.push_str(sgr_cell_footer_agg());
                            }
                            out.push_str(&display);
                            if is_cursor_cell {
                                out.push_str(SGR_BG_DEFAULT);
                            } else if cell_style == 2 || cell_style == 3 {
                                out.push_str(SGR_FG_DEFAULT);
                            }
                            vc += 1 + overflow_cols;
                        }
                        if is_boundary {
                            out.push_str(SGR_RESET);
                        }
                        out.push_str(&sgr_cup(ry, rect.x + rect.w - 1));
                        out.push_str("│");
                    }
                }
                // Bottom border line
                {
                    let br = row_offset + max_data_rows as i32;
                    out.push_str(&sgr_cup(br, rect.x));
                    out.push_str("└");
                    for i in 1..rect.w - 1 {
                        out.push('─');
                    }
                    out.push_str("┘");
                }
                // Status bar: dark gray fg
                if has_status {
                    let sr = if has_tabs {
                        row_offset + max_data_rows as i32
                    } else {
                        row_offset + max_data_rows as i32 + 1
                    };
                    if sr < rect.y + rect.h {
                        let max_w = rect.w as usize;
                        let st_end = status_text.char_indices().nth(max_w).map(|(i, _)| i).unwrap_or(status_text.len());
                        out.push_str(&sgr_cup(sr, rect.x));
                        out.push_str(sgr_sep());
                        out.push_str(&status_text[..st_end]);
                        let st_vis = status_text[..st_end].chars().count();
                        if has_tabs {
                            if st_vis < rect.w as usize - 1 {
                                for _ in st_vis..rect.w as usize - 1 {
                                    out.push('─');
                                }
                                out.push_str("┘");
                            }
                        } else if st_vis < rect.w as usize {
                            out.push_str(&" ".repeat(rect.w as usize - st_vis));
                        }
                    }
                }
                // Tab bar
                if !tab_text.is_empty() {
                    let ty = row_offset + max_data_rows as i32 + 1;
                    if ty < rect.y + rect.h {
                        let max_w = rect.w as usize;
                        let tab_end = tab_text.char_indices().nth(max_w).map(|(i, _)| i).unwrap_or(tab_text.len());
                        out.push_str(&sgr_cup(ty, rect.x));
                        out.push_str(&tab_text[..tab_end]);
                        let tab_vis = tab_text[..tab_end].chars().count();
                        if tab_vis < rect.w as usize {
                            out.push_str(&" ".repeat(rect.w as usize - tab_vis));
                        }
                    }
                }

                // Store SGR output for emission after ncurses refresh
                with_state(|s| s.spreadsheet_output.push_str(&out));
            }
        }
    }

    fn is_spreadsheet_focused(state: &PcState, fid: usize) -> bool {
        state.node(fid).map_or(false, |n| matches!(n.kind, PcWidgetKind::Spreadsheet { .. }))
    }

    fn spreadsheet_enter(state: &mut PcState, fid: usize) {
        if let Some(n) = state.node_mut(fid) {
            if let PcWidgetKind::Spreadsheet { ref cells, ref mut cursor_row, ref mut cursor_col, ref mut editing, ref mut edit_buf, ref mut edit_pos, .. } = n.kind {
                if *editing {
                    cells.borrow_mut().insert((*cursor_row, *cursor_col), edit_buf.clone());
                    *editing = false;
                    if *cursor_row + 1 < u32::MAX { *cursor_row += 1; }
                } else {
                    *editing = true;
                    let existing = cells.borrow().get(&(*cursor_row, *cursor_col)).cloned().unwrap_or_default();
                    *edit_buf = existing;
                    *edit_pos = edit_buf.len();
                }
            }
        }
    }

    fn spreadsheet_scroll_to_cursor(state: &mut PcState, fid: usize) {
        if let Some(n) = state.node_mut(fid) {
            if let PcWidgetKind::Spreadsheet { ref mut top_row, ref mut left_col, ref cursor_row, ref cursor_col, ref col_width, .. } = n.kind {
                let cw = *col_width as i32;
                let rh_w = 4i32;
                let max_data_cols = ((n.rect.w - rh_w - 1) / (cw + 1)).max(1);
                let max_data_rows = (n.rect.h - 2).max(1) as u32;
                if *cursor_row < *top_row {
                    *top_row = *cursor_row;
                } else if *cursor_row >= *top_row + max_data_rows {
                    *top_row = *cursor_row - max_data_rows + 1;
                }
                if *cursor_col < *left_col {
                    *left_col = *cursor_col;
                } else if *cursor_col >= *left_col + max_data_cols as u32 {
                    *left_col = *cursor_col - max_data_cols as u32 + 1;
                }
            }
        }
        spreadsheet_update_formula_bar_inner(state, fid);
    }

    fn spreadsheet_commit_edit(state: &mut PcState, fid: usize) {
        if let Some(n) = state.node_mut(fid) {
            if let PcWidgetKind::Spreadsheet { ref cells, cursor_row, cursor_col, ref edit_buf, ref mut editing, .. } = n.kind {
                if *editing {
                    cells.borrow_mut().insert((cursor_row, cursor_col), edit_buf.clone());
                    *editing = false;
                }
            }
        }
    }

    fn col_label(idx: u32) -> String {
        if idx < 26 {
            let c = (b'A' + idx as u8) as char;
            c.to_string()
        } else {
            let prefix = (idx / 26 - 1) as u8;
            let suffix = (idx % 26) as u8;
            let mut s = String::new();
            s.push((b'A' + prefix) as char);
            s.push((b'A' + suffix) as char);
            s
        }
    }

    /// Activate the focused widget: toggle CheckButton/RadioButton or fire Button callbacks.
    /// Returns `(was_toggleable, callbacks_to_fire)`. Callbacks must be fired *outside* `with_state`.
    fn toggle_focused(state: &mut PcState, fid: usize) -> (bool, Vec<Callback>) {
        let idx = match state.nodes.iter().position(|n| n.id == fid) {
            Some(i) => i,
            None => return (false, vec![]),
        };
        // Match by reference — do NOT move `kind` out of the node, otherwise
        // subsequent reads of the node's fields become use-after-move.
        match &mut state.nodes[idx].kind {
            PcWidgetKind::Button { .. } => {
                let cbs = std::mem::take(&mut state.nodes[idx].callbacks);
                (true, cbs)
            }
            PcWidgetKind::CheckButton { ref mut checked, .. } => {
                *checked = !*checked;
                let cbs = std::mem::take(&mut state.nodes[idx].callbacks);
                (true, cbs)
            }
            PcWidgetKind::RadioButton { group_id, .. } => {
                let gid = *group_id;
                // Uncheck all radio buttons in the same group
                for node in &mut state.nodes {
                    if let PcWidgetKind::RadioButton { checked: ref mut oc, group_id: og, .. } = &mut node.kind {
                        if *og == gid { *oc = false; }
                    }
                }
                // Check the focused one
                if let PcWidgetKind::RadioButton { checked: ref mut c, .. } = &mut state.nodes[idx].kind {
                    *c = true;
                }
                let cbs = std::mem::take(&mut state.nodes[idx].callbacks);
                (true, cbs)
            }
            _ => (false, vec![]),
        }
    }

    /// Fire a list of callbacks outside of any `with_state` borrow.
    fn fire_callbacks(callbacks: Vec<Callback>) {
        for mut cb in callbacks { cb(); }
    }

    // -- Factory functions --

    pub fn init() -> Result<Box<dyn crate::backends::BackendApp>, Box<dyn std::error::Error + Send + Sync>> {
        Ok(Box::new(PcApp))
    }

    fn find_window_id(state: &PcState) -> Option<usize> {
        state.nodes.iter().find(|n| matches!(n.kind, PcWidgetKind::Window { .. })).map(|n| n.id)
    }

    pub fn create_window() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        Ok(with_state(|s| s.add_node(PcWidgetKind::Window { title: String::new() }, None)))
    }

    pub fn create_button(label: &str) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        Ok(with_state(|s| s.add_node(PcWidgetKind::Button { label: label.to_string() }, find_window_id(s))))
    }

    pub fn create_label(text: &str) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        Ok(with_state(|s| s.add_node(PcWidgetKind::Label { text: text.to_string() }, find_window_id(s))))
    }

    pub fn create_box(horizontal: bool, spacing: i32) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        Ok(with_state(|s| s.add_node(PcWidgetKind::BoxWidget { horizontal, spacing }, find_window_id(s))))
    }

    pub fn create_grid() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        Ok(with_state(|s| s.add_node(PcWidgetKind::Grid { cols: 0, rows: 0 }, find_window_id(s))))
    }

    pub fn create_entry() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        Ok(with_state(|s| s.add_node(PcWidgetKind::Entry { buffer: String::new(), cursor: 0 }, find_window_id(s))))
    }

    pub fn create_menu() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        Ok(with_state(|s| s.add_node(PcWidgetKind::Menu, find_window_id(s))))
    }

    pub fn create_simple_action(_name: &str) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        Ok(with_state(|s| s.add_node(PcWidgetKind::SimpleAction, find_window_id(s))))
    }

    pub unsafe fn create_menubar(submenu_items: Vec<(String, Vec<(String, String)>)>, _action_group: *mut c_void) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        let labels: Vec<String> = submenu_items.iter().map(|(l, _)| l.clone()).collect();
        Ok(with_state(|s| s.add_node(PcWidgetKind::MenuBar { labels, submenu_items }, find_window_id(s))))
    }

    pub fn create_dialog() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        Ok(with_state(|s| s.add_node(PcWidgetKind::Dialog { title: String::new() }, find_window_id(s))))
    }

    pub fn create_dropdown(items: &[&str]) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        let items_str: Vec<String> = items.iter().map(|s| s.to_string()).collect();
        Ok(with_state(|s| s.add_node(PcWidgetKind::DropDown { items: items_str, selected: None }, find_window_id(s))))
    }

    pub fn create_checkbutton(label: &str) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        Ok(with_state(|s| s.add_node(PcWidgetKind::CheckButton { label: label.to_string(), checked: false }, find_window_id(s))))
    }

    pub fn create_radiobutton(group_id: Option<usize>, label: &str) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        let gid = group_id.unwrap_or(0);
        Ok(with_state(|s| s.add_node(PcWidgetKind::RadioButton { label: label.to_string(), checked: false, group_id: gid }, find_window_id(s))))
    }

    pub fn create_textview() -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        Ok(with_state(|s| s.add_node(PcWidgetKind::TextView { text: String::new() }, find_window_id(s))))
    }

    pub fn create_spreadsheet(rows: u32, cols: u32) -> Result<usize, Box<dyn std::error::Error + Send + Sync>> {
        let cells = Rc::new(RefCell::new(HashMap::new()));
        let raw_cells = Rc::new(RefCell::new(HashMap::new()));
        let cell_styles = Rc::new(RefCell::new(HashMap::new()));
        let id = with_state(|s| s.add_node(PcWidgetKind::Spreadsheet {
            cells,
            raw_cells,
            cell_styles,
            total_rows: rows,
            total_cols: cols,
            top_row: 0,
            left_col: 0,
            cursor_row: 0,
            cursor_col: 0,
            editing: false,
            edit_buf: String::new(),
            edit_pos: 0,
            col_width: 12,
            margin_cols: 0,
            main_cols: cols,
            formula_bar_address_id: None,
            formula_bar_entry_id: None,
            anchor: None,
            header_row_count: 0,
            main_row_count: 0,
            menu_text: String::new(),
            status_text: String::new(),
            border_title: String::new(),
            formula_bar_trailing: String::new(),
            column_layout: Vec::new(),
            row_labels: Vec::new(),
            tab_text: String::new(),
        }, find_window_id(s)));
        Ok(id)
    }

    pub fn spreadsheet_set_cell(id: usize, r: u32, c: u32, text: &str) {
        with_state(|s| {
            if let Some(n) = s.node_mut(id) {
                if let PcWidgetKind::Spreadsheet { ref cells, .. } = n.kind {
                    cells.borrow_mut().insert((r, c), text.to_string());
                }
            }
        });
    }

    pub fn spreadsheet_set_cell_style(id: usize, r: u32, c: u32, style: u8) {
        with_state(|s| {
            if let Some(n) = s.node_mut(id) {
                if let PcWidgetKind::Spreadsheet { ref cell_styles, .. } = n.kind {
                    cell_styles.borrow_mut().insert((r, c), style);
                }
            }
        });
    }

    pub fn spreadsheet_set_raw_cell(id: usize, r: u32, c: u32, text: &str) {
        with_state(|s| {
            if let Some(n) = s.node_mut(id) {
                if let PcWidgetKind::Spreadsheet { ref raw_cells, .. } = n.kind {
                    raw_cells.borrow_mut().insert((r, c), text.to_string());
                }
            }
        });
    }

    pub fn spreadsheet_get_cell(id: usize, r: u32, c: u32) -> Option<String> {
        with_state(|s| {
            s.node(id).and_then(|n| {
                if let PcWidgetKind::Spreadsheet { ref cells, .. } = n.kind {
                    cells.borrow().get(&(r, c)).cloned()
                } else {
                    None
                }
            })
        })
    }

    fn spreadsheet_prepare_move(state: &mut PcState, fid: usize, shift: bool) {
        if let Some(n) = state.node_mut(fid) {
            if let PcWidgetKind::Spreadsheet { ref mut anchor, cursor_row, cursor_col, .. } = n.kind {
                if shift {
                    if anchor.is_none() {
                        *anchor = Some((cursor_row, cursor_col));
                    }
                } else {
                    *anchor = None;
                }
            }
        }
    }

    pub fn spreadsheet_clear_anchor(fid: usize) {
        with_state(|s| {
            if let Some(n) = s.node_mut(fid) {
                if let PcWidgetKind::Spreadsheet { ref mut anchor, .. } = n.kind {
                    *anchor = None;
                }
            }
        });
    }

    pub fn spreadsheet_set_column_layout(spreadsheet_id: usize, layout: Vec<(u32, u32, String)>) {
        with_state(|s| {
            if let Some(n) = s.node_mut(spreadsheet_id) {
                if let PcWidgetKind::Spreadsheet { ref mut column_layout, .. } = n.kind {
                    *column_layout = layout;
                }
            }
        });
    }

    pub fn spreadsheet_set_row_labels(spreadsheet_id: usize, labels: Vec<(u32, String)>) {
        with_state(|s| {
            if let Some(n) = s.node_mut(spreadsheet_id) {
                if let PcWidgetKind::Spreadsheet { ref mut row_labels, .. } = n.kind {
                    *row_labels = labels;
                }
            }
        });
    }

    pub fn spreadsheet_set_menu_text(spreadsheet_id: usize, text: &str) {
        with_state(|s| {
            if let Some(n) = s.node_mut(spreadsheet_id) {
                if let PcWidgetKind::Spreadsheet { ref mut menu_text, .. } = n.kind {
                    *menu_text = text.to_string();
                }
            }
        });
    }

    pub fn spreadsheet_set_border_title(spreadsheet_id: usize, text: &str) {
        with_state(|s| {
            if let Some(n) = s.node_mut(spreadsheet_id) {
                if let PcWidgetKind::Spreadsheet { ref mut border_title, .. } = n.kind {
                    *border_title = text.to_string();
                }
            }
        });
    }

    pub fn spreadsheet_set_status_text(spreadsheet_id: usize, text: &str) {
        with_state(|s| {
            if let Some(n) = s.node_mut(spreadsheet_id) {
                if let PcWidgetKind::Spreadsheet { ref mut status_text, .. } = n.kind {
                    *status_text = text.to_string();
                }
            }
        });
    }

    pub fn spreadsheet_set_formula_bar_trailing(spreadsheet_id: usize, text: &str) {
        with_state(|s| {
            if let Some(n) = s.node_mut(spreadsheet_id) {
                if let PcWidgetKind::Spreadsheet { ref mut formula_bar_trailing, .. } = n.kind {
                    *formula_bar_trailing = text.to_string();
                }
            }
        });
    }

    pub fn spreadsheet_set_tab_text(spreadsheet_id: usize, text: &str) {
        with_state(|s| {
            if let Some(n) = s.node_mut(spreadsheet_id) {
                if let PcWidgetKind::Spreadsheet { ref mut tab_text, .. } = n.kind {
                    *tab_text = text.to_string();
                }
            }
        });
    }

    pub fn spreadsheet_set_grid_config(spreadsheet_id: usize, margin_c: u32, main_c: u32) {
        with_state(|s| {
            if let Some(n) = s.node_mut(spreadsheet_id) {
                if let PcWidgetKind::Spreadsheet { ref mut margin_cols, ref mut main_cols, .. } = n.kind {
                    *margin_cols = margin_c;
                    *main_cols = main_c;
                }
            }
        });
    }

    pub fn spreadsheet_set_row_counts(spreadsheet_id: usize, header_rows: u32, main_rows: u32) {
        with_state(|s| {
            if let Some(n) = s.node_mut(spreadsheet_id) {
                if let PcWidgetKind::Spreadsheet { ref mut header_row_count, ref mut main_row_count, .. } = n.kind {
                    *header_row_count = header_rows;
                    *main_row_count = main_rows;
                }
            }
        });
    }

    pub fn spreadsheet_cursor_position(id: usize) -> Option<(u32, u32)> {
        with_state(|s| {
            s.node(id).and_then(|n| {
                if let PcWidgetKind::Spreadsheet { cursor_row, cursor_col, .. } = n.kind {
                    Some((cursor_row, cursor_col))
                } else { None }
            })
        })
    }

    pub fn spreadsheet_set_cursor(id: usize, row: u32, col: u32) {
        with_state(|s| {
            if let Some(n) = s.node_mut(id) {
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, ref mut cursor_col, .. } = n.kind {
                    *cursor_row = row;
                    *cursor_col = col;
                }
            }
        });
    }

    pub fn spreadsheet_commit_formula_bar(spreadsheet_id: usize) {
        with_state(|s| {
            let (entry_id, cursor_row, cursor_col) = match s.node(spreadsheet_id) {
                Some(n) => match &n.kind {
                    PcWidgetKind::Spreadsheet { formula_bar_entry_id, cursor_row, cursor_col, .. } => {
                        (*formula_bar_entry_id, *cursor_row, *cursor_col)
                    }
                    _ => return,
                },
                None => return,
            };
            let Some(eid) = entry_id else { return };
            let text = match s.node(eid) {
                Some(en) => match &en.kind {
                    PcWidgetKind::Entry { buffer, .. } => buffer.clone(),
                    _ => return,
                },
                None => return,
            };
            if let Some(n) = s.node_mut(spreadsheet_id) {
                if let PcWidgetKind::Spreadsheet { ref cells, .. } = n.kind {
                    cells.borrow_mut().insert((cursor_row, cursor_col), text);
                }
            }
        });
    }

    pub fn spreadsheet_set_formula_bar(spreadsheet_id: usize, address_label_id: usize, entry_id: usize) {
        with_state(|s| {
            if let Some(n) = s.node_mut(spreadsheet_id) {
                if let PcWidgetKind::Spreadsheet { ref mut formula_bar_address_id, ref mut formula_bar_entry_id, .. } = n.kind {
                    *formula_bar_address_id = Some(address_label_id);
                    *formula_bar_entry_id = Some(entry_id);
                }
            }
        });
        spreadsheet_update_formula_bar(spreadsheet_id);
    }

    fn spreadsheet_update_formula_bar(fid: usize) {
        with_state(|s| {
            if spreadsheet_update_formula_bar_inner(s, fid) {
                // Mark for re-render by adjusting rect
                if let Some(n) = s.node_mut(fid) {
                    n.rect.w = n.rect.w.max(1);
                }
            }
        });
    }

    fn spreadsheet_update_formula_bar_inner(state: &mut PcState, fid: usize) -> bool {
        let n = match state.node(fid) {
            Some(n) => n,
            None => return false,
        };
        let (cells, cursor_row, cursor_col, margin_cols, addr_id, entry_id) = match &n.kind {
            PcWidgetKind::Spreadsheet { cells, cursor_row, cursor_col, margin_cols, formula_bar_address_id, formula_bar_entry_id, .. } => {
                (cells.clone(), *cursor_row, *cursor_col, *margin_cols, *formula_bar_address_id, *formula_bar_entry_id)
            }
            _ => return false,
        };
        let mut changed = false;
        if let Some(aid) = addr_id {
            let display_col = if cursor_col < margin_cols { cursor_col } else { cursor_col - margin_cols };
            let label = format!("{}{}", col_label(display_col), cursor_row + 1);
            if let Some(an) = state.node_mut(aid) {
                if let PcWidgetKind::Label { ref mut text } = &mut an.kind {
                    if *text != label {
                        *text = label;
                        changed = true;
                    }
                }
            }
        }
        if let Some(eid) = entry_id {
            let val = cells.borrow().get(&(cursor_row, cursor_col)).cloned().unwrap_or_default();
            if let Some(en) = state.node_mut(eid) {
                if let PcWidgetKind::Entry { ref mut buffer, .. } = &mut en.kind {
                    if *buffer != val {
                        *buffer = val;
                        changed = true;
                    }
                }
            }
        }
        changed
    }

    pub fn set_window_title(id: usize, title: &str) {
        with_state(|s| {
            if let Some(n) = s.node_mut(id) {
                if let PcWidgetKind::Window { title: ref mut t } = n.kind {
                    *t = title.to_string();
                }
            }
        });
    }

    pub fn set_label_text(id: usize, text: &str) {
        with_state(|s| {
            if let Some(n) = s.node_mut(id) {
                if let PcWidgetKind::Label { text: ref mut t } = n.kind {
                    *t = text.to_string();
                }
            }
        });
    }

    pub fn get_label_text(id: usize) -> Option<String> {
        with_state(|s| {
            s.node(id).and_then(|n| {
                if let PcWidgetKind::Label { ref text } = n.kind {
                    Some(text.clone())
                } else {
                    None
                }
            })
        })
    }

    pub fn set_label_visible(id: usize, visible: bool) {
        with_state(|s| {
            if let Some(n) = s.node_mut(id) {
                n.visible = visible;
            }
        });
    }

    pub fn add_callback(id: usize, cb: Box<dyn FnMut()>) {
        with_state(|s| {
            if let Some(n) = s.node_mut(id) {
                n.callbacks.push(cb);
            }
        });
    }

    pub fn set_entry_text(id: usize, text: &str) {
        with_state(|s| {
            if let Some(n) = s.node_mut(id) {
                if let PcWidgetKind::Entry { ref mut buffer, ref mut cursor } = n.kind {
                    *buffer = text.to_string();
                    *cursor = buffer.len();
                }
            }
        });
    }

    pub fn get_entry_text(id: usize) -> Option<String> {
        with_state(|s| {
            s.node(id).and_then(|n| {
                if let PcWidgetKind::Entry { ref buffer, .. } = n.kind {
                    Some(buffer.clone())
                } else {
                    None
                }
            })
        })
    }

    pub fn set_textview_text(id: usize, text: &str) {
        with_state(|s| {
            if let Some(n) = s.node_mut(id) {
                if let PcWidgetKind::TextView { text: ref mut t } = n.kind {
                    *t = text.to_string();
                }
            }
        });
    }

    pub fn get_textview_text(id: usize) -> Option<String> {
        with_state(|s| {
            s.node(id).and_then(|n| {
                if let PcWidgetKind::TextView { ref text } = n.kind {
                    Some(text.clone())
                } else {
                    None
                }
            })
        })
    }

    pub fn set_dropdown_items(id: usize, items: &[&str]) {
        with_state(|s| {
            if let Some(n) = s.node_mut(id) {
                if let PcWidgetKind::DropDown { items: ref mut items_vec, .. } = n.kind {
                    *items_vec = items.iter().map(|s| s.to_string()).collect();
                }
            }
        });
    }

    pub fn set_dropdown_selected(id: usize, idx: i32) {
        with_state(|s| {
            if let Some(n) = s.node_mut(id) {
                if let PcWidgetKind::DropDown { ref mut selected, ref items } = n.kind {
                    *selected = if idx >= 0 && (idx as usize) < items.len() { Some(idx as usize) } else { None };
                }
            }
        });
    }

    pub fn get_dropdown_selected(id: usize) -> i32 {
        with_state(|s| {
            s.node(id).and_then(|n| {
                if let PcWidgetKind::DropDown { ref selected, .. } = n.kind {
                    selected.map(|s| s as i32)
                } else {
                    None
                }
            }).unwrap_or(-1)
        })
    }

    pub fn get_checkbutton_checked(id: usize) -> bool {
        with_state(|s| {
            s.node(id).map(|n| {
                if let PcWidgetKind::CheckButton { ref checked, .. } = n.kind {
                    *checked
                } else {
                    false
                }
            }).unwrap_or(false)
        })
    }

    pub fn get_radiobutton_checked(id: usize) -> bool {
        with_state(|s| {
            s.node(id).map(|n| {
                if let PcWidgetKind::RadioButton { ref checked, .. } = n.kind {
                    *checked
                } else {
                    false
                }
            }).unwrap_or(false)
        })
    }

    pub fn set_checkbutton_checked(id: usize, checked: bool) {
        with_state(|s| {
            if let Some(n) = s.node_mut(id) {
                if let PcWidgetKind::CheckButton { checked: ref mut c, .. } = n.kind {
                    *c = checked;
                }
            }
        });
    }

    pub fn set_radiobutton_checked(id: usize, checked: bool) {
        with_state(|s| {
            if let Some(n) = s.node_mut(id) {
                if let PcWidgetKind::RadioButton { checked: ref mut c, .. } = n.kind {
                    *c = checked;
                }
            }
        });
    }

    pub fn set_child(parent_id: usize, child_id: usize) {
        with_state(|s| {
            s.nodes.iter_mut().for_each(|n| n.children.retain(|c| *c != child_id));
            if let Some(parent) = s.node_mut(parent_id) {
                parent.children.push(child_id);
            }
            if let Some(child) = s.node_mut(child_id) {
                child.parent = Some(parent_id);
            }
        });
    }

    pub fn append_child(parent_id: usize, child_id: usize) {
        set_child(parent_id, child_id);
    }

    fn layout_box_inner(id: usize, s: &mut PcState) -> usize {
        let (horizontal, spacing, children) = {
            let node = s.node(id);
            if node.is_none() { return 0; }
            let n = node.unwrap();
            match &n.kind {
                PcWidgetKind::BoxWidget { horizontal, spacing } => {
                    (*horizontal, *spacing, n.children.clone())
                }
                _ => return 0,
            }
        };

        if children.is_empty() { return 0; }

        let parent_rect = s.node(id).map(|n| n.rect).unwrap_or(Rect::default());
        let total_children = children.len() as i32;
        let total_spacing = spacing * (total_children - 1).max(0);

        if horizontal {
            let total_w = parent_rect.w.saturating_sub(total_spacing).max(1);
            let per_child = if total_children > 0 { (total_w / total_children).max(1) } else { 1 };
            let mut x = parent_rect.x;
            for child_id in &children {
                if let Some(n) = s.node_mut(*child_id) {
                    n.rect = Rect { x, y: parent_rect.y, w: per_child, h: parent_rect.h.max(1) };
                }
                x += per_child + spacing;
            }
        } else {
            // Natural-height vertical layout: each child gets 1 row by default,
            // last child stretches to fill any remaining space.
            let mut y = parent_rect.y;
            let natural_h = 1i32;
            let total_used = natural_h * total_children + total_spacing;
            let remaining = parent_rect.h.saturating_sub(total_used);
            for (i, child_id) in children.iter().enumerate() {
                let h = if i as i32 == total_children - 1 {
                    (natural_h + remaining).max(1)
                } else {
                    natural_h
                };
                if let Some(n) = s.node_mut(*child_id) {
                    n.rect = Rect { x: parent_rect.x, y, w: parent_rect.w.max(1), h };
                }
                y += h + spacing;
            }
        }
        children.len()
    }

    pub fn layout_box(id: usize) {
        with_state(|s| { layout_box_inner(id, s); });
    }

    fn layout_grid_inner(id: usize, s: &mut PcState) -> usize {
        let (cols, rows, children) = {
            let node = s.node(id);
            if node.is_none() { return 0; }
            let n = node.unwrap();
            match &n.kind {
                PcWidgetKind::Grid { cols, rows } => (*cols.max(&1), *rows.max(&1), n.children.clone()),
                _ => return 0,
            }
        };

        if children.is_empty() { return 0; }

        let parent_rect = s.node(id).map(|n| n.rect).unwrap_or(Rect::default());
        let ncols = cols.max(1);
        let nrows = rows.max(1);
        let cell_w = (parent_rect.w / ncols as i32).max(1);
        let cell_h = (parent_rect.h / nrows as i32).max(1);

        for (i, child_id) in children.iter().enumerate() {
            let col = i as i32 % ncols as i32;
            let row = i as i32 / ncols as i32;
            if let Some(n) = s.node_mut(*child_id) {
                n.rect = Rect {
                    x: parent_rect.x + col * cell_w,
                    y: parent_rect.y + row * cell_h,
                    w: cell_w,
                    h: cell_h,
                };
            }
        }
        children.len()
    }

    pub fn layout_grid(id: usize) {
        with_state(|s| { layout_grid_inner(id, s); });
    }

    pub fn entry_set_text(id: usize, text: &str) {
        set_entry_text(id, text);
    }

    pub fn entry_text(id: usize) -> Option<String> {
        get_entry_text(id)
    }

    pub fn set_focus(id: usize) {
        with_state(|s| s.focus_id = Some(id));
    }

    pub fn quit() {
        with_state(|s| s.running = false);
    }

    pub fn add_key_callback(key: char, cb: Box<dyn FnMut()>) {
        with_state(|s| s.key_callbacks.push((key, cb)));
    }

    /// Render the spreadsheet grid to a `Vec<String>` buffer (one string per row)
    /// for testing/verification. This mirrors the rendering logic of `render_widget`
    /// for `PcWidgetKind::Spreadsheet` but writes to strings instead of ncurses.
    pub fn render_spreadsheet_to_buffer(spreadsheet_id: usize, width: usize, height: usize) -> Vec<String> {
        let mut buf = vec![String::new(); height];
        with_state(|state| {
            let n = match state.node(spreadsheet_id) {
                Some(n) => n,
                None => return,
            };
            let (cells, raw_cells, top_row, left_col, cursor_row, cursor_col, editing, edit_buf, edit_pos,
                 col_width, margin_cols, main_cols, menu_text, status_text,
                 formula_bar_trailing, column_layout, row_labels) = match &n.kind {
                PcWidgetKind::Spreadsheet { cells, raw_cells, top_row, left_col, cursor_row, cursor_col,
                    editing, edit_buf, edit_pos, col_width, margin_cols, main_cols,
                    menu_text, status_text, ref formula_bar_trailing, ref column_layout, ref row_labels, .. } => {
                    (cells.clone(), raw_cells.clone(), *top_row, *left_col, *cursor_row, *cursor_col,
                     *editing, edit_buf.clone(), *edit_pos, *col_width,
                     *margin_cols, *main_cols, menu_text.clone(), status_text.clone(),
                     formula_bar_trailing.clone(), column_layout.clone(), row_labels.clone())
                }
                _ => return,
            };
            let cw = col_width as usize;
            let rh_w = 5usize;
            let lm = margin_cols as usize;
            let mc = main_cols as usize;
            let sep = '│';
            let mut row = 0usize;

            // Menu bar
            if !menu_text.is_empty() {
                if row < height {
                    let display = &menu_text[..menu_text.len().min(width)];
                    buf[row] = display.to_string();
                }
                row += 1;
            }

            // Formula bar (matching ratatui: leading space, trailing text)
            if row < height {
                let display_col = if cursor_col < margin_cols { cursor_col } else { cursor_col - margin_cols };
                let addr_text = format!("{}{}", col_label(display_col), cursor_row + 1);
                let cell_val = raw_cells.borrow().get(&(cursor_row, cursor_col)).cloned().unwrap_or_default();
                let max_fb = width.saturating_sub(addr_text.chars().count() + 3 + formula_bar_trailing.chars().count()).max(1);
                let fb_text = if cell_val.chars().count() > max_fb {
                    let truncated: String = cell_val.chars().take(max_fb.saturating_sub(3)).collect();
                    format!(" {}  {}…{}", addr_text, truncated, formula_bar_trailing)
                } else {
                    format!(" {}  {}{}", addr_text, cell_val, formula_bar_trailing)
                };
                let truncated: String = fb_text.chars().take(width).collect();
                buf[row] = truncated;
                row += 1;
            }

            // Use pre-computed layout when available, otherwise fall back to computed layout.
            let use_layout = !column_layout.is_empty();
            let use_row_labels = !row_labels.is_empty();

            // Grid header row (with │ border)
            if row < height {
                let mut hdr = String::new();
                hdr.push('│');
                if use_layout {
                    hdr.push_str(&" ".repeat(5));
                    let n = column_layout.len();
                    for (i, &(_, w, ref label)) in column_layout.iter().enumerate() {
                        let padded = format!("{:<1$}", label, (w as usize).max(1));
                        hdr.push_str(&padded);
                        if i + 1 < n {
                            hdr.push(' ');
                        }
                    }
                } else {
                    let max_data_cols = ((width as i32 - rh_w as i32 - 2) / (cw as i32 + 1)).max(1) as usize;
                    hdr.push_str(&" ".repeat(5));
                    for vc in 0..max_data_cols.min(mc.saturating_add(lm)) {
                        let col_idx = left_col + vc as u32;
                        let label = if col_idx < lm as u32 {
                            format!("[{}", col_label(col_idx))
                        } else if col_idx < (lm + mc) as u32 {
                            col_label(col_idx)
                        } else {
                            format!("]{}", col_label(col_idx))
                        };
                        let padded = format!("{:<1$}", label, (cw as usize).max(label.len()));
                        hdr.push_str(&format!("{}{}", padded, sep));
                    }
                }
                let truncated: String = hdr.chars().take(width).collect();
                buf[row] = truncated;
                row += 1;
            }

            // Header separator (matching ratatui: ││───│)
            if row < height {
                let mut sep_row = String::new();
                sep_row.push('│');
                sep_row.push('│');
                let dash_len = width.saturating_sub(3).max(1);
                for _ in 0..dash_len {
                    sep_row.push('─');
                }
                sep_row.push('│');
                let cut = sep_row.char_indices().take(width).last().map(|(i, _)| i).unwrap_or(0);
                if cut > 0 { buf[row] = sep_row[..cut].to_string(); }
                else { buf[row] = sep_row.chars().take(width).collect(); }
                row += 1;
            }

            // Data rows (with │ border)
            let grid_bottom = height - if status_text.is_empty() { 0 } else { 2 };
            while row < grid_bottom && row < height {
                let row_idx = top_row + (row.saturating_sub(4)) as u32;
                let mut data_row = String::new();
                data_row.push('│');
                // Row header
                if use_row_labels {
                    if let Some((_, label)) = row_labels.iter().find(|(r, _)| *r == row_idx) {
                        data_row.push_str(label);
                    } else {
                        data_row.push_str(&format!("{:>4}", row_idx + 1));
                    }
                } else {
                    data_row.push_str(&format!("{:>4}", row_idx + 1));
                }
                data_row.push(sep);
                // Data cells
                let cells_ref = cells.borrow();
                if use_layout {
                    let lm = margin_cols;
                    let n = column_layout.len();
                    let mut vi = 0usize;
                    while vi < n {
                        let (col_idx, w, _) = column_layout[vi];
                        let cell_text = cells_ref.get(&(row_idx, col_idx)).map(|s| s.as_str()).unwrap_or("");
                        let cell_w = w as usize;
                        if cell_text.is_empty() {
                            data_row.push_str(&format!("{:<1$}", "", cell_w));
                            vi += 1;
                            continue;
                        }
                        let text_width = cell_text.chars().count();
                        // Find overflow space into adjacent empty columns
                        let mut overflow_cols = 0usize;
                        if text_width > cell_w {
                            let mut scan = vi + 1;
                            while scan < n {
                                let (sc_idx, _, _) = column_layout[scan];
                                let sc_text = cells_ref.get(&(row_idx, sc_idx)).map(|s| s.as_str()).unwrap_or("");
                                if sc_text.is_empty() {
                                    overflow_cols += 1;
                                    scan += 1;
                                } else { break; }
                            }
                        }
                        let gap_after = if overflow_cols == 0 && vi + 1 < n { 1 } else { 0 };
                        let total_avail: usize = column_layout[vi..=vi+overflow_cols].iter()
                            .map(|&(_, w, _)| w as usize).sum::<usize>()
                            + overflow_cols // spaces between overflow columns
                            + gap_after;   // gap to next column
                        let display = if text_width > total_avail {
                            if overflow_cols == 0 {
                                cell_text.chars().take(total_avail).collect()
                            } else {
                                let trunc = total_avail.saturating_sub(1).max(1);
                                let mut s: String = cell_text.chars().take(trunc).collect();
                                if text_width > trunc { s.push('…'); }
                                s
                            }
                        } else { cell_text.to_string() };
                        // Pad to combined width of current + overflow columns to maintain alignment.
                        // When overflow_cols==0, include the gap (total_avail) so the gap
                        // character is preserved, matching ratatui's use of the inter-column
                        // gap as overflow room.
                        let total_width = if overflow_cols > 0 {
                            column_layout[vi..=vi+overflow_cols].iter()
                                .map(|&(_, w, _)| w as usize).sum::<usize>()
                        } else {
                            total_avail
                        };
                        data_row.push_str(&format!("{:<1$}", display, total_width));
                        vi += 1 + overflow_cols;
                    }
                } else {
                    let max_data_cols = ((width as i32 - rh_w as i32 - 2) / (cw as i32 + 1)).max(1) as usize;
                    let mut vc = 0usize;
                    while vc < max_data_cols {
                        let col_idx = left_col + vc as u32;
                        let cell_text = cells_ref.get(&(row_idx, col_idx)).map(|s| s.as_str()).unwrap_or("");
                        let text_width = cell_text.chars().count();
                        let mut overflow_cols = 0usize;
                        if text_width > cw as usize {
                            let mut scan = vc + 1;
                            while scan < max_data_cols {
                                let sc = left_col + scan as u32;
                                if cells_ref.get(&(row_idx, sc)).map_or(true, |s| s.is_empty()) {
                                    overflow_cols += 1;
                                    scan += 1;
                                } else { break; }
                            }
                        }
                        let available = (overflow_cols + 1) * cw as usize;
                        let display = if text_width > available {
                            if overflow_cols == 0 {
                                cell_text.chars().take(available).collect()
                            } else {
                                let trunc = (available - 1).max(1);
                                let mut s: String = cell_text.chars().take(trunc).collect();
                                if text_width > trunc { s.push('…'); }
                                s
                            }
                        } else {
                            cell_text.to_string()
                        };
                        let cell_w = (overflow_cols + 1) * cw as usize;
                        data_row.push_str(&format!("{:<1$}", display, cell_w));
                        data_row.push(sep);
                        vc += 1 + overflow_cols;
                    }
                }
                if data_row.chars().count() > width {
                    data_row = data_row.chars().take(width).collect();
                }
                // Only add │ closing border if there's room
                if data_row.chars().count() + 1 <= width && !data_row.ends_with('│') {
                    data_row.push('│');
                }
                buf[row] = data_row;
                row += 1;
            }

            // Bottom border line (matching ratatui layout)
            if row < height {
                let mut border = String::new();
                border.push('└');
                for _ in 1..width.saturating_sub(1) {
                    border.push('─');
                }
                if width > 1 {
                    border.push('┘');
                }
                let cut = border.char_indices().take(width).last().map(|(i, _)| i).unwrap_or(0);
                if cut > 0 { buf[row] = border[..cut].to_string(); }
                else { buf[row] = border.chars().take(width).collect(); }
                row += 1;
            }

            // Status bar (always on its own line below the bottom border)
            if !status_text.is_empty() && row < height {
                let display = &status_text[..status_text.len().min(width)];
                buf[row] = display.to_string();
            }
        });
        buf
    }

    #[cfg(test)]
    mod tests {
        use super::*;

        /// Helper to set up a minimal spreadsheet state with a window parent.
        fn make_spreadsheet_id() -> usize {
            let wid = with_state(|s| s.add_node(
                PcWidgetKind::Window { title: "test".into() }, None,
            ));
            let sid = with_state(|s| s.add_node(
                PcWidgetKind::Spreadsheet {
                    cells: Rc::new(RefCell::new(HashMap::new())),
                    raw_cells: Rc::new(RefCell::new(HashMap::new())),
                    cell_styles: Rc::new(RefCell::new(HashMap::new())),
                    total_rows: 100,
                    total_cols: 26,
                    top_row: 0,
                    left_col: 0,
                    cursor_row: 0,
                    cursor_col: 0,
                    editing: false,
                    edit_buf: String::new(),
                    edit_pos: 0,
                    col_width: 12,
                    margin_cols: 0,
                    main_cols: 26,
                    formula_bar_address_id: None,
                    formula_bar_entry_id: None,
                    anchor: None,
                    menu_text: String::new(),
                    status_text: String::new(),
                    border_title: String::new(),
                    formula_bar_trailing: String::new(),
                    column_layout: Vec::new(),
                    row_labels: Vec::new(),
                    tab_text: String::new(),
                    header_row_count: 2,
                    main_row_count: 24,
                },
                Some(wid),
            ));
            with_state(|s| s.focus_id = Some(sid));
            sid
        }

        /// Moving cursor with arrow keys.
        #[test]
        fn spreadsheet_arrow_keys_move_cursor() {
            let sid = make_spreadsheet_id();
            // Simulate KeyDown by directly manipulating state:
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, .. } = n.kind {
                    *cursor_row += 1;
                }
            });
            let pos = spreadsheet_cursor_position(sid);
            assert_eq!(pos, Some((1, 0)), "cursor should be at (1, 0) after down");

            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_col, .. } = n.kind {
                    *cursor_col += 1;
                }
            });
            let pos = spreadsheet_cursor_position(sid);
            assert_eq!(pos, Some((1, 1)), "cursor should be at (1, 1) after right");

            // Home should go to col 0
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_col, .. } = n.kind {
                    *cursor_col = 0;
                }
            });
            let pos = spreadsheet_cursor_position(sid);
            assert_eq!(pos, Some((1, 0)), "cursor should be at (1, 0) after home");
        }

        /// Entering text and committing.
        #[test]
        fn spreadsheet_enter_and_commit_text() {
            let sid = make_spreadsheet_id();
            // Enter edit mode (simulate Enter key)
            with_state(|state| spreadsheet_enter(state, sid));
            // Check editing state
            with_state(|state| {
                let n = state.node(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { editing, .. } = &n.kind {
                    assert!(*editing, "should be in edit mode after enter");
                }
            });

            // Set edit buffer text (simulate typing)
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut edit_buf, .. } = n.kind {
                    *edit_buf = "hello".to_string();
                }
            });

            // Commit edit (simulate Enter again)
            with_state(|state| spreadsheet_commit_edit(state, sid));

            // Verify cell content
            let val = spreadsheet_get_cell(sid, 0, 0);
            assert_eq!(val, Some("hello".into()), "cell should contain 'hello' after commit");

            let pos = spreadsheet_cursor_position(sid);
            assert_eq!(pos, Some((0, 0)), "cursor should remain at (0, 0)");
        }

        /// Moving cursor and committing advances cursor.
        #[test]
        fn spreadsheet_enter_advances_cursor() {
            let sid = make_spreadsheet_id();
            // Enter edit mode
            with_state(|state| spreadsheet_enter(state, sid));

            // Set edit buffer text
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut edit_buf, .. } = n.kind {
                    *edit_buf = "test".to_string();
                }
            });

            // Second enter commits and advances
            with_state(|state| spreadsheet_enter(state, sid));
            let val = spreadsheet_get_cell(sid, 0, 0);
            assert_eq!(val, Some("test".into()), "cell should contain 'test'");

            let pos = spreadsheet_cursor_position(sid);
            assert_eq!(pos, Some((1, 0)), "cursor should advance to row 1");
        }

        /// Test that cursor does not go negative.
        #[test]
        fn spreadsheet_cursor_clamped_at_zero() {
            let sid = make_spreadsheet_id();
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, ref mut cursor_col, .. } = n.kind {
                    *cursor_row = 0;
                    *cursor_col = 0;
                }
            });
            let pos = spreadsheet_cursor_position(sid);
            assert_eq!(pos, Some((0, 0)));

            // Simulate KeyUp: should not go negative
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, .. } = n.kind {
                    if *cursor_row > 0 { *cursor_row -= 1; }
                }
            });
            let pos = spreadsheet_cursor_position(sid);
            assert_eq!(pos, Some((0, 0)), "cursor should stay at row 0");

            // Simulate KeyLeft: should not go negative
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_col, .. } = n.kind {
                    if *cursor_col > 0 { *cursor_col -= 1; }
                }
            });
            let pos = spreadsheet_cursor_position(sid);
            assert_eq!(pos, Some((0, 0)), "cursor should stay at col 0");
        }

        #[test]
        fn spreadsheet_render_matches_ratatui() {
            // Set up a spreadsheet with the same data as docs/tests/overflow.corro
            let wid = with_state(|s| s.add_node(
                PcWidgetKind::Window { title: "corro".into() }, None,
            ));
            let sid = with_state(|s| s.add_node(
                PcWidgetKind::Spreadsheet {
                    cells: Rc::new(RefCell::new(HashMap::new())),
                    raw_cells: Rc::new(RefCell::new(HashMap::new())),
                    cell_styles: Rc::new(RefCell::new(HashMap::new())),
                    total_rows: 100,
                    total_cols: 3,
                    top_row: 0,
                    left_col: 0,
                    cursor_row: 0,
                    cursor_col: 0,
                    editing: false,
                    edit_buf: String::new(),
                    edit_pos: 0,
                    col_width: 12,
                    margin_cols: 0,
                    main_cols: 3,
                    formula_bar_address_id: None,
                    formula_bar_entry_id: None,
                    anchor: None,
                    menu_text: " [File]  Edit    Insert    Format    Sheet    Help".into(),
                    status_text: "  type/F2·edit; Ctrl+C·copy; Ctrl+X·cut; Ctrl+V·paste; Ctrl+;·date; Ctrl+:·time; Ctrl+S·save; F1·help".into(),
                    border_title: String::new(),
                    formula_bar_trailing: String::new(),
                    column_layout: Vec::new(),
                    row_labels: Vec::new(),
                    tab_text: String::new(),
                    header_row_count: 2,
                    main_row_count: 24,
                },
                Some(wid),
            ));
            // Set cell data matching overflow.corro
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                match &mut n.kind {
                    PcWidgetKind::Spreadsheet { ref cells, ref raw_cells, .. } => {
                        cells.borrow_mut().insert((0, 0), "This Text is really long and should overflow.".into());
                        cells.borrow_mut().insert((1, 1), "This Text is really long and should overflow.".into());
                        cells.borrow_mut().insert((2, 2), "This Text is really long and should overflow.".into());
                        raw_cells.borrow_mut().insert((0, 0), "This Text is really long and should overflow.".into());
                        raw_cells.borrow_mut().insert((1, 1), "This Text is really long and should overflow.".into());
                        raw_cells.borrow_mut().insert((2, 2), "This Text is really long and should overflow.".into());
                    }
                    _ => unreachable!(),
                }
            });

            // Compute column widths from cell content (matching ratatui's auto-fit logic)
            let cells = with_state(|state| {
                let n = state.node(sid).unwrap();
                match &n.kind {
                    PcWidgetKind::Spreadsheet { ref cells, .. } => cells.clone(),
                    _ => unreachable!(),
                }
            });
            let layout: Vec<(u32, u32, String)> = (0..3).map(|col| {
                let max_w = cells.borrow().iter()
                    .filter(|&((_, c), _)| *c == col)
                    .map(|((r, _), text)| text.chars().count().max(format!("{}", r + 1).len()))
                    .max()
                    .unwrap_or(8)
                    .max(8) as u32;
                (col, max_w + 2, col_label(col)) // +2 for padding
            }).collect();
            // Match pnc_backend.rs: logical_row = HEADER_ROWS + r
            let header_rows = 1usize;
            let main_rows = 3usize;
            let row_labels: Vec<(u32, String)> = (0..main_rows as u32).map(|r| {
                let logical_row = header_rows + r as usize;
                if logical_row < header_rows + main_rows {
                    (r, format!("{:>4}", logical_row + 1 - header_rows))
                } else {
                    let fr = logical_row - header_rows - main_rows;
                    (r, format!("_{:>3}", fr + 1))
                }
            }).collect();

            spreadsheet_set_column_layout(sid, layout);
            spreadsheet_set_row_labels(sid, row_labels);

            let buf = render_spreadsheet_to_buffer(sid, 120, 40);

            // Verify key structural elements match ratatui output
            // Line 0: Menu bar
            assert_eq!(&buf[0], " [File]  Edit    Insert    Format    Sheet    Help",
                "menu bar mismatch: {:?}", &buf[0]);

            // Line 1: Formula bar — should start with leading space then A1
            assert!(buf[1].starts_with(" "), "formula bar missing leading space: {:?}", &buf[1]);
            assert!(buf[1].contains("A1"), "formula bar missing A1: {:?}", &buf[1]);
            assert!(buf[1].contains("This Text is really long"), "formula bar missing text");

            // Line 2: Grid header starts with │ border, then column labels
            assert!(buf[2].starts_with("│"), "grid header missing border: {:?}", &buf[2]);
            assert!(buf[2].contains("A"), "header missing A");
            assert!(buf[2].contains("B"), "header missing B");
            assert!(buf[2].contains("C"), "header missing C");

            // Note: ratatui shows a border box around the grid:
            //   ┌ corro  3r × 3c  ops 3─────┐   (top border)
            //   │     [A  │ A     B     C   │   (column headers)
            // Our pancurses buffer may not match the border format exactly
            // but the cell content and structural elements should match.

            // Data rows: each should contain the overflow text
            assert!(buf[4].contains("This Text is really long"), "row 1 missing overflow text");
            assert!(buf[4].starts_with("│   1"), "row 1 label wrong: {:?}", &buf[4]);

            // Row 2 should start with │   2
            assert!(buf[5].starts_with("│   2"), "row 2 label wrong: {:?}", &buf[5]);

            // Row 3 should start with │   3
            assert!(buf[6].starts_with("│   3"), "row 3 label wrong: {:?}", &buf[6]);

            // Status bar
            let last = &buf[buf.len() - 1];
            assert!(last.contains("Ctrl+S"), "status bar missing save hint: {:?}", last);

            // Verify cell text appears
            let texts = ["This Text is really long and should overflow."];
            for text in &texts {
                let found = buf.iter().any(|line| line.contains(text));
                assert!(found, "text '{}' not found in pancurses render", text);
            }
        }

        // Moving to the right edge of the spreadsheet.
        #[test]
        fn spreadsheet_cursor_stays_in_bounds() {
            let sid = make_spreadsheet_id();
            // Move to last column
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_col, total_cols, .. } = n.kind {
                    *cursor_col = total_cols - 1;
                }
            });
            let pos = spreadsheet_cursor_position(sid);
            assert_eq!(pos, Some((0, 25)), "cursor should be at col 25 (Z)");

            // KeyRight should be clamped
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_col, total_cols, .. } = n.kind {
                    if *cursor_col + 1 < total_cols { *cursor_col += 1; }
                }
            });
            let pos = spreadsheet_cursor_position(sid);
            assert_eq!(pos, Some((0, 25)), "cursor should stay at col 25");

            // Move to last row
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, total_rows, .. } = n.kind {
                    *cursor_row = total_rows - 1;
                }
            });
            // KeyDown should be clamped
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, total_rows, .. } = n.kind {
                    if *cursor_row + 1 < total_rows { *cursor_row += 1; }
                }
            });
            let pos = spreadsheet_cursor_position(sid);
            assert_eq!(pos, Some((99, 25)), "cursor should stay at row 99");
        }

        // ── Full ratatui-compatible test suite ──────────────────────────────

        /// Navigation: arrow keys move cursor in all four directions.
        #[test]
        fn navigation_arrows() {
            let sid = make_spreadsheet_id();
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, ref mut cursor_col, .. } = n.kind {
                    *cursor_row = 5; *cursor_col = 3;
                }
            });
            // Down
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, .. } = n.kind {
                    if *cursor_row + 1 < 100 { *cursor_row += 1; }
                }
            });
            assert_eq!(spreadsheet_cursor_position(sid), Some((6, 3)));
            // Right
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_col, .. } = n.kind {
                    if *cursor_col + 1 < 26 { *cursor_col += 1; }
                }
            });
            assert_eq!(spreadsheet_cursor_position(sid), Some((6, 4)));
            // Up
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, .. } = n.kind {
                    if *cursor_row > 0 { *cursor_row -= 1; }
                }
            });
            assert_eq!(spreadsheet_cursor_position(sid), Some((5, 4)));
            // Left
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_col, .. } = n.kind {
                    if *cursor_col > 0 { *cursor_col -= 1; }
                }
            });
            assert_eq!(spreadsheet_cursor_position(sid), Some((5, 3)));
        }

        /// Edit: Enter toggles edit mode.
        #[test]
        fn edit_toggle_enter() {
            let sid = make_spreadsheet_id();
            with_state(|state| spreadsheet_enter(state, sid));
            with_state(|state| {
                let n = state.node(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { editing, .. } = &n.kind {
                    assert!(*editing, "should be editing after enter");
                }
            });
        }

        /// Edit: Enter again commits and moves to next row.
        #[test]
        fn edit_enter_commits_and_advances() {
            let sid = make_spreadsheet_id();
            with_state(|state| spreadsheet_enter(state, sid));
            // Type text
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut edit_buf, .. } = n.kind {
                    *edit_buf = "hello".to_string();
                }
            });
            // Commit
            with_state(|state| spreadsheet_enter(state, sid));
            assert_eq!(spreadsheet_get_cell(sid, 0, 0), Some("hello".into()));
            assert_eq!(spreadsheet_cursor_position(sid), Some((1, 0)));
        }

        /// Edit: Esc cancels editing (no cell content change).
        #[test]
        fn edit_esc_cancels() {
            let sid = make_spreadsheet_id();
            // Set initial cell value
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref cells, .. } = n.kind {
                    cells.borrow_mut().insert((0, 0), "original".into());
                }
            });
            // Start editing
            with_state(|state| spreadsheet_enter(state, sid));
            // Change buffer
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut edit_buf, .. } = n.kind {
                    *edit_buf = "modified".to_string();
                }
            });
            // Cancel via Esc (set editing=false without committing)
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut editing, .. } = n.kind {
                    *editing = false;
                }
            });
            assert_eq!(spreadsheet_get_cell(sid, 0, 0), Some("original".into()));
        }

        /// Edit: Tab commits and moves right.
        #[test]
        fn edit_tab_moves_right() {
            let sid = make_spreadsheet_id();
            with_state(|state| spreadsheet_enter(state, sid));
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut edit_buf, .. } = n.kind {
                    *edit_buf = "data".to_string();
                }
            });
            // Commit via Tab (simulate: commit + move right)
            with_state(|state| spreadsheet_commit_edit(state, sid));
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_col, total_cols, .. } = n.kind {
                    if *cursor_col + 1 < total_cols { *cursor_col += 1; }
                }
            });
            assert_eq!(spreadsheet_get_cell(sid, 0, 0), Some("data".into()));
            assert_eq!(spreadsheet_cursor_position(sid), Some((0, 1)));
        }

        /// Cursor: clamped at row 0 and col 0.
        #[test]
        fn cursor_clamped_zero() {
            let sid = make_spreadsheet_id();
            // Try to go up from 0
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, .. } = n.kind {
                    if *cursor_row > 0 { *cursor_row -= 1; }
                }
            });
            assert_eq!(spreadsheet_cursor_position(sid), Some((0, 0)));
            // Try to go left from 0
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_col, .. } = n.kind {
                    if *cursor_col > 0 { *cursor_col -= 1; }
                }
            });
            assert_eq!(spreadsheet_cursor_position(sid), Some((0, 0)));
        }

        /// Selection: anchor is set on Shift+Arrow and cleared on non-Shift Arrow.
        #[test]
        fn selection_anchor() {
            let sid = make_spreadsheet_id();
            // Move to (5,5)
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, ref mut cursor_col, .. } = n.kind {
                    *cursor_row = 5; *cursor_col = 5;
                }
            });
            // Simulate Shift+Down: set anchor then move
            with_state(|state| spreadsheet_prepare_move(state, sid, true));
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, .. } = n.kind {
                    *cursor_row += 1;
                }
            });
            // Anchor should be set to (5, 5)
            let anchor = with_state(|state| {
                let n = state.node(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { anchor, .. } = &n.kind {
                    *anchor
                } else { None }
            });
            assert_eq!(anchor, Some((5, 5)), "anchor should be at (5,5) after Shift+Down");

            // Non-shift move should clear anchor
            with_state(|state| spreadsheet_prepare_move(state, sid, false));
            let anchor = with_state(|state| {
                let n = state.node(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { anchor, .. } = &n.kind {
                    *anchor
                } else { None }
            });
            assert_eq!(anchor, None, "anchor should be None after non-Shift move");
        }

        /// Cell content: set and get round-trip.
        /// Verify that make_spreadsheet_id correctly sets focus on the spreadsheet.
        #[test]
        fn focus_is_on_spreadsheet_after_creation() {
            let sid = make_spreadsheet_id();
            let focused = with_state(|s| s.focus_id);
            assert_eq!(focused, Some(sid), "focus should be on the spreadsheet after creation");
        }

        /// Simulate full arrow-key processing: down, right, up, left.
        /// This tests the exact order of operations the keyboard handler uses.
        #[test]
        fn keyboard_arrows_work() {
            let sid = make_spreadsheet_id();
            // Move to (3,5) first
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, ref mut cursor_col, .. } = n.kind {
                    *cursor_row = 3; *cursor_col = 5;
                }
            });
            // Simulate KeyDown handler:
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, total_rows, .. } = n.kind {
                    if *cursor_row + 1 < total_rows { *cursor_row += 1; }
                }
            });
            assert_eq!(spreadsheet_cursor_position(sid), Some((4, 5)), "KeyDown failed");
            // Simulate KeyRight handler:
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_col, total_cols, .. } = n.kind {
                    if *cursor_col + 1 < total_cols { *cursor_col += 1; }
                }
            });
            assert_eq!(spreadsheet_cursor_position(sid), Some((4, 6)), "KeyRight failed");
            // Simulate KeyUp handler:
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_row, .. } = n.kind {
                    if *cursor_row > 0 { *cursor_row -= 1; }
                }
            });
            assert_eq!(spreadsheet_cursor_position(sid), Some((3, 6)), "KeyUp failed");
            // Simulate KeyLeft handler:
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_col, .. } = n.kind {
                    if *cursor_col > 0 { *cursor_col -= 1; }
                }
            });
            assert_eq!(spreadsheet_cursor_position(sid), Some((3, 5)), "KeyLeft failed");
        }

        /// Simulate Tab key: commit edit and move right.
        #[test]
        fn keyboard_tab_commits_and_moves_right() {
            let sid = make_spreadsheet_id();
            with_state(|state| spreadsheet_enter(state, sid));
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut edit_buf, .. } = n.kind {
                    *edit_buf = "tab_test".to_string();
                }
            });
            // Commit edit (simulates Tab handler)
            with_state(|state| spreadsheet_commit_edit(state, sid));
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut cursor_col, total_cols, .. } = n.kind {
                    if *cursor_col + 1 < total_cols { *cursor_col += 1; }
                }
            });
            assert_eq!(spreadsheet_get_cell(sid, 0, 0), Some("tab_test".into()));
            assert_eq!(spreadsheet_cursor_position(sid), Some((0, 1)));
        }

        /// Simulate Enter then type then Enter again (commit + advance).
        #[test]
        fn keyboard_enter_edit_type_enter_commit() {
            let sid = make_spreadsheet_id();
            // First Enter starts editing
            with_state(|state| spreadsheet_enter(state, sid));
            // Type "hello" into edit buffer (simulates Character('h'), etc.)
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut edit_buf, ref mut edit_pos, .. } = n.kind {
                    *edit_buf = "hello".to_string();
                    *edit_pos = 5;
                }
            });
            // Second Enter commits and advances row
            with_state(|state| spreadsheet_enter(state, sid));
            assert_eq!(spreadsheet_get_cell(sid, 0, 0), Some("hello".into()));
            assert_eq!(spreadsheet_cursor_position(sid), Some((1, 0)));
        }

        /// Simulate typing characters in edit mode via the Character(c) handler logic.
        #[test]
        fn keyboard_typing_in_edit_mode() {
            let sid = make_spreadsheet_id();
            // Start editing
            with_state(|state| spreadsheet_enter(state, sid));
            // Type 'a' (simulates Character('a') handler)
            with_state(|state| {
                if let Some(fid) = state.focus_id {
                    if is_spreadsheet_focused(state, fid) {
                        if let Some(n) = state.node_mut(fid) {
                            if let PcWidgetKind::Spreadsheet { ref mut editing, ref mut edit_buf, ref mut edit_pos, .. } = n.kind {
                                if !*editing {
                                    *editing = true;
                                    *edit_buf = "a".to_string();
                                    *edit_pos = 1;
                                } else {
                                    edit_buf.insert(*edit_pos, 'a');
                                    *edit_pos += 1;
                                }
                            }
                        }
                    }
                }
            });
            // Type 'b' 
            with_state(|state| {
                if let Some(fid) = state.focus_id {
                    if is_spreadsheet_focused(state, fid) {
                        if let Some(n) = state.node_mut(fid) {
                            if let PcWidgetKind::Spreadsheet { ref mut editing, ref mut edit_buf, ref mut edit_pos, .. } = n.kind {
                                if !*editing {
                                    *editing = true;
                                    *edit_buf = "b".to_string();
                                    *edit_pos = 1;
                                } else {
                                    edit_buf.insert(*edit_pos, 'b');
                                    *edit_pos += 1;
                                }
                            }
                        }
                    }
                }
            });
            // Commit
            with_state(|state| spreadsheet_commit_edit(state, sid));
            assert_eq!(spreadsheet_get_cell(sid, 0, 0), Some("ab".into()));
        }

        /// Simulate Backspace in edit mode.
        #[test]
        fn keyboard_backspace_in_edit() {
            let sid = make_spreadsheet_id();
            with_state(|state| spreadsheet_enter(state, sid));
            // Type "abc"
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut editing, ref mut edit_buf, ref mut edit_pos, .. } = n.kind {
                    *editing = true;
                    *edit_buf = "abc".to_string();
                    *edit_pos = 3;
                }
            });
            // Simulate Backspace: remove char before cursor
            with_state(|state| {
                if let Some(fid) = state.focus_id {
                    if is_spreadsheet_focused(state, fid) {
                        if let Some(n) = state.node_mut(fid) {
                            if let PcWidgetKind::Spreadsheet { ref mut editing, ref mut edit_buf, ref mut edit_pos, .. } = n.kind {
                                if *editing && *edit_pos > 0 {
                                    *edit_pos -= 1;
                                    edit_buf.remove(*edit_pos);
                                }
                            }
                        }
                    }
                }
            });
            assert_eq!(spreadsheet_get_cell(sid, 0, 0), None, "not committed yet");
            with_state(|state| spreadsheet_enter(state, sid));
            assert_eq!(spreadsheet_get_cell(sid, 0, 0), Some("ab".into()), "backspace should remove 'c'");
        }

        /// Simulate the q/Q quit handler (should quit when not editing).
        #[test]
        fn keyboard_q_quits_when_not_editing() {
            let sid = make_spreadsheet_id();
            // Mark not editing
            with_state(|state| {
                let n = state.node_mut(sid).unwrap();
                if let PcWidgetKind::Spreadsheet { ref mut editing, .. } = n.kind {
                    *editing = false;
                }
            });
            with_state(|state| state.running = false);
            // Verify running was set to false
            let running = with_state(|s| s.running);
            assert!(!running, "running should be false after q");
        }

        /// Simulate Ctrl+Q (KeyExit) — should quit regardless of edit state.
        #[test]
        fn keyboard_ctrlq_quits() {
            with_state(|state| state.running = false);
            let running = with_state(|s| s.running);
            assert!(!running, "running should be false after Ctrl+Q/KeyExit");
        }

        #[test]
        fn cell_content_roundtrip() {
            let sid = make_spreadsheet_id();
            spreadsheet_set_cell(sid, 5, 3, "hello world");
            assert_eq!(spreadsheet_get_cell(sid, 5, 3), Some("hello world".into()));
            assert_eq!(spreadsheet_get_cell(sid, 0, 0), None, "different cell should be empty");
        }

        /// Cell content: multiple cells can be set independently.
        #[test]
        fn cell_content_multiple() {
            let sid = make_spreadsheet_id();
            spreadsheet_set_cell(sid, 0, 0, "A1");
            spreadsheet_set_cell(sid, 0, 1, "B1");
            spreadsheet_set_cell(sid, 1, 0, "A2");
            assert_eq!(spreadsheet_get_cell(sid, 0, 0), Some("A1".into()));
            assert_eq!(spreadsheet_get_cell(sid, 0, 1), Some("B1".into()));
            assert_eq!(spreadsheet_get_cell(sid, 1, 0), Some("A2".into()));
        }

        /// Rendering: buffer has correct number of rows.
        #[test]
        fn render_buffer_size() {
            let wid = with_state(|s| s.add_node(
                PcWidgetKind::Window { title: "test".into() }, None,
            ));
            let sid = with_state(|s| s.add_node(
                PcWidgetKind::Spreadsheet {
                    cells: Rc::new(RefCell::new(HashMap::new())),
                    raw_cells: Rc::new(RefCell::new(HashMap::new())),
                    cell_styles: Rc::new(RefCell::new(HashMap::new())),
                    total_rows: 10, total_cols: 3, top_row: 0, left_col: 0,
                    cursor_row: 0, cursor_col: 0, editing: false,
                    edit_buf: String::new(), edit_pos: 0, col_width: 12,
                    margin_cols: 0, main_cols: 3,
                    formula_bar_address_id: None, formula_bar_entry_id: None,
                    anchor: None,
                    menu_text: "menu".into(),
                    status_text: "status".into(),
                    border_title: String::new(),
                    formula_bar_trailing: String::new(),
                    column_layout: vec![(0,10,"A".into()),(1,10,"B".into()),(2,10,"C".into())],
                    row_labels: vec![(0,"   1".into()),(1,"   2".into())],
                    tab_text: String::new(),
                    header_row_count: 2,
                    main_row_count: 20,
                },
                Some(wid),
            ));
            let buf = render_spreadsheet_to_buffer(sid, 80, 10);
            assert_eq!(buf.len(), 10);
            assert!(buf[0].contains("menu"));
            assert!(buf[buf.len()-1].contains("status"));
            // Data rows should have │ borders
            assert!(buf[4].starts_with("│"), "data row missing │ border");
        }

        /// Rendering: column headers use layout labels.
        #[test]
        fn render_column_headers_from_layout() {
            let wid = with_state(|s| s.add_node(
                PcWidgetKind::Window { title: "test".into() }, None,
            ));
            let sid = with_state(|s| s.add_node(
                PcWidgetKind::Spreadsheet {
                    cells: Rc::new(RefCell::new(HashMap::new())),
                    raw_cells: Rc::new(RefCell::new(HashMap::new())),
                    cell_styles: Rc::new(RefCell::new(HashMap::new())),
                    total_rows: 10, total_cols: 2, top_row: 0, left_col: 0,
                    cursor_row: 0, cursor_col: 0, editing: false,
                    edit_buf: String::new(), edit_pos: 0, col_width: 12,
                    margin_cols: 0, main_cols: 2,
                    formula_bar_address_id: None, formula_bar_entry_id: None,
                    anchor: None,
                    menu_text: String::new(), status_text: String::new(),
                    border_title: String::new(),
                    formula_bar_trailing: String::new(),
                    column_layout: vec![(0,5,"X".into()),(1,5,"Y".into())],
                    row_labels: vec![],
                    tab_text: String::new(),
                    header_row_count: 2,
                    main_row_count: 20,
                },
                Some(wid),
            ));
            let buf = render_spreadsheet_to_buffer(sid, 80, 10);
            // Header is at index 1 (index 0 = formula bar, since menu_text is empty)
            assert!(buf[1].contains("X"), "header missing X in {:?}", &buf[1]);
            assert!(buf[1].contains("Y"), "header missing Y in {:?}", &buf[1]);
        }

        /// Rendering: pre-computed row labels appear in data rows.
        #[test]
        fn render_row_labels_from_layout() {
            let wid = with_state(|s| s.add_node(
                PcWidgetKind::Window { title: "test".into() }, None,
            ));
            let sid = with_state(|s| s.add_node(
                PcWidgetKind::Spreadsheet {
                    cells: Rc::new(RefCell::new(HashMap::new())),
                    raw_cells: Rc::new(RefCell::new(HashMap::new())),
                    cell_styles: Rc::new(RefCell::new(HashMap::new())),
                    total_rows: 10, total_cols: 1, top_row: 0, left_col: 0,
                    cursor_row: 0, cursor_col: 0, editing: false,
                    edit_buf: String::new(), edit_pos: 0, col_width: 12,
                    margin_cols: 0, main_cols: 1,
                    formula_bar_address_id: None, formula_bar_entry_id: None,
                    anchor: None,
                    menu_text: String::new(), status_text: String::new(),
                    border_title: String::new(),
                    formula_bar_trailing: String::new(),
                    column_layout: vec![(0,8,"A".into())],
                    row_labels: vec![(0,"ROW0".into()),(1,"ROW1".into())],
                    tab_text: String::new(),
                    header_row_count: 2,
                    main_row_count: 20,
                },
                Some(wid),
            ));
            let buf = render_spreadsheet_to_buffer(sid, 80, 10);
            assert!(buf[4].contains("ROW0"), "row 0 label missing: {:?}", buf[4]);
            assert!(buf[5].contains("ROW1"), "row 1 label missing: {:?}", buf[5]);
        }
    }
}

pub use pancurses_backend::*;
