//! Mixed workload for LLVM PGO: ~20% wall time replaying `.corro` logs (default: `docs/` under the crate),
//! ~80% arrow navigation plus ratatui draw on a `TestBackend`.

use corro::grid::{HEADER_ROWS, MARGIN_COLS, SheetCursor};
use corro::ui::App;
use crossterm::event::{KeyCode, KeyEvent, KeyModifiers};
use ratatui::backend::TestBackend;
use ratatui::Terminal;
use std::fs;
use std::hint::black_box;
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant};

const DEFAULT_SECS: u64 = 90;
/// Target fraction of cumulative **CPU time** (replay phases vs arrow phases): ~20% replay.
const REPLAY_NUM: u128 = 1;
const REPLAY_DENOM: u128 = 5;
/// Key+draw batches per iteration when arrows are favored (replay work is comparatively cheap).
const ARROWS_PER_BATCH: usize = 8;
const TERMINAL_W: u16 = 120;
const TERMINAL_H: u16 = 28;

fn gather_corro_under(dir: &Path, out: &mut Vec<PathBuf>) -> std::io::Result<()> {
    if !dir.is_dir() {
        return Ok(());
    }
    for entry in fs::read_dir(dir)? {
        let p = entry?.path();
        if p.is_dir() {
            gather_corro_under(&p, out)?;
        } else if p
            .extension()
            .is_some_and(|e| e.eq_ignore_ascii_case("corro"))
        {
            out.push(p);
        }
    }
    Ok(())
}

fn load_log_corpus(paths: &[PathBuf]) -> std::io::Result<Vec<Vec<String>>> {
    let mut files = Vec::with_capacity(paths.len());
    for path in paths {
        let text = fs::read_to_string(path)?;
        let mut lines = Vec::new();
        for line in text.lines() {
            let t = line.trim();
            if !t.is_empty() {
                lines.push(t.to_string());
            }
        }
        if !lines.is_empty() {
            files.push(lines);
        }
    }
    Ok(files)
}

struct CorpusCursor {
    files: Vec<Vec<String>>,
    file_idx: usize,
    line_idx: usize,
}

impl CorpusCursor {
    fn new(files: Vec<Vec<String>>) -> Self {
        Self {
            files,
            file_idx: 0,
            line_idx: 0,
        }
    }

    fn next_line(&mut self) -> Option<String> {
        if self.files.is_empty() {
            return None;
        }
        let line = self.files[self.file_idx][self.line_idx].clone();
        self.line_idx += 1;
        if self.line_idx >= self.files[self.file_idx].len() {
            self.line_idx = 0;
            self.file_idx = (self.file_idx + 1) % self.files.len();
        }
        Some(line)
    }
}

fn parse_cli() -> (u64, PathBuf, bool) {
    let mut secs = DEFAULT_SECS;
    let manifest = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let mut scan_root = manifest.join("docs");
    let mut quiet = false;
    let mut args = std::env::args().skip(1);
    while let Some(a) = args.next() {
        match a.as_str() {
            "--duration" => {
                if let Some(n) = args.next().and_then(|s| s.parse().ok()) {
                    secs = n;
                }
            }
            "--docs-dir" => {
                if let Some(p) = args.next() {
                    scan_root = PathBuf::from(p);
                }
            }
            "--quiet" => quiet = true,
            _ => {}
        }
    }
    (secs, scan_root, quiet)
}

fn arrow_key_pattern() -> [KeyEvent; 4] {
    [
        KeyEvent::new(KeyCode::Down, KeyModifiers::empty()),
        KeyEvent::new(KeyCode::Right, KeyModifiers::empty()),
        KeyEvent::new(KeyCode::Up, KeyModifiers::empty()),
        KeyEvent::new(KeyCode::Left, KeyModifiers::empty()),
    ]
}

fn run_mix(duration: Duration, scan_root: PathBuf, quiet: bool) -> std::io::Result<()> {
    let mut paths = Vec::new();
    gather_corro_under(&scan_root, &mut paths)?;
    paths.sort();

    let corpus_files = load_log_corpus(&paths)?;
    let mut corpus = CorpusCursor::new(corpus_files);

    if corpus.files.is_empty() {
        eprintln!(
            "pgo_mix_benchmark: no .corro files under {}",
            scan_root.display()
        );
        std::process::exit(2);
    }

    let mut app = App::new(None);
    app.load_initial()
        .map_err(|e| std::io::Error::new(std::io::ErrorKind::Other, e.to_string()))?;
    app.state.grid.set_main_size(64, 48);
    app.cursor = SheetCursor {
        row: HEADER_ROWS + 12,
        col: MARGIN_COLS + 6,
    };

    let backend = TestBackend::new(TERMINAL_W, TERMINAL_H);
    let mut terminal = Terminal::new(backend).map_err(|e| {
        std::io::Error::new(std::io::ErrorKind::Other, format!("terminal: {}", e))
    })?;

    let pattern = arrow_key_pattern();
    let deadline = Instant::now() + duration;
    let mut replay_wall = Duration::ZERO;
    let mut arrow_wall = Duration::ZERO;
    let mut replay_iters: u64 = 0;
    let mut arrow_iters: u64 = 0;

    let wall_start = Instant::now();

    /// Replay until `replay_cpu / wall_clock` reaches `REPLAY_NUM/REPLAY_DENOM`, then arrows.
    ///
    /// `wall_clock` is wall time since benchmark start (`Instant::elapsed`); replay and arrow timers
    /// only count phases of this workload, keeping ~20%/80% time share when sustained.
    #[inline]
    fn prefer_replay(replay_cpu: Duration, wall_clock: Duration) -> bool {
        let w = wall_clock.as_nanos() as u128;
        w == 0
            || (replay_cpu.as_nanos() as u128).saturating_mul(REPLAY_DENOM)
                < w.saturating_mul(REPLAY_NUM)
    }

    while Instant::now() < deadline {
        let wall_clock = wall_start.elapsed();
        let want_replay = prefer_replay(replay_wall, wall_clock);

        if want_replay {
            let t0 = Instant::now();
            if let Some(line) = corpus.next_line() {
                let _ = app.bench_apply_corro_log_line(&line);
                replay_iters = replay_iters.saturating_add(1);
            }
            replay_wall += t0.elapsed();
        } else {
            let t0 = Instant::now();
            for i in 0..ARROWS_PER_BATCH {
                let k = pattern[(arrow_iters as usize + i) % pattern.len()];
                let _ = black_box(app.bench_handle_key(k));
                terminal.draw(|f| app.bench_draw(f)).map_err(|e| {
                    std::io::Error::new(std::io::ErrorKind::Other, format!("draw: {}", e))
                })?;
            }
            arrow_iters += ARROWS_PER_BATCH as u64;
            arrow_wall += t0.elapsed();
        }
    }

    let wall = wall_start.elapsed();
    let total_phase = replay_wall + arrow_wall;
    let replay_pct = if total_phase.is_zero() {
        0.0
    } else {
        100.0 * replay_wall.as_secs_f64() / total_phase.as_secs_f64()
    };

    let wall_ms = wall.as_secs_f64() * 1000.0;
    let replay_ms = replay_wall.as_secs_f64() * 1000.0;
    let arrow_ms = arrow_wall.as_secs_f64() * 1000.0;

    if quiet {
        println!(
            "PGO_MIX wall_ms={:.3} replay_ms={:.3} arrow_ms={:.3} replay_wall_pct={:.2} replay_lines={} arrow_key_draw_iters={}",
            wall_ms,
            replay_ms,
            arrow_ms,
            replay_pct,
            replay_iters,
            arrow_iters
        );
    } else {
        eprintln!(
            "pgo_mix_benchmark: wall {:.3} ms  replay {:.3} ms ({:.1}% of replay+arrow)  arrow {:.3} ms\n\
             replay log lines: {}  arrow key×draw iterations: {}  (.corro files: {})  scan: {}",
            wall_ms,
            replay_ms,
            replay_pct,
            arrow_ms,
            replay_iters,
            arrow_iters,
            paths.len(),
            scan_root.display()
        );
    }
    Ok(())
}

fn main() -> std::io::Result<()> {
    let (secs, scan_root, quiet) = parse_cli();
    let duration = Duration::from_secs(secs.max(1));
    run_mix(duration, scan_root, quiet)
}
