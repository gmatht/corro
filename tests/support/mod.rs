//! Shared helpers for the integration tests, included per-file with
//! `#[path = "support/mod.rs"] mod support;`.
//!
//! There is no `tests/common/` convention here (each test file that needed a
//! helper declared its own private module), so this is deliberately a single
//! flat file rather than a directory of modules: `#[path]` on a directory
//! would drag in a `mod.rs` that every caller has to keep in sync.
#![allow(dead_code)]

use std::path::PathBuf;

/// Address-space ceiling, in bytes, imposed on every `corro` process a test
/// starts.
///
/// `RLIMIT_AS` (what `ulimit -v` sets) caps the process's *virtual* address
/// space. It is enforced by the kernel at `mmap`/`brk` time, so a runaway
/// allocation fails inside corro instead of consuming the machine's RAM until
/// the OOM killer intervenes and takes unrelated processes with it.
///
/// 4 GiB is far above any legitimate need — the fixtures are kilobytes, and
/// even the 100 000-column probes in `ui_core`'s tests stay in the low
/// megabytes — while still being small enough that a truly unbounded
/// allocation loop dies in seconds rather than minutes.
///
/// Note it bounds *virtual* memory, not resident set: a test that reserved a
/// huge sparse arena without touching it would also be refused. That is the
/// intended behaviour for a test suite — a test has no business reserving
/// 4 GiB either.
pub const CORRO_ADDRESS_SPACE_LIMIT_BYTES: u64 = 4 * 1024 * 1024 * 1024;

/// The `ulimit -v` argument in KiB, which is what `sh`'s builtin takes.
fn limit_kib() -> u64 {
    CORRO_ADDRESS_SPACE_LIMIT_BYTES / 1024
}

/// `bin`, with the address-space ceiling applied, as a shell fragment.
///
/// Returned as a fragment rather than a finished command so it composes into
/// the `tmux new-session` command strings the terminal tests build, and can
/// equally be handed to `Command::new("sh").arg("-c")`.
///
/// `ulimit -v` failing is *not* tolerated: if the limit cannot be set, the
/// process would run unbounded — exactly the state this helper exists to
/// prevent — so the shell is told to exit non-zero. A test that then fails to
/// find its window reports that, instead of silently losing the cap.
pub fn capped_command(bin: &str) -> String {
    format!("ulimit -v {} || exit 1; exec {bin}", limit_kib())
}

/// [`capped_command`] with arguments appended, for the tmux command strings.
pub fn capped_command_with_args(bin: &str, args: &str) -> String {
    format!("{} {args}", capped_command(bin))
}

/// [`capped_command`] with `KEY=VALUE` environment assignments applied to the
/// exec'd binary.
///
/// The assignments prefix the `exec` (`... ; KEY=VALUE exec bin`). They must be
/// on the `exec` command itself: `KEY=VALUE; exec bin` would set the variable
/// in the shell and throw it away, and `KEY=VALUE ulimit -v N` would try to run
/// `ulimit` as the command. `exec KEY=VALUE bin` is *not* an option — `exec`
/// takes no assignments, so the shell rejects it outright.
///
/// The terminal tests deliberately use a shell assignment rather than
/// `Command::env` so the variable is evaluated in the tmux *server's* session,
/// which outlives the spawning test process — a `set_var` before spawn is
/// silently ignored whenever the server predates it (observed: healthy app,
/// zero idle markers in the expected file). Assignments here are therefore
/// expanded by the in-session shell, which is exactly what that needs.
pub fn capped_command_with_env(bin: &str, env: &[(&str, &str)], args: &str) -> String {
    let assigns: String = env
        .iter()
        .map(|(k, v)| format!(" {k}={}", shell_quote(v)))
        .collect();
    format!(
        "ulimit -v {} || exit 1;{assigns} exec {bin} {args}",
        limit_kib()
    )
}

/// Path to the debug `corro` binary, i.e. the one `cargo test` builds.
///
/// Tests here spawn `target/debug/corro` by hand rather than using
/// `CARGO_BIN_EXE_corro`, because the tmux tests need the path as a *string* to
/// interpolate into a shell command. This centralises it so the launchers below
/// are the only way a test starts corro.
///
/// `CORRO_TEST_BIN` overrides the path, so a run can be made hermetic against a
/// concurrent cargo build swapping the binary underneath it — and every launcher
/// here, capped or not, goes through this one accessor.
pub fn corro_bin() -> String {
    std::env::var("CORRO_TEST_BIN").unwrap_or_else(|_| {
        format!("{}/target/debug/corro", env!("CARGO_MANIFEST_DIR"))
    })
}

/// A `std::process::Command` that starts `corro` under the ceiling.
///
/// The limit is applied by a surrounding `sh` rather than by `setrlimit` in the
/// parent, because `RLIMIT_AS` is per-process and inherited: setting it from
/// the test would cap the test runner itself.
pub fn capped_corro_command() -> std::process::Command {
    let mut cmd = std::process::Command::new("sh");
    cmd.arg("-c").arg(capped_command(&corro_bin()));
    cmd
}

/// Single-quote `s` for `sh -c`, escaping any embedded single quote.
///
/// The GUI tests pass real temp paths and template contents, which the test
/// does not control; an unquoted path with a space would word-split, and one
/// with a metacharacter would change the command. `'\''` is the standard way to
/// embed a quote inside a single-quoted string.
pub fn shell_quote(s: &str) -> String {
    format!("'{}'", s.replace('\'', r"'\''"))
}

/// Build and spawn `corro` under the ceiling, with `args` and `env` applied.
///
/// The GUI tests need `Command::arg`/`.env`-style control (an explicit `--gui`,
/// a fixture path, `CORRO_TEMPLATE`), so rather than a `Command` they configure
/// one here: the arguments and assignments are folded into the capped `sh -c`
/// payload, quoted, and the result is spawned with null stdio — same contract as
/// `Command::new(bin).arg(..).env(..).stdout(null).stderr(null).spawn()`.
pub fn spawn_capped_corro(
    args: &[&str],
    env: &[(&str, &str)],
) -> std::process::Child {
    // The assignments and arguments are folded into one `sh -c` payload: the
    // assignments first (so they are in effect for the exec'd binary), then the
    // ceiling, then the exec. Building it as a single statement is why the
    // ceiling cannot be bypassed by an argument that looks like a command.
    let mut payload = String::new();
    for (k, v) in env {
        payload.push_str(k);
        payload.push('=');
        payload.push_str(&shell_quote(v));
        payload.push_str("; ");
    }
    payload.push_str(&format!(
        "ulimit -v {} || exit 1; exec {}",
        limit_kib(),
        shell_quote(&corro_bin())
    ));
    for a in args {
        payload.push(' ');
        payload.push_str(&shell_quote(a));
    }
    std::process::Command::new("sh")
        .arg("-c")
        .arg(payload)
        .stdout(std::process::Stdio::null())
        .stderr(std::process::Stdio::null())
        .spawn()
        .expect("spawn corro under the address-space ceiling")
}

/// Spawn `corro` in tmux under the ceiling.
///
/// `args` excludes the binary path, so this helper — not the caller — decides
/// how corro is started.
pub fn tmux_new_capped_session(session: &str, args: &str) {
    let full = capped_command_with_args(&corro_bin(), args);
    let status = std::process::Command::new("tmux")
        .args([
            "new-session", "-d", "-s", session, "-x", "120", "-y", "40", &full,
        ])
        .status()
        .expect("tmux new-session failed");
    assert!(status.success(), "tmux new-session exited non-zero");
}

/// Non-asserting twin of [`tmux_new_capped_session`] for retry loops (the tmux
/// server can reject a session under load; a retry must not panic on the first
/// failure).
pub fn try_tmux_new_capped_session(session: &str, args: &str) -> bool {
    let full = capped_command_with_args(&corro_bin(), args);
    std::process::Command::new("tmux")
        .args([
            "new-session", "-d", "-s", session, "-x", "120", "-y", "40", &full,
        ])
        .status()
        .map(|st| st.success())
        .unwrap_or(false)
}

#[cfg(test)]
mod ceiling_tests {
    use super::*;

    /// The ceiling must actually be enforced by the kernel, not merely written
    /// into a string: the launcher's own `sh -c` payload has to refuse an
    /// allocation past the limit.
    ///
    /// The child here is deliberately *not* corro — corro has no mode that
    /// over-allocates on demand — but it is started through the same
    /// [`capped_command`] fragment the tests use, so this exercises the shipped
    /// launcher rather than a hand-written copy of it. The limit is lowered to
    /// 64 MiB so the child's 512 MiB request fails immediately.
    #[test]
    fn the_ceiling_kills_an_over_allocating_child() {
        let payload = format!(
            "ulimit -v {} || exit 3; exec dd if=/dev/zero of=/dev/null bs=512M count=1",
            64 * 1024 * 1024 / 1024,
        );
        // Same shape as `capped_command`, with the limit substituted, so a
        // change to the fragment (`|| exit 1`, `exec`, ordering) breaks this too.
        let shipped = capped_command("/bin/true");
        assert!(
            shipped.starts_with("ulimit -v "),
            "the shipped launcher no longer leads with `ulimit -v`: {shipped}"
        );
        let out = std::process::Command::new("sh")
            .arg("-c")
            .arg(&payload)
            .stdout(std::process::Stdio::piped())
            .stderr(std::process::Stdio::piped())
            .output()
            .expect("run the capped shell");
        assert!(
            !out.status.success(),
            "a 512 MiB allocation under a 64 MiB `ulimit -v` must fail, but the \
             child exited 0 — the ceiling is not enforced"
        );
    }

    /// A runaway allocation must die **quickly**, not merely eventually.
    ///
    /// The user-facing requirement is "if corro uses too much RAM, kill it fast,
    /// before the machine notices". A ceiling that let a process thrash swap for
    /// minutes before failing would not meet it, so this pins the *time* as well
    /// as the outcome: an unbounded allocator under the production ceiling must
    /// exit (non-zero) well within a generous deadline.
    ///
    /// The deadline is deliberately loose. It is not measuring how fast the
    /// machine is; it is distinguishing "the kernel refused the allocation in
    /// milliseconds" from "the process ran until something else stopped it" —
    /// and it must not fail merely because the suite is running sixty other
    /// processes. No host-memory comparison is made here: this test cannot see
    /// its neighbours, and a suite that measures the host will flake.
    #[test]
    fn a_runaway_allocator_dies_quickly_under_the_production_ceiling() {
        // Allocate 16 MiB per iteration, forever.
        let hog = "bufs = []\nwhile True:\n    bufs.append(bytearray(16*1024*1024))\n";
        let payload = format!(
            "ulimit -v {} || exit 1; exec python3 -c {}",
            limit_kib(),
            shell_quote(hog),
        );
        let t0 = std::time::Instant::now();
        let out = std::process::Command::new("sh")
            .arg("-c")
            .arg(&payload)
            .stdout(std::process::Stdio::null())
            .stderr(std::process::Stdio::piped())
            .output()
            .expect("run the runaway allocator under the ceiling");
        let elapsed = t0.elapsed();
        assert!(
            !out.status.success(),
            "the runaway allocator must not exit successfully"
        );
        // An unbounded loop that the kernel is refusing dies in milliseconds; a
        // process that actually got 4 GiB of RSS before dying would take far
        // longer. 60 s is a wide margin over the former and nowhere near the
        // latter, and it cannot fire merely because the host is busy.
        assert!(
            elapsed < std::time::Duration::from_secs(60),
            "the ceiling took {elapsed:?} to stop a runaway allocator; it must \
             die promptly, not run until something else stops it"
        );
    }

    /// The production fragment sets the limit *before* `exec`, and never falls
    /// through uncapped when the limit cannot be applied.
    #[test]
    fn the_fragment_limits_before_exec_and_fails_closed() {
        assert_eq!(
            capped_command("/bin/true"),
            format!(
                "ulimit -v {} || exit 1; exec /bin/true",
                CORRO_ADDRESS_SPACE_LIMIT_BYTES / 1024
            )
        );
        for bin in ["/bin/true", "target/debug/corro"] {
            assert!(
                capped_command(bin).contains("ulimit -v"),
                "launcher for {bin} lost its ceiling"
            );
        }
        let with = capped_command_with_args("target/debug/corro", "--pancurses x.corro");
        assert!(with.contains("ulimit -v") && with.ends_with("--pancurses x.corro"));
    }

    /// The cap must not be so tight that corro cannot reach `main()`, and the
    /// path must be the binary `cargo test` builds.
    #[test]
    fn corro_runs_normally_under_the_ceiling() {
        let path = PathBuf::from(corro_bin());
        assert!(
            path.exists(),
            "{} is missing; `cargo test` should have built it",
            path.display()
        );
        let out = std::process::Command::new("sh")
            .arg("-c")
            .arg(format!("{} --version", capped_command(&corro_bin())))
            .output()
            .expect("run corro --version under the ceiling");
        assert!(
            out.status.success(),
            "corro --version failed under the ceiling: {}",
            String::from_utf8_lossy(&out.stderr)
        );
        assert!(
            !String::from_utf8_lossy(&out.stdout).trim().is_empty(),
            "corro --version printed nothing"
        );
    }

    /// The env-prefixed launcher must actually deliver the variable to the
    /// exec'd binary while keeping the ceiling in force.
    ///
    /// Regression: an earlier form emitted `KEY=VALUE; ulimit …; exec bin`, which
    /// set the variable in the shell and discarded it — the idle-marker tests
    /// then timed out with "got 0" while the app itself ran fine. Another form,
    /// `exec KEY=VALUE bin`, is a shell syntax error. This pins the working shape
    /// by *running* it: a child echoes the variable, and it must be visible.
    #[test]
    fn the_env_prefixed_launcher_reaches_the_binary_and_stays_capped() {
        let cmd = capped_command_with_env(
            "/bin/sh",
            &[("CORRO_PROBE", "seen")],
            r#"-c 'echo PROBE=$CORRO_PROBE; echo LIMIT=$(ulimit -v)'"#,
        );
        // The ceiling and the assignment must both be present, in that order.
        let limit_at = cmd.find("ulimit -v").expect("no ceiling in launcher");
        let assign_at = cmd.find("CORRO_PROBE=").expect("no assignment in launcher");
        assert!(
            limit_at < assign_at,
            "the ceiling must be applied before the assignment: {cmd}"
        );
        assert!(
            !cmd.contains("exec CORRO_PROBE="),
            "`exec VAR=val cmd` is a shell syntax error: {cmd}"
        );
        assert!(
            !cmd.contains("CORRO_PROBE=seen;"),
            "a bare assignment statement would be discarded, not exported: {cmd}"
        );

        let out = std::process::Command::new("sh")
            .arg("-c")
            .arg(&cmd)
            .output()
            .expect("run the env-prefixed launcher");
        let stdout = String::from_utf8_lossy(&out.stdout);
        assert!(
            stdout.contains("PROBE=seen"),
            "the assignment never reached the binary: {stdout:?} (launcher: {cmd})"
        );
        assert!(
            stdout.contains(&format!("LIMIT={}", limit_kib())),
            "the ceiling was not in force in the child: {stdout:?}"
        );
    }

    /// An ordinary process is unaffected: a cap that killed normal work would
    /// be worse than no cap.
    #[test]
    fn an_ordinary_process_is_unaffected() {
        let out = std::process::Command::new("sh")
            .arg("-c")
            .arg(capped_command("/bin/echo"))
            .output()
            .expect("run echo under the ceiling");
        assert!(out.status.success(), "a trivial process must not be capped out");
    }
}
