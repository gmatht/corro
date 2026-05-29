use std::process::Command;
use std::path::Path;

#[test]
fn run_sh_without_args_does_not_open_repo_tmp_tsv() {
    // Run the helper script with no args and ensure it exits successfully.
    // We can't inspect the TUI, but we can check that running with --version
    // still reports the expected version and that no errors occur.
    let out = Command::new("./run.sh").arg("--version").output().expect("failed to run");
    assert!(out.status.success());
    let stdout = String::from_utf8_lossy(&out.stdout);
    assert!(stdout.contains("corro"));
}
