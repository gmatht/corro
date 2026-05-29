use std::path::PathBuf;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let repo_root = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let path = repo_root.join("tmp.tsv");
    let mut app = corro::ui::App::new(Some(path));
    app.load_initial()?;

    // Print a minimal snapshot using only public fields so this binary can be
    // compiled as part of workspace build and tests.
    println!(
        "App cursor before: row={} col={}",
        app.cursor.row, app.cursor.col
    );

    Ok(())
}
