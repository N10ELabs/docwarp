use std::fs;
use std::io::Write;
use std::process::{Command, Output, Stdio};

use anyhow::{Context, Result, bail};
use tempfile::tempdir;

#[test]
fn guided_mode_converts_markdown_file_with_default_output() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("note.md");
    let output = temp.path().join("note.docx");
    fs::write(&input, "# Guided\n\nBody\n").context("failed writing markdown fixture")?;

    let run = run_guided_with_stdin(&format!("{}\n", input.display()))?;
    assert_command_status(&run, Some(0), "guided markdown conversion should succeed")?;
    let stdout = stdout_text(&run);

    assert!(
        output.is_file(),
        "expected guided output at {}",
        output.display()
    );
    assert_guided_startup_header(&stdout);
    assert!(
        !stdout.contains("Converting:"),
        "did not expect guided conversion preamble, got:\n{}",
        stdout
    );
    assert!(
        !stdout.contains("Output path"),
        "did not expect output-path prompt in guided mode, got:\n{}",
        stdout
    );
    assert!(
        stdout.contains("completed in "),
        "expected duration-based completion line, got:\n{}",
        stdout
    );

    Ok(())
}

#[test]
fn guided_mode_converts_docx_file_with_default_output() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let seed_md = temp.path().join("seed.md");
    let input_docx = temp.path().join("roundtrip.docx");
    let output_md = temp.path().join("roundtrip.md");

    fs::write(&seed_md, "# Source\n\nBody\n").context("failed writing seed markdown")?;
    let setup = run_docwarp([
        "md2docx",
        seed_md.to_string_lossy().as_ref(),
        "--output",
        input_docx.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&setup, Some(0), "setup md2docx should succeed")?;

    let run = run_guided_with_stdin(&format!("{}\n", input_docx.display()))?;
    assert_command_status(&run, Some(0), "guided docx conversion should succeed")?;
    assert!(
        output_md.is_file(),
        "expected guided output at {}",
        output_md.display()
    );

    let output = fs::read_to_string(&output_md).context("failed reading guided markdown output")?;
    assert!(
        output.contains("# Source"),
        "expected converted markdown heading, got:\n{output}"
    );

    Ok(())
}

#[test]
fn subcommand_output_includes_header_mode() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("note.md");
    let output = temp.path().join("note.docx");
    fs::write(&input, "# Header\n\nBody\n").context("failed writing markdown fixture")?;

    let run = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&run, Some(0), "md2docx should succeed")?;
    let stdout = stdout_text(&run);
    assert_startup_header(&stdout);
    assert!(
        !stdout.contains("Paste a path, or press Enter to choose."),
        "subcommand header should not include guided picker hint, got:\n{}",
        stdout
    );

    Ok(())
}

fn assert_startup_header(stdout: &str) {
    assert!(
        stdout.contains("╭─ ◉ docwarp v") || stdout.contains("+-- o docwarp v"),
        "expected startup box title, got:\n{stdout}"
    );
    assert!(
        stdout.contains("md ⇄ docx") || stdout.contains("md <--> docx"),
        "expected conversion label in startup header, got:\n{stdout}"
    );
}

fn assert_guided_startup_header(stdout: &str) {
    assert_startup_header(stdout);
    assert!(
        stdout.contains("Paste a path, or press Enter to choose."),
        "expected guided picker instruction in startup header, got:\n{stdout}"
    );
    assert!(
        stdout.contains("Type / for options."),
        "expected guided options instruction in startup header, got:\n{stdout}"
    );
}

fn run_docwarp<const N: usize>(args: [&str; N]) -> Result<Output> {
    Command::new(env!("CARGO_BIN_EXE_docwarp"))
        .args(args)
        .output()
        .context("failed running docwarp")
}

fn run_guided_with_stdin(stdin_payload: &str) -> Result<Output> {
    let mut child = Command::new(env!("CARGO_BIN_EXE_docwarp"))
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .context("failed spawning guided docwarp process")?;

    {
        let mut stdin = child
            .stdin
            .take()
            .context("failed opening guided stdin pipe")?;
        stdin
            .write_all(stdin_payload.as_bytes())
            .context("failed writing guided stdin payload")?;
    }

    child
        .wait_with_output()
        .context("failed waiting for guided process output")
}

fn assert_command_status(output: &Output, expected: Option<i32>, label: &str) -> Result<()> {
    if output.status.code() == expected {
        return Ok(());
    }

    bail!(
        "{} expected status {:?} but got {:?}\nstdout:\n{}\nstderr:\n{}",
        label,
        expected,
        output.status.code(),
        stdout_text(output),
        stderr_text(output)
    )
}

fn stdout_text(output: &Output) -> String {
    String::from_utf8_lossy(&output.stdout).into_owned()
}

fn stderr_text(output: &Output) -> String {
    String::from_utf8_lossy(&output.stderr).into_owned()
}
