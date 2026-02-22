use std::process::{Command, Output};

use anyhow::{Context, Result, bail};

#[test]
fn root_help_includes_examples() -> Result<()> {
    let output = run_docwarp(["--help"])?;
    assert_command_success(&output, "root --help should succeed")?;
    let stdout = stdout_text(&output);

    assert!(
        stdout.contains("Examples:"),
        "expected examples section in root help, got:\n{stdout}"
    );
    assert!(
        stdout.contains("docwarp"),
        "expected guided-mode example in root help, got:\n{stdout}"
    );
    assert!(
        stdout.contains("docwarp md2docx ./docs/spec.md --output ./build/spec.docx"),
        "expected md2docx example in root help, got:\n{stdout}"
    );
    assert!(
        stdout.contains("docwarp docx2md ./contracts/master.docx --output ./contracts/master.md"),
        "expected docx2md example in root help, got:\n{stdout}"
    );

    Ok(())
}

#[test]
fn md2docx_help_includes_examples() -> Result<()> {
    let output = run_docwarp(["md2docx", "--help"])?;
    assert_command_success(&output, "md2docx --help should succeed")?;
    let stdout = stdout_text(&output);

    assert!(
        stdout.contains("Examples:"),
        "expected examples section in md2docx help, got:\n{stdout}"
    );
    assert!(
        stdout.contains("--template ./brand.dotx --style-map ./style-map.yml"),
        "expected template/style-map example in md2docx help, got:\n{stdout}"
    );
    assert!(
        stdout.contains("--report ./report.json --strict"),
        "expected strict/report example in md2docx help, got:\n{stdout}"
    );
    assert!(
        stdout.contains("--glob \"**/*.md\""),
        "expected batch glob example in md2docx help, got:\n{stdout}"
    );

    Ok(())
}

#[test]
fn docx2md_help_includes_examples() -> Result<()> {
    let output = run_docwarp(["docx2md", "--help"])?;
    assert_command_success(&output, "docx2md --help should succeed")?;
    let stdout = stdout_text(&output);

    assert!(
        stdout.contains("Examples:"),
        "expected examples section in docx2md help, got:\n{stdout}"
    );
    assert!(
        stdout.contains("--assets-dir ./output_assets"),
        "expected assets-dir example in docx2md help, got:\n{stdout}"
    );
    assert!(
        stdout.contains("--config ./.docwarp.yml --report ./report.json"),
        "expected config/report example in docx2md help, got:\n{stdout}"
    );
    assert!(
        stdout.contains("--glob \"**/*.docx\""),
        "expected batch glob example in docx2md help, got:\n{stdout}"
    );

    Ok(())
}

fn run_docwarp<const N: usize>(args: [&str; N]) -> Result<Output> {
    Command::new(env!("CARGO_BIN_EXE_docwarp"))
        .args(args)
        .output()
        .context("failed running docwarp")
}

fn assert_command_success(output: &Output, label: &str) -> Result<()> {
    if output.status.success() {
        return Ok(());
    }

    bail!(
        "{label} failed with status {:?}\nstdout:\n{}\nstderr:\n{}",
        output.status.code(),
        stdout_text(output),
        String::from_utf8_lossy(&output.stderr)
    )
}

fn stdout_text(output: &Output) -> String {
    String::from_utf8_lossy(&output.stdout).into_owned()
}
