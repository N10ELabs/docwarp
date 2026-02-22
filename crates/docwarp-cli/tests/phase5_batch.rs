use std::fs;
use std::process::{Command, Output};

use anyhow::{Context, Result, bail};
use tempfile::tempdir;

#[test]
fn md2docx_directory_input_converts_markdown_tree() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input_root = temp.path().join("md-input");
    let output_root = temp.path().join("docx-output");

    fs::create_dir_all(input_root.join("nested")).context("failed creating nested input dir")?;
    fs::write(input_root.join("root.md"), "# Root\n\nBody\n").context("failed writing root.md")?;
    fs::write(input_root.join("nested/child.md"), "# Child\n\nBody\n")
        .context("failed writing child.md")?;
    fs::write(input_root.join("nested/ignore.txt"), "not markdown")
        .context("failed writing ignore.txt")?;

    let args = vec![
        "md2docx".to_string(),
        input_root.to_string_lossy().into_owned(),
        "--output".to_string(),
        output_root.to_string_lossy().into_owned(),
    ];
    let run = run_docwarp(&args)?;
    assert_command_status(&run, Some(0), "batch md2docx should succeed")?;

    assert!(output_root.join("root.docx").is_file());
    assert!(output_root.join("nested/child.docx").is_file());
    assert!(!output_root.join("nested/ignore.docx").exists());
    Ok(())
}

#[test]
fn docx2md_batch_glob_filters_inputs() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let md_root = temp.path().join("md-source");
    let docx_root = temp.path().join("docx-input");
    let output_root = temp.path().join("md-output");

    fs::create_dir_all(md_root.join("nested")).context("failed creating markdown nested dir")?;
    fs::write(md_root.join("top.md"), "# Top\n\nTop body\n").context("failed writing top.md")?;
    fs::write(md_root.join("nested/inner.md"), "# Inner\n\nInner body\n")
        .context("failed writing inner.md")?;

    fs::create_dir_all(docx_root.join("nested")).context("failed creating docx nested dir")?;

    let top_docx = docx_root.join("top.docx");
    let top_build_args = vec![
        "md2docx".to_string(),
        md_root.join("top.md").to_string_lossy().into_owned(),
        "--output".to_string(),
        top_docx.to_string_lossy().into_owned(),
    ];
    let top_build = run_docwarp(&top_build_args)?;
    assert_command_status(&top_build, Some(0), "setup top docx should succeed")?;

    let inner_docx = docx_root.join("nested/inner.docx");
    let inner_build_args = vec![
        "md2docx".to_string(),
        md_root
            .join("nested/inner.md")
            .to_string_lossy()
            .into_owned(),
        "--output".to_string(),
        inner_docx.to_string_lossy().into_owned(),
    ];
    let inner_build = run_docwarp(&inner_build_args)?;
    assert_command_status(&inner_build, Some(0), "setup inner docx should succeed")?;

    let batch_args = vec![
        "docx2md".to_string(),
        docx_root.to_string_lossy().into_owned(),
        "--output".to_string(),
        output_root.to_string_lossy().into_owned(),
        "--glob".to_string(),
        "nested/*.docx".to_string(),
    ];
    let batch_run = run_docwarp(&batch_args)?;
    assert_command_status(
        &batch_run,
        Some(0),
        "batch docx2md with glob should succeed",
    )?;

    assert!(output_root.join("nested/inner.md").is_file());
    assert!(!output_root.join("top.md").exists());
    Ok(())
}

#[test]
fn batch_strict_mode_returns_exit_code_2_when_warnings_exist() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input_root = temp.path().join("md-input");
    let output_root = temp.path().join("docx-output");

    fs::create_dir_all(&input_root).context("failed creating batch input dir")?;
    fs::write(input_root.join("ok.md"), "# Ok\n\nBody\n").context("failed writing ok.md")?;
    fs::write(
        input_root.join("warn.md"),
        "![remote](https://example.invalid/image.png)\n",
    )
    .context("failed writing warn.md")?;

    let args = vec![
        "md2docx".to_string(),
        input_root.to_string_lossy().into_owned(),
        "--output".to_string(),
        output_root.to_string_lossy().into_owned(),
        "--strict".to_string(),
    ];
    let run = run_docwarp(&args)?;
    assert_command_status(
        &run,
        Some(2),
        "strict batch conversion should return warning exit code",
    )?;
    assert!(stdout_text(&run).contains("[remote_image_blocked]"));

    assert!(output_root.join("ok.docx").is_file());
    assert!(output_root.join("warn.docx").is_file());
    Ok(())
}

#[test]
fn glob_requires_directory_input() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("single.md");
    let output = temp.path().join("single.docx");
    fs::write(&input, "# Single\n").context("failed writing single markdown input")?;

    let args = vec![
        "md2docx".to_string(),
        input.to_string_lossy().into_owned(),
        "--output".to_string(),
        output.to_string_lossy().into_owned(),
        "--glob".to_string(),
        "*.md".to_string(),
    ];
    let run = run_docwarp(&args)?;
    assert_command_status(&run, Some(1), "glob should fail for file input")?;
    assert!(stderr_text(&run).contains("--glob requires a directory input path"));
    Ok(())
}

fn run_docwarp(args: &[String]) -> Result<Output> {
    Command::new(env!("CARGO_BIN_EXE_docwarp"))
        .args(args)
        .output()
        .context("failed running docwarp")
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
