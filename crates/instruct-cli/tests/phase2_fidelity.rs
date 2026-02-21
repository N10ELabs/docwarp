use std::fs;
use std::io::{Read, Write};
use std::path::Path;
use std::process::{Command, Output};

use anyhow::{Context, Result, bail};
use instruct_md::parse_markdown;
use tempfile::tempdir;
use zip::ZipWriter;
use zip::read::ZipArchive;
use zip::write::SimpleFileOptions;

#[test]
fn mixed_list_type_transitions_survive_md_docx_md_roundtrip() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("input.md");
    let docx = temp.path().join("out.docx");
    let output = temp.path().join("out.md");

    fs::write(
        &input,
        "1. ordered one\n2. ordered two\n\n- bullet one\n- bullet two\n\n1. ordered three\n",
    )
    .context("failed writing input markdown")?;

    let md2docx = run_instruct([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        docx.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&md2docx, Some(0), "md2docx should succeed")?;

    let docx2md = run_instruct([
        "docx2md",
        docx.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
        "--assets-dir",
        "assets",
    ])?;
    assert_command_status(&docx2md, Some(0), "docx2md should succeed")?;

    let generated = fs::read_to_string(&output).context("failed reading roundtrip markdown")?;
    let (doc, warnings) =
        parse_markdown(&generated).context("failed parsing roundtrip markdown")?;
    assert!(
        warnings.is_empty(),
        "expected no parser warnings, got: {warnings:?}"
    );

    let list_kinds: Vec<bool> = doc
        .blocks
        .iter()
        .filter_map(|block| {
            if let instruct_core::Block::List { ordered, .. } = block {
                Some(*ordered)
            } else {
                None
            }
        })
        .collect();

    assert_eq!(list_kinds, vec![true, false, true]);
    Ok(())
}

#[test]
fn remote_images_are_blocked_by_default_without_network_fetch() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("remote.md");
    let output = temp.path().join("out.docx");
    fs::write(&input, "![remote](https://example.invalid/image.png)\n")
        .context("failed writing markdown input")?;

    let run = run_instruct([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
    ])?;

    assert_command_status(&run, Some(0), "md2docx should succeed with warning")?;
    let stdout = stdout_text(&run);
    assert!(
        stdout.contains("[remote_image_blocked]"),
        "expected remote_image_blocked warning, got:\n{stdout}"
    );
    assert!(
        stdout.contains("offline-by-default"),
        "warning should mention offline-by-default policy, got:\n{stdout}"
    );
    assert!(
        !stdout.contains("[image_load_failed]"),
        "default policy should block before any fetch attempt, got:\n{stdout}"
    );
    assert!(output.is_file(), "output docx should still be produced");
    Ok(())
}

#[test]
fn missing_relative_and_absolute_images_explain_resolution_behavior() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("input.md");
    let output = temp.path().join("out.docx");
    let missing_absolute = temp.path().join("missing-absolute.png");

    let markdown = format!(
        "![rel](missing-relative.png)\n\n![abs]({})\n",
        missing_absolute.to_string_lossy()
    );
    fs::write(&input, markdown).context("failed writing markdown input")?;

    let run = run_instruct([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
        "--strict",
    ])?;

    assert_command_status(&run, Some(2), "strict mode should return warning exit code")?;
    let stdout = stdout_text(&run);
    assert!(
        stdout.contains("relative local image"),
        "expected relative-path resolution detail, got:\n{stdout}"
    );
    assert!(
        stdout.contains("absolute local image"),
        "expected absolute-path resolution detail, got:\n{stdout}"
    );
    Ok(())
}

#[test]
fn dotx_template_is_applied_and_invalid_template_falls_back() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("input.md");
    fs::write(&input, "Body\n").context("failed writing markdown input")?;

    let template = temp.path().join("brand.dotx");
    write_template_zip(
        &template,
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="BrandStyle"><w:name w:val="BrandStyle"/></w:style>
</w:styles>"#,
    )?;

    let output_ok = temp.path().join("ok.docx");
    let ok_run = run_instruct([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output_ok.to_string_lossy().as_ref(),
        "--template",
        template.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&ok_run, Some(0), "valid .dotx template should succeed")?;
    let styles_ok = read_styles_xml(&output_ok)?;
    assert!(
        styles_ok.contains("BrandStyle"),
        "output should include template style definitions"
    );

    let broken_template = temp.path().join("broken.dotx");
    write_invalid_template_zip(&broken_template)?;

    let output_fallback = temp.path().join("fallback.docx");
    let fallback_run = run_instruct([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output_fallback.to_string_lossy().as_ref(),
        "--template",
        broken_template.to_string_lossy().as_ref(),
        "--strict",
    ])?;
    assert_command_status(
        &fallback_run,
        Some(2),
        "invalid template should warn and fall back under strict",
    )?;
    assert!(
        stdout_text(&fallback_run).contains("[invalid_template]"),
        "invalid template warning should be emitted"
    );

    let styles_fallback = read_styles_xml(&output_fallback)?;
    assert!(
        styles_fallback.contains("ListBullet"),
        "fallback should use built-in styles.xml"
    );

    Ok(())
}

fn run_instruct<const N: usize>(args: [&str; N]) -> Result<Output> {
    Command::new(env!("CARGO_BIN_EXE_instruct"))
        .args(args)
        .output()
        .context("failed running instruct")
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

fn write_template_zip(path: &Path, styles_xml: &[u8]) -> Result<()> {
    let file = fs::File::create(path)?;
    let mut zip = ZipWriter::new(file);
    zip.start_file("word/styles.xml", SimpleFileOptions::default())?;
    zip.write_all(styles_xml)?;
    zip.finish()?;
    Ok(())
}

fn write_invalid_template_zip(path: &Path) -> Result<()> {
    let file = fs::File::create(path)?;
    let mut zip = ZipWriter::new(file);
    zip.start_file("word/not-styles.xml", SimpleFileOptions::default())?;
    zip.write_all(b"placeholder")?;
    zip.finish()?;
    Ok(())
}

fn read_styles_xml(docx_path: &Path) -> Result<String> {
    let mut archive = ZipArchive::new(
        fs::File::open(docx_path)
            .with_context(|| format!("failed opening output docx: {}", docx_path.display()))?,
    )
    .context("failed reading output docx as zip archive")?;
    let mut styles = String::new();
    archive
        .by_name("word/styles.xml")
        .context("output docx missing word/styles.xml")?
        .read_to_string(&mut styles)
        .context("failed reading word/styles.xml")?;
    Ok(styles)
}
