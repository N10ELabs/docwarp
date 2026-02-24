use std::fs;
use std::io::{Read, Write};
use std::path::Path;
use std::process::{Command, Output};

use anyhow::{Context, Result, bail};
use docwarp_md::parse_markdown;
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

    let md2docx = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        docx.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&md2docx, Some(0), "md2docx should succeed")?;

    let docx2md = run_docwarp([
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
            if let docwarp_core::Block::List { ordered, .. } = block {
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

    let run = run_docwarp([
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

    let run = run_docwarp([
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
    let ok_run = run_docwarp([
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
    let fallback_run = run_docwarp([
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

#[test]
fn style_map_can_target_template_style_names_and_aliases_end_to_end() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("input.md");
    let output_docx = temp.path().join("out.docx");
    let output_md = temp.path().join("out.md");
    let template = temp.path().join("brand.dotx");
    let style_map = temp.path().join("style-map.yml");

    fs::write(&input, "# Heading\n\nBody with `code`.\n")
        .context("failed writing markdown input")?;

    write_template_zip(
        &template,
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="CorpHeading1">
    <w:name w:val="Corporate Heading 1"/>
    <w:aliases w:val="Corp H1"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CorpBody"><w:name w:val="Corporate Body"/></w:style>
  <w:style w:type="character" w:styleId="CorpCodeChar"><w:name w:val="Corporate Code"/></w:style>
  <w:style w:type="paragraph" w:styleId="CorpCodePara">
    <w:name w:val="Corporate Code Block"/>
    <w:link w:val="CorpCodeChar"/>
  </w:style>
</w:styles>"#,
    )?;

    fs::write(
        &style_map,
        r#"md_to_docx:
  h1: Corp H1
  paragraph: Corporate Body
  code: Corporate Code Block
docx_to_md:
  Corporate Heading 1: h1
  Corporate Body: paragraph
"#,
    )
    .context("failed writing style-map fixture")?;

    let md2docx = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output_docx.to_string_lossy().as_ref(),
        "--template",
        template.to_string_lossy().as_ref(),
        "--style-map",
        style_map.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&md2docx, Some(0), "md2docx should succeed")?;

    let document_xml = read_document_xml(&output_docx)?;
    assert!(
        document_xml.contains("<w:pStyle w:val=\"CorpHeading1\"/>"),
        "expected heading style name/alias to resolve to template styleId"
    );
    assert!(
        document_xml.contains("<w:pStyle w:val=\"CorpBody\"/>"),
        "expected paragraph style name to resolve to template styleId"
    );
    assert!(
        document_xml.contains("<w:rStyle w:val=\"CorpCodeChar\"/>"),
        "expected inline code to use linked template character style"
    );

    let docx2md = run_docwarp([
        "docx2md",
        output_docx.to_string_lossy().as_ref(),
        "--output",
        output_md.to_string_lossy().as_ref(),
        "--style-map",
        style_map.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&docx2md, Some(0), "docx2md should succeed")?;

    let roundtrip = fs::read_to_string(&output_md).context("failed reading markdown output")?;
    assert!(
        roundtrip.starts_with("# Heading"),
        "expected heading token mapping from template style name, got:\n{roundtrip}"
    );

    Ok(())
}

#[test]
fn equations_roundtrip_md_docx_md_with_native_omml() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("equations.md");
    let output_docx = temp.path().join("equations.docx");
    let output_md = temp.path().join("equations.roundtrip.md");

    fs::write(&input, "Inline equation: $x^2 + y^2$\n\n$$\nE=mc^2\n$$\n")
        .context("failed writing markdown with equations")?;

    let md2docx = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output_docx.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(
        &md2docx,
        Some(0),
        "md2docx should succeed for equation input",
    )?;

    let docx2md = run_docwarp([
        "docx2md",
        output_docx.to_string_lossy().as_ref(),
        "--output",
        output_md.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(
        &docx2md,
        Some(0),
        "docx2md should succeed for equation output",
    )?;

    let rendered = fs::read_to_string(&output_md).context("failed reading roundtrip markdown")?;
    assert!(
        rendered.contains("Inline equation: $x^2 + y^2$"),
        "expected inline equation in markdown output, got:\n{rendered}"
    );
    assert!(
        rendered.contains("$$\nE=mc^2\n$$"),
        "expected canonical display equation block in markdown output, got:\n{rendered}"
    );

    Ok(())
}

#[test]
fn strict_mode_returns_exit_code_2_for_unsupported_omml_equations() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input_docx = temp.path().join("unsupported-omml.docx");
    let output_md = temp.path().join("unsupported-omml.md");

    write_docx_with_document_xml(
        &input_docx,
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
    <w:body>
    <w:p>
      <m:oMath>
        <m:groupChr>
          <m:e><m:r><m:t>x</m:t></m:r></m:e>
        </m:groupChr>
      </m:oMath>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>"#,
    )?;

    let run = run_docwarp([
        "docx2md",
        input_docx.to_string_lossy().as_ref(),
        "--output",
        output_md.to_string_lossy().as_ref(),
        "--strict",
    ])?;

    assert_command_status(
        &run,
        Some(2),
        "strict docx2md should return exit code 2 on unsupported equation warning",
    )?;
    assert!(
        stdout_text(&run).contains("[unsupported_feature]"),
        "expected unsupported_feature warning in strict output, got:\n{}",
        stdout_text(&run)
    );

    Ok(())
}

fn run_docwarp<const N: usize>(args: [&str; N]) -> Result<Output> {
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

fn write_docx_with_document_xml(path: &Path, document_xml: &str) -> Result<()> {
    let file = fs::File::create(path)?;
    let mut zip = ZipWriter::new(file);
    zip.start_file("word/document.xml", SimpleFileOptions::default())?;
    zip.write_all(document_xml.as_bytes())?;
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

fn read_document_xml(docx_path: &Path) -> Result<String> {
    let mut archive = ZipArchive::new(
        fs::File::open(docx_path)
            .with_context(|| format!("failed opening output docx: {}", docx_path.display()))?,
    )
    .context("failed reading output docx as zip archive")?;
    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")
        .context("output docx missing word/document.xml")?
        .read_to_string(&mut document_xml)
        .context("failed reading word/document.xml")?;
    Ok(document_xml)
}
