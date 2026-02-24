use std::fs;
use std::io::{Read, Write};
use std::path::Path;
use std::process::{Command, Output};

use anyhow::{Context, Result, bail};
use docwarp_core::style_map;
use tempfile::tempdir;
use zip::ZipWriter;
use zip::read::ZipArchive;
use zip::write::SimpleFileOptions;

#[test]
fn template_map_generates_yaml_and_json_maps() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let template = temp.path().join("brand.dotx");
    let output_dir = temp.path().join("style-maps");

    write_company_template_zip(&template)?;

    let run = run_docwarp(&[
        "template-map".to_string(),
        template.to_string_lossy().into_owned(),
        "--output-dir".to_string(),
        output_dir.to_string_lossy().into_owned(),
        "--name".to_string(),
        "acme".to_string(),
    ])?;
    assert_command_status(&run, Some(0), "template-map should succeed")?;

    let yaml_path = output_dir.join("acme.yml");
    let json_path = output_dir.join("acme.json");
    assert!(yaml_path.exists(), "expected generated YAML map");
    assert!(json_path.exists(), "expected generated JSON map");

    let yaml_map = style_map::load_style_map(&yaml_path)
        .with_context(|| format!("generated YAML map should parse: {}", yaml_path.display()))?;
    let json_map = style_map::load_style_map(&json_path)
        .with_context(|| format!("generated JSON map should parse: {}", json_path.display()))?;
    assert_eq!(
        yaml_map, json_map,
        "YAML and JSON maps should be equivalent"
    );

    assert_eq!(
        yaml_map.md_to_docx.get("h1"),
        Some(&"CorpHeading1".to_string())
    );
    assert_eq!(
        yaml_map.md_to_docx.get("paragraph"),
        Some(&"CorpBody".to_string())
    );
    assert_eq!(
        yaml_map.md_to_docx.get("table"),
        Some(&"CorpTable".to_string())
    );
    assert_eq!(
        yaml_map.docx_to_md.get("Corp Heading 1"),
        Some(&"h1".to_string()),
        "display names should be included in reverse map"
    );

    Ok(())
}

#[test]
fn generated_template_map_is_immediately_usable_for_conversions() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let template = temp.path().join("brand.dotx");
    let maps_dir = temp.path().join("maps");
    let input_md = temp.path().join("input.md");
    let output_docx = temp.path().join("output.docx");
    let roundtrip_md = temp.path().join("roundtrip.md");

    write_company_template_zip(&template)?;
    fs::write(
        &input_md,
        "# Company Spec\n\nBody paragraph.\n\n- Item one\n\n| Col |\n| --- |\n| v |\n",
    )
    .context("failed writing input markdown")?;

    let extract = run_docwarp(&[
        "template-map".to_string(),
        template.to_string_lossy().into_owned(),
        "--output-dir".to_string(),
        maps_dir.to_string_lossy().into_owned(),
    ])?;
    assert_command_status(&extract, Some(0), "template-map should succeed")?;

    let generated_yaml = maps_dir.join("brand-style-map.yml");
    let generated_json = maps_dir.join("brand-style-map.json");
    assert!(generated_yaml.exists(), "expected generated YAML style map");
    assert!(generated_json.exists(), "expected generated JSON style map");

    let md2docx = run_docwarp(&[
        "md2docx".to_string(),
        input_md.to_string_lossy().into_owned(),
        "--output".to_string(),
        output_docx.to_string_lossy().into_owned(),
        "--template".to_string(),
        template.to_string_lossy().into_owned(),
        "--style-map".to_string(),
        generated_yaml.to_string_lossy().into_owned(),
    ])?;
    assert_command_status(
        &md2docx,
        Some(0),
        "md2docx should succeed with generated map",
    )?;

    let document_xml = read_document_xml(&output_docx)?;
    for expected_style in ["CorpHeading1", "CorpBody", "CorpBulletList", "CorpTable"] {
        assert!(
            document_xml.contains(expected_style),
            "expected generated map to apply template style `{expected_style}`\n{document_xml}"
        );
    }

    let docx2md = run_docwarp(&[
        "docx2md".to_string(),
        output_docx.to_string_lossy().into_owned(),
        "--output".to_string(),
        roundtrip_md.to_string_lossy().into_owned(),
        "--style-map".to_string(),
        generated_json.to_string_lossy().into_owned(),
        "--assets-dir".to_string(),
        temp.path().join("assets").to_string_lossy().into_owned(),
    ])?;
    assert_command_status(
        &docx2md,
        Some(0),
        "docx2md should succeed with generated reverse map",
    )?;

    let rendered = fs::read_to_string(&roundtrip_md).with_context(|| {
        format!(
            "failed reading roundtrip markdown: {}",
            roundtrip_md.display()
        )
    })?;
    assert!(
        rendered.contains("# Company Spec"),
        "heading should survive roundtrip with generated map:\n{rendered}"
    );

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
        "{label} failed with status {:?}\nstdout:\n{}\nstderr:\n{}",
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

fn read_document_xml(docx_path: &Path) -> Result<String> {
    let file = fs::File::open(docx_path)
        .with_context(|| format!("failed opening docx: {}", docx_path.display()))?;
    let mut archive = ZipArchive::new(file).context("failed reading docx as zip")?;
    let mut xml = String::new();
    archive
        .by_name("word/document.xml")
        .context("docx missing word/document.xml")?
        .read_to_string(&mut xml)
        .context("failed reading word/document.xml")?;
    Ok(xml)
}

fn write_company_template_zip(path: &Path) -> Result<()> {
    let styles_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="CorpTitle"><w:name w:val="Corp Title"/></w:style>
  <w:style w:type="paragraph" w:styleId="CorpHeading1"><w:name w:val="Corp Heading 1"/></w:style>
  <w:style w:type="paragraph" w:styleId="CorpHeading2"><w:name w:val="Corp Heading 2"/></w:style>
  <w:style w:type="paragraph" w:styleId="CorpBody"><w:name w:val="Corp Body Text"/></w:style>
  <w:style w:type="paragraph" w:styleId="CorpQuote"><w:name w:val="Corp Block Quote"/></w:style>
  <w:style w:type="paragraph" w:styleId="CorpCode"><w:name w:val="Corp Code Block"/></w:style>
  <w:style w:type="character" w:styleId="CorpEqInline"><w:name w:val="Corp Equation Inline"/></w:style>
  <w:style w:type="paragraph" w:styleId="CorpEqBlock"><w:name w:val="Corp Equation Block"/></w:style>
  <w:style w:type="paragraph" w:styleId="CorpBulletList"><w:name w:val="Corp Bullet List"/></w:style>
  <w:style w:type="paragraph" w:styleId="CorpNumberList"><w:name w:val="Corp Numbered List"/></w:style>
  <w:style w:type="table" w:styleId="CorpTable"><w:name w:val="Corp Table"/></w:style>
</w:styles>"#;

    let file = fs::File::create(path)
        .with_context(|| format!("failed creating template zip: {}", path.display()))?;
    let mut zip = ZipWriter::new(file);
    zip.start_file("word/styles.xml", SimpleFileOptions::default())?;
    zip.write_all(styles_xml.as_bytes())?;
    zip.finish()?;
    Ok(())
}
