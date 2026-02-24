use std::collections::BTreeSet;
use std::fs;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::process::{Command, Output};

use anyhow::{Context, Result, bail};
use docwarp_core::{Block, Document, Inline};
use docwarp_md::parse_markdown;
use tempfile::tempdir;
use zip::ZipWriter;
use zip::read::ZipArchive;
use zip::write::SimpleFileOptions;

#[test]
fn company_template_name_and_alias_style_map_applies_across_sample_docs() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let fixtures = company_fixtures_root();
    let template = temp.path().join("acme.dotx");
    let style_map = fixtures.join("style-maps/acme-by-name.yml");
    let output_root = temp.path().join("docx-output");

    write_acme_template_zip(&template, &fixtures)?;
    fs::create_dir_all(&output_root).context("failed creating docx output root")?;

    let expectations = vec![
        (
            "01-headings-and-body.md",
            vec!["AcmeHeading1", "AcmeHeading2", "AcmeHeading3", "AcmeBody"],
        ),
        (
            "02-nested-lists.md",
            vec!["AcmeHeading1", "AcmeBulletList", "AcmeNumberList"],
        ),
        (
            "03-code-and-quote.md",
            vec![
                "AcmeHeading1",
                "AcmeQuote",
                "AcmeCodeBlock",
                "AcmeCodeInline",
            ],
        ),
        (
            "04-tables-and-links.md",
            vec!["AcmeHeading1", "AcmeTable", "AcmeBody"],
        ),
        (
            "05-equations.md",
            vec!["AcmeHeading1", "AcmeEquationBlock", "AcmeBody"],
        ),
        (
            "06-comprehensive.md",
            vec![
                "AcmeHeading1",
                "AcmeHeading2",
                "AcmeBulletList",
                "AcmeQuote",
                "AcmeTable",
                "AcmeCodeInline",
                "AcmeEquationBlock",
            ],
        ),
    ];

    for (sample, markers) in expectations {
        let input = fixtures.join("md").join(sample);
        let output = output_root.join(sample.replace(".md", ".docx"));
        let run = run_docwarp(
            &[
                "md2docx".to_string(),
                input.to_string_lossy().into_owned(),
                "--output".to_string(),
                output.to_string_lossy().into_owned(),
                "--template".to_string(),
                template.to_string_lossy().into_owned(),
                "--style-map".to_string(),
                style_map.to_string_lossy().into_owned(),
            ],
            None,
        )?;
        assert_command_status(&run, Some(0), "sample md2docx should succeed")?;

        let document_xml = read_document_xml(&output)?;
        for marker in markers {
            assert!(
                document_xml.contains(marker),
                "expected marker `{marker}` in output for {sample}\n{}",
                document_xml
            );
        }
    }

    Ok(())
}

#[test]
fn company_template_list_numbering_uses_template_numids_and_levels() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let fixtures = company_fixtures_root();
    let template = temp.path().join("acme.dotx");
    let style_map = fixtures.join("style-maps/acme-by-name.yml");
    let input = fixtures.join("md/02-nested-lists.md");
    let output = temp.path().join("lists.docx");

    write_acme_template_zip(&template, &fixtures)?;

    let run = run_docwarp(
        &[
            "md2docx".to_string(),
            input.to_string_lossy().into_owned(),
            "--output".to_string(),
            output.to_string_lossy().into_owned(),
            "--template".to_string(),
            template.to_string_lossy().into_owned(),
            "--style-map".to_string(),
            style_map.to_string_lossy().into_owned(),
        ],
        None,
    )?;
    assert_command_status(&run, Some(0), "list sample md2docx should succeed")?;

    let document_xml = read_document_xml(&output)?;
    assert!(
        document_xml.contains("<w:pStyle w:val=\"AcmeBulletList\"/>"),
        "expected bullet list style from template"
    );
    assert!(
        document_xml.contains("<w:pStyle w:val=\"AcmeNumberList\"/>"),
        "expected number list style from template"
    );
    assert!(
        document_xml.contains("<w:numId w:val=\"77\"/>"),
        "expected bullet list numId from template style definition"
    );
    assert!(
        document_xml.contains("<w:numId w:val=\"88\"/>"),
        "expected number list numId from template style definition"
    );
    assert!(
        document_xml.contains("<w:ilvl w:val=\"1\"/>")
            && document_xml.contains("<w:ilvl w:val=\"2\"/>"),
        "expected nested bullet list levels offset from template base ilvl"
    );

    Ok(())
}

#[test]
fn company_template_roundtrip_with_alias_map_preserves_structure_signals() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let fixtures = company_fixtures_root();
    let template = temp.path().join("acme.dotx");
    let style_map_md2docx = fixtures.join("style-maps/acme-by-name.yml");
    let style_map_docx2md = fixtures.join("style-maps/acme-docx-aliases.yml");
    let output_docx_root = temp.path().join("docx-output");
    let output_md_root = temp.path().join("md-output");

    write_acme_template_zip(&template, &fixtures)?;
    fs::create_dir_all(&output_docx_root).context("failed creating docx output root")?;
    fs::create_dir_all(&output_md_root).context("failed creating markdown output root")?;

    for input in sample_markdown_paths(&fixtures)? {
        let name = input
            .file_name()
            .and_then(|value| value.to_str())
            .context("sample file name should be valid UTF-8")?;
        let output_docx = output_docx_root.join(name.replace(".md", ".docx"));
        let output_md = output_md_root.join(name);

        let md2docx_run = run_docwarp(
            &[
                "md2docx".to_string(),
                input.to_string_lossy().into_owned(),
                "--output".to_string(),
                output_docx.to_string_lossy().into_owned(),
                "--template".to_string(),
                template.to_string_lossy().into_owned(),
                "--style-map".to_string(),
                style_map_md2docx.to_string_lossy().into_owned(),
            ],
            None,
        )?;
        assert_command_status(&md2docx_run, Some(0), "md2docx sample should succeed")?;

        let docx2md_run = run_docwarp(
            &[
                "docx2md".to_string(),
                output_docx.to_string_lossy().into_owned(),
                "--output".to_string(),
                output_md.to_string_lossy().into_owned(),
                "--style-map".to_string(),
                style_map_docx2md.to_string_lossy().into_owned(),
                "--assets-dir".to_string(),
                output_md_root.join("assets").to_string_lossy().into_owned(),
            ],
            None,
        )?;
        assert_command_status(&docx2md_run, Some(0), "docx2md sample should succeed")?;

        let source = fs::read_to_string(&input)
            .with_context(|| format!("failed reading source markdown: {}", input.display()))?;
        let rendered = fs::read_to_string(&output_md).with_context(|| {
            format!("failed reading roundtrip markdown: {}", output_md.display())
        })?;

        let (source_doc, source_warnings) =
            parse_markdown(&source).context("failed parsing source markdown")?;
        let (roundtrip_doc, roundtrip_warnings) =
            parse_markdown(&rendered).context("failed parsing roundtrip markdown")?;

        assert!(
            source_warnings.is_empty(),
            "source markdown should parse cleanly for {name}"
        );
        assert!(
            roundtrip_warnings.is_empty(),
            "roundtrip markdown should parse cleanly for {name}"
        );

        assert_eq!(
            heading_levels(&source_doc),
            heading_levels(&roundtrip_doc),
            "heading levels should roundtrip for {name}"
        );
        assert_eq!(
            list_kinds(&source_doc),
            list_kinds(&roundtrip_doc),
            "ordered/unordered list grouping should roundtrip for {name}"
        );
        assert_eq!(
            block_kind_set(&source_doc),
            block_kind_set(&roundtrip_doc),
            "block kind coverage should roundtrip for {name}"
        );
        assert_eq!(
            equation_counts(&source_doc),
            equation_counts(&roundtrip_doc),
            "equation counts should roundtrip for {name}"
        );
    }

    Ok(())
}

#[test]
fn company_template_batch_conversion_covers_entire_sample_folder() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let fixtures = company_fixtures_root();
    let template = temp.path().join("acme.dotx");
    let style_map = fixtures.join("style-maps/acme-by-name.yml");
    let input_root = fixtures.join("md");
    let output_docx_root = temp.path().join("batch-docx");
    let output_md_root = temp.path().join("batch-md");

    write_acme_template_zip(&template, &fixtures)?;

    let md2docx = run_docwarp(
        &[
            "md2docx".to_string(),
            input_root.to_string_lossy().into_owned(),
            "--output".to_string(),
            output_docx_root.to_string_lossy().into_owned(),
            "--template".to_string(),
            template.to_string_lossy().into_owned(),
            "--style-map".to_string(),
            style_map.to_string_lossy().into_owned(),
        ],
        None,
    )?;
    assert_command_status(&md2docx, Some(0), "batch md2docx should succeed")?;

    let source_count = sample_markdown_paths(&fixtures)?.len();
    let docx_count = count_files_with_extension(&output_docx_root, "docx")?;
    assert_eq!(
        source_count, docx_count,
        "all sample markdown files should convert"
    );

    let docx2md = run_docwarp(
        &[
            "docx2md".to_string(),
            output_docx_root.to_string_lossy().into_owned(),
            "--output".to_string(),
            output_md_root.to_string_lossy().into_owned(),
            "--style-map".to_string(),
            style_map.to_string_lossy().into_owned(),
        ],
        None,
    )?;
    assert_command_status(&docx2md, Some(0), "batch docx2md should succeed")?;

    let md_count = count_files_with_extension(&output_md_root, "md")?;
    assert_eq!(
        source_count, md_count,
        "all converted DOCX files should convert back"
    );

    Ok(())
}

#[test]
fn company_template_defaults_work_via_config_file() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let fixtures = company_fixtures_root();
    let workdir = temp.path().join("workspace");
    fs::create_dir_all(&workdir).context("failed creating workspace directory")?;

    let template = workdir.join("acme.dotx");
    let style_map = workdir.join("acme-style-map.yml");
    let input = workdir.join("input.md");
    let output = workdir.join("output.docx");
    let config = workdir.join(".docwarp.yml");

    write_acme_template_zip(&template, &fixtures)?;
    fs::copy(fixtures.join("style-maps/acme-by-name.yml"), &style_map)
        .context("failed copying style-map fixture")?;
    fs::copy(fixtures.join("md/06-comprehensive.md"), &input)
        .context("failed copying markdown fixture")?;

    fs::write(
        &config,
        "markdown_flavor: gfm\nstyle_map: ./acme-style-map.yml\ndefault_template: ./acme.dotx\nunsupported_policy: warn_continue\n",
    )
    .context("failed writing config fixture")?;

    let run = run_docwarp(
        &[
            "md2docx".to_string(),
            input.to_string_lossy().into_owned(),
            "--output".to_string(),
            output.to_string_lossy().into_owned(),
        ],
        Some(&workdir),
    )?;
    assert_command_status(&run, Some(0), "config-driven md2docx should succeed")?;

    let document_xml = read_document_xml(&output)?;
    for marker in [
        "AcmeHeading1",
        "AcmeHeading2",
        "AcmeTable",
        "AcmeQuote",
        "AcmeCodeInline",
        "AcmeEquationBlock",
    ] {
        assert!(
            document_xml.contains(marker),
            "expected config-driven output to include `{marker}`"
        );
    }

    Ok(())
}

fn sample_markdown_paths(fixtures: &Path) -> Result<Vec<PathBuf>> {
    let mut files = Vec::new();
    for entry in fs::read_dir(fixtures.join("md")).context("failed reading sample markdown dir")? {
        let entry = entry.context("failed reading sample markdown entry")?;
        let path = entry.path();
        if path
            .extension()
            .and_then(|value| value.to_str())
            .is_some_and(|ext| ext.eq_ignore_ascii_case("md"))
        {
            files.push(path);
        }
    }
    files.sort();
    Ok(files)
}

fn block_kind_set(document: &Document) -> BTreeSet<&'static str> {
    document
        .blocks
        .iter()
        .map(|block| match block {
            Block::Title(_) => "title",
            Block::Heading { .. } => "heading",
            Block::Paragraph(_) => "paragraph",
            Block::BlockQuote(_) => "quote",
            Block::CodeBlock { .. } => "code",
            Block::List { .. } => "list",
            Block::Table { .. } => "table",
            Block::Image { .. } => "image",
            Block::ThematicBreak => "thematic_break",
        })
        .collect()
}

fn heading_levels(document: &Document) -> Vec<u8> {
    document
        .blocks
        .iter()
        .filter_map(|block| {
            if let Block::Heading { level, .. } = block {
                Some(*level)
            } else {
                None
            }
        })
        .collect()
}

fn list_kinds(document: &Document) -> Vec<bool> {
    document
        .blocks
        .iter()
        .filter_map(|block| {
            if let Block::List { ordered, .. } = block {
                Some(*ordered)
            } else {
                None
            }
        })
        .collect()
}

fn equation_counts(document: &Document) -> (usize, usize) {
    let mut inline_count = 0usize;
    let mut display_count = 0usize;
    for block in &document.blocks {
        match block {
            Block::Title(inlines)
            | Block::Paragraph(inlines)
            | Block::BlockQuote(inlines)
            | Block::Heading {
                content: inlines, ..
            } => count_equations_in_inlines(inlines, &mut inline_count, &mut display_count),
            Block::CodeBlock { .. }
            | Block::List { .. }
            | Block::Table { .. }
            | Block::Image { .. }
            | Block::ThematicBreak => {}
        }
    }
    (inline_count, display_count)
}

fn count_equations_in_inlines(inlines: &[Inline], inline: &mut usize, display: &mut usize) {
    for node in inlines {
        match node {
            Inline::Equation { display: true, .. } => *display += 1,
            Inline::Equation { display: false, .. } => *inline += 1,
            Inline::Emphasis(children)
            | Inline::Strong(children)
            | Inline::Link { text: children, .. } => {
                count_equations_in_inlines(children, inline, display);
            }
            Inline::Text(_) | Inline::Code(_) | Inline::LineBreak | Inline::Image { .. } => {}
        }
    }
}

fn count_files_with_extension(root: &Path, extension: &str) -> Result<usize> {
    let mut count = 0usize;
    for entry in fs::read_dir(root).with_context(|| format!("failed reading {}", root.display()))? {
        let entry = entry.with_context(|| format!("failed iterating {}", root.display()))?;
        let path = entry.path();
        if path
            .extension()
            .and_then(|value| value.to_str())
            .is_some_and(|ext| ext.eq_ignore_ascii_case(extension))
        {
            count += 1;
        }
    }
    Ok(count)
}

fn company_fixtures_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../fixtures/company_templates")
        .canonicalize()
        .expect("fixtures/company_templates should exist")
}

fn write_acme_template_zip(path: &Path, fixtures: &Path) -> Result<()> {
    let styles_xml = fs::read(fixtures.join("template_parts/acme-styles.xml"))
        .context("failed reading acme-styles.xml fixture")?;
    let numbering_xml = fs::read(fixtures.join("template_parts/acme-numbering.xml"))
        .context("failed reading acme-numbering.xml fixture")?;

    let file = fs::File::create(path)
        .with_context(|| format!("failed creating template zip: {}", path.display()))?;
    let mut zip = ZipWriter::new(file);
    zip.start_file("word/styles.xml", SimpleFileOptions::default())?;
    zip.write_all(&styles_xml)?;
    zip.start_file("word/numbering.xml", SimpleFileOptions::default())?;
    zip.write_all(&numbering_xml)?;
    zip.finish()?;
    Ok(())
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

fn run_docwarp(args: &[String], workdir: Option<&Path>) -> Result<Output> {
    let mut command = Command::new(env!("CARGO_BIN_EXE_docwarp"));
    command.args(args);
    if let Some(workdir) = workdir {
        command.current_dir(workdir);
    }
    command.output().context("failed running docwarp")
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
