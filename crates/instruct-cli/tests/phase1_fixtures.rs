use std::ffi::OsStr;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::{Command, Output};

use anyhow::{Context, Result, bail};
use instruct_core::{Block, Document, Inline, StyleMap};
use instruct_docx::{DocxReadOptions, read_docx};
use tempfile::tempdir;

const FIXTURE_BASENAMES: [&str; 10] = [
    "01-title-heading-paragraph",
    "02-unordered-list",
    "03-ordered-list",
    "04-blockquote-link",
    "05-fenced-code",
    "06-table",
    "07-inline-formatting",
    "08-image-local",
    "09-mixed-structure",
    "10-comprehensive",
];

#[test]
fn fixture_corpus_contains_expected_samples() -> Result<()> {
    let root = workspace_root();

    for base in FIXTURE_BASENAMES {
        let md = root.join("fixtures/md").join(format!("{base}.md"));
        let docx = root.join("fixtures/docx").join(format!("{base}.docx"));
        let expected_md = root
            .join("fixtures/expected/docx2md")
            .join(format!("{base}.md"));

        if !md.is_file() {
            bail!("missing markdown fixture: {}", md.display());
        }
        if !docx.is_file() {
            bail!("missing DOCX fixture: {}", docx.display());
        }
        if !expected_md.is_file() {
            bail!(
                "missing expected markdown fixture: {}",
                expected_md.display()
            );
        }
    }

    Ok(())
}

#[test]
fn md_to_docx_matches_golden_structure() -> Result<()> {
    let root = workspace_root();
    let temp = tempdir().context("tempdir should be created")?;

    for base in FIXTURE_BASENAMES {
        let md_input = root.join("fixtures/md").join(format!("{base}.md"));
        let golden_docx = root.join("fixtures/docx").join(format!("{base}.docx"));
        let generated_docx = temp.path().join(format!("{base}.docx"));

        let output = Command::new(env!("CARGO_BIN_EXE_instruct"))
            .arg("md2docx")
            .arg(&md_input)
            .arg("--output")
            .arg(&generated_docx)
            .output()
            .with_context(|| format!("failed running md2docx for fixture {base}"))?;
        assert_command_success(&output, &format!("md2docx fixture {base}"))?;

        let mut generated = read_docx(
            &generated_docx,
            &DocxReadOptions {
                assets_dir: temp.path().join(format!("generated-assets-{base}")),
                style_map: StyleMap::builtin(),
            },
        )
        .with_context(|| format!("failed reading generated DOCX for fixture {base}"))?
        .0;

        let mut golden = read_docx(
            &golden_docx,
            &DocxReadOptions {
                assets_dir: temp.path().join(format!("golden-assets-{base}")),
                style_map: StyleMap::builtin(),
            },
        )
        .with_context(|| format!("failed reading golden DOCX fixture {base}"))?
        .0;

        normalize_document(&mut generated);
        normalize_document(&mut golden);

        assert_eq!(
            generated, golden,
            "md->docx structural mismatch for fixture {base}"
        );
    }

    Ok(())
}

#[test]
fn docx_to_md_matches_golden_markdown() -> Result<()> {
    let root = workspace_root();
    let temp = tempdir().context("tempdir should be created")?;

    for base in FIXTURE_BASENAMES {
        let docx_input = root.join("fixtures/docx").join(format!("{base}.docx"));
        let expected_md_path = root
            .join("fixtures/expected/docx2md")
            .join(format!("{base}.md"));
        let output_md = temp.path().join(format!("{base}.md"));

        let output = Command::new(env!("CARGO_BIN_EXE_instruct"))
            .arg("docx2md")
            .arg(&docx_input)
            .arg("--output")
            .arg(&output_md)
            .arg("--assets-dir")
            .arg("assets")
            .output()
            .with_context(|| format!("failed running docx2md for fixture {base}"))?;
        assert_command_success(&output, &format!("docx2md fixture {base}"))?;

        let actual = fs::read_to_string(&output_md)
            .with_context(|| format!("failed reading generated markdown for fixture {base}"))?;
        let expected = fs::read_to_string(&expected_md_path)
            .with_context(|| format!("failed reading expected markdown for fixture {base}"))?;

        assert_eq!(
            normalize_markdown(&actual),
            normalize_markdown(&expected),
            "docx->md mismatch for fixture {base}"
        );
    }

    Ok(())
}

fn workspace_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()
        .expect("workspace root should be resolvable")
}

fn assert_command_success(output: &Output, label: &str) -> Result<()> {
    if output.status.success() {
        return Ok(());
    }

    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);

    bail!(
        "{label} failed with status {}\nstdout:\n{}\nstderr:\n{}",
        output.status,
        stdout,
        stderr
    )
}

fn normalize_markdown(markdown: &str) -> String {
    markdown.replace("\r\n", "\n").trim_end().to_string()
}

fn normalize_document(document: &mut Document) {
    for block in &mut document.blocks {
        match block {
            Block::Title(inlines)
            | Block::Paragraph(inlines)
            | Block::BlockQuote(inlines)
            | Block::Heading {
                content: inlines, ..
            } => normalize_inline_images(inlines),
            Block::List { items, .. } => {
                for item in items {
                    normalize_inline_images(item);
                }
            }
            Block::Table { headers, rows } => {
                for cell in headers {
                    normalize_inline_images(cell);
                }
                for row in rows {
                    for cell in row {
                        normalize_inline_images(cell);
                    }
                }
            }
            Block::Image { src, .. } => normalize_image_path(src),
            Block::CodeBlock { .. } | Block::ThematicBreak => {}
        }
    }
}

fn normalize_inline_images(inlines: &mut [Inline]) {
    for inline in inlines {
        match inline {
            Inline::Image { src, .. } => normalize_image_path(src),
            Inline::Emphasis(children)
            | Inline::Strong(children)
            | Inline::Link { text: children, .. } => normalize_inline_images(children),
            Inline::Text(_) | Inline::Code(_) | Inline::LineBreak => {}
        }
    }
}

fn normalize_image_path(src: &mut String) {
    if src.starts_with("http://") || src.starts_with("https://") {
        return;
    }

    if let Some(name) = Path::new(src).file_name().and_then(OsStr::to_str) {
        *src = name.to_string();
    }
}
