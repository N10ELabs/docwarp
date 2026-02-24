use std::ffi::OsStr;
use std::fs;
use std::io;
use std::path::{Path, PathBuf};
use std::process::{Command, Output};
use std::thread;
use std::time::Duration;

use anyhow::{Context, Result, bail};
use docwarp_core::{Block, Document, Inline, StyleMap, WarningCode};
use docwarp_docx::{DocxReadOptions, DocxWriteOptions, read_docx, write_docx};
use docwarp_md::parse_markdown;
use tempfile::tempdir;
use zip::ZipWriter;
use zip::read::ZipArchive;
use zip::write::SimpleFileOptions;

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
fn roundtrip_md_docx_md_preserves_structure() -> Result<()> {
    let root = workspace_root();
    let temp = tempdir().context("tempdir should be created")?;

    for base in FIXTURE_BASENAMES {
        let input_md_path = root.join("fixtures/md").join(format!("{base}.md"));
        let input_md = fs::read_to_string(&input_md_path)
            .with_context(|| format!("failed reading fixture markdown: {base}"))?;
        let mut original = parse_markdown(&input_md)
            .with_context(|| format!("failed parsing original markdown fixture: {base}"))?
            .0;

        let roundtrip_docx = temp.path().join(format!("{base}.docx"));
        let roundtrip_md = temp.path().join(format!("{base}.roundtrip.md"));

        let md2docx = run_docwarp([
            "md2docx",
            input_md_path.to_string_lossy().as_ref(),
            "--output",
            roundtrip_docx.to_string_lossy().as_ref(),
        ])?;
        assert_command_status(
            &md2docx,
            Some(0),
            &format!("md2docx roundtrip setup {base}"),
        )?;

        let docx2md = run_docwarp([
            "docx2md",
            roundtrip_docx.to_string_lossy().as_ref(),
            "--output",
            roundtrip_md.to_string_lossy().as_ref(),
            "--assets-dir",
            "assets",
        ])?;
        assert_command_status(&docx2md, Some(0), &format!("docx2md roundtrip {base}"))?;

        let output_md = fs::read_to_string(&roundtrip_md)
            .with_context(|| format!("failed reading roundtrip markdown: {base}"))?;
        let mut roundtripped = parse_markdown(&output_md)
            .with_context(|| format!("failed parsing roundtrip markdown fixture: {base}"))?
            .0;

        normalize_document_semantics(&mut original);
        normalize_document_semantics(&mut roundtripped);

        assert_eq!(
            original, roundtripped,
            "md->docx->md semantic mismatch for fixture {base}"
        );
    }

    Ok(())
}

#[test]
fn roundtrip_docx_md_docx_preserves_structure() -> Result<()> {
    let root = workspace_root();
    let temp = tempdir().context("tempdir should be created")?;

    for base in FIXTURE_BASENAMES {
        let input_docx = root.join("fixtures/docx").join(format!("{base}.docx"));
        let output_md = temp.path().join(format!("{base}.md"));
        let output_docx = temp.path().join(format!("{base}.roundtrip.docx"));

        let mut original = read_docx(
            &input_docx,
            &DocxReadOptions {
                assets_dir: temp.path().join(format!("assets-original-{base}")),
                style_map: StyleMap::builtin(),
                password: None,
            },
        )
        .with_context(|| format!("failed reading original DOCX fixture {base}"))?
        .0;

        let docx2md = run_docwarp([
            "docx2md",
            input_docx.to_string_lossy().as_ref(),
            "--output",
            output_md.to_string_lossy().as_ref(),
            "--assets-dir",
            "assets",
        ])?;
        assert_command_status(
            &docx2md,
            Some(0),
            &format!("docx2md roundtrip setup {base}"),
        )?;

        let md2docx = run_docwarp([
            "md2docx",
            output_md.to_string_lossy().as_ref(),
            "--output",
            output_docx.to_string_lossy().as_ref(),
        ])?;
        assert_command_status(&md2docx, Some(0), &format!("md2docx roundtrip {base}"))?;

        let mut roundtripped = read_docx(
            &output_docx,
            &DocxReadOptions {
                assets_dir: temp.path().join(format!("assets-roundtrip-{base}")),
                style_map: StyleMap::builtin(),
                password: None,
            },
        )
        .with_context(|| format!("failed reading roundtrip DOCX for fixture {base}"))?
        .0;

        normalize_document_semantics(&mut original);
        normalize_document_semantics(&mut roundtripped);

        assert_eq!(
            original, roundtripped,
            "docx->md->docx semantic mismatch for fixture {base}"
        );
    }

    Ok(())
}

#[test]
fn strict_mode_returns_exit_code_2_on_warnings() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("remote-image.md");
    let output = temp.path().join("out.docx");
    fs::write(&input, "![Remote](https://example.com/image.png)\n")
        .context("failed writing markdown input")?;

    let run = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
        "--strict",
    ])?;

    assert_command_status(&run, Some(2), "strict mode warning exit")?;
    assert!(
        stdout_text(&run).contains("[remote_image_blocked]"),
        "expected remote_image_blocked warning, got:\n{}",
        stdout_text(&run)
    );

    Ok(())
}

#[test]
fn corrupt_docx_returns_fatal_error() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("corrupt.docx");
    let output = temp.path().join("out.md");
    fs::write(&input, b"not a valid docx zip").context("failed writing corrupt DOCX")?;

    let run = run_docwarp([
        "docx2md",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
    ])?;

    assert_command_status(&run, Some(1), "corrupt DOCX should fail")?;
    assert!(
        stderr_text(&run).contains("failed opening DOCX zip archive"),
        "expected zip-archive error in stderr, got:\n{}",
        stderr_text(&run)
    );

    Ok(())
}

#[test]
fn missing_media_emits_warning() -> Result<()> {
    let root = workspace_root();
    let temp = tempdir().context("tempdir should be created")?;

    let source_docx = root.join("fixtures/docx/08-image-local.docx");
    let broken_docx = temp.path().join("missing-media.docx");
    remove_docx_entry(&source_docx, &broken_docx, "word/media/image1.png")?;

    let output_md = temp.path().join("out.md");
    let run = run_docwarp([
        "docx2md",
        broken_docx.to_string_lossy().as_ref(),
        "--output",
        output_md.to_string_lossy().as_ref(),
        "--assets-dir",
        "assets",
    ])?;

    assert_command_status(&run, Some(0), "missing media should warn and continue")?;
    assert!(
        stdout_text(&run).contains("[missing_media]"),
        "expected missing_media warning, got:\n{}",
        stdout_text(&run)
    );

    Ok(())
}

#[test]
fn md2docx_rejects_output_equal_to_input_path() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("input.md");
    fs::write(&input, "# Title\n\nBody\n").context("failed writing markdown")?;

    let run = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        input.to_string_lossy().as_ref(),
    ])?;

    assert_command_status(
        &run,
        Some(1),
        "md2docx should reject destructive input/output path reuse",
    )?;
    assert!(
        stderr_text(&run).contains("output path must differ from input path"),
        "expected destructive-overwrite guard in stderr, got:\n{}",
        stderr_text(&run)
    );

    Ok(())
}

#[test]
fn docx2md_rejects_output_equal_to_input_path() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("input.docx");
    let doc = Document {
        blocks: vec![Block::Paragraph(vec![Inline::Text("hello".into())])],
    };
    write_docx(
        &doc,
        temp.path(),
        &input,
        &DocxWriteOptions {
            allow_remote_images: false,
            style_map: StyleMap::builtin(),
            template: None,
        },
    )
    .context("failed creating DOCX input fixture")?;

    let run = run_docwarp([
        "docx2md",
        input.to_string_lossy().as_ref(),
        "--output",
        input.to_string_lossy().as_ref(),
    ])?;

    assert_command_status(
        &run,
        Some(1),
        "docx2md should reject destructive input/output path reuse",
    )?;
    assert!(
        stderr_text(&run).contains("output path must differ from input path"),
        "expected destructive-overwrite guard in stderr, got:\n{}",
        stderr_text(&run)
    );

    Ok(())
}

#[test]
fn md2docx_overwrite_creates_backup_by_default() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("input.md");
    let output = temp.path().join("out.docx");

    fs::write(&input, "# Title\n\nv1\n").context("failed writing markdown v1")?;
    let first = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&first, Some(0), "initial md2docx run should succeed")?;
    let original_output = fs::read(&output).context("failed reading original output DOCX")?;

    fs::write(&input, "# Title\n\nv2\n").context("failed writing markdown v2")?;
    let second = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&second, Some(0), "overwrite md2docx run should succeed")?;
    assert!(
        stdout_text(&second).contains("backup created ✓"),
        "expected backup creation log line, got:\n{}",
        stdout_text(&second)
    );

    let backup_dir = temp.path().join("docwarp_backups");
    let backups = collect_backup_files(&backup_dir, "out.docx")?;
    assert_eq!(
        backups.len(),
        1,
        "expected one backup file after first overwrite, found {} in {}",
        backups.len(),
        backup_dir.display()
    );

    let backup_bytes = fs::read(&backups[0]).context("failed reading backup file")?;
    assert_eq!(
        backups[0].extension().and_then(OsStr::to_str),
        Some("docx"),
        "backup should keep native DOCX extension"
    );
    assert_eq!(
        backup_bytes, original_output,
        "backup file should preserve bytes from pre-overwrite output"
    );

    Ok(())
}

#[test]
fn docx2md_overwrite_creates_backup_by_default() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("input.docx");
    let output = temp.path().join("out.md");
    let doc = Document {
        blocks: vec![Block::Paragraph(vec![Inline::Text("hello".into())])],
    };
    write_docx(
        &doc,
        temp.path(),
        &input,
        &DocxWriteOptions {
            allow_remote_images: false,
            style_map: StyleMap::builtin(),
            template: None,
        },
    )
    .context("failed creating DOCX input fixture")?;

    let first = run_docwarp([
        "docx2md",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&first, Some(0), "initial docx2md run should succeed")?;
    let original_output = fs::read(&output).context("failed reading original markdown output")?;

    let second = run_docwarp([
        "docx2md",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&second, Some(0), "overwrite docx2md run should succeed")?;
    assert!(
        stdout_text(&second).contains("backup created ✓"),
        "expected backup creation log line, got:\n{}",
        stdout_text(&second)
    );

    let backup_dir = temp.path().join("docwarp_backups");
    let backups = collect_backup_files(&backup_dir, "out.md")?;
    assert_eq!(
        backups.len(),
        1,
        "expected one backup file after overwrite, found {} in {}",
        backups.len(),
        backup_dir.display()
    );
    let backup_bytes = fs::read(&backups[0]).context("failed reading backup file")?;
    assert_eq!(
        backups[0].extension().and_then(OsStr::to_str),
        Some("md"),
        "backup should keep native Markdown extension"
    );
    assert_eq!(
        backup_bytes, original_output,
        "backup should preserve bytes from original markdown output"
    );

    Ok(())
}

#[test]
fn no_backup_flag_disables_overwrite_backups() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("input.md");
    let output = temp.path().join("out.docx");

    fs::write(&input, "# Title\n\nv1\n").context("failed writing markdown v1")?;
    let first = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&first, Some(0), "initial md2docx run should succeed")?;

    fs::write(&input, "# Title\n\nv2\n").context("failed writing markdown v2")?;
    let second = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
        "--no-backup",
    ])?;
    assert_command_status(
        &second,
        Some(0),
        "overwrite md2docx with no-backup should succeed",
    )?;
    assert!(
        !stdout_text(&second).contains("backup created ✓"),
        "did not expect backup creation log line, got:\n{}",
        stdout_text(&second)
    );

    let backups = collect_backup_files(&temp.path().join("docwarp_backups"), "out.docx")?;
    assert!(
        backups.is_empty(),
        "expected no backup files when --no-backup is set, found {}",
        backups.len()
    );

    Ok(())
}

#[test]
fn backup_retention_keeps_latest_n_files() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("input.md");
    let output = temp.path().join("out.docx");
    let backup_dir = temp.path().join("custom-backups");

    fs::write(&input, "# Title\n\nv0\n").context("failed writing markdown v0")?;
    let setup = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&setup, Some(0), "setup md2docx run should succeed")?;

    for idx in 1..=3 {
        fs::write(&input, format!("# Title\n\nv{idx}\n"))
            .with_context(|| format!("failed writing markdown v{idx}"))?;
        let run = run_docwarp([
            "md2docx",
            input.to_string_lossy().as_ref(),
            "--output",
            output.to_string_lossy().as_ref(),
            "--backup-dir",
            backup_dir.to_string_lossy().as_ref(),
            "--backup-keep",
            "2",
        ])?;
        assert_command_status(
            &run,
            Some(0),
            &format!("overwrite run {idx} with retention should succeed"),
        )?;
        thread::sleep(Duration::from_millis(1));
    }

    let backups = collect_backup_files(&backup_dir, "out.docx")?;
    assert_eq!(
        backups.len(),
        2,
        "expected retention to keep only 2 backups, found {} in {}",
        backups.len(),
        backup_dir.display()
    );

    Ok(())
}

#[test]
fn invalid_style_map_returns_fatal_error() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("input.md");
    let output = temp.path().join("out.docx");
    let style_map = temp.path().join("invalid-style-map.yml");

    fs::write(&input, "# Title\n\nBody\n").context("failed writing markdown")?;
    fs::write(&style_map, "docx_to_md: [\n").context("failed writing invalid style-map fixture")?;

    let run = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
        "--style-map",
        style_map.to_string_lossy().as_ref(),
    ])?;

    assert_command_status(&run, Some(1), "invalid style map should fail")?;
    assert!(
        stderr_text(&run).contains("invalid YAML style map"),
        "expected invalid YAML style map error, got:\n{}",
        stderr_text(&run)
    );

    Ok(())
}

#[test]
fn invalid_style_map_token_reports_actionable_diagnostic() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("input.md");
    let output = temp.path().join("out.docx");
    let style_map = temp.path().join("invalid-style-map-token.yml");

    fs::write(&input, "# Title\n\nBody\n").context("failed writing markdown")?;
    fs::write(&style_map, "md_to_docx:\n  paragrph: Normal\n")
        .context("failed writing invalid style-map token fixture")?;

    let run = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
        "--style-map",
        style_map.to_string_lossy().as_ref(),
    ])?;

    assert_command_status(&run, Some(1), "invalid style map token should fail")?;
    let stderr = stderr_text(&run);
    assert!(
        stderr.contains("md_to_docx.paragrph"),
        "expected style-map entry path diagnostic, got:\n{stderr}"
    );
    assert!(
        stderr.contains("Did you mean `paragraph`?"),
        "expected style-map token suggestion, got:\n{stderr}"
    );

    Ok(())
}

#[test]
fn invalid_template_in_strict_mode_exits_with_2_and_warning() -> Result<()> {
    let temp = tempdir().context("tempdir should be created")?;
    let input = temp.path().join("input.md");
    let output = temp.path().join("out.docx");

    fs::write(&input, "# Title\n\nBody\n").context("failed writing markdown")?;

    let run = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        output.to_string_lossy().as_ref(),
        "--template",
        temp.path()
            .join("missing-template.dotx")
            .to_string_lossy()
            .as_ref(),
        "--strict",
    ])?;

    assert_command_status(&run, Some(2), "invalid template should warn under strict")?;
    assert!(
        stdout_text(&run).contains("[invalid_template]"),
        "expected invalid_template warning, got:\n{}",
        stdout_text(&run)
    );

    Ok(())
}

#[test]
fn warning_catalog_docs_cover_all_codes() -> Result<()> {
    let root = workspace_root();
    let doc = fs::read_to_string(root.join("docs/warnings.md"))
        .context("failed reading warning-code docs")?;

    for code in WarningCode::ALL {
        let needle = format!("`{}`", code.as_str());
        assert!(doc.contains(&needle), "warning docs missing code {needle}");
    }

    Ok(())
}

#[test]
fn md_to_docx_output_is_deterministic() -> Result<()> {
    let root = workspace_root();
    let temp = tempdir().context("tempdir should be created")?;

    let input = root.join("fixtures/md/10-comprehensive.md");
    let out_a = temp.path().join("a.docx");
    let out_b = temp.path().join("b.docx");

    let first = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        out_a.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&first, Some(0), "first deterministic md2docx run")?;

    let second = run_docwarp([
        "md2docx",
        input.to_string_lossy().as_ref(),
        "--output",
        out_b.to_string_lossy().as_ref(),
    ])?;
    assert_command_status(&second, Some(0), "second deterministic md2docx run")?;

    let bytes_a = fs::read(&out_a).context("failed reading deterministic output a")?;
    let bytes_b = fs::read(&out_b).context("failed reading deterministic output b")?;
    assert_eq!(bytes_a, bytes_b, "md2docx output bytes should be stable");

    Ok(())
}

fn workspace_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .canonicalize()
        .expect("workspace root should be resolvable")
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

fn collect_backup_files(backup_dir: &Path, output_file_name: &str) -> Result<Vec<PathBuf>> {
    if !backup_dir.exists() {
        return Ok(Vec::new());
    }

    let output_path = Path::new(output_file_name);
    let output_name = output_path
        .file_name()
        .and_then(OsStr::to_str)
        .unwrap_or(output_file_name);
    let expected_stem = output_path
        .file_stem()
        .and_then(OsStr::to_str)
        .unwrap_or(output_name);
    let expected_extension_suffix = output_path
        .extension()
        .and_then(OsStr::to_str)
        .filter(|value| !value.is_empty())
        .map(|value| format!(".{value}"))
        .unwrap_or_default();
    let expected_prefix = format!("{expected_stem}--");
    let mut backups = Vec::new();
    for entry in fs::read_dir(backup_dir)
        .with_context(|| format!("failed reading backup directory: {}", backup_dir.display()))?
    {
        let entry = entry.with_context(|| {
            format!(
                "failed reading backup directory entry: {}",
                backup_dir.display()
            )
        })?;
        if !entry
            .file_type()
            .with_context(|| {
                format!(
                    "failed reading backup file type: {}",
                    entry.path().display()
                )
            })?
            .is_file()
        {
            continue;
        }

        let path = entry.path();
        let Some(file_name) = path.file_name().and_then(OsStr::to_str) else {
            continue;
        };
        if file_name.starts_with(&expected_prefix)
            && file_name.ends_with(&expected_extension_suffix)
        {
            backups.push(path);
        }
    }

    backups.sort();
    Ok(backups)
}

fn normalize_document_semantics(document: &mut Document) {
    for block in &mut document.blocks {
        match block {
            Block::Title(inlines)
            | Block::Paragraph(inlines)
            | Block::BlockQuote(inlines)
            | Block::Heading {
                content: inlines, ..
            } => normalize_inline_semantics(inlines),
            Block::List { items, .. } => {
                for item in items {
                    normalize_inline_semantics(item);
                }
            }
            Block::Table { headers, rows } => {
                for cell in headers {
                    normalize_inline_semantics(cell);
                }
                for row in rows {
                    for cell in row {
                        normalize_inline_semantics(cell);
                    }
                }
            }
            Block::Image { alt, src, title } => {
                normalize_image_path(src);
                alt.clear();
                *title = None;
            }
            Block::CodeBlock { language, .. } => {
                *language = None;
            }
            Block::ThematicBreak => {}
        }
    }

    // Treat legacy Word Title style as semantic Heading 1.
    for block in &mut document.blocks {
        let content = match block {
            Block::Title(inlines) => Some(inlines.clone()),
            _ => None,
        };
        if let Some(content) = content {
            *block = Block::Heading { level: 1, content };
        }
    }
}

fn normalize_inline_semantics(inlines: &mut [Inline]) {
    for inline in inlines {
        match inline {
            Inline::Image { alt, src, title } => {
                normalize_image_path(src);
                alt.clear();
                *title = None;
            }
            Inline::Emphasis(children)
            | Inline::Strong(children)
            | Inline::Link { text: children, .. } => normalize_inline_semantics(children),
            Inline::Text(_) | Inline::Equation { .. } | Inline::Code(_) | Inline::LineBreak => {}
        }
    }
}

fn normalize_image_path(src: &mut String) {
    if src.starts_with("http://") || src.starts_with("https://") {
        return;
    }

    if Path::new(src).file_name().and_then(OsStr::to_str).is_some() {
        *src = "__local_image__".to_string();
    }
}

fn remove_docx_entry(source: &Path, destination: &Path, removed_entry_name: &str) -> Result<()> {
    let source_file = fs::File::open(source)
        .with_context(|| format!("failed opening source DOCX: {}", source.display()))?;
    let mut archive = ZipArchive::new(source_file)
        .with_context(|| format!("failed reading source DOCX archive: {}", source.display()))?;

    let destination_file = fs::File::create(destination).with_context(|| {
        format!(
            "failed creating destination DOCX: {}",
            destination.display()
        )
    })?;
    let mut writer = ZipWriter::new(destination_file);

    let mut removed = false;

    for idx in 0..archive.len() {
        let mut file = archive
            .by_index(idx)
            .with_context(|| format!("failed reading source zip entry index {idx}"))?;
        let name = file.name().to_string();

        if name == removed_entry_name {
            removed = true;
            continue;
        }

        if file.is_dir() {
            writer
                .add_directory(name, SimpleFileOptions::default())
                .context("failed writing directory entry")?;
            continue;
        }

        writer
            .start_file(name, SimpleFileOptions::default())
            .context("failed writing file entry header")?;
        io::copy(&mut file, &mut writer).context("failed copying zip entry bytes")?;
    }

    writer
        .finish()
        .context("failed finalizing modified DOCX archive")?;

    if !removed {
        bail!(
            "entry '{}' was not found in source DOCX {}",
            removed_entry_name,
            source.display()
        );
    }

    Ok(())
}
