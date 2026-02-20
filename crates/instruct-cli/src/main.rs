use std::fs;
use std::path::{Path, PathBuf};
use std::process;
use std::time::Instant;

use anyhow::{Context, Result, anyhow};
use clap::{Parser, Subcommand};
use instruct_core::{
    AppConfig, ConversionDirection, ConversionReport, ConversionWarning, Document, StyleMap,
    UnsupportedPolicy, WarningCode, resolve_style_map, style_map,
};
use instruct_docx::{DocxReadOptions, DocxWriteOptions, read_docx, write_docx};
use instruct_md::{parse_markdown, render_markdown};

#[derive(Debug, Parser)]
#[command(name = "instruct")]
#[command(about = "Convert documentation between Markdown and DOCX")]
struct Cli {
    #[command(subcommand)]
    command: Commands,
}

#[derive(Debug, Subcommand)]
enum Commands {
    /// Convert Markdown to DOCX
    Md2docx {
        input: PathBuf,
        #[arg(short, long)]
        output: PathBuf,
        #[arg(long)]
        template: Option<PathBuf>,
        #[arg(long = "style-map")]
        style_map: Option<PathBuf>,
        #[arg(long)]
        config: Option<PathBuf>,
        #[arg(long)]
        report: Option<PathBuf>,
        #[arg(long)]
        strict: bool,
        #[arg(long = "allow-remote-images")]
        allow_remote_images: bool,
    },
    /// Convert DOCX to Markdown
    Docx2md {
        input: PathBuf,
        #[arg(short, long)]
        output: PathBuf,
        #[arg(long = "assets-dir")]
        assets_dir: Option<PathBuf>,
        #[arg(long = "style-map")]
        style_map: Option<PathBuf>,
        #[arg(long)]
        config: Option<PathBuf>,
        #[arg(long)]
        report: Option<PathBuf>,
        #[arg(long)]
        strict: bool,
    },
}

fn main() {
    let cli = Cli::parse();

    let exit_code = match run(cli) {
        Ok(code) => code,
        Err(err) => {
            eprintln!("error: {err:?}");
            1
        }
    };

    process::exit(exit_code);
}

fn run(cli: Cli) -> Result<i32> {
    match cli.command {
        Commands::Md2docx {
            input,
            output,
            template,
            style_map,
            config,
            report,
            strict,
            allow_remote_images,
        } => run_md2docx(
            input,
            output,
            template,
            style_map,
            config,
            report,
            strict,
            allow_remote_images,
        ),
        Commands::Docx2md {
            input,
            output,
            assets_dir,
            style_map,
            config,
            report,
            strict,
        } => run_docx2md(input, output, assets_dir, style_map, config, report, strict),
    }
}

fn run_md2docx(
    input: PathBuf,
    output: PathBuf,
    template: Option<PathBuf>,
    cli_style_map_path: Option<PathBuf>,
    config_path: Option<PathBuf>,
    report_path: Option<PathBuf>,
    strict_flag: bool,
    allow_remote_images: bool,
) -> Result<i32> {
    let started = Instant::now();
    let (config_file, config) = load_config_with_auto_discovery(config_path.as_deref())?;

    ensure_exists(&input)?;

    let input_data = fs::read_to_string(&input)
        .with_context(|| format!("failed reading markdown input: {}", input.display()))?;

    let (document, mut warnings) = parse_markdown(&input_data)?;

    let style_map = load_effective_style_map(
        config_file.as_deref(),
        config.style_map.as_deref(),
        cli_style_map_path.as_deref(),
    )?;

    let template_path =
        resolve_optional_configured_path(config_file.as_deref(), template).or_else(|| {
            resolve_optional_configured_path(
                config_file.as_deref(),
                config.default_template.clone(),
            )
        });

    let input_base = input
        .parent()
        .map(Path::to_path_buf)
        .unwrap_or_else(|| PathBuf::from("."));

    let mut write_warnings = write_docx(
        &document,
        &input_base,
        &output,
        &DocxWriteOptions {
            allow_remote_images,
            style_map,
            template: template_path,
        },
    )?;

    warnings.append(&mut write_warnings);

    let policy = config.unsupported_policy_or_default();
    let strict = strict_mode(policy, strict_flag);
    let duration = started.elapsed().as_millis();

    emit_summary(
        &ConversionDirection::MdToDocx,
        &input,
        &output,
        &warnings,
        strict,
    );
    write_report_if_requested(
        report_path.as_deref(),
        ConversionDirection::MdToDocx,
        &input,
        &output,
        duration,
        &document,
        &warnings,
    )?;

    Ok(exit_code_from_warnings(&warnings, strict))
}

fn run_docx2md(
    input: PathBuf,
    output: PathBuf,
    cli_assets_dir: Option<PathBuf>,
    cli_style_map_path: Option<PathBuf>,
    config_path: Option<PathBuf>,
    report_path: Option<PathBuf>,
    strict_flag: bool,
) -> Result<i32> {
    let started = Instant::now();
    let (config_file, config) = load_config_with_auto_discovery(config_path.as_deref())?;

    ensure_exists(&input)?;

    let output_parent = output
        .parent()
        .map(Path::to_path_buf)
        .unwrap_or_else(|| PathBuf::from("."));

    let configured_assets_dir = cli_assets_dir
        .or(config.assets_dir.clone())
        .unwrap_or_else(|| default_assets_dir_for_output(&output));

    let assets_dir = resolve_output_relative_path(&output_parent, configured_assets_dir);

    let style_map = load_effective_style_map(
        config_file.as_deref(),
        config.style_map.as_deref(),
        cli_style_map_path.as_deref(),
    )?;

    let (mut document, warnings) = read_docx(
        &input,
        &DocxReadOptions {
            assets_dir: assets_dir.clone(),
            style_map,
        },
    )?;

    rewrite_image_paths_relative_to_output(&mut document, &output_parent);

    let markdown = render_markdown(&document);
    if let Some(parent) = output.parent() {
        fs::create_dir_all(parent)
            .with_context(|| format!("failed creating output directory: {}", parent.display()))?;
    }
    fs::write(&output, markdown)
        .with_context(|| format!("failed writing markdown output: {}", output.display()))?;

    let policy = config.unsupported_policy_or_default();
    let strict = strict_mode(policy, strict_flag);
    let duration = started.elapsed().as_millis();

    emit_summary(
        &ConversionDirection::DocxToMd,
        &input,
        &output,
        &warnings,
        strict,
    );
    write_report_if_requested(
        report_path.as_deref(),
        ConversionDirection::DocxToMd,
        &input,
        &output,
        duration,
        &document,
        &warnings,
    )?;

    Ok(exit_code_from_warnings(&warnings, strict))
}

fn ensure_exists(path: &Path) -> Result<()> {
    if path.exists() {
        Ok(())
    } else {
        Err(anyhow!("input file does not exist: {}", path.display()))
    }
}

fn load_config_with_auto_discovery(path: Option<&Path>) -> Result<(Option<PathBuf>, AppConfig)> {
    if let Some(path) = path {
        return Ok((Some(path.to_path_buf()), AppConfig::load(path)?));
    }

    let default_path = PathBuf::from(".instruct.yml");
    if default_path.exists() {
        Ok((Some(default_path.clone()), AppConfig::load(&default_path)?))
    } else {
        Ok((None, AppConfig::default()))
    }
}

fn load_effective_style_map(
    config_file: Option<&Path>,
    config_style_map: Option<&Path>,
    cli_style_map: Option<&Path>,
) -> Result<StyleMap> {
    let config_style_map = match config_style_map {
        Some(path) => Some(style_map::load_style_map(&resolve_path_from_config(
            config_file,
            path,
        ))?),
        None => None,
    };

    let cli_style_map = match cli_style_map {
        Some(path) => Some(style_map::load_style_map(path)?),
        None => None,
    };

    Ok(resolve_style_map(config_style_map, cli_style_map))
}

fn resolve_path_from_config(config_file: Option<&Path>, configured: &Path) -> PathBuf {
    if configured.is_absolute() {
        return configured.to_path_buf();
    }

    if let Some(config_file) = config_file {
        if let Some(parent) = config_file.parent() {
            return parent.join(configured);
        }
    }

    configured.to_path_buf()
}

fn resolve_optional_configured_path(
    config_file: Option<&Path>,
    configured: Option<PathBuf>,
) -> Option<PathBuf> {
    configured.map(|value| {
        if value.is_absolute() {
            value
        } else if let Some(config_file) = config_file {
            config_file
                .parent()
                .map(|parent| parent.join(&value))
                .unwrap_or(value)
        } else {
            value
        }
    })
}

fn strict_mode(policy: UnsupportedPolicy, strict_flag: bool) -> bool {
    strict_flag || matches!(policy, UnsupportedPolicy::FailFast)
}

fn exit_code_from_warnings(warnings: &[ConversionWarning], strict: bool) -> i32 {
    if strict && !warnings.is_empty() { 2 } else { 0 }
}

fn emit_summary(
    direction: &ConversionDirection,
    input: &Path,
    output: &Path,
    warnings: &[ConversionWarning],
    strict: bool,
) {
    println!(
        "{} completed: {} -> {}",
        match direction {
            ConversionDirection::MdToDocx => "md2docx",
            ConversionDirection::DocxToMd => "docx2md",
        },
        input.display(),
        output.display()
    );

    if warnings.is_empty() {
        println!("warnings: 0");
        return;
    }

    println!("warnings: {}", warnings.len());
    for warning in warnings {
        let code = match warning.code {
            WarningCode::UnsupportedFeature => "unsupported_feature",
            WarningCode::ImageLoadFailed => "image_load_failed",
            WarningCode::RemoteImageBlocked => "remote_image_blocked",
            WarningCode::MissingMedia => "missing_media",
            WarningCode::InvalidStyleMap => "invalid_style_map",
            WarningCode::InvalidTemplate => "invalid_template",
            WarningCode::CorruptDocx => "corrupt_docx",
            WarningCode::NestedStructureSimplified => "nested_structure_simplified",
        };
        if let Some(location) = &warning.location {
            println!("- [{code}] {} ({location})", warning.message);
        } else {
            println!("- [{code}] {}", warning.message);
        }
    }

    if strict {
        println!("strict mode enabled: warnings will produce exit code 2");
    }
}

fn write_report_if_requested(
    report_path: Option<&Path>,
    direction: ConversionDirection,
    input: &Path,
    output: &Path,
    duration_ms: u128,
    document: &Document,
    warnings: &[ConversionWarning],
) -> Result<()> {
    let Some(report_path) = report_path else {
        return Ok(());
    };

    let report = ConversionReport::new(
        direction,
        input.display().to_string(),
        output.display().to_string(),
        duration_ms,
        document.stats(),
        warnings.to_vec(),
        true,
    );

    report.write_to_path(report_path)
}

fn default_assets_dir_for_output(output_markdown: &Path) -> PathBuf {
    let stem = output_markdown
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("assets");
    PathBuf::from(format!("{stem}_assets"))
}

fn resolve_output_relative_path(output_parent: &Path, path: PathBuf) -> PathBuf {
    if path.is_absolute() {
        path
    } else {
        output_parent.join(path)
    }
}

fn rewrite_image_paths_relative_to_output(document: &mut Document, output_parent: &Path) {
    for block in &mut document.blocks {
        match block {
            instruct_core::Block::Paragraph(inlines)
            | instruct_core::Block::BlockQuote(inlines)
            | instruct_core::Block::Title(inlines)
            | instruct_core::Block::Heading {
                content: inlines, ..
            } => {
                rewrite_inline_paths(inlines, output_parent);
            }
            instruct_core::Block::List { items, .. } => {
                for item in items {
                    rewrite_inline_paths(item, output_parent);
                }
            }
            instruct_core::Block::Table { headers, rows } => {
                for cell in headers {
                    rewrite_inline_paths(cell, output_parent);
                }
                for row in rows {
                    for cell in row {
                        rewrite_inline_paths(cell, output_parent);
                    }
                }
            }
            instruct_core::Block::Image { src, .. } => {
                if let Ok(rel) = make_relative_if_absolute(src, output_parent) {
                    *src = rel;
                }
            }
            instruct_core::Block::CodeBlock { .. } | instruct_core::Block::ThematicBreak => {}
        }
    }
}

fn rewrite_inline_paths(inlines: &mut [instruct_core::Inline], output_parent: &Path) {
    for inline in inlines {
        match inline {
            instruct_core::Inline::Image { src, .. } => {
                if let Ok(rel) = make_relative_if_absolute(src, output_parent) {
                    *src = rel;
                }
            }
            instruct_core::Inline::Emphasis(children)
            | instruct_core::Inline::Strong(children)
            | instruct_core::Inline::Link { text: children, .. } => {
                rewrite_inline_paths(children, output_parent)
            }
            instruct_core::Inline::Text(_)
            | instruct_core::Inline::Code(_)
            | instruct_core::Inline::LineBreak => {}
        }
    }
}

fn make_relative_if_absolute(path: &str, output_parent: &Path) -> Result<String> {
    let as_path = Path::new(path);
    if !as_path.is_absolute() {
        return Ok(path.to_string());
    }

    let relative = as_path
        .strip_prefix(output_parent)
        .map(|p| p.to_string_lossy().replace('\\', "/"))
        .with_context(|| {
            format!(
                "absolute image path is outside markdown output directory: {}",
                as_path.display()
            )
        })?;

    Ok(relative)
}
