use std::fs;
use std::io::{self, IsTerminal, Write};
use std::path::{Path, PathBuf};
use std::process;
use std::time::Instant;

use anyhow::{Context, Result, anyhow};
use clap::{Parser, Subcommand};
use glob::Pattern;
use instruct_core::{
    AppConfig, ConversionDirection, ConversionReport, ConversionWarning, Document, StyleMap,
    UnsupportedPolicy, resolve_style_map, style_map,
};
use instruct_docx::{DocxReadOptions, DocxWriteOptions, read_docx, write_docx};
use instruct_md::{parse_markdown, render_markdown};
use walkdir::WalkDir;

const CLI_LONG_ABOUT: &str = "Convert documentation between Markdown and DOCX.\n\
\n\
Run without arguments to use guided mode.\n\
\n\
The tool also supports directional subcommands:\n\
- md2docx: convert Markdown into DOCX\n\
- docx2md: convert DOCX into Markdown";

const CLI_AFTER_LONG_HELP: &str = "Examples:\n\
  instruct\n\
  instruct md2docx ./docs/spec.md --output ./build/spec.docx\n\
  instruct md2docx ./docs/spec.md --output ./build/spec.docx --strict --report ./build/report.json\n\
  instruct docx2md ./contracts/master.docx --output ./contracts/master.md --assets-dir ./contracts/assets\n\
\n\
Run command-specific help for detailed examples:\n\
  instruct md2docx --help\n\
  instruct docx2md --help";

const MD2DOCX_AFTER_LONG_HELP: &str = "Examples:\n\
  instruct md2docx ./input.md --output ./output.docx\n\
  instruct md2docx ./input.md --output ./output.docx --template ./brand.dotx --style-map ./style-map.yml\n\
  instruct md2docx ./docs --output ./build/docx --glob \"**/*.md\"\n\
  instruct md2docx ./input.md --output ./output.docx --config ./.instruct.yml\n\
  instruct md2docx ./input.md --output ./output.docx --report ./report.json --strict\n\
  instruct md2docx ./input.md --output ./output.docx --allow-remote-images";

const DOCX2MD_AFTER_LONG_HELP: &str = "Examples:\n\
  instruct docx2md ./input.docx --output ./output.md\n\
  instruct docx2md ./input.docx --output ./output.md --assets-dir ./output_assets\n\
  instruct docx2md ./contracts --output ./build/md --glob \"**/*.docx\"\n\
  instruct docx2md ./input.docx --output ./output.md --style-map ./style-map.json\n\
  instruct docx2md ./input.docx --output ./output.md --config ./.instruct.yml --report ./report.json\n\
  instruct docx2md ./input.docx --output ./output.md --strict";

#[derive(Debug, Parser)]
#[command(name = "instruct")]
#[command(about = "Convert documentation between Markdown and DOCX")]
#[command(long_about = CLI_LONG_ABOUT)]
#[command(after_long_help = CLI_AFTER_LONG_HELP)]
struct Cli {
    #[command(subcommand)]
    command: Option<Commands>,
}

#[derive(Debug, Subcommand)]
enum Commands {
    /// Convert Markdown to DOCX
    #[command(after_long_help = MD2DOCX_AFTER_LONG_HELP)]
    Md2docx {
        input: PathBuf,
        #[arg(short, long)]
        output: PathBuf,
        #[arg(
            long,
            value_name = "PATTERN",
            help = "Enable batch mode for directory input and filter files by glob pattern (for example, \"**/*.md\")"
        )]
        glob: Option<String>,
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
    #[command(after_long_help = DOCX2MD_AFTER_LONG_HELP)]
    Docx2md {
        input: PathBuf,
        #[arg(short, long)]
        output: PathBuf,
        #[arg(
            long,
            value_name = "PATTERN",
            help = "Enable batch mode for directory input and filter files by glob pattern (for example, \"**/*.docx\")"
        )]
        glob: Option<String>,
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
    emit_startup_header();

    match cli.command {
        Some(Commands::Md2docx {
            input,
            output,
            glob,
            template,
            style_map,
            config,
            report,
            strict,
            allow_remote_images,
        }) => run_md2docx(
            input,
            output,
            glob,
            template,
            style_map,
            config,
            report,
            strict,
            allow_remote_images,
        ),
        Some(Commands::Docx2md {
            input,
            output,
            glob,
            assets_dir,
            style_map,
            config,
            report,
            strict,
        }) => run_docx2md(
            input, output, glob, assets_dir, style_map, config, report, strict,
        ),
        None => run_guided_mode(),
    }
}

fn emit_startup_header() {
    let version = env!("CARGO_PKG_VERSION");
    let title = format!("instruct (v{version})");
    let width = title.chars().count();
    let horizontal = "─".repeat(width + 2);

    println!("╭{horizontal}╮");
    println!("│ {title} │");
    println!("╰{horizontal}╯");
    println!();
}

#[derive(Debug, Clone, Copy)]
enum GuidedDirection {
    MdToDocx,
    DocxToMd,
}

#[derive(Debug)]
enum NativePickerOutcome {
    Selected(PathBuf),
    Cancelled,
    Unavailable,
}

impl GuidedDirection {
    fn label(self) -> &'static str {
        match self {
            GuidedDirection::MdToDocx => "Markdown -> DOCX",
            GuidedDirection::DocxToMd => "DOCX -> Markdown",
        }
    }

    fn output_extension(self) -> &'static str {
        match self {
            GuidedDirection::MdToDocx => "docx",
            GuidedDirection::DocxToMd => "md",
        }
    }
}

fn run_guided_mode() -> Result<i32> {
    println!("Drop a Markdown or DOCX file/folder path and press Enter.");
    println!("Press Enter to open the file picker.");

    let input = prompt_for_input_path()?;
    let direction = detect_guided_direction(&input)?;
    let default_output = default_guided_output_path(&input, direction);
    let output = prompt_for_output_path(&input, default_output, direction)?;

    println!();
    println!("Converting: {}", direction.label());
    println!("input:  {}", input.display());
    println!("output: {}", output.display());
    println!();

    match direction {
        GuidedDirection::MdToDocx => {
            run_md2docx(input, output, None, None, None, None, None, false, false)
        }
        GuidedDirection::DocxToMd => {
            run_docx2md(input, output, None, None, None, None, None, false)
        }
    }
}

fn prompt_for_input_path() -> Result<PathBuf> {
    loop {
        let raw = prompt_line("Input path: ")?;
        let normalized = raw.trim();
        if normalized.is_empty() {
            if should_offer_native_picker() {
                match pick_path_with_native_explorer()? {
                    NativePickerOutcome::Selected(path) => return Ok(path),
                    NativePickerOutcome::Cancelled => continue,
                    NativePickerOutcome::Unavailable => return browse_for_path(),
                }
            }

            return browse_for_path();
        }

        let input = parse_user_path(&raw);
        if input.exists() {
            return Ok(input);
        }

        eprintln!("path does not exist: {}", input.display());
    }
}

fn should_offer_native_picker() -> bool {
    io::stdin().is_terminal() && io::stdout().is_terminal()
}

fn pick_path_with_native_explorer() -> Result<NativePickerOutcome> {
    #[cfg(target_os = "macos")]
    {
        let script = r#"
ObjC.import('AppKit');

function run() {
    const app = $.NSApplication.sharedApplication;
    app.setActivationPolicy($.NSApplicationActivationPolicyRegular);
    $.NSApp.activateIgnoringOtherApps(true);
    $.NSRunningApplication.currentApplication.activateWithOptions($.NSApplicationActivateIgnoringOtherApps);

    const panel = $.NSOpenPanel.openPanel;
    panel.setCanChooseFiles(true);
    panel.setCanChooseDirectories(true);
    panel.setAllowsMultipleSelection(false);
    panel.setCanCreateDirectories(false);
    panel.setPrompt("Select");
    panel.setMessage("Choose a Markdown file, DOCX file, or folder");

    const response = panel.runModal;
    if (response !== $.NSModalResponseOK) {
        return "";
    }

    return ObjC.unwrap(panel.URL.path);
}
"#;

        let output = match process::Command::new("osascript")
            .arg("-l")
            .arg("JavaScript")
            .arg("-e")
            .arg(script)
            .output()
        {
            Ok(output) => output,
            Err(err) if err.kind() == io::ErrorKind::NotFound => {
                return Ok(NativePickerOutcome::Unavailable);
            }
            Err(err) => {
                return Err(err).context("failed launching native file explorer dialog");
            }
        };

        if !output.status.success() {
            let stderr = String::from_utf8_lossy(&output.stderr);
            eprintln!(
                "native picker unavailable (osascript error), falling back to terminal browser: {}",
                stderr.trim()
            );
            return Ok(NativePickerOutcome::Unavailable);
        }

        let path_text = String::from_utf8_lossy(&output.stdout).trim().to_string();
        if path_text.is_empty() {
            return Ok(NativePickerOutcome::Cancelled);
        }

        let selected = PathBuf::from(path_text);
        if !selected.exists() {
            return Ok(NativePickerOutcome::Cancelled);
        }

        return Ok(NativePickerOutcome::Selected(selected));
    }

    #[cfg(not(target_os = "macos"))]
    {
        Ok(NativePickerOutcome::Unavailable)
    }
}

fn prompt_for_output_path(
    input: &Path,
    default_output: PathBuf,
    direction: GuidedDirection,
) -> Result<PathBuf> {
    let prompt = format!("Output path [{}]: ", default_output.display());
    let raw = prompt_line(&prompt)?;
    if raw.trim().is_empty() {
        return Ok(default_output);
    }

    Ok(normalize_file_output_path(
        input,
        parse_user_path(&raw),
        direction,
    ))
}

fn browse_for_path() -> Result<PathBuf> {
    let mut current = std::env::current_dir().context("failed reading current directory")?;

    loop {
        let mut entries = Vec::new();
        for entry in fs::read_dir(&current)
            .with_context(|| format!("failed reading directory {}", current.display()))?
        {
            let entry = entry?;
            let path = entry.path();
            let metadata = entry.metadata()?;
            let is_dir = metadata.is_dir();
            let name = entry.file_name().to_string_lossy().to_string();
            entries.push((is_dir, name, path));
        }

        entries.sort_by(|a, b| match (a.0, b.0) {
            (true, false) => std::cmp::Ordering::Less,
            (false, true) => std::cmp::Ordering::Greater,
            _ => a.1.to_lowercase().cmp(&b.1.to_lowercase()),
        });

        println!();
        println!("Browsing: {}", current.display());
        println!("0) Use this directory");
        println!("u) Go up one directory");
        println!("p) Paste path manually");
        println!("q) Cancel");
        for (index, (is_dir, name, _)) in entries.iter().enumerate() {
            let marker = if *is_dir { "/" } else { "" };
            println!("{}) {}{}", index + 1, name, marker);
        }

        let choice = prompt_line("Select item: ")?;
        let normalized = choice.trim();
        if normalized.eq_ignore_ascii_case("q") {
            return Err(anyhow!("guided mode cancelled"));
        }
        if normalized.eq_ignore_ascii_case("u") {
            if let Some(parent) = current.parent() {
                current = parent.to_path_buf();
            }
            continue;
        }
        if normalized.eq_ignore_ascii_case("p") {
            let raw = prompt_line("Path: ")?;
            if raw.trim().is_empty() {
                continue;
            }
            let path = parse_user_path(&raw);
            if !path.exists() {
                eprintln!("path does not exist: {}", path.display());
                continue;
            }
            return Ok(path);
        }
        if normalized == "0" {
            return Ok(current.clone());
        }

        let index: usize = match normalized.parse() {
            Ok(value) => value,
            Err(_) => {
                eprintln!("invalid selection: {normalized}");
                continue;
            }
        };
        if index == 0 {
            eprintln!("invalid selection: {normalized}");
            continue;
        }
        let Some((is_dir, _, selected)) = entries.get(index.saturating_sub(1)) else {
            eprintln!("selection out of range: {index}");
            continue;
        };

        if *is_dir {
            current = selected.clone();
        } else {
            return Ok(selected.clone());
        }
    }
}

fn detect_guided_direction(input: &Path) -> Result<GuidedDirection> {
    if input.is_file() {
        return detect_direction_from_file(input);
    }
    if input.is_dir() {
        return detect_direction_from_directory(input);
    }

    Err(anyhow!(
        "input path is neither file nor directory: {}",
        input.display()
    ))
}

fn detect_direction_from_file(path: &Path) -> Result<GuidedDirection> {
    let extension = path
        .extension()
        .and_then(|value| value.to_str())
        .map(|value| value.to_ascii_lowercase())
        .ok_or_else(|| {
            anyhow!(
                "unsupported input file type (missing extension): {}",
                path.display()
            )
        })?;

    match extension.as_str() {
        "md" | "markdown" | "mdown" | "mkd" => Ok(GuidedDirection::MdToDocx),
        "docx" => Ok(GuidedDirection::DocxToMd),
        _ => Err(anyhow!(
            "unsupported input file extension '.{}' for {} (supported: .md, .markdown, .docx)",
            extension,
            path.display()
        )),
    }
}

fn detect_direction_from_directory(path: &Path) -> Result<GuidedDirection> {
    let mut md_count = 0usize;
    let mut docx_count = 0usize;

    for entry in WalkDir::new(path) {
        let entry = entry.with_context(|| {
            format!(
                "failed traversing directory while detecting file type: {}",
                path.display()
            )
        })?;
        if !entry.file_type().is_file() {
            continue;
        }
        if let Some(extension) = entry.path().extension().and_then(|value| value.to_str()) {
            if extension.eq_ignore_ascii_case("docx") {
                docx_count += 1;
            } else if matches!(
                extension.to_ascii_lowercase().as_str(),
                "md" | "markdown" | "mdown" | "mkd"
            ) {
                md_count += 1;
            }
        }
    }

    match (md_count, docx_count) {
        (0, 0) => Err(anyhow!(
            "directory has no convertible files (.md/.markdown/.docx): {}",
            path.display()
        )),
        (_, 0) => Ok(GuidedDirection::MdToDocx),
        (0, _) => Ok(GuidedDirection::DocxToMd),
        _ => prompt_direction_for_mixed_directory(md_count, docx_count),
    }
}

fn prompt_direction_for_mixed_directory(
    md_count: usize,
    docx_count: usize,
) -> Result<GuidedDirection> {
    let default = if md_count >= docx_count {
        GuidedDirection::MdToDocx
    } else {
        GuidedDirection::DocxToMd
    };

    println!();
    println!("Found both Markdown and DOCX files:");
    println!("Markdown files: {md_count}");
    println!("DOCX files: {docx_count}");
    println!(
        "Select conversion direction [1/2] (default {}):",
        if matches!(default, GuidedDirection::MdToDocx) {
            "1"
        } else {
            "2"
        }
    );
    println!("1) Markdown -> DOCX");
    println!("2) DOCX -> Markdown");

    loop {
        let choice = prompt_line("Choice: ")?;
        let normalized = choice.trim();
        if normalized.is_empty() {
            return Ok(default);
        }
        match normalized {
            "1" => return Ok(GuidedDirection::MdToDocx),
            "2" => return Ok(GuidedDirection::DocxToMd),
            _ => eprintln!("invalid choice: {normalized}"),
        }
    }
}

fn default_guided_output_path(input: &Path, direction: GuidedDirection) -> PathBuf {
    if input.is_file() {
        let mut output = input.to_path_buf();
        output.set_extension(direction.output_extension());
        return output;
    }

    let parent = input.parent().unwrap_or_else(|| Path::new("."));
    let base = input
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or("output");
    let suffix = direction.output_extension();

    let mut candidate = parent.join(format!("{base}-{suffix}"));
    let mut index = 2usize;
    while candidate.exists() {
        candidate = parent.join(format!("{base}-{suffix}-{index}"));
        index += 1;
    }
    candidate
}

fn normalize_file_output_path(
    input: &Path,
    output: PathBuf,
    direction: GuidedDirection,
) -> PathBuf {
    if !input.is_file() {
        return output;
    }

    if output.exists() && output.is_dir() {
        let stem = input
            .file_stem()
            .and_then(|value| value.to_str())
            .unwrap_or("output");
        return output.join(format!("{stem}.{}", direction.output_extension()));
    }

    if output.extension().is_none() {
        let mut patched = output;
        patched.set_extension(direction.output_extension());
        return patched;
    }

    output
}

fn prompt_line(prompt: &str) -> Result<String> {
    print!("{prompt}");
    io::stdout().flush().context("failed flushing stdout")?;

    let mut line = String::new();
    let bytes = io::stdin()
        .read_line(&mut line)
        .context("failed reading stdin")?;
    if bytes == 0 {
        return Err(anyhow!(
            "interactive input was closed; provide input in guided mode or use subcommands"
        ));
    }

    Ok(line
        .trim_end_matches(|c| c == '\n' || c == '\r')
        .to_string())
}

fn parse_user_path(raw: &str) -> PathBuf {
    let trimmed = raw.trim();
    let unquoted = if trimmed.len() >= 2
        && ((trimmed.starts_with('"') && trimmed.ends_with('"'))
            || (trimmed.starts_with('\'') && trimmed.ends_with('\'')))
    {
        trimmed[1..trimmed.len() - 1].to_string()
    } else {
        trimmed.to_string()
    };

    let mut normalized = String::with_capacity(unquoted.len());
    let mut chars = unquoted.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch != '\\' {
            normalized.push(ch);
            continue;
        }

        let Some(next) = chars.peek().copied() else {
            normalized.push('\\');
            break;
        };

        if matches!(
            next,
            ' ' | '\\' | '"' | '\'' | '(' | ')' | '[' | ']' | '{' | '}'
        ) {
            normalized.push(next);
            chars.next();
        } else {
            normalized.push('\\');
        }
    }

    expand_tilde_path(PathBuf::from(normalized))
}

fn expand_tilde_path(path: PathBuf) -> PathBuf {
    let Some(raw) = path.to_str() else {
        return path;
    };
    if raw == "~" {
        return home_dir().unwrap_or(path);
    }

    if let Some(stripped) = raw.strip_prefix("~/").or_else(|| raw.strip_prefix("~\\")) {
        if let Some(home) = home_dir() {
            return home.join(stripped);
        }
    }

    path
}

fn home_dir() -> Option<PathBuf> {
    std::env::var_os("HOME")
        .map(PathBuf::from)
        .or_else(|| std::env::var_os("USERPROFILE").map(PathBuf::from))
}

fn run_md2docx(
    input: PathBuf,
    output: PathBuf,
    glob_pattern: Option<String>,
    template: Option<PathBuf>,
    cli_style_map_path: Option<PathBuf>,
    config_path: Option<PathBuf>,
    report_path: Option<PathBuf>,
    strict_flag: bool,
    allow_remote_images: bool,
) -> Result<i32> {
    let (config_file, config) = load_config_with_auto_discovery(config_path.as_deref())?;
    ensure_exists(&input)?;
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
    let policy = config.unsupported_policy_or_default();
    let strict = strict_mode(policy, strict_flag);

    if input.is_file() {
        if glob_pattern.is_some() {
            return Err(anyhow!("--glob requires a directory input path"));
        }

        let warnings = convert_md2docx_single(
            &input,
            &output,
            &style_map,
            template_path.as_deref(),
            report_path.as_deref(),
            strict,
            allow_remote_images,
        )?;
        return Ok(exit_code_from_warnings(&warnings, strict));
    }

    if !input.is_dir() {
        return Err(anyhow!(
            "input path is neither a file nor directory: {}",
            input.display()
        ));
    }

    ensure_batch_output_root(&output, "output")?;
    ensure_batch_output_root_optional(report_path.as_deref(), "report")?;

    let batch_inputs = collect_batch_inputs(&input, glob_pattern.as_deref(), "md")?;
    let mut outcome = BatchOutcome::default();

    for batch_input in batch_inputs {
        let batch_output = map_batch_output_path(&input, &batch_input, &output, "docx")?;
        let batch_report =
            map_batch_report_path(report_path.as_deref(), &input, &batch_input, "json")?;

        match convert_md2docx_single(
            &batch_input,
            &batch_output,
            &style_map,
            template_path.as_deref(),
            batch_report.as_deref(),
            strict,
            allow_remote_images,
        ) {
            Ok(warnings) => {
                outcome.converted += 1;
                outcome.warnings += warnings.len();
            }
            Err(err) => {
                outcome.failed += 1;
                eprintln!(
                    "md2docx failed: {} -> {}: {err:#}",
                    batch_input.display(),
                    batch_output.display()
                );
            }
        }
    }

    emit_batch_summary("md2docx", &input, &output, &outcome, strict);
    Ok(exit_code_from_batch(&outcome, strict))
}

fn run_docx2md(
    input: PathBuf,
    output: PathBuf,
    glob_pattern: Option<String>,
    cli_assets_dir: Option<PathBuf>,
    cli_style_map_path: Option<PathBuf>,
    config_path: Option<PathBuf>,
    report_path: Option<PathBuf>,
    strict_flag: bool,
) -> Result<i32> {
    let (config_file, config) = load_config_with_auto_discovery(config_path.as_deref())?;
    ensure_exists(&input)?;
    let style_map = load_effective_style_map(
        config_file.as_deref(),
        config.style_map.as_deref(),
        cli_style_map_path.as_deref(),
    )?;
    let policy = config.unsupported_policy_or_default();
    let strict = strict_mode(policy, strict_flag);

    if input.is_file() {
        if glob_pattern.is_some() {
            return Err(anyhow!("--glob requires a directory input path"));
        }

        let warnings = convert_docx2md_single(
            &input,
            &output,
            cli_assets_dir.as_deref(),
            config.assets_dir.as_deref(),
            &style_map,
            report_path.as_deref(),
            strict,
        )?;
        return Ok(exit_code_from_warnings(&warnings, strict));
    }

    if !input.is_dir() {
        return Err(anyhow!(
            "input path is neither a file nor directory: {}",
            input.display()
        ));
    }

    ensure_batch_output_root(&output, "output")?;
    ensure_batch_output_root_optional(report_path.as_deref(), "report")?;

    let batch_inputs = collect_batch_inputs(&input, glob_pattern.as_deref(), "docx")?;
    let mut outcome = BatchOutcome::default();

    for batch_input in batch_inputs {
        let batch_output = map_batch_output_path(&input, &batch_input, &output, "md")?;
        let batch_report =
            map_batch_report_path(report_path.as_deref(), &input, &batch_input, "json")?;

        match convert_docx2md_single(
            &batch_input,
            &batch_output,
            cli_assets_dir.as_deref(),
            config.assets_dir.as_deref(),
            &style_map,
            batch_report.as_deref(),
            strict,
        ) {
            Ok(warnings) => {
                outcome.converted += 1;
                outcome.warnings += warnings.len();
            }
            Err(err) => {
                outcome.failed += 1;
                eprintln!(
                    "docx2md failed: {} -> {}: {err:#}",
                    batch_input.display(),
                    batch_output.display()
                );
            }
        }
    }

    emit_batch_summary("docx2md", &input, &output, &outcome, strict);
    Ok(exit_code_from_batch(&outcome, strict))
}

fn ensure_exists(path: &Path) -> Result<()> {
    if path.exists() {
        Ok(())
    } else {
        Err(anyhow!("input path does not exist: {}", path.display()))
    }
}

#[derive(Debug, Default)]
struct BatchOutcome {
    converted: usize,
    failed: usize,
    warnings: usize,
}

fn convert_md2docx_single(
    input: &Path,
    output: &Path,
    style_map: &StyleMap,
    template_path: Option<&Path>,
    report_path: Option<&Path>,
    strict: bool,
    allow_remote_images: bool,
) -> Result<Vec<ConversionWarning>> {
    let started = Instant::now();
    let input_data = fs::read_to_string(input)
        .with_context(|| format!("failed reading markdown input: {}", input.display()))?;

    let (document, mut warnings) = parse_markdown(&input_data)?;
    let input_base = input
        .parent()
        .map(Path::to_path_buf)
        .unwrap_or_else(|| PathBuf::from("."));

    let mut write_warnings = write_docx(
        &document,
        &input_base,
        output,
        &DocxWriteOptions {
            allow_remote_images,
            style_map: style_map.clone(),
            template: template_path.map(Path::to_path_buf),
        },
    )?;
    warnings.append(&mut write_warnings);

    let duration = started.elapsed().as_millis();
    emit_summary(
        &ConversionDirection::MdToDocx,
        input,
        output,
        &warnings,
        strict,
    );
    write_report_if_requested(
        report_path,
        ConversionDirection::MdToDocx,
        input,
        output,
        duration,
        &document,
        &warnings,
    )?;

    Ok(warnings)
}

fn convert_docx2md_single(
    input: &Path,
    output: &Path,
    cli_assets_dir: Option<&Path>,
    config_assets_dir: Option<&Path>,
    style_map: &StyleMap,
    report_path: Option<&Path>,
    strict: bool,
) -> Result<Vec<ConversionWarning>> {
    let started = Instant::now();
    let output_parent = output
        .parent()
        .map(Path::to_path_buf)
        .unwrap_or_else(|| PathBuf::from("."));
    let configured_assets_dir = cli_assets_dir
        .map(Path::to_path_buf)
        .or_else(|| config_assets_dir.map(Path::to_path_buf))
        .unwrap_or_else(|| default_assets_dir_for_output(output));
    let assets_dir = resolve_output_relative_path(&output_parent, configured_assets_dir);

    let (mut document, warnings) = read_docx(
        input,
        &DocxReadOptions {
            assets_dir: assets_dir.clone(),
            style_map: style_map.clone(),
        },
    )?;

    rewrite_image_paths_relative_to_output(&mut document, &output_parent);
    let markdown = render_markdown(&document);

    if let Some(parent) = output.parent() {
        fs::create_dir_all(parent)
            .with_context(|| format!("failed creating output directory: {}", parent.display()))?;
    }
    fs::write(output, markdown)
        .with_context(|| format!("failed writing markdown output: {}", output.display()))?;

    let duration = started.elapsed().as_millis();
    emit_summary(
        &ConversionDirection::DocxToMd,
        input,
        output,
        &warnings,
        strict,
    );
    write_report_if_requested(
        report_path,
        ConversionDirection::DocxToMd,
        input,
        output,
        duration,
        &document,
        &warnings,
    )?;

    Ok(warnings)
}

fn collect_batch_inputs(
    input_root: &Path,
    glob_pattern: Option<&str>,
    default_extension: &str,
) -> Result<Vec<PathBuf>> {
    let matcher = match glob_pattern {
        Some(raw_pattern) => Some(parse_batch_glob_pattern(raw_pattern)?),
        None => None,
    };

    let mut inputs = Vec::new();
    for entry in WalkDir::new(input_root) {
        let entry = entry.with_context(|| {
            format!(
                "failed traversing batch input directory: {}",
                input_root.display()
            )
        })?;
        if !entry.file_type().is_file() {
            continue;
        }

        let include = if let Some(matcher) = &matcher {
            let relative = entry.path().strip_prefix(input_root).with_context(|| {
                format!(
                    "failed computing path relative to batch input root: {}",
                    input_root.display()
                )
            })?;
            let normalized = normalize_relative_path(relative);
            matcher.matches(&normalized)
        } else {
            entry
                .path()
                .extension()
                .and_then(|value| value.to_str())
                .map(|value| value.eq_ignore_ascii_case(default_extension))
                .unwrap_or(false)
        };

        if include {
            inputs.push(entry.path().to_path_buf());
        }
    }

    inputs.sort();
    if inputs.is_empty() {
        if let Some(pattern) = glob_pattern {
            return Err(anyhow!(
                "no input files matched pattern '{}' under {}",
                pattern,
                input_root.display()
            ));
        }

        return Err(anyhow!(
            "no .{} files found under {}",
            default_extension,
            input_root.display()
        ));
    }

    Ok(inputs)
}

fn parse_batch_glob_pattern(pattern: &str) -> Result<Pattern> {
    let normalized = pattern
        .trim()
        .replace('\\', "/")
        .trim_start_matches("./")
        .trim_start_matches('/')
        .to_string();

    if normalized.is_empty() {
        return Err(anyhow!("--glob pattern cannot be empty"));
    }

    Pattern::new(&normalized).with_context(|| format!("invalid --glob pattern: {pattern}"))
}

fn normalize_relative_path(path: &Path) -> String {
    path.to_string_lossy().replace('\\', "/")
}

fn ensure_batch_output_root(path: &Path, kind: &str) -> Result<()> {
    if path.exists() && !path.is_dir() {
        return Err(anyhow!(
            "batch {} path must be a directory: {}",
            kind,
            path.display()
        ));
    }

    fs::create_dir_all(path).with_context(|| {
        format!(
            "failed creating batch {} directory: {}",
            kind,
            path.display()
        )
    })
}

fn ensure_batch_output_root_optional(path: Option<&Path>, kind: &str) -> Result<()> {
    if let Some(path) = path {
        ensure_batch_output_root(path, kind)?;
    }
    Ok(())
}

fn map_batch_output_path(
    input_root: &Path,
    input_path: &Path,
    output_root: &Path,
    output_extension: &str,
) -> Result<PathBuf> {
    let relative = input_path.strip_prefix(input_root).with_context(|| {
        format!(
            "failed computing output path relative to input root {}",
            input_root.display()
        )
    })?;

    let mut output_path = output_root.join(relative);
    output_path.set_extension(output_extension);

    if let Some(parent) = output_path.parent() {
        fs::create_dir_all(parent)
            .with_context(|| format!("failed creating output directory: {}", parent.display()))?;
    }

    Ok(output_path)
}

fn map_batch_report_path(
    report_root: Option<&Path>,
    input_root: &Path,
    input_path: &Path,
    report_extension: &str,
) -> Result<Option<PathBuf>> {
    let Some(report_root) = report_root else {
        return Ok(None);
    };

    let relative = input_path.strip_prefix(input_root).with_context(|| {
        format!(
            "failed computing report path relative to input root {}",
            input_root.display()
        )
    })?;

    let mut report_path = report_root.join(relative);
    report_path.set_extension(report_extension);
    Ok(Some(report_path))
}

fn emit_batch_summary(
    command: &str,
    input_root: &Path,
    output_root: &Path,
    outcome: &BatchOutcome,
    strict: bool,
) {
    println!(
        "{} batch completed: {} -> {}",
        command,
        input_root.display(),
        output_root.display()
    );
    println!("converted: {}", outcome.converted);
    println!("failed: {}", outcome.failed);
    println!("warnings: {}", outcome.warnings);

    if strict && outcome.failed == 0 && outcome.warnings > 0 {
        println!("strict mode enabled: warnings will produce exit code 2");
    }
}

fn exit_code_from_batch(outcome: &BatchOutcome, strict: bool) -> i32 {
    if outcome.failed > 0 {
        1
    } else if strict && outcome.warnings > 0 {
        2
    } else {
        0
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
        let code = warning.code.as_str();
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
