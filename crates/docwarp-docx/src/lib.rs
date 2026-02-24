use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet};
use std::env;
use std::fs;
use std::io::{self, Read, Seek, Write};
use std::path::{Path, PathBuf};
use std::process;
use std::time::{SystemTime, UNIX_EPOCH};

use anyhow::{Context, Result, anyhow};
use docwarp_core::{
    Block, ConversionWarning, Document, Inline, StyleMap, WarningCode, model::inline_text,
};
use latex2mathml::{DisplayStyle, latex_to_mathml};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use sha2::{Digest, Sha256};
use ureq::Error as UreqError;
use zip::ZipArchive;
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

const OFFICE_REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const WORDPROCESSINGML_NUMBERING_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const LIST_BASE_INDENT_TWIPS: u32 = 720;
const LIST_INDENT_STEP_TWIPS: u32 = 360;
const LIST_MAX_LEVEL: u8 = 8;
const ORDERED_LIST_ABSTRACT_NUM_ID: u32 = 1;
const BULLET_LIST_ABSTRACT_NUM_ID: u32 = 2;
const ORDERED_LIST_NUM_ID: u32 = 1;
const BULLET_LIST_NUM_ID: u32 = 2;
const CODE_LANG_MARKER_PREFIX: &str = "[[docwarp-code-lang:";
const CODE_LANG_MARKER_SUFFIX: &str = "]]";
const EQUATION_MARKER_PREFIX: &str = "[[docwarp-eq:";
const EQUATION_MARKER_SUFFIX: &str = "]]";
const MSOFFCRYPTO_TOOL_VERSION: &str = "6.0.0";
const MSOFFCRYPTO_TOOL_WHEEL_SHA256: &str =
    "46c394ed5d9641e802fc79bf3fb0666a53748b23fa8c4aa634ae9d30d46fe397";

#[derive(Debug, Clone)]
pub struct DocxWriteOptions {
    pub allow_remote_images: bool,
    pub style_map: StyleMap,
    pub template: Option<PathBuf>,
}

#[derive(Debug, Clone)]
pub struct DocxReadOptions {
    pub assets_dir: PathBuf,
    pub style_map: StyleMap,
    pub password: Option<String>,
}

#[derive(Debug, Clone)]
struct Relationship {
    id: String,
    rel_type: String,
    target: String,
    target_mode: Option<String>,
}

#[derive(Debug, Clone)]
struct MediaFile {
    target: String,
    extension: String,
    content_type: String,
    bytes: Vec<u8>,
}

#[derive(Default)]
struct DocxBuildState {
    relationships: Vec<Relationship>,
    media_files: Vec<MediaFile>,
    next_rel_id: usize,
    next_media_index: usize,
    next_docpr_id: usize,
    reserved_media_targets: BTreeSet<String>,
}

impl DocxBuildState {
    fn from_template(template: Option<&TemplatePackage>) -> Self {
        let mut state = DocxBuildState::default();

        if let Some(template) = template {
            state.relationships = template.document_relationships.clone();
            state.next_rel_id = template
                .document_relationships
                .iter()
                .filter_map(|rel| parse_numeric_rel_id(&rel.id))
                .max()
                .unwrap_or(0);

            for path in template.entries.keys() {
                if let Some(target) = path.strip_prefix("word/") {
                    if target.starts_with("media/") {
                        state.reserved_media_targets.insert(target.to_string());
                    }
                }
            }
        }

        state
    }

    fn next_rel_id(&mut self) -> String {
        self.next_rel_id += 1;
        format!("rId{}", self.next_rel_id)
    }

    fn next_docpr_id(&mut self) -> usize {
        self.next_docpr_id += 1;
        self.next_docpr_id
    }

    fn add_hyperlink(&mut self, target: &str) -> String {
        let id = self.next_rel_id();
        self.relationships.push(Relationship {
            id: id.clone(),
            rel_type: format!("{OFFICE_REL_NS}/hyperlink"),
            target: target.to_string(),
            target_mode: Some("External".to_string()),
        });
        id
    }

    fn add_image(&mut self, extension: &str, content_type: &str, bytes: Vec<u8>) -> String {
        let image_index = self.next_available_media_index(extension);
        let rel_id = self.next_rel_id();
        let filename = format!("image{image_index}.{extension}");
        let target = format!("media/{filename}");

        self.relationships.push(Relationship {
            id: rel_id.clone(),
            rel_type: format!("{OFFICE_REL_NS}/image"),
            target: target.clone(),
            target_mode: None,
        });

        self.media_files.push(MediaFile {
            target,
            extension: extension.to_string(),
            content_type: content_type.to_string(),
            bytes,
        });

        rel_id
    }

    fn next_available_media_index(&mut self, extension: &str) -> usize {
        loop {
            self.next_media_index += 1;
            let candidate = format!("media/image{}.{}", self.next_media_index, extension);
            let conflict = self.reserved_media_targets.contains(&candidate)
                || self
                    .media_files
                    .iter()
                    .any(|media| media.target.eq_ignore_ascii_case(&candidate));

            if !conflict {
                return self.next_media_index;
            }
        }
    }
}

#[derive(Default)]
struct RunStyle {
    bold: bool,
    italic: bool,
    code: bool,
}

#[derive(Debug, Clone, Copy)]
struct ListNumbering {
    ordered: bool,
    level: u8,
}

#[derive(Default)]
struct ParseParagraph {
    style: Option<String>,
    indent_left: Option<u32>,
    inlines: Vec<Inline>,
}

#[derive(Default)]
struct ParseTable {
    rows: Vec<Vec<Vec<Inline>>>,
    current_row: Vec<Vec<Inline>>,
    current_cell: Vec<Inline>,
}

#[derive(Default)]
struct PendingList {
    ordered: bool,
    base_indent_left: Option<u32>,
    items: Vec<Vec<Inline>>,
    levels: Vec<u8>,
    item_ordered: Vec<bool>,
}

#[derive(Default)]
struct EquationCapture {
    display: bool,
    text: String,
    unsupported: bool,
    depth: usize,
}

#[derive(Debug, Clone, Default)]
struct MathMlNode {
    name: String,
    text: String,
    attributes: BTreeMap<String, String>,
    children: Vec<MathMlNode>,
}

#[derive(Debug, Clone)]
struct TemplatePackage {
    entries: BTreeMap<String, Vec<u8>>,
    document_relationships: Vec<Relationship>,
    section_properties_xml: Option<String>,
}

pub fn write_docx(
    document: &Document,
    markdown_base_dir: &Path,
    output_path: &Path,
    options: &DocxWriteOptions,
) -> Result<Vec<ConversionWarning>> {
    let mut warnings = Vec::new();
    let template = load_template_package(options.template.as_deref(), &mut warnings)?;

    if let Some(parent) = output_path.parent() {
        fs::create_dir_all(parent)
            .with_context(|| format!("failed creating output directory: {}", parent.display()))?;
    }

    let mut state = DocxBuildState::from_template(template.as_ref());
    ensure_styles_relationship(&mut state);
    ensure_numbering_relationship(&mut state);

    let document_xml = build_document_xml(
        document,
        markdown_base_dir,
        options,
        template
            .as_ref()
            .and_then(|package| package.section_properties_xml.as_deref()),
        &mut state,
        &mut warnings,
    )?;

    let styles_xml = resolve_styles_xml(template.as_ref());
    let numbering_xml = resolve_numbering_xml(template.as_ref());
    let template_content_types = template
        .as_ref()
        .and_then(|package| package.entries.get("[Content_Types].xml"))
        .map(|bytes| bytes.as_slice());
    let content_types_xml = build_content_types_xml(&state.media_files, template_content_types);
    let package_rels_xml = template
        .as_ref()
        .and_then(|package| package.entries.get("_rels/.rels"))
        .cloned()
        .unwrap_or_else(build_package_relationships_xml);
    let document_rels_xml = build_document_relationships_xml(&state.relationships);

    let mut output_entries = template
        .as_ref()
        .map(|package| package.entries.clone())
        .unwrap_or_default();

    output_entries.insert("[Content_Types].xml".to_string(), content_types_xml);
    output_entries.insert("_rels/.rels".to_string(), package_rels_xml);
    output_entries.insert("word/document.xml".to_string(), document_xml);
    output_entries.insert(
        "word/_rels/document.xml.rels".to_string(),
        document_rels_xml,
    );
    output_entries.insert("word/styles.xml".to_string(), styles_xml);
    output_entries.insert("word/numbering.xml".to_string(), numbering_xml);
    output_entries.insert("docProps/core.xml".to_string(), build_core_properties_xml());
    output_entries.insert("docProps/app.xml".to_string(), build_app_properties_xml());

    let output_parent = output_path.parent().unwrap_or_else(|| Path::new("."));
    let mut temp_output = tempfile::Builder::new()
        .prefix(".docwarp-")
        .suffix(".docx.tmp")
        .tempfile_in(output_parent)
        .with_context(|| {
            format!(
                "failed creating temporary DOCX output in {}",
                output_parent.display()
            )
        })?;

    {
        let mut zip = ZipWriter::new(temp_output.as_file_mut());
        let file_options = SimpleFileOptions::default();

        for (path, bytes) in output_entries {
            write_zip_entry(&mut zip, &path, &bytes, file_options)?;
        }

        for media in &state.media_files {
            let path = format!("word/{}", media.target);
            write_zip_entry(&mut zip, &path, &media.bytes, file_options)?;
        }

        zip.finish().context("failed finalizing DOCX zip")?;
    }

    let _ = temp_output.as_file_mut().sync_all();
    persist_tempfile_replace(temp_output, output_path, "DOCX output")?;

    Ok(warnings)
}

fn write_zip_entry<W: Write + Seek>(
    zip: &mut ZipWriter<W>,
    path: &str,
    bytes: &[u8],
    file_options: SimpleFileOptions,
) -> Result<()> {
    zip.start_file(path, file_options)
        .with_context(|| format!("failed writing zip entry: {path}"))?;
    zip.write_all(bytes)
        .with_context(|| format!("failed writing zip entry bytes: {path}"))?;
    Ok(())
}

fn persist_tempfile_replace(
    tempfile: tempfile::NamedTempFile,
    destination: &Path,
    label: &str,
) -> Result<()> {
    match tempfile.persist(destination) {
        Ok(_) => Ok(()),
        Err(err) => {
            if err.error.kind() != io::ErrorKind::AlreadyExists || !destination.exists() {
                return Err(anyhow!(
                    "failed writing {} at {}: {}",
                    label,
                    destination.display(),
                    err.error
                ));
            }

            let backup_path = temporary_backup_path(destination);
            fs::rename(destination, &backup_path).with_context(|| {
                format!(
                    "failed moving existing {} to backup before replacement: {} -> {}",
                    label,
                    destination.display(),
                    backup_path.display()
                )
            })?;

            match err.file.persist(destination) {
                Ok(_) => {
                    let _ = fs::remove_file(&backup_path);
                    Ok(())
                }
                Err(second_err) => {
                    let _ = fs::remove_file(second_err.file.path());
                    let restore_result = fs::rename(&backup_path, destination);
                    let restore_note = match restore_result {
                        Ok(_) => String::new(),
                        Err(restore_err) => format!(
                            " (also failed restoring original file from {}: {restore_err})",
                            backup_path.display()
                        ),
                    };
                    Err(anyhow!(
                        "failed replacing {} at {}: {}{}",
                        label,
                        destination.display(),
                        second_err.error,
                        restore_note
                    ))
                }
            }
        }
    }
}

fn temporary_backup_path(destination: &Path) -> PathBuf {
    let parent = destination.parent().unwrap_or_else(|| Path::new("."));
    let name = destination
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or("output");
    let now = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_default()
        .as_nanos();
    parent.join(format!(".{name}.docwarp-backup-{}-{now}", process::id()))
}

fn build_document_xml(
    document: &Document,
    markdown_base_dir: &Path,
    options: &DocxWriteOptions,
    section_properties_xml: Option<&str>,
    state: &mut DocxBuildState,
    warnings: &mut Vec<ConversionWarning>,
) -> Result<Vec<u8>> {
    let mut body = String::new();

    for (block_index, block) in document.blocks.iter().enumerate() {
        match block {
            Block::Title(content) => {
                body.push_str(&render_paragraph(
                    content,
                    &options.style_map.docx_style_for("title"),
                    if block_index > 0 { Some(240) } else { None },
                    Some(240),
                    None,
                    None,
                    None,
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                )?);
            }
            Block::Heading { level, content } => {
                let token = match *level {
                    1 => "h1",
                    2 => "h2",
                    3 => "h3",
                    4 => "h4",
                    5 => "h5",
                    _ => "h6",
                };

                body.push_str(&render_paragraph(
                    content,
                    &options.style_map.docx_style_for(token),
                    if block_index > 0 { Some(240) } else { None },
                    Some(240),
                    None,
                    None,
                    None,
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                )?);
            }
            Block::Paragraph(content) => {
                if let Some(tex) = single_display_equation(content) {
                    body.push_str(&render_equation_paragraph(
                        tex,
                        &options.style_map.docx_style_for("equation_block"),
                        &options.style_map.docx_style_for("equation_inline"),
                        None,
                        Some(240),
                        warnings,
                    ));
                } else {
                    body.push_str(&render_paragraph(
                        content,
                        &options.style_map.docx_style_for("paragraph"),
                        None,
                        Some(240),
                        None,
                        None,
                        None,
                        markdown_base_dir,
                        options,
                        state,
                        warnings,
                    )?);
                }
            }
            Block::BlockQuote(content) => {
                body.push_str(&render_paragraph(
                    content,
                    &options.style_map.docx_style_for("quote"),
                    None,
                    None,
                    None,
                    None,
                    None,
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                )?);
            }
            Block::CodeBlock { language, code } => {
                let mut code_inlines = Vec::new();
                for (idx, line) in code.lines().enumerate() {
                    if idx > 0 {
                        code_inlines.push(Inline::LineBreak);
                    }
                    code_inlines.push(Inline::Code(line.to_string()));
                }
                if code_inlines.is_empty() {
                    code_inlines.push(Inline::Code(String::new()));
                }

                body.push_str(&render_paragraph(
                    &code_inlines,
                    &options.style_map.docx_style_for("code"),
                    None,
                    None,
                    None,
                    None,
                    language.as_deref(),
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                )?);
            }
            Block::List {
                ordered,
                items,
                levels,
                item_ordered,
            } => {
                for (index, item) in items.iter().enumerate() {
                    let is_ordered = *item_ordered.get(index).unwrap_or(ordered);
                    let style = if is_ordered {
                        options.style_map.docx_style_for("list_number")
                    } else {
                        options.style_map.docx_style_for("list_bullet")
                    };
                    let level = *levels.get(index).unwrap_or(&0);
                    let clamped_level = level.min(LIST_MAX_LEVEL);
                    let level_twips = u32::from(clamped_level);
                    let indent_left = LIST_BASE_INDENT_TWIPS
                        .saturating_add(level_twips.saturating_mul(LIST_INDENT_STEP_TWIPS));

                    body.push_str(&render_paragraph(
                        item,
                        &style,
                        None,
                        None,
                        Some(indent_left),
                        Some(ListNumbering {
                            ordered: is_ordered,
                            level: clamped_level,
                        }),
                        None,
                        markdown_base_dir,
                        options,
                        state,
                        warnings,
                    )?);
                }
            }
            Block::Table { headers, rows } => {
                body.push_str(&render_table(
                    headers,
                    rows,
                    &options.style_map.docx_style_for("table"),
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                )?);
            }
            Block::Image { alt, src, .. } => {
                let inline = Inline::Image {
                    alt: alt.clone(),
                    src: src.clone(),
                    title: None,
                };
                body.push_str(&render_paragraph(
                    &[inline],
                    &options.style_map.docx_style_for("paragraph"),
                    None,
                    Some(240),
                    None,
                    None,
                    None,
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                )?);
            }
            Block::ThematicBreak => {
                body.push_str(&render_paragraph(
                    &[Inline::Text("---".to_string())],
                    &options.style_map.docx_style_for("paragraph"),
                    None,
                    None,
                    None,
                    None,
                    None,
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                )?);
            }
        }
    }

    match section_properties_xml {
        Some(section_properties_xml) => body.push_str(section_properties_xml),
        None => body.push_str(default_section_properties_xml()),
    }

    let xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"{OFFICE_REL_NS}\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 wp14\"><w:body>{body}</w:body></w:document>"
    );

    Ok(xml.into_bytes())
}

fn default_section_properties_xml() -> &'static str {
    "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/></w:sectPr>"
}

fn render_table(
    headers: &[Vec<Inline>],
    rows: &[Vec<Vec<Inline>>],
    style: &str,
    markdown_base_dir: &Path,
    options: &DocxWriteOptions,
    state: &mut DocxBuildState,
    warnings: &mut Vec<ConversionWarning>,
) -> Result<String> {
    let width = headers
        .len()
        .max(rows.iter().map(Vec::len).max().unwrap_or_default());
    let mut normalized_headers = headers.to_vec();
    let mut normalized_rows = rows.to_vec();
    normalize_table_dimensions(&mut normalized_headers, &mut normalized_rows, width);

    let mut out = String::new();
    out.push_str("<w:tbl><w:tblPr>");
    out.push_str(&format!(
        "<w:tblStyle w:val=\"{}\"/><w:tblW w:w=\"0\" w:type=\"auto\"/>",
        escape_xml(style)
    ));
    out.push_str(
        "<w:tblBorders><w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/><w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/></w:tblBorders>",
    );
    out.push_str("</w:tblPr>");

    if !normalized_headers.is_empty() {
        out.push_str("<w:tr>");
        for cell in &normalized_headers {
            out.push_str("<w:tc><w:p>");
            out.push_str(&render_inlines(
                cell,
                markdown_base_dir,
                options,
                state,
                warnings,
            )?);
            out.push_str("</w:p></w:tc>");
        }
        out.push_str("</w:tr>");
    }

    for row in &normalized_rows {
        out.push_str("<w:tr>");
        for cell in row {
            out.push_str("<w:tc><w:p>");
            out.push_str(&render_inlines(
                cell,
                markdown_base_dir,
                options,
                state,
                warnings,
            )?);
            out.push_str("</w:p></w:tc>");
        }
        out.push_str("</w:tr>");
    }

    out.push_str("</w:tbl>");
    Ok(out)
}

fn render_paragraph(
    inlines: &[Inline],
    style: &str,
    spacing_before_twips: Option<u32>,
    spacing_after_twips: Option<u32>,
    indent_left_twips: Option<u32>,
    list_numbering: Option<ListNumbering>,
    code_language: Option<&str>,
    markdown_base_dir: &Path,
    options: &DocxWriteOptions,
    state: &mut DocxBuildState,
    warnings: &mut Vec<ConversionWarning>,
) -> Result<String> {
    let mut out = String::new();
    out.push_str("<w:p><w:pPr>");
    out.push_str(&format!("<w:pStyle w:val=\"{}\"/>", escape_xml(style)));
    if spacing_before_twips.is_some() || spacing_after_twips.is_some() {
        let mut spacing = String::from("<w:spacing");
        if let Some(before) = spacing_before_twips {
            spacing.push_str(&format!(" w:before=\"{before}\""));
        }
        if let Some(after) = spacing_after_twips {
            spacing.push_str(&format!(" w:after=\"{after}\""));
        }
        spacing.push_str("/>");
        out.push_str(&spacing);
    }
    if let Some(indent_left) = indent_left_twips {
        out.push_str(&format!(
            "<w:ind w:left=\"{indent_left}\" w:hanging=\"360\"/>"
        ));
    }
    if let Some(list) = list_numbering {
        let num_id = if list.ordered {
            ORDERED_LIST_NUM_ID
        } else {
            BULLET_LIST_NUM_ID
        };
        out.push_str(&format!(
            "<w:numPr><w:ilvl w:val=\"{}\"/><w:numId w:val=\"{}\"/></w:numPr>",
            list.level, num_id
        ));
    }
    out.push_str("</w:pPr>");
    if let Some(language) = code_language.filter(|value| !value.trim().is_empty()) {
        out.push_str(&render_hidden_code_language_marker(language.trim()));
    }
    out.push_str(&render_inlines(
        inlines,
        markdown_base_dir,
        options,
        state,
        warnings,
    )?);
    out.push_str("</w:p>");
    Ok(out)
}

fn single_display_equation(inlines: &[Inline]) -> Option<&str> {
    match inlines {
        [Inline::Equation { tex, display: true }] => Some(tex.as_str()),
        _ => None,
    }
}

fn render_equation_paragraph(
    tex: &str,
    paragraph_style: &str,
    equation_inline_style: &str,
    spacing_before_twips: Option<u32>,
    spacing_after_twips: Option<u32>,
    warnings: &mut Vec<ConversionWarning>,
) -> String {
    let mut out = String::new();
    out.push_str("<w:p><w:pPr>");
    out.push_str(&format!(
        "<w:pStyle w:val=\"{}\"/>",
        escape_xml(paragraph_style)
    ));
    if spacing_before_twips.is_some() || spacing_after_twips.is_some() {
        let mut spacing = String::from("<w:spacing");
        if let Some(before) = spacing_before_twips {
            spacing.push_str(&format!(" w:before=\"{before}\""));
        }
        if let Some(after) = spacing_after_twips {
            spacing.push_str(&format!(" w:after=\"{after}\""));
        }
        spacing.push_str("/>");
        out.push_str(&spacing);
    }
    out.push_str("</w:pPr>");
    out.push_str("<m:oMathPara>");
    out.push_str(&render_omml(tex, equation_inline_style, warnings));
    out.push_str("</m:oMathPara>");
    out.push_str(&render_hidden_equation_marker(tex, true));
    out.push_str("</w:p>");
    out
}

fn render_inlines(
    inlines: &[Inline],
    markdown_base_dir: &Path,
    options: &DocxWriteOptions,
    state: &mut DocxBuildState,
    warnings: &mut Vec<ConversionWarning>,
) -> Result<String> {
    let mut out = String::new();
    for inline in inlines {
        render_inline(
            inline,
            RunStyle::default(),
            markdown_base_dir,
            options,
            state,
            warnings,
            &mut out,
        )?;
    }
    Ok(out)
}

fn render_inline(
    inline: &Inline,
    mut style: RunStyle,
    markdown_base_dir: &Path,
    options: &DocxWriteOptions,
    state: &mut DocxBuildState,
    warnings: &mut Vec<ConversionWarning>,
    out: &mut String,
) -> Result<()> {
    match inline {
        Inline::Text(text) => out.push_str(&render_text_run(text, &style)),
        Inline::LineBreak => out.push_str("<w:r><w:br/></w:r>"),
        Inline::Code(code) => {
            style.code = true;
            out.push_str(&render_text_run(code, &style));
        }
        Inline::Emphasis(children) => {
            style.italic = true;
            for child in children {
                render_inline(
                    child,
                    RunStyle {
                        bold: style.bold,
                        italic: style.italic,
                        code: style.code,
                    },
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                    out,
                )?;
            }
        }
        Inline::Strong(children) => {
            style.bold = true;
            for child in children {
                render_inline(
                    child,
                    RunStyle {
                        bold: style.bold,
                        italic: style.italic,
                        code: style.code,
                    },
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                    out,
                )?;
            }
        }
        Inline::Link { text, url } => {
            let rel_id = state.add_hyperlink(url);
            out.push_str(&format!("<w:hyperlink r:id=\"{}\">", escape_xml(&rel_id)));
            for child in text {
                render_inline(
                    child,
                    RunStyle {
                        bold: style.bold,
                        italic: style.italic,
                        code: style.code,
                    },
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                    out,
                )?;
            }
            out.push_str("</w:hyperlink>");
        }
        Inline::Image { alt, src, .. } => {
            if let Some(image) = load_image(
                src,
                markdown_base_dir,
                options.allow_remote_images,
                warnings,
            ) {
                let rel_id = state.add_image(&image.extension, &image.content_type, image.bytes);
                let docpr_id = state.next_docpr_id();
                out.push_str(&render_image_run(
                    &rel_id,
                    docpr_id,
                    &image.name,
                    alt,
                    image.width_emu,
                    image.height_emu,
                ));
            }
        }
        Inline::Equation { tex, display } => {
            out.push_str(&render_omml(
                tex,
                &options.style_map.docx_style_for("equation_inline"),
                warnings,
            ));
            out.push_str(&render_hidden_equation_marker(tex, *display));
        }
    }

    Ok(())
}

fn render_text_run(text: &str, style: &RunStyle) -> String {
    let mut run = String::new();
    run.push_str("<w:r>");
    if style.bold || style.italic || style.code {
        run.push_str("<w:rPr>");
        if style.bold {
            run.push_str("<w:b/>");
        }
        if style.italic {
            run.push_str("<w:i/>");
        }
        if style.code {
            run.push_str("<w:rStyle w:val=\"Code\"/>");
            run.push_str(
                "<w:rFonts w:ascii=\"Consolas\" w:hAnsi=\"Consolas\"/>\n<w:sz w:val=\"20\"/>",
            );
        }
        run.push_str("</w:rPr>");
    }

    run.push_str(&format!(
        "<w:t xml:space=\"preserve\">{}</w:t>",
        escape_xml(text)
    ));
    run.push_str("</w:r>");
    run
}

fn render_hidden_code_language_marker(language: &str) -> String {
    let marker = format!("{CODE_LANG_MARKER_PREFIX}{language}{CODE_LANG_MARKER_SUFFIX}");
    format!(
        "<w:r><w:rPr><w:vanish/></w:rPr><w:t xml:space=\"preserve\">{}</w:t></w:r>",
        escape_xml(&marker)
    )
}

fn render_hidden_equation_marker(tex: &str, display: bool) -> String {
    let kind = if display { "d" } else { "i" };
    let marker = format!(
        "{EQUATION_MARKER_PREFIX}{kind}:{}{EQUATION_MARKER_SUFFIX}",
        encode_hex(tex)
    );
    format!(
        "<w:r><w:rPr><w:vanish/></w:rPr><w:t xml:space=\"preserve\">{}</w:t></w:r>",
        escape_xml(&marker)
    )
}

fn encode_hex(input: &str) -> String {
    let mut out = String::with_capacity(input.len() * 2);
    for byte in input.as_bytes() {
        out.push_str(&format!("{byte:02x}"));
    }
    out
}

fn render_omml(
    tex: &str,
    equation_inline_style: &str,
    warnings: &mut Vec<ConversionWarning>,
) -> String {
    match render_structured_omml(tex, equation_inline_style) {
        Ok(Some(body)) if !body.trim().is_empty() => format!("<m:oMath>{body}</m:oMath>"),
        Ok(_) => render_linear_omml(tex, equation_inline_style),
        Err(err) => {
            warnings.push(ConversionWarning::new(
                WarningCode::UnsupportedFeature,
                format!(
                    "Unable to emit structured OMML for equation; using linear fallback: {err}"
                ),
            ));
            render_linear_omml(tex, equation_inline_style)
        }
    }
}

fn render_linear_omml(tex: &str, _equation_inline_style: &str) -> String {
    let mut out = String::new();
    out.push_str("<m:oMath><m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr>");
    out.push_str(&format!("<m:t>{}</m:t>", escape_xml(tex.trim())));
    out.push_str("</m:r></m:oMath>");
    out
}

fn render_structured_omml(tex: &str, _equation_inline_style: &str) -> Result<Option<String>> {
    let trimmed = tex.trim();
    if trimmed.is_empty() {
        return Ok(None);
    }

    let mathml = latex_to_mathml(trimmed, DisplayStyle::Inline)
        .map_err(|err| anyhow!("LaTeX parse failed: {err}"))?;
    if mathml.contains("[PARSE ERROR:") {
        return Err(anyhow!(
            "LaTeX expression contains unsupported commands for structured OMML conversion"
        ));
    }
    let root = parse_xml_node_tree(&mathml)?;

    let body = if root.name == "math" {
        render_mathml_nodes_to_omml(&root.children)
    } else {
        render_mathml_node_to_omml(&root)
    };

    if body.trim().is_empty() {
        Ok(None)
    } else {
        Ok(Some(body))
    }
}

fn parse_xml_node_tree(xml: &str) -> Result<MathMlNode> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut stack: Vec<MathMlNode> = Vec::new();
    let mut root: Option<MathMlNode> = None;
    let mut buf = Vec::new();

    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(start)) => {
                stack.push(parse_xml_start_node(&start));
            }
            Ok(Event::Empty(start)) => {
                let node = parse_xml_start_node(&start);
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(node);
                } else if root.is_none() {
                    root = Some(node);
                }
            }
            Ok(Event::Text(text)) => {
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decode_text(&reader, text)?);
                }
            }
            Ok(Event::CData(cdata)) => {
                if let Some(node) = stack.last_mut() {
                    let decoded = reader.decoder().decode(cdata.as_ref())?.into_owned();
                    node.text.push_str(&decoded);
                }
            }
            Ok(Event::End(_)) => {
                if let Some(node) = stack.pop() {
                    if let Some(parent) = stack.last_mut() {
                        parent.children.push(node);
                    } else {
                        root = Some(node);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(err) => return Err(anyhow!("failed parsing generated MathML: {err}")),
        }

        buf.clear();
    }

    root.ok_or_else(|| anyhow!("generated MathML did not contain a root node"))
}

fn parse_xml_start_node(start: &BytesStart<'_>) -> MathMlNode {
    let mut attributes = BTreeMap::new();
    for attribute in start.attributes().flatten() {
        let key = String::from_utf8_lossy(local_name(attribute.key.as_ref())).to_string();
        let value = String::from_utf8_lossy(attribute.value.as_ref()).to_string();
        attributes.insert(key, value);
    }

    MathMlNode {
        name: String::from_utf8_lossy(local_name(start.name().as_ref())).to_string(),
        text: String::new(),
        attributes,
        children: Vec::new(),
    }
}

fn render_mathml_nodes_to_omml(nodes: &[MathMlNode]) -> String {
    let mut out = String::new();
    let mut index = 0usize;
    while index < nodes.len() {
        if let Some(next_node) = nodes.get(index + 1) {
            if let Some(arg_extremum) = render_arg_extremum_pair(&nodes[index], next_node) {
                out.push_str(&arg_extremum);
                index += 2;
                continue;
            }
        }

        if let Some((operator_node, sub_node, sup_node)) = extract_nary_limits(&nodes[index]) {
            out.push_str(&render_omml_nary(
                &mathml_token_text(operator_node),
                sub_node.map(render_mathml_node_to_omml),
                sup_node.map(render_mathml_node_to_omml),
                render_mathml_nodes_to_omml(&nodes[index + 1..]),
            ));
            break;
        }

        out.push_str(&render_mathml_node_to_omml(&nodes[index]));
        index += 1;
    }
    out
}

fn render_mathml_node_to_omml(node: &MathMlNode) -> String {
    match node.name.as_str() {
        "math" | "mstyle" | "semantics" => render_mathml_nodes_to_omml(&node.children),
        "annotation" => String::new(),
        "mrow" => render_mathml_mrow(node),
        "mi" | "mn" | "mo" | "mtext" => {
            let text = mathml_token_text(node);
            if text.is_empty() {
                String::new()
            } else {
                render_omml_run(&text)
            }
        }
        "msup" => {
            let Some(base_node) = node.children.first() else {
                return render_mathml_nodes_to_omml(&node.children);
            };
            let Some(sup_node) = node.children.get(1) else {
                return render_mathml_nodes_to_omml(&node.children);
            };

            if is_nary_operator_node(base_node) {
                return render_omml_nary(
                    &mathml_token_text(base_node),
                    None,
                    Some(render_mathml_node_to_omml(sup_node)),
                    String::new(),
                );
            }
            if is_limit_like_operator_node(base_node) {
                let base = render_mathml_node_to_omml(base_node);
                let sup = render_mathml_node_to_omml(sup_node);
                return render_limit_like_operator(base, None, Some(sup));
            }

            let base = render_mathml_node_to_omml(base_node);
            let sup = render_mathml_node_to_omml(sup_node);
            format!(
                "<m:sSup>{}{}</m:sSup>",
                wrap_omml_arg("e", base),
                wrap_omml_arg("sup", sup)
            )
        }
        "msub" => {
            let Some(base_node) = node.children.first() else {
                return render_mathml_nodes_to_omml(&node.children);
            };
            let Some(sub_node) = node.children.get(1) else {
                return render_mathml_nodes_to_omml(&node.children);
            };

            if is_nary_operator_node(base_node) {
                return render_omml_nary(
                    &mathml_token_text(base_node),
                    Some(render_mathml_node_to_omml(sub_node)),
                    None,
                    String::new(),
                );
            }
            if is_limit_like_operator_node(base_node) {
                let base = render_mathml_node_to_omml(base_node);
                let sub = render_mathml_node_to_omml(sub_node);
                return render_limit_like_operator(base, Some(sub), None);
            }

            let base = render_mathml_node_to_omml(base_node);
            let sub = render_mathml_node_to_omml(sub_node);
            format!(
                "<m:sSub>{}{}</m:sSub>",
                wrap_omml_arg("e", base),
                wrap_omml_arg("sub", sub)
            )
        }
        "msubsup" => {
            let Some(base_node) = node.children.first() else {
                return render_mathml_nodes_to_omml(&node.children);
            };
            let Some(sub_node) = node.children.get(1) else {
                return render_mathml_nodes_to_omml(&node.children);
            };
            let Some(sup_node) = node.children.get(2) else {
                return render_mathml_nodes_to_omml(&node.children);
            };

            if is_nary_operator_node(base_node) {
                return render_omml_nary(
                    &mathml_token_text(base_node),
                    Some(render_mathml_node_to_omml(sub_node)),
                    Some(render_mathml_node_to_omml(sup_node)),
                    String::new(),
                );
            }
            if is_limit_like_operator_node(base_node) {
                let base = render_mathml_node_to_omml(base_node);
                let sub = render_mathml_node_to_omml(sub_node);
                let sup = render_mathml_node_to_omml(sup_node);
                return render_limit_like_operator(base, Some(sub), Some(sup));
            }

            let base = render_mathml_node_to_omml(base_node);
            let sub = render_mathml_node_to_omml(sub_node);
            let sup = render_mathml_node_to_omml(sup_node);
            format!(
                "<m:sSubSup>{}{}{}</m:sSubSup>",
                wrap_omml_arg("e", base),
                wrap_omml_arg("sub", sub),
                wrap_omml_arg("sup", sup)
            )
        }
        "mfrac" => {
            let Some(num_node) = node.children.first() else {
                return render_mathml_nodes_to_omml(&node.children);
            };
            let Some(den_node) = node.children.get(1) else {
                return render_mathml_nodes_to_omml(&node.children);
            };

            let mut out = String::new();
            out.push_str("<m:f>");
            if node
                .attributes
                .get("linethickness")
                .map(|value| value.trim() == "0")
                .unwrap_or(false)
            {
                out.push_str("<m:fPr><m:type m:val=\"noBar\"/></m:fPr>");
            }
            out.push_str(&wrap_omml_arg("num", render_mathml_node_to_omml(num_node)));
            out.push_str(&wrap_omml_arg("den", render_mathml_node_to_omml(den_node)));
            out.push_str("</m:f>");
            out
        }
        "msqrt" => {
            let content = render_mathml_nodes_to_omml(&node.children);
            format!(
                "<m:rad><m:radPr><m:degHide m:val=\"1\"/></m:radPr>{}</m:rad>",
                wrap_omml_arg("e", content)
            )
        }
        "mroot" => {
            let Some(base_node) = node.children.first() else {
                return render_mathml_nodes_to_omml(&node.children);
            };
            let Some(deg_node) = node.children.get(1) else {
                return render_mathml_nodes_to_omml(&node.children);
            };

            let base = render_mathml_node_to_omml(base_node);
            let degree = render_mathml_node_to_omml(deg_node);
            format!(
                "<m:rad>{}{}</m:rad>",
                wrap_omml_arg("deg", degree),
                wrap_omml_arg("e", base)
            )
        }
        "mfenced" => {
            let open = node
                .attributes
                .get("open")
                .map(String::as_str)
                .unwrap_or("(");
            let close = node
                .attributes
                .get("close")
                .map(String::as_str)
                .unwrap_or(")");
            render_mathml_delimited(open, close, render_mathml_nodes_to_omml(&node.children))
        }
        "mtable" => render_mathml_table(node, None, None),
        "mtr" | "mtd" => render_mathml_nodes_to_omml(&node.children),
        "mover" => {
            let Some(base_node) = node.children.first() else {
                return render_mathml_nodes_to_omml(&node.children);
            };
            let Some(over_node) = node.children.get(1) else {
                return render_mathml_nodes_to_omml(&node.children);
            };

            let base = render_mathml_node_to_omml(base_node);
            let over = render_mathml_node_to_omml(over_node);
            if over_node
                .attributes
                .get("accent")
                .map(|value| value == "true")
                .unwrap_or(false)
            {
                let accent_chr = mathml_token_text(over_node);
                if !accent_chr.is_empty() {
                    return format!(
                        "<m:acc><m:accPr><m:chr m:val=\"{}\"/></m:accPr>{}</m:acc>",
                        escape_xml(&accent_chr),
                        wrap_omml_arg("e", base)
                    );
                }
            }

            format!(
                "<m:sSup>{}{}</m:sSup>",
                wrap_omml_arg("e", base),
                wrap_omml_arg("sup", over)
            )
        }
        "munder" => {
            let Some(base_node) = node.children.first() else {
                return render_mathml_nodes_to_omml(&node.children);
            };
            let Some(under_node) = node.children.get(1) else {
                return render_mathml_nodes_to_omml(&node.children);
            };

            if is_nary_operator_node(base_node) {
                return render_omml_nary(
                    &mathml_token_text(base_node),
                    Some(render_mathml_node_to_omml(under_node)),
                    None,
                    String::new(),
                );
            }
            if is_limit_like_operator_node(base_node) {
                let base = render_mathml_node_to_omml(base_node);
                let sub = render_mathml_node_to_omml(under_node);
                return render_limit_like_operator(base, Some(sub), None);
            }

            let base = render_mathml_node_to_omml(base_node);
            let under = render_mathml_node_to_omml(under_node);
            format!(
                "<m:sSub>{}{}</m:sSub>",
                wrap_omml_arg("e", base),
                wrap_omml_arg("sub", under)
            )
        }
        "munderover" => {
            let Some(base_node) = node.children.first() else {
                return render_mathml_nodes_to_omml(&node.children);
            };
            let Some(under_node) = node.children.get(1) else {
                return render_mathml_nodes_to_omml(&node.children);
            };
            let Some(over_node) = node.children.get(2) else {
                return render_mathml_nodes_to_omml(&node.children);
            };

            if is_nary_operator_node(base_node) {
                return render_omml_nary(
                    &mathml_token_text(base_node),
                    Some(render_mathml_node_to_omml(under_node)),
                    Some(render_mathml_node_to_omml(over_node)),
                    String::new(),
                );
            }
            if is_limit_like_operator_node(base_node) {
                let base = render_mathml_node_to_omml(base_node);
                let sub = render_mathml_node_to_omml(under_node);
                let sup = render_mathml_node_to_omml(over_node);
                return render_limit_like_operator(base, Some(sub), Some(sup));
            }

            let base = render_mathml_node_to_omml(base_node);
            let under = render_mathml_node_to_omml(under_node);
            let over = render_mathml_node_to_omml(over_node);
            format!(
                "<m:sSubSup>{}{}{}</m:sSubSup>",
                wrap_omml_arg("e", base),
                wrap_omml_arg("sub", under),
                wrap_omml_arg("sup", over)
            )
        }
        "mspace" => {
            let width = node
                .attributes
                .get("width")
                .map(String::as_str)
                .unwrap_or("");
            if width.trim().is_empty() || width.trim() == "0" || width.trim() == "0em" {
                String::new()
            } else {
                render_omml_run(" ")
            }
        }
        _ => {
            if !node.children.is_empty() {
                render_mathml_nodes_to_omml(&node.children)
            } else {
                let text = normalize_math_token_text(&node.text, node.name.as_str());
                if text.is_empty() {
                    String::new()
                } else {
                    render_omml_run(&text)
                }
            }
        }
    }
}

fn render_mathml_mrow(node: &MathMlNode) -> String {
    let significant: Vec<&MathMlNode> = node
        .children
        .iter()
        .filter(|child| !is_mathml_whitespace_node(child))
        .collect();

    if significant.len() == 3 && significant[1].name == "mtable" {
        let open = mathml_token_text(significant[0]);
        let close = mathml_token_text(significant[2]);
        if !open.is_empty() && !close.is_empty() {
            return render_mathml_table(significant[1], Some(open.as_str()), Some(close.as_str()));
        }
    }

    let mut out = String::new();
    let mut index = 0usize;
    while index < significant.len() {
        if let Some(next_node) = significant.get(index + 1) {
            if let Some(arg_extremum) = render_arg_extremum_pair(significant[index], next_node) {
                out.push_str(&arg_extremum);
                index += 2;
                continue;
            }
        }

        if let Some((operator_node, sub_node, sup_node)) = extract_nary_limits(significant[index]) {
            out.push_str(&render_omml_nary(
                &mathml_token_text(operator_node),
                sub_node.map(render_mathml_node_to_omml),
                sup_node.map(render_mathml_node_to_omml),
                render_mathml_node_refs_to_omml(&significant[index + 1..]),
            ));
            return out;
        }

        out.push_str(&render_mathml_node_to_omml(significant[index]));
        index += 1;
    }

    if out.is_empty() {
        render_mathml_nodes_to_omml(&node.children)
    } else {
        out
    }
}

fn render_mathml_node_refs_to_omml(nodes: &[&MathMlNode]) -> String {
    let mut out = String::new();
    for node in nodes {
        out.push_str(&render_mathml_node_to_omml(node));
    }
    out
}

fn extract_nary_limits<'a>(
    node: &'a MathMlNode,
) -> Option<(
    &'a MathMlNode,
    Option<&'a MathMlNode>,
    Option<&'a MathMlNode>,
)> {
    match node.name.as_str() {
        "msub" | "munder" => {
            let base = node.children.first()?;
            let sub = node.children.get(1)?;
            if is_nary_operator_node(base) {
                Some((base, Some(sub), None))
            } else {
                None
            }
        }
        "msup" | "mover" => {
            let base = node.children.first()?;
            let sup = node.children.get(1)?;
            if is_nary_operator_node(base) {
                Some((base, None, Some(sup)))
            } else {
                None
            }
        }
        "msubsup" | "munderover" => {
            let base = node.children.first()?;
            let sub = node.children.get(1)?;
            let sup = node.children.get(2)?;
            if is_nary_operator_node(base) {
                Some((base, Some(sub), Some(sup)))
            } else {
                None
            }
        }
        _ if is_nary_operator_node(node) => Some((node, None, None)),
        _ => None,
    }
}

fn render_arg_extremum_pair(arg_node: &MathMlNode, limit_node: &MathMlNode) -> Option<String> {
    if !is_arg_prefix_node(arg_node) {
        return None;
    }

    let (operator, sub, sup) = extract_limit_like_limits(limit_node)?;
    let suffix = operator.to_ascii_lowercase();
    if !matches!(suffix.as_str(), "min" | "max") {
        return None;
    }

    let base = render_omml_run(&format!("arg{suffix}"));
    Some(render_limit_like_operator(
        base,
        sub.map(render_mathml_node_to_omml),
        sup.map(render_mathml_node_to_omml),
    ))
}

fn extract_limit_like_limits<'a>(
    node: &'a MathMlNode,
) -> Option<(String, Option<&'a MathMlNode>, Option<&'a MathMlNode>)> {
    match node.name.as_str() {
        "msub" | "munder" => {
            let base = node.children.first()?;
            let sub = node.children.get(1)?;
            if is_limit_like_operator_node(base) {
                Some((mathml_token_text(base), Some(sub), None))
            } else {
                None
            }
        }
        "msup" | "mover" => {
            let base = node.children.first()?;
            let sup = node.children.get(1)?;
            if is_limit_like_operator_node(base) {
                Some((mathml_token_text(base), None, Some(sup)))
            } else {
                None
            }
        }
        "msubsup" | "munderover" => {
            let base = node.children.first()?;
            let sub = node.children.get(1)?;
            let sup = node.children.get(2)?;
            if is_limit_like_operator_node(base) {
                Some((mathml_token_text(base), Some(sub), Some(sup)))
            } else {
                None
            }
        }
        _ if is_limit_like_operator_node(node) => Some((mathml_token_text(node), None, None)),
        _ => None,
    }
}

fn render_omml_nary(
    operator: &str,
    sub: Option<String>,
    sup: Option<String>,
    operand: String,
) -> String {
    let mut out = String::new();
    out.push_str("<m:nary><m:naryPr>");
    if !operator.is_empty() {
        out.push_str(&format!("<m:chr m:val=\"{}\"/>", escape_xml(operator)));
    }
    out.push_str("<m:limLoc m:val=\"undOvr\"/>");
    out.push_str("</m:naryPr>");
    out.push_str(&wrap_omml_arg("sub", sub.unwrap_or_default()));
    out.push_str(&wrap_omml_arg("sup", sup.unwrap_or_default()));
    out.push_str(&wrap_omml_arg("e", operand));
    out.push_str("</m:nary>");
    out
}

fn render_omml_lim_low(base: String, lim: String) -> String {
    format!(
        "<m:limLow>{}{}</m:limLow>",
        wrap_omml_arg("e", base),
        wrap_omml_arg("lim", lim)
    )
}

fn render_omml_lim_upp(base: String, lim: String) -> String {
    format!(
        "<m:limUpp>{}{}</m:limUpp>",
        wrap_omml_arg("e", base),
        wrap_omml_arg("lim", lim)
    )
}

fn render_limit_like_operator(base: String, sub: Option<String>, sup: Option<String>) -> String {
    match (sub, sup) {
        (Some(sub), None) => render_omml_lim_low(base, sub),
        (None, Some(sup)) => render_omml_lim_upp(base, sup),
        (Some(sub), Some(sup)) => format!(
            "<m:sSubSup>{}{}{}</m:sSubSup>",
            wrap_omml_arg("e", base),
            wrap_omml_arg("sub", sub),
            wrap_omml_arg("sup", sup)
        ),
        (None, None) => base,
    }
}

fn render_mathml_table(table: &MathMlNode, open: Option<&str>, close: Option<&str>) -> String {
    let rows: Vec<&MathMlNode> = table
        .children
        .iter()
        .filter(|child| child.name == "mtr")
        .collect();

    if rows.is_empty() {
        return render_mathml_nodes_to_omml(&table.children);
    }

    let mut matrix = String::new();
    matrix.push_str("<m:m>");

    for row in rows {
        let cells: Vec<&MathMlNode> = row
            .children
            .iter()
            .filter(|child| child.name == "mtd")
            .collect();

        matrix.push_str("<m:mr>");
        if cells.is_empty() {
            matrix.push_str(&wrap_omml_arg(
                "e",
                render_mathml_nodes_to_omml(&row.children),
            ));
        } else {
            for cell in cells {
                matrix.push_str(&wrap_omml_arg(
                    "e",
                    render_mathml_nodes_to_omml(&cell.children),
                ));
            }
        }
        matrix.push_str("</m:mr>");
    }

    matrix.push_str("</m:m>");

    if let Some(open) = open {
        let close = close.unwrap_or(open);
        render_mathml_delimited(open, close, matrix)
    } else {
        matrix
    }
}

fn render_mathml_delimited(open: &str, close: &str, content: String) -> String {
    let open = if open == "." { "" } else { open };
    let close = if close == "." { "" } else { close };
    format!(
        "<m:d><m:dPr><m:begChr m:val=\"{}\"/><m:endChr m:val=\"{}\"/></m:dPr>{}</m:d>",
        escape_xml(open),
        escape_xml(close),
        wrap_omml_arg("e", content)
    )
}

fn wrap_omml_arg(tag: &str, content: String) -> String {
    let body = if content.trim().is_empty() {
        empty_omml_expression()
    } else {
        content
    };
    format!("<m:{tag}>{body}</m:{tag}>")
}

fn empty_omml_expression() -> String {
    render_omml_run("")
}

fn render_omml_run(text: &str) -> String {
    let preserve = text
        .chars()
        .next()
        .map(char::is_whitespace)
        .unwrap_or(false)
        || text
            .chars()
            .next_back()
            .map(char::is_whitespace)
            .unwrap_or(false)
        || text.contains("  ");
    if preserve {
        format!(
            "<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t xml:space=\"preserve\">{}</m:t></m:r>",
            escape_xml(text)
        )
    } else {
        format!(
            "<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>{}</m:t></m:r>",
            escape_xml(text)
        )
    }
}

fn mathml_token_text(node: &MathMlNode) -> String {
    normalize_math_token_text(&collect_mathml_text(node), node.name.as_str())
}

fn is_nary_operator_node(node: &MathMlNode) -> bool {
    matches!(
        mathml_token_text(node).as_str(),
        "∑" | "∏"
            | "∐"
            | "⋃"
            | "⋂"
            | "⋁"
            | "⋀"
            | "⨁"
            | "⨂"
            | "⨀"
            | "∫"
            | "∮"
            | "∯"
            | "∰"
    )
}

fn is_limit_like_operator_node(node: &MathMlNode) -> bool {
    matches!(
        mathml_token_text(node).to_ascii_lowercase().as_str(),
        "min" | "max" | "lim" | "sup" | "inf"
    )
}

fn is_arg_prefix_node(node: &MathMlNode) -> bool {
    mathml_token_text(node).eq_ignore_ascii_case("arg")
}

fn collect_mathml_text(node: &MathMlNode) -> String {
    let mut out = node.text.clone();
    for child in &node.children {
        out.push_str(&collect_mathml_text(child));
    }
    out
}

fn normalize_math_token_text(value: &str, node_name: &str) -> String {
    let cleaned = value
        .replace('\u{2061}', "")
        .replace('\u{2062}', "")
        .replace('\u{2063}', "")
        .replace('\u{2064}', "")
        .replace('\u{00A0}', " ");
    if node_name == "mtext" {
        cleaned
    } else {
        cleaned.trim().to_string()
    }
}

fn is_mathml_whitespace_node(node: &MathMlNode) -> bool {
    if node.name == "mspace" {
        return true;
    }
    if node.name == "mtext" || node.name == "mi" || node.name == "mn" || node.name == "mo" {
        return mathml_token_text(node).trim().is_empty();
    }
    false
}

fn render_image_run(
    rel_id: &str,
    docpr_id: usize,
    image_name: &str,
    alt_text: &str,
    width_emu: i64,
    height_emu: i64,
) -> String {
    format!(
        "<w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"{width_emu}\" cy=\"{height_emu}\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/><wp:docPr id=\"{docpr_id}\" name=\"{}\" descr=\"{}\"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/></wp:cNvGraphicFramePr><a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"{}\" descr=\"{}\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"{}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{width_emu}\" cy=\"{height_emu}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>",
        escape_xml(image_name),
        escape_xml(alt_text),
        escape_xml(image_name),
        escape_xml(alt_text),
        escape_xml(rel_id)
    )
}

struct LoadedImage {
    name: String,
    extension: String,
    content_type: String,
    bytes: Vec<u8>,
    width_emu: i64,
    height_emu: i64,
}

fn load_image(
    src: &str,
    markdown_base_dir: &Path,
    allow_remote_images: bool,
    warnings: &mut Vec<ConversionWarning>,
) -> Option<LoadedImage> {
    let is_remote = src.starts_with("http://") || src.starts_with("https://");

    let bytes = if is_remote {
        if !allow_remote_images {
            warnings.push(
                ConversionWarning::new(
                    WarningCode::RemoteImageBlocked,
                    format!(
                        "Remote image blocked by offline-by-default policy. Re-run with --allow-remote-images: {src}"
                    ),
                )
                .with_location(src),
            );
            return None;
        }

        match ureq::get(src).call() {
            Ok(mut response) => match response.body_mut().read_to_vec() {
                Ok(data) => data,
                Err(err) => {
                    warnings.push(
                        ConversionWarning::new(
                            WarningCode::ImageLoadFailed,
                            format!("Failed reading remote image bytes: {err}"),
                        )
                        .with_location(src),
                    );
                    return None;
                }
            },
            Err(UreqError::StatusCode(status)) => {
                warnings.push(
                    ConversionWarning::new(
                        WarningCode::ImageLoadFailed,
                        format!("Failed downloading remote image: HTTP {status}"),
                    )
                    .with_location(src),
                );
                return None;
            }
            Err(err) => {
                warnings.push(
                    ConversionWarning::new(
                        WarningCode::ImageLoadFailed,
                        format!("Failed requesting remote image: {err}"),
                    )
                    .with_location(src),
                );
                return None;
            }
        }
    } else {
        let (candidate, source_kind) = resolve_local_image_path(markdown_base_dir, src);
        match fs::read(&candidate) {
            Ok(data) => data,
            Err(err) => {
                warnings.push(
                    ConversionWarning::new(
                        WarningCode::ImageLoadFailed,
                        format!(
                            "Failed reading {source_kind} local image '{src}' (resolved to '{}'): {err}",
                            candidate.display()
                        ),
                    )
                    .with_location(candidate.display().to_string()),
                );
                return None;
            }
        }
    };

    let (extension, content_type) = detect_image_type(src, &bytes)
        .unwrap_or_else(|| ("png".to_string(), "image/png".to_string()));

    let name = Path::new(src)
        .file_name()
        .and_then(|n| n.to_str())
        .map(str::to_string)
        .unwrap_or_else(|| format!("image.{}", extension));

    Some(LoadedImage {
        name,
        extension,
        content_type,
        bytes,
        width_emu: 2_400_000,
        height_emu: 1_800_000,
    })
}

fn resolve_local_image_path(markdown_base_dir: &Path, src: &str) -> (PathBuf, &'static str) {
    let candidate = PathBuf::from(src);
    if candidate.is_absolute() {
        (candidate, "absolute")
    } else {
        (markdown_base_dir.join(&candidate), "relative")
    }
}

fn detect_image_type(src: &str, bytes: &[u8]) -> Option<(String, String)> {
    let ext = Path::new(src)
        .extension()
        .and_then(|e| e.to_str())
        .map(|e| e.to_ascii_lowercase());

    if let Some(ext) = ext {
        match ext.as_str() {
            "png" => return Some(("png".to_string(), "image/png".to_string())),
            "jpg" | "jpeg" => return Some(("jpg".to_string(), "image/jpeg".to_string())),
            "gif" => return Some(("gif".to_string(), "image/gif".to_string())),
            _ => {}
        }
    }

    if bytes.starts_with(&[0x89, b'P', b'N', b'G']) {
        return Some(("png".to_string(), "image/png".to_string()));
    }
    if bytes.starts_with(&[0xFF, 0xD8, 0xFF]) {
        return Some(("jpg".to_string(), "image/jpeg".to_string()));
    }
    if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        return Some(("gif".to_string(), "image/gif".to_string()));
    }

    None
}

fn load_template_package(
    template_path: Option<&Path>,
    warnings: &mut Vec<ConversionWarning>,
) -> Result<Option<TemplatePackage>> {
    if let Some(template_path) = template_path {
        if !template_path.exists() {
            warnings.push(
                ConversionWarning::new(
                    WarningCode::InvalidTemplate,
                    format!("Template not found: {}", template_path.display()),
                )
                .with_location(template_path.display().to_string()),
            );
            return Ok(None);
        }

        match read_template_package(template_path) {
            Ok(template) => return Ok(Some(template)),
            Err(err) => {
                warnings.push(
                    ConversionWarning::new(
                        WarningCode::InvalidTemplate,
                        format!("Unable to use template: {err}"),
                    )
                    .with_location(template_path.display().to_string()),
                );
            }
        }
    }

    Ok(None)
}

fn read_template_package(template_path: &Path) -> Result<TemplatePackage> {
    let file = fs::File::open(template_path).context("failed opening template")?;
    let mut archive = ZipArchive::new(file).context("failed reading template as zip")?;
    let mut entries = BTreeMap::new();

    for index in 0..archive.len() {
        let mut entry = archive
            .by_index(index)
            .with_context(|| format!("failed reading template zip entry at index {index}"))?;
        let name = entry.name().to_string();
        if name.ends_with('/') {
            continue;
        }

        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes)?;
        entries.insert(name, bytes);
    }

    if !entries.contains_key("word/styles.xml") {
        return Err(anyhow!("template is missing word/styles.xml"));
    }

    let document_relationships = entries
        .get("word/_rels/document.xml.rels")
        .map(|bytes| parse_relationships_xml(bytes))
        .transpose()?
        .unwrap_or_default();
    let section_properties_xml = entries
        .get("word/document.xml")
        .and_then(|bytes| extract_last_sect_pr_xml(bytes));

    Ok(TemplatePackage {
        entries,
        document_relationships,
        section_properties_xml,
    })
}

fn resolve_styles_xml(template: Option<&TemplatePackage>) -> Vec<u8> {
    template
        .and_then(|package| package.entries.get("word/styles.xml").cloned())
        .unwrap_or_else(|| default_styles_xml().as_bytes().to_vec())
}

fn resolve_numbering_xml(template: Option<&TemplatePackage>) -> Vec<u8> {
    template
        .and_then(|package| package.entries.get("word/numbering.xml").cloned())
        .unwrap_or_else(default_numbering_xml)
}

fn ensure_styles_relationship(state: &mut DocxBuildState) {
    if state.relationships.iter().any(|rel| {
        rel.rel_type == format!("{OFFICE_REL_NS}/styles")
            && rel.target.eq_ignore_ascii_case("styles.xml")
    }) {
        return;
    }

    let style_rel_id = state.next_rel_id();
    state.relationships.push(Relationship {
        id: style_rel_id,
        rel_type: format!("{OFFICE_REL_NS}/styles"),
        target: "styles.xml".to_string(),
        target_mode: None,
    });
}

fn ensure_numbering_relationship(state: &mut DocxBuildState) {
    if state
        .relationships
        .iter()
        .any(|rel| rel.rel_type == format!("{OFFICE_REL_NS}/numbering"))
    {
        return;
    }

    let numbering_rel_id = state.next_rel_id();
    state.relationships.push(Relationship {
        id: numbering_rel_id,
        rel_type: format!("{OFFICE_REL_NS}/numbering"),
        target: "numbering.xml".to_string(),
        target_mode: None,
    });
}

fn parse_relationships_xml(bytes: &[u8]) -> Result<Vec<Relationship>> {
    let xml =
        String::from_utf8(bytes.to_vec()).context("template relationship XML is not UTF-8")?;
    let mut relationships = Vec::new();
    let mut reader = Reader::from_str(&xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(start)) | Ok(Event::Start(start)) => {
                if local_name(start.name().as_ref()) == b"Relationship" {
                    let Some(id) = attr_value(&start, b"Id") else {
                        continue;
                    };
                    let Some(rel_type) = attr_value(&start, b"Type") else {
                        continue;
                    };
                    let Some(target) = attr_value(&start, b"Target") else {
                        continue;
                    };
                    let target_mode = attr_value(&start, b"TargetMode");

                    relationships.push(Relationship {
                        id,
                        rel_type,
                        target,
                        target_mode,
                    });
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(err) => return Err(anyhow!("failed parsing template relationships: {err}")),
        }
        buf.clear();
    }

    Ok(relationships)
}

fn parse_numeric_rel_id(value: &str) -> Option<usize> {
    value
        .strip_prefix("rId")
        .and_then(|suffix| suffix.parse::<usize>().ok())
}

fn extract_last_sect_pr_xml(document_xml: &[u8]) -> Option<String> {
    let xml = String::from_utf8_lossy(document_xml);
    let start = xml.rfind("<w:sectPr")?;
    let trailing = &xml[start..];

    if let Some(end_offset) = trailing.find("</w:sectPr>") {
        let end = start + end_offset + "</w:sectPr>".len();
        return Some(xml[start..end].to_string());
    }

    if let Some(end_offset) = trailing.find("/>") {
        let end = start + end_offset + 2;
        return Some(xml[start..end].to_string());
    }

    None
}

fn default_styles_xml() -> &'static str {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">
  <w:docDefaults>
    <w:rPrDefault><w:rPr/></w:rPrDefault>
    <w:pPrDefault><w:pPr/></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"Title\"><w:name w:val=\"Title\"/><w:basedOn w:val=\"Normal\"/><w:qFormat/><w:rPr><w:b/><w:sz w:val=\"48\"/></w:rPr></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Normal\"/><w:qFormat/><w:rPr><w:b/><w:sz w:val=\"32\"/></w:rPr></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"Heading2\"><w:name w:val=\"heading 2\"/><w:basedOn w:val=\"Normal\"/><w:qFormat/><w:rPr><w:b/><w:sz w:val=\"28\"/></w:rPr></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"Heading3\"><w:name w:val=\"heading 3\"/><w:basedOn w:val=\"Normal\"/><w:qFormat/><w:rPr><w:b/><w:sz w:val=\"24\"/></w:rPr></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"Heading4\"><w:name w:val=\"heading 4\"/><w:basedOn w:val=\"Normal\"/><w:qFormat/><w:rPr><w:b/><w:sz w:val=\"22\"/></w:rPr></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"Heading5\"><w:name w:val=\"heading 5\"/><w:basedOn w:val=\"Normal\"/><w:qFormat/><w:rPr><w:b/><w:sz w:val=\"20\"/></w:rPr></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"Heading6\"><w:name w:val=\"heading 6\"/><w:basedOn w:val=\"Normal\"/><w:qFormat/><w:rPr><w:b/><w:sz w:val=\"18\"/></w:rPr></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"Quote\"><w:name w:val=\"Quote\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:ind w:left=\"720\"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"Code\"><w:name w:val=\"Code\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:spacing w:line=\"240\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Consolas\" w:hAnsi=\"Consolas\"/><w:sz w:val=\"20\"/></w:rPr></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"Equation\"><w:name w:val=\"Equation\"/><w:basedOn w:val=\"Normal\"/></w:style>
  <w:style w:type=\"character\" w:styleId=\"EquationInline\"><w:name w:val=\"Equation Inline\"/></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"ListBullet\"><w:name w:val=\"List Bullet\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:ind w:left=\"720\"/></w:pPr></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"ListNumber\"><w:name w:val=\"List Number\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:ind w:left=\"720\"/></w:pPr></w:style>
  <w:style w:type=\"table\" w:styleId=\"Table\"><w:name w:val=\"Table\"/></w:style>
</w:styles>"
}

fn default_numbering_xml() -> Vec<u8> {
    let mut xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\
<w:abstractNum w:abstractNumId=\"{BULLET_LIST_ABSTRACT_NUM_ID}\">"
    );

    for level in 0..=LIST_MAX_LEVEL {
        let left_indent =
            LIST_BASE_INDENT_TWIPS + u32::from(level).saturating_mul(LIST_INDENT_STEP_TWIPS);
        xml.push_str(&format!(
            "<w:lvl w:ilvl=\"{level}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"bullet\"/><w:lvlText w:val=\"•\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"{left_indent}\" w:hanging=\"360\"/></w:pPr></w:lvl>"
        ));
    }

    xml.push_str("</w:abstractNum>");
    xml.push_str(&format!(
        "<w:abstractNum w:abstractNumId=\"{ORDERED_LIST_ABSTRACT_NUM_ID}\">"
    ));

    for level in 0..=LIST_MAX_LEVEL {
        let left_indent =
            LIST_BASE_INDENT_TWIPS + u32::from(level).saturating_mul(LIST_INDENT_STEP_TWIPS);
        let level_text = ordered_level_text(level);
        xml.push_str(&format!(
            "<w:lvl w:ilvl=\"{level}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"{}\"/><w:lvlJc w:val=\"left\"/><w:pPr><w:ind w:left=\"{left_indent}\" w:hanging=\"360\"/></w:pPr></w:lvl>",
            escape_xml(&level_text)
        ));
    }

    xml.push_str("</w:abstractNum>");
    xml.push_str(&format!(
        "<w:num w:numId=\"{ORDERED_LIST_NUM_ID}\"><w:abstractNumId w:val=\"{ORDERED_LIST_ABSTRACT_NUM_ID}\"/></w:num>"
    ));
    xml.push_str(&format!(
        "<w:num w:numId=\"{BULLET_LIST_NUM_ID}\"><w:abstractNumId w:val=\"{BULLET_LIST_ABSTRACT_NUM_ID}\"/></w:num>"
    ));
    xml.push_str("</w:numbering>");
    xml.into_bytes()
}

fn ordered_level_text(level: u8) -> String {
    let mut text = String::new();
    for index in 0..=level {
        if index > 0 {
            text.push('.');
        }
        text.push('%');
        text.push_str(&(u32::from(index) + 1).to_string());
    }
    text.push('.');
    text
}

fn build_content_types_xml(
    media_files: &[MediaFile],
    template_content_types: Option<&[u8]>,
) -> Vec<u8> {
    if let Some(template_xml) = template_content_types {
        return merge_content_types_with_media_defaults(template_xml, media_files);
    }

    let mut defaults = BTreeSet::new();
    defaults.insert((
        "rels".to_string(),
        "application/vnd.openxmlformats-package.relationships+xml".to_string(),
    ));
    defaults.insert(("xml".to_string(), "application/xml".to_string()));

    for media in media_files {
        defaults.insert((media.extension.clone(), media.content_type.clone()));
    }

    let mut xml = String::new();
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    xml.push_str(&format!("<Types xmlns=\"{CONTENT_TYPES_NS}\">"));

    for (extension, content_type) in defaults {
        xml.push_str(&format!(
            "<Default Extension=\"{}\" ContentType=\"{}\"/>",
            escape_xml(&extension),
            escape_xml(&content_type)
        ));
    }

    for (part_name, content_type) in required_content_type_overrides() {
        xml.push_str(&format!(
            "<Override PartName=\"{}\" ContentType=\"{}\"/>",
            escape_xml(part_name),
            escape_xml(content_type)
        ));
    }
    xml.push_str("</Types>");

    xml.into_bytes()
}

fn merge_content_types_with_media_defaults(
    template_xml: &[u8],
    media_files: &[MediaFile],
) -> Vec<u8> {
    let mut xml = String::from_utf8_lossy(template_xml).to_string();
    let close_tag = "</Types>";
    if let Some(close_index) = xml.rfind(close_tag) {
        let mut additions = String::new();
        for media in media_files {
            let marker = format!("Extension=\"{}\"", escape_xml(&media.extension));
            if !xml.contains(&marker) {
                additions.push_str(&format!(
                    "<Default Extension=\"{}\" ContentType=\"{}\"/>",
                    escape_xml(&media.extension),
                    escape_xml(&media.content_type)
                ));
            }
        }
        for (part_name, content_type) in required_content_type_overrides() {
            let marker = format!("PartName=\"{}\"", escape_xml(part_name));
            if !xml.contains(&marker) {
                additions.push_str(&format!(
                    "<Override PartName=\"{}\" ContentType=\"{}\"/>",
                    escape_xml(part_name),
                    escape_xml(content_type)
                ));
            }
        }
        xml.insert_str(close_index, &additions);
    }

    xml.into_bytes()
}

fn required_content_type_overrides() -> [(&'static str, &'static str); 5] {
    [
        (
            "/word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        ),
        (
            "/word/styles.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        ),
        (
            "/word/numbering.xml",
            WORDPROCESSINGML_NUMBERING_CONTENT_TYPE,
        ),
        (
            "/docProps/core.xml",
            "application/vnd.openxmlformats-package.core-properties+xml",
        ),
        (
            "/docProps/app.xml",
            "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        ),
    ]
}

fn build_package_relationships_xml() -> Vec<u8> {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"{PACKAGE_REL_NS}\"><Relationship Id=\"rId1\" Type=\"{OFFICE_REL_NS}/officeDocument\" Target=\"word/document.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/></Relationships>"
    ).into_bytes()
}

fn build_document_relationships_xml(relationships: &[Relationship]) -> Vec<u8> {
    let mut xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"{PACKAGE_REL_NS}\">"
    );

    for rel in relationships {
        if let Some(target_mode) = &rel.target_mode {
            xml.push_str(&format!(
                "<Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\" TargetMode=\"{}\"/>",
                escape_xml(&rel.id),
                escape_xml(&rel.rel_type),
                escape_xml(&rel.target),
                escape_xml(target_mode),
            ));
        } else {
            xml.push_str(&format!(
                "<Relationship Id=\"{}\" Type=\"{}\" Target=\"{}\"/>",
                escape_xml(&rel.id),
                escape_xml(&rel.rel_type),
                escape_xml(&rel.target),
            ));
        }
    }

    xml.push_str("</Relationships>");
    xml.into_bytes()
}

fn build_core_properties_xml() -> Vec<u8> {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:title>docwarp output</dc:title><dc:creator>docwarp</dc:creator></cp:coreProperties>".as_bytes().to_vec()
}

fn build_app_properties_xml() -> Vec<u8> {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application>docwarp</Application></Properties>".as_bytes().to_vec()
}

pub fn read_docx(
    input_path: &Path,
    options: &DocxReadOptions,
) -> Result<(Document, Vec<ConversionWarning>)> {
    let mut warnings = Vec::new();

    let password = options
        .password
        .as_deref()
        .map(str::trim)
        .filter(|p| !p.is_empty());
    let is_password_protected = is_password_protected_docx(input_path)?;
    if is_password_protected && password.is_none() {
        return Err(anyhow!(
            "DOCX appears password-protected. Re-run with --password or guided mode password prompt."
        ));
    }

    let mut _decrypted_tempdir = None;
    let archive_input_path = if is_password_protected {
        let tempdir = tempfile::tempdir().context("failed creating temporary decrypt directory")?;
        let decrypted_path = tempdir.path().join("decrypted.docx");
        decrypt_password_protected_docx(
            input_path,
            &decrypted_path,
            password.expect("password is required for protected DOCX"),
        )?;
        _decrypted_tempdir = Some(tempdir);
        decrypted_path
    } else {
        input_path.to_path_buf()
    };

    let file = fs::File::open(&archive_input_path)
        .with_context(|| format!("failed opening DOCX file: {}", input_path.display()))?;
    let mut archive = ZipArchive::new(file).context("failed opening DOCX zip archive")?;

    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")
        .context("DOCX is missing word/document.xml")?
        .read_to_string(&mut document_xml)
        .context("failed reading word/document.xml")?;

    let relationships = read_relationships(&mut archive)?;
    let image_targets = extract_image_assets(
        &mut archive,
        &relationships,
        &options.assets_dir,
        &mut warnings,
    )?;

    let mut blocks = Vec::new();
    let mut paragraph: Option<ParseParagraph> = None;
    let mut table: Option<ParseTable> = None;
    let mut run_style = RunStyle::default();
    let mut current_hyperlink: Option<(String, Vec<Inline>)> = None;
    let mut pending_list: Option<PendingList> = None;
    let mut in_word_text_node = false;
    let mut in_math_text_node = false;
    let mut math_para_depth = 0usize;
    let mut current_equation: Option<EquationCapture> = None;
    let mut unsupported_equation_warning_emitted = false;

    let mut reader = Reader::from_str(&document_xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(start)) => {
                let name = start.name().as_ref().to_vec();
                if is_math_tag(&name, b"oMathPara") {
                    math_para_depth = math_para_depth.saturating_add(1);
                }
                if is_math_tag(&name, b"oMath") {
                    begin_equation_capture(&mut current_equation, math_para_depth > 0);
                }
                if is_math_tag(&name, b"t") && current_equation.is_some() {
                    in_math_text_node = true;
                }
                mark_equation_unsupported_if_needed(&name, &mut current_equation);

                match local_name(&name) {
                    b"t" => {
                        if is_word_tag(&name, b"t") {
                            in_word_text_node = true;
                        }
                    }
                    b"p" => {
                        if is_word_tag(&name, b"p") {
                            paragraph = Some(ParseParagraph::default());
                        }
                    }
                    b"pStyle" => {
                        if is_word_tag(&name, b"pStyle") {
                            if let Some(value) = attr_value(&start, b"val") {
                                if let Some(paragraph) = paragraph.as_mut() {
                                    paragraph.style = Some(value);
                                }
                            }
                        }
                    }
                    b"ind" => {
                        if is_word_tag(&name, b"ind") {
                            if let Some(paragraph) = paragraph.as_mut() {
                                let raw = attr_value(&start, b"left")
                                    .or_else(|| attr_value(&start, b"start"));
                                paragraph.indent_left =
                                    raw.and_then(|value| parse_twips_value(&value));
                            }
                        }
                    }
                    b"r" => {
                        if is_word_tag(&name, b"r") {
                            run_style = RunStyle::default();
                        }
                    }
                    b"b" => {
                        if is_word_tag(&name, b"b") {
                            run_style.bold = true;
                        }
                    }
                    b"i" => {
                        if is_word_tag(&name, b"i") {
                            run_style.italic = true;
                        }
                    }
                    b"rStyle" => {
                        if is_word_tag(&name, b"rStyle") {
                            if let Some(value) = attr_value(&start, b"val") {
                                if value.contains("Code") {
                                    run_style.code = true;
                                }
                            }
                        }
                    }
                    b"hyperlink" => {
                        if is_word_tag(&name, b"hyperlink") {
                            if let Some(rel_id) = attr_value(&start, b"id") {
                                if let Some(url) = relationships.get(&rel_id) {
                                    current_hyperlink = Some((url.clone(), Vec::new()));
                                }
                            }
                        }
                    }
                    b"br" => {
                        if is_word_tag(&name, b"br") {
                            push_inline_target(Inline::LineBreak, &mut paragraph, &mut table);
                        }
                    }
                    b"tbl" => {
                        if is_word_tag(&name, b"tbl") {
                            flush_pending_list(&mut pending_list, &mut blocks);
                            table = Some(ParseTable::default());
                        }
                    }
                    b"tr" => {
                        if is_word_tag(&name, b"tr") {
                            if let Some(table) = table.as_mut() {
                                table.current_row.clear();
                            }
                        }
                    }
                    b"tc" => {
                        if is_word_tag(&name, b"tc") {
                            if let Some(table) = table.as_mut() {
                                table.current_cell.clear();
                            }
                        }
                    }
                    b"blip" => {
                        if let Some(rel_id) = attr_value(&start, b"embed") {
                            if let Some(src) = image_targets.get(&rel_id) {
                                push_inline_target(
                                    Inline::Image {
                                        alt: String::new(),
                                        src: src.clone(),
                                        title: None,
                                    },
                                    &mut paragraph,
                                    &mut table,
                                );
                            } else {
                                warnings.push(
                                    ConversionWarning::new(
                                        WarningCode::MissingMedia,
                                        format!(
                                            "Image relationship missing extracted media target: {rel_id}"
                                        ),
                                    )
                                    .with_location(rel_id),
                                );
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Empty(start)) => {
                let name = start.name().as_ref().to_vec();
                if is_math_tag(&name, b"oMathPara") {
                    math_para_depth = math_para_depth.saturating_add(1);
                }
                if is_math_tag(&name, b"oMath") {
                    begin_equation_capture(&mut current_equation, math_para_depth > 0);
                }
                if is_math_tag(&name, b"t") && current_equation.is_some() {
                    in_math_text_node = true;
                }
                mark_equation_unsupported_if_needed(&name, &mut current_equation);

                match local_name(&name) {
                    b"pStyle" => {
                        if is_word_tag(&name, b"pStyle") {
                            if let Some(value) = attr_value(&start, b"val") {
                                if let Some(paragraph) = paragraph.as_mut() {
                                    paragraph.style = Some(value);
                                }
                            }
                        }
                    }
                    b"ind" => {
                        if is_word_tag(&name, b"ind") {
                            if let Some(paragraph) = paragraph.as_mut() {
                                let raw = attr_value(&start, b"left")
                                    .or_else(|| attr_value(&start, b"start"));
                                paragraph.indent_left =
                                    raw.and_then(|value| parse_twips_value(&value));
                            }
                        }
                    }
                    b"b" => {
                        if is_word_tag(&name, b"b") {
                            run_style.bold = true;
                        }
                    }
                    b"i" => {
                        if is_word_tag(&name, b"i") {
                            run_style.italic = true;
                        }
                    }
                    b"rStyle" => {
                        if is_word_tag(&name, b"rStyle") {
                            if let Some(value) = attr_value(&start, b"val") {
                                if value.contains("Code") {
                                    run_style.code = true;
                                }
                            }
                        }
                    }
                    b"br" => {
                        if is_word_tag(&name, b"br") {
                            push_inline_target(Inline::LineBreak, &mut paragraph, &mut table);
                        }
                    }
                    b"blip" => {
                        if let Some(rel_id) = attr_value(&start, b"embed") {
                            if let Some(src) = image_targets.get(&rel_id) {
                                push_inline_target(
                                    Inline::Image {
                                        alt: String::new(),
                                        src: src.clone(),
                                        title: None,
                                    },
                                    &mut paragraph,
                                    &mut table,
                                );
                            } else {
                                warnings.push(
                                    ConversionWarning::new(
                                        WarningCode::MissingMedia,
                                        format!(
                                            "Image relationship missing extracted media target: {rel_id}"
                                        ),
                                    )
                                    .with_location(rel_id),
                                );
                            }
                        }
                    }
                    _ => {}
                }

                if is_math_tag(&name, b"t") {
                    in_math_text_node = false;
                }
                if is_math_tag(&name, b"oMath") {
                    finalize_equation_capture(
                        &mut current_equation,
                        &mut paragraph,
                        &mut table,
                        &mut warnings,
                        &mut unsupported_equation_warning_emitted,
                    );
                    in_math_text_node = false;
                }
                if is_math_tag(&name, b"oMathPara") {
                    math_para_depth = math_para_depth.saturating_sub(1);
                }
            }
            Ok(Event::Text(text)) => {
                let decoded = decode_text(&reader, text)?;
                if decoded.is_empty() {
                    buf.clear();
                    continue;
                }

                if in_math_text_node {
                    if let Some(equation) = current_equation.as_mut() {
                        equation.text.push_str(&decoded);
                    }
                    buf.clear();
                    continue;
                }

                if !in_word_text_node {
                    buf.clear();
                    continue;
                }

                if let Some((display, tex)) = parse_equation_marker(&decoded) {
                    apply_equation_marker(display, tex, &mut paragraph, &mut table);
                    buf.clear();
                    continue;
                }

                let styled = apply_run_style(decoded, &run_style);
                if let Some((_, link_inlines)) = current_hyperlink.as_mut() {
                    link_inlines.push(styled);
                } else {
                    push_inline_target(styled, &mut paragraph, &mut table);
                }
            }
            Ok(Event::End(end)) => {
                let name = end.name().as_ref().to_vec();
                if is_math_tag(&name, b"oMath") {
                    finalize_equation_capture(
                        &mut current_equation,
                        &mut paragraph,
                        &mut table,
                        &mut warnings,
                        &mut unsupported_equation_warning_emitted,
                    );
                    in_math_text_node = false;
                }
                if is_math_tag(&name, b"oMathPara") {
                    math_para_depth = math_para_depth.saturating_sub(1);
                }

                match local_name(&name) {
                    b"t" => {
                        if is_word_tag(&name, b"t") {
                            in_word_text_node = false;
                        }
                        if is_math_tag(&name, b"t") {
                            in_math_text_node = false;
                        }
                    }
                    b"hyperlink" => {
                        if is_word_tag(&name, b"hyperlink") {
                            if let Some((url, text)) = current_hyperlink.take() {
                                push_inline_target(
                                    Inline::Link { text, url },
                                    &mut paragraph,
                                    &mut table,
                                );
                            }
                        }
                    }
                    b"p" => {
                        if is_word_tag(&name, b"p") {
                            if let Some(paragraph) = paragraph.take() {
                                if let Some(table) = table.as_mut() {
                                    if !table.current_cell.is_empty()
                                        && !paragraph.inlines.is_empty()
                                    {
                                        table.current_cell.push(Inline::LineBreak);
                                    }
                                    table.current_cell.extend(paragraph.inlines);
                                } else {
                                    classify_paragraph(
                                        paragraph,
                                        &options.style_map,
                                        &mut pending_list,
                                        &mut blocks,
                                    );
                                }
                            }
                        }
                    }
                    b"tc" => {
                        if is_word_tag(&name, b"tc") {
                            if let Some(table) = table.as_mut() {
                                table
                                    .current_row
                                    .push(std::mem::take(&mut table.current_cell));
                            }
                        }
                    }
                    b"tr" => {
                        if is_word_tag(&name, b"tr") {
                            if let Some(table) = table.as_mut() {
                                table.rows.push(std::mem::take(&mut table.current_row));
                            }
                        }
                    }
                    b"tbl" => {
                        if is_word_tag(&name, b"tbl") {
                            if let Some(table) = table.take() {
                                let mut rows = table.rows;
                                if !rows.is_empty() {
                                    let mut headers = rows.remove(0);
                                    normalize_table_dimensions(&mut headers, &mut rows, 0);
                                    blocks.push(Block::Table { headers, rows });
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(err) => {
                return Err(anyhow!("failed parsing word/document.xml: {err}"));
            }
        }

        buf.clear();
    }

    if current_equation.is_some() {
        if let Some(equation) = current_equation.as_mut() {
            equation.unsupported = true;
            equation.depth = 1;
        }
        finalize_equation_capture(
            &mut current_equation,
            &mut paragraph,
            &mut table,
            &mut warnings,
            &mut unsupported_equation_warning_emitted,
        );
    }

    flush_pending_list(&mut pending_list, &mut blocks);

    Ok((Document { blocks }, warnings))
}

pub fn is_password_protected_docx(input_path: &Path) -> Result<bool> {
    let mut file = fs::File::open(input_path)
        .with_context(|| format!("failed opening DOCX file: {}", input_path.display()))?;
    let mut magic = [0_u8; 8];
    let read = file
        .read(&mut magic)
        .with_context(|| format!("failed reading DOCX header: {}", input_path.display()))?;
    if read < magic.len() {
        return Ok(false);
    }

    Ok(magic == [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1])
}

fn decrypt_password_protected_docx(
    input_path: &Path,
    output_path: &Path,
    password: &str,
) -> Result<()> {
    match decrypt_password_protected_docx_with_python(input_path, output_path, password) {
        Ok(()) => Ok(()),
        Err(PythonDocxDecryptError::IncorrectPassword) => Err(anyhow!("incorrect DOCX password")),
        Err(PythonDocxDecryptError::PythonNotFound) => Err(anyhow!(
            "unable to decrypt password-protected DOCX: python3/python not found. Install Python 3 to enable managed decryptor bootstrap.",
        )),
        Err(PythonDocxDecryptError::BootstrapFailed { details }) => Err(anyhow!(
            "unable to prepare managed Python decryptor runtime: {details}",
        )),
        Err(PythonDocxDecryptError::HashMismatch { expected, actual }) => Err(anyhow!(
            "unable to prepare managed Python decryptor runtime: package hash mismatch (expected {expected}, got {actual})",
        )),
        Err(PythonDocxDecryptError::LaunchFailed { python, source }) => Err(anyhow!(
            "unable to decrypt password-protected DOCX: failed launching {python}: {source}",
        )),
        Err(PythonDocxDecryptError::Failed { python, details }) => Err(anyhow!(
            "unable to decrypt password-protected DOCX: failed with {python}: {details}",
        )),
        Err(PythonDocxDecryptError::InvalidOutput { python }) => Err(anyhow!(
            "unable to decrypt password-protected DOCX: {python} returned output that is not a valid DOCX archive",
        )),
    }
}

#[derive(Debug)]
enum PythonDocxDecryptError {
    IncorrectPassword,
    PythonNotFound,
    BootstrapFailed {
        details: String,
    },
    HashMismatch {
        expected: &'static str,
        actual: String,
    },
    LaunchFailed {
        python: String,
        source: io::Error,
    },
    Failed {
        python: String,
        details: String,
    },
    InvalidOutput {
        python: String,
    },
}

fn decrypt_password_protected_docx_with_python(
    input_path: &Path,
    output_path: &Path,
    password: &str,
) -> std::result::Result<(), PythonDocxDecryptError> {
    let python = ensure_managed_python_with_msoffcrypto()?;
    let python_label = python.display().to_string();

    let script = r#"
import os
import sys

import msoffcrypto

input_path = sys.argv[1]
output_path = sys.argv[2]
password = os.environ.get("DOCWARP_DOCX_PASSWORD", "")

try:
    with open(input_path, "rb") as source:
        office = msoffcrypto.OfficeFile(source)
        office.load_key(password=password)
        with open(output_path, "wb") as target:
            office.decrypt(target)
except Exception as exc:
    message = str(exc)
    sys.stderr.write(message)
    if "password" in message.lower():
        sys.exit(4)
    sys.exit(2)
"#;

    let output = process::Command::new(&python)
        .arg("-c")
        .arg(script)
        .arg(input_path)
        .arg(output_path)
        .env("DOCWARP_DOCX_PASSWORD", password)
        .output()
        .map_err(|source| PythonDocxDecryptError::LaunchFailed {
            python: python_label.clone(),
            source,
        })?;

    if output.status.success() {
        let decrypted_bytes = match fs::read(output_path) {
            Ok(bytes) => bytes,
            Err(err) => {
                return Err(PythonDocxDecryptError::Failed {
                    python: python_label,
                    details: format!("failed reading decrypted output: {err}"),
                });
            }
        };
        if !is_valid_decrypted_docx_archive(&decrypted_bytes) {
            return Err(PythonDocxDecryptError::InvalidOutput {
                python: python.display().to_string(),
            });
        }
        return Ok(());
    }

    let details = command_failure_details(&output);
    match output.status.code() {
        Some(4) => Err(PythonDocxDecryptError::IncorrectPassword),
        _ => Err(PythonDocxDecryptError::Failed {
            python: python.display().to_string(),
            details,
        }),
    }
}

fn ensure_managed_python_with_msoffcrypto() -> std::result::Result<PathBuf, PythonDocxDecryptError>
{
    let venv_dir = managed_python_venv_dir();
    let managed_python = managed_python_executable(&venv_dir);

    match probe_msoffcrypto_tool(&managed_python)? {
        MsoffcryptoToolProbe::Version(version) if version == MSOFFCRYPTO_TOOL_VERSION => {
            return Ok(managed_python);
        }
        _ => {}
    }

    let bootstrap_python = find_bootstrap_python()?;

    if !managed_python.exists() {
        if let Some(parent) = venv_dir.parent() {
            fs::create_dir_all(parent).map_err(|err| PythonDocxDecryptError::BootstrapFailed {
                details: format!(
                    "failed creating managed runtime directory '{}': {err}",
                    parent.display()
                ),
            })?;
        }

        let output = process::Command::new(&bootstrap_python)
            .arg("-m")
            .arg("venv")
            .arg(&venv_dir)
            .output()
            .map_err(|source| PythonDocxDecryptError::LaunchFailed {
                python: bootstrap_python.display().to_string(),
                source,
            })?;
        if !output.status.success() {
            return Err(PythonDocxDecryptError::BootstrapFailed {
                details: format!(
                    "failed creating managed python virtualenv at '{}': {}",
                    venv_dir.display(),
                    command_failure_details(&output)
                ),
            });
        }
    }

    install_msoffcrypto_with_hash(&managed_python)?;

    match probe_msoffcrypto_tool(&managed_python)? {
        MsoffcryptoToolProbe::Version(version) if version == MSOFFCRYPTO_TOOL_VERSION => {
            Ok(managed_python)
        }
        MsoffcryptoToolProbe::Version(version) => Err(PythonDocxDecryptError::BootstrapFailed {
            details: format!(
                "managed runtime at '{}' reports msoffcrypto-tool=={} after installation (expected {})",
                managed_python.display(),
                version,
                MSOFFCRYPTO_TOOL_VERSION
            ),
        }),
        MsoffcryptoToolProbe::Missing => Err(PythonDocxDecryptError::BootstrapFailed {
            details: format!(
                "managed runtime at '{}' does not expose msoffcrypto-tool after installation",
                managed_python.display()
            ),
        }),
    }
}

fn install_msoffcrypto_with_hash(
    managed_python: &Path,
) -> std::result::Result<(), PythonDocxDecryptError> {
    let tempdir = tempfile::tempdir().map_err(|err| PythonDocxDecryptError::BootstrapFailed {
        details: format!("failed creating temporary python bootstrap directory: {err}"),
    })?;

    let output = process::Command::new(managed_python)
        .arg("-m")
        .arg("pip")
        .arg("download")
        .arg("--disable-pip-version-check")
        .arg("--no-input")
        .arg("--no-deps")
        .arg("--only-binary=:all:")
        .arg("--dest")
        .arg(tempdir.path())
        .arg(format!("msoffcrypto-tool=={MSOFFCRYPTO_TOOL_VERSION}"))
        .output()
        .map_err(|source| PythonDocxDecryptError::LaunchFailed {
            python: managed_python.display().to_string(),
            source,
        })?;
    if !output.status.success() {
        return Err(PythonDocxDecryptError::BootstrapFailed {
            details: format!(
                "failed downloading managed decryptor package: {}",
                command_failure_details(&output)
            ),
        });
    }

    let wheel = find_downloaded_wheel(tempdir.path())?;
    let hash = sha256_hex_file(&wheel)?;
    if hash != MSOFFCRYPTO_TOOL_WHEEL_SHA256 {
        return Err(PythonDocxDecryptError::HashMismatch {
            expected: MSOFFCRYPTO_TOOL_WHEEL_SHA256,
            actual: hash,
        });
    }

    let output = process::Command::new(managed_python)
        .arg("-m")
        .arg("pip")
        .arg("install")
        .arg("--disable-pip-version-check")
        .arg("--no-input")
        .arg("--upgrade")
        .arg(&wheel)
        .output()
        .map_err(|source| PythonDocxDecryptError::LaunchFailed {
            python: managed_python.display().to_string(),
            source,
        })?;
    if !output.status.success() {
        return Err(PythonDocxDecryptError::BootstrapFailed {
            details: format!(
                "failed installing managed decryptor package: {}",
                command_failure_details(&output)
            ),
        });
    }

    Ok(())
}

fn find_downloaded_wheel(
    download_dir: &Path,
) -> std::result::Result<PathBuf, PythonDocxDecryptError> {
    let mut wheel = None;
    let entries =
        fs::read_dir(download_dir).map_err(|err| PythonDocxDecryptError::BootstrapFailed {
            details: format!(
                "failed reading managed decryptor download directory '{}': {err}",
                download_dir.display()
            ),
        })?;

    for entry in entries {
        let entry = entry.map_err(|err| PythonDocxDecryptError::BootstrapFailed {
            details: format!("failed iterating managed decryptor download directory: {err}"),
        })?;
        let path = entry.path();
        if path
            .extension()
            .and_then(|ext| ext.to_str())
            .is_some_and(|ext| ext.eq_ignore_ascii_case("whl"))
        {
            wheel = Some(path);
            break;
        }
    }

    wheel.ok_or_else(|| PythonDocxDecryptError::BootstrapFailed {
        details: "managed decryptor bootstrap did not download a wheel artifact".to_string(),
    })
}

fn sha256_hex_file(path: &Path) -> std::result::Result<String, PythonDocxDecryptError> {
    let mut file = fs::File::open(path).map_err(|err| PythonDocxDecryptError::BootstrapFailed {
        details: format!(
            "failed reading downloaded package '{}': {err}",
            path.display()
        ),
    })?;
    let mut hasher = Sha256::new();
    let mut buffer = [0_u8; 8192];
    loop {
        let read =
            file.read(&mut buffer)
                .map_err(|err| PythonDocxDecryptError::BootstrapFailed {
                    details: format!(
                        "failed hashing downloaded package '{}': {err}",
                        path.display()
                    ),
                })?;
        if read == 0 {
            break;
        }
        hasher.update(&buffer[..read]);
    }
    Ok(format!("{:x}", hasher.finalize()))
}

#[derive(Debug)]
enum MsoffcryptoToolProbe {
    Missing,
    Version(String),
}

fn probe_msoffcrypto_tool(
    python: &Path,
) -> std::result::Result<MsoffcryptoToolProbe, PythonDocxDecryptError> {
    if !python.exists() {
        return Ok(MsoffcryptoToolProbe::Missing);
    }

    let script = r#"
import sys
try:
    import importlib.metadata as md
except Exception:
    try:
        import importlib_metadata as md
    except Exception:
        sys.exit(3)
for name in ("msoffcrypto-tool", "msoffcrypto_tool"):
    try:
        sys.stdout.write(md.version(name))
        sys.exit(0)
    except Exception:
        pass
sys.exit(3)
"#;

    let output = process::Command::new(python)
        .arg("-c")
        .arg(script)
        .output()
        .map_err(|source| PythonDocxDecryptError::LaunchFailed {
            python: python.display().to_string(),
            source,
        })?;

    match output.status.code() {
        Some(0) => {
            let version = String::from_utf8_lossy(&output.stdout).trim().to_string();
            if version.is_empty() {
                Err(PythonDocxDecryptError::Failed {
                    python: python.display().to_string(),
                    details: "managed runtime package probe produced an empty version string"
                        .to_string(),
                })
            } else {
                Ok(MsoffcryptoToolProbe::Version(version))
            }
        }
        Some(3) => Ok(MsoffcryptoToolProbe::Missing),
        _ => Err(PythonDocxDecryptError::Failed {
            python: python.display().to_string(),
            details: command_failure_details(&output),
        }),
    }
}

fn find_bootstrap_python() -> std::result::Result<PathBuf, PythonDocxDecryptError> {
    for candidate in ["python3", "python"] {
        match process::Command::new(candidate)
            .arg("-c")
            .arg("import sys")
            .output()
        {
            Ok(output) if output.status.success() => return Ok(PathBuf::from(candidate)),
            Ok(_) => continue,
            Err(err) if err.kind() == io::ErrorKind::NotFound => continue,
            Err(source) => {
                return Err(PythonDocxDecryptError::LaunchFailed {
                    python: candidate.to_string(),
                    source,
                });
            }
        }
    }
    Err(PythonDocxDecryptError::PythonNotFound)
}

fn managed_python_venv_dir() -> PathBuf {
    managed_runtime_root()
        .join("python")
        .join(format!("msoffcrypto-tool-{MSOFFCRYPTO_TOOL_VERSION}"))
}

fn managed_runtime_root() -> PathBuf {
    if let Some(explicit) = env::var_os("DOCWARP_HOME") {
        return PathBuf::from(explicit);
    }

    #[cfg(windows)]
    if let Some(local_app_data) = env::var_os("LOCALAPPDATA") {
        return PathBuf::from(local_app_data).join("docwarp");
    }

    if let Some(xdg_data_home) = env::var_os("XDG_DATA_HOME") {
        return PathBuf::from(xdg_data_home).join("docwarp");
    }

    if let Some(home) = env::var_os("HOME") {
        return PathBuf::from(home)
            .join(".local")
            .join("share")
            .join("docwarp");
    }

    PathBuf::from(".docwarp")
}

#[cfg(windows)]
fn managed_python_executable(venv_dir: &Path) -> PathBuf {
    venv_dir.join("Scripts").join("python.exe")
}

#[cfg(not(windows))]
fn managed_python_executable(venv_dir: &Path) -> PathBuf {
    venv_dir.join("bin").join("python")
}

fn command_failure_details(output: &process::Output) -> String {
    let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
    let stdout = String::from_utf8_lossy(&output.stdout).trim().to_string();
    if !stderr.is_empty() {
        stderr
    } else if !stdout.is_empty() {
        stdout
    } else {
        format!("process exited with status {}", output.status)
    }
}

fn is_valid_decrypted_docx_archive(bytes: &[u8]) -> bool {
    let cursor = io::Cursor::new(bytes);
    let mut archive = match ZipArchive::new(cursor) {
        Ok(archive) => archive,
        Err(_) => return false,
    };

    archive.by_name("word/document.xml").is_ok()
}

fn read_relationships<R: Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
) -> Result<BTreeMap<String, String>> {
    let mut rels_xml = String::new();

    if let Ok(mut rels) = archive.by_name("word/_rels/document.xml.rels") {
        rels.read_to_string(&mut rels_xml)
            .context("failed reading word/_rels/document.xml.rels")?;
    } else {
        return Ok(BTreeMap::new());
    }

    let mut relationships = BTreeMap::new();
    let mut reader = Reader::from_str(&rels_xml);
    reader.config_mut().trim_text(true);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Empty(start)) | Ok(Event::Start(start)) => {
                if local_name(start.name().as_ref()) == b"Relationship" {
                    if let (Some(id), Some(target)) =
                        (attr_value(&start, b"Id"), attr_value(&start, b"Target"))
                    {
                        relationships.insert(id, target);
                    }
                }
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(err) => {
                return Err(anyhow!("failed parsing document relationships: {err}"));
            }
        }

        buf.clear();
    }

    Ok(relationships)
}

fn extract_image_assets<R: Read + std::io::Seek>(
    archive: &mut ZipArchive<R>,
    relationships: &BTreeMap<String, String>,
    assets_dir: &Path,
    warnings: &mut Vec<ConversionWarning>,
) -> Result<BTreeMap<String, String>> {
    fs::create_dir_all(assets_dir)
        .with_context(|| format!("failed creating assets directory: {}", assets_dir.display()))?;

    let mut rel_to_output = BTreeMap::new();

    for (rel_id, target) in relationships {
        if !target.contains("media/") {
            continue;
        }

        let normalized = normalize_docx_target(target);
        let source_path = format!("word/{normalized}");

        let mut bytes = Vec::new();
        match archive.by_name(&source_path) {
            Ok(mut file) => {
                file.read_to_end(&mut bytes)
                    .with_context(|| format!("failed reading media entry: {source_path}"))?;

                let file_name = Path::new(&normalized)
                    .file_name()
                    .and_then(|n| n.to_str())
                    .ok_or_else(|| anyhow!("invalid media path in DOCX: {normalized}"))?
                    .to_string();

                let output_path = assets_dir.join(file_name);
                fs::write(&output_path, bytes).with_context(|| {
                    format!("failed writing extracted media: {}", output_path.display())
                })?;

                rel_to_output.insert(rel_id.clone(), path_to_markdown_link(&output_path));
            }
            Err(_) => {
                warnings.push(
                    ConversionWarning::new(
                        WarningCode::MissingMedia,
                        format!("Missing referenced media file: {source_path}"),
                    )
                    .with_location(source_path),
                );
            }
        }
    }

    Ok(rel_to_output)
}

fn classify_paragraph(
    paragraph: ParseParagraph,
    style_map: &StyleMap,
    pending_list: &mut Option<PendingList>,
    blocks: &mut Vec<Block>,
) {
    let ParseParagraph {
        style,
        indent_left,
        inlines,
    } = paragraph;
    let style = style.unwrap_or_else(|| "Normal".to_string());
    let token = style_map.md_token_for(&style);

    match token.as_str() {
        "list_bullet" | "list_number" => {
            let ordered = token == "list_number";
            if let Some(list) = pending_list.as_mut() {
                if list.base_indent_left.is_none() {
                    list.base_indent_left = indent_left;
                }
                let base_indent = list.base_indent_left.unwrap_or(LIST_BASE_INDENT_TWIPS);
                let item_indent = indent_left.unwrap_or(base_indent);
                let level = list_level_from_indent(item_indent, base_indent);
                let prev_level = list.levels.last().copied().unwrap_or(0);
                let prev_ordered = list.item_ordered.last().copied().unwrap_or(list.ordered);

                // Preserve top-level ordered/unordered list transitions as separate list blocks.
                if ordered != prev_ordered && level == 0 && prev_level == 0 {
                    flush_pending_list(pending_list, blocks);
                    *pending_list = Some(PendingList {
                        ordered,
                        base_indent_left: indent_left,
                        items: vec![inlines],
                        levels: vec![0],
                        item_ordered: vec![ordered],
                    });
                    return;
                }

                list.items.push(inlines);
                list.levels.push(level);
                list.item_ordered.push(ordered);
                return;
            }

            *pending_list = Some(PendingList {
                ordered,
                base_indent_left: indent_left,
                items: vec![inlines],
                levels: vec![0],
                item_ordered: vec![ordered],
            });
        }
        _ => {
            flush_pending_list(pending_list, blocks);

            if inlines.len() == 1 {
                if let Inline::Image { alt, src, title } = &inlines[0] {
                    blocks.push(Block::Image {
                        alt: alt.clone(),
                        src: src.clone(),
                        title: title.clone(),
                    });
                    return;
                }
            }

            let block = match token.as_str() {
                "title" => Block::Title(inlines),
                "h1" => Block::Heading {
                    level: 1,
                    content: inlines,
                },
                "h2" => Block::Heading {
                    level: 2,
                    content: inlines,
                },
                "h3" => Block::Heading {
                    level: 3,
                    content: inlines,
                },
                "h4" => Block::Heading {
                    level: 4,
                    content: inlines,
                },
                "h5" => Block::Heading {
                    level: 5,
                    content: inlines,
                },
                "h6" => Block::Heading {
                    level: 6,
                    content: inlines,
                },
                "quote" => Block::BlockQuote(inlines),
                "code" => {
                    let raw = inline_text(&inlines);
                    let (language, code) = extract_code_language_marker(raw);
                    Block::CodeBlock { language, code }
                }
                _ => Block::Paragraph(inlines),
            };

            blocks.push(block);
        }
    }
}

fn flush_pending_list(pending_list: &mut Option<PendingList>, blocks: &mut Vec<Block>) {
    if let Some(list) = pending_list.take() {
        blocks.push(Block::List {
            ordered: list.ordered,
            items: list.items,
            levels: list.levels,
            item_ordered: list.item_ordered,
        });
    }
}

fn push_inline_target(
    inline: Inline,
    paragraph: &mut Option<ParseParagraph>,
    table: &mut Option<ParseTable>,
) {
    if let Some(paragraph) = paragraph.as_mut() {
        paragraph.inlines.push(inline);
    } else if let Some(table) = table.as_mut() {
        table.current_cell.push(inline);
    }
}

fn apply_equation_marker(
    display: bool,
    tex: String,
    paragraph: &mut Option<ParseParagraph>,
    table: &mut Option<ParseTable>,
) {
    if let Some(paragraph) = paragraph.as_mut() {
        if let Some(Inline::Equation {
            tex: existing_tex,
            display: existing_display,
        }) = paragraph.inlines.last_mut()
        {
            *existing_tex = tex;
            *existing_display = display;
            return;
        }

        paragraph.inlines.push(Inline::Equation { tex, display });
        return;
    }

    if let Some(table) = table.as_mut() {
        if let Some(Inline::Equation {
            tex: existing_tex,
            display: existing_display,
        }) = table.current_cell.last_mut()
        {
            *existing_tex = tex;
            *existing_display = display;
            return;
        }

        table.current_cell.push(Inline::Equation { tex, display });
    }
}

fn begin_equation_capture(current: &mut Option<EquationCapture>, display: bool) {
    if let Some(capture) = current.as_mut() {
        capture.depth = capture.depth.saturating_add(1);
        capture.unsupported = true;
        return;
    }

    *current = Some(EquationCapture {
        display,
        text: String::new(),
        unsupported: false,
        depth: 1,
    });
}

fn mark_equation_unsupported_if_needed(name: &[u8], current: &mut Option<EquationCapture>) {
    let Some(capture) = current.as_mut() else {
        return;
    };
    if !is_math_prefixed(name) {
        return;
    }

    if !matches!(
        local_name(name),
        b"oMath"
            | b"oMathPara"
            | b"oMathParaPr"
            | b"r"
            | b"rPr"
            | b"ctrlPr"
            | b"sty"
            | b"jc"
            | b"f"
            | b"fPr"
            | b"type"
            | b"num"
            | b"den"
            | b"rad"
            | b"radPr"
            | b"degHide"
            | b"deg"
            | b"e"
            | b"sSub"
            | b"sSup"
            | b"sSubSup"
            | b"sub"
            | b"sup"
            | b"d"
            | b"dPr"
            | b"begChr"
            | b"endChr"
            | b"m"
            | b"mPr"
            | b"mr"
            | b"nary"
            | b"naryPr"
            | b"limLoc"
            | b"limLow"
            | b"limUpp"
            | b"lim"
            | b"acc"
            | b"accPr"
            | b"chr"
            | b"t"
    ) {
        capture.unsupported = true;
    }
}

fn finalize_equation_capture(
    current: &mut Option<EquationCapture>,
    paragraph: &mut Option<ParseParagraph>,
    table: &mut Option<ParseTable>,
    warnings: &mut Vec<ConversionWarning>,
    unsupported_equation_warning_emitted: &mut bool,
) {
    let Some(capture) = current.as_mut() else {
        return;
    };
    capture.depth = capture.depth.saturating_sub(1);
    if capture.depth != 0 {
        return;
    }

    let capture = current.take().expect("equation capture exists");
    let tex = capture.text.trim().to_string();
    if !tex.is_empty() {
        push_inline_target(
            Inline::Equation {
                tex,
                display: capture.display,
            },
            paragraph,
            table,
        );
    }
    if capture.unsupported && !*unsupported_equation_warning_emitted {
        warnings.push(ConversionWarning::new(
            WarningCode::UnsupportedFeature,
            "Encountered unsupported OMML equation styling/structure; flattened to linear text. Source DOCX equation styling remains unchanged.",
        ));
        *unsupported_equation_warning_emitted = true;
    }
}

fn apply_run_style(text: String, style: &RunStyle) -> Inline {
    let base = if style.code {
        Inline::Code(text)
    } else {
        Inline::Text(text)
    };

    let mut result = base;
    if style.italic {
        result = Inline::Emphasis(vec![result]);
    }
    if style.bold {
        result = Inline::Strong(vec![result]);
    }

    result
}

fn decode_text(reader: &Reader<&[u8]>, text: quick_xml::events::BytesText<'_>) -> Result<String> {
    match text.unescape()? {
        Cow::Borrowed(raw) => Ok(reader.decoder().decode(raw.as_bytes())?.to_string()),
        Cow::Owned(raw) => Ok(raw),
    }
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|b| *b == b':').next().unwrap_or(name)
}

fn namespace_prefix(name: &[u8]) -> Option<&[u8]> {
    let mut parts = name.splitn(2, |byte| *byte == b':');
    let prefix = parts.next()?;
    parts.next().map(|_| prefix)
}

fn is_word_tag(name: &[u8], tag: &[u8]) -> bool {
    local_name(name) == tag && matches!(namespace_prefix(name), None | Some(b"w"))
}

fn is_math_tag(name: &[u8], tag: &[u8]) -> bool {
    local_name(name) == tag && namespace_prefix(name) == Some(b"m")
}

fn is_math_prefixed(name: &[u8]) -> bool {
    namespace_prefix(name) == Some(b"m")
}

fn attr_value(start: &BytesStart<'_>, local_key: &[u8]) -> Option<String> {
    start
        .attributes()
        .flatten()
        .find(|attr| local_name(attr.key.as_ref()) == local_key)
        .and_then(|attr| String::from_utf8(attr.value.as_ref().to_vec()).ok())
}

fn parse_twips_value(value: &str) -> Option<u32> {
    value
        .trim()
        .parse::<i64>()
        .ok()
        .map(|raw| raw.max(0))
        .and_then(|raw| u32::try_from(raw).ok())
}

fn list_level_from_indent(item_indent_left: u32, base_indent_left: u32) -> u8 {
    if item_indent_left <= base_indent_left {
        return 0;
    }

    let delta = item_indent_left.saturating_sub(base_indent_left);
    let rounded_steps = (delta + (LIST_INDENT_STEP_TWIPS / 2)) / LIST_INDENT_STEP_TWIPS;
    u8::try_from(rounded_steps).unwrap_or(u8::MAX)
}

fn extract_code_language_marker(raw: String) -> (Option<String>, String) {
    parse_code_language_marker(&raw, CODE_LANG_MARKER_PREFIX).unwrap_or((None, raw))
}

fn parse_code_language_marker(raw: &str, prefix: &str) -> Option<(Option<String>, String)> {
    let without_prefix = raw.strip_prefix(prefix)?;

    let Some(end) = without_prefix.find(CODE_LANG_MARKER_SUFFIX) else {
        return Some((None, raw.to_string()));
    };

    let language = without_prefix[..end].trim();
    let code_start = end + CODE_LANG_MARKER_SUFFIX.len();
    let code = without_prefix[code_start..].to_string();

    Some(if language.is_empty() {
        (None, code)
    } else {
        (Some(language.to_string()), code)
    })
}

fn parse_equation_marker(raw: &str) -> Option<(bool, String)> {
    let raw = raw.trim();
    let payload = raw
        .strip_prefix(EQUATION_MARKER_PREFIX)?
        .strip_suffix(EQUATION_MARKER_SUFFIX)?;
    let (kind, encoded_tex) = payload.split_once(':')?;
    let display = match kind {
        "d" => true,
        "i" => false,
        _ => return None,
    };

    let tex = decode_hex(encoded_tex)?;
    Some((display, tex))
}

fn decode_hex(value: &str) -> Option<String> {
    if value.len() % 2 != 0 {
        return None;
    }

    let mut bytes = Vec::with_capacity(value.len() / 2);
    for chunk in value.as_bytes().chunks_exact(2) {
        let piece = std::str::from_utf8(chunk).ok()?;
        let byte = u8::from_str_radix(piece, 16).ok()?;
        bytes.push(byte);
    }

    String::from_utf8(bytes).ok()
}

fn normalize_docx_target(target: &str) -> String {
    let replaced = target.replace('\\', "/");
    replaced.trim_start_matches("../").to_string()
}

fn normalize_table_dimensions(
    headers: &mut Vec<Vec<Inline>>,
    rows: &mut [Vec<Vec<Inline>>],
    forced_width: usize,
) {
    let width = if forced_width > 0 {
        forced_width
    } else {
        headers
            .len()
            .max(rows.iter().map(Vec::len).max().unwrap_or_default())
    };

    if width == 0 {
        return;
    }

    headers.resize_with(width, Vec::new);
    for row in rows {
        row.resize_with(width, Vec::new);
    }
}

fn path_to_markdown_link(path: &Path) -> String {
    path.to_string_lossy().replace('\\', "/")
}

fn escape_xml(input: &str) -> String {
    input
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    fn tiny_png() -> Vec<u8> {
        vec![
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48,
            0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00,
            0x00, 0x90, 0x77, 0x53, 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, 0x08,
            0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00, 0xC9, 0xFE, 0x92,
            0xEF, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
        ]
    }

    #[test]
    fn detects_cfb_header_as_password_protected_docx() {
        let dir = tempdir().expect("tempdir should be created");
        let cfb_path = dir.path().join("encrypted.docx");
        fs::write(
            &cfb_path,
            [0xD0_u8, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00],
        )
        .expect("signature file should be written");

        let detected =
            is_password_protected_docx(&cfb_path).expect("header detection should succeed");
        assert!(detected);
    }

    #[test]
    fn regular_docx_is_not_marked_password_protected() {
        let dir = tempdir().expect("tempdir should be created");
        let output_docx = dir.path().join("plain.docx");
        let doc = Document {
            blocks: vec![Block::Paragraph(vec![Inline::Text("plain".into())])],
        };

        write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: None,
            },
        )
        .expect("DOCX write should succeed");

        let detected =
            is_password_protected_docx(&output_docx).expect("header detection should succeed");
        assert!(!detected);
    }

    #[test]
    fn decrypted_docx_archive_validation_rejects_non_zip_bytes() {
        assert!(!is_valid_decrypted_docx_archive(b"not a zip archive"));
    }

    #[test]
    fn decrypted_docx_archive_validation_accepts_real_docx_zip() {
        let dir = tempdir().expect("tempdir should be created");
        let output_docx = dir.path().join("plain.docx");
        let doc = Document {
            blocks: vec![Block::Paragraph(vec![Inline::Text("plain".into())])],
        };

        write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: None,
            },
        )
        .expect("DOCX write should succeed");

        let bytes = fs::read(&output_docx).expect("DOCX should be readable as bytes");
        assert!(is_valid_decrypted_docx_archive(&bytes));
    }

    #[test]
    fn write_and_read_docx_roundtrip_core_blocks() {
        let dir = tempdir().expect("tempdir should be created");
        let image_path = dir.path().join("image.png");
        fs::write(&image_path, tiny_png()).expect("image should be written");

        let doc = Document {
            blocks: vec![
                Block::Title(vec![Inline::Text("Doc title".into())]),
                Block::Paragraph(vec![Inline::Text("Hello".into())]),
                Block::List {
                    ordered: false,
                    items: vec![
                        vec![Inline::Text("a".into())],
                        vec![Inline::Text("b".into())],
                    ],
                    levels: Vec::new(),
                    item_ordered: Vec::new(),
                },
                Block::Image {
                    alt: "logo".into(),
                    src: image_path.to_string_lossy().to_string(),
                    title: None,
                },
            ],
        };

        let output_docx = dir.path().join("out.docx");
        let write_warnings = write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: None,
            },
        )
        .expect("DOCX write should succeed");

        assert!(write_warnings.is_empty());

        let output_assets = dir.path().join("assets");
        let (read_doc, read_warnings) = read_docx(
            &output_docx,
            &DocxReadOptions {
                assets_dir: output_assets,
                style_map: StyleMap::builtin(),
                password: None,
            },
        )
        .expect("DOCX read should succeed");

        assert!(read_warnings.is_empty());
        assert!(!read_doc.blocks.is_empty());
        assert!(
            read_doc
                .blocks
                .iter()
                .any(|block| matches!(block, Block::List { .. }))
        );
    }

    #[test]
    fn write_and_read_docx_roundtrip_preserves_heading_levels_h4_to_h6() {
        let dir = tempdir().expect("tempdir should be created");
        let output_docx = dir.path().join("out.docx");

        let doc = Document {
            blocks: vec![
                Block::Heading {
                    level: 4,
                    content: vec![Inline::Text("Level 4".into())],
                },
                Block::Heading {
                    level: 5,
                    content: vec![Inline::Text("Level 5".into())],
                },
                Block::Heading {
                    level: 6,
                    content: vec![Inline::Text("Level 6".into())],
                },
            ],
        };

        let write_warnings = write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: None,
            },
        )
        .expect("DOCX write should succeed");

        assert!(write_warnings.is_empty());

        let (read_doc, read_warnings) = read_docx(
            &output_docx,
            &DocxReadOptions {
                assets_dir: dir.path().join("assets"),
                style_map: StyleMap::builtin(),
                password: None,
            },
        )
        .expect("DOCX read should succeed");

        assert!(read_warnings.is_empty());

        let heading_levels: Vec<u8> = read_doc
            .blocks
            .iter()
            .filter_map(|block| match block {
                Block::Heading { level, .. } => Some(*level),
                _ => None,
            })
            .collect();

        assert_eq!(heading_levels, vec![4, 5, 6]);
    }

    #[test]
    fn write_and_read_docx_roundtrip_preserves_nested_list_levels() {
        let dir = tempdir().expect("tempdir should be created");
        let output_docx = dir.path().join("out.docx");

        let doc = Document {
            blocks: vec![Block::List {
                ordered: false,
                items: vec![
                    vec![Inline::Text("parent bullet".into())],
                    vec![Inline::Text("child bullet".into())],
                    vec![Inline::Text("parent number".into())],
                    vec![Inline::Text("child number".into())],
                ],
                levels: vec![0, 1, 0, 1],
                item_ordered: vec![false, false, true, true],
            }],
        };

        let write_warnings = write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: None,
            },
        )
        .expect("DOCX write should succeed");
        assert!(write_warnings.is_empty());

        let (read_doc, read_warnings) = read_docx(
            &output_docx,
            &DocxReadOptions {
                assets_dir: dir.path().join("assets"),
                style_map: StyleMap::builtin(),
                password: None,
            },
        )
        .expect("DOCX read should succeed");
        assert!(read_warnings.is_empty());

        let Some(Block::List {
            levels,
            item_ordered,
            ..
        }) = read_doc.blocks.first()
        else {
            panic!("expected first block to be a list");
        };

        assert_eq!(levels, &vec![0, 1, 0, 1]);
        assert_eq!(item_ordered, &vec![false, false, true, true]);
    }

    #[test]
    fn write_and_read_docx_roundtrip_preserves_code_block_language() {
        let dir = tempdir().expect("tempdir should be created");
        let output_docx = dir.path().join("out.docx");

        let doc = Document {
            blocks: vec![Block::CodeBlock {
                language: Some("rust".into()),
                code: "fn main() {\n    println!(\"hi\");\n}".into(),
            }],
        };

        let write_warnings = write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: None,
            },
        )
        .expect("DOCX write should succeed");
        assert!(write_warnings.is_empty());

        let (read_doc, read_warnings) = read_docx(
            &output_docx,
            &DocxReadOptions {
                assets_dir: dir.path().join("assets"),
                style_map: StyleMap::builtin(),
                password: None,
            },
        )
        .expect("DOCX read should succeed");
        assert!(read_warnings.is_empty());

        let Some(Block::CodeBlock { language, code }) = read_doc.blocks.first() else {
            panic!("expected first block to be a code block");
        };

        assert_eq!(language.as_deref(), Some("rust"));
        assert_eq!(code, "fn main() {\n    println!(\"hi\");\n}");
    }

    #[test]
    fn writes_native_omml_and_applies_equation_style_mapping() {
        let dir = tempdir().expect("tempdir should be created");
        let output_docx = dir.path().join("out.docx");

        let document = Document {
            blocks: vec![
                Block::Paragraph(vec![
                    Inline::Text("Inline ".into()),
                    Inline::Equation {
                        tex: "x^2 + \\frac{1}{\\sqrt{y}}".into(),
                        display: false,
                    },
                ]),
                Block::Paragraph(vec![
                    Inline::Text("Optimize ".into()),
                    Inline::Equation {
                        tex: "\\min_{\\beta} f(\\beta)".into(),
                        display: false,
                    },
                ]),
                Block::Paragraph(vec![Inline::Equation {
                    tex: "\\left[\\begin{matrix} a & b \\\\ c & d \\end{matrix}\\right]".into(),
                    display: true,
                }]),
                Block::Paragraph(vec![Inline::Equation {
                    tex: "\\sum_{i=1}^{n} x_i".into(),
                    display: true,
                }]),
            ],
        };

        let mut style_map = StyleMap::builtin();
        style_map
            .md_to_docx
            .insert("equation_inline".to_string(), "EqInline".to_string());
        style_map
            .md_to_docx
            .insert("equation_block".to_string(), "EqBlock".to_string());

        let warnings = write_docx(
            &document,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map,
                template: None,
            },
        )
        .expect("DOCX write should succeed");
        assert!(warnings.is_empty());

        let mut archive = ZipArchive::new(
            fs::File::open(&output_docx).expect("written docx should be readable as zip"),
        )
        .expect("written docx should be a valid zip");
        let mut document_xml = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document.xml should exist")
            .read_to_string(&mut document_xml)
            .expect("document.xml should be readable");

        assert!(
            document_xml.contains("<m:oMath>"),
            "equations should be emitted as OMML"
        );
        assert!(
            document_xml.contains("<m:oMathPara>"),
            "display equations should be emitted as OMML paragraph equations"
        );
        assert!(
            document_xml.contains("<m:sSup>"),
            "superscript equations should be emitted with structured OMML scripts"
        );
        assert!(
            document_xml.contains("<m:f>"),
            "fraction equations should be emitted with structured OMML fractions"
        );
        assert!(
            document_xml.contains("<m:rad>"),
            "sqrt equations should be emitted with structured OMML radicals"
        );
        assert!(
            document_xml.contains("<m:d>") && document_xml.contains("<m:m>"),
            "matrix equations with delimiters should be emitted as delimiter-wrapped OMML matrices"
        );
        assert!(
            document_xml.contains("<m:nary>") && document_xml.contains("m:limLoc m:val=\"undOvr\""),
            "display summations should emit n-ary OMML with under/over limits"
        );
        assert!(
            document_xml.contains("<m:limLow>"),
            "limit-like operators should emit limLow for improved operator typography"
        );
        assert!(
            !document_xml.contains("<m:mPr><m:begChr"),
            "matrix delimiters must not be placed in m:mPr"
        );
        assert!(
            document_xml.contains("<w:pStyle w:val=\"EqBlock\"/>"),
            "display equation paragraph should apply equation_block style"
        );
        assert!(
            !document_xml.contains("<m:ctrlPr>"),
            "OMML runs should avoid invalid control-property placement"
        );
    }

    #[test]
    fn renders_argmin_as_unified_limit_operator() {
        let rendered = render_structured_omml("\\arg\\min_{\\beta} f(\\beta)", "EqInline")
            .expect("structured OMML conversion should succeed")
            .expect("expression should produce OMML body");

        assert!(
            rendered.contains("<m:t>argmin</m:t>"),
            "arg + min should be collapsed into a single argmin operator token"
        );
        assert!(
            rendered.contains("<m:limLow>"),
            "argmin lower bounds should render as limLow"
        );
        assert!(
            !rendered.contains("<m:t>arg</m:t>"),
            "arg should not be emitted as a standalone token before min limits"
        );
    }

    #[test]
    fn reads_inline_omml_equation_as_inline_math() {
        let dir = tempdir().expect("tempdir should be created");
        let input_docx = dir.path().join("inline-omml.docx");
        write_minimal_docx_with_document_xml(
            &input_docx,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p>
      <m:oMath>
        <m:r><m:t>x</m:t></m:r>
        <m:r><m:t>=</m:t></m:r>
        <m:r><m:t>1</m:t></m:r>
      </m:oMath>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>"#,
        )
        .expect("fixture docx should be written");

        let (document, warnings) = read_docx(
            &input_docx,
            &DocxReadOptions {
                assets_dir: dir.path().join("assets"),
                style_map: StyleMap::builtin(),
                password: None,
            },
        )
        .expect("DOCX read should succeed");

        assert!(warnings.is_empty());
        let Some(Block::Paragraph(inlines)) = document.blocks.first() else {
            panic!("expected paragraph block");
        };
        assert_eq!(
            inlines,
            &vec![Inline::Equation {
                tex: "x=1".to_string(),
                display: false
            }]
        );
    }

    #[test]
    fn reads_display_omml_equation_as_display_math() {
        let dir = tempdir().expect("tempdir should be created");
        let input_docx = dir.path().join("display-omml.docx");
        write_minimal_docx_with_document_xml(
            &input_docx,
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
  <w:body>
    <w:p>
      <m:oMathPara>
        <m:oMath>
          <m:r><m:t>E=mc^2</m:t></m:r>
        </m:oMath>
      </m:oMathPara>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>"#,
        )
        .expect("fixture docx should be written");

        let (document, warnings) = read_docx(
            &input_docx,
            &DocxReadOptions {
                assets_dir: dir.path().join("assets"),
                style_map: StyleMap::builtin(),
                password: None,
            },
        )
        .expect("DOCX read should succeed");

        assert!(warnings.is_empty());
        let Some(Block::Paragraph(inlines)) = document.blocks.first() else {
            panic!("expected paragraph block");
        };
        assert_eq!(
            inlines,
            &vec![Inline::Equation {
                tex: "E=mc^2".to_string(),
                display: true
            }]
        );
    }

    #[test]
    fn warns_and_flattens_unsupported_omml_structures() {
        let dir = tempdir().expect("tempdir should be created");
        let input_docx = dir.path().join("unsupported-omml.docx");
        write_minimal_docx_with_document_xml(
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
        )
        .expect("fixture docx should be written");

        let (document, warnings) = read_docx(
            &input_docx,
            &DocxReadOptions {
                assets_dir: dir.path().join("assets"),
                style_map: StyleMap::builtin(),
                password: None,
            },
        )
        .expect("DOCX read should succeed");

        assert!(
            warnings
                .iter()
                .any(|warning| warning.code == WarningCode::UnsupportedFeature),
            "unsupported OMML should emit unsupported_feature warning"
        );
        assert!(
            warnings.iter().any(|warning| {
                warning.code == WarningCode::UnsupportedFeature
                    && warning.message.contains("styling remains unchanged")
            }),
            "unsupported OMML warning should explain source styling is unchanged"
        );
        let Some(Block::Paragraph(inlines)) = document.blocks.first() else {
            panic!("expected paragraph block");
        };
        assert_eq!(
            inlines,
            &vec![Inline::Equation {
                tex: "x".to_string(),
                display: false
            }]
        );
    }

    #[test]
    fn warns_once_for_multiple_unsupported_omml_equations() {
        let dir = tempdir().expect("tempdir should be created");
        let input_docx = dir.path().join("multiple-unsupported-omml.docx");
        write_minimal_docx_with_document_xml(
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
    <w:p>
      <m:oMath>
        <m:groupChr>
          <m:e><m:r><m:t>y</m:t></m:r></m:e>
        </m:groupChr>
      </m:oMath>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>"#,
        )
        .expect("fixture docx should be written");

        let (_document, warnings) = read_docx(
            &input_docx,
            &DocxReadOptions {
                assets_dir: dir.path().join("assets"),
                style_map: StyleMap::builtin(),
                password: None,
            },
        )
        .expect("DOCX read should succeed");

        let unsupported_warning_count = warnings
            .iter()
            .filter(|warning| warning.code == WarningCode::UnsupportedFeature)
            .count();
        assert_eq!(
            unsupported_warning_count, 1,
            "multiple unsupported equations should emit a single unsupported_feature warning"
        );
    }

    #[test]
    fn extract_code_language_marker_accepts_docwarp_prefix() {
        let raw = "[[docwarp-code-lang:rust]]fn main() {}".to_string();
        let (language, code) = super::extract_code_language_marker(raw);

        assert_eq!(language.as_deref(), Some("rust"));
        assert_eq!(code, "fn main() {}");
    }

    #[test]
    fn remote_images_warn_when_disabled() {
        let dir = tempdir().expect("tempdir should be created");
        let output_docx = dir.path().join("out.docx");

        let doc = Document {
            blocks: vec![Block::Paragraph(vec![Inline::Image {
                alt: "remote".into(),
                src: "https://example.com/a.png".into(),
                title: None,
            }])],
        };

        let warnings = write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: None,
            },
        )
        .expect("DOCX write should succeed with warnings");

        assert!(
            warnings
                .iter()
                .any(|warning| warning.code == WarningCode::RemoteImageBlocked)
        );
        assert!(
            warnings
                .iter()
                .any(|warning| warning.message.contains("offline-by-default")),
            "remote-image warning should explain offline default policy"
        );
    }

    #[test]
    fn resolves_relative_image_paths_against_markdown_base_dir() {
        let dir = tempdir().expect("tempdir should be created");
        let nested = dir.path().join("assets");
        fs::create_dir_all(&nested).expect("assets directory should be created");
        let image_path = nested.join("tiny.png");
        fs::write(&image_path, tiny_png()).expect("image should be written");

        let mut warnings = Vec::new();
        let loaded = load_image("assets/tiny.png", dir.path(), false, &mut warnings);
        assert!(
            loaded.is_some(),
            "relative image should load from markdown base"
        );
        assert!(
            warnings.is_empty(),
            "loading relative image should not warn"
        );
    }

    #[test]
    fn loads_absolute_image_paths_without_base_joining() {
        let dir = tempdir().expect("tempdir should be created");
        let other_base = tempdir().expect("tempdir should be created");
        let image_path = dir.path().join("tiny.png");
        fs::write(&image_path, tiny_png()).expect("image should be written");

        let mut warnings = Vec::new();
        let loaded = load_image(
            image_path.to_string_lossy().as_ref(),
            other_base.path(),
            false,
            &mut warnings,
        );
        assert!(
            loaded.is_some(),
            "absolute image path should be read directly"
        );
        assert!(
            warnings.is_empty(),
            "loading absolute image should not warn"
        );
    }

    #[test]
    fn uses_dotx_template_styles_when_available() {
        let dir = tempdir().expect("tempdir should be created");
        let template_path = dir.path().join("custom.dotx");
        let output_docx = dir.path().join("out.docx");

        write_template_zip(
            &template_path,
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="BrandStyle"><w:name w:val="BrandStyle"/></w:style>
</w:styles>"#,
        )
        .expect("template should be written");

        let doc = Document {
            blocks: vec![Block::Paragraph(vec![Inline::Text("Body".into())])],
        };

        let warnings = write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: Some(template_path),
            },
        )
        .expect("DOCX write should succeed");
        assert!(warnings.is_empty(), "valid .dotx template should not warn");

        let mut archive = ZipArchive::new(
            fs::File::open(&output_docx).expect("written docx should be readable as zip"),
        )
        .expect("written docx should be a valid zip");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles.xml should exist")
            .read_to_string(&mut styles)
            .expect("styles.xml should be readable");

        assert!(
            styles.contains("BrandStyle"),
            "expected template styles to be copied into output DOCX"
        );
    }

    #[test]
    fn invalid_dotx_falls_back_to_builtin_styles_with_warning() {
        let dir = tempdir().expect("tempdir should be created");
        let template_path = dir.path().join("broken.dotx");
        let output_docx = dir.path().join("out.docx");

        write_invalid_template_zip(&template_path).expect("invalid template should be written");

        let doc = Document {
            blocks: vec![Block::Paragraph(vec![Inline::Text("Body".into())])],
        };

        let warnings = write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: Some(template_path),
            },
        )
        .expect("DOCX write should succeed with warning fallback");

        assert!(
            warnings
                .iter()
                .any(|warning| warning.code == WarningCode::InvalidTemplate),
            "invalid .dotx should emit invalid_template warning"
        );

        let mut archive = ZipArchive::new(
            fs::File::open(&output_docx).expect("written docx should be readable as zip"),
        )
        .expect("written docx should be a valid zip");
        let mut styles = String::new();
        archive
            .by_name("word/styles.xml")
            .expect("styles.xml should exist")
            .read_to_string(&mut styles)
            .expect("styles.xml should be readable");

        assert!(
            styles.contains("ListBullet"),
            "fallback styles should include built-in style definitions"
        );
    }

    #[test]
    fn preserves_template_sections_headers_footers_and_related_parts() {
        let dir = tempdir().expect("tempdir should be created");
        let template_path = dir.path().join("full-template.dotx");
        let output_docx = dir.path().join("out.docx");

        let mut entries = BTreeMap::new();
        entries.insert(
            "[Content_Types].xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"#
                .to_vec(),
        );
        entries.insert(
            "_rels/.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"#
                .to_vec(),
        );
        entries.insert("docProps/core.xml".to_string(), build_core_properties_xml());
        entries.insert("docProps/app.xml".to_string(), build_app_properties_xml());
        entries.insert(
            "word/styles.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"#
                .to_vec(),
        );
        entries.insert(
            "word/document.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>template body</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdHeaderDefault"/>
      <w:footerReference w:type="default" r:id="rIdFooterDefault"/>
      <w:pgSz w:w="12240" w:h="15840"/>
    </w:sectPr>
  </w:body>
</w:document>"#
                .to_vec(),
        );
        entries.insert(
            "word/_rels/document.xml.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rIdHeaderDefault" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdFooterDefault" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rIdSettings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rIdFontTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
  <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>"#
                .to_vec(),
        );
        entries.insert(
            "word/header1.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Template Header Marker</w:t></w:r></w:p>
</w:hdr>"#
                .to_vec(),
        );
        entries.insert(
            "word/_rels/header1.xml.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWatermarkImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/watermark.png"/>
</Relationships>"#
                .to_vec(),
        );
        entries.insert(
            "word/footer1.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Template Footer Marker</w:t></w:r></w:p>
</w:ftr>"#
                .to_vec(),
        );
        entries.insert(
            "word/theme/theme1.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="TemplateTheme"/>"#
                .to_vec(),
        );
        entries.insert(
            "word/settings.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="120"/></w:settings>"#
                .to_vec(),
        );
        entries.insert(
            "word/fontTable.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:font w:name="Calibri"/></w:fonts>"#
                .to_vec(),
        );
        entries.insert(
            "word/numbering.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:numbering>"#
                .to_vec(),
        );
        entries.insert("word/media/watermark.png".to_string(), tiny_png());

        write_template_entries_zip(&template_path, &entries).expect("template should be written");

        let doc = Document {
            blocks: vec![Block::Paragraph(vec![Inline::Text("Body".into())])],
        };

        let warnings = write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: Some(template_path),
            },
        )
        .expect("DOCX write should succeed");
        assert!(warnings.is_empty(), "valid template should not warn");

        let mut archive = ZipArchive::new(
            fs::File::open(&output_docx).expect("written docx should be readable as zip"),
        )
        .expect("written docx should be a valid zip");

        for required_part in [
            "word/header1.xml",
            "word/footer1.xml",
            "word/theme/theme1.xml",
            "word/settings.xml",
            "word/fontTable.xml",
            "word/numbering.xml",
            "word/_rels/header1.xml.rels",
            "word/media/watermark.png",
        ] {
            archive
                .by_name(required_part)
                .unwrap_or_else(|_| panic!("expected output to include {required_part}"));
        }

        let mut document_xml = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document.xml should exist")
            .read_to_string(&mut document_xml)
            .expect("document.xml should be readable");
        assert!(
            document_xml.contains("rIdHeaderDefault"),
            "document.xml should preserve template section header reference"
        );
        assert!(
            document_xml.contains("rIdFooterDefault"),
            "document.xml should preserve template section footer reference"
        );

        let mut rels = String::new();
        archive
            .by_name("word/_rels/document.xml.rels")
            .expect("document.xml.rels should exist")
            .read_to_string(&mut rels)
            .expect("document.xml.rels should be readable");
        assert!(
            rels.contains("Target=\"header1.xml\""),
            "document rels should preserve header relationship"
        );
        assert!(
            rels.contains("Target=\"footer1.xml\""),
            "document rels should preserve footer relationship"
        );
        assert!(
            rels.contains("Target=\"theme/theme1.xml\""),
            "document rels should preserve theme relationship"
        );
    }

    #[test]
    fn writes_document_relationship_for_styles_xml() {
        let dir = tempdir().expect("tempdir should be created");
        let output_docx = dir.path().join("out.docx");
        let doc = Document {
            blocks: vec![Block::Heading {
                level: 2,
                content: vec![Inline::Text("Overview".into())],
            }],
        };

        let warnings = write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: None,
            },
        )
        .expect("DOCX write should succeed");
        assert!(warnings.is_empty());

        let mut archive = ZipArchive::new(
            fs::File::open(&output_docx).expect("written docx should be readable as zip"),
        )
        .expect("written docx should be a valid zip");
        let mut rels = String::new();
        archive
            .by_name("word/_rels/document.xml.rels")
            .expect("document.xml.rels should exist")
            .read_to_string(&mut rels)
            .expect("document.xml.rels should be readable");

        assert!(
            rels.contains("relationships/styles"),
            "document relationships should include the styles relationship type"
        );
        assert!(
            rels.contains("Target=\"styles.xml\""),
            "document relationships should target styles.xml"
        );
    }

    #[test]
    fn lists_emit_numbering_metadata_and_parts() {
        let dir = tempdir().expect("tempdir should be created");
        let output_docx = dir.path().join("out.docx");
        let doc = Document {
            blocks: vec![
                Block::List {
                    ordered: false,
                    items: vec![vec![Inline::Text("bullet".into())]],
                    levels: vec![0],
                    item_ordered: vec![false],
                },
                Block::List {
                    ordered: true,
                    items: vec![vec![Inline::Text("numbered".into())]],
                    levels: vec![0],
                    item_ordered: vec![true],
                },
            ],
        };

        let warnings = write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: None,
            },
        )
        .expect("DOCX write should succeed");
        assert!(warnings.is_empty());

        let mut archive = ZipArchive::new(
            fs::File::open(&output_docx).expect("written docx should be readable as zip"),
        )
        .expect("written docx should be a valid zip");

        let mut document_xml = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document.xml should exist")
            .read_to_string(&mut document_xml)
            .expect("document.xml should be readable");
        assert!(
            document_xml.contains("<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr>"),
            "bullet list items should emit w:numPr with bullet numId"
        );
        assert!(
            document_xml.contains("<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>"),
            "ordered list items should emit w:numPr with ordered numId"
        );

        let mut numbering_xml = String::new();
        archive
            .by_name("word/numbering.xml")
            .expect("numbering.xml should exist")
            .read_to_string(&mut numbering_xml)
            .expect("numbering.xml should be readable");
        assert!(
            numbering_xml.contains("<w:num w:numId=\"1\"><w:abstractNumId w:val=\"1\"/></w:num>"),
            "ordered numbering definition should be present"
        );
        assert!(
            numbering_xml.contains("<w:num w:numId=\"2\"><w:abstractNumId w:val=\"2\"/></w:num>"),
            "bullet numbering definition should be present"
        );
        assert!(
            numbering_xml.contains("w:lvlText w:val=\"%1.%2.\""),
            "ordered level 2 should render hierarchical numbering text"
        );
        assert!(
            numbering_xml.contains("w:lvlText w:val=\"%1.%2.%3.\""),
            "ordered level 3 should render hierarchical numbering text"
        );

        let mut rels = String::new();
        archive
            .by_name("word/_rels/document.xml.rels")
            .expect("document.xml.rels should exist")
            .read_to_string(&mut rels)
            .expect("document.xml.rels should be readable");
        assert!(
            rels.contains("relationships/numbering"),
            "document relationships should include numbering relationship type"
        );
        assert!(
            rels.contains("Target=\"numbering.xml\""),
            "document relationships should include numbering.xml target"
        );

        let mut content_types = String::new();
        archive
            .by_name("[Content_Types].xml")
            .expect("[Content_Types].xml should exist")
            .read_to_string(&mut content_types)
            .expect("[Content_Types].xml should be readable");
        assert!(
            content_types.contains("PartName=\"/word/numbering.xml\""),
            "content types should include numbering override part name"
        );
        assert!(
            content_types.contains(WORDPROCESSINGML_NUMBERING_CONTENT_TYPE),
            "content types should include numbering override content type"
        );
    }

    #[test]
    fn normalizes_uneven_table_rows_after_roundtrip() {
        let dir = tempdir().expect("tempdir should be created");
        let output_docx = dir.path().join("out.docx");

        let doc = Document {
            blocks: vec![Block::Table {
                headers: vec![
                    vec![Inline::Text("A".into())],
                    vec![Inline::Text("B".into())],
                ],
                rows: vec![
                    vec![vec![Inline::Text("1".into())]],
                    vec![
                        vec![Inline::Text("2".into())],
                        vec![Inline::Text("3".into())],
                        vec![Inline::Text("4".into())],
                    ],
                ],
            }],
        };

        write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: None,
            },
        )
        .expect("DOCX write should succeed");

        let (read_doc, warnings) = read_docx(
            &output_docx,
            &DocxReadOptions {
                assets_dir: dir.path().join("assets"),
                style_map: StyleMap::builtin(),
                password: None,
            },
        )
        .expect("DOCX read should succeed");
        assert!(warnings.is_empty());

        let Some(Block::Table { headers, rows }) = read_doc.blocks.first() else {
            panic!("expected first block to be a table");
        };

        assert_eq!(headers.len(), 3);
        assert_eq!(rows[0].len(), 3);
        assert_eq!(rows[1].len(), 3);
    }

    #[test]
    fn markdown_paragraphs_emit_docx_spacing_to_preserve_blank_line_flow() {
        let dir = tempdir().expect("tempdir should be created");
        let output_docx = dir.path().join("out.docx");
        let doc = Document {
            blocks: vec![
                Block::Paragraph(vec![Inline::Text("first".into())]),
                Block::Paragraph(vec![Inline::Text("second".into())]),
            ],
        };

        let warnings = write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: None,
            },
        )
        .expect("DOCX write should succeed");
        assert!(warnings.is_empty());

        let mut archive = ZipArchive::new(
            fs::File::open(&output_docx).expect("written docx should be readable as zip"),
        )
        .expect("written docx should be a valid zip");
        let mut document_xml = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document.xml should exist")
            .read_to_string(&mut document_xml)
            .expect("document.xml should be readable");

        assert!(
            document_xml.contains("<w:spacing w:after=\"240\"/>"),
            "expected markdown paragraph spacing marker in DOCX output"
        );
    }

    #[test]
    fn markdown_headings_emit_docx_spacing_to_prevent_style_crunch() {
        let dir = tempdir().expect("tempdir should be created");
        let output_docx = dir.path().join("out.docx");
        let doc = Document {
            blocks: vec![
                Block::Paragraph(vec![Inline::Text("Lead".into())]),
                Block::Heading {
                    level: 2,
                    content: vec![Inline::Text("Section".into())],
                },
                Block::Paragraph(vec![Inline::Text("Body".into())]),
            ],
        };

        let warnings = write_docx(
            &doc,
            dir.path(),
            &output_docx,
            &DocxWriteOptions {
                allow_remote_images: false,
                style_map: StyleMap::builtin(),
                template: None,
            },
        )
        .expect("DOCX write should succeed");
        assert!(warnings.is_empty());

        let mut archive = ZipArchive::new(
            fs::File::open(&output_docx).expect("written docx should be readable as zip"),
        )
        .expect("written docx should be a valid zip");
        let mut document_xml = String::new();
        archive
            .by_name("word/document.xml")
            .expect("document.xml should exist")
            .read_to_string(&mut document_xml)
            .expect("document.xml should be readable");

        assert!(
            document_xml.contains(
                "<w:pStyle w:val=\"Heading2\"/><w:spacing w:before=\"240\" w:after=\"240\"/>"
            ) || document_xml.contains(
                "<w:pStyle w:val=\"Heading2\"/><w:spacing w:after=\"240\" w:before=\"240\"/>"
            ),
            "expected heading paragraph to include spacing markers above and below"
        );
    }

    fn write_template_zip(path: &Path, styles_xml: &[u8]) -> Result<()> {
        let file = fs::File::create(path)?;
        let mut zip = ZipWriter::new(file);
        zip.start_file("word/styles.xml", SimpleFileOptions::default())?;
        zip.write_all(styles_xml)?;
        zip.finish()?;
        Ok(())
    }

    fn write_template_entries_zip(path: &Path, entries: &BTreeMap<String, Vec<u8>>) -> Result<()> {
        let file = fs::File::create(path)?;
        let mut zip = ZipWriter::new(file);
        for (entry_path, bytes) in entries {
            zip.start_file(entry_path, SimpleFileOptions::default())?;
            zip.write_all(bytes)?;
        }
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

    fn write_minimal_docx_with_document_xml(path: &Path, document_xml: &str) -> Result<()> {
        let file = fs::File::create(path)?;
        let mut zip = ZipWriter::new(file);
        zip.start_file("word/document.xml", SimpleFileOptions::default())?;
        zip.write_all(document_xml.as_bytes())?;
        zip.finish()?;
        Ok(())
    }
}
