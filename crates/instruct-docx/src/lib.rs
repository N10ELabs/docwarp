use std::borrow::Cow;
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

use anyhow::{Context, Result, anyhow};
use instruct_core::{
    Block, ConversionWarning, Document, Inline, StyleMap, WarningCode, model::inline_text,
};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use reqwest::blocking::Client;
use zip::ZipArchive;
use zip::ZipWriter;
use zip::write::SimpleFileOptions;

const OFFICE_REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

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
    next_docpr_id: usize,
}

impl DocxBuildState {
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
        let image_index = self.media_files.len() + 1;
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
}

#[derive(Default)]
struct RunStyle {
    bold: bool,
    italic: bool,
    code: bool,
}

#[derive(Default)]
struct ParseParagraph {
    style: Option<String>,
    inlines: Vec<Inline>,
}

#[derive(Default)]
struct ParseTable {
    rows: Vec<Vec<Vec<Inline>>>,
    current_row: Vec<Vec<Inline>>,
    current_cell: Vec<Inline>,
}

pub fn write_docx(
    document: &Document,
    markdown_base_dir: &Path,
    output_path: &Path,
    options: &DocxWriteOptions,
) -> Result<Vec<ConversionWarning>> {
    let mut warnings = Vec::new();

    if let Some(parent) = output_path.parent() {
        fs::create_dir_all(parent)
            .with_context(|| format!("failed creating output directory: {}", parent.display()))?;
    }

    let mut state = DocxBuildState::default();
    let document_xml = build_document_xml(
        document,
        markdown_base_dir,
        options,
        &mut state,
        &mut warnings,
    )?;

    let styles_xml = load_styles_xml(options.template.as_deref(), &mut warnings)?;
    let content_types_xml = build_content_types_xml(&state.media_files);
    let package_rels_xml = build_package_relationships_xml();
    let document_rels_xml = build_document_relationships_xml(&state.relationships);
    let core_xml = build_core_properties_xml();
    let app_xml = build_app_properties_xml();

    let file = fs::File::create(output_path)
        .with_context(|| format!("failed creating output DOCX: {}", output_path.display()))?;
    let mut zip = ZipWriter::new(file);
    let file_options = SimpleFileOptions::default();

    write_zip_entry(
        &mut zip,
        "[Content_Types].xml",
        &content_types_xml,
        file_options,
    )?;
    write_zip_entry(&mut zip, "_rels/.rels", &package_rels_xml, file_options)?;
    write_zip_entry(&mut zip, "word/document.xml", &document_xml, file_options)?;
    write_zip_entry(
        &mut zip,
        "word/_rels/document.xml.rels",
        &document_rels_xml,
        file_options,
    )?;
    write_zip_entry(&mut zip, "word/styles.xml", &styles_xml, file_options)?;
    write_zip_entry(&mut zip, "docProps/core.xml", &core_xml, file_options)?;
    write_zip_entry(&mut zip, "docProps/app.xml", &app_xml, file_options)?;

    for media in &state.media_files {
        let path = format!("word/{}", media.target);
        write_zip_entry(&mut zip, &path, &media.bytes, file_options)?;
    }

    zip.finish().context("failed finalizing DOCX zip")?;

    Ok(warnings)
}

fn write_zip_entry(
    zip: &mut ZipWriter<fs::File>,
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

fn build_document_xml(
    document: &Document,
    markdown_base_dir: &Path,
    options: &DocxWriteOptions,
    state: &mut DocxBuildState,
    warnings: &mut Vec<ConversionWarning>,
) -> Result<Vec<u8>> {
    let mut body = String::new();

    for block in &document.blocks {
        match block {
            Block::Title(content) => {
                body.push_str(&render_paragraph(
                    content,
                    &options.style_map.docx_style_for("title"),
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
                    _ => "h3",
                };

                body.push_str(&render_paragraph(
                    content,
                    &options.style_map.docx_style_for(token),
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                )?);
            }
            Block::Paragraph(content) => {
                body.push_str(&render_paragraph(
                    content,
                    &options.style_map.docx_style_for("paragraph"),
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                )?);
            }
            Block::BlockQuote(content) => {
                body.push_str(&render_paragraph(
                    content,
                    &options.style_map.docx_style_for("quote"),
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                )?);
            }
            Block::CodeBlock { code, .. } => {
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
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                )?);
            }
            Block::List { ordered, items } => {
                let style = if *ordered {
                    options.style_map.docx_style_for("list_number")
                } else {
                    options.style_map.docx_style_for("list_bullet")
                };

                for item in items {
                    body.push_str(&render_paragraph(
                        item,
                        &style,
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
                    markdown_base_dir,
                    options,
                    state,
                    warnings,
                )?);
            }
        }
    }

    body.push_str(
        "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/></w:sectPr>",
    );

    let xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"{OFFICE_REL_NS}\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 wp14\"><w:body>{body}</w:body></w:document>"
    );

    Ok(xml.into_bytes())
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

    if !headers.is_empty() {
        out.push_str("<w:tr>");
        for cell in headers {
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

    for row in rows {
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
    markdown_base_dir: &Path,
    options: &DocxWriteOptions,
    state: &mut DocxBuildState,
    warnings: &mut Vec<ConversionWarning>,
) -> Result<String> {
    let mut out = String::new();
    out.push_str("<w:p><w:pPr>");
    out.push_str(&format!("<w:pStyle w:val=\"{}\"/>", escape_xml(style)));
    out.push_str("</w:pPr>");
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
                    format!("Remote image blocked (use --allow-remote-images): {src}"),
                )
                .with_location(src),
            );
            return None;
        }

        match Client::new().get(src).send() {
            Ok(response) => match response.error_for_status() {
                Ok(ok_response) => match ok_response.bytes() {
                    Ok(data) => data.to_vec(),
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
                Err(err) => {
                    warnings.push(
                        ConversionWarning::new(
                            WarningCode::ImageLoadFailed,
                            format!("Failed downloading remote image: {err}"),
                        )
                        .with_location(src),
                    );
                    return None;
                }
            },
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
        let candidate = markdown_base_dir.join(src);
        match fs::read(&candidate) {
            Ok(data) => data,
            Err(err) => {
                warnings.push(
                    ConversionWarning::new(
                        WarningCode::ImageLoadFailed,
                        format!("Failed reading local image: {err}"),
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

fn load_styles_xml(
    template_path: Option<&Path>,
    warnings: &mut Vec<ConversionWarning>,
) -> Result<Vec<u8>> {
    if let Some(template_path) = template_path {
        if !template_path.exists() {
            warnings.push(
                ConversionWarning::new(
                    WarningCode::InvalidTemplate,
                    format!("Template not found: {}", template_path.display()),
                )
                .with_location(template_path.display().to_string()),
            );
            return Ok(default_styles_xml().as_bytes().to_vec());
        }

        match fs::File::open(template_path)
            .context("failed opening template")
            .and_then(|file| {
                let mut archive =
                    ZipArchive::new(file).context("failed reading template as zip")?;
                let mut styles = String::new();
                let mut entry = archive
                    .by_name("word/styles.xml")
                    .context("template is missing word/styles.xml")?;
                entry
                    .read_to_string(&mut styles)
                    .context("failed reading template styles")?;
                Ok(styles.into_bytes())
            }) {
            Ok(bytes) => return Ok(bytes),
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

    Ok(default_styles_xml().as_bytes().to_vec())
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
  <w:style w:type=\"paragraph\" w:styleId=\"Quote\"><w:name w:val=\"Quote\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:ind w:left=\"720\"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"Code\"><w:name w:val=\"Code\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:spacing w:line=\"240\"/></w:pPr><w:rPr><w:rFonts w:ascii=\"Consolas\" w:hAnsi=\"Consolas\"/><w:sz w:val=\"20\"/></w:rPr></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"ListBullet\"><w:name w:val=\"List Bullet\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:ind w:left=\"720\"/></w:pPr></w:style>
  <w:style w:type=\"paragraph\" w:styleId=\"ListNumber\"><w:name w:val=\"List Number\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:ind w:left=\"720\"/></w:pPr></w:style>
  <w:style w:type=\"table\" w:styleId=\"Table\"><w:name w:val=\"Table\"/></w:style>
</w:styles>"
}

fn build_content_types_xml(media_files: &[MediaFile]) -> Vec<u8> {
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

    xml.push_str("<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>");
    xml.push_str("<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>");
    xml.push_str("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
    xml.push_str("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
    xml.push_str("</Types>");

    xml.into_bytes()
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
<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:title>instruct output</dc:title><dc:creator>instruct</dc:creator></cp:coreProperties>".as_bytes().to_vec()
}

fn build_app_properties_xml() -> Vec<u8> {
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>
<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><Application>instruct</Application></Properties>".as_bytes().to_vec()
}

pub fn read_docx(
    input_path: &Path,
    options: &DocxReadOptions,
) -> Result<(Document, Vec<ConversionWarning>)> {
    let mut warnings = Vec::new();

    let file = fs::File::open(input_path)
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
    let mut pending_list: Option<(bool, Vec<Vec<Inline>>)> = None;
    let mut in_text_node = false;

    let mut reader = Reader::from_str(&document_xml);
    reader.config_mut().trim_text(false);

    let mut buf = Vec::new();
    loop {
        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(start)) => {
                let name = start.name().as_ref().to_vec();
                match local_name(&name) {
                    b"t" => in_text_node = true,
                    b"p" => paragraph = Some(ParseParagraph::default()),
                    b"pStyle" => {
                        if let Some(value) = attr_value(&start, b"val") {
                            if let Some(paragraph) = paragraph.as_mut() {
                                paragraph.style = Some(value);
                            }
                        }
                    }
                    b"r" => run_style = RunStyle::default(),
                    b"b" => run_style.bold = true,
                    b"i" => run_style.italic = true,
                    b"rStyle" => {
                        if let Some(value) = attr_value(&start, b"val") {
                            if value.contains("Code") {
                                run_style.code = true;
                            }
                        }
                    }
                    b"hyperlink" => {
                        if let Some(rel_id) = attr_value(&start, b"id") {
                            if let Some(url) = relationships.get(&rel_id) {
                                current_hyperlink = Some((url.clone(), Vec::new()));
                            }
                        }
                    }
                    b"br" => push_inline_target(Inline::LineBreak, &mut paragraph, &mut table),
                    b"tbl" => {
                        flush_pending_list(&mut pending_list, &mut blocks);
                        table = Some(ParseTable::default());
                    }
                    b"tr" => {
                        if let Some(table) = table.as_mut() {
                            table.current_row.clear();
                        }
                    }
                    b"tc" => {
                        if let Some(table) = table.as_mut() {
                            table.current_cell.clear();
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
                match local_name(&name) {
                    b"pStyle" => {
                        if let Some(value) = attr_value(&start, b"val") {
                            if let Some(paragraph) = paragraph.as_mut() {
                                paragraph.style = Some(value);
                            }
                        }
                    }
                    b"b" => run_style.bold = true,
                    b"i" => run_style.italic = true,
                    b"rStyle" => {
                        if let Some(value) = attr_value(&start, b"val") {
                            if value.contains("Code") {
                                run_style.code = true;
                            }
                        }
                    }
                    b"br" => push_inline_target(Inline::LineBreak, &mut paragraph, &mut table),
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
            Ok(Event::Text(text)) => {
                if !in_text_node {
                    buf.clear();
                    continue;
                }

                let decoded = decode_text(&reader, text)?;
                if decoded.is_empty() {
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
                match local_name(&name) {
                    b"t" => in_text_node = false,
                    b"hyperlink" => {
                        if let Some((url, text)) = current_hyperlink.take() {
                            push_inline_target(
                                Inline::Link { text, url },
                                &mut paragraph,
                                &mut table,
                            );
                        }
                    }
                    b"p" => {
                        if let Some(paragraph) = paragraph.take() {
                            if let Some(table) = table.as_mut() {
                                if !table.current_cell.is_empty() && !paragraph.inlines.is_empty() {
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
                    b"tc" => {
                        if let Some(table) = table.as_mut() {
                            table
                                .current_row
                                .push(std::mem::take(&mut table.current_cell));
                        }
                    }
                    b"tr" => {
                        if let Some(table) = table.as_mut() {
                            table.rows.push(std::mem::take(&mut table.current_row));
                        }
                    }
                    b"tbl" => {
                        if let Some(table) = table.take() {
                            let mut rows = table.rows;
                            if !rows.is_empty() {
                                let headers = rows.remove(0);
                                blocks.push(Block::Table { headers, rows });
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

    flush_pending_list(&mut pending_list, &mut blocks);

    Ok((Document { blocks }, warnings))
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
    pending_list: &mut Option<(bool, Vec<Vec<Inline>>)>,
    blocks: &mut Vec<Block>,
) {
    let style = paragraph.style.unwrap_or_else(|| "Normal".to_string());
    let token = style_map.md_token_for(&style);

    match token.as_str() {
        "list_bullet" | "list_number" => {
            let ordered = token == "list_number";
            if let Some((existing_ordered, items)) = pending_list.as_mut() {
                if *existing_ordered == ordered {
                    items.push(paragraph.inlines);
                    return;
                }
            }

            flush_pending_list(pending_list, blocks);
            *pending_list = Some((ordered, vec![paragraph.inlines]));
        }
        _ => {
            flush_pending_list(pending_list, blocks);

            if paragraph.inlines.len() == 1 {
                if let Inline::Image { alt, src, title } = &paragraph.inlines[0] {
                    blocks.push(Block::Image {
                        alt: alt.clone(),
                        src: src.clone(),
                        title: title.clone(),
                    });
                    return;
                }
            }

            let block = match token.as_str() {
                "title" => Block::Title(paragraph.inlines),
                "h1" => Block::Heading {
                    level: 1,
                    content: paragraph.inlines,
                },
                "h2" => Block::Heading {
                    level: 2,
                    content: paragraph.inlines,
                },
                "h3" => Block::Heading {
                    level: 3,
                    content: paragraph.inlines,
                },
                "quote" => Block::BlockQuote(paragraph.inlines),
                "code" => Block::CodeBlock {
                    language: None,
                    code: inline_text(&paragraph.inlines),
                },
                _ => Block::Paragraph(paragraph.inlines),
            };

            blocks.push(block);
        }
    }
}

fn flush_pending_list(
    pending_list: &mut Option<(bool, Vec<Vec<Inline>>)>,
    blocks: &mut Vec<Block>,
) {
    if let Some((ordered, items)) = pending_list.take() {
        blocks.push(Block::List { ordered, items });
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

fn attr_value(start: &BytesStart<'_>, local_key: &[u8]) -> Option<String> {
    start
        .attributes()
        .flatten()
        .find(|attr| local_name(attr.key.as_ref()) == local_key)
        .and_then(|attr| String::from_utf8(attr.value.as_ref().to_vec()).ok())
}

fn normalize_docx_target(target: &str) -> String {
    let replaced = target.replace('\\', "/");
    replaced.trim_start_matches("../").to_string()
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
    }
}
