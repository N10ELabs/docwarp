use std::mem;

use anyhow::Result;
use instruct_core::{Block, ConversionWarning, Document, Inline, WarningCode, model::inline_text};
use pulldown_cmark::{CodeBlockKind, Event, HeadingLevel, Options, Parser, Tag, TagEnd};

#[derive(Debug)]
enum BlockContext {
    Paragraph(Vec<Inline>),
    Heading(u8, Vec<Inline>),
    BlockQuote(Vec<Inline>),
    Item(Vec<Inline>),
    CodeBlock {
        language: Option<String>,
        code: String,
    },
    Table(TableContext),
}

#[derive(Debug)]
struct TableContext {
    in_head: bool,
    headers: Vec<Vec<Inline>>,
    rows: Vec<Vec<Vec<Inline>>>,
    current_row: Vec<Vec<Inline>>,
    current_cell: Vec<Inline>,
}

#[derive(Debug)]
enum InlineContext {
    Emphasis(Vec<Inline>),
    Strong(Vec<Inline>),
    Link {
        url: String,
        text: Vec<Inline>,
    },
    Image {
        src: String,
        title: Option<String>,
        alt: Vec<Inline>,
    },
}

#[derive(Debug)]
struct ListContext {
    ordered: bool,
    items: Vec<Vec<Inline>>,
    trailing_lists: Vec<Block>,
}

pub fn parse_markdown(input: &str) -> Result<(Document, Vec<ConversionWarning>)> {
    let mut warnings = Vec::new();
    let mut blocks = Vec::new();
    let mut block_stack: Vec<BlockContext> = Vec::new();
    let mut inline_stack: Vec<InlineContext> = Vec::new();
    let mut list_stack: Vec<ListContext> = Vec::new();

    let options = Options::ENABLE_TABLES
        | Options::ENABLE_STRIKETHROUGH
        | Options::ENABLE_TASKLISTS
        | Options::ENABLE_HEADING_ATTRIBUTES
        | Options::ENABLE_FOOTNOTES;

    let parser = Parser::new_ext(input, options);

    for event in parser {
        match event {
            Event::Start(tag) => match tag {
                Tag::Paragraph => {
                    if !matches!(block_stack.last(), Some(BlockContext::BlockQuote(_)))
                        && !matches!(block_stack.last(), Some(BlockContext::Item(_)))
                    {
                        block_stack.push(BlockContext::Paragraph(Vec::new()));
                    }
                }
                Tag::Heading { level, .. } => {
                    block_stack.push(BlockContext::Heading(
                        heading_level_to_u8(level),
                        Vec::new(),
                    ));
                }
                Tag::BlockQuote(_) => {
                    block_stack.push(BlockContext::BlockQuote(Vec::new()));
                }
                Tag::List(start) => {
                    if !list_stack.is_empty() {
                        warnings.push(ConversionWarning::new(
                            WarningCode::NestedStructureSimplified,
                            "Nested lists are flattened in Markdown parsing",
                        ));
                    }
                    list_stack.push(ListContext {
                        ordered: start.is_some(),
                        items: Vec::new(),
                        trailing_lists: Vec::new(),
                    });
                }
                Tag::Item => {
                    block_stack.push(BlockContext::Item(Vec::new()));
                }
                Tag::CodeBlock(kind) => {
                    let language = match kind {
                        CodeBlockKind::Indented => None,
                        CodeBlockKind::Fenced(lang) => {
                            let trimmed = lang.trim();
                            if trimmed.is_empty() {
                                None
                            } else {
                                Some(trimmed.to_string())
                            }
                        }
                    };
                    block_stack.push(BlockContext::CodeBlock {
                        language,
                        code: String::new(),
                    });
                }
                Tag::Table(_) => {
                    block_stack.push(BlockContext::Table(TableContext {
                        in_head: false,
                        headers: Vec::new(),
                        rows: Vec::new(),
                        current_row: Vec::new(),
                        current_cell: Vec::new(),
                    }));
                }
                Tag::TableHead => {
                    if let Some(BlockContext::Table(table)) = block_stack.last_mut() {
                        table.in_head = true;
                    }
                }
                Tag::TableRow => {
                    if let Some(BlockContext::Table(table)) = block_stack.last_mut() {
                        table.current_row.clear();
                    }
                }
                Tag::TableCell => {
                    if let Some(BlockContext::Table(table)) = block_stack.last_mut() {
                        table.current_cell.clear();
                    }
                }
                Tag::Emphasis => inline_stack.push(InlineContext::Emphasis(Vec::new())),
                Tag::Strong => inline_stack.push(InlineContext::Strong(Vec::new())),
                Tag::Link { dest_url, .. } => inline_stack.push(InlineContext::Link {
                    url: dest_url.to_string(),
                    text: Vec::new(),
                }),
                Tag::Image {
                    dest_url, title, ..
                } => inline_stack.push(InlineContext::Image {
                    src: dest_url.to_string(),
                    title: if title.trim().is_empty() {
                        None
                    } else {
                        Some(title.to_string())
                    },
                    alt: Vec::new(),
                }),
                _ => {
                    warnings.push(ConversionWarning::new(
                        WarningCode::UnsupportedFeature,
                        "Encountered unsupported markdown tag during parsing",
                    ));
                }
            },
            Event::End(tag) => match tag {
                TagEnd::Paragraph => {
                    if matches!(block_stack.last(), Some(BlockContext::Paragraph(_))) {
                        if let Some(BlockContext::Paragraph(content)) =
                            pop_context(&mut block_stack)
                        {
                            if !content.is_empty() {
                                blocks.push(Block::Paragraph(content));
                            }
                        }
                    }
                }
                TagEnd::Heading(_) => {
                    if let Some(BlockContext::Heading(level, content)) =
                        pop_context(&mut block_stack)
                    {
                        let level = level.clamp(1, 6);
                        if level == 1 && blocks.is_empty() {
                            blocks.push(Block::Title(content));
                        } else {
                            blocks.push(Block::Heading { level, content });
                        }
                    }
                }
                TagEnd::BlockQuote(_) => {
                    if let Some(BlockContext::BlockQuote(content)) = pop_context(&mut block_stack) {
                        blocks.push(Block::BlockQuote(content));
                    }
                }
                TagEnd::Item => {
                    if let Some(BlockContext::Item(content)) = pop_context(&mut block_stack) {
                        if let Some(list) = list_stack.last_mut() {
                            list.items.push(content);
                        }
                    }
                }
                TagEnd::List(_) => {
                    if let Some(list) = list_stack.pop() {
                        let list_block = Block::List {
                            ordered: list.ordered,
                            items: list.items,
                        };

                        if let Some(parent) = list_stack.last_mut() {
                            parent.trailing_lists.push(list_block);
                            parent.trailing_lists.extend(list.trailing_lists);
                        } else {
                            blocks.push(list_block);
                            blocks.extend(list.trailing_lists);
                        }
                    }
                }
                TagEnd::CodeBlock => {
                    if let Some(BlockContext::CodeBlock { language, code }) =
                        pop_context(&mut block_stack)
                    {
                        blocks.push(Block::CodeBlock { language, code });
                    }
                }
                TagEnd::TableHead => {
                    if let Some(BlockContext::Table(table)) = block_stack.last_mut() {
                        if table.headers.is_empty() && !table.current_row.is_empty() {
                            table.headers = mem::take(&mut table.current_row);
                        }
                        table.in_head = false;
                    }
                }
                TagEnd::TableCell => {
                    if let Some(BlockContext::Table(table)) = block_stack.last_mut() {
                        let cell = mem::take(&mut table.current_cell);
                        table.current_row.push(cell);
                    }
                }
                TagEnd::TableRow => {
                    if let Some(BlockContext::Table(table)) = block_stack.last_mut() {
                        if table.in_head && table.headers.is_empty() {
                            table.headers = mem::take(&mut table.current_row);
                        } else {
                            table.rows.push(mem::take(&mut table.current_row));
                        }
                    }
                }
                TagEnd::Table => {
                    if let Some(BlockContext::Table(table)) = pop_context(&mut block_stack) {
                        let mut headers = table.headers;
                        let mut rows = table.rows;
                        if headers.is_empty() && !rows.is_empty() {
                            headers = rows.remove(0);
                        }
                        normalize_table_dimensions(&mut headers, &mut rows);
                        blocks.push(Block::Table { headers, rows });
                    }
                }
                TagEnd::Emphasis => close_inline_context(&mut inline_stack, &mut block_stack),
                TagEnd::Strong => close_inline_context(&mut inline_stack, &mut block_stack),
                TagEnd::Link => close_inline_context(&mut inline_stack, &mut block_stack),
                TagEnd::Image => close_inline_context(&mut inline_stack, &mut block_stack),
                _ => {}
            },
            Event::Text(text) => {
                if let Some(BlockContext::CodeBlock { code, .. }) = block_stack.last_mut() {
                    code.push_str(&text);
                } else {
                    push_inline(
                        Inline::Text(text.to_string()),
                        &mut inline_stack,
                        &mut block_stack,
                    );
                }
            }
            Event::Code(text) => {
                push_inline(
                    Inline::Code(text.to_string()),
                    &mut inline_stack,
                    &mut block_stack,
                );
            }
            Event::SoftBreak | Event::HardBreak => {
                push_inline(Inline::LineBreak, &mut inline_stack, &mut block_stack);
            }
            Event::Rule => blocks.push(Block::ThematicBreak),
            Event::TaskListMarker(checked) => {
                let marker = if checked { "[x] " } else { "[ ] " };
                push_inline(
                    Inline::Text(marker.to_string()),
                    &mut inline_stack,
                    &mut block_stack,
                );
            }
            Event::Html(raw) => {
                warnings.push(ConversionWarning::new(
                    WarningCode::UnsupportedFeature,
                    "Raw HTML is preserved as text",
                ));
                push_inline(
                    Inline::Text(raw.to_string()),
                    &mut inline_stack,
                    &mut block_stack,
                );
            }
            _ => {
                warnings.push(ConversionWarning::new(
                    WarningCode::UnsupportedFeature,
                    "Encountered unsupported markdown event",
                ));
            }
        }
    }

    if !block_stack.is_empty() {
        warnings.push(ConversionWarning::new(
            WarningCode::NestedStructureSimplified,
            "Unclosed markdown structures were simplified",
        ));
    }

    let mut normalized = Vec::with_capacity(blocks.len());
    for block in blocks {
        match block {
            Block::Paragraph(inlines) => {
                if inlines.len() == 1 {
                    if let Inline::Image { alt, src, title } = &inlines[0] {
                        normalized.push(Block::Image {
                            alt: alt.clone(),
                            src: src.clone(),
                            title: title.clone(),
                        });
                        continue;
                    }
                }
                normalized.push(Block::Paragraph(inlines));
            }
            _ => normalized.push(block),
        }
    }

    Ok((Document { blocks: normalized }, warnings))
}

pub fn render_markdown(document: &Document) -> String {
    let mut out = Vec::new();

    for block in &document.blocks {
        let rendered = match block {
            Block::Title(content) => format!("# {}", render_inlines(content)),
            Block::Heading { level, content } => {
                let level = (*level).clamp(1, 6);
                format!("{} {}", "#".repeat(level as usize), render_inlines(content))
            }
            Block::Paragraph(content) => render_inlines(content),
            Block::BlockQuote(content) => {
                let text = render_inlines(content);
                text.lines()
                    .map(|line| format!("> {line}"))
                    .collect::<Vec<_>>()
                    .join("\n")
            }
            Block::CodeBlock { language, code } => {
                let lang = language.clone().unwrap_or_default();
                format!("```{lang}\n{code}\n```")
            }
            Block::List { ordered, items } => items
                .iter()
                .enumerate()
                .map(|(idx, item)| {
                    if *ordered {
                        format!("{}. {}", idx + 1, render_inlines(item))
                    } else {
                        format!("- {}", render_inlines(item))
                    }
                })
                .collect::<Vec<_>>()
                .join("\n"),
            Block::Table { headers, rows } => render_table(headers, rows),
            Block::Image { alt, src, title } => {
                if let Some(title) = title {
                    format!("![{alt}]({src} \"{title}\")")
                } else {
                    format!("![{alt}]({src})")
                }
            }
            Block::ThematicBreak => "---".to_string(),
        };
        out.push(rendered);
    }

    out.join("\n\n")
}

fn render_table(headers: &[Vec<Inline>], rows: &[Vec<Vec<Inline>>]) -> String {
    let width = headers
        .len()
        .max(rows.iter().map(Vec::len).max().unwrap_or_default());
    if width == 0 {
        return String::new();
    }

    let mut normalized_headers = headers.to_vec();
    normalized_headers.resize_with(width, Vec::new);

    let mut out = String::new();
    out.push('|');
    for header in &normalized_headers {
        out.push(' ');
        out.push_str(&render_inlines(header));
        out.push(' ');
        out.push('|');
    }
    out.push('\n');

    out.push('|');
    for _ in 0..width {
        out.push_str(" --- |");
    }
    out.push('\n');

    for row in rows {
        let mut normalized_row = row.clone();
        normalized_row.resize_with(width, Vec::new);
        out.push('|');
        for cell in &normalized_row {
            out.push(' ');
            out.push_str(&render_inlines(cell));
            out.push(' ');
            out.push('|');
        }
        out.push('\n');
    }

    out.trim_end().to_string()
}

fn render_inlines(inlines: &[Inline]) -> String {
    let mut out = String::new();
    for inline in inlines {
        match inline {
            Inline::Text(text) => out.push_str(text),
            Inline::Emphasis(children) => out.push_str(&render_emphasis(children)),
            Inline::Strong(children) => out.push_str(&render_strong(children)),
            Inline::Code(code) => out.push_str(&render_code_span(code)),
            Inline::Link { text, url } => {
                out.push_str(&format!(
                    "[{}]({})",
                    render_inlines(text),
                    render_link_destination(url)
                ));
            }
            Inline::Image { alt, src, title } => {
                if let Some(title) = title {
                    out.push_str(&format!("![{alt}]({src} \"{title}\")"));
                } else {
                    out.push_str(&format!("![{alt}]({src})"));
                }
            }
            Inline::LineBreak => out.push_str("\\\n"),
        }
    }
    out
}

fn render_emphasis(children: &[Inline]) -> String {
    let inner = render_inlines(children);
    let delimiter = if inner.contains('*') && !inner.contains('_') {
        "_"
    } else {
        "*"
    };
    format!("{delimiter}{inner}{delimiter}")
}

fn render_strong(children: &[Inline]) -> String {
    let inner = render_inlines(children);
    let delimiter = if inner.contains("**") && !inner.contains("__") {
        "__"
    } else {
        "**"
    };
    format!("{delimiter}{inner}{delimiter}")
}

fn render_code_span(code: &str) -> String {
    let mut max_backtick_run = 0;
    let mut current_run = 0;
    for ch in code.chars() {
        if ch == '`' {
            current_run += 1;
            max_backtick_run = max_backtick_run.max(current_run);
        } else {
            current_run = 0;
        }
    }

    let fence = "`".repeat(max_backtick_run + 1);
    if code.starts_with('`') || code.ends_with('`') || code.starts_with(' ') || code.ends_with(' ')
    {
        format!("{fence} {code} {fence}")
    } else {
        format!("{fence}{code}{fence}")
    }
}

fn render_link_destination(url: &str) -> String {
    if url.contains([' ', '\t', '\n', '(', ')']) {
        format!("<{}>", url.replace('>', "%3E"))
    } else {
        url.to_string()
    }
}

fn heading_level_to_u8(level: HeadingLevel) -> u8 {
    match level {
        HeadingLevel::H1 => 1,
        HeadingLevel::H2 => 2,
        HeadingLevel::H3 => 3,
        HeadingLevel::H4 => 4,
        HeadingLevel::H5 => 5,
        HeadingLevel::H6 => 6,
    }
}

fn pop_context(block_stack: &mut Vec<BlockContext>) -> Option<BlockContext> {
    block_stack.pop()
}

fn close_inline_context(
    inline_stack: &mut Vec<InlineContext>,
    block_stack: &mut Vec<BlockContext>,
) {
    let Some(context) = inline_stack.pop() else {
        return;
    };

    let inline = match context {
        InlineContext::Emphasis(children) => Inline::Emphasis(children),
        InlineContext::Strong(children) => Inline::Strong(children),
        InlineContext::Link { url, text } => Inline::Link { text, url },
        InlineContext::Image { src, title, alt } => Inline::Image {
            alt: inline_text(&alt),
            src,
            title,
        },
    };

    push_inline(inline, inline_stack, block_stack);
}

fn push_inline(
    inline: Inline,
    inline_stack: &mut [InlineContext],
    block_stack: &mut Vec<BlockContext>,
) {
    if let Some(context) = inline_stack.last_mut() {
        match context {
            InlineContext::Emphasis(children)
            | InlineContext::Strong(children)
            | InlineContext::Link { text: children, .. }
            | InlineContext::Image { alt: children, .. } => {
                children.push(inline);
            }
        }
        return;
    }

    if let Some(block) = block_stack.last_mut() {
        match block {
            BlockContext::Paragraph(content)
            | BlockContext::Heading(_, content)
            | BlockContext::BlockQuote(content)
            | BlockContext::Item(content) => content.push(inline),
            BlockContext::Table(table) => table.current_cell.push(inline),
            BlockContext::CodeBlock { code, .. } => match inline {
                Inline::Text(text) | Inline::Code(text) => code.push_str(&text),
                Inline::LineBreak => code.push('\n'),
                other => {
                    code.push_str(&inline_text(&[other]));
                }
            },
        }
    }
}

fn normalize_table_dimensions(headers: &mut Vec<Vec<Inline>>, rows: &mut [Vec<Vec<Inline>>]) {
    let width = headers
        .len()
        .max(rows.iter().map(Vec::len).max().unwrap_or_default());

    if width == 0 {
        return;
    }

    headers.resize_with(width, Vec::new);
    for row in rows {
        row.resize_with(width, Vec::new);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_core_markdown_features() {
        let input = r#"# Title

## Heading

Paragraph with **bold** and [link](https://example.com).

- one
- two

| A | B |
|---|---|
| 1 | 2 |

![alt](img.png)
"#;

        let (doc, warnings) = parse_markdown(input).expect("parse should succeed");

        assert!(warnings.is_empty());
        assert!(matches!(doc.blocks.first(), Some(Block::Title(_))));
        assert!(
            doc.blocks
                .iter()
                .any(|block| matches!(block, Block::Heading { level: 2, .. }))
        );
        assert!(
            doc.blocks
                .iter()
                .any(|block| matches!(block, Block::List { ordered: false, .. }))
        );
        assert!(
            doc.blocks
                .iter()
                .any(|block| matches!(block, Block::Table { .. }))
        );
        assert!(
            doc.blocks
                .iter()
                .any(|block| matches!(block, Block::Image { .. }))
        );
    }

    #[test]
    fn renders_markdown_with_lists_and_code() {
        let document = Document {
            blocks: vec![
                Block::Heading {
                    level: 2,
                    content: vec![Inline::Text("Overview".into())],
                },
                Block::List {
                    ordered: true,
                    items: vec![
                        vec![Inline::Text("one".into())],
                        vec![Inline::Text("two".into())],
                    ],
                },
                Block::CodeBlock {
                    language: Some("rust".into()),
                    code: "fn main() {}".into(),
                },
            ],
        };

        let output = render_markdown(&document);

        assert!(output.contains("## Overview"));
        assert!(output.contains("1. one"));
        assert!(output.contains("```rust"));
    }

    #[test]
    fn normalizes_table_columns_for_uneven_rows() {
        let input = r#"
| A | B | C |
|---|---|---|
| 1 |
| 2 | 3 |
"#;

        let (doc, warnings) = parse_markdown(input).expect("parse should succeed");
        assert!(warnings.is_empty());

        let Some(Block::Table { headers, rows }) = doc.blocks.first() else {
            panic!("expected first block to be a table");
        };

        assert_eq!(headers.len(), 3);
        assert_eq!(rows.len(), 2);
        assert_eq!(rows[0].len(), 3);
        assert_eq!(rows[1].len(), 3);
        assert!(rows[0][1].is_empty());
        assert!(rows[0][2].is_empty());
        assert!(rows[1][2].is_empty());
    }

    #[test]
    fn preserves_list_type_transitions_in_order() {
        let input = r#"1. one
2. two

- three
- four

1. five"#;

        let (doc, _) = parse_markdown(input).expect("parse should succeed");
        let list_kinds: Vec<bool> = doc
            .blocks
            .iter()
            .filter_map(|block| {
                if let Block::List { ordered, .. } = block {
                    Some(*ordered)
                } else {
                    None
                }
            })
            .collect();

        assert_eq!(list_kinds, vec![true, false, true]);
    }

    #[test]
    fn renders_code_spans_with_backticks_safely() {
        let document = Document {
            blocks: vec![Block::Paragraph(vec![Inline::Code("a`b".into())])],
        };

        let output = render_markdown(&document);
        assert_eq!(output.trim(), "``a`b``");
    }

    #[test]
    fn line_breaks_render_as_hard_breaks() {
        let document = Document {
            blocks: vec![Block::Paragraph(vec![
                Inline::Text("line 1".into()),
                Inline::LineBreak,
                Inline::Text("line 2".into()),
            ])],
        };

        let output = render_markdown(&document);
        assert_eq!(output.trim(), "line 1\\\nline 2");
    }
}
