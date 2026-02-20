use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq)]
pub struct Document {
    pub blocks: Vec<Block>,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum Block {
    Title(Vec<Inline>),
    Heading {
        level: u8,
        content: Vec<Inline>,
    },
    Paragraph(Vec<Inline>),
    BlockQuote(Vec<Inline>),
    CodeBlock {
        language: Option<String>,
        code: String,
    },
    List {
        ordered: bool,
        items: Vec<Vec<Inline>>,
    },
    Table {
        headers: Vec<Vec<Inline>>,
        rows: Vec<Vec<Vec<Inline>>>,
    },
    Image {
        alt: String,
        src: String,
        title: Option<String>,
    },
    ThematicBreak,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
#[serde(tag = "kind", rename_all = "snake_case")]
pub enum Inline {
    Text(String),
    Emphasis(Vec<Inline>),
    Strong(Vec<Inline>),
    Code(String),
    Link {
        text: Vec<Inline>,
        url: String,
    },
    Image {
        alt: String,
        src: String,
        title: Option<String>,
    },
    LineBreak,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize, PartialEq, Eq)]
pub struct DocumentStats {
    pub block_count: usize,
    pub heading_count: usize,
    pub paragraph_count: usize,
    pub list_count: usize,
    pub list_item_count: usize,
    pub table_count: usize,
    pub image_count: usize,
    pub code_block_count: usize,
}

impl Document {
    pub fn stats(&self) -> DocumentStats {
        let mut stats = DocumentStats::default();
        stats.block_count = self.blocks.len();

        for block in &self.blocks {
            match block {
                Block::Title(_) | Block::Heading { .. } => stats.heading_count += 1,
                Block::Paragraph(_) | Block::BlockQuote(_) => stats.paragraph_count += 1,
                Block::CodeBlock { .. } => stats.code_block_count += 1,
                Block::List { items, .. } => {
                    stats.list_count += 1;
                    stats.list_item_count += items.len();
                }
                Block::Table { .. } => stats.table_count += 1,
                Block::Image { .. } => stats.image_count += 1,
                Block::ThematicBreak => {}
            }
        }

        stats
    }
}

pub fn inline_text(inlines: &[Inline]) -> String {
    let mut out = String::new();
    for inline in inlines {
        match inline {
            Inline::Text(t) => out.push_str(t),
            Inline::Code(t) => out.push_str(t),
            Inline::LineBreak => out.push('\n'),
            Inline::Emphasis(children) | Inline::Strong(children) => {
                out.push_str(&inline_text(children))
            }
            Inline::Link { text, .. } => out.push_str(&inline_text(text)),
            Inline::Image { alt, .. } => out.push_str(alt),
        }
    }
    out
}
