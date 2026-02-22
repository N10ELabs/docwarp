use std::collections::BTreeMap;
use std::fs;
use std::path::Path;

use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize, Default, PartialEq, Eq)]
pub struct StyleMap {
    #[serde(default)]
    pub docx_to_md: BTreeMap<String, String>,
    #[serde(default)]
    pub md_to_docx: BTreeMap<String, String>,
}

impl StyleMap {
    pub fn builtin() -> Self {
        let mut docx_to_md = BTreeMap::new();
        docx_to_md.insert("Title".to_string(), "title".to_string());
        docx_to_md.insert("Heading1".to_string(), "h1".to_string());
        docx_to_md.insert("Heading2".to_string(), "h2".to_string());
        docx_to_md.insert("Heading3".to_string(), "h3".to_string());
        docx_to_md.insert("Heading4".to_string(), "h4".to_string());
        docx_to_md.insert("Heading5".to_string(), "h5".to_string());
        docx_to_md.insert("Heading6".to_string(), "h6".to_string());
        docx_to_md.insert("Normal".to_string(), "paragraph".to_string());
        docx_to_md.insert("Quote".to_string(), "quote".to_string());
        docx_to_md.insert("Code".to_string(), "code".to_string());
        docx_to_md.insert("ListBullet".to_string(), "list_bullet".to_string());
        docx_to_md.insert("ListNumber".to_string(), "list_number".to_string());
        docx_to_md.insert("Table".to_string(), "table".to_string());

        let mut md_to_docx = BTreeMap::new();
        md_to_docx.insert("title".to_string(), "Title".to_string());
        md_to_docx.insert("h1".to_string(), "Heading1".to_string());
        md_to_docx.insert("h2".to_string(), "Heading2".to_string());
        md_to_docx.insert("h3".to_string(), "Heading3".to_string());
        md_to_docx.insert("h4".to_string(), "Heading4".to_string());
        md_to_docx.insert("h5".to_string(), "Heading5".to_string());
        md_to_docx.insert("h6".to_string(), "Heading6".to_string());
        md_to_docx.insert("paragraph".to_string(), "Normal".to_string());
        md_to_docx.insert("quote".to_string(), "Quote".to_string());
        md_to_docx.insert("code".to_string(), "Code".to_string());
        md_to_docx.insert("list_bullet".to_string(), "ListBullet".to_string());
        md_to_docx.insert("list_number".to_string(), "ListNumber".to_string());
        md_to_docx.insert("table".to_string(), "Table".to_string());

        Self {
            docx_to_md,
            md_to_docx,
        }
    }

    pub fn merge(&mut self, other: StyleMap) {
        for (k, v) in other.docx_to_md {
            self.docx_to_md.insert(k, v);
        }
        for (k, v) in other.md_to_docx {
            self.md_to_docx.insert(k, v);
        }
    }

    pub fn docx_style_for(&self, token: &str) -> String {
        self.md_to_docx
            .get(token)
            .cloned()
            .unwrap_or_else(|| "Normal".to_string())
    }

    pub fn md_token_for(&self, style: &str) -> String {
        self.docx_to_md
            .get(style)
            .cloned()
            .unwrap_or_else(|| "paragraph".to_string())
    }
}

pub fn load_style_map(path: &Path) -> Result<StyleMap> {
    let raw = fs::read_to_string(path)
        .with_context(|| format!("failed reading style map: {}", path.display()))?;

    match path.extension().and_then(|ext| ext.to_str()) {
        Some("json") => serde_json::from_str(&raw)
            .with_context(|| format!("invalid JSON style map: {}", path.display())),
        _ => serde_yaml::from_str(&raw)
            .with_context(|| format!("invalid YAML style map: {}", path.display())),
    }
}

pub fn resolve_style_map(config_map: Option<StyleMap>, cli_map: Option<StyleMap>) -> StyleMap {
    let mut merged = StyleMap::builtin();

    if let Some(config_map) = config_map {
        merged.merge(config_map);
    }

    if let Some(cli_map) = cli_map {
        merged.merge(cli_map);
    }

    merged
}

#[cfg(test)]
mod tests {
    use super::*;

    fn single_map(value: &str) -> StyleMap {
        let mut map = StyleMap::default();
        map.md_to_docx
            .insert("paragraph".to_string(), value.to_string());
        map
    }

    #[test]
    fn style_map_precedence_cli_over_config_over_builtin() {
        let merged = resolve_style_map(
            Some(single_map("ConfigParagraph")),
            Some(single_map("CliParagraph")),
        );

        assert_eq!(
            merged.md_to_docx.get("paragraph").map(String::as_str),
            Some("CliParagraph")
        );
    }

    #[test]
    fn style_map_uses_builtin_when_no_overrides() {
        let merged = resolve_style_map(None, None);

        assert_eq!(
            merged.md_to_docx.get("paragraph").map(String::as_str),
            Some("Normal")
        );
    }

    #[test]
    fn config_override_applies_when_no_cli_map() {
        let merged = resolve_style_map(Some(single_map("ConfigParagraph")), None);

        assert_eq!(
            merged.md_to_docx.get("paragraph").map(String::as_str),
            Some("ConfigParagraph")
        );
    }

    #[test]
    fn builtin_style_map_includes_h1_through_h6() {
        let merged = resolve_style_map(None, None);

        for (token, style) in [
            ("h1", "Heading1"),
            ("h2", "Heading2"),
            ("h3", "Heading3"),
            ("h4", "Heading4"),
            ("h5", "Heading5"),
            ("h6", "Heading6"),
        ] {
            assert_eq!(
                merged.md_to_docx.get(token).map(String::as_str),
                Some(style)
            );
        }
    }
}
