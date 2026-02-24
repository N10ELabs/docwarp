use std::collections::BTreeMap;
use std::fs;
use std::path::Path;

use anyhow::{Context, Result, bail};
use serde::{Deserialize, Serialize};

const DOCX_TO_MD_ALLOWED_TOKENS: [&str; 13] = [
    "title",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "paragraph",
    "quote",
    "code",
    "list_bullet",
    "list_number",
    "table",
];

const MD_TO_DOCX_ALLOWED_TOKENS: [&str; 15] = [
    "title",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "paragraph",
    "quote",
    "code",
    "equation_inline",
    "equation_block",
    "list_bullet",
    "list_number",
    "table",
];

#[derive(Debug, Clone, Serialize, Deserialize, Default, PartialEq, Eq)]
#[serde(deny_unknown_fields)]
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
        md_to_docx.insert("equation_inline".to_string(), "EquationInline".to_string());
        md_to_docx.insert("equation_block".to_string(), "Equation".to_string());

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

    parse_style_map(&raw, StyleMapFormat::from_path(path), path)
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StyleMapFormat {
    Json,
    Yaml,
}

impl StyleMapFormat {
    fn from_path(path: &Path) -> Self {
        match path.extension().and_then(|ext| ext.to_str()) {
            Some("json") => Self::Json,
            _ => Self::Yaml,
        }
    }

    fn name(self) -> &'static str {
        match self {
            Self::Json => "JSON",
            Self::Yaml => "YAML",
        }
    }
}

fn parse_style_map(raw: &str, format: StyleMapFormat, path: &Path) -> Result<StyleMap> {
    let style_map: StyleMap = match format {
        StyleMapFormat::Json => serde_json::from_str(raw)
            .with_context(|| format!("invalid JSON style map: {}", path.display()))?,
        StyleMapFormat::Yaml => serde_yaml::from_str(raw)
            .with_context(|| format!("invalid YAML style map: {}", path.display()))?,
    };

    validate_style_map(&style_map, path)
        .with_context(|| format!("invalid {} style map: {}", format.name(), path.display()))?;

    Ok(style_map)
}

fn validate_style_map(style_map: &StyleMap, path: &Path) -> Result<()> {
    for (style_name, token) in &style_map.docx_to_md {
        if style_name.trim().is_empty() {
            bail!(
                "found an empty DOCX style name key in `docx_to_md`; use a real style name like `Heading1`"
            );
        }

        if token.trim().is_empty() {
            let entry = format_map_entry("docx_to_md", style_name);
            bail!(
                "empty markdown token at `{entry}`; provide one of: {}",
                format_allowed_tokens(&DOCX_TO_MD_ALLOWED_TOKENS)
            );
        }

        ensure_allowed_token(
            path,
            "docx_to_md",
            style_name,
            token,
            &DOCX_TO_MD_ALLOWED_TOKENS,
            true,
        )?;
    }

    for (token, style_name) in &style_map.md_to_docx {
        if token.trim().is_empty() {
            bail!(
                "found an empty markdown token key in `md_to_docx`; use one of: {}",
                format_allowed_tokens(&MD_TO_DOCX_ALLOWED_TOKENS)
            );
        }

        ensure_allowed_token(
            path,
            "md_to_docx",
            token,
            token,
            &MD_TO_DOCX_ALLOWED_TOKENS,
            false,
        )?;

        if style_name.trim().is_empty() {
            let entry = format_map_entry("md_to_docx", token);
            bail!("empty DOCX style name at `{entry}`; provide a Word style such as `Normal`");
        }
    }

    Ok(())
}

fn ensure_allowed_token(
    path: &Path,
    section: &str,
    entry_key: &str,
    token: &str,
    allowed: &[&str],
    token_in_value_position: bool,
) -> Result<()> {
    if allowed.contains(&token) {
        return Ok(());
    }

    let suggestion = suggest_token(token, allowed)
        .map(|candidate| format!(" Did you mean `{candidate}`?"))
        .unwrap_or_default();
    let allowed_list = format_allowed_tokens(allowed);
    let entry = format_map_entry(section, entry_key);

    if token_in_value_position {
        bail!(
            "invalid markdown token `{token}` at `{entry}` in {}; values in `{section}` must be one of: {allowed_list}.{suggestion}",
            path.display()
        );
    }

    bail!(
        "invalid markdown token key `{token}` at `{entry}` in {}; keys in `{section}` must be one of: {allowed_list}.{suggestion}",
        path.display()
    );
}

fn format_allowed_tokens(tokens: &[&str]) -> String {
    tokens
        .iter()
        .map(|token| format!("`{token}`"))
        .collect::<Vec<_>>()
        .join(", ")
}

fn format_map_entry(section: &str, key: &str) -> String {
    if key
        .chars()
        .all(|ch| ch.is_ascii_alphanumeric() || ch == '_' || ch == '-')
    {
        format!("{section}.{key}")
    } else {
        let escaped = key.replace('"', "\\\"");
        format!("{section}[\"{escaped}\"]")
    }
}

fn suggest_token<'a>(actual: &str, allowed: &'a [&str]) -> Option<&'a str> {
    let normalized_actual = actual.to_ascii_lowercase();
    let mut best: Option<(&str, usize)> = None;

    for candidate in allowed {
        let distance = levenshtein(&normalized_actual, &candidate.to_ascii_lowercase());
        match best {
            None => best = Some((candidate, distance)),
            Some((_, best_distance)) if distance < best_distance => {
                best = Some((candidate, distance));
            }
            _ => {}
        }
    }

    best.and_then(|(candidate, distance)| (distance <= 3).then_some(candidate))
}

fn levenshtein(a: &str, b: &str) -> usize {
    if a == b {
        return 0;
    }
    if a.is_empty() {
        return b.chars().count();
    }
    if b.is_empty() {
        return a.chars().count();
    }

    let a_chars: Vec<char> = a.chars().collect();
    let b_chars: Vec<char> = b.chars().collect();
    let mut prev: Vec<usize> = (0..=b_chars.len()).collect();
    let mut curr = vec![0; b_chars.len() + 1];

    for (i, a_char) in a_chars.iter().enumerate() {
        curr[0] = i + 1;
        for (j, b_char) in b_chars.iter().enumerate() {
            let replace_cost = usize::from(a_char != b_char);
            let insert = curr[j] + 1;
            let delete = prev[j + 1] + 1;
            let replace = prev[j] + replace_cost;
            curr[j + 1] = insert.min(delete).min(replace);
        }
        std::mem::swap(&mut prev, &mut curr);
    }

    prev[b_chars.len()]
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

    #[test]
    fn builtin_style_map_includes_equation_tokens() {
        let merged = resolve_style_map(None, None);

        assert_eq!(
            merged.md_to_docx.get("equation_inline").map(String::as_str),
            Some("EquationInline")
        );
        assert_eq!(
            merged.md_to_docx.get("equation_block").map(String::as_str),
            Some("Equation")
        );
    }

    fn parse_yaml(raw: &str) -> Result<StyleMap> {
        parse_style_map(raw, StyleMapFormat::Yaml, Path::new("style-map.yml"))
    }

    #[test]
    fn style_map_validation_reports_unknown_md_to_docx_token_keys() {
        let err = parse_yaml("md_to_docx:\n  paragrph: Normal\n")
            .expect_err("invalid markdown token key should fail");
        let message = format!("{err:#}");

        assert!(
            message.contains("md_to_docx.paragrph"),
            "expected map entry path in error, got:\n{message}"
        );
        assert!(
            message.contains("Did you mean `paragraph`?"),
            "expected token suggestion in error, got:\n{message}"
        );
        assert!(
            message.contains("keys in `md_to_docx` must be one of"),
            "expected allowed-token guidance in error, got:\n{message}"
        );
    }

    #[test]
    fn style_map_validation_reports_unknown_docx_to_md_token_values() {
        let err = parse_yaml("docx_to_md:\n  Heading1: h7\n")
            .expect_err("invalid markdown token value should fail");
        let message = format!("{err:#}");

        assert!(
            message.contains("docx_to_md.Heading1"),
            "expected map entry path in error, got:\n{message}"
        );
        assert!(
            message.contains("values in `docx_to_md` must be one of"),
            "expected allowed-token guidance in error, got:\n{message}"
        );
    }

    #[test]
    fn style_map_validation_rejects_empty_docx_style_name_keys() {
        let err = parse_yaml("docx_to_md:\n  \"\": h1\n")
            .expect_err("empty docx style names should fail");
        let message = format!("{err:#}");

        assert!(
            message.contains("empty DOCX style name key"),
            "expected clear key diagnostic, got:\n{message}"
        );
    }
}
