use std::fs;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum MarkdownFlavor {
    Gfm,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum UnsupportedPolicy {
    WarnContinue,
    FailFast,
    BestEffortSilent,
}

#[derive(Debug, Clone, Serialize, Deserialize, Default, PartialEq, Eq)]
pub struct AppConfig {
    #[serde(default)]
    pub markdown_flavor: Option<MarkdownFlavor>,
    #[serde(default)]
    pub style_map: Option<PathBuf>,
    #[serde(default)]
    pub assets_dir: Option<PathBuf>,
    #[serde(default)]
    pub default_template: Option<PathBuf>,
    #[serde(default)]
    pub unsupported_policy: Option<UnsupportedPolicy>,
}

impl AppConfig {
    pub fn load(path: &Path) -> Result<Self> {
        let raw = fs::read_to_string(path)
            .with_context(|| format!("failed reading config file: {}", path.display()))?;

        match path.extension().and_then(|ext| ext.to_str()) {
            Some("json") => serde_json::from_str(&raw)
                .with_context(|| format!("invalid JSON config: {}", path.display())),
            _ => serde_yaml::from_str(&raw)
                .with_context(|| format!("invalid YAML config: {}", path.display())),
        }
    }

    pub fn load_optional(path: Option<&Path>) -> Result<Self> {
        match path {
            Some(path) => Self::load(path),
            None => Ok(Self::default()),
        }
    }

    pub fn markdown_flavor_or_default(&self) -> MarkdownFlavor {
        self.markdown_flavor.clone().unwrap_or(MarkdownFlavor::Gfm)
    }

    pub fn unsupported_policy_or_default(&self) -> UnsupportedPolicy {
        self.unsupported_policy
            .clone()
            .unwrap_or(UnsupportedPolicy::WarnContinue)
    }
}
