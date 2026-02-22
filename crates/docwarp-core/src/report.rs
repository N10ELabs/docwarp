use std::fs;
use std::path::Path;

use anyhow::{Context, Result};
use serde::{Deserialize, Serialize};

use crate::model::DocumentStats;
use crate::warning::ConversionWarning;

pub const REPORT_SCHEMA_VERSION: &str = "1.0.0";

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum ConversionDirection {
    MdToDocx,
    DocxToMd,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct ConversionReport {
    pub version: String,
    pub direction: ConversionDirection,
    pub input_path: String,
    pub output_path: String,
    pub duration_ms: u128,
    pub stats: DocumentStats,
    pub warnings: Vec<ConversionWarning>,
    pub success: bool,
}

impl ConversionReport {
    pub fn new(
        direction: ConversionDirection,
        input_path: impl Into<String>,
        output_path: impl Into<String>,
        duration_ms: u128,
        stats: DocumentStats,
        warnings: Vec<ConversionWarning>,
        success: bool,
    ) -> Self {
        Self {
            version: REPORT_SCHEMA_VERSION.to_string(),
            direction,
            input_path: input_path.into(),
            output_path: output_path.into(),
            duration_ms,
            stats,
            warnings,
            success,
        }
    }

    pub fn write_to_path(&self, path: &Path) -> Result<()> {
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent).with_context(|| {
                format!("failed creating report directory: {}", parent.display())
            })?;
        }

        let data =
            serde_json::to_string_pretty(self).context("failed serializing conversion report")?;
        fs::write(path, data)
            .with_context(|| format!("failed writing conversion report: {}", path.display()))
    }
}
