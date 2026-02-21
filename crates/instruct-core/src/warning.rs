use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "snake_case")]
pub enum WarningCode {
    UnsupportedFeature,
    ImageLoadFailed,
    RemoteImageBlocked,
    MissingMedia,
    InvalidStyleMap,
    InvalidTemplate,
    CorruptDocx,
    NestedStructureSimplified,
}

impl WarningCode {
    pub const ALL: [WarningCode; 8] = [
        WarningCode::UnsupportedFeature,
        WarningCode::ImageLoadFailed,
        WarningCode::RemoteImageBlocked,
        WarningCode::MissingMedia,
        WarningCode::InvalidStyleMap,
        WarningCode::InvalidTemplate,
        WarningCode::CorruptDocx,
        WarningCode::NestedStructureSimplified,
    ];

    pub fn as_str(&self) -> &'static str {
        match self {
            WarningCode::UnsupportedFeature => "unsupported_feature",
            WarningCode::ImageLoadFailed => "image_load_failed",
            WarningCode::RemoteImageBlocked => "remote_image_blocked",
            WarningCode::MissingMedia => "missing_media",
            WarningCode::InvalidStyleMap => "invalid_style_map",
            WarningCode::InvalidTemplate => "invalid_template",
            WarningCode::CorruptDocx => "corrupt_docx",
            WarningCode::NestedStructureSimplified => "nested_structure_simplified",
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct ConversionWarning {
    pub code: WarningCode,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub location: Option<String>,
}

impl ConversionWarning {
    pub fn new(code: WarningCode, message: impl Into<String>) -> Self {
        Self {
            code,
            message: message.into(),
            location: None,
        }
    }

    pub fn with_location(mut self, location: impl Into<String>) -> Self {
        self.location = Some(location.into());
        self
    }
}

#[cfg(test)]
mod tests {
    use super::WarningCode;

    #[test]
    fn warning_code_catalog_is_stable() {
        let actual: Vec<&str> = WarningCode::ALL.iter().map(WarningCode::as_str).collect();
        let expected = vec![
            "unsupported_feature",
            "image_load_failed",
            "remote_image_blocked",
            "missing_media",
            "invalid_style_map",
            "invalid_template",
            "corrupt_docx",
            "nested_structure_simplified",
        ];

        assert_eq!(actual, expected);
    }
}
