# Warning Codes

`instruct` emits stable warning codes in machine-readable reports and CLI summaries.

- `unsupported_feature`: Encountered markdown or DOCX content outside MVP support and continued.
- `image_load_failed`: Could not read a local image or fetch/parse a remote image.
- `remote_image_blocked`: Remote image embedding was blocked by default policy.
- `missing_media`: A DOCX relationship pointed to media that could not be extracted.
- `invalid_style_map`: Style map configuration is invalid.
- `invalid_template`: Provided template path/content could not be used; built-in styles were used.
- `corrupt_docx`: DOCX package or required XML parts are invalid/corrupt.
- `nested_structure_simplified`: Nested or unclosed structures were simplified during parsing.
