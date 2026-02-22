# JSON Report Schema

When `--report <path>` is provided, `docwarp` writes a machine-readable JSON report.

Current schema version: `1.0.0`

## Top-Level Fields

- `version` (`string`): report schema version
- `direction` (`"md_to_docx"` or `"docx_to_md"`): conversion direction
- `input_path` (`string`): input file path
- `output_path` (`string`): output file path
- `duration_ms` (`number`): elapsed conversion time in milliseconds
- `stats` (`object`): document block counters
- `warnings` (`array`): warning entries
- `success` (`boolean`): whether conversion completed successfully

## `stats` Object

- `block_count`
- `heading_count`
- `paragraph_count`
- `list_count`
- `list_item_count`
- `table_count`
- `image_count`
- `code_block_count`

All stats fields are numeric counters.

## `warnings` Items

Each warning item has:

- `code` (`string`): stable warning code (see `docs/warnings.md`)
- `message` (`string`): human-readable description
- `location` (`string`, optional): source location/path if available

## Sample Report

```json
{
  "version": "1.0.0",
  "direction": "md_to_docx",
  "input_path": "fixtures/md/10-comprehensive.md",
  "output_path": "build/10-comprehensive.docx",
  "duration_ms": 42,
  "stats": {
    "block_count": 8,
    "heading_count": 2,
    "paragraph_count": 2,
    "list_count": 2,
    "list_item_count": 4,
    "table_count": 1,
    "image_count": 1,
    "code_block_count": 1
  },
  "warnings": [
    {
      "code": "remote_image_blocked",
      "message": "Remote image blocked by offline-by-default policy. Re-run with --allow-remote-images: https://example.invalid/image.png",
      "location": "https://example.invalid/image.png"
    }
  ],
  "success": true
}
```
