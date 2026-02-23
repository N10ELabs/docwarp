# Configuration and Style Maps

`docwarp` can load runtime defaults from `.docwarp.yml` and can override style mapping with YAML or JSON files.

## Config File

When `--config` is omitted, `docwarp` auto-loads `.docwarp.yml` from the current working directory if it exists.

Example `.docwarp.yml`:

```yaml
markdown_flavor: gfm
style_map: ./styles/project-style-map.yml
assets_dir: ./build/assets
default_template: ./styles/brand.dotx
unsupported_policy: warn_continue
```

Supported fields:

- `markdown_flavor`: currently `gfm`
- `style_map`: path to a YAML or JSON style-map file
- `assets_dir`: default extraction directory for `docx2md`
- `default_template`: default `.dotx` template for `md2docx`
- `unsupported_policy`: `warn_continue`, `fail_fast`, or `best_effort_silent`

Path behavior:

- `style_map` and `default_template` paths in config are resolved relative to the config file directory unless already absolute.
- `assets_dir` is resolved relative to the Markdown output path for `docx2md`.

Precedence:

- Built-in defaults
- Config file values
- CLI flags (`--style-map`, `--template`, `--assets-dir`, `--strict`, etc.)

## Agent Mapping Reference

Use these canonical mapping tokens when an agent prepares a style map:

- Markdown heading tokens: `h1`, `h2`, `h3`, `h4`, `h5`, `h6`
- Equation tokens (for `md_to_docx`): `equation_inline`, `equation_block`
- DOCX heading styles: `Heading1`, `Heading2`, `Heading3`, `Heading4`, `Heading5`, `Heading6`

Default heading mapping:

- `h1` -> `Heading1`
- `h2` -> `Heading2`
- `h3` -> `Heading3`
- `h4` -> `Heading4`
- `h5` -> `Heading5`
- `h6` -> `Heading6`

Reverse mapping for `docx2md`:

- `Heading1` -> `h1`
- `Heading2` -> `h2`
- `Heading3` -> `h3`
- `Heading4` -> `h4`
- `Heading5` -> `h5`
- `Heading6` -> `h6`

Agent guidance:

- Preserve these tokens exactly (case-sensitive) when generating `md_to_docx` and `docx_to_md`.
- If a custom style is needed (for example, `BrandHeading4`), map it to the nearest heading token in `docx_to_md` and keep a corresponding `md_to_docx` entry.
- Keep heading mappings symmetric unless you intentionally want lossy round-trips.

## Style-Map Examples

YAML style map:

```yaml
docx_to_md:
  Title: title
  Heading1: h1
  Heading2: h2
  Heading3: h3
  Heading4: h4
  Heading5: h5
  Heading6: h6
  Normal: paragraph
  Quote: quote
  Code: code
  ListBullet: list_bullet
  ListNumber: list_number
  Table: table
md_to_docx:
  title: Title
  h1: Heading1
  h2: Heading2
  h3: Heading3
  h4: Heading4
  h5: Heading5
  h6: Heading6
  paragraph: Normal
  quote: Quote
  code: Code
  equation_inline: EquationInline
  equation_block: Equation
  list_bullet: ListBullet
  list_number: ListNumber
  table: Table
```

JSON style map:

```json
{
  "docx_to_md": {
    "Title": "title",
    "Heading1": "h1",
    "Heading2": "h2",
    "Heading3": "h3",
    "Heading4": "h4",
    "Heading5": "h5",
    "Heading6": "h6",
    "Normal": "paragraph",
    "Quote": "quote",
    "Code": "code",
    "ListBullet": "list_bullet",
    "ListNumber": "list_number",
    "Table": "table"
  },
  "md_to_docx": {
    "title": "Title",
    "h1": "Heading1",
    "h2": "Heading2",
    "h3": "Heading3",
    "h4": "Heading4",
    "h5": "Heading5",
    "h6": "Heading6",
    "paragraph": "Normal",
    "quote": "Quote",
    "code": "Code",
    "equation_inline": "EquationInline",
    "equation_block": "Equation",
    "list_bullet": "ListBullet",
    "list_number": "ListNumber",
    "table": "Table"
  }
}
```
