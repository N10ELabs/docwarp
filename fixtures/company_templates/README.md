# Company Template Sample Corpus

This fixture pack stress-tests branded template/style-map behavior for `docwarp`.

## Contents

- `md/`: sample Markdown inputs covering headings, body text, lists, quotes, code, tables, links, equations, images, and mixed structures.
- `style-maps/`: style-map variants for name-based, alias-based, and styleId-based mappings.
- `template_parts/`: OpenXML parts used to assemble an ACME-branded `.dotx` at test/runtime.
- `assets/`: local image assets referenced by sample Markdown files.

## Intended Coverage

- `md_to_docx` style resolution by style display name, alias, and styleId.
- `docx_to_md` style-token resolution when mappings reference display names/aliases.
- linked character style application for inline code runs.
- template list style `numId`/`ilvl` inheritance for nested markdown list levels.
- batch conversion and config-driven defaults with branded templates.
