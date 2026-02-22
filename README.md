# docwarp

`docwarp` is a lightweight open-source CLI for converting documents between GitHub-Flavored Markdown and Microsoft Word-compatible DOCX.

## Status

MVP scaffold implemented for `v0.1.0`:

- `docwarp md2docx` for Markdown -> DOCX
- `docwarp docx2md` for DOCX -> Markdown
- warning-first conversion policy with optional strict mode
- optional JSON report output
- style-map + config support

## Quick start

```bash
docwarp
cargo run -p docwarp-cli -- md2docx input.md --output output.docx
cargo run -p docwarp-cli -- docx2md input.docx --output output.md
```

Guided mode:

- Run `docwarp` with no arguments.
- Drag a file/folder path into the terminal (or browse interactively).
- `docwarp` detects Markdown vs DOCX and runs the matching conversion automatically.

## Commands

```text
docwarp
docwarp md2docx <input.md|input_dir> --output <output.docx|output_dir> [--glob <pattern>] [--template <template.dotx>] [--style-map <map.yml>] [--config <docwarp.yml>] [--report <report.json|report_dir>] [--strict] [--allow-remote-images]
docwarp docx2md <input.docx|input_dir> --output <output.md|output_dir> [--glob <pattern>] [--assets-dir <dir>] [--style-map <map.yml>] [--config <docwarp.yml>] [--report <report.json|report_dir>] [--strict]
```

Batch mode:

```bash
docwarp md2docx ./docs --output ./build/docx
docwarp docx2md ./contracts --output ./build/md --glob "**/*.docx"
```

Run command-specific help for full examples:

```bash
docwarp --help
docwarp md2docx --help
docwarp docx2md --help
```

## Docs

- Install guide: `docs/install.md`
- Configuration and style maps: `docs/configuration.md`
- Strict mode and CI guidance: `docs/strict-mode.md`
- JSON report schema: `docs/report-schema.md`
- Warning code catalog: `docs/warnings.md`
- Homebrew/core submission guide: `docs/homebrew-core.md`
- Release runbook: `docs/release.md`
- Changelog: `CHANGELOG.md`

## Warning Codes

See `docs/warnings.md` for the stable warning-code catalog.

## License

Apache-2.0
