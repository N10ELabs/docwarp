# instruct

`instruct` is a lightweight open-source CLI for converting documents between GitHub-Flavored Markdown and Microsoft Word-compatible DOCX.

## Status

MVP scaffold implemented for `v0.1.0`:

- `instruct md2docx` for Markdown -> DOCX
- `instruct docx2md` for DOCX -> Markdown
- warning-first conversion policy with optional strict mode
- optional JSON report output
- style-map + config support

## Quick start

```bash
instruct
cargo run -p instruct-cli -- md2docx input.md --output output.docx
cargo run -p instruct-cli -- docx2md input.docx --output output.md
```

Guided mode:

- Run `instruct` with no arguments.
- Drag a file/folder path into the terminal (or browse interactively).
- `instruct` detects Markdown vs DOCX and runs the matching conversion automatically.

## Commands

```text
instruct
instruct md2docx <input.md|input_dir> --output <output.docx|output_dir> [--glob <pattern>] [--template <template.dotx>] [--style-map <map.yml>] [--config <instruct.yml>] [--report <report.json|report_dir>] [--strict] [--allow-remote-images]
instruct docx2md <input.docx|input_dir> --output <output.md|output_dir> [--glob <pattern>] [--assets-dir <dir>] [--style-map <map.yml>] [--config <instruct.yml>] [--report <report.json|report_dir>] [--strict]
```

Batch mode:

```bash
instruct md2docx ./docs --output ./build/docx
instruct docx2md ./contracts --output ./build/md --glob "**/*.docx"
```

Run command-specific help for full examples:

```bash
instruct --help
instruct md2docx --help
instruct docx2md --help
```

## Docs

- Install guide: `docs/install.md`
- Configuration and style maps: `docs/configuration.md`
- Strict mode and CI guidance: `docs/strict-mode.md`
- JSON report schema: `docs/report-schema.md`
- Warning code catalog: `docs/warnings.md`
- Release runbook: `docs/release.md`
- Changelog: `CHANGELOG.md`

## Warning Codes

See `docs/warnings.md` for the stable warning-code catalog.

## License

Apache-2.0
