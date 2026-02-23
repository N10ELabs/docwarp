# docwarp

> Agentic CLI Usage References:
> 1. [`AGENTS.md`](AGENTS.md) as the canonical hub for agents using the `docwarp` CLI.
> 2. `README.md` (this file) for commands and conversion flow.
> 3. [`docs/strict-mode.md`](docs/strict-mode.md) for CI/quality gates.

`docwarp` is a lightweight open-source CLI for converting documents between GitHub-Flavored Markdown and Microsoft Word-compatible DOCX.

## Install (Homebrew)

```bash
brew install n10elabs/tap/docwarp
```

## Current Status

`v0.1.0` is available and supports:

- `docwarp md2docx` for Markdown -> DOCX
- `docwarp docx2md` for DOCX -> Markdown
- guided mode when run without arguments
- warning-first conversion with optional `--strict` exit behavior
- optional JSON conversion reports
- config file + style-map support
- batch conversion via directory input + `--glob`
- native Word equation round-tripping for `$...$` and `$$...$$` with equation style-map tokens

## Quick Start

```bash
docwarp md2docx ./input.md --output ./output.docx
docwarp docx2md ./input.docx --output ./output.md
```

Guided mode:

- Run `docwarp` with no arguments.
- Choose a file/folder interactively (or drag a path into the terminal).
- `docwarp` auto-detects direction and runs the matching conversion.

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

Run command-specific help for detailed examples:

```bash
docwarp --help
docwarp md2docx --help
docwarp docx2md --help
```

## Docs

- Install guide: `docs/install.md`
- Configuration and style maps: `docs/configuration.md`
- Agent instruction hub (canonical): `AGENTS.md`
- Strict mode and CI guidance: `docs/strict-mode.md`
- JSON report schema: `docs/report-schema.md`
- Warning code catalog: `docs/warnings.md`
- Homebrew/core submission guide: `docs/homebrew-core.md`
- Release runbook: `docs/release.md`
- Changelog: `CHANGELOG.md`

## License

Apache-2.0
