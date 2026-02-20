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
cargo run -p instruct-cli -- md2docx input.md --output output.docx
cargo run -p instruct-cli -- docx2md input.docx --output output.md
```

## Commands

```text
instruct md2docx <input.md> --output <output.docx> [--template <template.dotx>] [--style-map <map.yml>] [--config <instruct.yml>] [--report <report.json>] [--strict] [--allow-remote-images]
instruct docx2md <input.docx> --output <output.md> [--assets-dir <dir>] [--style-map <map.yml>] [--config <instruct.yml>] [--report <report.json>] [--strict]
```

## License

Apache-2.0
