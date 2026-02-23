# Agent CLI Instruction Hub

This is the canonical instruction hub for any agent using `docwarp` as a CLI tool.

It is intentionally scoped to authoring and style behavior for Markdown/DOCX conversion, not repository contribution workflow.

If guidance appears in multiple files, start here. If conflicts exist, `AGENTS.md` is the source of truth.

## Reference Docs (Top-Level)

Core authoring references:

- [`docs/agent-equations.md`](docs/agent-equations.md) for detailed Markdown equation + list authoring rules.
- [`docs/agent-template-pack.md`](docs/agent-template-pack.md) for reusable prompt templates.
- [`docs/configuration.md`](docs/configuration.md) for style-map/config tokens and examples.

Operational references:

- [`docs/strict-mode.md`](docs/strict-mode.md) for CI strict-mode behavior and exit codes.
- [`docs/report-schema.md`](docs/report-schema.md) for JSON report format.
- [`docs/warnings.md`](docs/warnings.md) for stable warning codes.
- [`docs/install.md`](docs/install.md) for installation paths and binaries.
- [`docs/release.md`](docs/release.md) for release runbook steps.
- [`docs/homebrew-core.md`](docs/homebrew-core.md) for Homebrew/core submission workflow.

## `/docs` Inventory

- `agent-equations.md`: detailed equations/lists conversion-safe rules and examples.
- `agent-template-pack.md`: copy/paste prompts for upstream agents.
- `configuration.md`: runtime config + style-map token contract.
- `strict-mode.md`: CI gate semantics (`--strict`, exit code behavior).
- `report-schema.md`: machine-readable output contract for `--report`.
- `warnings.md`: stable warning taxonomy used in docs/tests/report interpretation.
- `install.md`: user/operator installation instructions.
- `release.md`: maintainer release checklist and validation flow.
- `homebrew-core.md`: maintainer formula submission process.

## Purpose

Use these conventions when generating Markdown that will be converted with `docwarp` (`md2docx`) in any environment.

## Required Markdown Conventions

### Lists

- Unordered lists: use `-` as the marker.
- Ordered lists: use Markdown numbering syntax (`1.`, `2.`, `3.`), not hand-typed outline labels.
- Nested ordered lists: use indentation to express levels (minimum two spaces per level).
- Do not manually type labels like `1.1` or `1.1.1` in item text.
- If switching between ordered and unordered groups, separate groups with a blank line.
- Use real headings (`##`, `###`) for section labels above lists.

### Equations

- Inline math: `$...$` only.
- Display math: standalone multiline `$$...$$` blocks only.
- One display equation per block.
- Do not use `\(...\)`, `\[...\]`, `align`, `equation`, or custom macros.

## DOCX Style Mapping Conventions

When generating style maps, preserve these exact `md_to_docx` tokens:

- `equation_inline`
- `equation_block`
- `list_bullet`
- `list_number`

If list items must look like headings in DOCX, override list tokens to heading styles while keeping Markdown list semantics.

Example:

```yaml
md_to_docx:
  list_number: Heading2
  list_bullet: ListBullet
```
