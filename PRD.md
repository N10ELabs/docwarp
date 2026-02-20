# instruct PRD (MVP v0.1.0)

## Product overview and problem statement.
`instruct` is an open-source, lightweight CLI that converts documents bidirectionally between GitHub-Flavored Markdown (GFM) and Microsoft Word-compatible DOCX. The product solves a growing workflow gap between AI/agent-native markdown authoring and enterprise/client-facing Word deliverables.

The core problem is friction and quality loss when moving documents between markdown-centric tooling and Word-centric collaboration. Users need deterministic structure-preserving conversion with explicit warnings for unsupported features, while keeping distribution simple (single binary, no required external runtime).

## Goals and non-goals.
### Goals
- Provide two-way conversion:
  - `md -> docx` for polished Word output from markdown sources.
  - `docx -> md` for AI/agent-friendly transformation and editing.
- Guarantee structure-preserving conversion for core document constructs.
- Preserve style intent through built-in + optional user-defined style mapping.
- Handle images in both directions:
  - embed images into DOCX on `md -> docx`.
  - extract images to filesystem and rewrite markdown links on `docx -> md`.
- Emit both human-readable conversion summaries and optional JSON reports.
- Maintain warning-first behavior and support strict exit semantics.
- Ship cross-platform binaries (macOS, Linux, Windows) and Homebrew install path.

### Non-goals (MVP)
- Full visual/page-layout fidelity with Word.
- Tracked changes, comments, SmartArt, macros, equations, mail merge artifact support.
- Guaranteed TOC field round-trip fidelity.
- Batch/directory conversion mode.
- External conversion-engine dependency (for example, Pandoc).

## Personas and primary use cases.
### Persona 1: AI-first technical writer
- Writes docs in markdown (often generated/refined by an LLM) and needs Word output for stakeholders.
- Uses `instruct md2docx` with optional template/style map for brand alignment.

### Persona 2: Operations/legal/business editor
- Receives legacy or poorly structured DOCX files and needs markdown for agent-assisted revision.
- Uses `instruct docx2md`, edits markdown via AI workflow, converts back to DOCX.

### Persona 3: OSS/web documentation maintainer
- Converts Word documents into markdown for GitHub/web publishing.
- Uses `instruct docx2md` to migrate content and extracted media.

## Functional requirements.
1. CLI must expose explicit directional subcommands.
2. `md2docx` must parse GFM, map supported structures to DOCX OpenXML, and write valid DOCX zip packages.
3. `docx2md` must parse DOCX OpenXML, map supported structures to canonical document model, and render markdown.
4. Supported structure set (quality-guaranteed):
- headings/title
- paragraphs
- ordered/unordered lists
- inline emphasis/strong/code
- links
- blockquotes
- fenced code blocks
- tables
- images
5. Warning-first unsupported handling:
- default behavior is warn-and-continue.
- warnings include machine-readable codes.
6. Strict mode:
- if enabled and any warnings occur, process exits with code `2`.
7. Exit code contract:
- `0`: success
- `1`: fatal error
- `2`: success with warnings under strict mode
8. `md2docx` image rules:
- local images are embedded when readable.
- remote images are blocked by default and warned.
- remote embedding only occurs with explicit `--allow-remote-images`.
9. `docx2md` image rules:
- referenced media is extracted to assets directory.
- markdown image links point to extracted assets.
10. Optional JSON report output includes duration, stats, warnings, and result metadata.
11. Config + style-map precedence:
- built-ins < config file < CLI `--style-map`.

## CLI interface contract.
### Command 1
```bash
instruct md2docx <input.md> --output <output.docx> [--template <template.dotx>] [--style-map <map.yml>] [--config <instruct.yml>] [--report <report.json>] [--strict] [--allow-remote-images]
```

### Command 2
```bash
instruct docx2md <input.docx> --output <output.md> [--assets-dir <dir>] [--style-map <map.yml>] [--config <instruct.yml>] [--report <report.json>] [--strict]
```

### Runtime behavior
- Auto-load `.instruct.yml` from working directory when `--config` is omitted.
- Print conversion summary and warning list to stdout.
- Write report only when `--report` is provided.

## Config and style map schema.
### Config file (`.instruct.yml` or `--config`)
```yaml
markdown_flavor: gfm
style_map: ./style-map.yml
assets_dir: ./output_assets
default_template: ./template.dotx
unsupported_policy: warn_continue
```

### Config fields
- `markdown_flavor`: currently `gfm`.
- `style_map`: path to YAML/JSON style map.
- `assets_dir`: default extraction directory for `docx2md`.
- `default_template`: optional default `.dotx` for `md2docx`.
- `unsupported_policy`: `warn_continue | fail_fast | best_effort_silent`.

### Style map schema
```yaml
docx_to_md:
  Heading1: h1
  Heading2: h2
  Title: title
md_to_docx:
  h1: Heading1
  h2: Heading2
  title: Title
```

### Built-in style defaults
- `Title`, `Heading1`, `Heading2`, `Heading3`, `Normal`, `Quote`, `Code`, `ListBullet`, `ListNumber`, `Table`.

## Conversion fidelity and unsupported-feature policy.
### Fidelity target
- Structure-preserving round trips, not pixel-perfect rendering parity.
- Document intent and semantic blocks are prioritized over page-layout exactness.

### Unsupported policy
- Default: warn and continue.
- Warnings are emitted with stable codes:
  - `unsupported_feature`
  - `image_load_failed`
  - `remote_image_blocked`
  - `missing_media`
  - `invalid_style_map`
  - `invalid_template`
  - `corrupt_docx`
  - `nested_structure_simplified`
- Strict mode upgrades warnings to exit code `2` without hiding output artifacts.

## Architecture and module boundaries.
### Workspace layout
- `crates/instruct-cli`: command parsing, config loading, orchestration, reporting, exit behavior.
- `crates/instruct-core`: canonical document model, warnings, config types, report schema, style-map merge logic.
- `crates/instruct-md`: GFM parse/render adapter using `pulldown-cmark`.
- `crates/instruct-docx`: DOCX OpenXML read/write adapter using `zip` and `quick-xml`.

### Canonical model
- A shared intermediate `Document` model is mandatory between format adapters.
- Both directional conversions map through this model to enforce deterministic behavior.

### Dependency decisions
- CLI: `clap`
- Data/config/report: `serde`, `serde_json`, `serde_yaml`
- Markdown: `pulldown-cmark`
- DOCX/OpenXML: `zip`, `quick-xml`
- Error handling: `anyhow`, `thiserror`
- Optional remote image fetch: `reqwest` (blocking) behind explicit flag usage

### Security defaults
- No telemetry.
- No network activity unless `--allow-remote-images` is provided.

## Test strategy and acceptance criteria.
### Unit tests
- Style map precedence validation (`CLI > config > built-in`).
- Markdown parser/render behavior for core structures.
- DOCX adapter behavior for warning and round-trip core scenarios.

### Acceptance criteria
1. `md2docx` converts core markdown fixtures into Word-openable `.docx` files.
2. `docx2md` converts supported DOCX structures into valid markdown with extracted images.
3. `md -> docx -> md` preserves structure and content intent.
4. `docx -> md -> docx` preserves block-type composition and media references.
5. Strict mode yields exit code `2` when warnings are present.
6. JSON report output matches schema contract:
- `version`
- `direction`
- `input_path`
- `output_path`
- `duration_ms`
- `stats`
- `warnings`
- `success`

## Release, packaging, and licensing.
- Initial release target: `v0.1.0`.
- Distribution:
  - GitHub Releases with binaries for macOS, Linux, Windows.
  - Homebrew formula for install convenience.
- License: Apache-2.0.
- Release metadata includes compliance files and changelog references.

## Risks, mitigations, and phased milestones.
### Key risks
- OpenXML edge-case variation across Word-generated documents.
- Markdown feature mismatches that cannot cleanly round-trip.
- Image/path portability across environments.

### Mitigations
- Canonical intermediate model with explicit warning codes.
- Conservative MVP scope with documented non-goals.
- Strict mode for CI/automation workflows requiring warning-free outputs.
- Golden/fixture-based regression tests for core content patterns.

### Milestones
1. Foundation
- Workspace scaffold, core model, CLI interface, config/report contracts.
2. Conversion core
- Implement GFM parser/renderer and DOCX read/write pipelines for supported blocks.
3. Quality gates
- Add unit and round-trip tests, warning-code stabilization, strict-mode verification.
4. Release prep
- CI matrix, release automation, Homebrew packaging, v0.1.0 publish.
