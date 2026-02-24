# Changelog

All notable changes to this project are documented in this file.

## [Unreleased]

## [0.1.1] - 2026-02-24

### Added

- Homebrew/core submission guide at `docs/homebrew-core.md`.
- Homebrew/core formula generation script at `scripts/generate_homebrew_core_formula.sh`.
- Updated installation guidance for release formula assets and tap-based install flow.
- `template-map` command for extracting reusable YAML/JSON style maps from DOCX/DOTX templates.
- Overwrite backup controls for both directions: `--no-backup`, `--backup-dir`, and `--backup-keep`.
- Native Word equation round-tripping for Markdown math (`$...$`, `$$...$$`) with style-map tokens (`equation_inline`, `equation_block`).
- Managed decryptor bootstrap for protected DOCX conversions: private Python venv, pinned `msoffcrypto-tool`, and wheel hash verification.
- Company-template fixture pack and integration tests for style-id/style-name alias maps and template numbering behavior.
- Agent CLI authoring docs: `AGENTS.md`, `docs/agent-equations.md`, and `docs/agent-template-pack.md`.
- Dependabot configuration for Cargo and GitHub Actions dependency updates.

### Changed

- Homebrew formula metadata and checksums were updated for published release artifacts.
- README/install docs were aligned with release distribution paths and naming.
- DOCX style-map behavior and diagnostics were refined for template compatibility and clearer validation feedback.
- CLI UX and docs were polished with updated examples, startup copy, and feature summaries.
- Repository tracking was cleaned to remove generated review artifacts and local scratch conversion files.

## [0.1.1-rc.1] - 2026-02-22

### Added

- Pre-release candidate cut for validating the `v0.1.1` release workflow and distribution assets.

## [0.1.0] - 2026-02-21

### Added

- Workspace scaffold for `docwarp-cli`, `docwarp-core`, `docwarp-md`, and `docwarp-docx`.
- Bidirectional CLI commands: `docwarp md2docx` and `docwarp docx2md`.
- Guided mode default when running `docwarp` with no subcommand.
- Batch conversion support for directory inputs with `--glob`.
- Config loading with auto-discovery of `.docwarp.yml`.
- Style-map loading and merge precedence across built-in defaults, config file values, and CLI overrides.
- Stable warning-code catalog and strict-mode exit behavior (`--strict`).
- Machine-readable JSON conversion report output (`--report`).
- Password flag support for protected DOCX input (`docx2md --password`).
- Fixture corpus, golden tests, round-trip tests, and reliability integration tests.
- Fidelity improvements for uneven table rows, empty cells, list transitions, inline formatting edge cases, line break normalization, image path handling, and `.dotx` fallback behavior.
- Expanded CLI help and docs coverage for config/style maps, strict mode CI usage, and JSON report schema.
- Release automation for tagged artifacts, checksum generation, release-time Homebrew formula publishing, and cross-platform binary smoke tests.

### Changed

- Heading-style fidelity was improved by preserving Markdown heading levels `h1` through `h6` with distinct DOCX heading mappings.
- Guided-mode defaults were improved for path selection and output placement.
- Homebrew formula metadata now targets `N10ELabs/docwarp` URLs.
- Project docs now include install instructions for binary and Homebrew flows.
