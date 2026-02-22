# docwarp Roadmap

Last updated: 2026-02-22

## Current State
- MVP scaffold is implemented and passes local tests.
- Bidirectional CLI exists:
  - `docwarp md2docx ...`
  - `docwarp docx2md ...`
- Core model, style-map precedence, warnings, strict mode, and JSON reports are in place.
- CI/release/homebrew scaffolding exists but release metadata is not finalized.
- Release plan is now to push through all `P1` and likely `P2` work before publishing `v0.1.0`.

## Definition of Done for `v0.1.0`
- Reliable conversion for MVP-supported structures: headings, paragraphs, lists, tables, links, blockquotes, code blocks, images.
- Stable warning behavior with documented codes.
- Cross-platform build + test matrix green on macOS, Linux, Windows.
- Release artifacts published and install flow validated (including Homebrew).
- User-facing docs sufficient for first external users.
- `P0` is fully complete, `P1` is complete, and `P2` is driven as far as feasible pre-release (remaining items can roll to `v0.1.x`).

## Priority Lanes
- `P0`: hard release gate for `v0.1.0`.
- `P1`: planned for completion before `v0.1.0`.
- `P2`: stretch scope likely to be pulled into `v0.1.0` where feasible; remaining items roll forward.

## Phase 1: Reliability Hardening (`P0`)
- [x] Create fixture corpus under `fixtures/md` and `fixtures/docx` for all in-scope structures.
- [x] Add golden tests for `md -> docx` structure expectations.
- [x] Add golden tests for `docx -> md` markdown expectations.
- [x] Add round-trip tests for:
  - [x] `md -> docx -> md`
  - [x] `docx -> md -> docx`
- [x] Add explicit failure tests:
  - [x] corrupt docx
  - [x] missing media
  - [x] invalid style map
  - [x] invalid template path
  - [x] unsupported feature warnings
- [x] Lock warning code catalog in docs and tests.
- [x] Improve parse/write determinism (stable ordering where applicable).

## Phase 2: Fidelity and Edge Cases (`P0`)
- [x] Improve table handling for uneven rows and empty cells.
- [x] Improve list fidelity for mixed ordered/unordered transitions.
- [x] Preserve nested list hierarchy in round-trips by carrying per-item list levels/types in the core model.
- [x] Improve inline formatting edge cases (nested emphasis/strong/code/link combinations).
- [x] Preserve/normalize hard and soft line breaks predictably.
- [x] Preserve blockquote paragraph breaks (multi-paragraph quote formatting) in markdown round-trips.
- [x] Preserve fenced code language labels for docwarp-generated DOCX round-trips.
- [x] Improve image handling:
  - [x] explicit behavior for absolute vs relative paths
  - [x] clearer remote image warning messages
  - [x] enforce offline-by-default in tests
- [x] Validate `.dotx` template integration with fallback behavior.

## Phase 3: CLI and Config UX (`P0`)
- [x] Expand `--help` examples for both commands.
- [x] Add config examples for `.docwarp.yml` in docs.
- [x] Add style-map examples (YAML + JSON) in docs.
- [x] Add `--strict` behavior examples and CI integration guidance.
- [x] Add machine-readable report schema documentation.

## Phase 4: Release Readiness (`P0`)
- [x] Finalize project metadata:
  - [x] replace `OWNER` placeholders
  - [x] add Homebrew `sha256` values per artifact
- [ ] Validate GitHub Actions release workflow end-to-end on a prerelease tag.
- [x] Add smoke tests for published binaries on macOS/Linux/Windows.
- [x] Add install docs for:
  - [x] binary download
  - [x] Homebrew tap
## Phase 5: Pre-Release Enhancements (`P1`)
- [x] Add batch conversion mode (`--glob` or directory input) while keeping single-file mode default.
- [X] Simplify CLI interface significantly. Take inspiration from great CLI apps like Codex
- [x] Work on CLI defaults
  - [x] `docwarp` with no subcommand enters guided mode.
  - [x] Guided flow supports drag/paste path or native picker.
  - [x] Input type is auto-detected (`.md` vs `.docx`) and conversion direction selected automatically.
  - [x] Default output is auto-written next to input without prompting for output path.
  - [x] Added guided slash-config menu for template/profile/remote-image policy toggles.
- [x] Add configs
- [x] Add ability to enter passwords for protected documents
  - [x] Add `--password` for `docx2md` and guided password prompt immediately after input selection for protected DOCX.
  - [x] Add protected-file detection and decryption path (requires Python `msoffcrypto-tool` when decrypting encrypted Office containers).
  - [ ] Test password functionality
  - [x] Rename project, docs, and crates to `docwarp`
- [ ] Improve style-map validation with actionable error diagnostics.
- [x] Improve heading-style fidelity in `md2docx`: preserve Markdown heading levels `h1`-`h6` as distinct DOCX heading styles (or configurable mappings) instead of collapsing `h4`-`h6`.
- [ ] Native Word equation support for bidirectional `md <-> docx`, including configurable equation style mapping so users do not need manual Insert Equation restyling (for example, statistical-method appendices).
- [ ] Add performance benchmark suite for large documents.
- [ ] Add regression test pack from real-world anonymized docs.
- [ ] Publish `v0.1.0` and changelog.

## Phase 6: Stretch Pre-Release Scope (`P2`)
- [ ] Add config profiles for common use cases (e.g., academic, technical, business)
- [ ] Add ability to load company templates and style-maps containing company branding and docx styles
- [ ] Footnotes/endnotes.
- [ ] Headers/footers/page breaks.
- [ ] Better TOC/section handling.
- [ ] Comment/track-changes awareness (at least warning-grade support).
- [ ] Optional plugin/provider architecture for custom mappings.

## Suggested Weekly Cadence
- Monday: plan and lock weekly target tasks.
- Midweek: conversion fidelity and tests first, docs second.
- Friday: release-health check (CI status, unresolved P0 issues, risk log update).

## Task Tracking Conventions
- Use GitHub issues with labels:
  - `roadmap:P0`, `roadmap:P1`, `roadmap:P2`
  - `area:cli`, `area:core`, `area:md`, `area:docx`, `area:release`, `area:docs`
- Use milestone `v0.1.0` for all open `P0` and in-scope `P1`/`P2` tasks targeted for the release.
- Keep this file updated when:
  - a phase starts
  - a phase completes
  - priorities change

## Immediate Next 10 Tasks
- [x] Add fixture directories and first 10 canonical samples.
- [x] Add golden test harness helper utilities.
- [x] Add strict-mode integration test for warning exit code `2`.
- [x] Add invalid-style-map integration test.
- [x] Add invalid-template integration test.
- [x] Add missing-media integration test.
- [x] Add docs for JSON report schema and sample output.
- [x] Finalize release workflow placeholders.
- [x] Draft `CHANGELOG.md` and release template.
- [ ] Create GitHub issues for all `P0` checkboxes.
