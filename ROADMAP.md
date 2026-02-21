# instruct Roadmap

Last updated: 2026-02-20

## Current State
- MVP scaffold is implemented and passes local tests.
- Bidirectional CLI exists:
  - `instruct md2docx ...`
  - `instruct docx2md ...`
- Core model, style-map precedence, warnings, strict mode, and JSON reports are in place.
- CI/release/homebrew scaffolding exists but release metadata is not finalized.

## Definition of Done for `v0.1.0`
- Reliable conversion for MVP-supported structures: headings, paragraphs, lists, tables, links, blockquotes, code blocks, images.
- Stable warning behavior with documented codes.
- Cross-platform build + test matrix green on macOS, Linux, Windows.
- Release artifacts published and install flow validated (including Homebrew).
- User-facing docs sufficient for first external users.

## Priority Lanes
- `P0`: must-have for `v0.1.0`.
- `P1`: strongly recommended for `v0.1.x`.
- `P2`: post-MVP enhancements.

## Phase 1: Reliability Hardening (`P0`)
- [x] Create fixture corpus under `fixtures/md` and `fixtures/docx` for all in-scope structures.
- [x] Add golden tests for `md -> docx` structure expectations.
- [x] Add golden tests for `docx -> md` markdown expectations.
- [ ] Add round-trip tests for:
  - [ ] `md -> docx -> md`
  - [ ] `docx -> md -> docx`
- [ ] Add explicit failure tests:
  - [ ] corrupt docx
  - [ ] missing media
  - [ ] invalid style map
  - [ ] invalid template path
  - [ ] unsupported feature warnings
- [ ] Lock warning code catalog in docs and tests.
- [ ] Improve parse/write determinism (stable ordering where applicable).

## Phase 2: Fidelity and Edge Cases (`P0`)
- [ ] Improve table handling for uneven rows and empty cells.
- [ ] Improve list fidelity for mixed ordered/unordered transitions.
- [ ] Improve inline formatting edge cases (nested emphasis/strong/code/link combinations).
- [ ] Preserve/normalize hard and soft line breaks predictably.
- [ ] Improve image handling:
  - [ ] explicit behavior for absolute vs relative paths
  - [ ] clearer remote image warning messages
  - [ ] enforce offline-by-default in tests
- [ ] Validate `.dotx` template integration with fallback behavior.

## Phase 3: CLI and Config UX (`P0`)
- [ ] Expand `--help` examples for both commands.
- [ ] Add config examples for `.instruct.yml` in docs.
- [ ] Add style-map examples (YAML + JSON) in docs.
- [ ] Add `--strict` behavior examples and CI integration guidance.
- [ ] Add machine-readable report schema documentation.

## Phase 4: Release Readiness (`P0`)
- [ ] Finalize project metadata:
  - [ ] replace `OWNER` placeholders
  - [ ] add Homebrew `sha256` values per artifact
- [ ] Validate GitHub Actions release workflow end-to-end on a prerelease tag.
- [ ] Add smoke tests for published binaries on macOS/Linux/Windows.
- [ ] Add install docs for:
  - [ ] binary download
  - [ ] Homebrew tap
- [ ] Publish `v0.1.0` and changelog.

## Phase 5: Short-Term Enhancements (`P1`)
- [ ] Add batch conversion mode (`--glob` or directory input) while keeping single-file mode default.
- [ ] Add compatibility mode flags (for example, stricter markdown output for docs platforms).
- [ ] Simplify CLI usage with config profiles and short command aliases so common permission-related flags/defaults (for example, remote image policy, strict mode, template/style-map paths) do not need to be repeated.
- [ ] Improve style-map validation with actionable error diagnostics.
- [ ] Native Word equation support for bidirectional `md <-> docx`, including configurable equation style mapping so users do not need manual Insert Equation restyling (for example, statistical-method appendices).
- [ ] Add performance benchmark suite for large documents.
- [ ] Add regression test pack from real-world anonymized docs.

## Phase 6: Post-MVP Scope (`P2`)
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
- Use milestone `v0.1.0` for all open `P0` tasks.
- Keep this file updated when:
  - a phase starts
  - a phase completes
  - priorities change

## Immediate Next 10 Tasks
- [x] Add fixture directories and first 10 canonical samples.
- [x] Add golden test harness helper utilities.
- [ ] Add strict-mode integration test for warning exit code `2`.
- [ ] Add invalid-style-map integration test.
- [ ] Add invalid-template integration test.
- [ ] Add missing-media integration test.
- [ ] Add docs for JSON report schema and sample output.
- [ ] Finalize release workflow placeholders.
- [ ] Draft `CHANGELOG.md` and release template.
- [ ] Create GitHub issues for all `P0` checkboxes.
