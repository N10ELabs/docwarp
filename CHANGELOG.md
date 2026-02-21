# Changelog

All notable changes to this project are documented in this file.

## [0.1.0] - Pending Release

### Added

- Workspace scaffold for `instruct-cli`, `instruct-core`, `instruct-md`, and `instruct-docx`.
- Bidirectional CLI:
  - `instruct md2docx`
  - `instruct docx2md`
- Config loading with auto-discovery of `.instruct.yml`.
- Style-map loading/merging with precedence:
  - built-in defaults
  - config file
  - CLI overrides
- Stable warning-code catalog and strict-mode exit behavior.
- Machine-readable JSON conversion report output (`--report`).
- Fixture corpus, golden tests, round-trip tests, and reliability integration tests.
- Fidelity improvements for:
  - uneven table rows and empty cells
  - mixed ordered/unordered list transitions
  - inline formatting edge cases
  - normalized line break rendering
  - image path handling and offline-by-default warning clarity
  - `.dotx` template usage with fallback behavior
- Expanded CLI help examples and Phase 3 documentation:
  - configuration/style-map examples
  - strict mode and CI guidance
  - JSON report schema reference
- Release automation improvements:
  - tagged release artifact publishing
  - artifact checksum generation
  - release-time Homebrew formula generation
  - cross-platform published-binary smoke tests

### Changed

- Homebrew formula metadata now targets `N10ELabs/instruct` URLs.
- Project docs now include install instructions for binary and Homebrew flows.
