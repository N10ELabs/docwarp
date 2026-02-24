# Homebrew/Core Submission

This runbook prepares `docwarp` for global install as:

```bash
brew install docwarp
```

That only works after `docwarp` is merged into `homebrew/core`.

## Prerequisites

- GitHub repository visibility is **public**.
- A release tag exists (for example `v0.1.1`).
- You can download the source archive URL:
  - `https://github.com/N10ELabs/docwarp/archive/refs/tags/v0.1.1.tar.gz`

## Generate a Core-Candidate Formula

From repo root:

```bash
scripts/generate_homebrew_core_formula.sh \
  N10ELabs \
  docwarp \
  v0.1.1 \
  packaging/homebrew-core/docwarp.rb
```

This computes `sha256` from the public tag archive and writes
`packaging/homebrew-core/docwarp.rb`.
If the repo is still private and archive download fails, make it public first
and rerun.

## Local Validation

```bash
brew tap --force homebrew/core
mkdir -p "$(brew --repository homebrew/core)/Formula/d"
cp packaging/homebrew-core/docwarp.rb "$(brew --repository homebrew/core)/Formula/d/docwarp.rb"
HOMEBREW_NO_INSTALL_FROM_API=1 brew install --build-from-source homebrew/core/docwarp
brew test homebrew/core/docwarp
HOMEBREW_NO_INSTALL_FROM_API=1 brew audit --new --strict homebrew/core/docwarp
```

If `docwarp` is already installed, reinstall from source:

```bash
HOMEBREW_NO_INSTALL_FROM_API=1 brew reinstall --build-from-source homebrew/core/docwarp
```

## Submit to Homebrew/Core

1. Fork `Homebrew/homebrew-core`.
2. Create a branch in your fork.
3. Copy formula into your fork at `Formula/d/docwarp.rb`.
4. Commit and open a PR to `Homebrew/homebrew-core`.
5. Address maintainer feedback and CI checks.

References:

- [Acceptable Formulae](https://docs.brew.sh/Acceptable-Formulae)
- [Formula Cookbook](https://docs.brew.sh/Formula-Cookbook)
- [How to Open a Homebrew Pull Request](https://docs.brew.sh/How-To-Open-a-Homebrew-Pull-Request)
