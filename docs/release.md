# Release Runbook

## One-Command Quick Action

After updating `[workspace.package].version` in `Cargo.toml`:

```bash
scripts/release_quick.sh 0.1.2
```

What this does:

- validates Cargo version + CLI `--version`
- runs workspace tests (unless `--skip-tests`)
- commits and pushes release prep to `main`
- pushes prerelease tag `vX.Y.Z-rc.1` (unless `--skip-rc`)
- pushes stable tag `vX.Y.Z` (triggers GitHub Release workflow)
- waits for release `docwarp.rb` and updates `N10ELabs/homebrew-tap`

## Prerelease Validation

Use a prerelease tag to validate the end-to-end release workflow before `v0.1.1`.

Example:

```bash
git tag v0.1.1-rc.1
git push origin v0.1.1-rc.1
```

Expected workflow behavior:

- Builds release binaries for Linux/macOS/Windows.
- Publishes GitHub release marked as prerelease.
- Uploads release assets:
  - binaries
  - `checksums.txt`
  - `docwarp.rb` with per-platform Homebrew checksums
- Runs smoke tests against the published assets on all three OSes.

Manual checks after workflow completion:

1. Open the prerelease page and confirm all assets exist.
2. Download one asset and run `--help`.
3. Verify `checksums.txt` matches downloaded binary.
4. Validate Homebrew formula asset installs and runs.

## Stable Release

After prerelease validation succeeds, create the stable tag:

```bash
git tag v0.1.1
git push origin v0.1.1
```

Then:

1. Confirm release smoke job passed.
2. Confirm `docs/install.md` commands work for the new tag.
3. Confirm `CHANGELOG.md` entry is final and published with the release.
