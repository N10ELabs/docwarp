# Release Runbook

## Prerelease Validation

Use a prerelease tag to validate the end-to-end release workflow before `v0.1.0`.

Example:

```bash
git tag v0.1.0-rc.1
git push origin v0.1.0-rc.1
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
git tag v0.1.0
git push origin v0.1.0
```

Then:

1. Confirm release smoke job passed.
2. Confirm `docs/install.md` commands work for the new tag.
3. Confirm `CHANGELOG.md` entry is final and published with the release.
