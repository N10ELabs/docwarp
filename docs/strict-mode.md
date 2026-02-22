# Strict Mode and CI

`--strict` turns warnings into a non-zero success exit code.

Exit-code contract:

- `0`: success (or success with warnings when strict is not enabled)
- `1`: fatal error
- `2`: conversion succeeded, but warnings were emitted and strict mode is active

## CLI Examples

Markdown to DOCX:

```bash
docwarp md2docx ./input.md --output ./output.docx --strict
```

DOCX to Markdown:

```bash
docwarp docx2md ./input.docx --output ./output.md --strict
```

With report output:

```bash
docwarp md2docx ./input.md --output ./output.docx --strict --report ./report.json
```

## CI Integration Guidance

Use strict mode in CI to block merges on unsupported or degraded conversions.

Example GitHub Actions step:

```yaml
- name: Verify docs conversion (strict)
  run: |
    set -euo pipefail
    docwarp md2docx ./fixtures/md/10-comprehensive.md \
      --output ./tmp/ci-check.docx \
      --strict \
      --report ./tmp/ci-check-report.json
```

If the command exits with `2`, the job fails and the warning summary plus JSON report can be used for diagnosis.

Recommended pattern:

- run conversion checks with `--strict`
- upload `--report` JSON as a CI artifact
- fail fast on any warning regressions
