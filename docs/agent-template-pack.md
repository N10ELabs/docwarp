# Agent Template Pack (Markdown -> docwarp -> DOCX)

This pack gives reusable prompt templates for agents that generate Markdown intended for `docwarp`.

Canonical entrypoint: start at [`AGENTS.md`](../AGENTS.md). This file is a detailed supporting reference for agents using the `docwarp` CLI. If any rule conflicts with `AGENTS.md`, follow `AGENTS.md`.

Use these templates when you want consistent equation formatting, predictable round-tripping, and minimal manual cleanup in Word.

## Universal Equation Contract

Apply this contract in every template:

- Use inline math as `$...$`.
- Use display math as multiline `$$` blocks only.
- Put each display equation in its own paragraph.
- Use one display equation per `$$...$$` block.
- Do not use `\(...\)` or `\[...\]`.
- Do not use `align`, `equation`, or custom macro definitions.

## Universal List Contract

Apply these list rules in every template:

- Use `-` for unordered list items.
- Use standard Markdown ordered list syntax (`1.`, `2.`, `3.`) for ordered lists.
- For nested ordered lists, use indentation only; do not manually type outline labels like `1.1` or `1.1.1`.
- Keep nesting indentation consistent (minimum two spaces per level).
- If switching between ordered and unordered groups, separate list groups with a blank line.
- Use real headings (`##`, `###`) for section labels above lists rather than styling list text manually.

DOCX note:

- Word bullet shape and nested numbering patterns are controlled by DOCX numbering definitions; agents should focus on Markdown structure, not visual glyph hacks.

Example display equation block:

```markdown
$$
\hat{\beta} = \arg\min_{\beta} \| y - X\beta \|_2^2
$$
```

## Template 1: Statistical Methods Appendix

Use when generating methods-heavy content for research papers, audits, or technical appendices.

```text
You are writing Markdown for conversion with docwarp.

Task:
Write a statistical methods appendix for: [PROJECT OR STUDY NAME].

Audience:
[TECHNICAL AUDIENCE DESCRIPTION].

Requirements:
- Follow this equation contract exactly:
  - inline math: $...$
  - display math: multiline $$ blocks only
  - one display equation per block and in its own paragraph
  - no \(...\), no \[...\], no align/equation environments, no custom macros
- Follow this list contract exactly:
  - unordered lists use '-'
  - ordered lists use Markdown numbering syntax
  - nested ordered lists use indentation (no manual "1.1" text)
- Use concise technical prose.
- Prefer standard LaTeX math commands (\frac, \sqrt, \sum, \int, subscripts/superscripts, matrix, \arg\min/\arg\max).
- Include these sections in order:
  1) Model Setup
  2) Estimation Objective
  3) Distributional Assumptions
  4) Inference Procedure
  5) Diagnostics and Robustness Checks
  6) Symbol Table
- Include at least:
  - one inline equation
  - four display equations
  - one matrix expression
  - one argmin/argmax objective

Output:
Return only valid GitHub-Flavored Markdown.
```

## Template 2: Research Report With Embedded Math

Use when drafting a full report that mixes narrative and formal equations.

```text
You are writing Markdown for conversion with docwarp.

Task:
Draft a research report about: [TOPIC].

Audience:
[AUDIENCE].

Structure:
1) Executive Summary
2) Problem Statement
3) Method
4) Results
5) Limitations
6) Conclusion

Math and formatting rules:
- inline math: $...$
- display math: multiline $$ blocks only, each in its own paragraph
- no \(...\), no \[...\], no align/equation environments
- no custom macro definitions
- equations must use standard LaTeX commands only
- unordered lists use '-'
- ordered/nested ordered lists use Markdown numbering + indentation
- do not manually type outline labels like 1.1 in body text
- use clear variable definitions before first use

Style:
- professional and concise
- avoid unnecessary jargon
- keep equations tightly aligned to nearby prose explanations

Output:
Return Markdown only.
```

## Template 3: Technical Spec (Engineering + Quantitative Logic)

Use when writing implementation docs that include formulas, constraints, and scoring functions.

```text
You are writing Markdown for conversion with docwarp.

Task:
Write a technical specification for: [SYSTEM OR FEATURE].

Must include:
1) Scope
2) Definitions and Symbols
3) Core Formulation
4) Constraints
5) Optimization Objective
6) Edge Cases
7) Validation Plan

Equation rules:
- inline math: $...$
- display math: multiline $$ blocks in standalone paragraphs
- no \(...\), no \[...\], no align/equation environments, no custom macros
- use \arg\min or \arg\max for objectives where applicable
- include at least one summation and one matrix equation if relevant

List rules:
- unordered lists use '-'
- ordered lists use Markdown numbering syntax
- nested ordered lists rely on indentation only

Output constraints:
- Markdown only
- deterministic section ordering
- explicit variable definitions
- no code fences for equations
```

## Template 4: Minimal Conversion-Safe Math Snippet

Use when an upstream agent only needs to generate equation content (not full report prose).

```text
Generate Markdown equations only.

Rules:
- inline math with $...$
- display math with multiline $$ blocks only
- no \(...\), no \[...\], no align/equation environments, no custom macros
- if lists are included, use '-' for unordered and indentation-based nested numbering
- output must be valid GitHub-Flavored Markdown

Produce:
- 3 inline equations
- 3 display equations
- 1 matrix equation
- 1 argmin equation
```

## Optional Style-Map Reminder For Agent Pipelines

When an agent also generates style-map config for `md2docx`, keep these tokens:

- `md_to_docx.equation_inline`
- `md_to_docx.equation_block`
- `md_to_docx.list_bullet`
- `md_to_docx.list_number`

Example:

```yaml
md_to_docx:
  equation_inline: EquationInline
  equation_block: Equation
  list_bullet: ListBullet
  list_number: ListNumber
```

## Recommended Pipeline Pattern

1. Generate Markdown using one template above.
2. Validate math delimiters quickly (`$...$`, multiline `$$...$$`).
3. Run `docwarp md2docx`.
4. If strict mode is desired in CI, run with `--strict` and fail on warnings.
