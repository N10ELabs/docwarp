# Agent Markdown Authoring Contract (Equations + Lists)

Use this contract when an LLM or agent generates Markdown that will be converted by `docwarp`.

Canonical entrypoint: start at [`AGENTS.md`](../AGENTS.md). This file is a detailed supporting reference for agents using the `docwarp` CLI. If any rule conflicts with `AGENTS.md`, follow `AGENTS.md`.

For reusable prompt templates built on this contract, see [Agent Template Pack](agent-template-pack.md).

## Required List Conventions

- Unordered lists: use `-` as the Markdown marker.
- Ordered lists: use Markdown ordered list syntax with `1.` / `2.` / `3.` (do not hand-type outline labels like `1.1`).
- Nested lists: indent child items consistently (minimum two spaces per level).
- Keep list-type transitions explicit. If switching between ordered and unordered, separate logical groups with a blank line.
- Use headings around lists when semantic section labels are needed; do not fake headings by hand-formatting list item text.

DOCX rendering notes:

- Unordered list marker shape in Word is controlled by `word/numbering.xml`, not whether Markdown used `-`, `*`, or `+`.
- `docwarp` default numbering emits round bullets for unordered lists and hierarchical numbering for nested ordered lists (`1.1`, `1.1.1`, ...).
- If heading-like visual styling is required on numbered or bulleted list items, map `list_number` and/or `list_bullet` to a heading style in `md_to_docx` while keeping list semantics.

Example style-map override:

```yaml
md_to_docx:
  list_number: Heading2
  list_bullet: ListBullet
```

## Required Delimiters

- Inline math must use `$...$`.
- Display math must use a standalone multiline `$$` block:

```markdown
$$
E = mc^2
$$
```

- Do not use `\(...\)` or `\[...\]`.

## Display Equation Placement Rules

- Put display equations in their own paragraph.
- Keep exactly one display equation per `$$...$$` block.
- Do not place sentence text on the same paragraph line as the display block.

Good:

```markdown
Regression objective:

$$
\hat{\beta} = \arg\min_{\beta} \| y - X\beta \|_2^2
$$
```

Avoid:

```markdown
Regression objective: $$ \hat{\beta} = \arg\min_{\beta} \| y - X\beta \|_2^2 $$
```

## TeX Patterns That Convert Cleanly

Use standard LaTeX math commands like these:

- Fractions and roots: `\frac{a}{b}`, `\sqrt{x}`
- Scripts: `x_i`, `x^2`, `x_i^2`
- Summation/product/integral: `\sum_{i=1}^{n}`, `\prod_{k=1}^{m}`, `\int_a^b`
- Limits/operators: `\min_{x}`, `\max_{x}`, `\lim_{n \to \infty}`, `\arg\min_{\beta}`
- Matrices: `\left[\begin{matrix} a & b \\ c & d \end{matrix}\right]`

## Patterns To Avoid In Agent Output

- `\begin{equation}...\end{equation}` and `\begin{align}...\end{align}`
- Custom macros like `\newcommand`, `\def`, or template-local command aliases
- Embedding equations in code fences
- Mixing markdown emphasis delimiters into math content

If an expression is outside the supported conversion path, `docwarp` can still emit an equation object but may fall back to linear text and raise `unsupported_feature`.

## Style Mapping Tokens

When generating a style map for equations, preserve these exact keys:

- `md_to_docx.equation_inline`
- `md_to_docx.equation_block`

Example:

```yaml
md_to_docx:
  equation_inline: EquationInline
  equation_block: Equation
```

## Copy/Paste Prompt Snippet For Upstream Agents

```text
When writing Markdown for docwarp:
- Use inline math as $...$ only.
- Use display math as multiline $$ blocks only.
- Put each display equation in its own paragraph with no surrounding sentence text on the same line.
- Prefer standard LaTeX math commands (\frac, \sqrt, \sum, \int, subscripts/superscripts, matrix, argmin/argmax).
- Do not use \(...\), \[...\], align/equation environments, or custom macros.
- Use '-' for unordered lists and Markdown numbering syntax for ordered lists.
- For nested ordered lists, use indentation levels instead of manually typing labels like '1.1' in text.
```
