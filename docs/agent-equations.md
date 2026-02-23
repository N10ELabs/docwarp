# Agent Equation Authoring Contract

Use this contract when an LLM or agent generates Markdown that will be converted by `docwarp`.

For reusable prompt templates built on this contract, see [Agent Template Pack](agent-template-pack.md).

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
```
