# Markdown Conversion Benchmark

Use this file to benchmark Markdown -> DOCX conversion fidelity.

## 1. Headings

# Heading 1
## Heading 2
### Heading 3
#### Heading 4
##### Heading 5
###### Heading 6

## 2. Paragraph And Inline Styles

This paragraph includes **bold**, *italic*, ***bold+italic***, `inline code`, and a [link to example.com](https://example.com).

This line has an escaped asterisk: \*not italic\* and escaped backticks: \`not code\`.

## 3. Hard Line Breaks

Line one with two trailing spaces.  
Line two after hard break.

Line one with a backslash hard break.\
Line two after backslash break.

## 4. Blockquote

> This is a blockquote.
>
> It spans multiple lines and includes **inline formatting**.

## 5. Lists

### Unordered

- Item A
- Item B
- Item C

### Ordered

1. First
2. Second
3. Third

### Nested (may be simplified depending on parser)

- Parent 1
  - Child 1.1
  - Child 1.2
- Parent 2

### Task List

- [ ] Open task
- [x] Completed task

## 6. Code Blocks

```bash
echo "hello from fenced code"
for i in 1 2 3; do
  echo "$i"
done
```

```rust
fn main() {
    println!("hello rust");
}
```

    Indented code block line 1
    Indented code block line 2

## 7. Table

| Left | Center | Right |
|:-----|:------:|------:|
| a    |   b    |     c |
| 10   |   20   |    30 |

## 8. Images

### Local image

![Benchmark local image](fixtures/assets/benchmark.png)

## 9. Horizontal Rule

---

## 10. Mixed Content

Paragraph before list.

- Bullet with `inline code`
- Bullet with [link](https://openai.com)

> Quote after list.

Final paragraph after quote.

## 11. Raw HTML (expected warning/fallback behavior)

<span style="color:red">Raw HTML span</span>

## 12. Footnote Syntax (parser support dependent)

Footnote reference[^1].

[^1]: Footnote text.
