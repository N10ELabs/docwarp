# Fixture Corpus

This directory contains canonical conversion fixtures used by P0 reliability tests.

- `md/`: markdown source fixtures (10 canonical samples).
- `docx/`: paired DOCX fixtures generated from markdown samples.
- `expected/docx2md/`: golden markdown outputs from DOCX fixtures.
- `assets/`: local assets referenced by fixtures.
- `company_templates/`: branded template/style-map stress corpus for company-style workflows.

Coverage includes all MVP in-scope structures:

- title/headings
- paragraphs
- ordered and unordered lists
- links
- blockquotes
- fenced code blocks
- tables
- images
