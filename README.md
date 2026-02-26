# html2docx

A CLI tool that converts HTML files to DOCX (Word) documents, powered by [html-to-docx](https://github.com/nicetry99/html-to-docx).

## Installation

```bash
# Clone the project repository
git clone https://github.com/paulopes/html2docx.git

# Navigate to the project directory
cd html2docx

# Install dependencies
npm install

# Link the command globally so you can run `html2docx` from anywhere
npm link
```

After running `npm link`, the `html2docx` command is available system-wide. To remove it later, run `npm unlink -g html2docx`.

You can also run it without installing globally:

```bash
node bin/html2docx.js <input>
```

## Usage

```
html2docx [options] <input>
```

The simplest invocation converts an HTML file to a DOCX file in the same directory:

```bash
html2docx report.html          # creates report.docx
```

## Options

### Output

| Flag | Description |
|---|---|
| `-o, --output <path>` | Output DOCX file path. If omitted, replaces the `.html` extension with `.docx`. |

```bash
html2docx report.html -o output/final.docx
```

### Layout

| Flag | Description |
|---|---|
| `--landscape` | Landscape orientation (default: portrait). |
| `--margin <inches>` | Uniform page margin on all four sides, in inches. |
| `--margin-top <inches>` | Top margin in inches. Overrides `--margin`. |
| `--margin-bottom <inches>` | Bottom margin in inches. Overrides `--margin`. |
| `--margin-left <inches>` | Left margin in inches. Overrides `--margin`. |
| `--margin-right <inches>` | Right margin in inches. Overrides `--margin`. |

```bash
# Landscape with 0.75-inch margins everywhere
html2docx slides.html --landscape --margin 0.75

# 1-inch margins with a wider left gutter for binding
html2docx report.html --margin 1 --margin-left 1.5
```

Per-edge margin flags override the uniform `--margin` value, so you can set a baseline and then adjust individual sides.

### Typography

| Flag | Description |
|---|---|
| `--font <name>` | Default font name (default: Times New Roman). |
| `--font-size <points>` | Default font size in points (default: 11). |

```bash
html2docx report.html --font "Calibri" --font-size 12
```

### Header & Footer

| Flag | Description |
|---|---|
| `--header <file-or-html>` | Header HTML. Accepts a file path or an inline HTML string. Appears on every page. |
| `--footer-html <file-or-html>` | Footer HTML. Accepts a file path or an inline HTML string. |
| `--no-page-numbers` | Disable automatic page numbers in the footer. |

The `--header` and `--footer-html` flags are flexible: if the value is a path to an existing file, the file contents are used; otherwise the value is treated as inline HTML.

```bash
# Inline header
html2docx report.html --header "<p><b>ACME Corp</b> - Confidential</p>"

# Header from a file
html2docx report.html --header templates/header.html

# Disable page numbers
html2docx report.html --no-page-numbers
```

### Metadata

| Flag | Description |
|---|---|
| `--title <text>` | Document title (appears in file properties). |
| `--creator <text>` | Document author metadata. |

```bash
html2docx report.html --title "Q4 Report" --creator "Jane Doe"
```

### Miscellaneous

| Flag | Description |
|---|---|
| `--line-numbers` | Enable line numbering in the document. |
| `--lang <code>` | Language locale for spell checking (default: `en-US`). |
| `-V, --version` | Print the version number. |
| `-h, --help` | Display help. |

## Defaults

Out of the box (with no flags), the tool produces a document with:

- **Portrait** orientation
- **U.S. Letter** page size (8.5 x 11 in)
- **1-inch** top/bottom margins, **1.25-inch** left/right margins
- **Times New Roman** font
- **Page numbers** enabled in the footer
- Table rows that **don't split** across pages

## HTML Tips

### Page Breaks

Insert a page break anywhere in your HTML:

```html
<div class="page-break" style="page-break-after: always;"></div>
```

### Ordered List Styles

Customize list numbering with CSS `list-style-type`:

```html
<ol style="list-style-type: lower-alpha;">
  <li>First item</li>   <!-- a. First item -->
  <li>Second item</li>  <!-- b. Second item -->
</ol>
```

Supported types: `decimal`, `upper-alpha`, `lower-alpha`, `upper-roman`, `lower-roman`, `lower-alpha-bracket-end`, `decimal-bracket-end`, `decimal-bracket`.

You can also set a custom start number with `data-start`:

```html
<ol data-start="5">
  <li>Starts at 5</li>
</ol>
```

## License

This project is licensed under the [ISC License](https://opensource.org/licenses/ISC).

Copyright 2025 Paulo Lopes

Permission to use, copy, modify, and/or distribute this software for any
purpose with or without fee is hereby granted, provided that the above
copyright notice and this permission notice appear in all copies.

THE SOFTWARE IS PROVIDED "AS IS" AND THE AUTHOR DISCLAIMS ALL WARRANTIES
WITH REGARD TO THIS SOFTWARE INCLUDING ALL IMPLIED WARRANTIES OF
MERCHANTABILITY AND FITNESS. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY
SPECIAL, DIRECT, INDIRECT, OR CONSEQUENTIAL DAMAGES OR ANY DAMAGES
WHATSOEVER RESULTING FROM LOSS OF USE, DATA OR PROFITS, WHETHER IN AN ACTION
OF CONTRACT, NEGLIGENCE OR OTHER TORTIOUS ACTION, ARISING OUT OF OR IN
CONNECTION WITH THE USE OR PERFORMANCE OF THIS SOFTWARE.

