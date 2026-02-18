````markdown
# Markdown to Modern Submission DOCX Converter

A Node.js script that converts a Markdown file into a properly formatted DOCX file using modern short-story manuscript submission standards (Shunn-style format).

The output is designed to open cleanly in Microsoft Word and Apple Pages, with:

- US Letter page size  
- 1-inch margins  
- Times New Roman, 12pt  
- Double-spaced body text  
- 0.5" first-line paragraph indent  
- Header on pages 2–N (`LASTNAME / Short Title / Page#`)  
- Title centered and in ALL CAPS  
- Author/contact block on first page  
- Word count displayed as “about N words” (rounded to nearest 100)  

---

# Installation

Requires:

- Node.js v18+ (tested on Node 22)
- npm

Initialize project:

```bash
npm init -y
````

Install dependencies:

```bash
npm install docx unified remark-parse remark-gfm unist-util-visit
```

---

# Usage

```bash
node md2sub.mjs input.md [output.docx]
```

If `output.docx` is omitted, the file `submission.docx` is created.

Example:

```bash
node md2sub.mjs story.md story.docx
```

**NOTE**: See [benchmark.md](benchmark.md) for feature test input and [submission.docx](submission.docx) for result.

---

# Input File Format

At the top of your Markdown file, define metadata as plain `Key: Value` lines:

```md
Author: Robert Raisch
LastName: Raisch
Address: 15 Mount Vernon Terrace
        3rd Floor
City: Newton
State: MA
PostalCode: 02465
Phone: (617) 331-0222
Email: raisch+words@gmail.com
ShortTitle: MD Converter
```

Leave one blank line after the metadata block.

---

# Example Input

```md
Author: Robert Raisch
LastName: Raisch
Address: 15 Mount Vernon Terrace
        3rd Floor
City: Newton
State: MA
PostalCode: 02465
Phone: (617) 331-0222
Email: raisch+words@gmail.com
ShortTitle: MD Converter

# The Markdown to Modern Submission Format Converter

## By Robert Raisch writing as R. L. Raisch

This is the first paragraph. This is the first paragraph.

This is the second paragraph.

***

This is the third paragraph.
```

---

# Features

## 1. Manuscript Formatting

* US Letter (8.5" x 11")
* 1" margins (1440 twips)
* Times New Roman 12pt
* Double-spaced body (`line = 480`)
* 0.5" first-line indent (720 twips)
* No extra space between paragraphs

---

## 2. First Page Layout

* Contact block at top-left (single-spaced)
* First line includes:

```
Author Name                                  about 3,200 words
```

* Title centered, in ALL CAPS
* Byline centered below title
* Body begins two lines below byline

---

## 3. Headers (Pages 2–N Only)

Header format:

```
LASTNAME / Short Title / Page#
```

* Right-aligned
* Page 1 header suppressed
* Pages 2+ automatically numbered

---

## 4. Supported Markdown Features

### Headings

```md
# Title
## Byline
### Section Heading
```

* `#` → Centered, bold, ALL CAPS
* Lower headings → Centered

---

### Paragraph Formatting

```md
This is **bold** text.
This is *italic* text.
```

* Bold and italics supported
* Inline code treated as plain text

---

### Scene Breaks

Markdown thematic breaks:

```md
---
***
```

Converted to centered manuscript scene break:

```
#
```

---

### Lists

```md
- Item one
- Item two

1. First
2. Second
```

Rendered with hanging indent in manuscript style.

---

### Tables (GFM)

```md
| Name | Value |
|------|-------|
| A    | 1     |
| B    | 2     |
```

Converted to DOCX tables with fixed-width columns compatible with Pages.

---

## 5. Word Count

* Counts words from the story body (excluding metadata/title/byline)
* Rounded to nearest 100
* Displayed as:

```
about 3,200 words
```

---

# Limitations

* No support for:

  * Images
  * Footnotes
  * Complex nested Markdown features
* Designed for short fiction manuscript submission format, not layout publishing

---

# License

Use freely for manuscript preparation.
