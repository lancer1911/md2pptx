# MD2PPTX

Convert structured Markdown files into PowerPoint presentations using a `.pptx` template. A single Python script with no web service, no cloud dependency, and no GUI — just a command-line tool you run locally.

---

## What it does

You write your slides in Markdown. The script reads your template, reproduces its chrome (background, colours, decorative shapes, footer) on every slide, and fills in your content. The output is a `.pptx` file you open directly in PowerPoint or Keynote.

Supported Markdown features:

| Syntax | Renders as |
| --- | --- |
| `# Title` + `Speaker \| Year` | Cover slide (uses template cover layout) |
| `## Slide title` | New content slide |
| `### Section heading` | Bold dark-blue sub-heading |
| `- bullet` / `* bullet` | Bullet point |
| `  - sub-bullet` | Indented bullet |
| `1. item` | Numbered list |
| `> quote` | Block quote with left grey bar |
| `**bold**` / `*italic*` / `~~strike~~` | Inline formatting |
| `` `code` `` | Inline monospace |
| `[text](url)` | Hyperlink (text only) |
| `![alt](path.png)` | Embedded local image |
| ` ```lang … ``` ` | Code block |
| `\| table \| … \|` | Table |
| `> **Speaker note:** …` | Written to PowerPoint Notes Pane |
| `---` | Slide separator |

---

## Requirements

- Python 3.10 or later
- `python-pptx`, `lxml`, `chardet`

---

## Installation

```bash
git clone https://github.com/yourname/md2pptx.git
cd md2pptx
bash setup.sh          # creates ./env and installs dependencies
```

After setup, the directory looks like this:

```
md2pptx/
├── env/                   # virtual environment (created by setup.sh)
├── md_to_pptx.py          # main script
├── template.pptx          # default template
├── run.sh                 # convenience wrapper
├── setup.sh               # one-time setup
├── layout_example.json    # optional layout config reference
├── INSTRUCTION_md_format.md
└── INSTRUCTION-convert_existing_presentation_to_MD.md
```

---

## Usage

```bash
# output filename defaults to <input>.pptx
bash run.sh my_slides.md

# explicit output filename
bash run.sh my_slides.md output.pptx

# use a different template
python3 md_to_pptx.py my_slides.md output.pptx --template other_template.pptx
```

---

## Markdown format

### Cover slide

```markdown
# Presentation Title
Speaker Name | 2025
```

The line immediately after `#` (no blank line between) is the speaker / year line. Keep it under ~20 characters to avoid wrapping. The `|` separator renders in cyan; the `|` itself renders in white, matching the template design.

### Content slides

```markdown
## Slide Title

### Section Heading

- First bullet point.
- Second bullet point.

> Key conclusion or quote that deserves visual emphasis.

---

## Next Slide
```

Use `---` on its own line to separate slides explicitly. The script also splits on `##` headings if `---` is absent.

### Speaker notes

```markdown
> **Speaker note:** This is only visible in the presenter view.
> Continue on the next line if needed.
```

Anything starting with `> **Speaker note:**` (case-insensitive) is written to the PowerPoint Notes Pane and does not appear on the slide.

---

## Customising the layout

If your template has different dimensions or shape positions, place a JSON config file next to it with the same base name (e.g. `my_template.json`). All fields are optional:

```json
{
  "slide_w": 13.333,
  "slide_h": 7.5,
  "title_l": 0.5,   "title_t": 0.3,   "title_w": 12.0,  "title_h": 0.8,
  "content_l": 0.5, "content_t": 1.2,  "content_w": 12.0, "content_h": 5.5,
  "font_main": "Microsoft YaHei",
  "title_color": "00B0F0",
  "h3_color": "005A87"
}
```

See `layout_example.json` for the full list of fields.

---

## Using AI to write the Markdown

Two prompt files are included to help you prepare Markdown with an AI assistant (Claude, ChatGPT, etc.):

- **`INSTRUCTION_md_format.md`** — paste this as your prompt when writing new slides from scratch or from unstructured material.
- **`INSTRUCTION-convert_existing_presentation_to_MD.md`** — paste this when converting an existing presentation into Markdown.

---

## Capacity guidelines

The content area of the default template holds comfortably:

| Item | Recommended | Hard limit |
| --- | --- | --- |
| Section headings per slide | 2–3 | 4 |
| Bullet points per slide | 5–8 | 10 |
| Characters per bullet | 30–50 | 60 |
| Table columns | 2–4 | 5 |
| Table rows | 3–6 | 8 |

When in doubt, split into two slides rather than cramming one.

---

## Limitations

- Remote image URLs are not fetched; use local file paths only.
- Clickable hyperlinks are rendered as plain text.
- The cover slide layout is detected automatically from the template; templates with only one slide are treated as content-only.
- Complex nested lists beyond two levels may not render as expected.

---

## License

MIT License

Copyright (c) 2025

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
