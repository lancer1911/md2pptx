# Instruction: Convert an Existing Presentation into MD2PPTX Markdown

Convert the supplied PPTX or PDF presentation into Markdown for the local `md_to_pptx.py` tool.

The goal is not pixel-perfect reproduction. The goal is to preserve the presentation's logical structure, wording, numbers, legal conclusions, tables, images, diagrams, and reusable slide sequence in Markdown that can be regenerated as PPTX.

---

## 1. Required slide structure

- Use `#` only for the cover slide title.
- Use the next plain line after `#` for `Speaker | year` when the source has such information.
- Use `##` for every normal slide title.
- Separate slides with `---`.
- Preserve the original slide order unless the source is clearly disordered.
- If a dense slide must be split, keep the original title for the first split slide and use a clear suffix for later slides, such as `(continued)`, `Part 2`, or a meaningful subtopic.

Example:

```markdown
# Presentation Title
Speaker | 2026

---

## First Content Slide

- Key point.

---

## First Content Slide (continued)

- Remaining key point.
```

---

## 2. Supported Markdown elements

Use only the following Markdown elements unless necessary:

- `###` to `######` for section headings.
- Plain paragraphs for short explanations.
- `-` for bullets, with at most two levels where possible.
- Numbered lists only where sequence matters.
- Markdown pipe tables for real tables.
- `>` for conclusions, case holdings, or emphasized quotations.
- Fenced code blocks for code, commands, or configuration examples.
- `![alt](relative/path.png)` for images or screenshots that should be inserted later.
- `> **Speaker note:** ...` for speaker-only notes.
- `:::smartart` blocks for editable pseudo-SmartArt diagrams.

Do not use HTML, repeated blank lines, manual spacing, Mermaid, footnotes, or fake indentation to reproduce visual positions.

---

## 3. Fidelity rules

- Do not shorten legal conclusions, technical limitations, numerical figures, names, dates, case names, statute names, patent numbers, claim language, or party names.
- Preserve all numbers, percentages, amounts, time periods, docket numbers, patent numbers, dates, and citations.
- Preserve all tables as tables whenever possible.
- Preserve every readable label inside a diagram or shape.
- If a slide is too dense, split it into multiple `##` slides rather than reducing substance.
- Keep wording concise, but do not lose material legal, factual, technical, or procedural distinctions.
- Do not include recurring template elements such as page numbers, logos, confidentiality footers, decorative lines, or background text unless they are substantive content unique to that slide.

---

## 4. Common slide elements

### 4.1 Paragraphs and bullets

- Convert normal body text into short paragraphs or bullets.
- Use bullets when the source slide contains bullet-like statements or separated text boxes.
- Do not merge separate legal or technical points merely to save space.

### 4.2 Quotes and emphasized conclusions

Use block quotes for visually emphasized conclusions, court holdings, or key warnings:

```markdown
> Key conclusion or holding from the source slide.
```

If multiple quote lines form one visual block in the source, keep them as consecutive `>` lines.

### 4.3 Code, command, or configuration text

Use fenced code blocks:

````markdown
```bash
python md_to_pptx.py input.md template.pptx output.pptx
```
````

---

## 5. SmartArt / diagram conversion

When a slide contains a visual diagram made of boxes, arrows, circles, hierarchy nodes, timeline markers, or matrix quadrants, convert it into `:::smartart` whenever possible.

Use:

- `type="process"` for linear steps, arrows, or stage flows.
- `type="cycle"` for circular, iterative, or repeated-loop diagrams.
- `type="hierarchy"` for organization charts, team structures, tree structures, or parent-child relationships.
- `type="matrix"` for 2×2 or 2×3 quadrant charts.
- `type="pyramid"` for priority layers, value tiers, or progressive levels.
- `type="timeline"` for dated, staged, or milestone-based horizontal timelines.

Preserve the text in each shape exactly. Do not omit node labels merely because they are visually small.

### 5.1 Process

```markdown
:::smartart type="process" style="chevron" theme="blue"
- Evidence Collection
- Claim Mapping
- Complaint Filing
- Hearing
- Decision
:::
```

For vertical step lists or numbered cards:

```markdown
:::smartart type="process" style="numbered" theme="cyan"
- Complete pre-suit assessment
- Preserve samples and webpages
- File complaint and evidence list
:::
```

### 5.2 Cycle

```markdown
:::smartart type="cycle" theme="blue"
- Monitor
- Analyze
- Act
- Track
- Review
:::
```

### 5.3 Hierarchy

Use two spaces for each indentation level. Do not mix tabs and spaces.

```markdown
:::smartart type="hierarchy" theme="blue"
- Project Team
  - Litigation Team
    - Complaint Drafting
    - Hearing Preparation
  - Invalidity Team
    - Prior Art Search
    - Invalidity Petition
:::
```

### 5.4 Matrix

```markdown
:::smartart type="matrix" size="2x2" theme="cyan"
- High Impact / Low Probability / Watch Closely
- High Impact / High Probability / Act Immediately
- Low Impact / Low Probability / Monitor
- Low Impact / High Probability / Manage
:::
```

### 5.5 Pyramid

```markdown
:::smartart type="pyramid" theme="purple"
- Immediate Action
- Prepare
- Monitor
- Archive
:::
```

### 5.6 Timeline

```markdown
:::smartart type="timeline" theme="blue"
- Week 1 | Evidence Preservation
- Week 3 | Complaint Filed
- Month 2 | Invalidity Request
- Month 4 | Hearing
:::
```

Use `Label | Description` for each milestone.

### 5.7 When not to convert a diagram into SmartArt

If a diagram is too complex to convert reliably into `:::smartart`, use an image placeholder:

```markdown
![Original diagram: short description](images/slide-XX-diagram.png)
```

Then add bullets below only if the source slide itself contains readable text outside the image.

---

## 6. Table conversion rules

- Convert real tables into Markdown pipe tables.
- Preserve row headers and column headers.
- Do not merge cells using HTML.
- If the source table has merged cells, repeat the merged heading text in each relevant column or row.
- If a table is too large for one slide, split it into multiple slides and repeat the header row.
- Do not convert tables into screenshots unless the table is purely visual or impossible to read reliably.

Example:

```markdown
| Stage | Work | Output |
| --- | --- | --- |
| Assessment | Claim mapping and target confirmation | Initial memo |
| Evidence | Sample purchase and webpage notarization | Notarial certificate |
```

---

## 7. Images and screenshots

Use local image placeholders for images that should be inserted later:

```markdown
![Product comparison image](images/slide-05-product-comparison.png)
```

- Use meaningful alt text.
- Preserve captions if the source slide has captions.
- Do not describe a screenshot in excessive prose if an image placeholder is more faithful.
- If the image contains important readable labels, include those labels in bullets or SmartArt when possible.

---

## 8. Speaker notes

If the source presentation includes speaker notes, convert them into `> **Speaker note:** ...` at the end of the corresponding slide.

```markdown
> **Speaker note:** Explain that this slide summarizes the pre-suit evidence strategy.
```

- Do not mix speaker notes into visible slide body content.
- Preserve important instructions, timing notes, and oral explanation cues.

---

## 9. Splitting dense slides

When splitting a dense slide:

- Keep the original slide title for the first split slide.
- For subsequent split slides, use the same title plus `(continued)`, `Part 2`, or a meaningful subtopic.
- Do not remove content to fit one slide.
- If the original slide combines text, table, and diagram, consider splitting into separate slides for text/table/diagram.

Examples:

```markdown
## Evidence Strategy

...

---

## Evidence Strategy (continued)

...
```

or:

```markdown
## Evidence Strategy: Sample Preservation

...

---

## Evidence Strategy: Sales Data

...
```

---

## 10. Output requirements

- Output only Markdown.
- Do not add explanatory comments outside the Markdown.
- Do not include file names unless they are part of image placeholders or source content.
- Use `---` between slides.
- Preserve the source presentation's language unless instructed otherwise.
- If information is unreadable or uncertain, mark it as `[unclear]` rather than guessing.


> 注意：项目模板的最后一页会自动作为结束页/感谢页追加。除非用户特别要求，不要在 Markdown 正文中额外生成 Thank you / 感谢页。


## SmartArt 同页多行排列提示

转换器会对过长 SmartArt 进行同页多行排列：

- `process` 超过 5 个顶层节点时在同一页换行；
- `timeline` 超过 6 个时间节点时在同一页换行；
- `cycle` 超过 6 个节点时在同一页换行；
- `pyramid` 超过 5 层时在同一页换行；
- `matrix` 与 `hierarchy` 默认不自动换行。

即使支持同页多行排列，写作时仍建议尽量控制每页信息量。  
如果内容具有明确逻辑阶段，优先由作者主动拆成多页。

补充：同页多行排列会避免最后一行只有 1 个节点，例如 6 个节点排为 4+2，11 个节点排为 5+4+2。
