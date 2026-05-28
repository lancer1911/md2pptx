#!/usr/bin/env python3
"""
md_to_pptx.py — Convert a full-featured Markdown file to PPTX using a template.

Supported Markdown elements
────────────────────────────
  # H1              → Cover/title slide (title centred)
  ## H2             → New slide title
  ### H3            → Section heading (dark-blue bold)
  #### H4           → Sub-section heading (medium-blue)
  ##### H5          → Sub-sub heading (grey)
  - / * bullet      → Bullet point (•) with inline formatting
    - / * indent    → Sub-bullet (◦)
  1. ordered list   → Numbered list item
     1. indent      → Numbered sub-item
  **bold**          → Bold inline
  *italic*          → Italic inline
  ***bold+italic*** → Bold+italic inline
  `code`            → Monospace inline (Courier New)
  ~~strikethrough~~ → Strikethrough inline
  > **Speaker note:** … → Written to slide notes pane only
  > blockquote      → Block quote with a light-grey vertical bar
  ```lang … ```     → Fenced code block (monospace)
  | table |         → Real PPTX table with header + alternating rows
  ---               → Horizontal rule / section divider
  [text](url)       → Hyperlink (underlined blue)
  ![alt](url)       → Image placeholder caption

Key improvements over v1:
  ✓ Real python-pptx table shapes (not tab-hacked text)
  ✓ Auto font-size based on content density
  ✓ y-cursor layout: text and tables flow in document order
  ✓ Speaker notes written to PPTX notes pane
  ✓ Microsoft YaHei font throughout
  ✓ Template decorative elements recreated cleanly

Usage:
    python md_to_pptx.py input.md template.pptx output.pptx
"""

import sys
import re
import argparse
from pathlib import Path
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu

# ── Layout globals (set at runtime by load_layout()) ─────────────────────────
# Defaults match the bundled template.pptx; overridden when a different
# template is supplied or when a layout.json config file is present.
SLIDE_W   = Inches(11.42)
SLIDE_H   = Inches(6.42)
TITLE_L   = Inches(0.473)
TITLE_T   = Inches(0.202)
TITLE_W   = Inches(9.482)
TITLE_H   = Inches(0.637)
CONTENT_L = Inches(0.473)
CONTENT_T = Inches(1.354)
CONTENT_W = Inches(10.662)
CONTENT_H = Inches(3.466)
CONTENT_B = CONTENT_T + CONTENT_H

# Decorative shape defaults (bundled template)
RECT_L = Inches(0.519);  RECT_T = Inches(0.947);  RECT_W = Inches(0.717);  RECT_H = Inches(0.086)
LINE_L = Inches(0.519);  LINE_T = Inches(0.947);  LINE_W = Inches(4.863);  LINE_H = 0
SLIDE_NUM_L = Inches(8.672);  SLIDE_NUM_T = Inches(5.951)
SLIDE_NUM_W = Inches(2.665);  SLIDE_NUM_H = Inches(0.342)

# Populated by load_layout()
_DECO_ELEMENTS: list = []        # legacy fallback elements (content slides)
_COVER_DECO_ELEMENTS: list = []  # decorative elements for cover slides
_LAYOUT_SOURCE = None            # slide layout for content slides (from template)
_COVER_LAYOUT_SOURCE = None      # slide layout for cover slides (from template)

# Cover slide shape IDs (detected from template slide[0] when present)
# These are the named text boxes that hold the big white title and speaker/year
_COVER_TITLE_SHAPE_ID   = None   # e.g. "23" — big white title on cover
_COVER_SPEAKER_SHAPE_ID = None   # e.g. "29" — "Speaker | year" line
_TEMPLATE_HAS_COVER     = False  # True when template slide[0] is a real cover
_TEMPLATE_PATH          = None   # stored so _add_decorative can re-open it

# ── Design tokens ─────────────────────────────────────────────────────────────
FONT_MAIN = "Microsoft YaHei"
FONT_MONO = "Courier New"

C_TITLE      = RGBColor(0x00, 0xB0, 0xF0)
C_H3         = RGBColor(0x00, 0x5A, 0x87)
C_H4         = RGBColor(0x1A, 0x7D, 0xB5)
C_H5         = RGBColor(0x6B, 0x72, 0x80)
C_BODY       = RGBColor(0x1F, 0x29, 0x37)
C_QUOTE      = RGBColor(0x1F, 0x29, 0x37)
C_QUOTE_BAR  = RGBColor(0xC8, 0xCD, 0xD4)
C_CODE_FG    = RGBColor(0x1F, 0x29, 0x37)
C_LINK       = RGBColor(0x25, 0x63, 0xEB)
C_DIVIDER    = RGBColor(0xCB, 0xD5, 0xE1)
C_RECT       = RGBColor(0x76, 0xCA, 0xF2)   # template accent rect colour
C_LINE       = RGBColor(0x9E, 0x9E, 0x9F)   # template line colour
C_SLIDE_NUM  = RGBColor(0xAA, 0xAA, 0xAA)

# Table colours (from micknoise — proven beautiful)
C_TABLE_HDR     = RGBColor(0x1A, 0x1A, 0x2E)
C_TABLE_HDR_TXT = RGBColor(0xFF, 0xFF, 0xFF)
C_TABLE_EVEN    = RGBColor(0xEF, 0xF6, 0xFB)
C_TABLE_ODD     = RGBColor(0xFF, 0xFF, 0xFF)

SZ_H3   = 15
SZ_H4   = 13
SZ_H5   = 12
SZ_SUB  = 12
SZ_QUOT = 13
SZ_CODE = 11

TABLE_ROW_H = Inches(0.40)

def auto_body_size(n_lines: int) -> int:
    """Auto-scale body font based on content density (from micknoise)."""
    if n_lines <= 6:  return 18
    if n_lines <= 10: return 16
    if n_lines <= 14: return 14
    return 13


# ═══════════════════════════════════════════════════════════════════════════════
# Markdown parser
# ═══════════════════════════════════════════════════════════════════════════════

def parse_md(md_text: str) -> list:
    """
    Parse markdown into slide dicts.

    Slide boundaries are detected in two ways (both work, can be mixed):
      A) Explicit:  a line containing only "---" (with optional surrounding blanks)
      B) Implicit:  a ## or # heading starts a new slide when no "---" is used

    Title extraction also handles:
      "## Slide 3 — Title text"  →  "Title text"
      "## Slide 3 - Title text"  →  "Title text"
      "## Slide 3｜Title text"   →  "Title text"
      "## 3. Title text"         →  "Title text"
    """
    slides     = []
    current    = None
    in_code    = False
    code_buf   = []
    note_lines = []
    in_note    = False

    _SPKR      = re.compile(r'^>\s*\*\*Speaker\s+note:\*\*\s*(.*)', re.IGNORECASE)
    # Strip "Slide N — ", "Slide N - ", "Slide N｜", "N. " prefixes from titles
    _SLIDE_PFX = re.compile(
        r'^(?:##?\s*)?(?:Slide\s+\d+\s*[—–\-｜]+\s*|\d+[.、]\s*)', re.IGNORECASE)

    def _clean_title(raw_title: str) -> str:
        return _SLIDE_PFX.sub("", raw_title).strip()

    def _new_slide(title: str, cover: bool):
        nonlocal current, in_note, note_lines
        if current is not None and note_lines:
            current["note"] = " ".join(note_lines)
        current = {"title": _clean_title(title), "cover": cover,
                   "body_lines": [], "note": "", "speaker": ""}
        in_note = False; note_lines = []
        slides.append(current)

    for raw in md_text.splitlines():
        line = raw.strip()

        # ── Fenced code block ────────────────────────────────────────
        if line.startswith("```"):
            if not in_code:
                in_code = True; code_buf = [raw]
            else:
                in_code = False; code_buf.append(raw)
                if current is not None:
                    current["body_lines"].extend(code_buf)
                code_buf = []
            continue
        if in_code:
            code_buf.append(raw); continue

        # ── Explicit slide separator "---" ───────────────────────────
        # Treat a lone --- as a slide break (body_lines may hold
        # a pending title from the previous section; here we just
        # mark the boundary by not creating a new slide — the next
        # heading will do that). We simply skip the line so it doesn't
        # end up as a stray hr in body content.
        if re.match(r'^-{3,}$', line) and current is not None:
            # Flush any trailing blanks, then mark separator seen.
            # Next ## will start the new slide normally.
            while current["body_lines"] and not current["body_lines"][-1].strip():
                current["body_lines"].pop()
            # Also flush pending speaker note
            if note_lines:
                current["note"] = " ".join(note_lines); note_lines = []
            in_note = False
            continue

        # ── H1 → cover slide ─────────────────────────────────────────
        if re.match(r"^# [^#]", line):
            _new_slide(line[2:].strip(), cover=True)
            # Peek at next non-blank line: if it looks like "Name | Year"
            # (no # or - prefix, short, possibly contains |), treat it as
            # the speaker/year subtitle and store it in the slide dict.
            continue

        # ── Line immediately after a cover H1 ────────────────────────
        # Capture speaker/year line: plain text after a cover slide start
        if (current is not None and current["cover"]
                and not current["body_lines"]   # nothing in body yet
                and line
                and not line.startswith("#")
                and not line.startswith("-")
                and not line.startswith("*")
                and not line.startswith("|")
                and not line.startswith(">")):
            current["speaker"] = line
            continue

        # ── H2 → new slide ───────────────────────────────────────────
        if re.match(r"^## [^#]", line):
            _new_slide(line[3:].strip(), cover=False)
            continue

        # ── Body content ─────────────────────────────────────────────
        if current is not None:
            mn = _SPKR.match(line)
            if mn:
                note_lines = [mn.group(1).strip()]; in_note = True; continue
            if in_note and line.startswith(">"):
                note_lines.append(line.lstrip("> ").strip()); continue
            in_note = False
            if note_lines:
                current["note"] = " ".join(note_lines); note_lines = []
            current["body_lines"].append(raw)

    if current is not None and note_lines:
        current["note"] = " ".join(note_lines)

    for s in slides:
        while s["body_lines"] and not s["body_lines"][-1].strip():
            s["body_lines"].pop()

    return slides


# ═══════════════════════════════════════════════════════════════════════════════
# Body classifier → segments
# ═══════════════════════════════════════════════════════════════════════════════

def classify_body(body_lines: list) -> list:
    """Split body_lines into typed segments for y-cursor rendering."""
    segments  = []
    cur_text  = []
    cur_table = None
    cur_code  = None
    cur_quote = None
    in_code   = False
    ord_cnt   = {}

    def flush_text():
        if cur_text:
            segments.append(("text", list(cur_text)))
            cur_text.clear()

    def flush_table():
        nonlocal cur_table
        if cur_table:
            segments.append(("table", list(cur_table)))
            cur_table = None

    def flush_code():
        nonlocal cur_code
        if cur_code is not None:
            segments.append(("code", list(cur_code)))
            cur_code = None

    def flush_quote():
        nonlocal cur_quote
        if cur_quote:
            segments.append(("quote", "\n".join(cur_quote)))
            cur_quote = None

    _IMG_LINE = re.compile(r'^!\[([^\]]*)\]\(([^)]*)\)$')   # image on its own line

    for raw in body_lines:
        line = raw.strip()

        # ── Fenced code block ─────────────────────────────────────────
        if line.startswith("```"):
            if not in_code:
                in_code = True; cur_code = []; flush_text(); flush_table()
            else:
                in_code = False; flush_code()
            continue
        if in_code:
            if cur_code is not None: cur_code.append(raw)
            continue

        # ── Blockquote ───────────────────────────────────────────────
        # Render contiguous quote lines as a dedicated block with a
        # vertical bar on the left, instead of mixing them into a regular
        # text box. This gives the PPTX the same visual cue as Markdown
        # quote blocks in modern editors.
        if line.startswith("> "):
            flush_text(); flush_table()
            if cur_quote is None: cur_quote = []
            cur_quote.append(line[2:].strip())
            continue
        else:
            flush_quote()

        # ── Table row ────────────────────────────────────────────────
        if line.startswith("|"):
            flush_text()
            if cur_table is None: cur_table = []
            if not re.match(r'^\|[-| :]+\|$', line):
                cur_table.append(line)
            continue
        else:
            flush_table()

        # ── Image on its own line → image block ──────────────────────
        mi = _IMG_LINE.match(line)
        if mi:
            cur_text.append({"type": "image", "alt": mi.group(1), "src": mi.group(2)})
            continue

        # ── Headings H3–H6 (all render as page sub-headings) ─────────
        m = re.match(r"^(#{3,6}) (.+)$", line)
        if m:
            level = min(len(m.group(1)), 5)   # cap at H5 visually
            cur_text.append({"type": "heading", "text": m.group(2).strip(), "level": level})
            ord_cnt = {}; continue   # reset numbering per section

        # ── Horizontal rule (within body — not a slide separator) ─────
        if re.match(r"^[-*_]{3,}$", line):
            cur_text.append({"type": "hr"}); continue

        # ── Ordered list: "1." or "1)" style ─────────────────────────
        m = re.match(r"^( {0,3})(\d+)[.)]\s+(.+)$", raw)
        if m:
            ind = 1 if len(m.group(1)) >= 2 else 0
            ord_cnt[ind] = ord_cnt.get(ind, 0) + 1
            cur_text.append({"type": "ordered", "text": m.group(3).strip(),
                             "indent": ind, "number": ord_cnt[ind]}); continue

        # ── Unordered bullet: "-", "*", "+" ──────────────────────────
        m = re.match(r"^( {0,3})[-*+] (.+)$", raw)
        if m:
            ind = 1 if len(m.group(1)) >= 2 else 0
            cur_text.append({"type": "bullet", "text": m.group(2).strip(), "indent": ind}); continue

        # ── Blank line ───────────────────────────────────────────────
        if not line:
            cur_text.append({"type": "blank"}); continue

        # ── Plain paragraph ──────────────────────────────────────────
        cur_text.append({"type": "para", "text": line})

    flush_text(); flush_table(); flush_code(); flush_quote()
    return segments


# ═══════════════════════════════════════════════════════════════════════════════
# Inline markdown parser
# ═══════════════════════════════════════════════════════════════════════════════

_INLINE_RE = re.compile(
    r"!\[([^\]]*)\]\(([^)]*)\)"
    r"|\[([^\]]+)\]\(([^)]+)\)"
    r"|\*\*\*(.+?)\*\*\*"
    r"|\*\*(.+?)\*\*"
    r"|\*(.+?)\*"
    r"|~~(.+?)~~"
    r"|`(.+?)`"
)

def parse_inline(text: str) -> list:
    """Return list of (text, bold, italic, strike, code, url) tuples."""
    runs = []
    pos  = 0
    for m in _INLINE_RE.finditer(text):
        if m.start() > pos:
            runs.append((text[pos:m.start()], False, False, False, False, None))
        g = m.groups()
        if   g[0] is not None: runs.append((f"[图片: {g[0]}]", False, True, False, False, None))
        elif g[2] is not None: runs.append((g[2], False, False, False, False, g[3]))
        elif g[4] is not None: runs.append((g[4], True, True, False, False, None))
        elif g[5] is not None: runs.append((g[5], True, False, False, False, None))
        elif g[6] is not None: runs.append((g[6], False, True, False, False, None))
        elif g[7] is not None: runs.append((g[7], False, False, True, False, None))
        elif g[8] is not None: runs.append((g[8], False, False, False, True, None))
        pos = m.end()
    if pos < len(text):
        runs.append((text[pos:], False, False, False, False, None))
    return runs or [(text, False, False, False, False, None)]


# ═══════════════════════════════════════════════════════════════════════════════
# python-pptx text helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _apply_run(run, text, sz_pt, bold=False, italic=False, strike=False,
               code=False, color: RGBColor = None, url=None):
    run.text = text
    run.font.name      = FONT_MONO if code else FONT_MAIN
    run.font.size      = Pt(SZ_CODE if code else sz_pt)
    run.font.bold      = bold
    run.font.italic    = italic
    run.font.color.rgb = color or C_BODY
    if strike:
        run._r.get_or_add_rPr().set("strike", "sngStrike")
    if url:
        run.hyperlink.address = url
        run.font.underline = True
        run.font.color.rgb = C_LINK


def _add_runs(para, text: str, sz_pt, color=None, bold_all=False, italic_all=False):
    for (t, b, i, s, code, url) in parse_inline(text):
        r = para.add_run()
        c = C_LINK if url else (C_CODE_FG if code else (color or C_BODY))
        _apply_run(r, t, sz_pt,
                   bold=bold_all or b, italic=italic_all or i,
                   strike=s, code=code, color=c, url=url)


def _new_tf(slide, left, top, width, height):
    """Add a word-wrap textbox and return its text_frame."""
    tb = slide.shapes.add_textbox(int(left), int(top), int(width), int(height))
    tf = tb.text_frame
    tf.word_wrap = True
    txBody = tf._txBody
    for p in txBody.findall(qn("a:p")):
        txBody.remove(p)
    return tf


def _no_bullet(para):
    pPr = para._p.get_or_add_pPr()
    if pPr.find(qn("a:buNone")) is None:
        etree.SubElement(pPr, qn("a:buNone"))


def _heading_para(tf, text, level, sz_override=None):
    col = {3: C_H3, 4: C_H4, 5: C_H5}.get(level, C_H3)
    sz  = sz_override or {3: SZ_H3, 4: SZ_H4, 5: SZ_H5}.get(level, SZ_H3)
    p   = tf.add_paragraph()
    p.space_before = Pt(6)
    _no_bullet(p)
    _add_runs(p, text, sz, col, bold_all=(level == 3))
    return p


def _bullet_para(tf, text, indent, sz):
    MAR0, MAR1, IND = 342900, 685800, -342900
    p   = tf.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(MAR1 if indent else MAR0))
    pPr.set("indent", str(IND))
    for el in pPr.findall(qn("a:buNone")): pPr.remove(el)
    buChar = etree.SubElement(pPr, qn("a:buChar"))
    buChar.set("char", "◦" if indent else "•")
    buFont = etree.SubElement(pPr, qn("a:buFont"))
    buFont.set("typeface", FONT_MAIN)
    p.space_before = Pt(3)
    _add_runs(p, text, sz)
    return p


def _ordered_para(tf, text, indent, number, sz):
    MAR0, MAR1, IND = 342900, 685800, -342900
    p   = tf.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(MAR1 if indent else MAR0))
    pPr.set("indent", str(IND))
    for el in pPr.findall(qn("a:buNone")): pPr.remove(el)
    buAN = etree.SubElement(pPr, qn("a:buAutoNum"))
    buAN.set("type", "arabicPeriod")
    buAN.set("startAt", str(number))
    buFont = etree.SubElement(pPr, qn("a:buFont"))
    buFont.set("typeface", FONT_MAIN)
    p.space_before = Pt(3)
    _add_runs(p, text, sz)
    return p


def _quote_para(tf, text, sz):
    # Fallback for old text-box rendering paths. The main renderer uses
    # _add_quote_block(), which draws a real vertical bar shape.
    p   = tf.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(274638))
    pPr.set("indent", str(0))
    _no_bullet(p)
    p.space_before = Pt(4)
    p.space_after  = Pt(4)
    _add_runs(p, text, sz, color=C_QUOTE)
    return p


def _estimate_quote_height(text: str, width_emu, sz_pt: int) -> int:
    """Approximate quote box height for y-cursor layout."""
    width_in = max(width_emu / 914400, 1.0)
    # For Chinese-heavy text, chars are visually wider; this value is
    # intentionally conservative to avoid clipping.
    chars_per_line = max(int(width_in * (54 / 9.5) * (16 / max(sz_pt, 1))), 18)
    logical_lines = 0
    for part in (text or " ").splitlines():
        logical_lines += max(1, (len(part) + chars_per_line - 1) // chars_per_line)
    return int(Pt(sz_pt + 9) * logical_lines + Pt(6))


def _add_quote_block(slide, quote_text: str, left, top, width, sz_pt: int) -> float:
    """Render a Markdown blockquote as a PPT block with a vertical bar.

    This uses a real PowerPoint shape for the bar, rather than relying on
    paragraph-border XML that is not consistently honoured by PowerPoint.
    """
    bar_w   = Inches(0.055)
    gap     = Inches(0.22)
    txt_l   = int(left + bar_w + gap)
    txt_w   = int(width - bar_w - gap)
    height  = _estimate_quote_height(quote_text, txt_w, sz_pt)

    bar = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, int(left), int(top + Pt(4)), int(bar_w), int(height - Pt(8))
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_QUOTE_BAR
    bar.line.fill.background()

    tf = _new_tf(slide, txt_l, int(top), txt_w, height)
    for idx, line in enumerate(quote_text.splitlines() or [""]):
        p = tf.add_paragraph()
        _no_bullet(p)
        p.space_before = Pt(0 if idx == 0 else 3)
        _add_runs(p, line, sz_pt + 2, color=C_QUOTE)

    return int(top) + height


def _plain_para(tf, text, sz):
    p = tf.add_paragraph()
    _no_bullet(p)
    _add_runs(p, text, sz)
    return p


def _spacer_para(tf, sz_pt=4):
    p = tf.add_paragraph()
    _no_bullet(p)
    r = p.add_run()
    r.font.size = Pt(sz_pt)
    return p


def _render_text_blocks(tf, blocks: list, sz_pt: float):
    last = None
    for block in blocks:
        bt = block.get("type")
        if bt == "blank":
            if last not in (None, "blank"): _spacer_para(tf, 4)
            last = bt; continue
        if bt == "heading":
            _heading_para(tf, block["text"], block["level"])
        elif bt == "bullet":
            _bullet_para(tf, block["text"], block["indent"], sz_pt)
        elif bt == "ordered":
            _ordered_para(tf, block["text"], block["indent"], block["number"], sz_pt)
        elif bt == "quote":
            _quote_para(tf, block["text"], sz_pt)
        elif bt == "para":
            _plain_para(tf, block["text"], sz_pt)
        elif bt == "image":
            alt = block.get("alt") or block.get("src", "")
            p   = tf.add_paragraph()
            _no_bullet(p)
            r = p.add_run()
            r.text           = f"[图片: {alt}]" if alt else "[图片]"
            r.font.name      = FONT_MAIN
            r.font.size      = Pt(sz_pt - 1)
            r.font.italic    = True
            r.font.color.rgb = C_QUOTE
        elif bt == "hr":
            p = tf.add_paragraph()
            _no_bullet(p)
            r = p.add_run()
            r.text = "─" * 55
            r.font.size = Pt(8)
            r.font.color.rgb = C_DIVIDER
        last = bt


# ═══════════════════════════════════════════════════════════════════════════════
# Real PPTX table (from micknoise — proven beautiful)
# ═══════════════════════════════════════════════════════════════════════════════

def _add_pptx_table(slide, table_rows: list, left, top, width) -> float:
    parsed = []
    for row in table_rows:
        cells = [c.strip() for c in row.strip().strip("|").split("|")]
        parsed.append(cells)
    if not parsed: return int(top)

    ncols  = max(len(r) for r in parsed)
    nrows  = len(parsed)
    height = int(TABLE_ROW_H * nrows)

    tbl = slide.shapes.add_table(
        nrows, ncols, int(left), int(top), int(width), height
    ).table

    col_w = int(width / ncols)
    for ci in range(ncols):
        tbl.columns[ci].width = col_w

    for ri, row in enumerate(parsed):
        is_hdr = (ri == 0)
        for ci in range(ncols):
            txt  = row[ci] if ci < len(row) else ""
            cell = tbl.cell(ri, ci)
            cell.fill.solid()
            if is_hdr:
                cell.fill.fore_color.rgb = C_TABLE_HDR
            elif ri % 2 == 0:
                cell.fill.fore_color.rgb = C_TABLE_EVEN
            else:
                cell.fill.fore_color.rgb = C_TABLE_ODD
            tf = cell.text_frame
            tf.word_wrap = True
            p  = tf.paragraphs[0]
            for (t, b, i, s, code, url) in parse_inline(txt):
                run = p.add_run()
                tc  = C_TABLE_HDR_TXT if is_hdr else C_BODY
                _apply_run(run, t, 13, bold=is_hdr or b, italic=i,
                           strike=s, code=code, color=tc, url=url)

    return int(top) + height


# ═══════════════════════════════════════════════════════════════════════════════
# Code block textbox
# ═══════════════════════════════════════════════════════════════════════════════

def _add_code_block(slide, code_lines: list, left, top, width) -> float:
    n      = len(code_lines) or 1
    height = int(Pt(SZ_CODE + 4) * n + Pt(8))
    tf     = _new_tf(slide, int(left), int(top), int(width), height)
    for line in code_lines:
        p = tf.add_paragraph()
        _no_bullet(p)
        p.space_before = Pt(0)
        r = p.add_run()
        r.text           = line or " "
        r.font.name      = FONT_MONO
        r.font.size      = Pt(SZ_CODE)
        r.font.color.rgb = C_CODE_FG
    return int(top) + height


# ═══════════════════════════════════════════════════════════════════════════════
# Slide decorative elements (replicate template visuals)
# ═══════════════════════════════════════════════════════════════════════════════

def _copy_layout_chrome(src_layout, dst_slide):
    """
    Copy non-placeholder shapes from a slide layout onto a slide.
    Handles picture relationships correctly (copies the media part).
    This is how template logos, footers, watermarks etc. are preserved.
    """
    import copy as _copy
    P  = "http://schemas.openxmlformats.org/presentationml/2006/main"
    A  = "http://schemas.openxmlformats.org/drawingml/2006/main"
    Rn = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    spTree_src = src_layout._element.find(f"{{{P}}}cSld/{{{P}}}spTree")
    if spTree_src is None:
        spTree_src = src_layout._element.find(f".//{{{P}}}spTree")
    spTree_dst = dst_slide._element.find(f".//{{{P}}}spTree")
    if spTree_src is None or spTree_dst is None:
        return

    for child in spTree_src:
        tag = child.tag.split("}")[1] if "}" in child.tag else child.tag

        # Only copy real content shapes, not spTree metadata
        if tag == "sp":
            # Skip placeholders (they inherit from layout automatically)
            if child.find(f".//{{{P}}}ph") is not None:
                continue
        elif tag not in ("pic", "cxnSp", "graphicFrame"):
            continue

        el = _copy.deepcopy(child)

        # Remap picture relationship IDs
        blip = el.find(f".//{{{A}}}blip")
        if blip is not None:
            old_rId = blip.get(f"{{{Rn}}}embed")
            if old_rId:
                src_rel = src_layout.part.rels.get(old_rId)
                if src_rel:
                    new_rId = dst_slide.part.relate_to(
                        src_rel.target_part, src_rel.reltype)
                    blip.set(f"{{{Rn}}}embed", new_rId)

        spTree_dst.append(el)


def _add_decorative(slide, slide_num: int, cover: bool = False):
    """
    Stamp template chrome onto a slide.

    For cover slides: copy all shapes from the template cover slide verbatim
    (background image, coloured rectangles, etc.) then write title/speaker
    into the detected text box IDs — no separate title textbox is drawn.

    For content slides: copy non-placeholder shapes from the content slide
    layout, then add a slide number.
    """
    import copy as _copy

    if cover and _TEMPLATE_HAS_COVER:
        # Clone every shape from the template cover slide into this slide
        P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
        A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
        Rn   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        spTree_dst = slide._element.find(f".//{{{P_NS}}}spTree")

        # We need the source slide's part to remap picture relationships
        # _COVER_DECO_ELEMENTS are deep-copies of the XML; we need the
        # original slide part for relationship lookup.
        # Re-open template here to get the live part.
        import copy as _copy
        from pptx import Presentation as _Prs
        _tprs       = _Prs(_TEMPLATE_PATH)
        src_slide   = _tprs.slides[0]

        for shape in src_slide.shapes:
            el     = _copy.deepcopy(shape._element)
            # Remap picture (blip) relationship IDs
            blip = el.find(f".//{{{A_NS}}}blip")
            if blip is not None:
                old_rId = blip.get(f"{{{Rn}}}embed")
                if old_rId:
                    src_rel = src_slide.part.rels.get(old_rId)
                    if src_rel:
                        new_rId = slide.part.relate_to(
                            src_rel.target_part, src_rel.reltype)
                        blip.set(f"{{{Rn}}}embed", new_rId)
            spTree_dst.append(el)
        # No slide number on cover
        return

    # ── Content slide chrome ──────────────────────────────────────────
    if _LAYOUT_SOURCE is not None:
        _copy_layout_chrome(_LAYOUT_SOURCE, slide)
    elif _DECO_ELEMENTS:
        spTree = slide._element.find(
            ".//{http://schemas.openxmlformats.org/presentationml/2006/main}spTree")
        for el in _DECO_ELEMENTS:
            spTree.append(_copy.deepcopy(el))
    else:
        # Hard-coded fallback for bundled template (single-slide templates)
        rect = slide.shapes.add_shape(1, RECT_L, RECT_T, RECT_W, RECT_H)
        rect.fill.solid()
        rect.fill.fore_color.rgb = C_RECT
        rect.line.fill.background()
        line_xml = (
            f'''<p:cxnSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvCxnSpPr><p:cNvPr id="900" name="Line"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{int(LINE_L)}" y="{int(LINE_T)}"/><a:ext cx="{int(LINE_W)}" cy="0"/></a:xfrm>
    <a:prstGeom prst="line"><a:avLst/></a:prstGeom>
    <a:ln w="6350"><a:solidFill><a:srgbClr val="9E9E9F"/></a:solidFill></a:ln>
  </p:spPr>
</p:cxnSp>'''
        )
        spTree = slide._element.find(
            ".//{http://schemas.openxmlformats.org/presentationml/2006/main}spTree")
        if spTree is not None:
            spTree.append(etree.fromstring(line_xml))

    # Slide number — bottom-right on content slides
    tf = _new_tf(slide, SLIDE_NUM_L, SLIDE_NUM_T, SLIDE_NUM_W, SLIDE_NUM_H)
    p  = tf.add_paragraph()
    p.alignment = PP_ALIGN.RIGHT
    _no_bullet(p)
    r = p.add_run()
    r.text           = str(slide_num)
    r.font.name      = FONT_MAIN
    r.font.size      = Pt(10)
    r.font.color.rgb = C_SLIDE_NUM

# ═══════════════════════════════════════════════════════════════════════════════
# Title textbox
# ═══════════════════════════════════════════════════════════════════════════════

def _add_title(slide, title: str, cover=False):
    """
    Add the slide title.

    For cover slides with a real template cover (detected shape IDs):
      Write the title directly into the template text box so the font,
      colour, and position match the designer's intent exactly.

    For content slides (or cover slides without a dedicated template cover):
      Draw a new textbox at the standard title position.
    """
    if cover and _TEMPLATE_HAS_COVER and _COVER_TITLE_SHAPE_ID is not None:
        # Find the cloned title shape already on the slide and overwrite its text
        P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
        A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
        for sp in slide._element.findall(f".//{{{P_NS}}}sp"):
            cNvPr = sp.find(f".//{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr")
            if cNvPr is None:
                cNvPr = sp.find(
                    f"{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr")
            if cNvPr is not None and cNvPr.get("id") == _COVER_TITLE_SHAPE_ID:
                txBody = sp.find(f"{{{P_NS}}}txBody")
                if txBody is None: continue

                # Extract the rPr from the first <a:r> in the original paragraph
                # so we inherit the template's font, colour, size exactly.
                orig_rPr = txBody.find(f".//{{{A_NS}}}rPr")
                import copy as _copy_inner
                rPr_el = _copy_inner.deepcopy(orig_rPr) if orig_rPr is not None else None

                # Clear existing paragraphs
                for p_el in txBody.findall(f"{{{A_NS}}}p"):
                    txBody.remove(p_el)

                # Build new paragraph with original rPr
                p_el  = etree.SubElement(txBody, f"{{{A_NS}}}p")
                pPr   = etree.SubElement(p_el,  f"{{{A_NS}}}pPr")
                lnSpc = etree.SubElement(pPr,   f"{{{A_NS}}}lnSpc")
                etree.SubElement(lnSpc, f"{{{A_NS}}}spcPct", val="130000")
                r_el  = etree.SubElement(p_el,  f"{{{A_NS}}}r")
                if rPr_el is not None:
                    r_el.append(rPr_el)
                t_el  = etree.SubElement(r_el,  f"{{{A_NS}}}t")
                t_el.text = title
                return
        # Shape not found — fall through to draw a textbox
    elif cover and _TEMPLATE_HAS_COVER:
        return  # cover chrome already has a title box; skip

    # Standard textbox title (content slides + fallback cover)
    sz = 28 if cover else 24
    tf = _new_tf(slide, TITLE_L, TITLE_T, TITLE_W, TITLE_H)
    p  = tf.add_paragraph()
    if cover: p.alignment = PP_ALIGN.CENTER
    _no_bullet(p)
    _add_runs(p, title, sz, color=C_TITLE, bold_all=True)


# ═══════════════════════════════════════════════════════════════════════════════
# Cover content
# ═══════════════════════════════════════════════════════════════════════════════

def _render_cover(slide, segments, speaker_line: str = ""):
    """
    Render cover slide body content.

    If the template has a dedicated cover slide and a speaker shape ID,
    write the speaker/year text directly into that shape.
    Body content (paragraphs, headings, etc.) is rendered into the
    standard content area — useful for cover slides without a template
    cover (single-layout templates).
    """
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"

    # ── Write speaker line into template shape (if present) ───────────
    # Strategy: re-open the template, deep-copy the original <a:r> elements
    # from the speaker shape (which carry sz=1301, colour=#59E3FA, the correct
    # fonts), then replace only <a:t> text.  This preserves the designer's
    # exact style without hard-coding any values.
    if (_TEMPLATE_HAS_COVER and _COVER_SPEAKER_SHAPE_ID is not None
            and speaker_line):
        import copy as _copy_spk
        from pptx import Presentation as _PrsSpk
        _tprs_spk  = _PrsSpk(_TEMPLATE_PATH)
        _src_slide = _tprs_spk.slides[0]

        # Collect the three original <a:r> rPr elements from the template shape
        # (cyan "Speaker", white "|", cyan " year")
        orig_rPrs = []
        for _sp in _src_slide.shapes:
            if str(_sp.shape_id) == _COVER_SPEAKER_SHAPE_ID:
                for _r in _sp._element.findall(f".//{{{A_NS}}}r"):
                    rpr = _r.find(f"{{{A_NS}}}rPr")
                    orig_rPrs.append(_copy_spk.deepcopy(rpr) if rpr is not None else None)
                break
        # orig_rPrs[0] = cyan (left part), [1] = white "|", [2] = cyan (right part)

        for sp in slide._element.findall(f".//{{{P_NS}}}sp"):
            cNvPr = sp.find(f"{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr")
            if cNvPr is not None and cNvPr.get("id") == _COVER_SPEAKER_SHAPE_ID:
                txBody = sp.find(f"{{{P_NS}}}txBody")
                if txBody is None: break

                # Widen the text box so the speaker line fits on one line.
                # sz=1301 ≈ 10.01 pt; each character ≈ 10 pt wide on screen.
                # We expand leftward (keep right edge fixed) to avoid overflow.
                spPr  = sp.find(f"{{{P_NS}}}spPr")
                if spPr is None:
                    spPr = sp.find(f".//{{{A_NS}}}spPr")  # fallback
                xfrm  = spPr.find(f"{{{A_NS}}}xfrm") if spPr is not None else None
                if xfrm is not None:
                    off = xfrm.find(f"{{{A_NS}}}off")
                    ext = xfrm.find(f"{{{A_NS}}}ext")
                    if off is not None and ext is not None:
                        # Estimate needed width: ~9144 EMU per character at sz=1301
                        char_count = len(speaker_line) + 2   # +2 for padding
                        needed_w   = max(int(char_count * 9600), int(ext.get("cx", 0)))
                        right_edge = int(off.get("x", 0)) + int(ext.get("cx", 0))
                        new_x      = max(0, right_edge - needed_w)
                        off.set("x",  str(new_x))
                        ext.set("cx", str(right_edge - new_x))

                for p_el in txBody.findall(f"{{{A_NS}}}p"):
                    txBody.remove(p_el)

                # Build new paragraph preserving original pPr (algn=r, lnSpc)
                p_el  = etree.SubElement(txBody, f"{{{A_NS}}}p")
                pPr   = etree.SubElement(p_el, f"{{{A_NS}}}pPr")
                pPr.set("algn", "r")
                lnSpc = etree.SubElement(pPr, f"{{{A_NS}}}lnSpc")
                etree.SubElement(lnSpc, f"{{{A_NS}}}spcPts", val="3758")

                def _make_run(text, rpr_template):
                    r_el = etree.SubElement(p_el, f"{{{A_NS}}}r")
                    if rpr_template is not None:
                        r_el.append(_copy_spk.deepcopy(rpr_template))
                    t_el = etree.SubElement(r_el, f"{{{A_NS}}}t")
                    t_el.text = text

                if "|" in speaker_line:
                    left, right = speaker_line.split("|", 1)
                    rpr_cyan  = orig_rPrs[0] if len(orig_rPrs) > 0 else None
                    rpr_white = orig_rPrs[1] if len(orig_rPrs) > 1 else None
                    _make_run(left.strip() + " ", rpr_cyan)
                    _make_run("|",               rpr_white)
                    _make_run(" " + right.strip(), rpr_cyan)
                else:
                    rpr_cyan = orig_rPrs[0] if orig_rPrs else None
                    _make_run(speaker_line, rpr_cyan)
                break

    # ── Render body content (for single-layout templates or extra text) ─
    if _TEMPLATE_HAS_COVER:
        return   # cover template manages its own layout; don't add extra boxes

    all_blocks = []
    for st, sd in segments:
        if st == "text": all_blocks.extend(sd)
    if not all_blocks: return

    tf = _new_tf(slide, CONTENT_L, CONTENT_T, CONTENT_W, CONTENT_H)
    for block in all_blocks:
        bt = block.get("type")
        if bt in ("blank", "hr"):
            _spacer_para(tf, 6)
        elif bt == "heading":
            p = tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            _no_bullet(p)
            col = {3: C_H3, 4: C_H4, 5: C_H5}.get(block["level"], C_H3)
            _add_runs(p, block["text"], SZ_H4, color=col, bold_all=True)
        else:
            p = tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            _no_bullet(p)
            _add_runs(p, block.get("text", ""), 16)


# ═══════════════════════════════════════════════════════════════════════════════
# Body segments with y-cursor (from micknoise approach)
# ═══════════════════════════════════════════════════════════════════════════════

def _render_body(slide, segments, body_lines):
    n_lines = sum(1 for l in body_lines
                  if l.strip() and not l.strip().startswith("|")
                  and not l.strip().startswith("```"))
    sz = auto_body_size(n_lines)

    y    = CONTENT_T
    gap  = Pt(10)

    for seg_type, seg_data in segments:
        if y >= CONTENT_B - Pt(20): break
        remaining = CONTENT_B - y

        if seg_type == "table":
            bottom = _add_pptx_table(slide, seg_data, CONTENT_L, y, CONTENT_W)
            y = bottom + gap

        elif seg_type == "code":
            bottom = _add_code_block(slide, seg_data,
                                     CONTENT_L + Inches(0.15), y,
                                     CONTENT_W - Inches(0.15))
            y = bottom + gap

        elif seg_type == "quote":
            bottom = _add_quote_block(slide, seg_data,
                                      CONTENT_L + Inches(0.05), y,
                                      CONTENT_W - Inches(0.05), sz)
            y = bottom + gap

        elif seg_type == "text":
            if all(b.get("type") == "blank" for b in seg_data): continue
            tf = _new_tf(slide, CONTENT_L, int(y), CONTENT_W, int(remaining))
            _render_text_blocks(tf, seg_data, sz)
            # Approximate height to advance cursor
            n_c = sum(1 for b in seg_data if b.get("type") != "blank")
            n_b = sum(1 for b in seg_data if b.get("type") == "blank")
            y   = int(y) + int(Pt(sz + 6) * n_c + Pt(4) * n_b + gap)


# ═══════════════════════════════════════════════════════════════════════════════
# Speaker notes
# ═══════════════════════════════════════════════════════════════════════════════

def _set_notes(slide, note_text: str):
    tf  = slide.notes_slide.notes_text_frame
    tf.clear()
    p   = tf.paragraphs[0]
    run = p.add_run()
    run.text           = note_text
    run.font.name      = FONT_MAIN
    run.font.size      = Pt(11)
    run.font.color.rgb = C_BODY


# ═══════════════════════════════════════════════════════════════════════════════
# Main conversion
# ═══════════════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════════════
# Template layout detection
# ═══════════════════════════════════════════════════════════════════════════════

def load_layout(template_path: str, config_path: str = None):
    """
    Auto-detect layout dimensions from any PPTX template and populate the
    global layout variables.  Override order (highest priority first):

      1. JSON config file (explicit overrides — most flexible)
      2. Auto-detection from the template slides' shapes
      3. Module-level defaults (bundled template dimensions)

    If the template contains 2+ slides, slide[0] is treated as a dedicated
    cover slide and slide[1] as the content slide.  Their decorative elements
    are stored separately so each type of output slide gets the right chrome.

    JSON config format (save as e.g. my_layout.json):
    {
      "slide_w": 13.333,    "slide_h": 7.5,
      "title_l": 0.5,       "title_t": 0.3,
      "title_w": 12.0,      "title_h": 0.8,
      "content_l": 0.5,     "content_t": 1.2,
      "content_w": 12.0,    "content_h": 5.5,
      "slide_num_l": 11.5,  "slide_num_t": 7.1,
      "slide_num_w": 1.5,   "slide_num_h": 0.3,
      "copy_deco": true,
      "font_main": "Microsoft YaHei",
      "title_color": "00B0F0",
      "h3_color": "005A87"
    }
    All fields are optional — omit any you want to keep at their detected value.
    """
    global SLIDE_W, SLIDE_H
    global TITLE_L, TITLE_T, TITLE_W, TITLE_H
    global CONTENT_L, CONTENT_T, CONTENT_W, CONTENT_H, CONTENT_B
    global SLIDE_NUM_L, SLIDE_NUM_T, SLIDE_NUM_W, SLIDE_NUM_H
    global _DECO_ELEMENTS, _COVER_DECO_ELEMENTS
    global _LAYOUT_SOURCE, _COVER_LAYOUT_SOURCE
    global _COVER_TITLE_SHAPE_ID, _COVER_SPEAKER_SHAPE_ID, _TEMPLATE_HAS_COVER
    global _TEMPLATE_PATH
    global FONT_MAIN, C_TITLE, C_H3

    import copy as _copy
    import json

    _TEMPLATE_PATH = template_path
    prs = Presentation(template_path)

    # ── Step 1: read slide size from template ─────────────────────────
    SLIDE_W = prs.slide_width
    SLIDE_H = prs.slide_height

    def _in(emu): return emu / 914400   # EMU → inches

    # ── Step 2: detect whether template has a dedicated cover slide ───
    # Heuristic: if slide[0] has no shape named "ContentBox" and no
    # obvious tall content area, it's a cover; slide[1] is content.
    n_slides = len(prs.slides)

    def _looks_like_cover(slide):
        """True if slide appears to be a cover (no tall content area WITH text)."""
        for shape in slide.shapes:
            if not hasattr(shape, "text"): continue
            if not shape.text.strip(): continue   # ignore empty shapes
            h = _in(shape.height)
            t = _in(shape.top)
            w = _in(shape.width)
            # A content area: tall, wide, not near top, and has actual text
            # that looks like body content (more than a title's worth)
            txt = shape.text.strip()
            if h > 1.5 and w > 4 and t >= 0.8 and len(txt) > 40:
                return False   # has a populated content box → not a cover
        return True

    if n_slides >= 2 and _looks_like_cover(prs.slides[0]):
        cover_slide   = prs.slides[0]
        content_slide = prs.slides[1]
        _TEMPLATE_HAS_COVER = True
        print("  Template: cover slide detected (slide 1) + content slide (slide 2)")
    else:
        cover_slide   = None
        content_slide = prs.slides[0]
        _TEMPLATE_HAS_COVER = False
        print("  Template: single-layout (no dedicated cover slide)")

    # ── Step 3: detect cover text box IDs ────────────────────────────
    # Look for the two named text boxes on the cover slide:
    #   largest text box  → main title (white bold)
    #   smaller text box with '|' in its text → speaker | year
    if cover_slide is not None:
        _COVER_LAYOUT_SOURCE = cover_slide.slide_layout
        best_title_area   = -1
        best_speaker_area = -1
        for shape in cover_slide.shapes:
            if not hasattr(shape, "text"): continue
            area = _in(shape.width) * _in(shape.height)
            txt  = shape.text.strip()
            sid  = str(shape.shape_id)
            # Speaker box: contains '|' or keywords like 'Speaker'
            if "|" in txt or "Speaker" in txt or "year" in txt.lower():
                if area > best_speaker_area:
                    best_speaker_area        = area
                    _COVER_SPEAKER_SHAPE_ID  = sid
            # Title box: largest text area on the slide (excluding speaker)
            elif txt and area > best_title_area:
                best_title_area         = area
                _COVER_TITLE_SHAPE_ID   = sid
        print(f"  Cover shape IDs: title={_COVER_TITLE_SHAPE_ID}  "
              f"speaker={_COVER_SPEAKER_SHAPE_ID}")
    else:
        _COVER_LAYOUT_SOURCE        = None
        _COVER_TITLE_SHAPE_ID       = None
        _COVER_SPEAKER_SHAPE_ID     = None

    # ── Step 4: auto-detect title and content boxes from content slide ─
    title_shape   = None
    content_shape = None
    deco_shapes   = []

    for shape in content_slide.shapes:
        has_text = hasattr(shape, "text")
        t = _in(shape.top);  h = _in(shape.height);  w = _in(shape.width)

        if has_text and t < 1.5 and h < 1.5 and w > 4:
            if title_shape is None or t < _in(title_shape.top):
                title_shape = shape
            continue
        if has_text and t >= 0.8 and h > 1.5 and w > 4:
            if content_shape is None or h > _in(content_shape.height):
                content_shape = shape
            continue
        deco_shapes.append(shape)

    if title_shape:
        TITLE_L = title_shape.left;  TITLE_T = title_shape.top
        TITLE_W = title_shape.width; TITLE_H = title_shape.height

    if content_shape:
        CONTENT_L = content_shape.left;  CONTENT_T = content_shape.top
        CONTENT_W = content_shape.width; CONTENT_H = content_shape.height
        CONTENT_B = CONTENT_T + CONTENT_H

    for shape in deco_shapes:
        t = _in(shape.top)
        if t > _in(SLIDE_H) * 0.8:
            SLIDE_NUM_L = shape.left;  SLIDE_NUM_T = shape.top
            SLIDE_NUM_W = shape.width; SLIDE_NUM_H = shape.height
            break

    # ── Step 5: store layout references and deco element lists ────────
    _LAYOUT_SOURCE = content_slide.slide_layout

    _DECO_ELEMENTS = []
    for shape in deco_shapes:
        t = _in(shape.top)
        if t > _in(SLIDE_H) * 0.8 and hasattr(shape, "text"):
            continue
        _DECO_ELEMENTS.append(_copy.deepcopy(shape._element))

    # Cover deco: copy all shapes from cover slide (they ARE the cover chrome)
    _COVER_DECO_ELEMENTS = []
    if cover_slide is not None:
        for shape in cover_slide.shapes:
            _COVER_DECO_ELEMENTS.append(_copy.deepcopy(shape._element))

    # ── Step 6: JSON config overrides ────────────────────────────────
    cfg = {}
    if config_path and Path(config_path).exists():
        cfg_file = config_path
    else:
        cfg_file = Path(template_path).with_suffix(".json")
        if not cfg_file.exists():
            cfg_file = None

    if cfg_file:
        try:
            cfg = json.loads(Path(cfg_file).read_text(encoding="utf-8"))
            print(f"  Config loaded from {cfg_file}")
        except Exception as e:
            print(f"  Warning: could not read config {cfg_file}: {e}")

    def _i(key, default):
        return Inches(cfg[key]) if key in cfg else default

    SLIDE_W     = _i("slide_w",     SLIDE_W)
    SLIDE_H     = _i("slide_h",     SLIDE_H)
    TITLE_L     = _i("title_l",     TITLE_L)
    TITLE_T     = _i("title_t",     TITLE_T)
    TITLE_W     = _i("title_w",     TITLE_W)
    TITLE_H     = _i("title_h",     TITLE_H)
    CONTENT_L   = _i("content_l",   CONTENT_L)
    CONTENT_T   = _i("content_t",   CONTENT_T)
    CONTENT_W   = _i("content_w",   CONTENT_W)
    CONTENT_H   = _i("content_h",   CONTENT_H)
    CONTENT_B   = CONTENT_T + CONTENT_H
    SLIDE_NUM_L = _i("slide_num_l", SLIDE_NUM_L)
    SLIDE_NUM_T = _i("slide_num_t", SLIDE_NUM_T)
    SLIDE_NUM_W = _i("slide_num_w", SLIDE_NUM_W)
    SLIDE_NUM_H = _i("slide_num_h", SLIDE_NUM_H)

    if "font_main"    in cfg: FONT_MAIN = cfg["font_main"]
    if "title_color"  in cfg:
        v = cfg["title_color"].lstrip("#")
        C_TITLE = RGBColor(int(v[0:2],16), int(v[2:4],16), int(v[4:6],16))
    if "h3_color"     in cfg:
        v = cfg["h3_color"].lstrip("#")
        C_H3 = RGBColor(int(v[0:2],16), int(v[2:4],16), int(v[4:6],16))
    if cfg.get("copy_deco") is False:
        _DECO_ELEMENTS = []

    def _fmt(emu): return f'{emu/914400:.3f}"'
    print(f"  Layout: slide={_fmt(SLIDE_W)}×{_fmt(SLIDE_H)}"
          f"  title=({_fmt(TITLE_L)},{_fmt(TITLE_T)}) {_fmt(TITLE_W)}×{_fmt(TITLE_H)}"
          f"  content=({_fmt(CONTENT_L)},{_fmt(CONTENT_T)}) {_fmt(CONTENT_W)}×{_fmt(CONTENT_H)}"
          f"  deco_shapes={len(_DECO_ELEMENTS)}")


# ═══════════════════════════════════════════════════════════════════════════════
# Main conversion
# ═══════════════════════════════════════════════════════════════════════════════

def convert(md_path: str, template_path: str, output_path: str,
            config_path: str = None):
    import chardet

    # Auto-detect layout from the supplied template
    print(f"Loading template: {template_path}")
    load_layout(template_path, config_path=config_path)

    raw      = Path(md_path).read_bytes()
    encoding = chardet.detect(raw).get("encoding") or "utf-8"
    md_text  = raw.decode(encoding)

    slides_data = parse_md(md_text)
    if not slides_data:
        print("ERROR: No slides found.", file=sys.stderr)
        sys.exit(1)

    print(f"Parsed {len(slides_data)} slide(s).")

    # Build output from a blank presentation (avoids XML-cloning issues)
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H
    blank_layout = prs.slide_layouts[6]   # completely blank

    for idx, sd in enumerate(slides_data):
        slide    = prs.slides.add_slide(blank_layout)
        cover    = sd.get("cover", False)
        speaker  = sd.get("speaker", "")
        segments = classify_body(sd["body_lines"])

        _add_decorative(slide, idx + 1, cover=cover)
        _add_title(slide, sd["title"], cover=cover)

        if cover:
            _render_cover(slide, segments, speaker_line=speaker)
        else:
            _render_body(slide, segments, sd["body_lines"])

        if sd["note"]:
            _set_notes(slide, sd["note"])

        print(f"  Slide {idx+1}: {sd['title'][:60]}")

    prs.save(output_path)
    print(f"\nSaved → {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Convert full-featured Markdown to PPTX using any template.")
    parser.add_argument("md",       help="Input Markdown file")
    parser.add_argument("template", help="PPTX template file")
    parser.add_argument("output",   help="Output PPTX file")
    parser.add_argument("--config", "-c", default=None,
                        help="Optional JSON layout config (default: <template>.json if exists)")
    args = parser.parse_args()
    convert(args.md, args.template, args.output, config_path=args.config)


if __name__ == "__main__":
    main()
