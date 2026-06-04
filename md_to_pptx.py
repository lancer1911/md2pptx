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
  ✓ Source Han Sans SC (Noto Sans CJK SC) font throughout — SIL OFL, free for commercial use
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
from dataclasses import dataclass, field
from typing import Optional

# ── LayoutConfig dataclass ────────────────────────────────────────────────────
# All layout state is held in one instance passed through the call chain.
# This eliminates mutable module globals, makes the code thread-safe, and
# allows easy unit-testing of individual functions.

@dataclass
class LayoutConfig:
    # Slide dimensions
    slide_w:   int = field(default_factory=lambda: Inches(11.42))
    slide_h:   int = field(default_factory=lambda: Inches(6.42))
    # Title box
    title_l:   int = field(default_factory=lambda: Inches(0.473))
    title_t:   int = field(default_factory=lambda: Inches(0.202))
    title_w:   int = field(default_factory=lambda: Inches(9.482))
    title_h:   int = field(default_factory=lambda: Inches(0.637))
    # Content box
    content_l: int = field(default_factory=lambda: Inches(0.473))
    content_t: int = field(default_factory=lambda: Inches(1.354))
    content_w: int = field(default_factory=lambda: Inches(10.662))
    content_h: int = field(default_factory=lambda: Inches(3.466))
    # Slide number box
    slide_num_l: int = field(default_factory=lambda: Inches(8.672))
    slide_num_t: int = field(default_factory=lambda: Inches(5.951))
    slide_num_w: int = field(default_factory=lambda: Inches(2.665))
    slide_num_h: int = field(default_factory=lambda: Inches(0.342))
    # Decorative fallback dimensions (bundled template)
    rect_l: int = field(default_factory=lambda: Inches(0.519))
    rect_t: int = field(default_factory=lambda: Inches(0.947))
    rect_w: int = field(default_factory=lambda: Inches(0.717))
    rect_h: int = field(default_factory=lambda: Inches(0.086))
    line_l: int = field(default_factory=lambda: Inches(0.519))
    line_t: int = field(default_factory=lambda: Inches(0.947))
    line_w: int = field(default_factory=lambda: Inches(4.863))
    # Fonts and colours (overridable via JSON)
    font_main:   str = "Source Han Sans SC"
    title_color: RGBColor = field(default_factory=lambda: RGBColor(0x00, 0xB0, 0xF0))
    h3_color:    RGBColor = field(default_factory=lambda: RGBColor(0x00, 0x5A, 0x87))
    # Template metadata (populated by load_layout)
    template_path:          Optional[str]    = None
    template_has_cover:     bool             = False
    cover_title_shape_id:   Optional[str]    = None
    cover_speaker_shape_id: Optional[str]    = None
    layout_source:          object           = None   # python-pptx SlideLayout
    cover_layout_source:    object           = None
    deco_elements:          list             = field(default_factory=list)
    cover_deco_elements:    list             = field(default_factory=list)

    @property
    def content_b(self) -> int:
        return self.content_t + self.content_h

# ── Design tokens ─────────────────────────────────────────────────────────────
# Colour and size constants that never change between templates.
# Per-template overrides (font_main, title_color, h3_color) live in LayoutConfig.
FONT_MONO = "Courier New"

C_H4         = RGBColor(0x1A, 0x7D, 0xB5)
C_H5         = RGBColor(0x6B, 0x72, 0x80)
C_BODY       = RGBColor(0x1F, 0x29, 0x37)
C_QUOTE      = RGBColor(0x1F, 0x29, 0x37)
C_QUOTE_BAR  = RGBColor(0xC8, 0xCD, 0xD4)
C_CODE_BG    = RGBColor(0xF3, 0xF4, 0xF6)   # light grey code-block background
C_CODE_FG    = RGBColor(0x1F, 0x29, 0x37)
C_LINK       = RGBColor(0x25, 0x63, 0xEB)
C_DIVIDER    = RGBColor(0xCB, 0xD5, 0xE1)
C_RECT       = RGBColor(0x76, 0xCA, 0xF2)   # template accent rect colour
C_LINE       = RGBColor(0x9E, 0x9E, 0x9F)   # template line colour
C_SLIDE_NUM  = RGBColor(0xAA, 0xAA, 0xAA)

# Table colours
C_TABLE_HDR     = RGBColor(0x1A, 0x1A, 0x2E)
C_TABLE_HDR_TXT = RGBColor(0xFF, 0xFF, 0xFF)
C_TABLE_EVEN    = RGBColor(0xEF, 0xF6, 0xFB)
C_TABLE_ODD     = RGBColor(0xFF, 0xFF, 0xFF)

# SmartArt colour palettes  (index 0 = darkest / most prominent)
_SA_PALETTES = {
    "blue":  [RGBColor(0x00, 0x5A, 0x87), RGBColor(0x00, 0x7B, 0xB5),
              RGBColor(0x00, 0xB0, 0xF0), RGBColor(0x9D, 0xD9, 0xF5),
              RGBColor(0xCC, 0xEE, 0xFA), RGBColor(0xE6, 0xF6, 0xFD)],
    "teal":  [RGBColor(0x02, 0x80, 0x90), RGBColor(0x00, 0xA8, 0x96),
              RGBColor(0x02, 0xC3, 0x9A), RGBColor(0x9D, 0xE0, 0xD8),
              RGBColor(0xCC, 0xF0, 0xEC), RGBColor(0xE5, 0xF8, 0xF6)],
    "green": [RGBColor(0x2C, 0x5F, 0x2D), RGBColor(0x5A, 0x96, 0x47),
              RGBColor(0x97, 0xBC, 0x62), RGBColor(0xC5, 0xDC, 0xA0),
              RGBColor(0xE2, 0xEF, 0xCF), RGBColor(0xF0, 0xF7, 0xE7)],
    "orange":[RGBColor(0xB8, 0x50, 0x42), RGBColor(0xE0, 0x7B, 0x39),
              RGBColor(0xF9, 0xA8, 0x25), RGBColor(0xFC, 0xD3, 0x91),
              RGBColor(0xFE, 0xEC, 0xC8), RGBColor(0xFF, 0xF6, 0xE4)],
    "purple":[RGBColor(0x4A, 0x14, 0x86), RGBColor(0x7B, 0x2F, 0xBE),
              RGBColor(0xAB, 0x6D, 0xE0), RGBColor(0xD3, 0xB3, 0xF0),
              RGBColor(0xEC, 0xD9, 0xF8), RGBColor(0xF6, 0xEC, 0xFC)],
    "mono":  [RGBColor(0x1F, 0x29, 0x37), RGBColor(0x37, 0x41, 0x51),
              RGBColor(0x6B, 0x72, 0x80), RGBColor(0x9C, 0xA3, 0xAF),
              RGBColor(0xD1, 0xD5, 0xDB), RGBColor(0xF3, 0xF4, 0xF6)],
}
_SA_DEFAULT_PALETTE = "blue"

SZ_H3   = 15
SZ_H4   = 13
SZ_H5   = 12
SZ_SUB  = 12
SZ_QUOT = 13
SZ_CODE = 11

TABLE_ROW_H = Inches(0.40)

def _is_cjk_dominant(body_lines: list) -> bool:
    """Return True when the majority of non-space characters are CJK."""
    cjk = latin = 0
    for line in body_lines:
        for ch in line:
            if '\u4e00' <= ch <= '\u9fff' or '\u3000' <= ch <= '\u303f':
                cjk += 1
            elif ch.isalpha():
                latin += 1
    return cjk >= latin   # CJK or mixed → True; Latin-dominant → False


def auto_body_size(n_lines: int, cjk: bool = True) -> int:
    """Auto-scale body font based on content density.

    English text needs a smaller font than Chinese at the same logical
    line count because English words are wider relative to the slide width.
    A Latin line at 14pt fits ~80 chars; the same line in Chinese fits
    ~40 chars.  We apply a 1.6× effective-line multiplier for Latin text.
    """
    effective = int(n_lines * (1.0 if cjk else 1.6))
    if effective <= 6:  return 18
    if effective <= 10: return 16
    if effective <= 14: return 14
    return 13


# ── Template cache (fix for issue #3: avoid re-opening the file per cover slide) ─
_TEMPLATE_CACHE: dict = {}   # path → Presentation

def _get_cached_template(lc: LayoutConfig):
    """Return a cached Presentation for lc.template_path, opening it once."""
    from pptx import Presentation as _Prs
    path = lc.template_path
    if path not in _TEMPLATE_CACHE:
        _TEMPLATE_CACHE[path] = _Prs(path)
    return _TEMPLATE_CACHE[path]


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

    in_quote = False   # track whether we are inside a blockquote run

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
            in_quote = False
            continue
        if in_code:
            code_buf.append(raw); continue

        # Track blockquote context so that "---" inside a quote is NOT
        # treated as a slide separator (fix for edge-case #8).
        if line.startswith("> ") or line == ">":
            in_quote = True
        elif line:
            in_quote = False

        # ── Explicit slide separator "---" ───────────────────────────
        # Only treat a lone --- as a slide break when we are NOT inside
        # a blockquote run.  This prevents "--- " acting as a divider
        # when it appears between consecutive "> " lines.
        if re.match(r'^-{3,}$', line) and current is not None and not in_quote:
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
        # Capture speaker/year line: plain text after a cover slide start.
        # Tolerate one blank line between # and the speaker line.
        if (current is not None and current["cover"]
                and not current.get("speaker")   # speaker not yet set
                and (not current["body_lines"]    # nothing in body yet
                     or current["body_lines"] == [''])  # only a blank so far
                and line
                and not line.startswith("#")
                and not line.startswith("-")
                and not line.startswith("*")
                and not line.startswith("|")
                and not line.startswith(">")
                and not line.startswith("---")):
            current["speaker"] = line
            current["body_lines"] = []  # discard the blank
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
            cur_text.append({"type": "quote", "text": "\n".join(cur_quote)})
            cur_quote = None

    _IMG_LINE = re.compile(r'^!\[([^\]]*)\]\(([^)]*)\)$')   # image on its own line

    # ── SmartArt accumulator ──────────────────────────────────────────
    in_smartart  = False
    cur_smartart = None   # dict with type/style/theme/items

    for raw in body_lines:
        line = raw.strip()

        # ── SmartArt fence open/close ─────────────────────────────────
        if not in_smartart and line.startswith(":::smartart"):
            flush_text(); flush_table()
            in_smartart = True
            attrs = {}
            for m in re.finditer(r'(\w+)\s*=\s*"([^"]*)"', line):
                attrs[m.group(1)] = m.group(2)
            cur_smartart = {
                "type":  attrs.get("type",  "process"),
                "style": attrs.get("style", ""),
                "theme": attrs.get("theme", ""),
                "size":  attrs.get("size",  ""),
                "items": [],
            }
            continue
        if in_smartart:
            if line == ":::":
                in_smartart = False
                segments.append(("smartart", cur_smartart))
                cur_smartart = None
            else:
                bm = re.match(r'^( *)[-*+] (.+)$', raw)
                if bm:
                    indent = len(bm.group(1)) // 2
                    cur_smartart["items"].append({"text": bm.group(2).strip(), "level": indent})
            continue

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
        # Contiguous "> " lines are merged into a single block with one
        # vertical bar.  This matches how Markdown editors render multi-
        # line block quotes and makes the PPTX easier to read.
        # A bare ">" (no trailing space) is also treated as a quote line.
        if line.startswith("> ") or line == ">":
            flush_table()
            if cur_quote is None: cur_quote = []
            cur_quote.append(line[2:].strip() if line.startswith("> ") else "")
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

        # ── Image on its own line → dedicated segment ────────────────
        # Images are promoted out of the text flow so _render_body can
        # embed them as real picture shapes (not placeholder captions).
        mi = _IMG_LINE.match(line)
        if mi:
            flush_text(); flush_table()
            segments.append(("image", {"alt": mi.group(1), "src": mi.group(2)}))
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
            # Reset counter when the author restarts at "1." (new list begins).
            # This prevents a second list on the same slide continuing from
            # the number where the first list stopped (fix for issue #9).
            if int(m.group(2)) == 1:
                ord_cnt[ind] = 0
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

    flush_table(); flush_code(); flush_quote(); flush_text()
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

def _apply_run(run, text, sz_pt, lc: LayoutConfig, bold=False, italic=False,
               strike=False, code=False, color: RGBColor = None, url=None):
    run.text = text
    run.font.name      = FONT_MONO if code else lc.font_main
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


def _add_runs(para, text: str, sz_pt, lc: LayoutConfig,
              color=None, bold_all=False, italic_all=False):
    for (t, b, i, s, code, url) in parse_inline(text):
        r = para.add_run()
        c = C_LINK if url else (C_CODE_FG if code else (color or C_BODY))
        _apply_run(r, t, sz_pt, lc,
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


def _heading_para(tf, text, level, lc: LayoutConfig, sz_override=None):
    col = {3: lc.h3_color, 4: C_H4, 5: C_H5}.get(level, lc.h3_color)
    sz  = sz_override or {3: SZ_H3, 4: SZ_H4, 5: SZ_H5}.get(level, SZ_H3)
    p   = tf.add_paragraph()
    p.space_before = Pt(6)
    _no_bullet(p)
    _add_runs(p, text, sz, lc, col, bold_all=(level == 3))
    return p


def _bullet_para(tf, text, indent, sz, lc: LayoutConfig):
    MAR0, MAR1, IND = 342900, 685800, -342900
    p   = tf.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(MAR1 if indent else MAR0))
    pPr.set("indent", str(IND))
    for el in pPr.findall(qn("a:buNone")): pPr.remove(el)
    buChar = etree.SubElement(pPr, qn("a:buChar"))
    buChar.set("char", "◦" if indent else "•")
    buFont = etree.SubElement(pPr, qn("a:buFont"))
    buFont.set("typeface", lc.font_main)
    p.space_before = Pt(3)
    _add_runs(p, text, sz, lc)
    return p


def _ordered_para(tf, text, indent, number, sz, lc: LayoutConfig):
    # Use explicit "N.\t" prefix rather than buAutoNum so that numbering
    # resets correctly at each new section (buAutoNum is slide-scoped in pptx).
    MAR0, MAR1 = 342900, 685800
    HANG        = -285750   # ~0.31" hanging indent
    p   = tf.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(MAR1 if indent else MAR0))
    pPr.set("indent", str(HANG))
    # Suppress any automatic bullet
    for el in pPr.findall(qn("a:buNone")): pPr.remove(el)
    etree.SubElement(pPr, qn("a:buNone"))
    p.space_before = Pt(3)
    # Number run
    nr   = p.add_run()
    nr.text      = f"{number}.  "
    nr.font.name = lc.font_main
    nr.font.size = Pt(sz)
    # Content run
    _add_runs(p, text, sz, lc)
    return p


def _quote_para(tf, text, sz, lc: LayoutConfig):
    # Fallback for old text-box rendering paths. The main renderer uses
    # _add_quote_block(), which draws a real vertical bar shape.
    p   = tf.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(274638))
    pPr.set("indent", str(0))
    _no_bullet(p)
    p.space_before = Pt(4)
    p.space_after  = Pt(4)
    _add_runs(p, text, sz, lc, color=C_QUOTE)
    return p


def _estimate_quote_height(text: str, width_emu, sz_pt: int) -> int:
    """Approximate quote box height for y-cursor layout."""
    width_in = max(width_emu / 914400, 1.0)
    # CJK chars are ~1 em wide; Latin chars average ~0.5 em.
    # Detect language from the quote text itself.
    cjk_ch = sum(1 for ch in text if '\u4e00' <= ch <= '\u9fff')
    lat_ch = sum(1 for ch in text if ch.isalpha() and ord(ch) < 0x300)
    is_cjk = cjk_ch >= lat_ch
    # CJK: ~54 chars per line at 16pt across 9.5"; Latin: ~95 chars
    density = 54 if is_cjk else 95
    chars_per_line = max(int(width_in * (density / 9.5) * (16 / max(sz_pt, 1))), 18)
    logical_lines = 0
    for part in (text or " ").splitlines():
        logical_lines += max(1, (len(part) + chars_per_line - 1) // chars_per_line)
    return int(Pt(sz_pt + 4) * logical_lines + Pt(6))


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
        _add_runs(p, line, sz_pt + 2, lc, color=C_QUOTE)

    return int(top) + height


def _plain_para(tf, text, sz, lc: LayoutConfig):
    p = tf.add_paragraph()
    _no_bullet(p)
    _add_runs(p, text, sz, lc)
    return p


def _spacer_para(tf, sz_pt=4):
    p = tf.add_paragraph()
    _no_bullet(p)
    r = p.add_run()
    r.font.size = Pt(sz_pt)
    return p


def _render_text_blocks(tf, blocks: list, sz_pt: float, lc: LayoutConfig):
    last = None
    for block in blocks:
        bt = block.get("type")
        if bt == "blank":
            if last not in (None, "blank"): _spacer_para(tf, 4)
            last = bt; continue
        if bt == "heading":
            _heading_para(tf, block["text"], block["level"], lc)
        elif bt == "bullet":
            _bullet_para(tf, block["text"], block["indent"], sz_pt, lc)
        elif bt == "ordered":
            _ordered_para(tf, block["text"], block["indent"], block["number"], sz_pt, lc)
        elif bt == "quote":
            _quote_para(tf, block["text"], sz_pt, lc)
        elif bt == "para":
            _plain_para(tf, block["text"], sz_pt, lc)
        elif bt == "image":
            # Images promoted to top-level segments are handled by _render_body;
            # inline image syntax (e.g. inside a paragraph) falls through here.
            alt = block.get("alt") or block.get("src", "")
            p   = tf.add_paragraph()
            _no_bullet(p)
            r = p.add_run()
            r.text           = f"[图片: {alt}]" if alt else "[图片]"
            r.font.name      = lc.font_main
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

def _add_pptx_table(slide, table_rows: list, left, top, width, lc: LayoutConfig) -> float:
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
                _apply_run(run, t, 13, lc, bold=is_hdr or b, italic=i,
                           strike=s, code=code, color=tc, url=url)

    return int(top) + height


# ═══════════════════════════════════════════════════════════════════════════════
# Code block textbox
# ═══════════════════════════════════════════════════════════════════════════════

def _add_code_block(slide, code_lines: list, left, top, width) -> float:
    n      = len(code_lines) or 1
    pad    = int(Pt(6))
    height = int(Pt(SZ_CODE + 4) * n + pad * 2)

    # Grey background rectangle drawn behind the text box
    bg = slide.shapes.add_shape(1, int(left), int(top), int(width), height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_CODE_BG
    bg.line.fill.background()

    tf = _new_tf(slide, int(left) + pad, int(top) + pad,
                 int(width) - pad * 2, height - pad * 2)
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


def _add_decorative(slide, slide_num: int, lc: LayoutConfig, cover: bool = False):
    """
    Stamp template chrome onto a slide.

    For cover slides: copy all shapes from the template cover slide verbatim
    (background image, coloured rectangles, etc.) then write title/speaker
    into the detected text box IDs — no separate title textbox is drawn.

    For content slides: copy non-placeholder shapes from the content slide
    layout, then add a slide number.

    The template is opened at most once per conversion run by using a
    module-level cache (_CACHED_TEMPLATE_PRS) so we avoid re-reading the
    file for every cover slide (fix for issue #3).
    """
    import copy as _copy

    if cover and lc.template_has_cover:
        # Clone every shape from the template cover slide into this slide.
        # Use the cached Presentation to avoid re-opening the file.
        P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
        A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
        Rn   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        spTree_dst = slide._element.find(f".//{{{P_NS}}}spTree")

        src_slide = _get_cached_template(lc).slides[0]

        for shape in src_slide.shapes:
            el     = _copy.deepcopy(shape._element)
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
    if lc.layout_source is not None:
        _copy_layout_chrome(lc.layout_source, slide)
    elif lc.deco_elements:
        spTree = slide._element.find(
            ".//{http://schemas.openxmlformats.org/presentationml/2006/main}spTree")
        for el in lc.deco_elements:
            spTree.append(_copy.deepcopy(el))
    else:
        # Hard-coded fallback for bundled template (single-slide templates)
        rect = slide.shapes.add_shape(1, lc.rect_l, lc.rect_t, lc.rect_w, lc.rect_h)
        rect.fill.solid()
        rect.fill.fore_color.rgb = C_RECT
        rect.line.fill.background()
        line_xml = (
            f'''<p:cxnSp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvCxnSpPr><p:cNvPr id="900" name="Line"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{int(lc.line_l)}" y="{int(lc.line_t)}"/><a:ext cx="{int(lc.line_w)}" cy="0"/></a:xfrm>
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
    tf = _new_tf(slide, lc.slide_num_l, lc.slide_num_t, lc.slide_num_w, lc.slide_num_h)
    p  = tf.add_paragraph()
    p.alignment = PP_ALIGN.RIGHT
    _no_bullet(p)
    r = p.add_run()
    r.text           = str(slide_num)
    r.font.name      = lc.font_main
    r.font.size      = Pt(10)
    r.font.color.rgb = C_SLIDE_NUM

# ═══════════════════════════════════════════════════════════════════════════════
# Title textbox
# ═══════════════════════════════════════════════════════════════════════════════

def _add_title(slide, title: str, lc: LayoutConfig, cover=False):
    """
    Add the slide title.

    For cover slides with a real template cover (detected shape IDs):
      Write the title directly into the template text box so the font,
      colour, and position match the designer's intent exactly.

    For content slides (or cover slides without a dedicated template cover):
      Draw a new textbox at the standard title position.
    """
    if cover and lc.template_has_cover and lc.cover_title_shape_id is not None:
        # Find the cloned title shape already on the slide and overwrite its text
        P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
        A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
        for sp in slide._element.findall(f".//{{{P_NS}}}sp"):
            cNvPr = sp.find(f".//{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr")
            if cNvPr is None:
                cNvPr = sp.find(
                    f"{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr")
            if cNvPr is not None and cNvPr.get("id") == lc.cover_title_shape_id:
                txBody = sp.find(f"{{{P_NS}}}txBody")
                if txBody is None: continue

                orig_rPr = txBody.find(f".//{{{A_NS}}}rPr")
                import copy as _copy_inner
                rPr_el = _copy_inner.deepcopy(orig_rPr) if orig_rPr is not None else None

                for p_el in txBody.findall(f"{{{A_NS}}}p"):
                    txBody.remove(p_el)

                p_el  = etree.SubElement(txBody, f"{{{A_NS}}}p")
                pPr   = etree.SubElement(p_el,  f"{{{A_NS}}}pPr")
                lnSpc = etree.SubElement(pPr,   f"{{{A_NS}}}lnSpc")
                etree.SubElement(lnSpc, f"{{{A_NS}}}spcPct", val="100000")
                r_el  = etree.SubElement(p_el,  f"{{{A_NS}}}r")
                if rPr_el is not None:
                    r_el.append(rPr_el)
                t_el  = etree.SubElement(r_el,  f"{{{A_NS}}}t")
                t_el.text = title
                return
        # Shape not found — fall through to draw a textbox
    elif cover and lc.template_has_cover:
        return  # cover chrome already has a title box; skip

    # Standard textbox title (content slides + fallback cover)
    sz = 28 if cover else 24
    tf = _new_tf(slide, lc.title_l, lc.title_t, lc.title_w, lc.title_h)
    p  = tf.add_paragraph()
    if cover: p.alignment = PP_ALIGN.CENTER
    _no_bullet(p)
    _add_runs(p, title, sz, lc, color=lc.title_color, bold_all=True)


# ═══════════════════════════════════════════════════════════════════════════════
# Cover content
# ═══════════════════════════════════════════════════════════════════════════════

def _render_cover(slide, segments, lc: LayoutConfig, speaker_line: str = ""):
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
    # Strategy: use the cached template, deep-copy the original <a:r> elements
    # from the speaker shape (which carry sz=1301, colour=#59E3FA, the correct
    # fonts), then replace only <a:t> text.  This preserves the designer's
    # exact style without hard-coding any values.
    if (lc.template_has_cover and lc.cover_speaker_shape_id is not None
            and speaker_line):
        import copy as _copy_spk
        _src_slide = _get_cached_template(lc).slides[0]

        # Collect the three original <a:r> rPr elements from the template shape
        # (cyan "Speaker", white "|", cyan " year")
        orig_rPrs = []
        for _sp in _src_slide.shapes:
            if str(_sp.shape_id) == lc.cover_speaker_shape_id:
                for _r in _sp._element.findall(f".//{{{A_NS}}}r"):
                    rpr = _r.find(f"{{{A_NS}}}rPr")
                    orig_rPrs.append(_copy_spk.deepcopy(rpr) if rpr is not None else None)
                break
        # orig_rPrs[0] = cyan (left part), [1] = white "|", [2] = cyan (right part)

        for sp in slide._element.findall(f".//{{{P_NS}}}sp"):
            cNvPr = sp.find(f"{{{P_NS}}}nvSpPr/{{{P_NS}}}cNvPr")
            if cNvPr is not None and cNvPr.get("id") == lc.cover_speaker_shape_id:
                txBody = sp.find(f"{{{P_NS}}}txBody")
                if txBody is None: break

                spPr  = sp.find(f"{{{P_NS}}}spPr")
                if spPr is None:
                    spPr = sp.find(f".//{{{A_NS}}}spPr")
                xfrm  = spPr.find(f"{{{A_NS}}}xfrm") if spPr is not None else None
                if xfrm is not None:
                    off = xfrm.find(f"{{{A_NS}}}off")
                    ext = xfrm.find(f"{{{A_NS}}}ext")
                    if off is not None and ext is not None:
                        char_count = len(speaker_line) + 2
                        needed_w   = max(int(char_count * 9600), int(ext.get("cx", 0)))
                        right_edge = int(off.get("x", 0)) + int(ext.get("cx", 0))
                        new_x      = max(0, right_edge - needed_w)
                        off.set("x",  str(new_x))
                        ext.set("cx", str(right_edge - new_x))

                for p_el in txBody.findall(f"{{{A_NS}}}p"):
                    txBody.remove(p_el)

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
    if lc.template_has_cover:
        return   # cover template manages its own layout; don't add extra boxes

    all_blocks = []
    for st, sd in segments:
        if st == "text": all_blocks.extend(sd)
    if not all_blocks: return

    tf = _new_tf(slide, lc.content_l, lc.content_t, lc.content_w, lc.content_h)
    for block in all_blocks:
        bt = block.get("type")
        if bt in ("blank", "hr"):
            _spacer_para(tf, 6)
        elif bt == "heading":
            p = tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            _no_bullet(p)
            col = {3: lc.h3_color, 4: C_H4, 5: C_H5}.get(block["level"], lc.h3_color)
            _add_runs(p, block["text"], SZ_H4, lc, color=col, bold_all=True)
        else:
            p = tf.add_paragraph()
            p.alignment = PP_ALIGN.CENTER
            _no_bullet(p)
            _add_runs(p, block.get("text", ""), 16, lc)


# ═══════════════════════════════════════════════════════════════════════════════
# SmartArt renderers
# ═══════════════════════════════════════════════════════════════════════════════

import math

def _sa_palette(theme: str) -> list:
    """Return the colour list for the requested theme (falls back to blue)."""
    key = (theme or "").lower().strip()
    return _SA_PALETTES.get(key, _SA_PALETTES[_SA_DEFAULT_PALETTE])


def _sa_text_tf(slide, left, top, width, height, text, font_name,
                sz_pt=11, bold=False, color=None, align_center=True):
    """Add a transparent text-frame with horizontal and vertical centring.

    SmartArt labels are drawn as separate transparent text boxes above the
    coloured shapes.  PowerPoint honours vertical centring most reliably when
    both the python-pptx API value and the underlying DrawingML anchor are set,
    and when the default text margins / paragraph spacing are removed.
    """
    from pptx.util import Pt
    from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR, MSO_AUTO_SIZE

    txb = slide.shapes.add_textbox(int(left), int(top), int(width), int(height))
    txb.line.fill.background()

    tf = txb.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE

    # Remove the built-in text box padding.  Otherwise short Chinese labels
    # can look slightly high even when the frame itself is middle-anchored.
    tf.margin_top = 0
    tf.margin_bottom = 0
    tf.margin_left = 0
    tf.margin_right = 0

    # Extra compatibility for PowerPoint / LibreOffice XML readers.
    try:
        txb._txBody.get_or_add_bodyPr().set("anchor", "ctr")
    except Exception:
        pass

    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER if align_center else PP_ALIGN.LEFT
    p.space_before = Pt(0)
    p.space_after = Pt(0)

    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(sz_pt)
    run.font.bold = bold
    run.font.color.rgb = color or RGBColor(0xFF, 0xFF, 0xFF)
    return txb


def _sa_add_connector(slide, x1, y1, x2, y2, color):
    """Draw a thin line between two points (parent→child connector)."""
    from pptx.util import Pt
    from pptx.oxml.ns import qn
    import lxml.etree as etree
    # Use a thin rectangle as a connector (python-pptx connectors are fragile)
    if abs(x2 - x1) < abs(y2 - y1):   # mostly vertical
        lw = int(Inches(0.018))
        cx = (x1 + x2) / 2
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            int(cx - lw / 2), int(min(y1, y2)),
            lw, int(abs(y2 - y1))
        )
    else:                               # mostly horizontal
        lh = int(Inches(0.018))
        cy = (y1 + y2) / 2
        line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            int(min(x1, x2)), int(cy - lh / 2),
            int(abs(x2 - x1)), lh
        )
    line.fill.solid()
    line.fill.fore_color.rgb = color
    line.line.fill.background()


def _render_smartart_process(slide, sa, left, top, width, height, lc):
    """Chevron / numbered-card process diagram — auto-sizes to item count."""
    from pptx.util import Pt
    items = [it for it in sa["items"] if it["level"] == 0]
    if not items:
        return top + height
    pal   = _sa_palette(sa.get("theme", ""))
    style = (sa.get("style") or "chevron").lower()
    n     = len(items)

    if style == "chevron":
        # ── True pentagon/chevron via PENTAGON shape ──────────────────
        # Shape height: 55–65% of available height, vertically centred
        shape_h   = min(height * 0.60, Inches(1.4))
        shape_top = top + (height - shape_h) / 2

        # Each chevron overlaps the next one by `overlap` to form a chain
        overlap   = int(Inches(0.22))
        total_w   = width
        step_w    = (total_w + (n - 1) * overlap) / n   # gross width per tile

        for i, item in enumerate(items):
            color = pal[i % len(pal)]
            x     = left + i * (step_w - overlap)
            # Last tile is a plain rectangle cap; others get the arrow notch via
            # PENTAGON (points right).  We fake it with CHEVRON which gives a
            # left-notch: for shape i>0 use CHEVRON, for i==0+last use different.
            if n == 1:
                sh_type = MSO_SHAPE.ROUNDED_RECTANGLE
            elif i == 0:
                sh_type = MSO_SHAPE.PENTAGON
            elif i == n - 1:
                sh_type = MSO_SHAPE.CHEVRON
            else:
                sh_type = MSO_SHAPE.CHEVRON

            box = slide.shapes.add_shape(
                sh_type, int(x), int(shape_top), int(step_w), int(shape_h)
            )
            box.fill.solid()
            box.fill.fore_color.rgb = color
            box.line.fill.background()

            # Text: centre in the shape body (skip the arrow notch area ~18%)
            body_w = step_w * (0.78 if i < n - 1 else 0.88)
            body_x = x + (step_w - body_w) * 0.5 + (overlap * 0.3 if i > 0 else 0)
            sz = max(9, min(13, int(100 / max(n, 1))))
            _sa_text_tf(slide, body_x, shape_top, body_w, shape_h,
                        item["text"], lc.font_main, sz_pt=sz,
                        bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

        return int(shape_top + shape_h)

    else:  # numbered card style
        # ── Vertically centred list of number-bubble + card rows ──────
        row_h  = min(Inches(0.68), height / n)
        gap    = int(Inches(0.10))
        total  = n * row_h + (n - 1) * gap
        y0     = top + (height - total) / 2   # vertically centred

        bub_d  = int(row_h * 0.82)            # bubble diameter
        pad    = int(Inches(0.12))

        for i, item in enumerate(items):
            color  = pal[i % len(pal)]
            y      = y0 + i * (row_h + gap)
            by     = y + (row_h - bub_d) / 2

            # Number bubble
            bub = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, int(left), int(by), bub_d, bub_d
            )
            bub.fill.solid(); bub.fill.fore_color.rgb = color
            bub.line.fill.background()
            _sa_text_tf(slide, left, by, bub_d, bub_d,
                        str(i + 1), lc.font_main, sz_pt=11,
                        bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

            # Card: use a very light tint of the colour
            card_x = left + bub_d + pad
            card_w = width - bub_d - pad
            card   = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                int(card_x), int(y), int(card_w), int(row_h)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = pal[min(i + 3, len(pal) - 1)]
            card.line.fill.background()
            try: card.adjustments[0] = 0.05
            except Exception: pass

            # Text vertically centred inside card
            sz = max(9, min(12, int(110 / max(n, 1))))
            _sa_text_tf(slide, card_x + pad, y, card_w - 2 * pad, row_h,
                        item["text"], lc.font_main, sz_pt=sz,
                        bold=False, color=C_BODY, align_center=False)

        return int(y0 + total)


def _render_smartart_cycle(slide, sa, left, top, width, height, lc):
    """Circular cycle diagram — elliptical orbit, equal-sized nodes."""
    items = [it for it in sa["items"] if it["level"] == 0]
    if not items:
        return top + height
    pal = _sa_palette(sa.get("theme", ""))
    n   = len(items)

    # Centre of the available area
    cx = left + width  / 2
    cy = top  + height / 2

    # Compute node diameter: large enough to read, small enough to fit n nodes
    # Ensure orbit radius keeps nodes inside the bounding box
    max_node_d = min(width / (n * 0.9 + 0.5), height * 0.44, Inches(1.1))
    node_d     = max(int(Inches(0.55)), int(max_node_d))

    # Orbit radii — cap at 75% of half-dimensions so nodes have clear margins
    rx = min(width  / 2 * 0.78 - node_d / 2, width  / 2 - node_d * 0.7)
    ry = min(height / 2 * 0.84 - node_d / 2, height / 2 - node_d * 0.7)
    rx = max(rx, node_d * 0.5)
    ry = max(ry, node_d * 0.5)

    def _contrast(rgb):
        """Return white or dark text colour for best contrast on rgb."""
        r, g, b = rgb
        lum = 0.299*r + 0.587*g + 0.114*b
        return RGBColor(0xFF,0xFF,0xFF) if lum < 160 else RGBColor(0x1F,0x29,0x37)

    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    sz    = max(8, min(12, int(72 / max(n, 1))))

    for i, item in enumerate(items):
        angle = math.radians(-90 + i * 360 / n)
        nx    = cx + rx * math.cos(angle)
        ny    = cy + ry * math.sin(angle)
        color = pal[i % len(pal)]
        txt_color = _contrast(color)

        # Draw node circle
        circ = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            int(nx - node_d / 2), int(ny - node_d / 2),
            int(node_d), int(node_d)
        )
        circ.fill.solid(); circ.fill.fore_color.rgb = color
        circ.line.fill.background()

        # Label
        _sa_text_tf(slide,
                    nx - node_d / 2, ny - node_d / 2,
                    node_d, node_d, item["text"],
                    lc.font_main, sz_pt=sz, bold=True, color=txt_color)

        # Small direction dot between this node and the next (inside orbit)
        a_mid  = math.radians(-90 + (i + 0.5) * 360 / n)
        arr_r  = (rx + ry) / 2 * 0.55    # place well inside the orbit ring
        ax     = cx + arr_r * math.cos(a_mid)
        ay     = cy + arr_r * math.sin(a_mid)
        arr_d2 = int(node_d * 0.18)
        arr    = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            int(ax - arr_d2 / 2), int(ay - arr_d2 / 2),
            int(arr_d2), int(arr_d2)
        )
        arr.fill.solid(); arr.fill.fore_color.rgb = pal[2]
        arr.line.fill.background()

    # Subtle centre circle
    c_r = int(min(rx, ry) * 0.26)
    if c_r > int(Inches(0.1)):
        cc = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            int(cx - c_r), int(cy - c_r), int(c_r * 2), int(c_r * 2)
        )
        cc.fill.solid(); cc.fill.fore_color.rgb = pal[-1]
        cc.line.fill.background()

    return int(top + height)


def _render_smartart_hierarchy(slide, sa, left, top, width, height, lc):
    """Top-down tree with connecting lines, adaptive node sizing.

    Render in two passes so connector lines stay underneath the node boxes.
    PowerPoint uses creation order as the z-order, so if connectors are added
    after child boxes they visually cut across those boxes.  We therefore first
    compute all node positions and connector segments, then draw all connectors,
    and only then draw the rounded node boxes and their labels.
    """
    items = sa["items"]
    if not items:
        return top + height
    pal = _sa_palette(sa.get("theme", ""))

    # ── Build tree ────────────────────────────────────────────────────
    def build_tree(items):
        root  = {"text": "", "children": []}
        stack = [(root, -1)]
        for it in items:
            node = {"text": it["text"], "children": []}
            lvl  = it["level"]
            while len(stack) > 1 and stack[-1][1] >= lvl:
                stack.pop()
            stack[-1][0]["children"].append(node)
            stack.append((node, lvl))
        return root["children"]

    def max_depth(nodes, d=0):
        if not nodes:
            return d
        return max(max_depth(c["children"], d + 1) for c in nodes)

    def count_leaves(nodes):
        if not nodes:
            return 1
        return sum(count_leaves(c["children"]) for c in nodes)

    tree     = build_tree(items)
    depth    = max_depth(tree) + 1
    n_leaves = count_leaves(tree)

    # Node dimensions — adaptive to depth and leaf count
    node_h = min(height / (depth * 1.45), Inches(0.56))
    node_w = min(width  / max(n_leaves, 1) * 0.82, Inches(1.9))
    v_step = height / depth
    LINE_C = pal[min(1, len(pal) - 1)]
    txt_sz = max(7, min(11, int(90 / max(n_leaves, 1))))

    node_specs = []
    connectors = []

    def layout_level(nodes, level, x_start, x_end):
        if not nodes:
            return []
        total_leaves = count_leaves(nodes)
        x_cur        = x_start
        centres      = []
        for node in nodes:
            lc_count = count_leaves(node["children"]) if node["children"] else 1
            share    = (x_end - x_start) * lc_count / total_leaves
            nx       = x_cur + share / 2
            ny       = top + level * v_step + (v_step - node_h) / 2
            color    = pal[min(level, len(pal) - 1)]

            node_specs.append({
                "x": nx - node_w / 2,
                "y": ny,
                "w": node_w,
                "h": node_h,
                "text": node["text"],
                "level": level,
                "color": color,
            })
            centres.append((nx, ny, ny + node_h))

            if node["children"]:
                child_centres = layout_level(node["children"], level + 1,
                                             x_cur, x_cur + share)
                parent_cx = nx
                parent_by = ny + node_h
                mid_y     = ny + v_step * 0.5

                connectors.append((parent_cx, parent_by, parent_cx, mid_y))
                if len(child_centres) > 1:
                    xs = [c[0] for c in child_centres]
                    connectors.append((min(xs), mid_y, max(xs), mid_y))
                for ccx, cny, _ in child_centres:
                    connectors.append((ccx, mid_y, ccx, cny))

            x_cur += share
        return centres

    layout_level(tree, 0, left, left + width)

    # Pass 1: connectors first, so they stay behind the boxes.
    for x1, y1, x2, y2 in connectors:
        _sa_add_connector(slide, x1, y1, x2, y2, LINE_C)

    # Pass 2: node boxes and labels above the connector lines.
    for spec in node_specs:
        box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            int(spec["x"]), int(spec["y"]),
            int(spec["w"]), int(spec["h"])
        )
        box.fill.solid()
        box.fill.fore_color.rgb = spec["color"]
        box.line.fill.background()
        try:
            box.adjustments[0] = 0.12
        except Exception:
            pass

        _sa_text_tf(slide,
                    spec["x"], spec["y"], spec["w"], spec["h"],
                    spec["text"], lc.font_main,
                    sz_pt=txt_sz, bold=(spec["level"] == 0),
                    color=RGBColor(0xFF, 0xFF, 0xFF))

    return int(top + height)


def _render_smartart_matrix(slide, sa, left, top, width, height, lc):
    """2×2 (or 2×3) matrix with centred labels and quadrant axis labels."""
    items = [it for it in sa["items"] if it["level"] == 0]
    if not items:
        return top + height
    pal  = _sa_palette(sa.get("theme", ""))
    size = sa.get("size", "2x2")
    try:
        rows, cols = [int(x) for x in size.lower().split("x")]
    except Exception:
        rows = cols = 2

    n     = rows * cols
    items = items[:n]

    gap   = int(Inches(0.07))
    cell_w = (width  - gap * (cols + 1)) / cols
    cell_h = (height - gap * (rows + 1)) / rows

    WHITE = RGBColor(0xFF, 0xFF, 0xFF)

    for idx, item in enumerate(items):
        r   = idx // cols
        c   = idx %  cols
        x   = left + gap + c * (cell_w + gap)
        y   = top  + gap + r * (cell_h + gap)
        color = pal[idx % len(pal)]

        box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            int(x), int(y), int(cell_w), int(cell_h)
        )
        box.fill.solid(); box.fill.fore_color.rgb = color
        box.line.fill.background()
        try: box.adjustments[0] = 0.08
        except Exception: pass

        # Text: split on " / " or " : " for two-line labels
        label = item["text"].replace(" / ", "\n").replace("：", "\n")
        sz    = max(9, min(13, int(120 / max(n, 1))))
        _sa_text_tf(slide, x, y, cell_w, cell_h, label,
                    lc.font_main, sz_pt=sz, bold=True, color=WHITE)

    return int(top + height)


def _render_smartart_pyramid(slide, sa, left, top, width, height, lc):
    """True trapezoid pyramid using ISOSCELES_TRIANGLE clipping trick.

    Each tier is rendered as a coloured trapezoid approximated by overlapping
    a rectangle with left/right triangles cut out — or more simply as a
    wide-to-narrow stack of rectangles with symmetric indentation.
    """
    items = [it for it in sa["items"] if it["level"] == 0]
    if not items:
        return top + height
    pal = _sa_palette(sa.get("theme", ""))
    n   = len(items)

    # Adaptive sizing: use most of the available height, centred
    tier_h   = min(height / n, Inches(0.72))
    total_h  = tier_h * n
    y0       = top + (height - total_h) / 2    # vertically centred
    gap_v    = int(Inches(0.04))

    # Pyramid spans full width at base, tapers to ~30% at apex
    base_w   = width
    apex_frac = 0.28   # top tier width as fraction of base

    WHITE = RGBColor(0xFF, 0xFF, 0xFF)

    for i, item in enumerate(items):
        # i=0 → top (narrowest apex), i=n-1 → base (widest)
        frac   = apex_frac + (1.0 - apex_frac) * i / max(n - 1, 1)
        tier_w = base_w * frac
        x      = left + (base_w - tier_w) / 2
        y      = y0 + i * (tier_h - gap_v)   # slight overlap for seamless look
        color  = pal[i % len(pal)]

        # Draw trapezoid using python-pptx freeform (available via add_shape
        # with custom XML, or approximate with a rounded rect)
        # Use ROUNDED_RECTANGLE for visual cleanliness
        box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            int(x), int(y), int(tier_w), int(tier_h - gap_v)
        )
        box.fill.solid(); box.fill.fore_color.rgb = color
        box.line.fill.background()
        try: box.adjustments[0] = 0.05
        except Exception: pass

        sz = max(8, min(12, int(100 / n)))
        _sa_text_tf(slide, x, y, tier_w, tier_h - gap_v,
                    item["text"], lc.font_main,
                    sz_pt=sz, bold=(i == 0), color=WHITE)

    return int(y0 + total_h)


def _render_smartart_timeline(slide, sa, left, top, width, height, lc):
    """Horizontal timeline — adaptive label/card layout, fully centred."""
    items = [it for it in sa["items"] if it["level"] == 0]
    if not items:
        return top + height
    pal = _sa_palette(sa.get("theme", ""))
    n   = len(items)

    # ── Layout constants ──────────────────────────────────────────────
    has_desc = any(" | " in it["text"] for it in items)

    label_h = int(Inches(0.45))          # area above the line for date labels
    dot_d   = int(Inches(0.26))          # milestone dot diameter
    line_h  = int(Inches(0.05))          # spine thickness
    card_h  = int(Inches(1.0)) if has_desc else 0
    gap     = int(Inches(0.12))

    # Total used height; centre vertically
    used_h  = label_h + gap + dot_d + (gap + card_h if has_desc else 0)
    y0      = top + (height - used_h) / 2

    line_y  = y0 + label_h + gap               # spine top
    dot_y   = line_y + (line_h - dot_d) / 2    # dot top (centred on spine)
    card_y  = line_y + line_h + gap            # card top

    step    = width / n
    WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
    sz_lbl  = max(8, min(11, int(80 / n)))
    sz_desc = max(7, min(10, int(75 / n)))

    # ── Spine line ────────────────────────────────────────────────────
    spine = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        int(left), int(line_y), int(width), int(line_h)
    )
    spine.fill.solid(); spine.fill.fore_color.rgb = pal[1]
    spine.line.fill.background()

    for i, item in enumerate(items):
        xc    = left + (i + 0.5) * step
        color = pal[i % max(len(pal) - 1, 1)]

        # ── Milestone dot ─────────────────────────────────────────────
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            int(xc - dot_d / 2), int(dot_y),
            int(dot_d), int(dot_d)
        )
        dot.fill.solid(); dot.fill.fore_color.rgb = color
        dot.line.fill.background()

        # ── Parse label | description ─────────────────────────────────
        if " | " in item["text"]:
            label, desc = item["text"].split(" | ", 1)
        else:
            label, desc = item["text"], ""
        label = label.strip(); desc = desc.strip()

        # ── Label above spine ─────────────────────────────────────────
        lw = step * 0.88
        _sa_text_tf(slide, xc - lw / 2, y0, lw, label_h,
                    label, lc.font_main, sz_pt=sz_lbl,
                    bold=True, color=color, align_center=True)

        # ── Vertical stem from dot to card ───────────────────────────
        if has_desc:
            stem_w = int(Inches(0.018))
            stem = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                int(xc - stem_w / 2), int(dot_y + dot_d),
                stem_w, int(card_y - dot_y - dot_d)
            )
            stem.fill.solid(); stem.fill.fore_color.rgb = pal[2]
            stem.line.fill.background()

        # ── Description card below spine ─────────────────────────────
        if has_desc:
            cw   = step * 0.84
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                int(xc - cw / 2), int(card_y),
                int(cw), int(card_h)
            )
            card.fill.solid(); card.fill.fore_color.rgb = pal[min(i + 2, len(pal) - 1)]
            card.line.fill.background()
            try: card.adjustments[0] = 0.08
            except Exception: pass

            if desc:
                _sa_text_tf(slide, xc - cw / 2, card_y, cw, card_h,
                            desc, lc.font_main, sz_pt=sz_desc,
                            bold=False, color=C_BODY, align_center=True)

    return int(y0 + used_h)


def _render_smartart(slide, sa: dict, left, top, width, height, lc) -> int:
    """Dispatch to the correct SmartArt renderer. Returns bottom y."""
    sa_type = (sa.get("type") or "process").lower()
    if sa_type == "cycle":
        return _render_smartart_cycle(slide, sa, left, top, width, height, lc)
    elif sa_type == "hierarchy":
        return _render_smartart_hierarchy(slide, sa, left, top, width, height, lc)
    elif sa_type == "matrix":
        return _render_smartart_matrix(slide, sa, left, top, width, height, lc)
    elif sa_type == "pyramid":
        return _render_smartart_pyramid(slide, sa, left, top, width, height, lc)
    elif sa_type == "timeline":
        return _render_smartart_timeline(slide, sa, left, top, width, height, lc)
    else:  # process (default)
        return _render_smartart_process(slide, sa, left, top, width, height, lc)


# ═══════════════════════════════════════════════════════════════════════════════
# Body segments with y-cursor (from micknoise approach)
# ═══════════════════════════════════════════════════════════════════════════════

def _plain_text_len(text: str) -> int:
    """Approximate visible text length by stripping markdown markup."""
    if not text:
        return 0
    try:
        parts = parse_inline(text)
        return sum(len(t or "") for (t, *_rest) in parts)
    except Exception:
        return len(text)


def _estimate_wrapped_lines(text: str, width_emu: int, font_pt: int) -> int:
    """Estimate how many visual lines a markdown text block will occupy."""
    width_in = max(width_emu / 914400, 0.8)
    vis_text = text or " "
    cjk = sum(1 for ch in vis_text if '一' <= ch <= '鿿')
    lat = sum(1 for ch in vis_text if ch.isalpha() and ord(ch) < 0x300)
    density = 54 if cjk >= lat else 95
    chars_per_line = max(int(width_in * (density / 9.5) * (16 / max(font_pt, 1))), 12)
    n_wrap = 0
    for part in vis_text.splitlines() or [" "]:
        plen = max(_plain_text_len(part), 1)
        n_wrap += max(1, (plen + chars_per_line - 1) // chars_per_line)
    return max(n_wrap, 1)


def _estimate_flow_text_height(block: dict, sz_pt: int, lc: LayoutConfig) -> int:
    """Estimate the vertical space consumed by one text block in the body flow.

    The text itself is still rendered in one combined textbox for reliable
    bullets/inline formatting.  This helper is used by the y-cursor that places
    tables, images, code blocks and SmartArt, so those objects start after the
    preceding paragraphs instead of being drawn on top of them.

    Earlier versions used a fixed one-line estimate for normal paragraphs,
    which caused overlap whenever a long sentence wrapped into two or more
    visual lines before a following table / image / SmartArt.  This estimator
    now approximates wrapping based on content width and font size.
    """
    bt = block.get("type")
    if bt == "blank":
        return int(Pt(4))
    if bt == "hr":
        return int(Pt(10))
    if bt == "image":
        return int(Pt(sz_pt + 8))
    if bt == "quote":
        txt_w = int(lc.content_w - Inches(0.055) - Inches(0.18))
        q_text = block.get("text", "") or " "
        n_wrap = _estimate_wrapped_lines(q_text, txt_w, sz_pt)
        return int(Pt(sz_pt + 4) * n_wrap + Pt(8))
    if bt == "heading":
        level = block.get("level", 3)
        hsz = {3: SZ_H3, 4: SZ_H4, 5: SZ_H5}.get(level, SZ_H3)
        txt = block.get("text", "") or " "
        n_wrap = _estimate_wrapped_lines(txt, int(lc.content_w), hsz)
        return int(Pt(hsz + 4) * n_wrap + Pt(8))
    if bt in ("para", "bullet", "ordered"):
        txt = block.get("text", "") or " "
        indent = int(block.get("indent", 0) or 0)
        mar_in = 0.38 if bt == "ordered" else (0.75 if indent else 0.38)
        width_emu = max(int(lc.content_w - Inches(mar_in)), int(Inches(1.0)))
        n_wrap = _estimate_wrapped_lines(txt, width_emu, sz_pt)
        line_h = Pt(sz_pt + 5)
        extra = Pt(3 if bt in ("bullet", "ordered") else 2)
        return int(line_h * n_wrap + extra)
    return 0

def _render_body(slide, segments, body_lines, lc: LayoutConfig, md_dir: str = ""):
    n_lines = sum(1 for l in body_lines
                  if l.strip() and not l.strip().startswith("|")
                  and not l.strip().startswith("```"))
    cjk = _is_cjk_dominant(body_lines)
    sz  = auto_body_size(n_lines, cjk=cjk)

    # ── Collect all text blocks across every text segment into one list.
    #    Tables, code blocks and images interrupt the text flow; handle
    #    them via a y-cursor that advances around them.
    all_text_blocks = []
    for seg_type, seg_data in segments:
        if seg_type == "text":
            all_text_blocks.extend(seg_data)
        else:
            # Non-text sentinel so we know where tables/code/images appear
            all_text_blocks.append({"type": "_break_", "seg_type": seg_type, "data": seg_data})

    # ── Overflow estimation ───────────────────────────────────────────
    # Count estimated lines to warn when content will overflow the slide.
    est_lines = 0
    for block in all_text_blocks:
        bt = block.get("type")
        if bt == "_break_":
            seg_type = block["seg_type"]
            if seg_type == "table":
                est_lines += max(len(block["data"]), 1) * 1.5
            elif seg_type in ("code", "image"):
                est_lines += 3
            elif seg_type == "smartart":
                est_lines += 6   # SmartArt takes roughly a full content area
        elif bt not in ("blank", "hr", None):
            est_lines += 1
    limit = {18: 6, 16: 10, 14: 14, 13: 16}.get(sz, 10)
    if est_lines > limit:
        print(f"  ⚠ 内容可能溢出（预估 {int(est_lines)} 行，字号 {sz}pt 上限约 {limit} 行）"
              f"  —— 建议拆分为两张幻灯片")

    # ── Pass 1: draw quote bar shapes (behind the tf in z-order).
    #    Walk all_text_blocks to accumulate estimated y-offsets.
    bar_w    = Inches(0.055)
    y_cursor = int(lc.content_t)
    gap      = int(Pt(10))

    for block in all_text_blocks:
        if y_cursor >= int(lc.content_b): break
        bt = block.get("type")
        if bt == "_break_":
            # Rough height estimate for table/code/image to advance cursor
            y_cursor += int(Pt(18) * max(
                len(block["data"]) if isinstance(block["data"], list) else 3, 3
            ) + gap)
        elif bt == "quote":
            bh = _estimate_flow_text_height(block, sz, lc)
            bar = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                int(lc.content_l + Inches(0.05)),
                y_cursor + int(Pt(2)),
                int(bar_w),
                max(bh - int(Pt(6)), int(Pt(8)))
            )
            bar.fill.solid()
            bar.fill.fore_color.rgb = C_QUOTE_BAR
            bar.line.fill.background()
            y_cursor += bh
        else:
            y_cursor += _estimate_flow_text_height(block, sz, lc)

    # ── Pass 2: create ONE full-height textbox for all text content,
    #    then render table/code/image blocks via the y-cursor.
    text_blocks_only = [b for b in all_text_blocks if b.get("type") != "_break_"]
    if text_blocks_only and not all(b.get("type") == "blank" for b in text_blocks_only):
        tf = _new_tf(slide, lc.content_l, int(lc.content_t), lc.content_w, int(lc.content_h))
        _render_text_blocks(tf, text_blocks_only, sz, lc)

    # ── Pass 3: render table/code/image blocks at their approximate positions.
    y    = int(lc.content_t)
    gap  = int(Pt(10))
    for block in all_text_blocks:
        bt = block.get("type")

        # Keep the object y-cursor in the same document flow as the text.
        # Previously only non-text breaks advanced this cursor; therefore a
        # SmartArt block placed after a paragraph was still drawn at the top of
        # the content area and could cover that paragraph.
        if bt != "_break_":
            y += _estimate_flow_text_height(block, sz, lc)
            continue

        if y >= int(lc.content_b): break
        seg_type = block["seg_type"]
        seg_data = block["data"]
        if seg_type == "table":
            bottom = _add_pptx_table(slide, seg_data, lc.content_l, y, lc.content_w, lc)
            y = int(bottom) + gap
        elif seg_type == "code":
            bottom = _add_code_block(slide, seg_data,
                                     lc.content_l + Inches(0.15), y,
                                     lc.content_w - Inches(0.15))
            y = int(bottom) + gap
        elif seg_type == "image":
            # Attempt to embed local image; fall back to caption text
            src = seg_data.get("src", "")
            alt = seg_data.get("alt", "") or src
            img_path = Path(src) if Path(src).is_absolute() else Path(md_dir) / src
            avail_h  = int(lc.content_b) - y - gap
            img_w    = int(lc.content_w)
            if img_path.exists() and img_path.suffix.lower() in (
                    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif"):
                try:
                    pic = slide.shapes.add_picture(
                        str(img_path), int(lc.content_l), y, img_w
                    )
                    # Scale down if the picture overflows the content area
                    if pic.height > avail_h:
                        ratio = avail_h / pic.height
                        pic.height = avail_h
                        pic.width  = int(pic.width * ratio)
                    y = int(y + pic.height + gap)
                    continue
                except Exception as e:
                    print(f"  ⚠ 图片嵌入失败 ({img_path}): {e}，改为占位文本")
            # Fallback: placeholder caption
            cap_h = int(Pt(sz + 4) * 2)
            tf = _new_tf(slide, int(lc.content_l), y, int(lc.content_w), cap_h)
            p  = tf.add_paragraph()
            _no_bullet(p)
            r = p.add_run()
            r.text           = f"[图片: {alt}]" if alt else "[图片]"
            r.font.name      = lc.font_main
            r.font.size      = Pt(sz - 1)
            r.font.italic    = True
            r.font.color.rgb = C_QUOTE
            y += cap_h + gap
        elif seg_type == "smartart":
            avail_h = int(lc.content_b) - y - gap
            # SmartArt fills remaining available height (min 1.4", max 3.2")
            sa_h    = max(int(Inches(1.4)), min(avail_h, int(Inches(3.2))))
            # Use the full slide width (not the narrow column width of some templates)
            sa_w    = max(int(lc.content_w), int(lc.slide_w - 2 * lc.content_l))
            bottom  = _render_smartart(slide, seg_data,
                                       int(lc.content_l), y,
                                       int(sa_w), sa_h, lc)
            y = int(bottom) + gap


# ═══════════════════════════════════════════════════════════════════════════════
# Speaker notes
# ═══════════════════════════════════════════════════════════════════════════════

def _set_notes(slide, note_text: str, lc: LayoutConfig):
    tf  = slide.notes_slide.notes_text_frame
    tf.clear()
    p   = tf.paragraphs[0]
    run = p.add_run()
    run.text           = note_text
    run.font.name      = lc.font_main
    run.font.size      = Pt(11)
    run.font.color.rgb = C_BODY



def _deduplicate_pptx_zip(pptx_path: str):
    """Rewrite a PPTX package with duplicate ZIP entries removed."""
    import zipfile
    import os
    from pathlib import Path

    src = Path(pptx_path)
    tmp = src.with_suffix(src.suffix + ".dedup.tmp")
    seen = set()

    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename in seen:
                continue
            seen.add(info.filename)
            zout.writestr(info, zin.read(info.filename))

    os.replace(tmp, src)


# ═══════════════════════════════════════════════════════════════════════════════
# Main conversion
# ═══════════════════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════════════════
# Template layout detection
# ═══════════════════════════════════════════════════════════════════════════════

def load_layout(template_path: str, config_path: str = None) -> LayoutConfig:
    """
    Auto-detect layout dimensions from a PPTX template and return a
    LayoutConfig.  Override order (highest priority first):

      1. JSON config file (explicit overrides — most flexible)
      2. Auto-detection from the template slides' shapes
      3. LayoutConfig defaults (bundled template dimensions)

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
      "font_main": "Source Han Sans SC",
      "title_color": "00B0F0",
      "h3_color": "005A87"
    }
    All fields are optional — omit any you want to keep at their detected value.
    """
    import copy as _copy
    import json

    lc  = LayoutConfig(template_path=template_path)
    prs = Presentation(template_path)
    # Also pre-warm the cache so _add_decorative doesn't re-open the file
    _TEMPLATE_CACHE[template_path] = prs

    # ── Step 1: read slide size from template ─────────────────────────
    lc.slide_w = prs.slide_width
    lc.slide_h = prs.slide_height

    def _in(emu): return emu / 914400   # EMU → inches

    # ── Step 2: detect whether template has a dedicated cover slide ───
    n_slides = len(prs.slides)

    def _looks_like_cover(slide):
        """True if slide appears to be a cover (no tall content area WITH text)."""
        for shape in slide.shapes:
            if not hasattr(shape, "text"): continue
            if not shape.text.strip(): continue
            h = _in(shape.height)
            t = _in(shape.top)
            w = _in(shape.width)
            txt = shape.text.strip()
            if h > 1.5 and w > 4 and t >= 0.8 and len(txt) > 40:
                return False
        return True

    if n_slides >= 2 and _looks_like_cover(prs.slides[0]):
        cover_slide   = prs.slides[0]
        content_slide = prs.slides[1]
        lc.template_has_cover = True
        print("  Template: cover slide detected (slide 1) + content slide (slide 2)")
    else:
        cover_slide   = None
        content_slide = prs.slides[0]
        lc.template_has_cover = False
        print("  Template: single-layout (no dedicated cover slide)")

    # ── Step 3: detect cover text box IDs ────────────────────────────
    if cover_slide is not None:
        lc.cover_layout_source = cover_slide.slide_layout
        best_title_area   = -1
        best_speaker_area = -1
        for shape in cover_slide.shapes:
            if not hasattr(shape, "text"): continue
            area = _in(shape.width) * _in(shape.height)
            txt  = shape.text.strip()
            sid  = str(shape.shape_id)
            if "|" in txt or "Speaker" in txt or "year" in txt.lower():
                if area > best_speaker_area:
                    best_speaker_area           = area
                    lc.cover_speaker_shape_id   = sid
            elif txt and area > best_title_area:
                best_title_area         = area
                lc.cover_title_shape_id = sid
        print(f"  Cover shape IDs: title={lc.cover_title_shape_id}  "
              f"speaker={lc.cover_speaker_shape_id}")

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
        lc.title_l = title_shape.left;  lc.title_t = title_shape.top
        lc.title_w = title_shape.width; lc.title_h = title_shape.height

    if content_shape:
        lc.content_l = content_shape.left;  lc.content_t = content_shape.top
        lc.content_w = content_shape.width; lc.content_h = content_shape.height

    for shape in deco_shapes:
        t = _in(shape.top)
        if t > _in(lc.slide_h) * 0.8:
            lc.slide_num_l = shape.left;  lc.slide_num_t = shape.top
            lc.slide_num_w = shape.width; lc.slide_num_h = shape.height
            break

    # ── Step 5: store layout references and deco element lists ────────
    lc.layout_source = content_slide.slide_layout

    lc.deco_elements = []
    for shape in deco_shapes:
        t = _in(shape.top)
        if t > _in(lc.slide_h) * 0.8 and hasattr(shape, "text"):
            continue
        lc.deco_elements.append(_copy.deepcopy(shape._element))

    lc.cover_deco_elements = []
    if cover_slide is not None:
        for shape in cover_slide.shapes:
            lc.cover_deco_elements.append(_copy.deepcopy(shape._element))

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

    lc.slide_w     = _i("slide_w",     lc.slide_w)
    lc.slide_h     = _i("slide_h",     lc.slide_h)
    lc.title_l     = _i("title_l",     lc.title_l)
    lc.title_t     = _i("title_t",     lc.title_t)
    lc.title_w     = _i("title_w",     lc.title_w)
    lc.title_h     = _i("title_h",     lc.title_h)
    lc.content_l   = _i("content_l",   lc.content_l)
    lc.content_t   = _i("content_t",   lc.content_t)
    lc.content_w   = _i("content_w",   lc.content_w)
    lc.content_h   = _i("content_h",   lc.content_h)
    lc.slide_num_l = _i("slide_num_l", lc.slide_num_l)
    lc.slide_num_t = _i("slide_num_t", lc.slide_num_t)
    lc.slide_num_w = _i("slide_num_w", lc.slide_num_w)
    lc.slide_num_h = _i("slide_num_h", lc.slide_num_h)

    if "font_main" in cfg:
        lc.font_main = cfg["font_main"]
    if "title_color" in cfg:
        v = cfg["title_color"].lstrip("#")
        lc.title_color = RGBColor(int(v[0:2],16), int(v[2:4],16), int(v[4:6],16))
    if "h3_color" in cfg:
        v = cfg["h3_color"].lstrip("#")
        lc.h3_color = RGBColor(int(v[0:2],16), int(v[2:4],16), int(v[4:6],16))
    if cfg.get("copy_deco") is False:
        lc.deco_elements = []

    def _fmt(emu): return f'{emu/914400:.3f}"'
    print(f"  Layout: slide={_fmt(lc.slide_w)}×{_fmt(lc.slide_h)}"
          f"  title=({_fmt(lc.title_l)},{_fmt(lc.title_t)}) {_fmt(lc.title_w)}×{_fmt(lc.title_h)}"
          f"  content=({_fmt(lc.content_l)},{_fmt(lc.content_t)}) {_fmt(lc.content_w)}×{_fmt(lc.content_h)}"
          f"  deco_shapes={len(lc.deco_elements)}")
    return lc



def _deduplicate_pptx_zip(pptx_path: str):
    """Rewrite a PPTX package with duplicate ZIP entries removed."""
    import zipfile
    import os
    from pathlib import Path

    src = Path(pptx_path)
    tmp = src.with_suffix(src.suffix + ".dedup.tmp")
    seen = set()

    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename in seen:
                continue
            seen.add(info.filename)
            zout.writestr(info, zin.read(info.filename))

    os.replace(tmp, src)


# ═══════════════════════════════════════════════════════════════════════════════
# Main conversion
# ═══════════════════════════════════════════════════════════════════════════════

def convert(md_path: str, template_path: str, output_path: str,
            config_path: str = None):
    import chardet

    # Auto-detect layout from the supplied template; returns a LayoutConfig
    print(f"Loading template: {template_path}")
    lc = load_layout(template_path, config_path=config_path)

    raw      = Path(md_path).read_bytes()
    encoding = chardet.detect(raw).get("encoding") or "utf-8"
    md_text  = raw.decode(encoding)
    md_dir   = str(Path(md_path).parent)   # for resolving local image paths

    slides_data = parse_md(md_text)
    if not slides_data:
        print("ERROR: No slides found.", file=sys.stderr)
        sys.exit(1)

    print(f"Parsed {len(slides_data)} slide(s).")

    # Build output from a blank presentation (the original behaviour).
    #
    # Do NOT generate normal pages from the template presentation itself:
    # some templates contain title/content placeholders, which can appear as
    # "Click here to add title/text" over generated content. We only use the
    # template to detect layout and to clone its last slide as a closing page.
    prs = Presentation()
    prs.slide_width  = lc.slide_w
    prs.slide_height = lc.slide_h
    blank_layout = prs.slide_layouts[6]   # completely blank

    def _append_template_closing_slide(prs_obj, template_path_str: str):
        """Append the last slide of template_path as a closing slide.

        The generated content slides remain exactly as before: they are created
        from a blank presentation and do not inherit template placeholders. The
        template is opened separately only for copying its last slide's shapes.
        """
        import copy
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT

        try:
            template_prs = Presentation(template_path_str)
            if len(template_prs.slides) == 0:
                return False
            src_slide = template_prs.slides[-1]
        except Exception as e:
            print(f"  ⚠ Unable to read template closing slide: {e}")
            return False

        dst_slide = prs_obj.slides.add_slide(blank_layout)

        # Copy slide background when present.
        try:
            src_bg = src_slide._element.cSld.find(qn("p:bg"))
            if src_bg is not None:
                dst_cSld = dst_slide._element.cSld
                dst_bg = dst_cSld.find(qn("p:bg"))
                if dst_bg is not None:
                    dst_cSld.remove(dst_bg)
                dst_cSld.insert(0, copy.deepcopy(src_bg))
        except Exception:
            pass

        # Copy non-layout relationships and update relationship IDs in cloned
        # XML. This supports images/media in the closing slide when present.
        rel_map = {}
        for rId, rel in list(src_slide.part.rels.items()):
            if rel.reltype == RT.SLIDE_LAYOUT:
                continue
            try:
                target = rel.target_ref if rel.is_external else rel.target_part
                new_rId = dst_slide.part.rels._add_relationship(
                    rel.reltype, target, rel.is_external
                )
                rel_map[rId] = new_rId
            except Exception:
                pass

        def _replace_rel_ids(el):
            if not rel_map:
                return
            for node in el.iter():
                for attr_name, attr_val in list(node.attrib.items()):
                    if attr_val in rel_map:
                        node.attrib[attr_name] = rel_map[attr_val]

        # Copy all shapes from the closing slide.
        for shape in src_slide.shapes:
            try:
                new_el = copy.deepcopy(shape.element)
                _replace_rel_ids(new_el)
                dst_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")
            except Exception as e:
                print(f"  ⚠ Unable to copy a closing-slide shape: {e}")

        return True

    for idx, sd in enumerate(slides_data):
        slide    = prs.slides.add_slide(blank_layout)
        cover    = sd.get("cover", False)
        speaker  = sd.get("speaker", "")
        segments = classify_body(sd["body_lines"])

        _add_decorative(slide, idx + 1, lc, cover=cover)
        _add_title(slide, sd["title"], lc, cover=cover)

        if cover:
            _render_cover(slide, segments, lc, speaker_line=speaker)
        else:
            _render_body(slide, segments, sd["body_lines"], lc, md_dir=md_dir)

        if sd["note"]:
            _set_notes(slide, sd["note"], lc)

        print(f"  Slide {idx+1}: {sd['title'][:60]}")

    if _append_template_closing_slide(prs, template_path):
        print("  Closing slide: appended from the last slide of the template")

    prs.save(output_path)
    _deduplicate_pptx_zip(output_path)
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
