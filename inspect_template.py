#!/usr/bin/env python3
"""
inspect_template.py — 分析 PPTX 模板结构，输出人类可读报告并可选生成 JSON 配置。

Usage:
    python inspect_template.py template.pptx              # 仅打印报告
    python inspect_template.py template.pptx --json out.json  # 同时生成配置
"""

import sys
import json
import argparse
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches

def emu_to_in(emu):
    return round(emu / 914400, 4)

def inspect(template_path: str, json_out: str = None):
    prs   = Presentation(template_path)
    slide = prs.slides[0]

    sw = emu_to_in(prs.slide_width)
    sh = emu_to_in(prs.slide_height)

    print(f"\n{'='*60}")
    print(f"模板：{template_path}")
    print(f"尺寸：{sw}\" × {sh}\"")
    print(f"{'='*60}\n")
    print(f"{'ID':<5} {'名称':<22} {'类型':<18} {'左':<7} {'上':<7} {'宽':<7} {'高':<7}  文字预览")
    print("-" * 95)

    for s in slide.shapes:
        l = emu_to_in(s.left); t = emu_to_in(s.top)
        w = emu_to_in(s.width); h = emu_to_in(s.height)
        txt = (s.text.replace("\n"," ")[:35] if hasattr(s,"text") else "")
        print(f"{s.shape_id:<5} {s.name[:20]:<22} {str(s.shape_type)[:16]:<18} "
              f"{l:<7} {t:<7} {w:<7} {h:<7}  {txt}")

    # ── Auto-detect title and content boxes ──────────────────────────
    title_shape   = None
    content_shape = None
    deco_shapes   = []

    for s in slide.shapes:
        has_text = hasattr(s, "text")
        l = emu_to_in(s.left); t = emu_to_in(s.top)
        w = emu_to_in(s.width); h = emu_to_in(s.height)

        if has_text and t < 1.5 and h < 1.5 and w > 4:
            if title_shape is None or t < emu_to_in(title_shape.top):
                title_shape = s
        elif has_text and t >= 0.8 and h > 1.5 and w > 4:
            if content_shape is None or h > emu_to_in(content_shape.height):
                content_shape = s
        else:
            deco_shapes.append(s)

    # Slide number: near bottom
    slide_num_shape = None
    for s in deco_shapes:
        if emu_to_in(s.top) > sh * 0.8:
            slide_num_shape = s
            break

    print(f"\n{'─'*60}")
    print("自动识别结果：")
    print(f"  标题框  → {'Shape ' + str(title_shape.shape_id) + ' ' + repr(title_shape.name) if title_shape else '⚠️  未识别'}")
    print(f"  内容框  → {'Shape ' + str(content_shape.shape_id) + ' ' + repr(content_shape.name) if content_shape else '⚠️  未识别'}")
    print(f"  装饰元素→ {len(deco_shapes)} 个")

    # ── Build config dict ─────────────────────────────────────────────
    cfg = {
        "_note": "此文件由 inspect_template.py 自动生成，可手动修改后重新运行转换。",
        "slide_w": sw,
        "slide_h": sh,
    }

    if title_shape:
        cfg.update({
            "title_l": emu_to_in(title_shape.left),
            "title_t": emu_to_in(title_shape.top),
            "title_w": emu_to_in(title_shape.width),
            "title_h": emu_to_in(title_shape.height),
        })
    else:
        print("\n  ⚠️  未能自动识别标题框，请在生成的 JSON 中手动填写 title_l/t/w/h")

    if content_shape:
        cfg.update({
            "content_l": emu_to_in(content_shape.left),
            "content_t": emu_to_in(content_shape.top),
            "content_w": emu_to_in(content_shape.width),
            "content_h": emu_to_in(content_shape.height),
        })
    else:
        print("  ⚠️  未能自动识别内容框，请在生成的 JSON 中手动填写 content_l/t/w/h")

    if slide_num_shape:
        cfg.update({
            "slide_num_l": emu_to_in(slide_num_shape.left),
            "slide_num_t": emu_to_in(slide_num_shape.top),
            "slide_num_w": emu_to_in(slide_num_shape.width),
            "slide_num_h": emu_to_in(slide_num_shape.height),
        })

    cfg.update({
        "copy_deco": True,
        "font_main": "Microsoft YaHei",
        "title_color": "00B0F0",
        "h3_color": "005A87",
    })

    # ── Print config suggestion ───────────────────────────────────────
    print(f"\n{'─'*60}")
    print("生成的 JSON 配置：")
    print(json.dumps(cfg, ensure_ascii=False, indent=2))

    # ── Write JSON if requested ───────────────────────────────────────
    if json_out:
        Path(json_out).write_text(
            json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\n✅ 配置已写入：{json_out}")

    print(f"\n{'='*60}")
    print("提示：如识别有误，直接编辑 JSON 文件中的数值再重新运行转换即可。")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="分析 PPTX 模板结构")
    parser.add_argument("template", help="PPTX 模板文件路径")
    parser.add_argument("--json", "-j", default=None,
                        help="将配置输出到此 JSON 文件路径（可选）")
    args = parser.parse_args()
    inspect(args.template, json_out=args.json)
