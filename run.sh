#!/bin/bash
# run.sh — Convert Markdown to PPTX
#
# Usage:
#   ./run.sh input.md                              # output: input.pptx (same dir)
#   ./run.sh input.md output.pptx                  # explicit output path
#   ./run.sh input.md output.pptx template.pptx    # explicit template
#
# Positional arguments:
#   $1  input .md file        (required)
#   $2  output .pptx file     (optional; defaults to <input>.pptx in same dir)
#   $3  template .pptx file   (optional; defaults to template.pptx next to this script)

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CALL_DIR="$(pwd)"

# ── 参数检查 ─────────────────────────────────────────────────────────
if [ -z "$1" ]; then
  echo "用法：./run.sh input.md [output.pptx] [template.pptx]"
  exit 1
fi

# ── 解析路径（相对路径转绝对路径，相对于调用目录）──────────────────
resolve() {
  local p="$1"
  [[ "$p" = /* ]] && echo "$p" || echo "$CALL_DIR/$p"
}

MD_FILE="$(resolve "$1")"

if [ ! -f "$MD_FILE" ]; then
  echo "❌ 找不到输入文件：$MD_FILE"
  exit 1
fi

# ── 确定输出路径 ─────────────────────────────────────────────────────
if [ -n "$2" ]; then
  OUTPUT="$(resolve "$2")"
else
  # Default: same directory as the .md file, same base name
  MD_DIR="$(dirname "$MD_FILE")"
  MD_BASE="$(basename "$MD_FILE" .md)"
  OUTPUT="$MD_DIR/${MD_BASE}.pptx"
fi

# ── 确定模板路径 ─────────────────────────────────────────────────────
USING_CUSTOM_TEMPLATE=false

if [ -n "$3" ]; then
  TEMPLATE="$(resolve "$3")"
  if [ ! -f "$TEMPLATE" ]; then
    echo "❌ 找不到模板文件：$TEMPLATE"
    exit 1
  fi
  USING_CUSTOM_TEMPLATE=true
else
  if [ -f "$SCRIPT_DIR/template.pptx" ]; then
    TEMPLATE="$SCRIPT_DIR/template.pptx"
  elif [ -f "$CALL_DIR/template.pptx" ]; then
    TEMPLATE="$CALL_DIR/template.pptx"
  else
    echo "❌ 未找到 template.pptx。"
    echo "   请将模板放在 $SCRIPT_DIR/ 或当前目录，或通过第三个参数指定。"
    exit 1
  fi
fi

# ── 激活虚拟环境 ─────────────────────────────────────────────────────
[ -f "$SCRIPT_DIR/env/bin/activate" ] && source "$SCRIPT_DIR/env/bin/activate"

# ── 自定义模板：自动 inspect 并生成 .json 配置 ───────────────────────
TEMPLATE_JSON="${TEMPLATE%.pptx}.json"

if [ "$USING_CUSTOM_TEMPLATE" = true ]; then
  if [ ! -f "$TEMPLATE_JSON" ]; then
    echo "🔍 检测新模板，生成配置文件..."
    python3 "$SCRIPT_DIR/inspect_template.py" "$TEMPLATE" --json "$TEMPLATE_JSON"
    echo "✅ 配置已保存到：$TEMPLATE_JSON"
  else
    echo "📋 使用已有配置：$TEMPLATE_JSON"
  fi
fi

# ── 转换 ─────────────────────────────────────────────────────────────
echo "📄 输入：$MD_FILE"
echo "🎨 模板：$TEMPLATE"
echo "💾 输出：$OUTPUT"
echo ""

CONFIG_ARG=""
[ -f "$TEMPLATE_JSON" ] && CONFIG_ARG="--config $TEMPLATE_JSON"

python3 "$SCRIPT_DIR/md_to_pptx.py" "$MD_FILE" "$TEMPLATE" "$OUTPUT" $CONFIG_ARG

echo ""
echo "✅ 完成：$OUTPUT"
