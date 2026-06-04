#!/bin/bash
# setup.sh — Create virtualenv and install dependencies
set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# ── Python 版本检查（需要 3.10+）────────────────────────────────────
python3 -c "
import sys
if sys.version_info < (3, 10):
    print(f'❌ 需要 Python 3.10 或更新版本，当前版本：{sys.version}')
    sys.exit(1)
print(f'✅ Python {sys.version_info.major}.{sys.version_info.minor} 检测通过')
"

echo "Creating virtual environment..."
python3 -m venv env
echo "Installing dependencies..."
env/bin/pip install --upgrade pip -q
env/bin/pip install python-pptx lxml chardet -q
echo ""
echo "Done. To convert: bash run.sh input.md"
