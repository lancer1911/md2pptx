#!/bin/bash
# setup.sh — Create virtualenv and install dependencies
set -e
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"
echo "Creating virtual environment..."
python3 -m venv env
echo "Installing dependencies..."
env/bin/pip install --upgrade pip -q
env/bin/pip install python-pptx lxml chardet -q
echo ""
echo "Done. To convert: bash run.sh input.md"
