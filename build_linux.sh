#!/usr/bin/env bash
# Build helper for Linux (cannot produce a native Windows .exe reliably).
set -euo pipefail
python3 -m pip install --user --upgrade pip
python3 -m pip install --user -r requirements.txt
python3 -m pip install --user pyinstaller

# Build a one-file Linux executable
pyinstaller --noconfirm --clean --onefile --name dynamic_payroll main.py

echo "Build complete. Output in dist/"
