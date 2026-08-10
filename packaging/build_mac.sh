#!/usr/bin/env bash
# Builds the macOS .app bundle for PTA Treasurer.
#
# Builds in a fresh venv (not the dev .venv) with a non-editable `pip
# install .` — this catches "works in dev venv but missing from a real
# install" issues (e.g. a runtime-read file that isn't declared as
# package data) that an editable install would mask, and avoids
# PyInstaller's import analysis following editable-install .pth
# indirection.
#
# Usage: packaging/build_mac.sh
# Output: dist/PTA Treasurer.app

set -euo pipefail
cd "$(dirname "$0")/.."

BUILD_VENV=".venv-build-mac"

# Use the same interpreter as the dev .venv (>=3.10 required by
# pyproject.toml) rather than whatever `python3` resolves to on PATH --
# on this machine that's an unrelated Anaconda-provided Python 3.8.
BUILD_PYTHON="$(cd "$(dirname "$0")/.." && .venv/bin/python3 -c 'import sys; print(sys.executable)')"

rm -rf "$BUILD_VENV" build dist
"$BUILD_PYTHON" -m venv "$BUILD_VENV"
source "$BUILD_VENV/bin/activate"

pip install --upgrade pip
pip install .
pip install "pyinstaller==6.21.0" "pyinstaller-hooks-contrib"

pyinstaller packaging/pyinstaller_mac.spec --noconfirm --clean

deactivate
echo ""
echo "Built: dist/PTA Treasurer.app"
