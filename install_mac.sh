#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# NJN Comparator — macOS Setup & Launcher
# Double-click this file (or run it from Terminal) to install and open the app.
# ─────────────────────────────────────────────────────────────────────────────

set -e
DIR="$(cd "$(dirname "$0")" && pwd)"

echo ""
echo "  ╔══════════════════════════════════╗"
echo "  ║      NJN COMPARATOR SETUP        ║"
echo "  ╚══════════════════════════════════╝"
echo ""

# ── 1. Find or install Python 3 ──────────────────────────────────────────────
PYTHON=""
for candidate in python3 /usr/local/bin/python3 /opt/homebrew/bin/python3; do
  if command -v "$candidate" &>/dev/null && "$candidate" -c "import sys; sys.exit(0 if sys.version_info >= (3,9) else 1)" 2>/dev/null; then
    PYTHON="$candidate"
    break
  fi
done

if [ -z "$PYTHON" ]; then
  echo "  Python 3.9+ not found. Installing via Homebrew..."
  if ! command -v brew &>/dev/null; then
    echo "  Installing Homebrew first (you may be asked for your password)..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    eval "$(/opt/homebrew/bin/brew shellenv)" 2>/dev/null || true
    eval "$(/usr/local/bin/brew shellenv)"   2>/dev/null || true
  fi
  brew install python3
  PYTHON="$(brew --prefix)/bin/python3"
fi

echo "  ✓ Python: $PYTHON ($($PYTHON --version))"

# ── 2. Install pip packages ───────────────────────────────────────────────────
echo "  Installing required packages (first time only)..."
"$PYTHON" -m pip install --quiet --upgrade customtkinter openpyxl
echo "  ✓ Packages ready"

# ── 3. Launch ─────────────────────────────────────────────────────────────────
echo "  Launching NJN Comparator..."
echo ""
cd "$DIR"
"$PYTHON" app.py
