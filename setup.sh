#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "${BASH_SOURCE[0]}")"

VENV_PYTHON=".venv/bin/python"

if [ ! -x "$VENV_PYTHON" ]; then
    echo "Creating local Python environment in .venv..."

    if command -v python3.12 >/dev/null 2>&1; then
        python3.12 -m venv .venv
    elif command -v python3 >/dev/null 2>&1; then
        python3 -m venv .venv
    else
        echo "No python3 found. Install Python 3 (e.g. 'brew install python') and re-run this script." >&2
        exit 1
    fi
else
    echo "Using existing .venv."
fi

echo "Installing project dependencies..."
"$VENV_PYTHON" -m pip install -r requirements.txt

echo ""
echo "Setup complete."
echo "In VS Code, open this folder and use the integrated terminal."
echo "Start the local frontend:"
echo "./.venv/bin/python frontend_server.py"
echo ""
echo "Note: WeCom / WhatsApp paste automation and group-chat detection (paste_sender.py)"
echo "are Windows-only for now and are disabled in the frontend on macOS."
