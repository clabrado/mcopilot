#!/usr/bin/env bash
# Creates mcopilot.app on the Desktop — a green-terminal shortcut that
# launches the mcopilot alias on click.
set -euo pipefail

REPO_DIR="${1:-$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)}"
APP_NAME="mcopilot"
DEST="$HOME/Desktop/${APP_NAME}.app"
OC_CONFIG="$REPO_DIR/opencode.json"
OPENCODE_BIN="$REPO_DIR/bin/opencode"

# ---------------------------------------------------------------------------
# Determine the shim command — use the mcopilot alias if already in .zshrc,
# otherwise build the command from the repo path.
# ---------------------------------------------------------------------------
SHIM_CMD="cd '$REPO_DIR' && ./shim.py --ensure-token && OPENCODE_CONFIG='$OC_CONFIG' '$OPENCODE_BIN'"

# ---------------------------------------------------------------------------
# Compile the AppleScript app with the "Homebrew" (green/black) profile.
# ---------------------------------------------------------------------------
SCRIPT=$(cat <<ASEOF
tell application "Terminal"
    set newTab to do script "source ~/.zshrc; $SHIM_CMD"
    delay 0.3
    set current settings of newTab to settings set "Homebrew"
    activate
end tell
ASEOF
)

rm -rf "$DEST"
osacompile -o "$DEST" - <<< "$SCRIPT"

echo "Created: $DEST"
echo "Tip: Drag it to your Dock to add a toolbar shortcut."
