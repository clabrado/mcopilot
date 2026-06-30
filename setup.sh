#!/usr/bin/env bash
# mcopilot setup — M365 Copilot via OpenCode on macOS
# Prompts for your Entra username and wires SSO + desktop icon.
set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CONFIG_DIR="$HOME/.config/mcopilot"
CONFIG_FILE="$CONFIG_DIR/config.json"
CACHE_DIR="$HOME/.cache/mcopilot"

# ── ANSI colors ──────────────────────────────────────────────────────────────
green()  { printf '\033[0;32m%s\033[0m\n' "$*"; }
yellow() { printf '\033[0;33m%s\033[0m\n' "$*"; }
red()    { printf '\033[0;31m%s\033[0m\n' "$*"; }
bold()   { printf '\033[1m%s\033[0m\n' "$*"; }

bold "═══════════════════════════════════════════"
bold "  mcopilot — M365 Copilot + OpenCode Setup "
bold "═══════════════════════════════════════════"
echo ""

# ── macOS check ──────────────────────────────────────────────────────────────
if [[ "$(uname -s)" != "Darwin" ]]; then
    red "ERROR: mcopilot requires macOS. Detected: $(uname -s)"
    exit 1
fi

# ── Prerequisites check ───────────────────────────────────────────────────────
echo "Checking prerequisites..."

MISSING=()

if ! command -v python3 &>/dev/null; then
    MISSING+=("python3 (install via https://python.org or 'brew install python@3.12')")
fi

PYTHON_VER=$(python3 -c 'import sys; print(sys.version_info >= (3, 11))' 2>/dev/null || echo "False")
if [[ "$PYTHON_VER" != "True" ]]; then
    MISSING+=("python3 >= 3.11 (current: $(python3 --version 2>&1))")
fi

if ! command -v uv &>/dev/null; then
    MISSING+=("uv (install via 'curl -LsSf https://astral.sh/uv/install.sh | sh')")
fi

if [[ ${#MISSING[@]} -gt 0 ]]; then
    red "Missing prerequisites:"
    for item in "${MISSING[@]}"; do
        echo "  • $item"
    done
    exit 1
fi

green "Prerequisites OK"
echo ""

# ── Username (REQUIRED) ───────────────────────────────────────────────────────
bold "Microsoft 365 Account"
echo ""
echo "Enter your Microsoft 365 / Entra email address."
echo "This is used as a login hint so SSO uses the correct identity."
echo "For Dynatrace: your.name@dynatrace.com"
echo ""

USERNAME=""
while [[ -z "$USERNAME" ]]; do
    read -rp "Email address: " USERNAME
    USERNAME="$(echo "$USERNAME" | tr '[:upper:]' '[:lower:]' | xargs)"

    if [[ -z "$USERNAME" ]]; then
        yellow "Email address is required."
        USERNAME=""
        continue
    fi

    if [[ "$USERNAME" != *"@"*"."* ]]; then
        yellow "Invalid format. Must be a valid email (e.g., you@company.com)."
        USERNAME=""
        continue
    fi

    echo ""
    echo "  Username: $USERNAME"
    read -rp "Correct? [y/n]: " CONFIRM
    echo ""
    if [[ "$CONFIRM" != "y" && "$CONFIRM" != "Y" ]]; then
        USERNAME=""
        yellow "Let's try again."
        echo ""
    fi
done

# ── Save config ───────────────────────────────────────────────────────────────
mkdir -p "$CONFIG_DIR" "$CACHE_DIR"
cat > "$CONFIG_FILE" <<EOF
{
  "username": "$USERNAME",
  "repo_dir": "$REPO_DIR"
}
EOF
chmod 600 "$CONFIG_FILE"
green "Config saved → $CONFIG_FILE"
echo ""

# ── Python dependencies ───────────────────────────────────────────────────────
echo "Installing Python dependencies..."
uv pip install --quiet msal httpx fastapi uvicorn 2>&1 | tail -3
green "Dependencies installed"
echo ""

# ── OpenCode binary ───────────────────────────────────────────────────────────
OPENCODE_BIN="$REPO_DIR/bin/opencode"

if [[ ! -x "$OPENCODE_BIN" ]]; then
    yellow "OpenCode binary not found at $OPENCODE_BIN"
    echo ""
    echo "Install OpenCode (https://opencode.ai):"
    echo "  curl -fsSL https://opencode.ai/install.sh | sh"
    echo ""
    echo "Then copy or symlink the binary:"
    echo "  mkdir -p $REPO_DIR/bin"
    echo "  ln -sf \$(which opencode) $REPO_DIR/bin/opencode"
    echo ""
    read -rp "OpenCode binary path (or Enter to skip): " OC_PATH
    if [[ -n "$OC_PATH" && -x "$OC_PATH" ]]; then
        mkdir -p "$REPO_DIR/bin"
        ln -sf "$OC_PATH" "$OPENCODE_BIN"
        green "Linked: $OC_PATH → $OPENCODE_BIN"
    else
        yellow "Skipping OpenCode link. Add it later: ln -sf \$(which opencode) $REPO_DIR/bin/opencode"
    fi
    echo ""
fi

# ── opencode.json ─────────────────────────────────────────────────────────────
OC_CONFIG="$REPO_DIR/opencode.json"
if [[ ! -f "$OC_CONFIG" ]]; then
    cat > "$OC_CONFIG" <<'OCEOF'
{
  "$schema": "https://opencode.ai/config.json",
  "enabled_providers": ["mcopilot"],
  "model": "mcopilot/m365-copilot",
  "provider": {
    "mcopilot": {
      "npm": "@ai-sdk/openai-compatible",
      "name": "M365 Copilot",
      "options": {
        "baseURL": "http://127.0.0.1:8765/v1",
        "apiKey": "not-needed"
      },
      "models": {
        "m365-copilot": { "name": "M365 Copilot (grounded)" }
      }
    }
  }
}
OCEOF
    green "Created opencode.json"
fi

# ── Shell alias ───────────────────────────────────────────────────────────────
SHELL_RC="$HOME/.zshrc"
ALIAS_LINE="alias mcopilot='(cd $REPO_DIR && ./shim.py --ensure-token && OPENCODE_CONFIG=$OC_CONFIG $OPENCODE_BIN)'"

if ! grep -qF "alias mcopilot=" "$SHELL_RC" 2>/dev/null; then
    echo "" >> "$SHELL_RC"
    echo "# mcopilot — M365 Copilot via OpenCode" >> "$SHELL_RC"
    echo "$ALIAS_LINE" >> "$SHELL_RC"
    green "Added 'mcopilot' alias to $SHELL_RC"
else
    yellow "mcopilot alias already in $SHELL_RC (not modified)"
fi
echo ""

# ── First-time auth ───────────────────────────────────────────────────────────
bold "First-time sign-in"
echo ""
echo "A browser window will open. Sign in with:"
echo "  $USERNAME"
echo ""
echo "If you're already signed in via SSO, this may complete automatically."
echo ""

python3 "$REPO_DIR/shim.py" --ensure-token
echo ""
green "Authentication successful"
echo ""

# ── Desktop icon ─────────────────────────────────────────────────────────────
read -rp "Create Desktop shortcut (green terminal, auto-launches mcopilot)? [y/n]: " CREATE_ICON
if [[ "$CREATE_ICON" == "y" || "$CREATE_ICON" == "Y" ]]; then
    bash "$REPO_DIR/create-desktop-icon.sh" "$REPO_DIR"
fi

# ── Done ──────────────────────────────────────────────────────────────────────
echo ""
bold "═══════════════════════════════"
green "  Setup complete!"
bold "═══════════════════════════════"
echo ""
echo "  Start:   mcopilot    (after reloading shell: source ~/.zshrc)"
echo "  Re-auth: cd $REPO_DIR && ./shim.py --ensure-token"
echo "  Status:  curl http://127.0.0.1:8765/healthz"
echo ""
