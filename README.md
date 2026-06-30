# mcopilot

M365 Copilot in your terminal via [OpenCode](https://opencode.ai).

Wraps the Microsoft 365 Copilot Chat API (`/beta/copilot`) behind an
OpenAI-compatible endpoint on localhost so OpenCode can talk to it like any
other model. Every connection is validated against a Microsoft-only allowlist —
no data reaches Anthropic, OpenAI, or any third party.

```
you (terminal)
    └─ OpenCode
         └─ http://127.0.0.1:8765/v1  (this shim)
              └─ M365 Copilot Chat API  (graph.microsoft.com)
                   └─ your enterprise data (Mail, Teams, SharePoint, People)
```

**macOS only.** Requires Python 3.11+, `uv`, and an OpenCode binary.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| macOS 13+ | AppleScript desktop icon uses macOS Terminal.app |
| Python 3.11+ | Check: `python3 --version` |
| `uv` | `curl -LsSf https://astral.sh/uv/install.sh \| sh` |
| OpenCode binary | [opencode.ai](https://opencode.ai) — download and place in `bin/opencode` |
| M365 Copilot license | The account must have a Microsoft 365 Copilot add-on license |

---

## Setup

```bash
git clone https://github.com/clabrado/mcopilot.git
cd mcopilot

# Download OpenCode and link it
curl -fsSL https://opencode.ai/install.sh | sh
mkdir -p bin && ln -sf $(which opencode) bin/opencode

# Run setup (prompts for your email, installs deps, authenticates)
bash setup.sh
```

The setup script will:

1. **Ask for your Microsoft 365 email address.** This is required — it's used
   as a login hint to ensure SSO uses the correct Entra identity. You must
   type it explicitly; it cannot be inferred.

2. Install Python dependencies (`msal`, `httpx`, `fastapi`, `uvicorn`).

3. Open a browser SSO window. Sign in with the email you entered. If you're
   already signed in via corporate SSO, this often completes with zero clicks.

4. Optionally create a Desktop shortcut (green Homebrew terminal).

---

## Usage

After setup, launch with:

```bash
mcopilot
```

This:
1. Silently refreshes the M365 token (or reopens SSO if expired)
2. Starts the shim server on `http://127.0.0.1:8765/v1`
3. Launches OpenCode pointed at the shim

Inside OpenCode, all queries go to M365 Copilot grounded in your enterprise
data. Web search is off by default (internal data only).

### Re-authenticate

```bash
cd ~/path/to/mcopilot
./shim.py --ensure-token
```

Silently refreshes if the token is still valid. If expired, opens a browser
SSO popup. No device codes, no manual URLs.

### Verify

```bash
curl http://127.0.0.1:8765/healthz
```

---

## Configuration

| Env var | Default | Description |
|---|---|---|
| `M365_SHIM_PORT` | `8765` | Port for the local shim server |
| `M365_SHIM_STREAM` | `pseudo` | `pseudo` / `real` / `auto` streaming mode |
| `M365_SHIM_WEB` | `0` | Set `1` to enable web grounding by default |
| `M365_SHIM_STYLE` | *(terse prose)* | System tone directive sent to Copilot |
| `M365_SHIM_CITATIONS` | `0` | Set `1` to always include source links |
| `M365_SHIM_FORWARD_SYSTEM` | `0` | Set `1` to forward OpenCode's system prompt |
| `AZURE_TENANT_ID` | `organizations` | Your Azure AD tenant ID (multi-tenant default) |
| `M365COPILOT_CACHE_PATH` | `~/.cache/mcopilot/token_cache.json` | MSAL token cache path |
| `M365COPILOT_ACCESS_TOKEN` | *(none)* | Override: use this bearer token directly |

Username (login hint) is stored in `~/.config/mcopilot/config.json` by `setup.sh`.

---

## Auth flow

```
setup.sh  →  shim.py --ensure-token  →  browser SSO popup
                                              ↓
                                    ~/.cache/mcopilot/token_cache.json
                                              ↓
shim.py (server)  →  _shim_token()  →  msal silent refresh
                                          (re-read cache from disk)
                                              ↓
                                    POST /beta/copilot/conversations/{id}/chat
```

The shim requests **7 read scopes** only — `Sites.Read.All`, `Mail.Read`,
`People.Read.All`, `OnlineMeetingTranscript.Read.All`, `Chat.Read`,
`ChannelMessage.Read.All`, `ExternalItem.Read.All`. `Mail.Send` is excluded to
avoid admin-consent walls.

Auth uses the **Microsoft Graph CLI public client** (`14d82eec-...`), the same
pre-authorized first-party client used by `az` and `msgraph` CLI tools. No
custom Entra app registration is required.

---

## Data boundary

The egress guard (installed before any import) wraps `socket.getaddrinfo` and
blocks every hostname outside the Microsoft allowlist. Any attempt to contact
`api.anthropic.com`, `api.openai.com`, or any non-Microsoft endpoint raises
immediately. Every resolved hostname is also logged to
`~/.cache/mcopilot/egress.log`.

Run `./shim.py --selftest` to verify the boundary holds.

---

## Files

| File | Purpose |
|---|---|
| `shim.py` | OpenAI-compatible server + CLI entry points |
| `m365_auth.py` | Standalone MSAL auth + Graph API helpers |
| `setup.sh` | One-shot interactive setup |
| `create-desktop-icon.sh` | Creates `mcopilot.app` on the Desktop |
| `opencode.json` | OpenCode provider config (points at the shim) |
| `bin/opencode` | OpenCode binary (gitignored — download separately) |

---

## Architecture

See [ARCHITECTURE.md](ARCHITECTURE.md) for a detailed design overview.
