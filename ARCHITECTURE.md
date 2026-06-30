# mcopilot — Architecture

## Overview

```
┌─────────────────────────────────────────────────────────────────┐
│  macOS (Terminal — Homebrew green profile)                       │
│                                                                   │
│  ┌────────────┐   OpenAI chat/completions   ┌────────────────┐  │
│  │  OpenCode  │ ─────────────────────────► │   shim.py      │  │
│  │  (TUI)     │ ◄──────────────────────── │   :8765/v1     │  │
│  └────────────┘   SSE chunks / JSON resp   └───────┬────────┘  │
│                                                      │           │
│                                             egress guard         │
│                                         (Microsoft only)         │
│                                                      │           │
└──────────────────────────────────────────────────────┼───────────┘
                                                        │ HTTPS
                                               ┌────────▼────────┐
                                               │  graph.microsoft │
                                               │  .com/beta/      │
                                               │  copilot         │
                                               └────────┬─────────┘
                                                        │
                                               ┌────────▼─────────┐
                                               │  M365 Copilot    │
                                               │  (enterprise     │
                                               │   grounding)     │
                                               │                  │
                                               │ Mail  Teams  SP  │
                                               └──────────────────┘
```

## Components

### `shim.py` — OpenAI-compatible API server

A FastAPI + uvicorn server that translates OpenAI `chat/completions` requests
into M365 Copilot Chat API calls.

**Key responsibilities:**
- Egress guard (installed before any import): wraps `socket.getaddrinfo` to
  block every non-Microsoft hostname
- Token management: `_shim_token()` tries in-memory cache → disk reload →
  MSAL silent refresh; raises on expiry (caller gets 401 + re-auth hint)
- Message translation: OpenAI `messages[]` → Copilot `conversations/{id}/chat`
- Session continuity: SHA-256 of conversation history maps to Copilot
  conversation IDs, persisted in `~/.cache/mcopilot/sessions.json`
- Streaming: real `chatOverStream` SSE (revision-safe delta logic) with
  pseudo-stream fallback
- Utility short-circuit: OpenCode's title-generator calls are answered locally
  (~0ms vs ~12s Copilot round-trip)

**CLI modes:**
```
./shim.py --ensure-token   # silent refresh or SSO popup; used by mcopilot alias
./shim.py                  # start server
./shim.py --login          # device-code fallback (headless envs)
./shim.py --selftest       # verify egress guard, no network
```

### `m365_auth.py` — Auth + Graph API helpers

Standalone module (no external MCP dependency). Provides:
- MSAL `PublicClientApplication` using the Graph CLI public client
  (`14d82eec-...`) — first-party, no app registration required
- `SerializableTokenCache` stored at `~/.cache/mcopilot/token_cache.json`
- `token_silent()` — MSAL cache lookup + silent refresh
- `token_interactive(login_hint)` — SSO browser popup, uses Entra session
- `post()` / `stream_post()` — Graph API calls with token injection
- `shape_answer()` — extracts text + citations from `copilotConversation`

### `setup.sh` — One-shot interactive setup

Forces the user to explicitly input their Microsoft 365 email address before
any auth attempt. This is a hard requirement: the email is the login hint that
ensures SSO authenticates the correct Entra identity rather than a personal
Microsoft account or a stale cached session.

Setup flow:
1. Validate macOS + Python 3.11+ + uv
2. **Prompt for email** (validates format; requires confirmation)
3. Save to `~/.config/mcopilot/config.json`
4. Install Python deps via `uv pip install`
5. Run `./shim.py --ensure-token` (browser SSO popup)
6. Optionally create Desktop icon

## Auth flow in detail

```
setup.sh
  │
  ├─ user types: you@company.com
  │    saved → ~/.config/mcopilot/config.json
  │
  └─ ./shim.py --ensure-token
       │
       ├─ _shim_token() → try token_silent()
       │    → reload MSAL cache from disk
       │    → _app.acquire_token_silent(SCOPES)
       │
       └─ RuntimeError (expired/missing)
            │
            └─ _do_sso_login(login_hint="you@company.com")
                 │
                 └─ _app.acquire_token_interactive(SCOPES, login_hint=email)
                       browser window opens; existing Entra SSO session
                       → refresh token written to ~/.cache/mcopilot/token_cache.json

[later, when shim.py server runs]

shim.py
  └─ _shim_token()
       ├─ in-memory cache hit → return
       ├─ auth.reload_cache() → re-read disk (catches --ensure-token writes)
       └─ token_silent() → MSAL refresh (silent, ~90-day token lifetime)
```

## Scopes (read-only)

| Scope | Used for |
|---|---|
| `Sites.Read.All` | SharePoint grounding |
| `Mail.Read` | Email context |
| `People.Read.All` | People/org data |
| `OnlineMeetingTranscript.Read.All` | Meeting transcripts |
| `Chat.Read` | Teams messages |
| `ChannelMessage.Read.All` | Teams channel content |
| `ExternalItem.Read.All` | Connected external content |

`Mail.Send` is excluded. Requesting it triggers admin-consent walls on most
tenants, and this tool never sends mail.

## Data boundary

The egress guard wraps `socket.getaddrinfo` at import time. Allowed suffixes:

```
*.microsoft.com
*.microsoftonline.com
*.microsoftonline-p.com
*.windows.net
*.msauth.net  *.msftauth.net
*.office.com  *.cloud.microsoft
localhost / loopback
```

Any other hostname → `RuntimeError` (connection never established).
Every successful resolution is appended to `~/.cache/mcopilot/egress.log`.
Run `./shim.py --selftest` to verify the guard is active.

## Multi-process token sharing

`--ensure-token` runs in a separate process (spawned by the shell alias before
OpenCode starts). The MSAL `SerializableTokenCache` is an in-memory object;
the `--ensure-token` process writes a fresh token to disk, then exits. When
the shim server calls `_shim_token()` and the in-memory cache is stale, it
calls `auth.reload_cache()` to re-read the disk file before attempting silent
refresh. This ensures the two processes share the token without IPC.

## Conversation continuity

OpenCode (like all OpenAI clients) sends the full conversation history on every
request. The shim hashes the history prefix (everything except the latest user
turn) with SHA-256 and maps it to a Copilot conversation ID. After each
response, it re-keys the map including the assistant's reply, so the next
request (which will include it) resolves to the same conversation. Persisted
to `~/.cache/mcopilot/sessions.json` (no secrets — only opaque IDs).
