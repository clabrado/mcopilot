#!/usr/bin/env python3
"""Standalone M365 Copilot auth + Graph Copilot API helpers.

Uses Microsoft's Graph CLI public client (14d82eec-...) — a pre-authorized
first-party public client that supports delegated Graph access. No custom
Entra app registration required.

Scopes: Chat API read scopes only. Mail.Send is intentionally excluded to
avoid the admin-consent wall most tenants apply to write scopes.
"""
import atexit
import base64
import json
import os
import time
from pathlib import Path

try:
    import msal
    from msal import PublicClientApplication, SerializableTokenCache
except ImportError as e:
    raise SystemExit(f"[m365-auth] msal not installed: {e}\n  → uv pip install msal") from e

try:
    import httpx
except ImportError as e:
    raise SystemExit(f"[m365-auth] httpx not installed: {e}\n  → uv pip install httpx") from e

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
GRAPH_RESOURCE = "https://graph.microsoft.com"
BASE = "https://graph.microsoft.com/beta/copilot"

# Microsoft Graph CLI public client — first-party, no registration needed.
# Same client as `az` and `graph` CLI tools; pre-authorized for delegated Graph.
GRAPH_CLI_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"

TENANT = os.environ.get("AZURE_TENANT_ID", "organizations")
CLIENT = os.environ.get("M365COPILOT_CLIENT_ID", GRAPH_CLI_CLIENT_ID)
STATIC_TOKEN: str | None = os.environ.get("M365COPILOT_ACCESS_TOKEN") or None
DEFAULT_TZ = os.environ.get("M365COPILOT_TIMEZONE", "America/New_York")
EXPIRY_MARGIN_SEC = 60

# Chat API read scopes. Mail.Send excluded — this shim is read-only Q&A.
CHAT_READ_SCOPES = [
    "Sites.Read.All",
    "Mail.Read",
    "People.Read.All",
    "OnlineMeetingTranscript.Read.All",
    "Chat.Read",
    "ChannelMessage.Read.All",
    "ExternalItem.Read.All",
]
SCOPES = [f"{GRAPH_RESOURCE}/{s}" for s in CHAT_READ_SCOPES]

# Token cache — shared path so the CLI setup tool and the shim server can
# exchange tokens without a new interactive login.
CACHE_PATH = Path(
    os.environ.get("M365COPILOT_CACHE_PATH",
                   str(Path.home() / ".cache" / "mcopilot" / "token_cache.json"))
)
CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)

_msal_cache = SerializableTokenCache()
if CACHE_PATH.exists():
    _msal_cache.deserialize(CACHE_PATH.read_text())

_app = PublicClientApplication(
    CLIENT,
    authority=f"https://login.microsoftonline.com/{TENANT}",
    token_cache=_msal_cache,
)

# ---------------------------------------------------------------------------
# JWT helpers
# ---------------------------------------------------------------------------

def decode_jwt(token: str) -> dict:
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload))
    except Exception:
        return {}


def is_valid(token: str | None) -> bool:
    return bool(token) and time.time() + EXPIRY_MARGIN_SEC < int(decode_jwt(token).get("exp", 0))


# ---------------------------------------------------------------------------
# Token cache persistence
# ---------------------------------------------------------------------------

def persist_cache() -> None:
    if _msal_cache.has_state_changed:
        CACHE_PATH.write_text(_msal_cache.serialize())
        CACHE_PATH.chmod(0o600)


atexit.register(persist_cache)


def reload_cache() -> None:
    """Re-read the token cache from disk. Call this when another process may
    have written a fresh token (e.g., the setup --ensure-token step)."""
    if CACHE_PATH.exists():
        _msal_cache.deserialize(CACHE_PATH.read_text())


# ---------------------------------------------------------------------------
# Token acquisition
# ---------------------------------------------------------------------------

def token_silent() -> str | None:
    """Try silent token acquisition from the MSAL cache (no user interaction)."""
    reload_cache()
    accts = _app.get_accounts()
    if not accts:
        return None
    res = _app.acquire_token_silent(SCOPES, account=accts[0])
    if res and "access_token" in res:
        persist_cache()
        return res["access_token"]
    return None


def token_interactive(login_hint: str | None = None) -> str:
    """SSO browser popup — reuses an existing Entra/Azure AD session.

    No device code, no manual URL entry. Opens a browser window; the user
    clicks through if already signed in (often zero clicks for SSO).

    Args:
        login_hint: The user's Entra email. Pre-fills the sign-in page and
            ensures the correct identity is used. Captured during setup and
            stored in ~/.config/mcopilot/config.json.
    """
    kwargs: dict = {}
    if login_hint:
        kwargs["login_hint"] = login_hint
    elif (accts := _app.get_accounts()):
        kwargs["login_hint"] = accts[0].get("username", "")
    res = _app.acquire_token_interactive(SCOPES, **kwargs)
    if "access_token" not in res:
        raise RuntimeError(res.get("error_description", str(res)))
    persist_cache()
    return res["access_token"]


# ---------------------------------------------------------------------------
# Graph API helpers
# ---------------------------------------------------------------------------

def _headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def post(path: str, body: dict, token: str, timeout: float = 120) -> dict:
    r = httpx.post(f"{BASE}{path}", headers=_headers(token), json=body, timeout=timeout)
    if r.status_code == 403:
        claims = decode_jwt(token)
        raise RuntimeError(
            f"403 Forbidden from Copilot Chat API. "
            f"Account: {claims.get('upn', '?')}. "
            f"Scopes: {claims.get('scp', '(none)')}. "
            f"Ensure the account has a Microsoft 365 Copilot license. "
            f"Server: {r.text[:300]}"
        )
    r.raise_for_status()
    return r.json() if r.content else {}


def stream_post(path: str, body: dict, token: str, timeout: float = 180):
    """Context manager for SSE streaming from the Copilot API."""
    return httpx.stream(
        "POST", f"{BASE}{path}",
        headers={**_headers(token), "Accept": "text/event-stream"},
        json=body, timeout=timeout,
    )


def shape_answer(convo: dict) -> dict:
    """Extract answer text and citations from a copilotConversation response."""
    messages = convo.get("messages") or []
    answer_msg = messages[-1] if messages else {}
    citations = [
        {
            "type": a.get("attributionType"),
            "title": a.get("providerDisplayName"),
            "url": a.get("seeMoreWebUrl"),
            "source": a.get("attributionSource"),
        }
        for a in (answer_msg.get("attributions") or [])
    ]
    return {
        "conversation_id": convo.get("id"),
        "answer": answer_msg.get("text", ""),
        "citations": citations,
        "turn_count": convo.get("turnCount"),
    }
