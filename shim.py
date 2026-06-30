#!/usr/bin/env python3
"""M365 Copilot → OpenAI-compatible shim.

Starts an OpenAI-compatible HTTP server on localhost:8765. OpenCode (or any
OpenAI client) connects to it and talks to Microsoft 365 Copilot's Chat API
through the standard chat/completions endpoint.

Data boundary: every outbound TCP connection is validated against a Microsoft-
only allowlist. No Anthropic, no OpenAI, no third-party endpoint is contacted.

Usage:
    ./shim.py --ensure-token   # silent refresh or SSO popup, then exit
    ./shim.py                  # start the API server (OpenCode connects here)
    ./shim.py --login          # device-code sign-in (fallback)
    ./shim.py --selftest       # verify egress guard + imports, no network
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import re
import socket
import sys
import time
import uuid
from pathlib import Path
from typing import Any, Iterable

# ---------------------------------------------------------------------------
# 0. EGRESS ALLOWLIST — installed before any network-capable import.
#    Every outbound TCP connection resolves through socket.getaddrinfo; we
#    wrap it to permit ONLY Microsoft + loopback. This covers httpx AND msal.
# ---------------------------------------------------------------------------

_ALLOWED_SUFFIXES = (
    ".microsoft.com",
    ".microsoftonline.com",
    ".microsoftonline-p.com",
    ".windows.net",
    ".msauth.net",
    ".msftauth.net",
    ".office.com",
    ".cloud.microsoft",
)
_ALLOWED_EXACT = {
    "microsoft.com", "graph.microsoft.com", "login.microsoftonline.com",
    "localhost", "localhost.localdomain",
}
_LOOPBACK = {"127.0.0.1", "0.0.0.0", "::1", "::", ""}

EGRESS_LOG: list[str] = []
_EGRESS_LOG_PATH = Path(
    os.environ.get("M365_SHIM_EGRESS_LOG",
                   str(Path.home() / ".cache" / "mcopilot" / "egress.log"))
)


def _host_allowed(host: str | None) -> bool:
    if host is None:
        return True
    h = host.strip().lower().rstrip(".")
    if h in _LOOPBACK or h in _ALLOWED_EXACT:
        return True
    if any(h == s.lstrip(".") or h.endswith(s) for s in _ALLOWED_SUFFIXES):
        return True
    return False


def _record_egress(host: str | None) -> None:
    if not host:
        return
    h = str(host).lower().rstrip(".")
    if h in _LOOPBACK:
        return
    EGRESS_LOG.append(h)
    try:
        _EGRESS_LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with _EGRESS_LOG_PATH.open("a") as fh:
            fh.write(f"{int(time.time())} {h}\n")
    except Exception:
        pass


def _install_egress_guard() -> None:
    _orig = socket.getaddrinfo

    def _guarded(host, *args, **kwargs):
        if not _host_allowed(host):
            raise RuntimeError(
                f"EGRESS BLOCKED: '{host}' is outside the Microsoft data boundary. "
                "This shim contacts Microsoft endpoints + loopback only."
            )
        _record_egress(host)
        return _orig(host, *args, **kwargs)

    socket.getaddrinfo = _guarded  # type: ignore[assignment]


_install_egress_guard()

# Warn (don't crash) if an Anthropic key is present — the shim never uses it,
# but flagging it makes the "no Anthropic traffic" claim unambiguous.
for _k in ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN", "CLAUDE_API_KEY"):
    if os.environ.get(_k):
        print(f"[mcopilot] NOTE: {_k} is set — not used by this shim (no Anthropic calls).",
              file=sys.stderr)

# ---------------------------------------------------------------------------
# 1. Auth + Graph API layer
# ---------------------------------------------------------------------------
_HERE = Path(__file__).parent
sys.path.insert(0, str(_HERE))

try:
    import m365_auth as auth
except ImportError as e:
    raise SystemExit(f"[mcopilot] Cannot import m365_auth: {e}\n"
                     f"  Expected at: {_HERE / 'm365_auth.py'}") from e

import httpx  # noqa: E402 (after guard)

# ---------------------------------------------------------------------------
# 2. Config — read username set during setup for SSO login_hint
# ---------------------------------------------------------------------------
_CONFIG_PATH = Path(
    os.environ.get("M365COPILOT_CONFIG",
                   str(Path.home() / ".config" / "mcopilot" / "config.json"))
)

def _load_config() -> dict:
    try:
        return json.loads(_CONFIG_PATH.read_text()) if _CONFIG_PATH.exists() else {}
    except Exception:
        return {}


HOST = "127.0.0.1"  # loopback only — never expose to LAN
PORT = int(os.environ.get("M365_SHIM_PORT", "8765"))
MODEL_ID = os.environ.get("M365_SHIM_MODEL", "m365-copilot")

# Streaming strategy:
#   pseudo (DEFAULT) — fetch the final answer, emit as deltas.
#       Best UX: Copilot drafts and rewrites mid-stream; pseudo shows only
#       the final version, no flickering.
#   real   — map chatOverStream SSE live (revision-safe logic included).
#   auto   — try real, fall back to pseudo on error.
STREAM_MODE = os.environ.get("M365_SHIM_STREAM", "pseudo").lower()

# Web grounding default OFF — grounds in internal enterprise data only.
# Override per-request via OpenAI metadata.web_search, or globally M365_SHIM_WEB=1.
WEB_DEFAULT = os.environ.get("M365_SHIM_WEB", "0") == "1"

# Tone directive — nudges Copilot away from its default "headers + emoji +
# bullet list + Action Items: None required" style.
STYLE_PROMPT = os.environ.get(
    "M365_SHIM_STYLE",
    "Be terse. Open with a one-line executive summary, then add only the few "
    "supporting points that matter. Plain prose — no section-header banners, "
    "no emoji, no bulleted lists unless I ask. Skip boilerplate.",
)
_LINK_OFF = " Do not include hyperlinks, URLs, or citation markers."
_LINK_ON = " Include the source link for each item you reference."

# OpenCode sends a large coding-oriented system prompt — don't forward it to
# an M365 Q&A endpoint; it confuses Copilot. Set M365_SHIM_FORWARD_SYSTEM=1 to enable.
FORWARD_SYSTEM = os.environ.get("M365_SHIM_FORWARD_SYSTEM", "0") == "1"
CITATIONS_DEFAULT = os.environ.get("M365_SHIM_CITATIONS", "0") == "1"

# ---------------------------------------------------------------------------
# 3. Token management
#    The shim requests ONLY the 7 read scopes (no Mail.Send) to avoid the
#    admin-consent wall most tenants put on write scopes.
# ---------------------------------------------------------------------------
_tokcache: dict = {"token": None}


def _shim_token() -> str:
    """Return a valid Graph bearer covering the Chat API read scopes."""
    if auth.is_valid(_tokcache["token"]):
        return _tokcache["token"]  # type: ignore[return-value]
    if auth.is_valid(auth.STATIC_TOKEN):
        _tokcache["token"] = auth.STATIC_TOKEN
        return auth.STATIC_TOKEN  # type: ignore[return-value]

    # Reload from disk — the setup --ensure-token step runs in a separate
    # process and writes a fresh token; we must re-read to see it.
    tok = auth.token_silent()
    if tok and auth.is_valid(tok):
        _tokcache["token"] = tok
        return tok

    raise RuntimeError(
        "M365 Copilot token is missing or expired.\n"
        "  Run: ./shim.py --ensure-token\n"
        "  Or:  ./shim.py --login  (device-code fallback)"
    )


# ---------------------------------------------------------------------------
# 4. Persisted session map — conversation-prefix hash → Copilot conversation id
# ---------------------------------------------------------------------------
_SESS_PATH = Path.home() / ".cache" / "mcopilot" / "sessions.json"
_SESS_PATH.parent.mkdir(parents=True, exist_ok=True)
try:
    _SESSIONS: dict[str, str] = json.loads(_SESS_PATH.read_text())
except Exception:
    _SESSIONS = {}


def _save_sessions() -> None:
    try:
        _SESS_PATH.write_text(json.dumps(_SESSIONS))
        _SESS_PATH.chmod(0o600)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 5. OpenAI ↔ Copilot translation helpers
# ---------------------------------------------------------------------------

def _prefix_key(messages: list[dict]) -> str:
    payload = json.dumps(messages, sort_keys=True, ensure_ascii=False)
    return hashlib.sha256(payload.encode()).hexdigest()


def _content_text(content: Any) -> str:
    if content is None:
        return ""
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for p in content:
            if isinstance(p, dict):
                parts.append(p.get("text") or p.get("content") or "")
            else:
                parts.append(str(p))
        return "\n".join(x for x in parts if x)
    return str(content)


def _split_messages(messages: list[dict]) -> tuple[list[str], str, list[dict]]:
    systems = [_content_text(m.get("content")) for m in messages if m.get("role") == "system"]
    latest = ""
    for m in reversed(messages):
        if m.get("role") == "user":
            latest = _content_text(m.get("content"))
            break
    if not latest and messages:
        latest = _content_text(messages[-1].get("content"))
    return systems, latest, messages[:-1]


def _resolve_conversation(messages: list[dict]) -> tuple[str, bool]:
    key = _prefix_key(messages[:-1]) if len(messages) > 1 else None
    if key and key in _SESSIONS:
        return _SESSIONS[key], False
    token = _shim_token()
    created = auth.post("/conversations", {}, token)
    cid = created.get("id")
    if not cid:
        raise RuntimeError(f"Failed to create Copilot conversation: {created}")
    return cid, True


def _build_chat_body(prompt: str, systems: list[str], web: bool, want_cites: bool = False) -> dict:
    style = (STYLE_PROMPT + (_LINK_ON if want_cites else _LINK_OFF)) if STYLE_PROMPT else ""
    text = f"{style}\n\nMy request: {prompt}" if style else prompt
    body: dict = {
        "message": {"text": text},
        "locationHint": {"timeZone": auth.DEFAULT_TZ},
    }
    if FORWARD_SYSTEM and systems:
        body["additionalContext"] = [{"text": s} for s in systems if s]
    ctx: dict = {}
    if not web:
        ctx["webContext"] = {"isWebEnabled": False}
    if ctx:
        body["contextualResources"] = ctx
    return body


def _wants_sources(prompt: str) -> bool:
    if CITATIONS_DEFAULT:
        return True
    return bool(re.search(
        r"\b(link|links|url|urls|source|sources|cite|citation|citations|reference|references)\b",
        prompt or "", re.I,
    ))


_RE_NUM_CITE = re.compile(r"\s*\[\d+\]\(https?://[^)]*\)")
_RE_MD_LINK = re.compile(r"\[([^\]]+)\]\(https?://[^)]*\)")
_RE_BARE_URL = re.compile(r"https?://\S+")
_RE_MULTISPACE = re.compile(r"[ \t]{2,}")


def _strip_links(text: str) -> str:
    text = _RE_NUM_CITE.sub("", text)
    text = _RE_MD_LINK.sub(r"\1", text)
    text = _RE_BARE_URL.sub("", text)
    text = _RE_MULTISPACE.sub(" ", text)
    return re.sub(r" +\n", "\n", text).strip()


def _citation_block(cites: list[dict]) -> str:
    if not cites:
        return ""
    lines = ["", "", "---", "Sources (M365 Copilot grounding):"]
    for i, c in enumerate(cites, 1):
        title = c.get("title") or c.get("source") or "source"
        url = c.get("url") or ""
        lines.append(f"  [{i}] {title}{(' — ' + url) if url else ''}")
    return "\n".join(lines)


def _format_answer(shaped: dict, want_cites: bool = True) -> str:
    answer = shaped.get("answer", "") or ""
    if not want_cites:
        return _strip_links(answer)
    return answer + _citation_block(shaped.get("citations") or [])


def _approx_tokens(s: str) -> int:
    return max(1, len(s) // 4)


# ---------------------------------------------------------------------------
# 6. Real SSE streaming from chatOverStream
# ---------------------------------------------------------------------------

def _extract_stream_text(evt: dict) -> str | None:
    if not isinstance(evt, dict):
        return None
    for path in (("text",), ("message", "text"), ("delta", "text"), ("item", "text")):
        cur: Any = evt
        ok = True
        for k in path:
            if isinstance(cur, dict) and k in cur:
                cur = cur[k]
            else:
                ok = False
                break
        if ok and isinstance(cur, str):
            return cur
    msgs = evt.get("messages")
    if isinstance(msgs, list) and msgs:
        last = msgs[-1]
        if isinstance(last, dict) and isinstance(last.get("text"), str):
            return last["text"]
    return None


def _iter_copilot_stream(cid: str, body: dict, want_cites: bool = True) -> Iterable[str]:
    token = _shim_token()
    emitted = ""
    final_text = ""
    last_attribs: list[dict] = []
    with auth.stream_post(f"/conversations/{cid}/chatOverStream", body, token) as r:
        if r.status_code >= 400:
            r.read()
            raise RuntimeError(f"chatOverStream HTTP {r.status_code}: {r.text[:300]}")
        for raw in r.iter_lines():
            if not raw:
                continue
            line = raw.decode() if isinstance(raw, bytes) else raw
            if not line.startswith("data:"):
                continue
            data = line[5:].strip()
            if not data or data == "[DONE]":
                continue
            try:
                evt = json.loads(data)
            except Exception:
                continue
            msg = None
            msgs = evt.get("messages")
            if isinstance(msgs, list) and msgs and isinstance(msgs[-1], dict):
                msg = msgs[-1]
            chunk = msg.get("text") if isinstance(msg, dict) else _extract_stream_text(evt)
            if isinstance(msg, dict) and msg.get("attributions"):
                last_attribs = msg["attributions"]
            if not isinstance(chunk, str):
                continue
            # Copilot streams CUMULATIVE text and may REWRITE it mid-stream.
            # Emit only forward deltas to avoid duplicating or flickering.
            if chunk.startswith(emitted) and len(chunk) > len(emitted):
                delta = chunk[len(emitted):]
                emitted = chunk
                if delta:
                    yield delta
            final_text = chunk
    if final_text and final_text != emitted and final_text.startswith(emitted):
        yield final_text[len(emitted):]
    elif final_text and not final_text.startswith(emitted):
        lcp = 0
        for a, b in zip(emitted, final_text):
            if a != b:
                break
            lcp += 1
        tail = final_text[lcp:]
        if tail:
            yield ("\n\n" if emitted else "") + tail
    if want_cites:
        cites = [{"type": a.get("attributionType"), "title": a.get("providerDisplayName"),
                  "url": a.get("seeMoreWebUrl")} for a in last_attribs]
        block = _citation_block(cites)
        if block:
            yield block


# ---------------------------------------------------------------------------
# 7. FastAPI server
# ---------------------------------------------------------------------------
from fastapi import FastAPI, Request  # noqa: E402
from fastapi.responses import JSONResponse, StreamingResponse  # noqa: E402

app = FastAPI(title="mcopilot", version="1.0.0")


@app.get("/healthz")
def healthz() -> dict:
    return {"ok": True, "model": MODEL_ID, "boundary": "microsoft-only"}


@app.get("/v1/models")
@app.get("/models")
def models() -> dict:
    return {
        "object": "list",
        "data": [{"id": MODEL_ID, "object": "model", "created": int(time.time()),
                  "owned_by": "microsoft"}],
    }


def _sse(obj: dict) -> str:
    return f"data: {json.dumps(obj, ensure_ascii=False)}\n\n"


def _chunk_obj(cid: str, created: int, delta: dict, finish: str | None = None) -> dict:
    return {
        "id": cid, "object": "chat.completion.chunk", "created": created,
        "model": MODEL_ID,
        "choices": [{"index": 0, "delta": delta, "finish_reason": finish}],
    }


def _is_utility_request(systems: list[str]) -> str | None:
    blob = " ".join(systems).lower()
    if "title generator" in blob or "thread title" in blob or "generate a brief title" in blob:
        return "title"
    return None


def _make_title(user_text: str) -> str:
    line = (user_text or "").strip().splitlines()[0] if (user_text or "").strip() else ""
    line = re.sub(r"\s+", " ", line).strip(" \"'")
    if not line:
        return "M365 Copilot chat"
    return (line[:47] + "…") if len(line) > 48 else line


def _instant_reply(text: str, stream: bool):
    cid = "chatcmpl-" + uuid.uuid4().hex
    created = int(time.time())
    if not stream:
        return JSONResponse(content={
            "id": cid, "object": "chat.completion", "created": created, "model": MODEL_ID,
            "choices": [{"index": 0, "message": {"role": "assistant", "content": text},
                         "finish_reason": "stop"}],
            "usage": {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0},
        })

    def _gen():
        yield _sse(_chunk_obj(cid, created, {"role": "assistant"}))
        yield _sse(_chunk_obj(cid, created, {"content": text}))
        yield _sse(_chunk_obj(cid, created, {}, finish="stop"))
        yield "data: [DONE]\n\n"
    return StreamingResponse(_gen(), media_type="text/event-stream",
                             headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


def _auth_error(exc: Exception) -> JSONResponse:
    return JSONResponse(
        status_code=401,
        content={"error": {
            "message": (
                "M365 Copilot session is not authenticated or token expired. "
                "Run: ./shim.py --ensure-token  Detail: " + str(exc)[:300]
            ),
            "type": "invalid_request_error",
            "code": "m365_reauth_required",
        }},
    )


def _looks_like_auth(exc: Exception) -> bool:
    s = str(exc).lower()
    return any(k in s for k in ("401", "token", "expired", "unauthor", "missing or expired"))


def _remember(messages: list[dict], answer: str, cid: str) -> None:
    future = list(messages) + [{"role": "assistant", "content": answer}]
    _SESSIONS[_prefix_key(future)] = cid
    if messages:
        _SESSIONS[_prefix_key(messages)] = cid
    _save_sessions()


def _chunk_text(s: str, n: int) -> Iterable[str]:
    for i in range(0, len(s), n):
        yield s[i:i + n]


@app.post("/v1/chat/completions")
@app.post("/chat/completions")
async def chat_completions(request: Request):
    try:
        payload = await request.json()
    except Exception:
        return JSONResponse(status_code=400, content={"error": {"message": "invalid JSON body"}})

    messages = payload.get("messages") or []
    if not isinstance(messages, list) or not messages:
        return JSONResponse(status_code=400, content={"error": {"message": "messages required"}})

    stream = bool(payload.get("stream"))
    web = WEB_DEFAULT or bool((payload.get("metadata") or {}).get("web_search"))
    systems, latest, _hist = _split_messages(messages)

    if _is_utility_request(systems) == "title":
        return _instant_reply(_make_title(latest), stream)

    try:
        _shim_token()
    except Exception as exc:
        return _auth_error(exc)

    try:
        cid, _new = _resolve_conversation(messages)
    except Exception as exc:
        if _looks_like_auth(exc):
            return _auth_error(exc)
        return JSONResponse(status_code=502, content={"error": {"message": str(exc)[:500]}})

    want_cites = _wants_sources(latest)
    body = _build_chat_body(latest, systems, web, want_cites)
    completion_id = "chatcmpl-" + uuid.uuid4().hex
    created = int(time.time())

    if not stream:
        try:
            token = _shim_token()
            convo = auth.post(f"/conversations/{cid}/chat", body, token)
            shaped = auth.shape_answer(convo)
        except Exception as exc:
            if _looks_like_auth(exc):
                return _auth_error(exc)
            return JSONResponse(status_code=502, content={"error": {"message": str(exc)[:500]}})
        text = _format_answer(shaped, want_cites)
        _remember(messages, text, cid)
        pt = _approx_tokens(" ".join(_content_text(m.get("content")) for m in messages))
        ct = _approx_tokens(text)
        return JSONResponse(content={
            "id": completion_id, "object": "chat.completion", "created": created,
            "model": MODEL_ID,
            "choices": [{"index": 0, "message": {"role": "assistant", "content": text},
                         "finish_reason": "stop"}],
            "usage": {"prompt_tokens": pt, "completion_tokens": ct, "total_tokens": pt + ct,
                      "_note": "approximate — Copilot Chat API does not report usage"},
        })

    def _stream_gen():
        yield _sse(_chunk_obj(completion_id, created, {"role": "assistant"}))
        full = ""
        used_real = False
        if STREAM_MODE in ("auto", "real"):
            try:
                for delta in _iter_copilot_stream(cid, body, want_cites):
                    used_real = True
                    full += delta
                    yield _sse(_chunk_obj(completion_id, created, {"content": delta}))
            except Exception as exc:
                if STREAM_MODE == "real":
                    yield _sse(_chunk_obj(completion_id, created,
                                          {"content": f"\n[stream error: {str(exc)[:200]}]"}))
                    yield _sse(_chunk_obj(completion_id, created, {}, finish="stop"))
                    yield "data: [DONE]\n\n"
                    return
                used_real = False
        if not used_real:
            try:
                token = _shim_token()
                convo = auth.post(f"/conversations/{cid}/chat", body, token)
                shaped = auth.shape_answer(convo)
                full = _format_answer(shaped, want_cites)
            except Exception as exc:
                full = f"[error: {str(exc)[:200]}]"
            for piece in _chunk_text(full, 24):
                yield _sse(_chunk_obj(completion_id, created, {"content": piece}))
                time.sleep(0.01)
        _remember(messages, full, cid)
        yield _sse(_chunk_obj(completion_id, created, {}, finish="stop"))
        yield "data: [DONE]\n\n"

    return StreamingResponse(
        _stream_gen(), media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


# ---------------------------------------------------------------------------
# 8. CLI
# ---------------------------------------------------------------------------

def _do_sso_login(login_hint: str | None = None) -> int:
    """SSO browser popup — uses your existing Entra session. No device code."""
    cfg = _load_config()
    hint = login_hint or cfg.get("username") or None
    if hint:
        print(f"[mcopilot] Opening browser SSO for {hint}…")
    else:
        print("[mcopilot] Opening browser SSO… (set username in ~/.config/mcopilot/config.json)")
    try:
        token = auth.token_interactive(login_hint=hint)
        _tokcache["token"] = token
    except Exception as exc:
        print(f"[mcopilot] Browser SSO failed: {exc}", file=sys.stderr)
        return 1
    c = auth.decode_jwt(token)
    print(f"[mcopilot] Authenticated: {c.get('upn')}  "
          f"exp_in={int(c.get('exp', 0)) - int(time.time())}s")
    return 0


def _do_login() -> int:
    """Device-code re-auth — fallback for headless environments."""
    from msal import PublicClientApplication
    print("[mcopilot] Device-code sign-in (Chat API read scopes only)…")
    try:
        flow = auth._app.initiate_device_flow(scopes=auth.SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"device flow init failed: {flow}")
        print(f"\n  → Open: {flow['verification_uri']}\n  → Code: {flow['user_code']}\n")
        res = auth._app.acquire_token_by_device_flow(flow)
        if "access_token" not in res:
            raise RuntimeError(res.get("error_description", res))
        auth.persist_cache()
        _tokcache["token"] = res["access_token"]
    except Exception as exc:
        print(f"[mcopilot] Login failed: {exc}", file=sys.stderr)
        return 1
    c = auth.decode_jwt(_tokcache["token"])
    print(f"[mcopilot] Authenticated: {c.get('upn')}  "
          f"exp_in={int(c.get('exp', 0)) - int(time.time())}s")
    return 0


def _do_selftest() -> int:
    print("[mcopilot] self-test (no network)…")
    assert _host_allowed("graph.microsoft.com")
    assert _host_allowed("login.microsoftonline.com")
    assert _host_allowed("127.0.0.1")
    assert not _host_allowed("api.anthropic.com")
    assert not _host_allowed("api.openai.com")
    assert not _host_allowed("evil.example.com")
    blocked = False
    try:
        socket.getaddrinfo("api.anthropic.com", 443)
    except RuntimeError:
        blocked = True
    assert blocked, "egress guard did not block anthropic!"
    assert "Mail.Send" not in auth.CHAT_READ_SCOPES
    print(f"[mcopilot] OK — egress guard active; BASE={auth.BASE}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="M365 Copilot OpenAI-compatible shim")
    ap.add_argument("--ensure-token", action="store_true",
                    help="silent refresh if valid, SSO popup if expired, then exit")
    ap.add_argument("--login", action="store_true",
                    help="device-code re-auth (fallback for headless envs), then exit")
    ap.add_argument("--selftest", action="store_true",
                    help="verify egress guard + imports (no network)")
    ap.add_argument("--port", type=int, default=PORT)
    args = ap.parse_args()

    if args.selftest:
        return _do_selftest()
    if args.login:
        return _do_login()
    if args.ensure_token:
        try:
            _shim_token()
            print("[mcopilot] Token is valid.")
            return 0
        except RuntimeError:
            return _do_sso_login()

    import uvicorn
    print(f"[mcopilot] OpenAI-compatible endpoint: http://{HOST}:{args.port}/v1")
    print(f"[mcopilot] model={MODEL_ID}  stream={STREAM_MODE}  web_default={WEB_DEFAULT}")
    print(f"[mcopilot] Data boundary: Microsoft + loopback only")
    uvicorn.run(app, host=HOST, port=args.port, log_level="info")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
