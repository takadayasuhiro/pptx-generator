import logging
import os
import ssl
import time
from urllib.parse import urlparse

import httpx
import requests as _requests
from openai import OpenAI

from app.config import (
    GOOGLE_API_KEY,
    OLLAMA_BASE_URL,
    AI_MODEL,
    AI_TIMEOUT,
    PROJECT_ROOT,
    SSL_CA_BUNDLE_PATH,
    SSL_RELAX_X509_STRICT,
)

logger = logging.getLogger(__name__)

_API_KEYS: dict[str, str] = {
    "GOOGLE_API_KEY": GOOGLE_API_KEY,
}

# OpenAI 互換 API（Google AI Studio / Gemini）
# gemini-2.0-flash は新規キー向けに提供終了。2.5 系を既定にする。
STATIC_MODELS: list[dict] = [
    {
        "id": "gemini-2.5-flash",
        "label": "Gemini 2.5 Flash",
        "provider": "google",
        "base_url": "https://generativelanguage.googleapis.com/v1beta/openai/",
        "model_id": "gemini-2.5-flash",
        "api_key_env": "GOOGLE_API_KEY",
        "supports_vision": True,
        "description": "Google Gemini 2.5（高速・推奨）",
    },
    {
        "id": "gemini-2.5-flash-lite",
        "label": "Gemini 2.5 Flash-Lite",
        "provider": "google",
        "base_url": "https://generativelanguage.googleapis.com/v1beta/openai/",
        "model_id": "gemini-2.5-flash-lite",
        "api_key_env": "GOOGLE_API_KEY",
        "supports_vision": True,
        "description": "Google Gemini 2.5 Flash-Lite（低コスト）",
    },
]

# チャット非対応・除外する Ollama モデル名のパターン（小文字）
_OLLAMA_CHAT_SKIP_SUBSTR = (
    "bge-", "bge/", "embed", "nomic-embed", "mxbai-embed", "snowflake-arctic-embed",
)

_ollama_cache: dict = {"models": [], "ts": 0.0}
OLLAMA_CACHE_TTL = 30


def _ollama_entry_id(ollama_name: str) -> str:
    """qwen3.5:9b → ollama-qwen3.5-9b（UI・AI_MODEL 用の一意 ID）。"""
    safe = ollama_name.replace(":", "-").replace("/", "-")
    return f"ollama-{safe}"


def _is_ollama_chat_model(ollama_name: str) -> bool:
    low = ollama_name.lower()
    return not any(s in low for s in _OLLAMA_CHAT_SKIP_SUBSTR)


def _detect_ollama() -> list[dict]:
    now = time.time()
    if now - _ollama_cache["ts"] < OLLAMA_CACHE_TTL:
        return _ollama_cache["models"]

    candidates: list[str] = []
    if OLLAMA_BASE_URL:
        candidates.append(OLLAMA_BASE_URL)
    candidates.extend([
        "http://host.docker.internal:11434/v1",
        "http://172.17.0.1:11434/v1",
        "http://172.18.0.1:11434/v1",
        "http://172.19.0.1:11434/v1",
        "http://localhost:11434/v1",
    ])

    seen = set()
    candidates = [u for u in candidates if u and not (u in seen or seen.add(u))]

    for base in candidates:
        api_base = base.replace("/v1", "").rstrip("/")
        try:
            resp = _requests.get(f"{api_base}/api/tags", timeout=2)
            if resp.status_code != 200:
                continue

            data = resp.json()
            models = []
            for m in data.get("models", []):
                name = m.get("name", "")
                if not name or not _is_ollama_chat_model(name):
                    continue
                short = name.split(":")[0]
                vision = any(
                    kw in short.lower()
                    for kw in ("llava", "vision", "bakllava", "moondream")
                )
                entry_id = _ollama_entry_id(name)
                if short.lower() == "qwen3.5":
                    label = f"{name} (Ollama · Qwen 3.5)"
                    desc = "ローカル・Qwen 3.5"
                else:
                    label = f"{name} (Ollama)"
                    desc = f"ローカル · {short}"
                models.append({
                    "id": entry_id,
                    "label": label,
                    "provider": "ollama",
                    "base_url": base,
                    "model_id": name,
                    "api_key_env": None,
                    "supports_vision": vision,
                    "description": desc,
                })

            _ollama_cache.update(models=models, ts=now)
            return models

        except Exception as e:
            logger.debug("Ollama 検出スキップ (%s): %s", base, e)
            continue

    _ollama_cache.update(models=[], ts=now)
    return []


def get_available_models() -> list[dict]:
    result = []
    for m in STATIC_MODELS:
        api_key = _API_KEYS.get(m["api_key_env"], "")
        result.append({**m, "available": bool(api_key)})

    for m in _detect_ollama():
        result.append({**m, "available": True})

    return result


def _resolve_key(model_info: dict) -> str:
    env = model_info.get("api_key_env")
    if not env:
        return "ollama"
    return _API_KEYS.get(env, "")


def _ssl_context_with_custom_ca(ca_file: str) -> ssl.SSLContext:
    """システムの信頼ストアに社内 PEM を追加。verify=PEM のみよりチェーン構築に有利なことがある。"""
    ctx = ssl.create_default_context()
    ctx.load_verify_locations(cafile=ca_file)
    if SSL_RELAX_X509_STRICT and hasattr(ssl, "VERIFY_X509_STRICT"):
        ctx.verify_flags &= ~ssl.VERIFY_X509_STRICT
    return ctx


def _make_openai_client(*, base_url: str, api_key: str, timeout: float) -> OpenAI:
    """社内 CA がある場合は TLS 検証にその PEM を使う（SSL インスペクション対策）。"""
    if SSL_CA_BUNDLE_PATH:
        http_client = httpx.Client(
            timeout=httpx.Timeout(timeout),
            verify=_ssl_context_with_custom_ca(SSL_CA_BUNDLE_PATH),
        )
        return OpenAI(
            base_url=base_url,
            api_key=api_key,
            http_client=http_client,
        )
    return OpenAI(
        base_url=base_url,
        api_key=api_key,
        timeout=timeout,
    )


def _auto_pick_target(available: list[dict]) -> dict:
    """model_id == auto 用。"""

    def by_id(mid: str) -> dict | None:
        return next((m for m in available if m["id"] == mid), None)

    target = None
    if AI_MODEL and str(AI_MODEL).strip():
        am = AI_MODEL.strip()
        target = by_id(am)
        if not target:
            for m in available:
                if m["provider"] == "ollama" and (
                    m["id"] == am or m["id"].startswith(am + "-")
                ):
                    target = m
                    break
        if not target:
            target = by_id(am.lower()) or by_id(am)

    if not target:
        for pref in ("gemini-2.5-flash", "gemini-2.5-flash-lite"):
            t = by_id(pref)
            if t:
                target = t
                break

    if not target:
        ollama_m = [m for m in available if m["provider"] == "ollama"]
        qwen_first = [m for m in ollama_m if "qwen" in m["id"].lower()]
        olist = qwen_first + [x for x in ollama_m if x not in qwen_first]
        target = olist[0] if olist else None

    if not target:
        target = available[0]

    return target


def get_client(model_id: str = "auto") -> tuple[OpenAI, str, dict]:
    """Return (client, model_name_for_api, model_info)."""
    models = get_available_models()
    available = [m for m in models if m["available"]]

    if not available:
        raise RuntimeError(
            "利用可能なAIモデルがありません。.env に GOOGLE_API_KEY または Ollama を用意してください。"
        )

    if model_id == "auto":
        target = _auto_pick_target(available)
    else:
        target = next((m for m in available if m["id"] == model_id), None)
        if not target:
            logger.warning("モデル '%s' は利用不可 → デフォルトにフォールバック", model_id)
            target = _auto_pick_target(available)

    client = _make_openai_client(
        base_url=target["base_url"],
        api_key=_resolve_key(target),
        timeout=float(AI_TIMEOUT),
    )
    return client, target["model_id"], target


def get_vision_client(preferred: str = "auto") -> tuple[OpenAI, str, dict]:
    """Vision 対応モデルを取得。非対応モデルが指定された場合は自動フォールバック。"""
    client, model_name, info = get_client(preferred)
    if info.get("supports_vision"):
        return client, model_name, info

    models = get_available_models()
    vision = [m for m in models if m["available"] and m.get("supports_vision")]
    if not vision:
        raise RuntimeError(
            "Vision対応のAIモデルがありません。Gemini を .env で有効にするか、"
            "Vision 対応の Ollama モデルを追加してください。"
        )

    target = vision[0]
    fb_client = _make_openai_client(
        base_url=target["base_url"],
        api_key=_resolve_key(target),
        timeout=float(AI_TIMEOUT),
    )
    return fb_client, target["model_id"], target


def clear_ollama_cache():
    _ollama_cache["ts"] = 0.0


def get_api_key_env_status() -> dict[str, bool]:
    """診断用: 各 API キー環境変数が非空か（値は返さない）。"""
    return {k: bool(v and str(v).strip()) for k, v in _API_KEYS.items()}


def _redact_diagnostics_value(obj, depth: int = 0):
    if depth > 8:
        return "…"
    if isinstance(obj, dict):
        out = {}
        for k, v in obj.items():
            lk = str(k).lower()
            if lk in ("api_key", "authorization", "access_token", "secret") and not isinstance(
                v, (dict, list)
            ):
                out[k] = "(redacted)"
            else:
                out[k] = _redact_diagnostics_value(v, depth + 1)
        return out
    if isinstance(obj, list):
        return [_redact_diagnostics_value(x, depth + 1) for x in obj[:100]]
    if isinstance(obj, str) and len(obj) > 4000:
        return obj[:4000] + "…"
    return obj


def _collect_exception_chain(exc: BaseException, max_depth: int = 8) -> list[dict]:
    out: list[dict] = []
    seen: set[int] = set()
    cur: BaseException | None = exc
    depth = 0
    while cur is not None and depth < max_depth:
        if id(cur) in seen:
            break
        seen.add(id(cur))
        entry: dict = {"type": type(cur).__name__, "message": str(cur)[:2000]}
        eno = getattr(cur, "errno", None)
        if eno is not None:
            entry["errno"] = eno
        out.append(entry)
        nxt = cur.__cause__ or cur.__context__
        cur = nxt if nxt is not cur else None
        depth += 1
    return out


def _exception_messages_flat(exc: BaseException) -> str:
    """上位が Connection error. だけでも、__cause__ 側の SSL メッセージをヒント判定に使う。"""
    parts: list[str] = []
    seen: set[int] = set()
    cur: BaseException | None = exc
    depth = 0
    while cur is not None and depth < 12:
        if id(cur) in seen:
            break
        seen.add(id(cur))
        parts.append(str(cur))
        nxt = cur.__cause__ or cur.__context__
        cur = nxt if nxt is not cur else None
        depth += 1
    return " ".join(parts)


def _build_ping_error_diagnostics(exc: BaseException, target: dict) -> dict:
    base = target.get("base_url") or ""
    host = None
    try:
        host = urlparse(base).hostname
    except Exception:
        pass
    d: dict = {
        "exception_type": type(exc).__name__,
        "message": str(exc)[:4000],
        "request_base_url": base,
        "request_host": host,
        "request_model_id": target.get("model_id"),
        "ssl_ca_bundle_used": bool(SSL_CA_BUNDLE_PATH),
        "ssl_ca_bundle_file": os.path.basename(SSL_CA_BUNDLE_PATH)
        if SSL_CA_BUNDLE_PATH
        else None,
        "ssl_ca_search_root": PROJECT_ROOT,
        "ssl_ca_default_pem": os.path.join(PROJECT_ROOT, "SecurityAppliance_SSL_CA.pem"),
        "ssl_relax_x509_strict": SSL_RELAX_X509_STRICT,
    }
    status = getattr(exc, "status_code", None)
    if status is not None:
        d["http_status"] = status
    body = getattr(exc, "body", None)
    if body is not None:
        if isinstance(body, (dict, list)):
            d["api_error_body"] = _redact_diagnostics_value(body)
        elif isinstance(body, str) and body.strip():
            d["api_error_body_text"] = body[:3000]
    req_id = getattr(exc, "request_id", None)
    if req_id:
        d["request_id"] = str(req_id)[:200]
    resp = getattr(exc, "response", None)
    if resp is not None:
        try:
            headers = getattr(resp, "headers", None)
            if headers:
                hdr_out = {}
                for key in ("x-request-id", "x-goog-request-id", "cf-ray", "via"):
                    v = headers.get(key)
                    if not v and key == "x-request-id":
                        v = headers.get("X-Request-Id")
                    if v:
                        hdr_out[key] = str(v)[:200]
                if hdr_out:
                    d["response_headers"] = hdr_out
            sc = getattr(resp, "status_code", None)
            if sc is not None:
                d["http_status"] = sc
            if body is None and not d.get("api_error_body_text"):
                txt = (getattr(resp, "text", None) or "")[:3000]
                if txt.strip():
                    d["response_text_preview"] = txt
        except Exception:
            pass
    chain = _collect_exception_chain(exc)
    if len(chain) > 1:
        d["exception_chain"] = chain
    hints: list[str] = []
    etn = type(exc).__name__
    msg_l = _exception_messages_flat(exc).lower()
    http_st = d.get("http_status")
    cert_fail = (
        "certificate_verify_failed" in msg_l
        or "sslcertverificationerror" in msg_l
        or "unable to get local issuer certificate" in msg_l
    )
    if not cert_fail and (
        etn in ("APIConnectionError",)
        or "connection error" in msg_l
        or "connection refused" in msg_l
    ):
        hints.append(
            "接続に失敗しています。社内 UTM / ファイアウォールが generativelanguage.googleapis.com "
            "への HTTPS をブロックしていないか、プロキシが必要か、DNS が解決できるかを確認してください。"
        )
    if etn == "APITimeoutError" or "timeout" in msg_l:
        hints.append(
            "タイムアウト: ネットワーク遅延・フィルタによる遅延、または応答が返っていない可能性があります。"
        )
    if http_st == 403:
        hints.append(
            "HTTP 403: API キー制限（IP・リファラー）、Generative Language API の無効、課金・権限を確認してください。"
        )
    if http_st == 401:
        hints.append("HTTP 401: API キーが無効か、.env の値が誤っている可能性があります。")
    if http_st == 429:
        hints.append("HTTP 429: レート制限に達している可能性があります。")
    if http_st == 404 and (
        "no longer available" in msg_l
        or "not_found" in msg_l
        or "not found" in msg_l
    ):
        hints.append(
            "HTTP 404: そのモデル ID は Google 側で新規利用不可または廃止の可能性があります。"
            "既定の gemini-2.5-flash へ更新するか、AI_MODEL とモデル一覧を最新のアプリに合わせてください。"
        )
    if "ssl" in msg_l or "certificate" in msg_l or "tls" in msg_l:
        hints.append(
            "SSL/TLS 関連: プロキシの証明書検証に失敗している場合は、信頼する CA を OS/コンテナに追加する必要があることがあります。"
        )
    if cert_fail:
        if SSL_CA_BUNDLE_PATH:
            hints.append(
                "SSL 検証エラー: SecurityAppliance_SSL_CA.pem（SSL_CA_BUNDLE）は読み込み済みです。"
                "PEM が正しいルート CA か、中間証明書の連結が必要かをインフラ担当に確認してください。"
            )
        else:
            hints.append(
                "SSL 検証エラー（社内 UTM の SSL インスペクションが典型）: ssl_ca_search_root 直下に "
                "SecurityAppliance_SSL_CA.pem を置くか、.env に SSL_CA_BUNDLE=（PEM の絶対パス）を指定してアプリを再起動。"
                "Docker の場合はホストに PEM を置いてビルドに含めるか、-v でマウントし SSL_CA_BUNDLE=/app/SecurityAppliance_SSL_CA.pem を設定。"
            )
    if "missing authority key identifier" in msg_l:
        if SSL_RELAX_X509_STRICT:
            hints.append(
                "Missing Authority Key Identifier: ssl_relax_x509_strict が既に有効でも失敗する場合は、"
                "PEM がルート CA 単体で足りないことがあります。可能なら中間 CA の連結をインフラに依頼してください。"
            )
        else:
            hints.append(
                "Missing Authority Key Identifier: 社内 PEM 利用時は既定で SSL_RELAX_X509_STRICT が有効のはずです。"
                ".env に SSL_RELAX_X509_STRICT が false と書かれていないか確認し、アプリを再起動してください。"
            )
    if "name or service not known" in msg_l or "nodename nor servname" in msg_l:
        hints.append("DNS 解決に失敗しています。ネットワーク設定を確認してください。")
    if hints:
        d["operator_hints"] = list(dict.fromkeys(hints))
    return d


def ping_model_minimal(model_id: str) -> dict:
    """診断用: 指定モデルIDへ最小チャットを送り疎通確認。フォールバックは行わない。"""
    models = get_available_models()
    target = next((m for m in models if m["id"] == model_id), None)
    if not target:
        return {"ok": False, "error": "モデルIDがカタログにありません"}
    if not target["available"]:
        env = target.get("api_key_env")
        detail = f".env の {env} が未設定です" if env else "接続条件を満たしていません"
        return {"ok": False, "error": f"UI上は利用不可: {detail}"}

    ping_timeout = min(120.0, float(AI_TIMEOUT))
    client = _make_openai_client(
        base_url=target["base_url"],
        api_key=_resolve_key(target),
        timeout=ping_timeout,
    )
    t0 = time.perf_counter()
    try:
        r = client.chat.completions.create(
            model=target["model_id"],
            messages=[{"role": "user", "content": "Reply with exactly: OK"}],
            max_tokens=16,
            temperature=0,
        )
        ms = int((time.perf_counter() - t0) * 1000)
        text = (r.choices[0].message.content or "").strip()
        return {
            "ok": True,
            "latency_ms": ms,
            "response_preview": text[:120],
            "api_model_name": target["model_id"],
        }
    except Exception as e:
        elapsed_ms = int((time.perf_counter() - t0) * 1000)
        diag = _build_ping_error_diagnostics(e, target)
        diag["elapsed_ms"] = elapsed_ms
        summary = str(e)[:600]
        hs = diag.get("http_status")
        if hs is not None:
            summary = f"HTTP {hs}: {summary}"
        logger.warning(
            "ping_model_minimal NG model_id=%s type=%s http=%s",
            model_id,
            type(e).__name__,
            hs,
        )
        return {"ok": False, "error": summary, "diagnostics": diag}
