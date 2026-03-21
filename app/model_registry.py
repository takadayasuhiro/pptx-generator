import logging
import time

import requests as _requests
from openai import OpenAI

from app.config import (
    GITHUB_TOKEN,
    OPENAI_API_KEY,
    ANTHROPIC_API_KEY,
    GOOGLE_API_KEY,
    OLLAMA_BASE_URL,
)

logger = logging.getLogger(__name__)

_API_KEYS: dict[str, str] = {
    "GITHUB_TOKEN": GITHUB_TOKEN,
    "OPENAI_API_KEY": OPENAI_API_KEY,
    "ANTHROPIC_API_KEY": ANTHROPIC_API_KEY,
    "GOOGLE_API_KEY": GOOGLE_API_KEY,
}

STATIC_MODELS: list[dict] = [
    {
        "id": "gpt-4o",
        "label": "GPT-4o",
        "provider": "github",
        "base_url": "https://models.github.ai/inference",
        "model_id": "openai/gpt-4o",
        "api_key_env": "GITHUB_TOKEN",
        "supports_vision": True,
        "description": "高精度・JSON安定（GitHub Models）",
    },
    {
        "id": "gpt-4o-mini",
        "label": "GPT-4o mini",
        "provider": "github",
        "base_url": "https://models.github.ai/inference",
        "model_id": "openai/gpt-4o-mini",
        "api_key_env": "GITHUB_TOKEN",
        "supports_vision": True,
        "description": "高速・低コスト（GitHub Models）",
    },
    {
        "id": "openai-gpt-4o",
        "label": "GPT-4o (OpenAI)",
        "provider": "openai",
        "base_url": "https://api.openai.com/v1",
        "model_id": "gpt-4o",
        "api_key_env": "OPENAI_API_KEY",
        "supports_vision": True,
        "description": "OpenAI API 直接利用",
    },
    {
        "id": "openai-gpt-4o-mini",
        "label": "GPT-4o mini (OpenAI)",
        "provider": "openai",
        "base_url": "https://api.openai.com/v1",
        "model_id": "gpt-4o-mini",
        "api_key_env": "OPENAI_API_KEY",
        "supports_vision": True,
        "description": "OpenAI API 直接利用・低コスト",
    },
]

_ollama_cache: dict = {"models": [], "ts": 0.0}
OLLAMA_CACHE_TTL = 30


def _detect_ollama() -> list[dict]:
    if not OLLAMA_BASE_URL:
        return []

    now = time.time()
    if now - _ollama_cache["ts"] < OLLAMA_CACHE_TTL:
        return _ollama_cache["models"]

    api_base = OLLAMA_BASE_URL.replace("/v1", "").rstrip("/")

    try:
        resp = _requests.get(f"{api_base}/api/tags", timeout=3)
        if resp.status_code != 200:
            _ollama_cache.update(models=[], ts=now)
            return []

        data = resp.json()
        models = []
        for m in data.get("models", []):
            name = m.get("name", "")
            short = name.split(":")[0]
            vision = any(
                kw in short.lower()
                for kw in ("llava", "vision", "bakllava", "moondream")
            )
            models.append({
                "id": f"ollama-{short}",
                "label": f"{short} (Ollama)",
                "provider": "ollama",
                "base_url": OLLAMA_BASE_URL,
                "model_id": name,
                "api_key_env": None,
                "supports_vision": vision,
                "description": "ローカル実行",
            })

        _ollama_cache.update(models=models, ts=now)
        return models

    except Exception as e:
        logger.debug("Ollama 検出スキップ: %s", e)
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


def get_client(model_id: str = "auto") -> tuple[OpenAI, str, dict]:
    """Return (client, model_name_for_api, model_info)."""
    models = get_available_models()
    available = [m for m in models if m["available"]]

    if not available:
        raise RuntimeError(
            "利用可能なAIモデルがありません。.env にAPIキーを設定してください。"
        )

    target = None
    if model_id == "auto":
        for pref in ("gpt-4o", "gpt-4o-mini"):
            target = next((m for m in available if m["id"] == pref), None)
            if target:
                break
        if not target:
            target = available[0]
    else:
        target = next((m for m in available if m["id"] == model_id), None)
        if not target:
            logger.warning("モデル '%s' は利用不可 → デフォルトにフォールバック", model_id)
            target = available[0]

    client = OpenAI(
        base_url=target["base_url"],
        api_key=_resolve_key(target),
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
            "Vision対応のAIモデルがありません。GPT-4o 等のAPIキーを設定してください。"
        )

    target = vision[0]
    fb_client = OpenAI(
        base_url=target["base_url"],
        api_key=_resolve_key(target),
    )
    return fb_client, target["model_id"], target


def clear_ollama_cache():
    _ollama_cache["ts"] = 0.0
