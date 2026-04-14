"""モデル診断 API 用スナップショット組み立て。"""

import os

from app.config import (
    AI_MODEL,
    AI_MAX_TOKENS,
    AI_TIMEOUT,
    OLLAMA_BASE_URL,
    PEXELS_API_KEY,
    PROJECT_ROOT,
    SSL_CA_BUNDLE_PATH,
    SSL_RELAX_X509_STRICT,
)
from app.model_registry import (
    clear_ollama_cache,
    get_api_key_env_status,
    get_available_models,
    get_client,
)


def build_snapshot(*, refresh_ollama: bool = False) -> dict:
    if refresh_ollama:
        clear_ollama_cache()

    models = get_available_models()
    key_status = get_api_key_env_status()

    rows = []
    for m in models:
        env = m.get("api_key_env")
        env_name = env or ""
        env_configured = key_status.get(env_name, True) if env_name else True
        rows.append({
            "id": m["id"],
            "label": m["label"],
            "provider": m["provider"],
            "api_key_env": env,
            "env_configured": env_configured,
            "selectable_available": m["available"],
            "base_url": m["base_url"],
            "api_model_name": m["model_id"],
            "supports_vision": m.get("supports_vision", False),
        })

    auto_id = None
    auto_api_name = None
    auto_error = None
    try:
        _, api_name, info = get_client("auto")
        auto_id = info.get("id")
        auto_api_name = api_name
    except Exception as e:
        auto_error = str(e)[:500]

    integration_notes = []
    if key_status.get("GOOGLE_API_KEY") and not any(
        m.get("api_key_env") == "GOOGLE_API_KEY" for m in models
    ):
        integration_notes.append(
            "GOOGLE_API_KEY は .env にありますが、登録モデルで未使用です（現状は STATIC に未接続）。"
        )

    return {
        "environment": {
            "AI_MODEL": AI_MODEL,
            "AI_MAX_TOKENS": AI_MAX_TOKENS,
            "AI_TIMEOUT": AI_TIMEOUT,
            "OLLAMA_BASE_URL": OLLAMA_BASE_URL or "(未設定・内部既定URLを試行)",
            "pexels_configured": bool(
                PEXELS_API_KEY and str(PEXELS_API_KEY).strip()
            ),
            "api_keys_configured": key_status,
            "ssl_ca_bundle_active": bool(SSL_CA_BUNDLE_PATH),
            "ssl_ca_bundle_file": os.path.basename(SSL_CA_BUNDLE_PATH)
            if SSL_CA_BUNDLE_PATH
            else None,
            "ssl_ca_search_root": PROJECT_ROOT,
            "ssl_ca_bundle_env_set": bool(os.getenv("SSL_CA_BUNDLE", "").strip()),
            "ssl_relax_x509_strict": SSL_RELAX_X509_STRICT,
        },
        "auto_selection": {
            "resolved_model_id": auto_id,
            "resolved_api_model_name": auto_api_name,
            "error": auto_error,
        },
        "integration_notes": integration_notes,
        "models": rows,
    }
