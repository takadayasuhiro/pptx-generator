import logging
import os
from dotenv import load_dotenv

logger = logging.getLogger(__name__)

_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# cwd に依存しないよう、まずプロジェクト直下の .env を読む（続けてカレントの .env で上書き可）
load_dotenv(os.path.join(_PROJECT_ROOT, ".env"))
load_dotenv()

# 診断表示・パス解決用（リポジトリ / コンテナの /app 等）
PROJECT_ROOT = _PROJECT_ROOT


def _init_ssl_ca_bundle_path() -> str | None:
    """社内 UTM 用の PEM。SSL_CA_BUNDLE またはプロジェクト直下 SecurityAppliance_SSL_CA.pem。"""
    raw = os.getenv("SSL_CA_BUNDLE", "").strip()
    candidates: list[str] = []
    if raw:
        if os.path.isabs(raw):
            candidates.append(raw)
        else:
            candidates.append(os.path.join(_PROJECT_ROOT, raw))
            candidates.append(os.path.join(os.getcwd(), raw))
    else:
        candidates.append(
            os.path.join(_PROJECT_ROOT, "SecurityAppliance_SSL_CA.pem")
        )
    for p in candidates:
        if p and os.path.isfile(p):
            return os.path.normpath(os.path.abspath(p))
    if raw:
        logger.warning(
            "SSL_CA_BUNDLE が指定されていますがファイルが見つかりません: %s",
            raw[:240],
        )
    return None


SSL_CA_BUNDLE_PATH: str | None = _init_ssl_ca_bundle_path()


def _resolve_ssl_relax_x509_strict() -> bool:
    """VERIFY_X509_STRICT を外すか。社内 PEM 使用時は既定 True（フルチェーンが無くても通しやすい）。明示で false にできる。"""
    raw = os.getenv("SSL_RELAX_X509_STRICT")
    if raw is None or not str(raw).strip():
        return bool(SSL_CA_BUNDLE_PATH)
    s = str(raw).strip().lower()
    if s in ("1", "true", "yes", "on"):
        return True
    if s in ("0", "false", "no", "off"):
        return False
    return bool(SSL_CA_BUNDLE_PATH)


# OpenSSL 3: Missing Authority Key Identifier 等の回避。PEM 利用時は未設定で自動有効
SSL_RELAX_X509_STRICT: bool = _resolve_ssl_relax_x509_strict()

AI_MODEL = os.getenv("AI_MODEL", "gemini-2.5-flash")
AI_TIMEOUT = int(os.getenv("AI_TIMEOUT", "300"))  # 秒（Ollama等の長時間応答用）
AI_MAX_TOKENS = int(os.getenv("AI_MAX_TOKENS", "2000"))  # 1応答あたりの上限（繰り返し暴走防止）
_excel_cap = os.getenv("AI_EXCEL_MAX_TOKENS", "").strip()
# Excel 生成は表・チャート入りの JSON が長くなりやすい（切り捨てで Unterminated string 等になりがち）
AI_EXCEL_MAX_TOKENS = int(_excel_cap) if _excel_cap else max(AI_MAX_TOKENS, 16384)
PEXELS_API_KEY = os.getenv("PEXELS_API_KEY", "")
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY", "")
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "")

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "output")
TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "templates")
PPTX_TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "pptx-template")

os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(PPTX_TEMPLATE_DIR, exist_ok=True)
