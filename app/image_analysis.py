import io
import os
import base64
import logging

from PIL import Image

from app.config import AI_MAX_TOKENS
from app.model_registry import get_vision_client

logger = logging.getLogger(__name__)

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HEIC_SUPPORTED = True
except ImportError:
    HEIC_SUPPORTED = False

MAX_DIMENSION = 2048
MAX_IMAGE_SIZE = 20 * 1024 * 1024  # 20MB

IMAGE_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif",
    ".webp", ".gif", ".ico", ".heic", ".heif",
}

VISION_PROMPT = (
    "この画像の内容を詳しく日本語で説明してください。"
    "テキスト、数値、グラフ、表、図形などがあれば正確に読み取ってください。"
    "ビジネス文書や資料の場合は、重要なデータポイントを抽出してください。"
)


async def analyze_image(content: bytes, filename: str,
                        model_id: str = "auto") -> str:
    if len(content) > MAX_IMAGE_SIZE:
        raise ValueError("画像サイズが上限（20MB）を超えています。")

    ext = os.path.splitext(filename)[1].lower()

    if ext in (".heic", ".heif") and not HEIC_SUPPORTED:
        raise ValueError(
            "HEIC/HEIF 形式を処理するには pillow-heif が必要です。"
            "pip install pillow-heif を実行してください。"
        )

    try:
        img = Image.open(io.BytesIO(content))
    except Exception:
        raise ValueError("画像ファイルを開けませんでした。ファイルが破損している可能性があります。")

    if max(img.size) > MAX_DIMENSION:
        img.thumbnail((MAX_DIMENSION, MAX_DIMENSION), Image.LANCZOS)

    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()

    client, model_name, _ = get_vision_client(model_id)

    try:
        response = client.chat.completions.create(
            model=model_name,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "text", "text": VISION_PROMPT},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/png;base64,{b64}"},
                    },
                ],
            }],
            max_tokens=AI_MAX_TOKENS,
        )
    except Exception as e:
        raise ValueError(f"画像分析APIエラー: {str(e)}")

    result = response.choices[0].message.content
    if not result:
        raise ValueError("画像の分析結果が空でした。")

    return result
