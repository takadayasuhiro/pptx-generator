import os
import hashlib
import requests
from app.config import PEXELS_API_KEY

CACHE_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "assets", "cache")
os.makedirs(CACHE_DIR, exist_ok=True)


def _cache_path(keyword: str) -> str:
    h = hashlib.md5(keyword.encode()).hexdigest()
    return os.path.join(CACHE_DIR, f"{h}.jpg")


def fetch_image(keyword: str, width: int = 1200, height: int = 800) -> str | None:
    if not keyword:
        return None

    cached = _cache_path(keyword)
    if os.path.exists(cached):
        return cached

    if not PEXELS_API_KEY:
        return None

    try:
        resp = requests.get(
            "https://api.pexels.com/v1/search",
            headers={"Authorization": PEXELS_API_KEY},
            params={"query": keyword, "per_page": 1, "orientation": "landscape"},
            timeout=10,
        )
        resp.raise_for_status()
        data = resp.json()

        if not data.get("photos"):
            return None

        img_url = data["photos"][0]["src"]["large"]
        img_resp = requests.get(img_url, timeout=15)
        img_resp.raise_for_status()

        with open(cached, "wb") as f:
            f.write(img_resp.content)

        return cached
    except Exception:
        return None
