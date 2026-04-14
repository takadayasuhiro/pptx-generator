#!/usr/bin/env python3
"""Ollama 到達性チェック（コンテナ内から実行用）"""
import requests
import sys

urls = [
    "http://172.17.0.1:11434/api/tags",
    "http://172.18.0.1:11434/api/tags",
    "http://172.19.0.1:11434/api/tags",
    "http://host.docker.internal:11434/api/tags",
    "http://localhost:11434/api/tags",
]
for url in urls:
    try:
        r = requests.get(url, timeout=3)
        if r.status_code == 200:
            data = r.json()
            names = [m.get("name", "") for m in data.get("models", [])]
            print(f"OK: {url} -> {names[:5]}")
            sys.exit(0)
    except Exception as e:
        print(f"FAIL: {url} -> {e}")
print("No working Ollama URL found")
sys.exit(1)
