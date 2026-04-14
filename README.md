# PPTX Auto Generator

AI（Google Gemini / ローカル Ollama）を使って、高品質なプレゼンテーション資料（PPTX）を自動生成する Web アプリケーションです。

## 機能

- **テキストからPPTX生成** - トピックを入力するだけでAIがスライド構成・コンテンツを自動生成
- **CSV分析 → PPTX** - CSVファイルをアップロードすると、統計分析・グラフ付きのプレゼンを自動生成
- **PDF要約 → PPTX** - PDFの内容を読み取り、要約プレゼンを生成
- **分析サマリーページ** - ファイル添付時にデータの概要・統計・トレンドを自動要約したページを先頭に追加
- **8種類のカラーテーマ** - Ocean Blue, Midnight, Forest, Sunset, Lavender, Monochrome, Coral, Emerald
- **画像の自動取得** - Pexels API でスライドに合った画像を自動挿入
- **多彩なレイアウト** - タイトル、コンテンツ、2カラム、画像付き、チャート、統計カード 等

## 技術スタック

| カテゴリ | 技術 |
|---|---|
| バックエンド | Python 3.13 / FastAPI / Uvicorn |
| AI | Gemini（OpenAI 互換 API）、ローカル Ollama（OpenAI 互換） |
| PPTX生成 | python-pptx |
| 画像 | Pexels API |
| PDF解析 | PyMuPDF |
| CSV分析 | pandas / numpy |
| フロントエンド | HTML / CSS / JavaScript |
| インフラ | Docker / Docker Compose |

## セットアップ

### 必要なもの

- Docker & Docker Compose（または Python + pip）
- **Google AI Studio** の API キー（[Gemini](https://aistudio.google.com/apikey)）
- Pexels API キー（[Pexels](https://www.pexels.com/api/) で無料取得）
- （任意）**Ollama** を同じマシンで起動し、チャット用モデルを `ollama pull` しておく（埋め込み専用モデルは一覧に出ません）

### 手順

```bash
# 1. クローン
git clone https://github.com/YOUR_USERNAME/pptx-generator.git
cd pptx-generator

# 2. 環境変数を設定
cp .env.example .env
# .env を編集して GOOGLE_API_KEY、PEXELS_API_KEY を設定

# 3. 起動
docker compose up -d --build

# 4. ブラウザでアクセス
# http://localhost （ポート 80）
```

### Docker を使わない場合

```bash
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 80
```

ホストで直接起動する場合、ポート 80 には管理者権限が必要な OS があります（`docker compose` での利用はコンテナ内でリッスンするため通常は問題ありません）。

## 環境変数

| 変数名 | 説明 |
|---|---|
| `GOOGLE_API_KEY` | Google AI Studio（Gemini）用 API キー |
| `AI_MODEL` | 使用する AI（例: `gemini-2.5-flash`、`ollama-qwen3.5-9b`） |
| `AI_TIMEOUT` | API タイムアウト秒（任意） |
| `AI_MAX_TOKENS` | 1 応答あたりの最大トークン（任意） |
| `OLLAMA_BASE_URL` | Ollama の OpenAI 互換ベース URL（任意、既定は `http://host.docker.internal:11434/v1` 等） |
| `PEXELS_API_KEY` | Pexels API キー（画像取得用） |

## プロジェクト構造

```
├── app/
│   ├── main.py           # FastAPI エントリーポイント
│   ├── config.py          # 設定管理
│   ├── models.py          # Pydantic データモデル
│   ├── ai_client.py       # AI API クライアント
│   ├── model_registry.py  # モデル一覧・フォールバック順
│   ├── pptx_builder.py    # PPTX 生成エンジン
│   ├── image_service.py   # Pexels 画像取得
│   ├── pdf_service.py     # PDF テキスト抽出
│   └── csv_service.py     # CSV 分析
├── static/
│   ├── index.html         # Web UI
│   ├── style.css
│   └── script.js
├── Dockerfile
├── docker-compose.yml
├── requirements.txt
└── .env.example
```

## ライセンス

MIT
