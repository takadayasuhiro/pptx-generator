# PPTX Auto Generator

AI（GitHub Models / GPT-4o）を使って、高品質なプレゼンテーション資料（PPTX）を自動生成する Web アプリケーションです。

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
| AI | GitHub Models (GPT-4o) via OpenAI SDK |
| PPTX生成 | python-pptx |
| 画像 | Pexels API |
| PDF解析 | PyMuPDF |
| CSV分析 | pandas / numpy |
| フロントエンド | HTML / CSS / JavaScript |
| インフラ | Docker / Docker Compose |

## セットアップ

### 必要なもの

- Docker & Docker Compose
- GitHub トークン（[GitHub Models](https://github.com/marketplace/models) 用）
- Pexels API キー（[Pexels](https://www.pexels.com/api/) で無料取得）

### 手順

```bash
# 1. クローン
git clone https://github.com/YOUR_USERNAME/pptx-generator.git
cd pptx-generator

# 2. 環境変数を設定
cp .env.example .env
# .env を編集して API キーを設定

# 3. 起動
docker compose up -d --build

# 4. ブラウザでアクセス
# http://localhost:8000
```

### Docker を使わない場合

```bash
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

## 環境変数

| 変数名 | 説明 |
|---|---|
| `GITHUB_TOKEN` | GitHub Models API のアクセストークン |
| `AI_MODEL` | 使用するAIモデル（デフォルト: `openai/gpt-4o`） |
| `PEXELS_API_KEY` | Pexels API キー（画像取得用） |

## プロジェクト構造

```
├── app/
│   ├── main.py           # FastAPI エントリーポイント
│   ├── config.py          # 設定管理
│   ├── models.py          # Pydantic データモデル
│   ├── ai_client.py       # AI API クライアント
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
