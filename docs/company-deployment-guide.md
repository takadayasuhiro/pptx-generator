# PPTX Auto Generator — 会社サーバー導入ガイド

## 1. サーバースペックと能力評価

| 項目 | スペック | 評価 |
|---|---|---|
| OS | Windows 11 Pro | Docker Desktop / WSL2 対応 |
| CPU | AMD Ryzen 7 9800X3D（8コア/16スレッド、96MB L3キャッシュ） | LLM推論にも有利な大容量キャッシュ |
| GPU | NVIDIA GeForce RTX 5080（16GB VRAM） | 8Bモデルが秒速、14Bも快適 |
| メモリ | 64GB DDR5-5600 | 複数モデル同時ロードでも余裕 |
| ストレージ | 2TB NVMe Gen4 | 数万ファイル生成でも問題なし |

### リソース配分の目安

| 用途 | メモリ消費 | GPU VRAM |
|---|---|---|
| OS + 基本サービス | ~4GB | — |
| Docker + FastAPI（4ワーカー） | ~2GB | — |
| Ollama（8Bモデル × 2同時ロード） | ~10GB | ~10GB |
| ファイル処理バッファ | ~2GB | — |
| **残り余裕** | **~46GB** | **~6GB** |

---

## 2. 利用可能な AI モデル構成

### クラウド API（インターネット接続必須）

| モデル | 提供元 | API キー | 特徴 |
|---|---|---|---|
| gemini-2.5-flash | Google AI（Gemini） | `GOOGLE_API_KEY` | 高速・Vision 対応（推奨） |
| gemini-2.5-flash-lite | Google AI（Gemini） | `GOOGLE_API_KEY` | 低コスト・低遅延 |

アプリは **OpenAI 互換 SDK** で Gemini のエンドポイントに接続します（実体は Google AI Studio のキー）。

### ローカル LLM（Ollama 経由・インターネット不要）

| モデル | VRAM使用 | 推論速度（RTX 5080） | JSON安定性 | 推奨用途 |
|---|---|---|---|---|
| Qwen3.5 4B | ~3GB | 60〜100 tok/秒 | △ | 高速要約・多人数時 |
| **Qwen3.5 8B** | **~5GB** | **30〜60 tok/秒** | **○** | **汎用（推奨）** |
| Qwen3.5 14B | ~9GB | 20〜35 tok/秒 | ◎ | 高品質レポート |
| Qwen3.5 32B | ~18GB | CPU補助必要 | ◎ | 非推奨（VRAM超過） |

> **推奨**: 通常は **Qwen3.5 8B** をメインに、速度重視時に 4B、品質重視時に 14B を切替。

---

## 3. 利用人数別の運用想定

### Gemini（API）メインの場合

| 同時利用者 | サーバー負荷 | 制約 | 対策 |
|---|---|---|---|
| 1〜5人 | ほぼゼロ | なし | 現状のまま |
| 5〜10人 | ほぼゼロ | API レート制限に注意 | ワーカー数増加・`AI_TIMEOUT` 調整 |
| 10〜20人 | 低い | クォータ到達の可能性 | Google AI の利用枠・複数プロジェクトで分散 |
| 20人以上 | 低い | クォータ | Ollama 併用・キューイング |

### Ollama（Qwen3.5 8B）メインの場合

| 同時利用者 | 体感速度 | 対策 |
|---|---|---|
| 1〜2人 | 8〜15秒（快適） | 特になし |
| 3〜5人 | 15〜20秒（実用的） | `OLLAMA_NUM_PARALLEL=4` |
| 5〜10人 | 待ち発生の可能性 | 4Bモデル併用 or API フォールバック |
| 10人以上 | 渋滞 | API メイン + ローカルは機密データ専用 |

### 推奨ハイブリッド運用

```
通常業務        → Gemini（API）で高品質・並列処理
機密データ分析  → Qwen3.5 8B（Ollama）でローカル処理
API 制限到達時  → アプリは利用可能モデルへ順次フォールバック（Gemini 間 → Ollama）
```

---

## 4. セットアップ手順

### 4-1. 基本セットアップ

```bash
# 1. リポジトリのクローン
git clone https://github.com/<your-org>/pptx.git
cd pptx

# 2. .env の設定
cp .env.example .env
# .env を編集して API キーを設定

# 3. Docker で起動
docker compose up --build -d

# 4. ブラウザでアクセス
# http://localhost （ポート 80）
```

### 4-2. .env の設定項目

```env
# 必須: Google AI Studio（Gemini）— https://aistudio.google.com/apikey
GOOGLE_API_KEY=your_key_here

# 必須: 画像検索（PPTX生成用）
PEXELS_API_KEY=XXXXXXXXXXXX

# 任意: 既定モデル（例: gemini-2.5-flash、ollama-qwen3.5-8b）
# AI_MODEL=gemini-2.5-flash

# 任意: ローカル LLM（Ollama・OpenAI 互換ベース URL）
OLLAMA_BASE_URL=http://host.docker.internal:11434/v1
```

> **注意**: Docker 経由で Ollama に接続する場合は `localhost` ではなく
> `host.docker.internal` を使用してください。

### 4-3. Ollama セットアップ（ローカル LLM）

```bash
# WSL2 内で実行

# 1. Ollama インストール（未インストールの場合）
curl -fsSL https://ollama.com/install.sh | sh

# 2. サービス起動
ollama serve

# 3. モデルのインストール
ollama pull qwen3.5:8b     # 推奨メインモデル
ollama pull qwen3.5:4b     # 軽量・高速用（任意）

# 4. 動作確認
ollama list
```

インストールすると、本アプリのモデルセレクターに自動的に表示されます。

### 4-4. Ollama 並列処理の設定（多人数利用時）

```bash
# WSL2 内で環境変数を設定してから起動
OLLAMA_NUM_PARALLEL=4 OLLAMA_MAX_LOADED_MODELS=2 ollama serve
```

| 設定 | 説明 | 推奨値 |
|---|---|---|
| `OLLAMA_NUM_PARALLEL` | 同時処理リクエスト数 | 2〜4 |
| `OLLAMA_MAX_LOADED_MODELS` | 同時ロードモデル数 | 1〜2 |

---

## 5. 用途別モデル選定ガイド

| 用途 | 推奨モデル | 理由 |
|---|---|---|
| テキスト質問・相談 | Gemini Flash / Qwen3.5 8B | どちらでも十分な品質 |
| CSV/Excel データ分析 | Gemini Flash / Qwen3.5 14B | 数値理解力が重要 |
| PPTX 生成 | Gemini Flash | JSON 構造出力の安定性が最重要（不安定時は Ollama 14B） |
| Excel 生成 | Gemini Flash / Qwen3.5 8B | JSON 構造出力が必要 |
| 機密データの分析 | Qwen3.5 8B（ローカル） | データが外部に出ない |
| 短い要約（100文字等） | Gemini Flash | 指示遵守がしやすい |
| 大量リクエスト | Qwen3.5 4B（ローカル） | API クォータを節約 |

---

## 6. スケーリング対策ロードマップ

### Phase 1: 即効対策（作業 30分・コスト 0）

- [ ] Uvicorn ワーカー数を 4 に増加（Dockerfile 1行変更）
- [ ] Ollama `NUM_PARALLEL=4` 設定

```dockerfile
# Dockerfile の CMD を変更
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "80", "--workers", "4"]
```

**効果**: 同時 4リクエスト並列処理。5〜10人の利用に対応。

### Phase 2: 中規模対策（作業 1〜2日・コストは利用量次第）

- [ ] Google AI のクォータ・課金プランの確認（必要なら上限引き上げ）
- [ ] ジョブキュー方式の導入（進捗バー表示）
- [ ] 結果キャッシュの実装（同一入力の再利用）
- [ ] 複数 API キーのラウンドロビン

**効果**: 10〜20人の頻繁な利用に対応。

### Phase 3: 本格スケーリング（作業 3〜5日）

- [ ] Nginx リバースプロキシ + ロードバランサ
- [ ] Docker レプリカ（`replicas: 3` で12並列処理）
- [ ] 処理の分離（AI / ファイル生成 / 画像検索）

**効果**: 20人以上の本格運用に対応。

---

## 7. API キー取得先一覧

| キー | 取得先 | 費用 | 用途 |
|---|---|---|---|
| `GOOGLE_API_KEY` | https://aistudio.google.com/apikey | 無料枠あり（従量あり） | Gemini（クラウド推論） |
| `PEXELS_API_KEY` | https://www.pexels.com/api/ | 無料 | PPTX用画像検索 |

> `GOOGLE_API_KEY` は組織のポリシーに応じてプロジェクト単位で管理してください。
> `PEXELS_API_KEY` は環境ごとに取得推奨。

---

## 8. セキュリティに関する注意事項

- `.env` ファイルには API キーが含まれるため、**Git にコミットしない**こと（`.gitignore` で除外済み）
- Ollama を使用すればデータは**ローカルで完全に処理**され、外部に送信されない
- Gemini（API）使用時はデータが Google のサーバーに送信される点を認識すること
- 機密性の高いデータの分析には **Ollama（ローカルモデル）の使用を推奨**

---

## 9. ネットワーク構成と開発環境

### 9-1. 社内ネットワーク構成

```
┌──────────────────────────┐         ┌──────────────────────────────┐
│  自席PC (192.168.1.186)  │         │  AIサーバー (192.168.1.10)   │
│  Windows 11              │         │  Windows 11 Pro              │
│  WSL2（GPUなし）          │  SSH    │  WSL2 + Docker + Ollama      │
│  Cursor / VS Code        │ ──────► │  Ryzen 9800X3D + RTX 5080    │
│  ブラウザ                │         │  64GB RAM                    │
└──────────────────────────┘         └──────────────────────────────┘
         開発端末                           実行サーバー

  社内ユーザー（全員）
  ブラウザ → http://192.168.1.10 でアクセス（ポート 80）
```

### 9-2. 開発方式: Remote SSH（推奨）

自席 PC の Cursor から AI サーバーに SSH 接続し、サーバー上で直接編集・実行する。

| 項目 | 場所 |
|---|---|
| コード編集 | 自席 PC の Cursor（SSH 経由でサーバー上のファイルを編集） |
| Docker 実行 | AI サーバー |
| Ollama 実行 | AI サーバー |
| 動作確認 | 自席 PC のブラウザ → `http://192.168.1.10` |
| Git 操作 | Cursor のターミナル → サーバー上で実行 |

### 9-3. Remote SSH セットアップ

#### AIサーバー側（192.168.1.10）の準備

```bash
# WSL2 Ubuntu 内で SSH サーバーを有効化
sudo apt update && sudo apt install -y openssh-server
sudo service ssh start

# WSL2 起動時に自動起動させる
echo "sudo service ssh start" >> ~/.bashrc
```

```powershell
# Windows 側（管理者 PowerShell）でファイアウォール許可
New-NetFirewallRule -DisplayName "SSH for WSL2" -Direction Inbound -LocalPort 22 -Protocol TCP -Action Allow
```

#### 自席PC側（192.168.1.186）の準備

```bash
# SSH キー生成（未作成の場合）
ssh-keygen -t ed25519

# 公開鍵をサーバーに転送
ssh-copy-id ユーザー名@192.168.1.10

# 接続テスト
ssh ユーザー名@192.168.1.10
```

`~/.ssh/config`（Windows: `C:\Users\ユーザー名\.ssh\config`）に追記:

```
Host ai-server
    HostName 192.168.1.10
    User ユーザー名
    IdentityFile ~/.ssh/id_ed25519
    ForwardAgent yes
```

#### Cursor での接続

1. `Ctrl+Shift+P` → 「Remote-SSH: Connect to Host...」
2. `ai-server` を選択
3. サーバー上のプロジェクトフォルダを開く（例: `/home/user/pptx`）

以後はローカルと同じ感覚で編集・ターミナル操作が可能。

### 9-4. 本番と開発の分離

Remote SSH で直接編集すると「編集 = 即本番反映」になるため、
大幅改修時は **ディレクトリ分離方式** で本番を守る。

#### ディレクトリ構成

```
/home/user/pptx/          ← 本番用（main ブランチ固定・ポート 80）
/home/user/pptx-dev/      ← 開発用（feature ブランチ・ポート 8001）
```

```bash
# 初回セットアップ
git clone https://github.com/xxx/pptx.git pptx       # 本番
git clone https://github.com/xxx/pptx.git pptx-dev   # 開発
```

#### docker-compose.dev.yml（開発用・ポート 8001）

```yaml
services:
  app-dev:
    build: .
    ports:
      - "8001:80"
    env_file: .env
    volumes:
      - ./output:/app/output
    container_name: pptx-dev
```

#### 並走運用

```
AIサーバー (192.168.1.10)
├─ [本番] pptx-app   :80  ← main ブランチ（ユーザーが利用中）
└─ [開発] pptx-dev   :8001  ← feature ブランチ（開発者のみテスト）
```

| 操作 | コマンド |
|---|---|
| 本番起動 | `cd ~/pptx && docker compose up -d` |
| 開発起動 | `cd ~/pptx-dev && docker compose -f docker-compose.dev.yml up --build -d` |
| 開発停止 | `cd ~/pptx-dev && docker compose -f docker-compose.dev.yml down` |
| 本番にマージ | `cd ~/pptx-dev && git push` → `cd ~/pptx && git pull && docker compose up --build -d` |

#### 日常の開発フロー

```
小さな修正:
  ~/pptx で直接編集 → docker compose restart → 即反映

大幅改修:
  1. cd ~/pptx-dev
  2. git checkout -b feature/new-feature
  3. コードを変更
  4. docker compose -f docker-compose.dev.yml up --build -d
  5. ブラウザで http://192.168.1.10:8001 で検証
     （本番 :80 はそのまま稼働中）
  6. テスト完了 → git push → ~/pptx で git pull → 本番再ビルド
```

#### 方式の使い分けガイド

| 変更の規模 | 方式 | 手順 |
|---|---|---|
| CSS 調整・テキスト修正 | 本番直接編集 | 編集 → restart |
| バグ修正（1〜2ファイル） | 本番直接編集 | 編集 → rebuild |
| 新機能追加 | 開発ディレクトリ | ブランチ → :8001 でテスト → マージ |
| 大幅リファクタリング | 開発ディレクトリ | ブランチ → 長期テスト → マージ |

### 9-5. ポート構成一覧

| サービス | ポート | 用途 | アクセス元 |
|---|---|---|---|
| 本番アプリ | 80 | ユーザー向け | 社内全員 `http://192.168.1.10` |
| 開発アプリ | 8001 | テスト用 | 開発者のみ `http://192.168.1.10:8001` |
| Ollama API | 11434 | LLM 推論 | Docker 内部 |
| SSH | 22 | Remote SSH | 開発者のみ |

---

## 10. トラブルシューティング

> セクション 9 で追加した開発環境に関する問題も含む。

| 問題 | 原因 | 対策 |
|---|---|---|
| モデルセレクターに Ollama モデルが表示されない | Ollama が起動していない / URL が間違い | `ollama serve` 実行確認、`.env` の `OLLAMA_BASE_URL` 確認 |
| Docker から Ollama に接続できない | localhost では Docker→ホスト間通信不可 | `OLLAMA_BASE_URL=http://host.docker.internal:11434/v1` に変更 |
| 「APIのレート制限に達しました」等 | Gemini のクォータ超過 | 時間を置く・別キー・Ollama に切替 |
| PPTX/Excel 生成で JSON パースエラー | モデルの JSON 出力が不正 | `gemini-2.5-flash` のまま再試行、または Ollama の大きめモデルに切替 |
| 画像が PPTX に含まれない | Pexels API キー未設定 / 制限 | `.env` の `PEXELS_API_KEY` を確認 |
| 応答が非常に遅い | Ollama のモデルコールドスタート | 初回のみ。2回目以降は高速（5分以内の再利用時） |
| SSH 接続できない | WSL2 の SSH が起動していない / FW未許可 | `sudo service ssh start` 実行、Windows FW でポート 22 許可 |
| Remote SSH でファイルが見えない | 開いたフォルダが違う | Cursor で正しいパス（`/home/user/pptx`）を指定 |
| 開発ポート 8001 にアクセスできない | 開発コンテナが起動していない | `docker compose -f docker-compose.dev.yml up -d` で起動 |
| 本番と開発で .env が競合 | 同じ .env を共有している | 各ディレクトリに個別の `.env` を配置（内容は同一でOK） |
