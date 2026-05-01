# Claude への引き継ぎメモ

---

## 【最新セッション】爆サイ掲示板スクレイパー（2026-05-01）

### 何をしたか

爆サイ（bakusai.com）の掲示板を一括取得してAIに読ませるスクレイパーを作成した。

- ファイル: `bakusai_scraper.py`
- ブランチ: `claude/bulletin-board-scraper-KolsA`（push済み）

### ローカルで動かすための手順

```bash
# 1. リポジトリをクローン（初回のみ）
git clone https://github.com/maedaatsushi5908-design/work
cd work

# 2. ブランチを切り替える
git checkout claude/bulletin-board-scraper-KolsA

# 3. 依存パッケージをインストール
pip install -r requirements.txt
# ← beautifulsoup4 と requests が必要（requirements.txtに記載済み）

# 4. 実行
python bakusai_scraper.py "板のURL" --threads 10 --pages 3 --merge
```

### 板URLの調べ方

1. ブラウザで `https://bakusai.com` を開く
2. 気になる板名をクリック
3. アドレスバーのURLをコピーしてそのまま使う

例：
```bash
# 特定の板（URLはブラウザで確認）
python bakusai_scraper.py "https://bakusai.com/bbs_list/c27/" --threads 5 --pages 3 --merge
```

### 出力ファイル

```
output/
├── 板名_スレタイ.txt        ← スレッドごとのテキスト
└── merged_for_ai.txt        ← 全スレッドまとめ（--merge指定時）
```

`merged_for_ai.txt` をClaudeやChatGPTに貼り付ければそのまま分析できる。

### オプション一覧

| オプション | デフォルト | 意味 |
|-----------|-----------|------|
| `--threads` | 10 | 取得するスレッド数 |
| `--pages` | 3 | スレッドあたりのページ数 |
| `--output` | output | 保存先ディレクトリ |
| `--format` | text | `text` or `json` |
| `--merge` | なし | 全テキストを1ファイルにまとめる |

### 注意点・既知の問題

- 爆サイはBot対策が強いため、403エラーが出る場合がある
- その際は Selenium（ブラウザ自動化）への切り替えが必要 → 必要なら次のセッションで対応
- リクエスト間隔はデフォルト2秒（サーバー負荷軽減のため変えないこと）

---


## このプロジェクトの概要

kenmo氏の「新高値ブレイク投資法」を自動化するPythonプログラムです。
GitHub: https://github.com/maedaatsushi5908-design/work

主な機能は2つ：
1. `main.py` — 過去データでバックテストを実行してグラフを生成する
2. `scanner.py` — 毎朝実行して「今日買える銘柄」「今日損切りする銘柄」をDiscordに通知する

ユーザーはいつもVS CodeでPythonを使っていて、毎朝Discordに通知が届く状態を目指しています。

---

## ファイル構成

```
work/
├── main.py            # バックテスト実行（python main.py）
├── scanner.py         # デイリースキャナー（python scanner.py --notify）
├── notify.py          # 通知モジュール（LINE / Discord / Slack / メール）
├── config.py          # 戦略パラメータ・対象銘柄リスト
├── strategy.py        # 新高値ブレイクシグナル生成
├── backtest.py        # バックテストエンジン
├── data.py            # 株価データ取得（yfinance、オフライン時は合成データ）
├── synthetic_data.py  # テスト用合成株価データ生成
├── report.py          # グラフ・CSV出力
├── positions.json     # 保有ポジション記録ファイル（scanner.pyが読み書き）
├── requirements.txt   # 依存パッケージ
├── .env.example       # 通知設定のテンプレート
├── .env               # 通知設定（Discordなどの秘密情報を記入済み・要更新）
└── .github/
    └── workflows/
        └── daily_scan.yml  # GitHub Actions（毎朝8時に自動実行）
```

---

## 現在の状況

- コードは完成してGitHubにpush済み（ブランチ: `claude/backtest-kenmo-strategy-RwzQm`）
- `.env` にDiscordウェブフックURLを設定済みだが、**URLをチャットに貼ってしまったため再発行が必要**
- ユーザーのPC上で `pip install -r requirements.txt` がエラーになっている（詳細不明）

---

## 次にやること

### 優先度1: pip install のエラーを解決する

エラーメッセージを見て対処する。よくある原因：

```bash
# Pythonバージョンの確認（3.10以上が必要）
python --version
python3 --version

# pipを最新にしてから再試行
pip install --upgrade pip
pip install -r requirements.txt

# それでもダメな場合、個別にインストール
pip install yfinance pandas numpy matplotlib seaborn tqdm
```

### 優先度2: Discordウェブフックを再発行して .env を更新

Discord → チャンネル設定 → 連携サービス → ウェブフック → 削除して新規作成
→ 新しいURLを `.env` の `DISCORD_WEBHOOK_URL=` に貼り付けて保存

### 優先度3: 動作確認

```bash
python scanner.py --notify
```

Discordに通知が届けば完了。

### 優先度4（任意）: 毎朝自動化

**Macの場合（cron）:**
```bash
crontab -e
# 以下を追加（パスは実際のフォルダに合わせる）
0 8 * * 1-5 cd /Users/あなたの名前/work && python scanner.py --notify
```

**GitHub Actionsを使う場合:**
リポジトリをGitHubにpushして、Settings → Secrets に `DISCORD_WEBHOOK_URL` を登録するだけで毎朝8時に自動実行される。

---

## scanner.py の使い方

```bash
# 朝のスキャン + Discord通知
python scanner.py --notify

# 株を買ったらポジションを記録
python scanner.py add --ticker 7203.T --price 2500 --shares 100

# 売ったら削除
python scanner.py remove --ticker 7203.T

# 保有一覧
python scanner.py list
```

---

## 戦略のパラメータ（config.py で変更可能）

| パラメータ | デフォルト | 意味 |
|-----------|-----------|------|
| high_period | 260 | 新高値の判定期間（営業日）≒ 52週 |
| stop_loss_pct | -0.08 | 損切りライン（-8%）|
| trailing_stop_pct | -0.15 | トレーリングストップ（高値から-15%）|
| max_positions | 5 | 最大保有銘柄数 |

対象銘柄はNikkei225主要30銘柄（config.py の `NIKKEI_225_TICKERS` で変更可能）。
