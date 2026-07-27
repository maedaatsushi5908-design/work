# Claude への引き継ぎメモ

## このプロジェクトの概要

「新高値ブレイク投資法」に **移動平均線フィルター** を重ねた売買システムです。
GitHub: https://github.com/maedaatsushi5908-design/work

主な機能は2つ：
1. `main.py` — 過去データでバックテストを実行してグラフを生成する
2. `scanner.py` — 毎朝実行して「今日買える銘柄」「今日損切りする銘柄」をDiscordに通知する

ユーザーはいつもVS CodeでPythonを使っていて、毎朝Discordに通知が届く状態を目指しています。

### 移動平均線フィルターの位置づけ

新高値ブレイクだけだと下降トレンド中の「戻り高値」でも買ってしまうため、
移動平均線で3段階の絞り込みをかけています。

| 段階 | 内容 | デフォルト |
|------|------|-----------|
| エントリーフィルター | 終値>25/75/200日MA・パーフェクトオーダー・200日MAが上向き | 有効 |
| 地合いフィルター | 日経225が200日MAより上のときだけ新規建て | 有効 |
| MA割れエグジット | 終値が75日MAを2日連続で下回ったら手仕舞い | **無効**（`--ma-exit` で有効化）|

効果を確認するには `python main.py --compare` を実行すると、
フィルターを1段ずつ足した4パターンが並んで比較できます。

⚠️ **重要**: この環境（Claude Code on the web）では Yahoo Finance に接続できないため、
バックテストは合成データ（幾何ブラウン運動）で動きます。合成データにはトレンドの
継続性が無いので、**フィルターの有効性の数字は当てになりません**。
ユーザーのPCなど実データが取れる環境で `--compare` を回して判断してください。

---

## ファイル構成

```
work/
├── main.py            # バックテスト実行（python main.py / --compare で比較）
├── scanner.py         # デイリースキャナー（python scanner.py --notify）
├── notify.py          # 通知モジュール（LINE / Discord / Slack / メール）
├── config.py          # 戦略パラメータ・対象銘柄リスト
├── strategy.py        # 新高値ブレイク＋移動平均線フィルターのシグナル生成
├── test_ma_filter.py  # 移動平均線フィルターの単体テスト（python test_ma_filter.py）
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
| high_period | 195 | 新高値の判定期間（営業日）≒ 39週 |
| stop_loss_pct | -0.10 | 損切りライン（-10%）|
| trailing_stop_pct | -0.20 | トレーリングストップ（高値から-20%）|
| max_positions | 5 | 最大保有銘柄数 |
| ma_short_period | 25 | 短期移動平均（営業日）|
| ma_mid_period | 75 | 中期移動平均（営業日）|
| ma_long_period | 200 | 長期移動平均（営業日）|
| ma_filter | True | 移動平均フィルターを使うか |
| ma_exit | False | MA割れで手仕舞うか |
| market_filter | True | 地合いフィルターを使うか |
| market_ma_period | 200 | 地合い判定の移動平均（営業日）|

対象銘柄は品質フィルター済みの116銘柄（config.py の `CURATED_TICKERS` で変更可能）。
`scanner.py` のCLIデフォルトも `config.py` の `StrategyConfig` と揃えてあるので、
バックテストとスキャナーで同じルールが走ります。
