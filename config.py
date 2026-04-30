"""
kenmo氏の新高値ブレイク投資法 バックテスト設定
"""

from dataclasses import dataclass, field
from typing import List


@dataclass
class StrategyConfig:
    # 新高値の判定期間（営業日）
    # 39週 = 約195営業日（最適化により260→195に変更）
    high_period: int = 195

    # 損切りライン（エントリー価格からの下落率）
    stop_loss_pct: float = -0.10  # -10%

    # トレーリングストップ（高値からの下落率）
    trailing_stop_pct: float = -0.20  # -20%

    # 最大保有銘柄数
    max_positions: int = 5

    # 1ポジションあたりの資金割合
    position_size_pct: float = 0.2  # 20%

    # エントリー: 翌日始値でエントリー
    entry_on_next_open: bool = True

    # 出来高フィルター: 前日比較で出来高が増加していること
    volume_filter: bool = True


@dataclass
class BacktestConfig:
    # バックテスト期間
    start_date: str = "2015-01-01"
    end_date: str = "2024-12-31"

    # 初期資金
    initial_capital: float = 1_000_000  # 100万円

    # 売買手数料（片道）
    commission_pct: float = 0.001  # 0.1%

    # スリッページ（片道）
    slippage_pct: float = 0.001  # 0.1%


# Nikkei 225 主要銘柄（東証コード + ".T" サフィックス）
NIKKEI_225_TICKERS: List[str] = [
    "7203.T",  # トヨタ自動車
    "6758.T",  # ソニーグループ
    "6861.T",  # キーエンス
    "8306.T",  # 三菱UFJフィナンシャル・グループ
    "9432.T",  # 日本電信電話
    "6902.T",  # デンソー
    "7267.T",  # 本田技研工業
    "9984.T",  # ソフトバンクグループ
    "4063.T",  # 信越化学工業
    "6954.T",  # ファナック
    "8035.T",  # 東京エレクトロン
    "4543.T",  # テルモ
    "2914.T",  # 日本たばこ産業
    "9433.T",  # KDDI
    "8058.T",  # 三菱商事
    "6367.T",  # ダイキン工業
    "4502.T",  # 武田薬品工業
    "6981.T",  # 村田製作所
    "7751.T",  # キヤノン
    "8802.T",  # 三菱地所
    "6501.T",  # 日立製作所
    "8316.T",  # 三井住友フィナンシャルグループ
    "9022.T",  # 東海旅客鉄道
    "4661.T",  # オリエンタルランド
    "2802.T",  # 味の素
    "7733.T",  # オリンパス
    "4519.T",  # 中外製薬
    "6503.T",  # 三菱電機
    "8031.T",  # 三井物産
    "7741.T",  # HOYA
]

# バックテスト対象ティッカー（デフォルト: Nikkei 225主要銘柄）
DEFAULT_TICKERS = NIKKEI_225_TICKERS
