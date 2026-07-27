"""
合成株価データ生成モジュール

yfinance が使えない環境向けに、幾何ブラウン運動（GBM）ベースの
リアルな日本株OHLCV時系列を生成する。
"""

from __future__ import annotations

import zlib

import numpy as np
import pandas as pd
from typing import Dict, List, Optional


# 銘柄ごとのシミュレーションパラメータ（年率）
TICKER_PARAMS: Dict[str, Dict] = {
    "7203.T": {"mu": 0.08, "sigma": 0.22, "start_price": 2000},   # トヨタ
    "6758.T": {"mu": 0.12, "sigma": 0.28, "start_price": 8000},   # ソニー
    "6861.T": {"mu": 0.15, "sigma": 0.25, "start_price": 40000},  # キーエンス
    "8306.T": {"mu": 0.07, "sigma": 0.25, "start_price": 700},    # 三菱UFJ
    "9432.T": {"mu": 0.04, "sigma": 0.15, "start_price": 4000},   # NTT
    "6902.T": {"mu": 0.09, "sigma": 0.22, "start_price": 5000},   # デンソー
    "7267.T": {"mu": 0.06, "sigma": 0.20, "start_price": 3500},   # ホンダ
    "9984.T": {"mu": 0.10, "sigma": 0.35, "start_price": 6000},   # SBG
    "4063.T": {"mu": 0.10, "sigma": 0.22, "start_price": 12000},  # 信越化学
    "6954.T": {"mu": 0.08, "sigma": 0.23, "start_price": 20000},  # ファナック
    "8035.T": {"mu": 0.18, "sigma": 0.32, "start_price": 15000},  # 東京エレクトロン
    "4543.T": {"mu": 0.09, "sigma": 0.20, "start_price": 4000},   # テルモ
    "2914.T": {"mu": 0.03, "sigma": 0.15, "start_price": 2500},   # JT
    "9433.T": {"mu": 0.05, "sigma": 0.14, "start_price": 3500},   # KDDI
    "8058.T": {"mu": 0.11, "sigma": 0.22, "start_price": 3000},   # 三菱商事
    "6367.T": {"mu": 0.12, "sigma": 0.24, "start_price": 20000},  # ダイキン
    "4502.T": {"mu": 0.04, "sigma": 0.20, "start_price": 4000},   # 武田薬品
    "6981.T": {"mu": 0.10, "sigma": 0.25, "start_price": 7000},   # 村田製作所
    "7751.T": {"mu": 0.05, "sigma": 0.18, "start_price": 3000},   # キヤノン
    "8802.T": {"mu": 0.06, "sigma": 0.20, "start_price": 2000},   # 三菱地所
    "6501.T": {"mu": 0.14, "sigma": 0.26, "start_price": 3000},   # 日立
    "8316.T": {"mu": 0.08, "sigma": 0.24, "start_price": 4000},   # 三井住友FG
    "9022.T": {"mu": 0.05, "sigma": 0.16, "start_price": 16000},  # JR東海
    "4661.T": {"mu": 0.07, "sigma": 0.22, "start_price": 4000},   # OLC
    "2802.T": {"mu": 0.06, "sigma": 0.18, "start_price": 2000},   # 味の素
    "7733.T": {"mu": 0.08, "sigma": 0.22, "start_price": 3000},   # オリンパス
    "4519.T": {"mu": 0.11, "sigma": 0.22, "start_price": 3500},   # 中外製薬
    "6503.T": {"mu": 0.10, "sigma": 0.22, "start_price": 2500},   # 三菱電機
    "8031.T": {"mu": 0.11, "sigma": 0.22, "start_price": 3000},   # 三井物産
    "7741.T": {"mu": 0.10, "sigma": 0.22, "start_price": 15000},  # HOYA
}

DEFAULT_PARAMS = {"mu": 0.08, "sigma": 0.22, "start_price": 3000}


def _generate_ohlcv(
    dates: pd.DatetimeIndex,
    mu: float,
    sigma: float,
    start_price: float,
    seed: int,
) -> pd.DataFrame:
    """GBM で日次 OHLCV を生成する。"""
    rng = np.random.default_rng(seed)
    n = len(dates)
    dt = 1 / 252

    # 日次リターン（GBM）
    drift = (mu - 0.5 * sigma**2) * dt
    diffusion = sigma * np.sqrt(dt) * rng.standard_normal(n)
    log_returns = drift + diffusion

    # 終値
    closes = start_price * np.exp(np.cumsum(log_returns))

    # 日中変動を模擬して OHLV を生成
    daily_range_pct = np.abs(rng.normal(0, sigma * np.sqrt(dt), n)) + 0.002
    opens = closes * np.exp(rng.normal(0, sigma * np.sqrt(dt) * 0.3, n))
    highs = np.maximum(closes, opens) * (1 + daily_range_pct * rng.uniform(0.3, 1.0, n))
    lows = np.minimum(closes, opens) * (1 - daily_range_pct * rng.uniform(0.3, 1.0, n))

    # 出来高: 価格変動幅と相関させる
    base_volume = 1_000_000
    volume = (base_volume * (1 + 3 * daily_range_pct) * rng.uniform(0.5, 1.5, n)).astype(int)

    return pd.DataFrame(
        {
            "Open": opens,
            "High": highs,
            "Low": lows,
            "Close": closes,
            "Volume": volume,
        },
        index=dates,
    )


def generate_stock_data(
    ticker: str,
    start_date: str,
    end_date: str,
) -> pd.DataFrame:
    """指定銘柄の合成 OHLCV データを生成する。"""
    params = TICKER_PARAMS.get(ticker, DEFAULT_PARAMS)
    dates = pd.bdate_range(start=start_date, end=end_date)

    # seed をティッカー文字列から決定論的に生成
    # 組み込み hash() は実行ごとに変わる（PYTHONHASHSEED）ため crc32 を使う
    seed = zlib.crc32(ticker.encode("utf-8")) % (2**31)

    return _generate_ohlcv(
        dates,
        mu=params["mu"],
        sigma=params["sigma"],
        start_price=params["start_price"],
        seed=seed,
    )


def generate_all_stocks(
    tickers: List[str],
    start_date: str,
    end_date: str,
) -> Dict[str, pd.DataFrame]:
    """複数銘柄の合成データを生成する。"""
    print(f"合成株価データを生成中... ({len(tickers)}銘柄)")
    result = {}
    for ticker in tickers:
        result[ticker] = generate_stock_data(ticker, start_date, end_date)
    print(f"生成完了: {len(result)} 銘柄")
    return result


def generate_benchmark(start_date: str, end_date: str) -> pd.DataFrame:
    """日経225相当のベンチマーク合成データを生成する。"""
    dates = pd.bdate_range(start=start_date, end=end_date)
    return _generate_ohlcv(dates, mu=0.07, sigma=0.18, start_price=20000, seed=225)
