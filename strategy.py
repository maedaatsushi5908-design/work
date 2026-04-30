"""
kenmo氏の新高値ブレイク投資法 シグナル生成モジュール

【戦略ルール】
1. エントリー条件:
   - 終値がN日間（デフォルト: 260営業日 ≈ 52週）の新高値を更新した日
   - （オプション）出来高が増加していること（前日比）

2. エントリー価格:
   - 翌日の始値（entry_on_next_open=True の場合）
   - シグナル当日の終値（entry_on_next_open=False の場合）

3. エグジット条件（以下のいずれかを満たした場合）:
   - 損切り: エントリー価格からの下落率が stop_loss_pct を下回った
   - トレーリングストップ: 保有中の最高値からの下落率が trailing_stop_pct を下回った
"""

import pandas as pd
import numpy as np
from config import StrategyConfig


def compute_signals(df: pd.DataFrame, config: StrategyConfig) -> pd.DataFrame:
    """
    各日付に対してエントリーシグナルを計算する。

    Returns:
        シグナル列を追加したDataFrame:
          - rolling_high: N日間の終値の最高値（当日を除く前N日間）
          - new_high: 当日終値が前N日間の高値を更新したか
          - volume_ok: 出来高フィルターを通過したか
          - signal: エントリーシグナル（Trueのとき買い検討）
    """
    result = df.copy()

    # 前N日間の終値の最高値（当日を含まない: shift(1)してから rolling）
    result["rolling_high"] = (
        result["Close"].shift(1).rolling(window=config.high_period).max()
    )

    # 終値が前N日間の最高値を上回った = 新高値ブレイク
    result["new_high"] = result["Close"] > result["rolling_high"]

    # 出来高フィルター: 前日比で出来高が増加していること
    if config.volume_filter:
        result["volume_ok"] = result["Volume"] > result["Volume"].shift(1)
    else:
        result["volume_ok"] = True

    # エントリーシグナル
    result["signal"] = result["new_high"] & result["volume_ok"]

    return result


def get_entry_price(df: pd.DataFrame, signal_date: pd.Timestamp, config: StrategyConfig) -> float:
    """シグナル発生日に対するエントリー価格を返す。"""
    if config.entry_on_next_open:
        # 翌営業日の始値
        future_dates = df.index[df.index > signal_date]
        if len(future_dates) == 0:
            return float("nan")
        next_date = future_dates[0]
        return float(df.loc[next_date, "Open"])
    else:
        return float(df.loc[signal_date, "Close"])


def get_entry_date(df: pd.DataFrame, signal_date: pd.Timestamp, config: StrategyConfig) -> pd.Timestamp:
    """エントリー実行日を返す（翌営業日始値の場合は翌日）。"""
    if config.entry_on_next_open:
        future_dates = df.index[df.index > signal_date]
        if len(future_dates) == 0:
            return pd.NaT
        return future_dates[0]
    else:
        return signal_date
