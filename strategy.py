"""
新高値ブレイク × 移動平均線フィルター シグナル生成モジュール

【戦略ルール】
1. エントリー条件（すべて満たしたとき買い検討）:
   - 終値がN日間（デフォルト: 195営業日 ≈ 39週）の新高値を更新した
   - 出来高が増加していること（前日比）※volume_filter
   - 移動平均線フィルターを通過していること ※ma_filter
       a. 終値が短期・中期・長期の全移動平均線より上
       b. パーフェクトオーダー（短期MA > 中期MA > 長期MA）
       c. 長期移動平均が上向き（ma_slope_days 日前より高い）
   - 地合いフィルターを通過していること ※market_filter
       指数（日経225）が長期移動平均より上にある

2. エントリー価格:
   - 翌日の始値（entry_on_next_open=True の場合）
   - シグナル当日の終値（entry_on_next_open=False の場合）

3. エグジット条件（以下のいずれかを満たした場合）:
   - 損切り: エントリー価格からの下落率が stop_loss_pct を下回った
   - トレーリングストップ: 保有中の最高値からの下落率が trailing_stop_pct を下回った
   - 移動平均割れ: 終値が移動平均線を ma_exit_confirm_days 日連続で下回った ※ma_exit

【移動平均線をフィルターとして使う狙い】
   新高値ブレイクは「上昇トレンドの初動」を捉える手法だが、下降トレンドの
   途中で起きる一時的な戻り高値でもシグナルが出てしまう。移動平均線を
   フィルターに重ねることで、トレンドが上を向いている局面に限定して
   エントリーし、トレンドが崩れた時点で早めに撤退する。
"""

from typing import Optional

import pandas as pd
import numpy as np

from config import StrategyConfig


def compute_moving_averages(df: pd.DataFrame, config: StrategyConfig) -> pd.DataFrame:
    """
    終値の単純移動平均（SMA）を計算して列を追加する。

    追加される列:
      - ma_short / ma_mid / ma_long : 短期・中期・長期の移動平均
      - ma_exit_line                : エグジット判定に使う移動平均
    """
    result = df.copy()
    close = result["Close"]

    result["ma_short"] = close.rolling(window=config.ma_short_period).mean()
    result["ma_mid"] = close.rolling(window=config.ma_mid_period).mean()
    result["ma_long"] = close.rolling(window=config.ma_long_period).mean()
    result["ma_exit_line"] = close.rolling(window=config.ma_exit_period).mean()

    return result


def compute_ma_filter(df: pd.DataFrame, config: StrategyConfig) -> pd.Series:
    """
    移動平均線フィルターの判定結果を返す（True = エントリー可）。

    ma_filter=False のときは常に True を返す。
    移動平均が計算できない期間（データ不足）は False になる。
    """
    if not config.ma_filter:
        return pd.Series(True, index=df.index)

    ok = pd.Series(True, index=df.index)

    # a. 終値がすべての移動平均線より上にある（上昇トレンドの中にいる）
    if config.ma_require_above:
        ok &= df["Close"] > df["ma_short"]
        ok &= df["Close"] > df["ma_mid"]
        ok &= df["Close"] > df["ma_long"]

    # b. パーフェクトオーダー（短期 > 中期 > 長期）
    if config.ma_require_alignment:
        ok &= df["ma_short"] > df["ma_mid"]
        ok &= df["ma_mid"] > df["ma_long"]

    # c. 長期移動平均が上向き（横ばい・下向きの局面を除外）
    if config.ma_require_slope_up:
        ok &= df["ma_long"] > df["ma_long"].shift(config.ma_slope_days)

    # 移動平均が未計算（NaN）の期間はフィルター不通過として扱う
    ma_ready = df["ma_long"].notna() & df["ma_mid"].notna() & df["ma_short"].notna()

    return (ok & ma_ready).fillna(False)


def compute_ma_exit(df: pd.DataFrame, config: StrategyConfig) -> pd.Series:
    """
    移動平均割れによるエグジットシグナルを返す（True = 手仕舞い）。

    終値が ma_exit_line を ma_exit_confirm_days 日連続で下回った日に True。
    ma_exit=False のときは常に False。
    """
    if not config.ma_exit:
        return pd.Series(False, index=df.index)

    below = df["Close"] < df["ma_exit_line"]
    confirm = max(1, config.ma_exit_confirm_days)

    # 直近 confirm 日がすべて「MA割れ」なら手仕舞い
    streak = below.rolling(window=confirm).sum()
    exit_signal = (streak >= confirm) & df["ma_exit_line"].notna()

    return exit_signal.fillna(False)


def compute_signals(df: pd.DataFrame, config: StrategyConfig) -> pd.DataFrame:
    """
    各日付に対してエントリーシグナルを計算する。

    Returns:
        シグナル列を追加したDataFrame:
          - rolling_high     : N日間の終値の最高値（当日を除く前N日間）
          - new_high         : 当日終値が前N日間の高値を更新したか
          - volume_ok        : 出来高フィルターを通過したか
          - ma_short/mid/long: 各移動平均線
          - ma_ok            : 移動平均線フィルターを通過したか
          - ma_exit_signal   : 移動平均割れのエグジットシグナル
          - signal           : エントリーシグナル（Trueのとき買い検討）
    """
    result = compute_moving_averages(df, config)

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

    # 移動平均線フィルター
    result["ma_ok"] = compute_ma_filter(result, config)

    # 移動平均割れエグジット
    result["ma_exit_signal"] = compute_ma_exit(result, config)

    # エントリーシグナル
    result["signal"] = result["new_high"] & result["volume_ok"] & result["ma_ok"]

    return result


def compute_market_regime(
    benchmark: Optional[pd.DataFrame],
    config: StrategyConfig,
) -> Optional[pd.Series]:
    """
    地合い（マーケットレジーム）を判定する。

    指数の終値が market_ma_period 日移動平均より上なら True（新規エントリー可）。

    Returns:
        日付をインデックスとする bool の Series。
        market_filter=False またはベンチマークが無い場合は None。
    """
    if not config.market_filter or benchmark is None or benchmark.empty:
        return None

    close = benchmark["Close"]
    ma = close.rolling(window=config.market_ma_period).mean()

    regime = (close > ma) & ma.notna()
    return regime.fillna(False)


def is_market_ok(regime: Optional[pd.Series], date: pd.Timestamp) -> bool:
    """
    指定日の地合いが「エントリー可」かを返す。

    指数に該当日のデータが無い場合は直近の判定結果を使う。
    regime が None（地合いフィルター無効）のときは常に True。
    """
    if regime is None:
        return True
    if date in regime.index:
        return bool(regime.loc[date])

    # 該当日が無ければ直前の営業日の判定を引き継ぐ
    past = regime.index[regime.index <= date]
    if len(past) == 0:
        return False
    return bool(regime.loc[past[-1]])


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
