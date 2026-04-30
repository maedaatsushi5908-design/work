"""
パラメータ最適化スクリプト

grid search で各パラメータの組み合わせをバックテストし、
総合スコアが高い上位10件を表示する。
"""

import itertools
import sys
import warnings
from dataclasses import replace

import pandas as pd

warnings.filterwarnings("ignore")

# Windows stdout を UTF-8 に
if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from backtest import Backtester
from config import BacktestConfig, StrategyConfig
from data import download_all_stocks
from config import DEFAULT_TICKERS

# ──────────────────────────────────────────
# 探索グリッド
# ──────────────────────────────────────────
GRID = {
    "high_period":       [130, 195, 260],
    "stop_loss_pct":     [-0.05, -0.07, -0.10],
    "trailing_stop_pct": [-0.10, -0.15, -0.20],
    "volume_filter":     [True, False],
}

BCFG = BacktestConfig()
MIN_TRADES = 15  # 試行回数が少なすぎる結果は除外


def score(r) -> float:
    """
    総合スコア = シャープレシオ × プロフィットファクター × (1 + 総リターン)
                 ÷ (1 + |最大ドローダウン|)
    複数の指標をバランスよく評価する。
    """
    if r.total_trades < MIN_TRADES:
        return -999.0
    return (
        r.sharpe_ratio
        * r.profit_factor
        * (1 + r.total_return)
        / (1 + abs(r.max_drawdown))
    )


def main():
    print("株価データを取得中（1回だけ）...")
    stock_data = download_all_stocks(DEFAULT_TICKERS, BCFG.start_date, BCFG.end_date)
    print(f"取得完了: {len(stock_data)} 銘柄\n")

    keys = list(GRID.keys())
    combinations = list(itertools.product(*[GRID[k] for k in keys]))
    total = len(combinations)
    print(f"組み合わせ数: {total} パターンを検証します\n")

    results = []
    base_cfg = StrategyConfig()

    for i, values in enumerate(combinations, 1):
        params = dict(zip(keys, values))
        cfg = replace(base_cfg, **params)

        bt = Backtester(stock_data, cfg, BCFG)
        r = bt.run()
        s = score(r)

        results.append({
            "スコア":           round(s, 4),
            "総リターン":       f"{r.total_return * 100:.1f}%",
            "シャープ":         round(r.sharpe_ratio, 2),
            "PF":               round(r.profit_factor, 2),
            "最大DD":           f"{r.max_drawdown * 100:.1f}%",
            "勝率":             f"{r.win_rate * 100:.0f}%",
            "取引数":           r.total_trades,
            "平均保有日":       round(r.avg_holding_days),
            "high_period":      params["high_period"],
            "stop_loss":        f"{params['stop_loss_pct']*100:.0f}%",
            "trailing_stop":    f"{params['trailing_stop_pct']*100:.0f}%",
            "volume_filter":    "あり" if params["volume_filter"] else "なし",
        })

        if i % 10 == 0 or i == total:
            print(f"  進捗: {i}/{total}")

    df = pd.DataFrame(results).sort_values("スコア", ascending=False)

    print("\n" + "=" * 90)
    print("  上位 10 パターン")
    print("=" * 90)
    print(df.head(10).to_string(index=False))

    print("\n" + "=" * 90)
    print("  現在の設定（参考）")
    print("=" * 90)
    current = df[
        (df["high_period"] == 260)
        & (df["stop_loss"] == "-8%")
        & (df["trailing_stop"] == "-15%")
        & (df["volume_filter"] == "あり")
    ]
    if not current.empty:
        print(current.to_string(index=False))

    # ベスト設定を表示
    best = df.iloc[0]
    print("\n" + "=" * 90)
    print("  推奨パラメータ（スコア最高）")
    print("=" * 90)
    print(f"  high_period    : {best['high_period']} 日")
    print(f"  stop_loss      : {best['stop_loss']}")
    print(f"  trailing_stop  : {best['trailing_stop']}")
    print(f"  volume_filter  : {best['volume_filter']}")
    print(f"  → 総リターン {best['総リターン']}  シャープ {best['シャープ']}  "
          f"最大DD {best['最大DD']}  勝率 {best['勝率']}")


if __name__ == "__main__":
    main()
