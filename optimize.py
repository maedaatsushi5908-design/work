"""
パラメータ最適化スクリプト

grid search で各パラメータの組み合わせをバックテストし、
総合スコアが高い上位10件を表示する。

使い方:
    python optimize.py            # 売買パラメータを探索
    python optimize.py --mode ma  # 移動平均線フィルターのパラメータを探索
"""

import argparse
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
from data import download_all_stocks, download_benchmark
from config import DEFAULT_TICKERS

# ──────────────────────────────────────────
# 探索グリッド
# ──────────────────────────────────────────

# 売買ルール（損切り・トレーリング・新高値期間）の探索
BASE_GRID = {
    "high_period":       [130, 195, 260],
    "stop_loss_pct":     [-0.05, -0.07, -0.10],
    "trailing_stop_pct": [-0.10, -0.15, -0.20],
    "volume_filter":     [True, False],
}

# 移動平均線フィルターの探索
# 短期/中期/長期の組み合わせは3本セットで動かす
MA_GRID = {
    "ma_set":               [(5, 25, 75), (25, 75, 200), (20, 60, 120), (25, 50, 200)],
    "ma_require_alignment": [True, False],
    "ma_exit":              [True, False],
    "market_filter":        [True, False],
}

# 後方互換用（旧バージョンの GRID を参照しているスクリプト向け）
GRID = BASE_GRID

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


def _metrics(r) -> dict:
    """結果の共通指標を辞書にまとめる。"""
    return {
        "スコア":     round(score(r), 4),
        "総リターン": f"{r.total_return * 100:.1f}%",
        "シャープ":   round(r.sharpe_ratio, 2),
        "PF":         round(r.profit_factor, 2),
        "最大DD":     f"{r.max_drawdown * 100:.1f}%",
        "勝率":       f"{r.win_rate * 100:.0f}%",
        "取引数":     r.total_trades,
        "平均保有日": round(r.avg_holding_days),
    }


def run_base_search(stock_data, benchmark) -> pd.DataFrame:
    """売買パラメータ（新高値期間・損切り・トレーリング）を探索する。"""
    keys = list(BASE_GRID.keys())
    combinations = list(itertools.product(*[BASE_GRID[k] for k in keys]))
    total = len(combinations)
    print(f"組み合わせ数: {total} パターンを検証します\n")

    results = []
    base_cfg = StrategyConfig()

    for i, values in enumerate(combinations, 1):
        params = dict(zip(keys, values))
        cfg = replace(base_cfg, **params)

        r = Backtester(stock_data, cfg, BCFG, benchmark=benchmark).run()
        results.append({
            **_metrics(r),
            "high_period":   params["high_period"],
            "stop_loss":     f"{params['stop_loss_pct']*100:.0f}%",
            "trailing_stop": f"{params['trailing_stop_pct']*100:.0f}%",
            "volume_filter": "あり" if params["volume_filter"] else "なし",
        })

        if i % 10 == 0 or i == total:
            print(f"  進捗: {i}/{total}")

    return pd.DataFrame(results).sort_values("スコア", ascending=False)


def run_ma_search(stock_data, benchmark) -> pd.DataFrame:
    """移動平均線フィルターのパラメータを探索する。"""
    keys = list(MA_GRID.keys())
    combinations = list(itertools.product(*[MA_GRID[k] for k in keys]))
    total = len(combinations)
    print(f"組み合わせ数: {total} パターンを検証します\n")

    results = []
    base_cfg = StrategyConfig()

    # 比較の基準としてMAフィルターなしも計測する
    no_ma = replace(base_cfg, ma_filter=False, ma_exit=False, market_filter=False)
    r_base = Backtester(stock_data, no_ma, BCFG, benchmark=benchmark).run()
    results.append({
        **_metrics(r_base),
        "MA(短/中/長)": "なし",
        "並び条件":     "-",
        "MA割れ売り":   "-",
        "地合い":       "-",
    })

    for i, values in enumerate(combinations, 1):
        params = dict(zip(keys, values))
        short, mid, long = params["ma_set"]
        cfg = replace(
            base_cfg,
            ma_filter=True,
            ma_short_period=short,
            ma_mid_period=mid,
            ma_long_period=long,
            ma_exit_period=mid,
            ma_require_alignment=params["ma_require_alignment"],
            ma_exit=params["ma_exit"],
            market_filter=params["market_filter"],
        )

        r = Backtester(stock_data, cfg, BCFG, benchmark=benchmark).run()
        results.append({
            **_metrics(r),
            "MA(短/中/長)": f"{short}/{mid}/{long}",
            "並び条件":     "あり" if params["ma_require_alignment"] else "なし",
            "MA割れ売り":   "あり" if params["ma_exit"] else "なし",
            "地合い":       "あり" if params["market_filter"] else "なし",
        })

        if i % 10 == 0 or i == total:
            print(f"  進捗: {i}/{total}")

    return pd.DataFrame(results).sort_values("スコア", ascending=False)


def main():
    parser = argparse.ArgumentParser(description="パラメータ最適化")
    parser.add_argument(
        "--mode",
        choices=["base", "ma"],
        default="base",
        help="base: 売買パラメータ / ma: 移動平均線フィルター",
    )
    args = parser.parse_args()

    print("株価データを取得中（1回だけ）...")
    stock_data = download_all_stocks(DEFAULT_TICKERS, BCFG.start_date, BCFG.end_date)
    print(f"取得完了: {len(stock_data)} 銘柄")

    benchmark = download_benchmark(BCFG.start_date, BCFG.end_date)
    if benchmark is None:
        print("警告: 指数データを取得できないため地合いフィルターは無効になります。")
    print()

    if args.mode == "ma":
        df = run_ma_search(stock_data, benchmark)
        title = "移動平均線フィルター 上位 10 パターン"
    else:
        df = run_base_search(stock_data, benchmark)
        title = "売買パラメータ 上位 10 パターン"

    print("\n" + "=" * 100)
    print(f"  {title}")
    print("=" * 100)
    print(df.head(10).to_string(index=False))

    best = df.iloc[0]
    print("\n" + "=" * 100)
    print("  推奨パラメータ（スコア最高）")
    print("=" * 100)
    for col in df.columns:
        if col in ("スコア", "総リターン", "シャープ", "PF", "最大DD", "勝率", "取引数", "平均保有日"):
            continue
        print(f"  {col:<14}: {best[col]}")
    print(f"  → 総リターン {best['総リターン']}  シャープ {best['シャープ']}  "
          f"最大DD {best['最大DD']}  勝率 {best['勝率']}  取引数 {best['取引数']}")


if __name__ == "__main__":
    main()
