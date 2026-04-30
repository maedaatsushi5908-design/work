"""
バックテスト結果のレポート・可視化モジュール
"""

from __future__ import annotations

import warnings
from pathlib import Path
from typing import Optional

import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

try:
    import japanize_matplotlib  # noqa: F401
    JAPANIZE = True
except ImportError:
    JAPANIZE = False

OUTPUT_DIR = Path("output")


def _ensure_output_dir() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)


def print_summary(result, benchmark: Optional[pd.DataFrame] = None) -> None:
    """コンソールにパフォーマンスサマリーを表示する。"""
    from backtest import BacktestResult

    r: BacktestResult = result

    print("\n" + "=" * 60)
    print("  kenmo 新高値ブレイク投資法 バックテスト結果")
    print("=" * 60)
    print(f"  総取引数          : {r.total_trades:>8} 回")
    print(f"  勝率              : {r.win_rate:>8.1%}")
    print(f"  平均リターン/取引 : {r.avg_return_per_trade:>8.2%}")
    print(f"  平均利益（勝ち）  : {r.avg_win:>8.2%}")
    print(f"  平均損失（負け）  : {r.avg_loss:>8.2%}")
    print(f"  プロフィットファクタ: {r.profit_factor:>6.2f}")
    print(f"  平均保有日数      : {r.avg_holding_days:>8.1f} 日")
    print("-" * 60)
    print(f"  総リターン        : {r.total_return:>8.2%}")
    print(f"  最大ドローダウン  : {r.max_drawdown:>8.2%}")
    print(f"  シャープレシオ    : {r.sharpe_ratio:>8.2f}")
    print(f"  期末資産          : {r.equity_curve.iloc[-1]:>12,.0f} 円")
    print("=" * 60)

    if r.closed_trades:
        exits = {}
        for t in r.closed_trades:
            exits[t.exit_reason] = exits.get(t.exit_reason, 0) + 1
        print("\n  エグジット内訳:")
        for reason, count in exits.items():
            label = {
                "stop_loss": "損切り（固定）",
                "trailing_stop": "トレーリングストップ",
                "end_of_period": "期末クローズ",
            }.get(reason, reason)
            print(f"    {label:20s}: {count} 回")

    print()


def plot_equity_curve(
    result,
    benchmark: Optional[pd.DataFrame] = None,
    save: bool = True,
    show: bool = False,
) -> None:
    """資産推移グラフを描画する。"""
    _ensure_output_dir()
    fig, axes = plt.subplots(2, 1, figsize=(14, 9), gridspec_kw={"height_ratios": [3, 1]})

    equity = result.equity_curve
    initial = result.initial_capital
    normalized = equity / initial * 100  # 100を基準に正規化

    ax1 = axes[0]
    ax1.plot(normalized.index, normalized.values, label="kenmo戦略", color="royalblue", linewidth=2)

    if benchmark is not None:
        bm = benchmark["Close"].reindex(equity.index, method="ffill").dropna()
        bm_start = bm.iloc[0]
        bm_norm = bm / bm_start * 100
        ax1.plot(bm_norm.index, bm_norm.values, label="日経225", color="tomato", linewidth=1.5, alpha=0.8)

    ax1.axhline(100, color="gray", linestyle="--", linewidth=0.8, alpha=0.6)
    ax1.set_title("kenmo 新高値ブレイク投資法 バックテスト", fontsize=14, fontweight="bold")
    ax1.set_ylabel("資産推移（開始=100）")
    ax1.legend(loc="upper left")
    ax1.grid(True, alpha=0.3)
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y/%m"))
    ax1.xaxis.set_major_locator(mdates.YearLocator())

    # ドローダウン
    ax2 = axes[1]
    peak = equity.cummax()
    drawdown = (equity - peak) / peak * 100
    ax2.fill_between(drawdown.index, drawdown.values, 0, color="tomato", alpha=0.5, label="ドローダウン")
    ax2.set_ylabel("ドローダウン（%）")
    ax2.set_xlabel("日付")
    ax2.legend(loc="lower left")
    ax2.grid(True, alpha=0.3)
    ax2.xaxis.set_major_formatter(mdates.DateFormatter("%Y/%m"))
    ax2.xaxis.set_major_locator(mdates.YearLocator())

    plt.tight_layout()

    if save:
        path = OUTPUT_DIR / "equity_curve.png"
        plt.savefig(path, dpi=150, bbox_inches="tight")
        print(f"  グラフ保存: {path}")

    if show:
        plt.show()

    plt.close()


def plot_trade_distribution(result, save: bool = True, show: bool = False) -> None:
    """取引リターン分布と保有日数ヒストグラムを描画する。"""
    _ensure_output_dir()
    trades = result.closed_trades
    if not trades:
        return

    returns = [t.return_pct * 100 for t in trades]
    holding = [t.holding_days for t in trades]

    fig, axes = plt.subplots(1, 2, figsize=(12, 5))

    # リターン分布
    ax1 = axes[0]
    bins = np.linspace(min(returns) - 1, max(returns) + 1, 40)
    ax1.hist([r for r in returns if r > 0], bins=bins, color="royalblue", alpha=0.7, label="勝ちトレード")
    ax1.hist([r for r in returns if r <= 0], bins=bins, color="tomato", alpha=0.7, label="負けトレード")
    ax1.axvline(0, color="black", linewidth=1)
    ax1.set_title("取引リターン分布")
    ax1.set_xlabel("リターン（%）")
    ax1.set_ylabel("取引数")
    ax1.legend()
    ax1.grid(True, alpha=0.3)

    # 保有日数分布
    ax2 = axes[1]
    ax2.hist(holding, bins=30, color="mediumseagreen", alpha=0.8)
    ax2.set_title("保有日数分布")
    ax2.set_xlabel("保有日数（日）")
    ax2.set_ylabel("取引数")
    ax2.grid(True, alpha=0.3)

    plt.suptitle("kenmo 新高値ブレイク投資法 取引分析", fontsize=13, fontweight="bold")
    plt.tight_layout()

    if save:
        path = OUTPUT_DIR / "trade_distribution.png"
        plt.savefig(path, dpi=150, bbox_inches="tight")
        print(f"  グラフ保存: {path}")

    if show:
        plt.show()

    plt.close()


def plot_monthly_returns(result, save: bool = True, show: bool = False) -> None:
    """月次リターンのヒートマップを描画する。"""
    _ensure_output_dir()
    equity = result.equity_curve
    monthly = equity.resample("ME").last()
    monthly_returns = monthly.pct_change().dropna() * 100

    pivot = pd.DataFrame({
        "year": monthly_returns.index.year,
        "month": monthly_returns.index.month,
        "return": monthly_returns.values,
    }).pivot(index="year", columns="month", values="return")

    month_labels = ["1月", "2月", "3月", "4月", "5月", "6月",
                    "7月", "8月", "9月", "10月", "11月", "12月"]

    fig, ax = plt.subplots(figsize=(14, max(4, len(pivot) * 0.6 + 1)))

    vmax = max(abs(pivot.values[~np.isnan(pivot.values)].max()),
               abs(pivot.values[~np.isnan(pivot.values)].min()))
    im = ax.imshow(pivot.values, cmap="RdYlGn", aspect="auto", vmin=-vmax, vmax=vmax)

    ax.set_xticks(range(len(month_labels)))
    ax.set_xticklabels(month_labels)
    ax.set_yticks(range(len(pivot.index)))
    ax.set_yticklabels(pivot.index)

    for i in range(len(pivot.index)):
        for j in range(pivot.shape[1]):
            val = pivot.iloc[i, j]
            if not np.isnan(val):
                color = "white" if abs(val) > vmax * 0.6 else "black"
                ax.text(j, i, f"{val:.1f}%", ha="center", va="center", fontsize=8, color=color)

    plt.colorbar(im, ax=ax, label="月次リターン（%）")
    ax.set_title("月次リターン ヒートマップ", fontsize=13, fontweight="bold")
    plt.tight_layout()

    if save:
        path = OUTPUT_DIR / "monthly_returns.png"
        plt.savefig(path, dpi=150, bbox_inches="tight")
        print(f"  グラフ保存: {path}")

    if show:
        plt.show()

    plt.close()


def save_trade_log(result, filename: str = "trade_log.csv") -> None:
    """取引ログをCSVに保存する。"""
    _ensure_output_dir()
    trades = result.closed_trades
    if not trades:
        return

    rows = []
    for t in trades:
        rows.append({
            "銘柄": t.ticker,
            "エントリー日": t.entry_date.strftime("%Y-%m-%d"),
            "エントリー価格": round(t.entry_price, 2),
            "エグジット日": t.exit_date.strftime("%Y-%m-%d") if t.exit_date else "",
            "エグジット価格": round(t.exit_price, 2) if t.exit_price else "",
            "保有日数": t.holding_days,
            "リターン(%)": round(t.return_pct * 100, 2),
            "損益(円)": round(t.pnl, 0),
            "エグジット理由": {
                "stop_loss": "損切り",
                "trailing_stop": "トレーリングストップ",
                "end_of_period": "期末クローズ",
            }.get(t.exit_reason, t.exit_reason or ""),
        })

    df = pd.DataFrame(rows)
    path = OUTPUT_DIR / filename
    df.to_csv(path, index=False, encoding="utf-8-sig")
    print(f"  取引ログ保存: {path}")
