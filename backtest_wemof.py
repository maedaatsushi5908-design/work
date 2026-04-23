#!/usr/bin/env python3
"""
WEMOF (Wakkyai's Entry Model Of FX) バックテスト
わっきゃい氏考案の逆張りスキャルピング手法

手法概要:
  - 銘柄: USD/JPY (ドル円)
  - 時間足: 5分足
  - エントリー: ボリンジャーバンド ±3σ タッチで逆張り
  - フィルター: 高ボラティリティ（ニュース・ブレイクアウト）時はスキップ
  - 利確: 3pips
  - 損切り: 30pips（最大保有: 2時間=24本）

注意: ネットワーク制限のため、リアルなUSDJPY統計パラメータに基づく
      合成データを使用しています。実データへの切り替えも可能です。
"""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.gridspec as gridspec
from datetime import datetime, timedelta
import warnings

warnings.filterwarnings("ignore")

# ============================================================
# パラメータ設定
# ============================================================

# ボリンジャーバンド
BB_PERIOD = 20
BB_SIGMA = 3.0

# エントリーフィルター: ATR比率がこの値を超えたらニュース/ブレイクアウトと判断してスキップ
VOLATILITY_THRESHOLD = 2.5

# 決済パラメータ
TAKE_PROFIT_PIPS = 3
STOP_LOSS_PIPS = 30
PIP_SIZE = 0.01        # USDJPY: 1pip = 0.01円
MAX_HOLD_CANDLES = 24  # 最大保有 24本 = 2時間

# 合成データ設定
SEED = 42
N_CANDLES = 17280      # 60日分 (288本/日 × 60日)
USDJPY_BASE = 150.0    # 基準レート
ANNUAL_VOL = 0.085     # 年率ボラティリティ 8.5%
ATR_PERIOD = 14


# ============================================================
# 合成USDJPYデータ生成
# ============================================================

def generate_synthetic_usdjpy(n_candles: int, seed: int = 42) -> pd.DataFrame:
    """
    幾何ブラウン運動 + ファットテール(ニュースショック)で
    リアルなUSDJPY 5分足データを生成する。
    """
    rng = np.random.default_rng(seed)

    # 5分足のボラティリティ (年率→5分)
    bars_per_year = 252 * 288
    bar_vol = ANNUAL_VOL / np.sqrt(bars_per_year)

    # 時間帯別ボラティリティ倍率（東京/ロンドン/NY時間）
    # 1日288本: 各時間のインデックス (JST基準)
    def vol_multiplier(bar_idx: int) -> float:
        hour_jst = (bar_idx % 288) * 5 // 60
        if 8 <= hour_jst < 12:    # 東京時間
            return 1.2
        elif 15 <= hour_jst < 18:  # ロンドン時間
            return 1.4
        elif 21 <= hour_jst < 24:  # NY時間
            return 1.6
        elif 3 <= hour_jst < 6:    # 深夜
            return 0.5
        return 1.0

    # GBMによるClose生成
    closes = np.zeros(n_candles)
    closes[0] = USDJPY_BASE

    # ニュースショック: ランダムに発生する大きな動き
    shock_indices = rng.choice(n_candles, size=n_candles // 150, replace=False)
    shock_sizes = rng.choice([-1, 1], size=len(shock_indices)) * rng.uniform(0.15, 0.50, size=len(shock_indices))

    shock_map = dict(zip(shock_indices, shock_sizes))

    for i in range(1, n_candles):
        vm = vol_multiplier(i)
        # 厚い裾分布 (t分布, df=5) で日常的な動き
        daily_ret = rng.standard_t(df=5) * bar_vol * vm
        shock = shock_map.get(i, 0.0)
        closes[i] = closes[i - 1] * np.exp(daily_ret + shock * bar_vol * 15)

    # OHLC生成
    bar_range_base = USDJPY_BASE * bar_vol * 1.5
    opens = np.zeros(n_candles)
    highs = np.zeros(n_candles)
    lows = np.zeros(n_candles)

    opens[0] = closes[0]
    for i in range(n_candles):
        vm = vol_multiplier(i)
        rng_size = abs(rng.normal(0, bar_range_base * vm * 0.8)) + bar_range_base * vm * 0.3
        opens[i] = closes[i - 1] if i > 0 else USDJPY_BASE
        direction = 1 if closes[i] >= opens[i] else -1

        # ヒゲの生成
        upper_wick = abs(rng.exponential(rng_size * 0.4))
        lower_wick = abs(rng.exponential(rng_size * 0.4))

        highs[i] = max(opens[i], closes[i]) + upper_wick
        lows[i] = min(opens[i], closes[i]) - lower_wick

    # 日時インデックス (2025-10-01 00:00 JST から60日間)
    start_dt = datetime(2025, 10, 1, 0, 0)
    index = pd.date_range(start=start_dt, periods=n_candles, freq="5min")

    df = pd.DataFrame({
        "Open":  opens,
        "High":  highs,
        "Low":   lows,
        "Close": closes,
    }, index=index)

    return df


# ============================================================
# インジケーター計算
# ============================================================

def add_indicators(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # ボリンジャーバンド (±1σ, ±2σ, ±3σ)
    df["BB_Mid"]    = df["Close"].rolling(BB_PERIOD).mean()
    df["BB_Std"]    = df["Close"].rolling(BB_PERIOD).std(ddof=0)
    df["BB_Upper1"] = df["BB_Mid"] + 1 * df["BB_Std"]
    df["BB_Lower1"] = df["BB_Mid"] - 1 * df["BB_Std"]
    df["BB_Upper2"] = df["BB_Mid"] + 2 * df["BB_Std"]
    df["BB_Lower2"] = df["BB_Mid"] - 2 * df["BB_Std"]
    df["BB_Upper3"] = df["BB_Mid"] + 3 * df["BB_Std"]
    df["BB_Lower3"] = df["BB_Mid"] - 3 * df["BB_Std"]

    # Bollinger %B (±3σ基準)
    band_width = df["BB_Upper3"] - df["BB_Lower3"]
    df["PctB"] = (df["Close"] - df["BB_Lower3"]) / band_width.replace(0, np.nan)

    # ATR (ボラティリティフィルター用)
    df["ATR"]          = (df["High"] - df["Low"]).rolling(ATR_PERIOD).mean()
    df["Candle_Range"] = df["High"] - df["Low"]
    df["ATR_Ratio"]    = df["Candle_Range"] / df["ATR"].replace(0, np.nan)

    return df.dropna()


# ============================================================
# シグナル生成
# ============================================================

def generate_signals(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # ±3σ タッチ かつ 低ボラティリティ(ニュース/ブレイクアウト除外)
    calm = df["ATR_Ratio"] < VOLATILITY_THRESHOLD

    df["Long_Signal"]  = (df["Low"]  <= df["BB_Lower3"]) & calm
    df["Short_Signal"] = (df["High"] >= df["BB_Upper3"]) & calm

    return df


# ============================================================
# バックテスト実行
# ============================================================

def run_backtest(df: pd.DataFrame) -> pd.DataFrame:
    closes = df["Close"].values
    highs  = df["High"].values
    lows   = df["Low"].values
    long_sig  = df["Long_Signal"].values
    short_sig = df["Short_Signal"].values
    dates  = df.index

    trades = []
    last_exit_idx = -1  # 同時複数エントリー防止

    for i in range(len(df)):
        if i <= last_exit_idx:
            continue

        if long_sig[i]:
            direction = 1
        elif short_sig[i]:
            direction = -1
        else:
            continue

        entry_price = closes[i]

        if direction == 1:   # ロング: ±3σ下限タッチ→上方向へ平均回帰
            tp_price = entry_price + TAKE_PROFIT_PIPS * PIP_SIZE
            sl_price = entry_price - STOP_LOSS_PIPS * PIP_SIZE
        else:                # ショート: ±3σ上限タッチ→下方向へ平均回帰
            tp_price = entry_price - TAKE_PROFIT_PIPS * PIP_SIZE
            sl_price = entry_price + STOP_LOSS_PIPS * PIP_SIZE

        result     = "Time"
        exit_price = closes[min(i + MAX_HOLD_CANDLES, len(df) - 1)]
        exit_idx   = min(i + MAX_HOLD_CANDLES, len(df) - 1)

        for j in range(i + 1, min(i + MAX_HOLD_CANDLES + 1, len(df))):
            if direction == 1:
                if highs[j] >= tp_price:
                    result, exit_price, exit_idx = "TP", tp_price, j
                    break
                if lows[j] <= sl_price:
                    result, exit_price, exit_idx = "SL", sl_price, j
                    break
            else:
                if lows[j] <= tp_price:
                    result, exit_price, exit_idx = "TP", tp_price, j
                    break
                if highs[j] >= sl_price:
                    result, exit_price, exit_idx = "SL", sl_price, j
                    break

        if result == "TP":
            pnl_pips = TAKE_PROFIT_PIPS
        elif result == "SL":
            pnl_pips = -STOP_LOSS_PIPS
        else:
            pnl_pips = (exit_price - entry_price) * direction / PIP_SIZE

        trades.append({
            "entry_date":  dates[i],
            "exit_date":   dates[exit_idx],
            "direction":   "Long" if direction == 1 else "Short",
            "entry_price": entry_price,
            "exit_price":  exit_price,
            "result":      result,
            "pnl_pips":    pnl_pips,
            "hold_candles": exit_idx - i,
        })
        last_exit_idx = exit_idx

    trades_df = pd.DataFrame(trades)
    if not trades_df.empty:
        trades_df["cumulative_pips"] = trades_df["pnl_pips"].cumsum()
    return trades_df


# ============================================================
# 統計サマリー
# ============================================================

def compute_stats(trades_df: pd.DataFrame, df: pd.DataFrame) -> dict:
    total = len(trades_df)
    if total == 0:
        return {}

    wins   = (trades_df["pnl_pips"] > 0).sum()
    losses = (trades_df["pnl_pips"] <= 0).sum()

    gross_profit = trades_df.loc[trades_df["pnl_pips"] > 0, "pnl_pips"].sum()
    gross_loss   = abs(trades_df.loc[trades_df["pnl_pips"] <= 0, "pnl_pips"].sum())

    rolling_max  = trades_df["cumulative_pips"].cummax()
    max_drawdown = (trades_df["cumulative_pips"] - rolling_max).min()

    # 連勝・連敗
    streak, max_win_streak, max_loss_streak = 0, 0, 0
    for p in trades_df["pnl_pips"]:
        if p > 0:
            streak = max(streak + 1, 1) if streak >= 0 else 1
            max_win_streak = max(max_win_streak, streak)
        else:
            streak = min(streak - 1, -1) if streak <= 0 else -1
            max_loss_streak = max(max_loss_streak, abs(streak))

    return {
        "total":          total,
        "wins":           wins,
        "losses":         losses,
        "win_rate":       wins / total * 100,
        "total_pips":     trades_df["pnl_pips"].sum(),
        "avg_win":        trades_df.loc[trades_df["pnl_pips"] > 0, "pnl_pips"].mean(),
        "avg_loss":       trades_df.loc[trades_df["pnl_pips"] <= 0, "pnl_pips"].mean(),
        "profit_factor":  gross_profit / gross_loss if gross_loss > 0 else float("inf"),
        "max_drawdown":   max_drawdown,
        "tp_count":       (trades_df["result"] == "TP").sum(),
        "sl_count":       (trades_df["result"] == "SL").sum(),
        "time_count":     (trades_df["result"] == "Time").sum(),
        "max_win_streak": max_win_streak,
        "max_loss_streak": max_loss_streak,
        "avg_hold":       trades_df["hold_candles"].mean(),
        "period_start":   df.index[0].strftime("%Y-%m-%d"),
        "period_end":     df.index[-1].strftime("%Y-%m-%d"),
    }


# ============================================================
# 結果表示
# ============================================================

def print_stats(stats: dict):
    print()
    print("=" * 55)
    print("       WEMOF バックテスト結果 (USD/JPY 5分足)")
    print("=" * 55)
    print(f"  期間            : {stats['period_start']} ～ {stats['period_end']}")
    print(f"  総トレード数    : {stats['total']}")
    print(f"  勝率            : {stats['win_rate']:.1f}%  ({stats['wins']}勝 {stats['losses']}敗)")
    print(f"  合計損益        : {stats['total_pips']:+.1f} pips")
    print(f"  平均利益        : {stats['avg_win']:.1f} pips")
    print(f"  平均損失        : {stats['avg_loss']:.1f} pips")
    print(f"  プロフィットF   : {stats['profit_factor']:.3f}")
    print(f"  最大DD          : {stats['max_drawdown']:.1f} pips")
    print(f"  TP / SL / 時間  : {stats['tp_count']} / {stats['sl_count']} / {stats['time_count']}")
    print(f"  最大連勝        : {stats['max_win_streak']}")
    print(f"  最大連敗        : {stats['max_loss_streak']}")
    print(f"  平均保有 (本数) : {stats['avg_hold']:.1f}")
    print("=" * 55)


# ============================================================
# チャート描画
# ============================================================

def plot_results(trades_df: pd.DataFrame, df: pd.DataFrame, stats: dict):
    fig = plt.figure(figsize=(16, 14))
    fig.patch.set_facecolor("#0d1117")

    gs = gridspec.GridSpec(3, 2, figure=fig, hspace=0.45, wspace=0.35)

    DARK_BG   = "#0d1117"
    PANEL_BG  = "#161b22"
    GREEN     = "#2ea043"
    RED       = "#f85149"
    BLUE      = "#58a6ff"
    GOLD      = "#e3b341"
    TEXT_COL  = "#c9d1d9"
    GRID_COL  = "#30363d"

    def style_ax(ax):
        ax.set_facecolor(PANEL_BG)
        ax.tick_params(colors=TEXT_COL, labelsize=8)
        ax.spines["bottom"].set_color(GRID_COL)
        ax.spines["top"].set_color(GRID_COL)
        ax.spines["left"].set_color(GRID_COL)
        ax.spines["right"].set_color(GRID_COL)
        ax.grid(True, color=GRID_COL, linewidth=0.5, alpha=0.6)
        ax.title.set_color(TEXT_COL)

    fig.suptitle(
        "WEMOF (Wakkyai's Entry Model Of FX)  バックテスト結果\n"
        "USD/JPY 5分足  |  ±3σ逆張りスキャルピング",
        fontsize=13, color=TEXT_COL, y=0.98,
    )

    # ── 1. 累積損益曲線 ──────────────────────────────────────────
    ax1 = fig.add_subplot(gs[0, :])
    style_ax(ax1)

    x = trades_df["entry_date"]
    y = trades_df["cumulative_pips"]

    ax1.plot(x, y, color=BLUE, linewidth=1.4, zorder=3)
    ax1.fill_between(x, y, 0, where=(y >= 0), color=GREEN, alpha=0.25)
    ax1.fill_between(x, y, 0, where=(y < 0),  color=RED,   alpha=0.25)
    ax1.axhline(0, color=TEXT_COL, linewidth=0.6, linestyle="--")

    final_val = y.iloc[-1]
    color_fin = GREEN if final_val >= 0 else RED
    ax1.annotate(
        f"{final_val:+.1f} pips",
        xy=(x.iloc[-1], final_val),
        xytext=(-60, 12), textcoords="offset points",
        color=color_fin, fontsize=9, fontweight="bold",
        arrowprops=dict(arrowstyle="->", color=color_fin, lw=1),
    )

    ax1.set_title("累積損益 (pips)", fontsize=10)
    ax1.set_ylabel("Pips", color=TEXT_COL, fontsize=8)
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%m/%d"))
    ax1.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))

    # ── 2. トレード別損益 ────────────────────────────────────────
    ax2 = fig.add_subplot(gs[1, 0])
    style_ax(ax2)

    colors = [GREEN if p > 0 else RED for p in trades_df["pnl_pips"]]
    ax2.bar(range(len(trades_df)), trades_df["pnl_pips"], color=colors, alpha=0.8, width=0.8)
    ax2.axhline(0, color=TEXT_COL, linewidth=0.6, linestyle="--")
    ax2.set_title("トレード別損益 (pips)", fontsize=10)
    ax2.set_xlabel("トレード番号", color=TEXT_COL, fontsize=8)
    ax2.set_ylabel("Pips", color=TEXT_COL, fontsize=8)

    # ── 3. ドローダウン ─────────────────────────────────────────
    ax3 = fig.add_subplot(gs[1, 1])
    style_ax(ax3)

    rolling_max = trades_df["cumulative_pips"].cummax()
    drawdown = trades_df["cumulative_pips"] - rolling_max
    ax3.fill_between(range(len(drawdown)), drawdown, 0, color=RED, alpha=0.5)
    ax3.plot(range(len(drawdown)), drawdown, color=RED, linewidth=0.8)
    ax3.axhline(0, color=TEXT_COL, linewidth=0.6, linestyle="--")
    ax3.set_title("ドローダウン (pips)", fontsize=10)
    ax3.set_xlabel("トレード番号", color=TEXT_COL, fontsize=8)
    ax3.set_ylabel("Pips", color=TEXT_COL, fontsize=8)

    # ── 4. 決済種別内訳 (円グラフ) ──────────────────────────────
    ax4 = fig.add_subplot(gs[2, 0])
    ax4.set_facecolor(PANEL_BG)
    ax4.title.set_color(TEXT_COL)

    labels  = ["TP (利確)", "SL (損切)", "時間切れ"]
    sizes   = [stats["tp_count"], stats["sl_count"], stats["time_count"]]
    pie_colors = [GREEN, RED, GOLD]
    non_zero = [(l, s, c) for l, s, c in zip(labels, sizes, pie_colors) if s > 0]
    if non_zero:
        labels_nz, sizes_nz, colors_nz = zip(*non_zero)
        wedges, texts, autotexts = ax4.pie(
            sizes_nz, labels=labels_nz, colors=colors_nz,
            autopct="%1.1f%%", startangle=90,
            textprops={"color": TEXT_COL, "fontsize": 8},
        )
        for at in autotexts:
            at.set_color(DARK_BG)
            at.set_fontweight("bold")
    ax4.set_title("決済種別内訳", fontsize=10, color=TEXT_COL)

    # ── 5. 統計テーブル ─────────────────────────────────────────
    ax5 = fig.add_subplot(gs[2, 1])
    ax5.set_facecolor(PANEL_BG)
    ax5.axis("off")

    table_data = [
        ["総トレード",     f"{stats['total']}"],
        ["勝率",          f"{stats['win_rate']:.1f}%"],
        ["合計損益",       f"{stats['total_pips']:+.1f} pips"],
        ["平均利益",       f"{stats['avg_win']:.1f} pips"],
        ["平均損失",       f"{stats['avg_loss']:.1f} pips"],
        ["PF",            f"{stats['profit_factor']:.3f}"],
        ["最大DD",        f"{stats['max_drawdown']:.1f} pips"],
        ["最大連勝/連敗", f"{stats['max_win_streak']} / {stats['max_loss_streak']}"],
    ]

    table = ax5.table(
        cellText=table_data,
        colLabels=["指標", "値"],
        cellLoc="center",
        loc="center",
        bbox=[0.0, 0.0, 1.0, 1.0],
    )
    table.auto_set_font_size(False)
    table.set_fontsize(9)

    for (row, col), cell in table.get_celld().items():
        cell.set_facecolor(PANEL_BG if row % 2 == 0 else "#21262d")
        cell.set_text_props(color=TEXT_COL)
        cell.set_edgecolor(GRID_COL)
        if row == 0:
            cell.set_facecolor("#1f6feb")
            cell.set_text_props(color="white", fontweight="bold")

    ax5.set_title("パフォーマンス指標", fontsize=10, color=TEXT_COL)

    out_path = "wemof_backtest_results.png"
    plt.savefig(out_path, dpi=150, bbox_inches="tight", facecolor=DARK_BG)
    print(f"\nチャートを保存: {out_path}")
    return out_path


# ============================================================
# メイン
# ============================================================

def main():
    print("=" * 55)
    print("  WEMOF バックテスト開始")
    print(f"  TP={TAKE_PROFIT_PIPS}pips / SL={STOP_LOSS_PIPS}pips / BB={BB_PERIOD}期間{BB_SIGMA}σ")
    print("=" * 55)

    print("\n[1/4] 合成USDJPYデータ生成中...")
    df_raw = generate_synthetic_usdjpy(N_CANDLES, seed=SEED)
    print(f"      {len(df_raw)}本  ({df_raw.index[0].date()} ～ {df_raw.index[-1].date()})")

    print("[2/4] インジケーター計算中...")
    df = add_indicators(df_raw)
    df = generate_signals(df)

    long_count  = df["Long_Signal"].sum()
    short_count = df["Short_Signal"].sum()
    print(f"      シグナル数: ロング={long_count}  ショート={short_count}")

    print("[3/4] バックテスト実行中...")
    trades_df = run_backtest(df)

    if trades_df.empty:
        print("      トレードが発生しませんでした。パラメータを確認してください。")
        return

    stats = compute_stats(trades_df, df)
    print_stats(stats)

    print("[4/4] チャート描画中...")
    plot_results(trades_df, df, stats)

    # CSVエクスポート
    csv_path = "wemof_trades.csv"
    trades_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    print(f"トレード履歴を保存: {csv_path}")

    # 総評
    print()
    pf = stats["profit_factor"]
    wr = stats["win_rate"]
    total_pips = stats["total_pips"]

    print("【総評】")
    if pf >= 1.5 and total_pips > 0:
        print("  ✓ プロフィットファクター≥1.5かつ黒字。統計的に有望な手法です。")
    elif pf >= 1.0 and total_pips > 0:
        print("  △ 黒字ですがPF<1.5。エッジはあるが安定性に課題があります。")
    else:
        print("  ✗ 赤字またはPF<1.0。パラメータ調整もしくは追加フィルターが必要です。")

    print(f"  勝率 {wr:.1f}%・PF {pf:.2f}・最大DD {stats['max_drawdown']:.1f}pips")
    print()
    print("【注意事項】")
    print("  - 本結果は合成データによるシミュレーションです。")
    print("  - スプレッドコスト・スリッページを含んでいません。")
    print("  - Wemof手法のコアフィルター「気持ち悪い動きか」は裁量判断であり、")
    print("    ATR比率による近似フィルターは完全な再現ではありません。")
    print("  - 過去のパフォーマンスは将来の結果を保証しません。")


if __name__ == "__main__":
    main()
