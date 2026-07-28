"""
韓国指数（KOSPI）との連動チェック

「最近は日本株が韓国指数と連動している」という体感を数字で検証する。

検証する4つの問い:

  1. 直近の連動は本当に過去より強いのか？
     → ローリング相関・過去との比較・Fisher z検定

  2. それは「韓国要因」なのか、単に米国株という共通要因の影響か？
     → 米国株（S&P500・SOX）を統制した偏相関
     相関が高くても、両市場が同じ米国発の材料に反応しているだけなら
     「韓国を見る」意味はない。ここが一番大事な検証。

  3. 韓国は日本に「先行」するのか？
     → リードラグ分析
     韓国(KRX)と日本(TSE)はどちらもUTC+9の09:00〜15:30でほぼ同時に開いている。
     同時刻の相関がいくら高くてもトレードには使えない。
     使えるのは lag>=1（前日までの韓国の値動き→翌日の日本株）だけ。

  4. どの銘柄が韓国に敏感なのか？
     → 銘柄別の対KOSPI感応度（日経の影響を除いた偏相関つき）

使い方:
    python korea_correlation.py                       # 指数レベルの検証
    python korea_correlation.py --start 2018-01-01
    python korea_correlation.py --window 60 --recent 60
    python korea_correlation.py --scan                # 銘柄別の感応度ランキング
    python korea_correlation.py --trades              # 取引をKOSPIレジームで層別
    python korea_correlation.py --no-plot

注意:
    このツールは実データのみを使う。合成データにはフォールバックしない。
    （ランダムウォークの合成データで相関を測っても意味がないため）
    データが取れない場合はエラーで止まる。
"""

from __future__ import annotations

import argparse
import sys
import warnings
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

import corrstats as cs
from config import BacktestConfig, DEFAULT_TICKERS, StrategyConfig
from data import download_index

warnings.filterwarnings("ignore")

OUTPUT_DIR = Path("output")

# ──────────────────────────────────────────
# 対象指数
# ──────────────────────────────────────────

KOSPI = "KOSPI"
NIKKEI = "日経225"

# 日本・韓国の株価指数
MARKET_INDICES: Dict[str, str] = {
    KOSPI: "^KS11",
    "KOSDAQ": "^KQ11",
    NIKKEI: "^N225",
    "TOPIX(ETF)": "1306.T",
}

# 偏相関の統制変数に使う「グローバル共通要因」
# 米国市場は日本・韓国が引けた後に動くため、日本のt日に効くのは
# 米国のt-1セッション（前夜）のリターン。US_LAG_DEFAULT で1日ずらす。
GLOBAL_FACTORS: Dict[str, str] = {
    "S&P500": "^GSPC",
    "SOX(米半導体)": "^SOX",
}

# 参考指標（相関表に載せるが統制変数には使わない）
REFERENCE_SERIES: Dict[str, str] = {
    "USDKRW": "KRW=X",
    "USDJPY": "JPY=X",
}

US_LAG_DEFAULT = 1

# dataviz パレット（categorical slot 1-3 + chrome）
C_BLUE = "#2a78d6"
C_ORANGE = "#eb6834"
C_AQUA = "#1baf7a"
C_MUTED = "#898781"
C_GRID = "#e1e0d9"
C_AXIS = "#c3c2b7"
C_INK = "#0b0b0b"
C_INK2 = "#52514e"
C_SURFACE = "#fcfcfb"


# ──────────────────────────────────────────
# データ取得
# ──────────────────────────────────────────

def fetch_closes(
    tickers: Dict[str, str],
    start: str,
    end: str,
    use_cache: bool = True,
) -> Tuple[Dict[str, pd.Series], List[str]]:
    """指定した銘柄・指数の終値系列を取得する。

    Returns:
        (取得できた {表示名: 終値Series}, 取得できなかった表示名のリスト)
    """
    closes: Dict[str, pd.Series] = {}
    failed: List[str] = []

    for name, ticker in tickers.items():
        df = download_index(ticker, start, end, use_cache=use_cache)
        if df is None or df.empty or "Close" not in df.columns:
            failed.append(f"{name} ({ticker})")
            continue
        closes[name] = df["Close"].astype(float)

    return closes, failed


def build_return_frame(
    closes: Dict[str, pd.Series],
    us_lag: int,
    global_names: List[str],
) -> pd.DataFrame:
    """終値から対数リターンを作り、共通日付で揃えたDataFrameを返す。

    米国系の系列は us_lag 日ずらす（前夜のセッション → 翌日のアジア）。
    """
    returns: Dict[str, pd.Series] = {}
    for name, close in closes.items():
        r = cs.log_returns(close)
        if name in global_names and us_lag:
            r = r.shift(us_lag)
        returns[name] = r
    return cs.align_returns(returns)


# ──────────────────────────────────────────
# 分析 1: 直近の連動は過去より強いか
# ──────────────────────────────────────────

def analyze_recent_vs_history(
    rets: pd.DataFrame,
    base_col: str,
    recent_days: int,
) -> pd.DataFrame:
    """base_col と他の各系列について、直近ウィンドウと過去を比較する。"""
    if len(rets) <= recent_days:
        raise ValueError(
            f"データが不足しています（{len(rets)}日）。"
            f"--recent {recent_days} より長い期間が必要です。"
        )

    recent = rets.iloc[-recent_days:]
    past = rets.iloc[:-recent_days]

    rows = []
    for col in rets.columns:
        if col == base_col:
            continue
        r_recent, p_recent, n_recent = cs.corr_with_p(recent[base_col], recent[col])
        r_past, p_past, n_past = cs.corr_with_p(past[base_col], past[col])
        z, p_diff = cs.compare_correlations(r_recent, n_recent, r_past, n_past)

        rows.append({
            "対象": col,
            "直近相関": r_recent,
            "直近p値": p_recent,
            "直近n": n_recent,
            "過去相関": r_past,
            "過去n": n_past,
            "差": r_recent - r_past,
            "差のz": z,
            "差のp値": p_diff,
        })

    return pd.DataFrame(rows)


def rolling_corr_history(
    rets: pd.DataFrame,
    a: str,
    b: str,
    window: int,
    recent_days: int,
) -> Tuple[pd.Series, float, float]:
    """ローリング相関と、最新値・そのパーセンタイル順位を返す。

    パーセンタイルは「直近ウィンドウが始まる前」までのローリング相関に対して
    測る（自分自身と比較しないため）。
    """
    roll = cs.rolling_correlation(rets[a], rets[b], window).dropna()
    if roll.empty:
        return roll, float("nan"), float("nan")

    latest = float(roll.iloc[-1])
    prior = roll.iloc[:-recent_days] if len(roll) > recent_days else roll
    pct = cs.percentile_rank(latest, prior.tolist())
    return roll, latest, pct


# ──────────────────────────────────────────
# 分析 2: 偏相関（米国要因を除く）
# ──────────────────────────────────────────

def analyze_partial(
    rets: pd.DataFrame,
    a: str,
    b: str,
    control_cols: List[str],
    recent_days: int,
) -> Dict[str, object]:
    """a と b の相関から control_cols の影響を除く。"""
    controls = [c for c in control_cols if c in rets.columns]
    recent = rets.iloc[-recent_days:]

    raw_r, raw_p, raw_n = cs.corr_with_p(recent[a], recent[b])

    if controls:
        part_r, part_p, part_n = cs.partial_correlation(
            recent[a], recent[b], recent[controls]
        )
    else:
        part_r, part_p, part_n = float("nan"), float("nan"), 0

    # 各統制変数を1つずつ入れた場合の偏相関（どれが効いているかを見る）
    per_control = {}
    for c in controls:
        r_c, _, _ = cs.partial_correlation(recent[a], recent[b], recent[[c]])
        per_control[c] = r_c

    return {
        "raw": raw_r,
        "raw_p": raw_p,
        "n": raw_n,
        "partial": part_r,
        "partial_p": part_p,
        "controls": controls,
        "per_control": per_control,
        "explained_share": (
            (raw_r - part_r) / raw_r if np.isfinite(raw_r) and raw_r != 0 and np.isfinite(part_r)
            else float("nan")
        ),
    }


def partial_rolling(
    rets: pd.DataFrame,
    a: str,
    b: str,
    control_cols: List[str],
    window: int,
) -> pd.Series:
    """ローリング偏相関（統制変数の影響を除いた連動の推移）。"""
    controls = [c for c in control_cols if c in rets.columns]
    if not controls:
        return pd.Series(dtype=float)

    values = {}
    for i in range(window, len(rets) + 1):
        chunk = rets.iloc[i - window : i]
        r, _, _ = cs.partial_correlation(chunk[a], chunk[b], chunk[controls])
        values[rets.index[i - 1]] = r
    return pd.Series(values).dropna()


# ──────────────────────────────────────────
# 分析 4: 銘柄別の対KOSPI感応度
# ──────────────────────────────────────────

def scan_ticker_sensitivity(
    kospi_ret: pd.Series,
    nikkei_ret: pd.Series,
    tickers: List[str],
    start: str,
    end: str,
    recent_days: int,
    use_cache: bool = True,
) -> pd.DataFrame:
    """各銘柄の直近の対KOSPI相関・ベータ・日経を除いた偏相関を計算する。

    「日経を除いた偏相関」が高い銘柄は、単に日本市場全体と一緒に動いて
    いるのではなく、韓国固有の材料に反応している銘柄。
    """
    from tqdm import tqdm

    from data import download_stock_data

    rows = []
    for ticker in tqdm(tickers, desc="銘柄スキャン"):
        df = download_stock_data(
            ticker, start, end, use_cache=use_cache, use_synthetic_fallback=False
        )
        if df is None or df.empty:
            continue

        stock_ret = cs.log_returns(df["Close"].astype(float))
        aligned = cs.align_returns({
            "kospi": kospi_ret,
            "nikkei": nikkei_ret,
            "stock": stock_ret,
        })
        if len(aligned) <= recent_days:
            continue

        recent = aligned.iloc[-recent_days:]
        r_kospi, p_kospi, n = cs.corr_with_p(recent["kospi"], recent["stock"])
        r_nikkei, _, _ = cs.corr_with_p(recent["nikkei"], recent["stock"])
        b_kospi = cs.beta(recent["kospi"], recent["stock"])
        r_partial, p_partial, _ = cs.partial_correlation(
            recent["kospi"], recent["stock"], recent[["nikkei"]]
        )

        rows.append({
            "ticker": ticker,
            "対KOSPI相関": r_kospi,
            "対KOSPI p値": p_kospi,
            "対KOSPIベータ": b_kospi,
            "対日経相関": r_nikkei,
            "日経除く偏相関": r_partial,
            "偏相関p値": p_partial,
            "n": n,
        })

    out = pd.DataFrame(rows)
    if not out.empty:
        out = out.sort_values("日経除く偏相関", ascending=False).reset_index(drop=True)
    return out


# ──────────────────────────────────────────
# 分析 5: 取引をKOSPIレジームで層別
# ──────────────────────────────────────────

def _asof_value(series: pd.Series, when: pd.Timestamp) -> float:
    """when より前の最後の値を返す（当日は使わない = 先読み防止）。"""
    prior = series.index[series.index < when]
    if len(prior) == 0:
        return float("nan")
    return float(series.loc[prior[-1]])


def analyze_trade_regimes(
    kospi_close: pd.Series,
    tickers: List[str],
    start: str,
    end: str,
    ma_period: int = 25,
    use_cache: bool = True,
) -> Optional[pd.DataFrame]:
    """バックテストの各取引を、エントリー前日のKOSPIレジームで層別する。

    レジーム判定にはエントリー日より前のKOSPI終値だけを使うので、
    そのままフィルターとして実運用できる形になっている。
    """
    from backtest import Backtester
    from data import download_all_stocks

    stock_data = download_all_stocks(
        tickers, start, end, use_cache=use_cache, use_synthetic_fallback=False
    )
    if not stock_data:
        print("  → 株価データを取得できなかったため取引層別はスキップします")
        return None

    scfg = StrategyConfig()
    scfg.position_size_pct = 1.0 / scfg.max_positions
    bcfg = BacktestConfig(start_date=start, end_date=end)

    result = Backtester(stock_data, scfg, bcfg).run()
    closed = result.closed_trades
    if not closed:
        print("  → 取引が発生しなかったため層別はスキップします")
        return None

    kospi_ma = kospi_close.rolling(ma_period).mean()

    rows = []
    for t in closed:
        close_prev = _asof_value(kospi_close, t.entry_date)
        ma_prev = _asof_value(kospi_ma, t.entry_date)
        if not (np.isfinite(close_prev) and np.isfinite(ma_prev)):
            continue
        rows.append({
            "ticker": t.ticker,
            "entry_date": t.entry_date,
            "return_pct": t.return_pct,
            "pnl": t.pnl,
            "regime": f"KOSPI > {ma_period}日MA" if close_prev > ma_prev else f"KOSPI <= {ma_period}日MA",
        })

    if not rows:
        return None

    trades_df = pd.DataFrame(rows)

    summary = []
    for regime, grp in trades_df.groupby("regime"):
        wins = grp[grp["pnl"] > 0]
        losses = grp[grp["pnl"] < 0]
        gross_profit = wins["pnl"].sum()
        gross_loss = abs(losses["pnl"].sum())
        summary.append({
            "レジーム": regime,
            "取引数": len(grp),
            "勝率": len(wins) / len(grp) if len(grp) else float("nan"),
            "平均リターン": grp["return_pct"].mean(),
            "PF": gross_profit / gross_loss if gross_loss > 0 else float("inf"),
        })

    return pd.DataFrame(summary).sort_values("レジーム").reset_index(drop=True)


# ──────────────────────────────────────────
# 出力: コンソール
# ──────────────────────────────────────────

def _dwidth(text: str) -> int:
    """全角文字を2桁として数えた表示幅。"""
    import unicodedata
    return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in text)


def _pad(text: str, width: int, align: str = "left") -> str:
    """表示幅を揃えてパディングする。

    str.format の幅指定は文字数で数えるため、日本語（全角）が混ざると
    表の列がずれる。表示幅ベースで詰めることで揃える。
    """
    gap = max(width - _dwidth(text), 0)
    return (" " * gap + text) if align == "right" else (text + " " * gap)


def _fmt(v: float, spec: str = "+.3f") -> str:
    if v is None or (isinstance(v, float) and not np.isfinite(v)):
        return "  n/a"
    return format(v, spec)


def _p_mark(p: float) -> str:
    """有意水準の記号。"""
    if not np.isfinite(p):
        return ""
    if p < 0.001:
        return "***"
    if p < 0.01:
        return "**"
    if p < 0.05:
        return "*"
    return ""


def print_report(
    rets: pd.DataFrame,
    comparison: pd.DataFrame,
    roll: pd.Series,
    latest: float,
    pct: float,
    partial_info: Dict[str, object],
    lag_table: pd.DataFrame,
    window: int,
    recent_days: int,
) -> None:
    recent_start = rets.index[-recent_days].date()
    recent_end = rets.index[-1].date()

    print("\n" + "=" * 66)
    print("  韓国指数（KOSPI）との連動チェック")
    print("=" * 66)
    print(f"  データ期間      : {rets.index[0].date()} 〜 {rets.index[-1].date()}（{len(rets)}日）")
    print(f"  直近ウィンドウ  : {recent_start} 〜 {recent_end}（{recent_days}営業日）")
    print(f"  ローリング窓    : {window}営業日")
    print(f"  対象系列        : {', '.join(rets.columns)}")
    print("  有意水準        : * p<0.05, ** p<0.01, *** p<0.001")

    # --- 1 ---
    print("\n" + "-" * 66)
    print("【1】KOSPIとの日次リターン相関: 直近 vs それ以前")
    print("-" * 66)
    print(
        "  " + _pad("対象", 16) + _pad("直近", 9, "right") + _pad("", 4)
        + _pad("過去", 10, "right") + _pad("差", 10, "right")
        + _pad("差のp値", 11, "right")
    )
    for _, r in comparison.iterrows():
        print(
            "  " + _pad(str(r["対象"]), 16)
            + _pad(_fmt(r["直近相関"]), 9, "right") + _pad(_p_mark(r["直近p値"]), 4)
            + _pad(_fmt(r["過去相関"]), 10, "right")
            + _pad(_fmt(r["差"]), 10, "right")
            + _pad(_fmt(r["差のp値"], ".4f"), 11, "right") + _p_mark(r["差のp値"])
        )

    print(f"\n  {NIKKEI}との{window}日ローリング相関")
    print(f"    最新値          : {_fmt(latest)}")
    if np.isfinite(pct):
        print(f"    過去内での位置  : {pct:.1f} パーセンタイル")
    if not roll.empty:
        print(f"    過去の中央値    : {_fmt(float(roll.median()))}")
        print(f"    過去の最大/最小 : {_fmt(float(roll.max()))} / {_fmt(float(roll.min()))}")

    # --- 2 ---
    print("\n" + "-" * 66)
    print("【2】偏相関: 米国要因を除いても韓国との連動は残るか")
    print("-" * 66)
    controls = partial_info["controls"]
    if not controls:
        print("  統制変数が取得できなかったためスキップ")
    else:
        print(f"  統制変数: {', '.join(controls)}（前夜の米国セッション）")
        print(f"  生の相関   (KOSPI vs {NIKKEI}) : "
              f"{_fmt(partial_info['raw'])}{_p_mark(partial_info['raw_p'])}")
        print(f"  偏相関     (米国要因を除去)     : "
              f"{_fmt(partial_info['partial'])}{_p_mark(partial_info['partial_p'])}")
        share = partial_info["explained_share"]
        if np.isfinite(share):
            print(f"  → 生の相関のうち約 {share*100:.0f}% は米国要因で説明できる")
        print("\n  統制変数を1つずつ入れた場合:")
        for c, v in partial_info["per_control"].items():
            print(f"    {c:<16} を除くと {_fmt(v)}")

    # --- 3 ---
    print("\n" + "-" * 66)
    print("【3】リードラグ: KOSPIは日本株に先行するか")
    print("-" * 66)
    print("  韓国(KRX)と日本(TSE)はどちらもUTC+9の09:00〜15:30でほぼ同時に立会。")
    print("  lag=0 は「同時刻」なので売買には使えない。")
    print("  実際に使えるのは lag>=1（前日までの韓国 → 翌日の日本）のみ。")
    print(f"\n  {'lag':>5}  {'相関':>8}   {'p値':>8}   {'意味'}")
    for lag, r in lag_table.iterrows():
        if lag > 0:
            meaning = f"KOSPIが{lag}日先行 → トレード可"
        elif lag == 0:
            meaning = "同時刻 → トレード不可"
        else:
            meaning = f"日本が{-lag}日先行"
        print(
            f"  {lag:>5}  {_fmt(r['corr']):>8}   "
            f"{_fmt(r['p_value'], '.4f'):>8}{_p_mark(r['p_value']):<3} {meaning}"
        )

    tradable = lag_table[lag_table.index >= 1]
    if not tradable.empty:
        best_lag = int(tradable["corr"].abs().idxmax())
        best_r = float(tradable.loc[best_lag, "corr"])
        best_p = float(tradable.loc[best_lag, "p_value"])
        print(f"\n  トレード可能なlagの最大相関: lag={best_lag} で {_fmt(best_r)}"
              f"（p={_fmt(best_p, '.4f')}{_p_mark(best_p)}）")
        if np.isfinite(best_r):
            print(f"  → 説明力 R^2 = {best_r**2*100:.1f}%")


def print_conclusion(
    comparison: pd.DataFrame,
    latest: float,
    pct: float,
    partial_info: Dict[str, object],
    lag_table: pd.DataFrame,
) -> None:
    print("\n" + "=" * 66)
    print("  まとめ")
    print("=" * 66)

    nikkei_row = comparison[comparison["対象"] == NIKKEI]

    # 問い1: 連動は強まったか
    if not nikkei_row.empty:
        row = nikkei_row.iloc[0]
        diff, p_diff = row["差"], row["差のp値"]
        if np.isfinite(p_diff) and p_diff < 0.05 and diff > 0:
            print(f"  ① 連動は強まった: 直近 {_fmt(row['直近相関'])} vs 過去 "
                  f"{_fmt(row['過去相関'])}（p={p_diff:.4f}）→ 体感は正しい")
        elif np.isfinite(p_diff) and p_diff < 0.05 and diff < 0:
            print(f"  ① 連動はむしろ弱まっている: 直近 {_fmt(row['直近相関'])} vs 過去 "
                  f"{_fmt(row['過去相関'])}（p={p_diff:.4f}）")
        else:
            print(f"  ① 連動の強まりは統計的にはっきりしない: 直近 {_fmt(row['直近相関'])} vs "
                  f"過去 {_fmt(row['過去相関'])}（p={_fmt(p_diff, '.4f')}）")
    if np.isfinite(pct):
        print(f"     ローリング相関は過去の {pct:.0f} パーセンタイル水準")

    # 問い2: 韓国固有の要因か
    part = partial_info.get("partial")
    part_p = partial_info.get("partial_p")
    if partial_info.get("controls") and np.isfinite(part):
        if np.isfinite(part_p) and part_p < 0.05 and abs(part) > 0.2:
            print(f"  ② 米国要因を除いても偏相関 {_fmt(part)} が残る（p={part_p:.4f}）")
            print("     → 米国株の裏返しだけでは説明できない、韓国固有の連動がある")
        else:
            print(f"  ② 米国要因を除くと偏相関は {_fmt(part)} まで落ちる")
            print("     → 連動の大半は「日韓がともに米国株に反応している」だけ")
            print("     → KOSPIを見るより米国株・SOXを見たほうが直接的")

    # 問い3: トレードに使えるか
    tradable = lag_table[lag_table.index >= 1]
    if not tradable.empty:
        best_lag = int(tradable["corr"].abs().idxmax())
        best_r = float(tradable.loc[best_lag, "corr"])
        best_p = float(tradable.loc[best_lag, "p_value"])
        contemporaneous = float(lag_table.loc[0, "corr"]) if 0 in lag_table.index else float("nan")

        print(f"  ③ 同時刻の相関 {_fmt(contemporaneous)} に対し、"
              f"先行lagの最大は {_fmt(best_r)}（lag={best_lag}）")
        if np.isfinite(best_p) and best_p < 0.05 and abs(best_r) >= 0.2:
            print(f"     → 前日の韓国の値動きに翌日の日本株への説明力あり"
                  f"（R^2={best_r**2*100:.1f}%）。フィルター候補になる")
        else:
            print("     → 先行性はほぼ無い。連動は『同時に動いている』だけで、")
            print("       KOSPIを見てから翌日売買してもエッジにならない")

    print("\n  注意: 日次リターンには変動の固まり（ボラティリティ・クラスタリング）が")
    print("        あるため、p値は実際よりやや小さく（有意に）出る傾向がある。")
    print("=" * 66)


# ──────────────────────────────────────────
# 出力: グラフ
# ──────────────────────────────────────────

def _setup_font() -> bool:
    """日本語フォントを設定する。使えなければ False を返す。"""
    import logging

    import matplotlib
    from matplotlib import font_manager

    # 日本語フォントにボールドが無い環境で出る findfont 警告を抑える
    logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)

    try:
        import japanize_matplotlib  # noqa: F401
        return True
    except ImportError:
        pass

    candidates = [
        "Noto Sans CJK JP", "Noto Sans JP", "Hiragino Sans", "Hiragino Kaku Gothic Pro",
        "Yu Gothic", "Meiryo", "MS Gothic", "IPAexGothic", "IPAGothic", "TakaoGothic",
        "VL Gothic", "Source Han Sans JP",
    ]
    available = {f.name for f in font_manager.fontManager.ttflist}
    for name in candidates:
        if name in available:
            matplotlib.rcParams["font.family"] = name
            matplotlib.rcParams["axes.unicode_minus"] = False
            return True
    return False


def plot_report(
    rets: pd.DataFrame,
    roll: pd.Series,
    roll_partial: pd.Series,
    lag_table: pd.DataFrame,
    window: int,
    recent_days: int,
    show: bool = False,
) -> Optional[Path]:
    import matplotlib
    if not show:
        matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    has_jp = _setup_font()

    def L(ja: str, en: str) -> str:
        return ja if has_jp else en

    OUTPUT_DIR.mkdir(exist_ok=True)

    fig = plt.figure(figsize=(13, 9), facecolor=C_SURFACE)
    gs = fig.add_gridspec(2, 2, height_ratios=[1.15, 1], hspace=0.42, wspace=0.24)
    ax1 = fig.add_subplot(gs[0, :])
    ax2 = fig.add_subplot(gs[1, 0])
    ax3 = fig.add_subplot(gs[1, 1])

    for ax in (ax1, ax2, ax3):
        ax.set_facecolor(C_SURFACE)
        for side in ("top", "right"):
            ax.spines[side].set_visible(False)
        for side in ("left", "bottom"):
            ax.spines[side].set_color(C_AXIS)
            ax.spines[side].set_linewidth(1.0)
        ax.tick_params(colors=C_MUTED, labelsize=9, length=3, width=1.0)
        ax.grid(True, color=C_GRID, linewidth=1.0, alpha=1.0)
        ax.set_axisbelow(True)

    # ── パネル1: ローリング相関の推移 ──
    ax1.axhline(0, color=C_AXIS, linewidth=1.0)
    ax1.plot(
        roll.index, roll.values, color=C_BLUE, linewidth=2.0,
        label=L(f"KOSPI vs 日経225（{window}日ローリング相関）",
                f"KOSPI vs Nikkei ({window}d rolling corr)"),
    )
    if not roll_partial.empty:
        ax1.plot(
            roll_partial.index, roll_partial.values, color=C_ORANGE, linewidth=2.0,
            label=L("同・米国要因を除いた偏相関", "partial corr (US factor removed)"),
        )

    if len(rets) > recent_days:
        ax1.axvspan(
            rets.index[-recent_days], rets.index[-1],
            color=C_BLUE, alpha=0.07, zorder=0,
        )
        ax1.annotate(
            L(f"直近{recent_days}日", f"last {recent_days}d"),
            xy=(rets.index[-recent_days], 1.0), xytext=(4, -2),
            textcoords="offset points", fontsize=9, color=C_INK2, va="top",
        )

    if not roll.empty:
        med = float(roll.median())
        ax1.axhline(med, color=C_MUTED, linewidth=1.0, linestyle=(0, (4, 3)))
        # 折れ線と重なっても読めるように、地色のbboxを敷く
        ax1.annotate(
            L(f"中央値 {med:+.2f}", f"median {med:+.2f}"),
            xy=(roll.index[0], med), xytext=(2, 5), textcoords="offset points",
            fontsize=9, color=C_INK2,
            bbox=dict(facecolor=C_SURFACE, edgecolor="none", pad=1.5),
        )
        ax1.annotate(
            f"{float(roll.iloc[-1]):+.2f}",
            xy=(roll.index[-1], float(roll.iloc[-1])), xytext=(6, 0),
            textcoords="offset points", fontsize=10, fontweight="bold", color=C_BLUE,
            va="center",
        )

    ax1.set_ylim(-1.05, 1.05)
    ax1.set_title(
        L("日韓の連動はいつ強まったか", "When did the Japan-Korea linkage tighten?"),
        fontsize=13, color=C_INK, loc="left", pad=12,
    )
    ax1.set_ylabel(L("相関係数", "correlation"), fontsize=10, color=C_INK2)
    ax1.legend(frameon=False, fontsize=9, loc="lower left", labelcolor=C_INK2)

    # ── パネル2: リードラグ ──
    lags = list(lag_table.index)
    corrs = lag_table["corr"].values
    colors = [C_BLUE if lag >= 1 else C_MUTED for lag in lags]

    bars = ax2.bar(lags, corrs, color=colors, width=0.68, zorder=2)
    for bar in bars:
        bar.set_linewidth(2.0)
        bar.set_edgecolor(C_SURFACE)
    ax2.axhline(0, color=C_AXIS, linewidth=1.0, zorder=3)

    # ラベルと凡例が収まるように上下に余白をとる
    finite = corrs[np.isfinite(corrs)]
    if len(finite):
        lo, hi = float(np.min(finite)), float(np.max(finite))
        span = max(hi - lo, 0.1)
        ax2.set_ylim(min(lo, 0) - span * 0.18, max(hi, 0) + span * 0.42)

    def _label_bar(lag: int, value: float, color: str, bold: bool) -> None:
        """棒の外側に値を書く（軸外にはみ出さないようにする）。"""
        if not np.isfinite(value):
            return
        offset = 5 if value >= 0 else -13
        ax2.annotate(
            f"{value:+.2f}", xy=(lag, value), xytext=(0, offset),
            textcoords="offset points", ha="center", fontsize=10,
            color=color, **({"fontweight": "bold"} if bold else {}),
        )

    # トレード可能なlagのうち最大のものだけ直接ラベル。
    # ただし全部ノイズ水準のときはラベルを出さない（意味のない数字を強調しない）
    tradable = lag_table[lag_table.index >= 1]
    if not tradable.empty and np.isfinite(tradable["corr"]).any():
        bl = int(tradable["corr"].abs().idxmax())
        bv = float(tradable.loc[bl, "corr"])
        if abs(bv) >= 0.05:
            _label_bar(bl, bv, C_INK, bold=True)
    if 0 in lag_table.index:
        _label_bar(0, float(lag_table.loc[0, "corr"]), C_INK2, bold=False)

    from matplotlib.patches import Patch
    # lag=0 の棒が中央で最も高くなりがちなので、凡例は右上に置いて衝突を避ける
    ax2.legend(
        handles=[
            Patch(facecolor=C_BLUE, label=L("lag>=1: 前日の韓国→翌日の日本（使える）",
                                           "lag>=1: tradable")),
            Patch(facecolor=C_MUTED, label=L("lag<=0: 同時刻/日本先行（使えない）",
                                            "lag<=0: not tradable")),
        ],
        frameon=False, fontsize=8, loc="upper right", labelcolor=C_INK2,
    )
    ax2.set_title(
        L("KOSPIは日本株に先行するか", "Does KOSPI lead Japan?"),
        fontsize=12, color=C_INK, loc="left", pad=10,
    )
    ax2.set_xlabel(L("KOSPIのラグ（営業日）", "KOSPI lag (trading days)"),
                   fontsize=10, color=C_INK2)
    ax2.set_ylabel(L("相関係数", "correlation"), fontsize=10, color=C_INK2)
    ax2.set_xticks(lags)

    # ── パネル3: 直近ウィンドウの散布図 ──
    recent = rets.iloc[-recent_days:]
    x = recent[KOSPI].values * 100
    y = recent[NIKKEI].values * 100

    ax3.axhline(0, color=C_AXIS, linewidth=1.0)
    ax3.axvline(0, color=C_AXIS, linewidth=1.0)
    ax3.scatter(
        x, y, s=34, color=C_BLUE, alpha=0.75,
        edgecolors=C_SURFACE, linewidths=2.0, zorder=2,
    )

    if len(x) >= 3 and np.isfinite(x).all() and np.isfinite(y).all():
        slope, intercept = np.polyfit(x, y, 1)
        xs = np.linspace(x.min(), x.max(), 50)
        ax3.plot(xs, slope * xs + intercept, color=C_ORANGE, linewidth=2.0, zorder=3)
        r = float(np.corrcoef(x, y)[0, 1])
        ax3.annotate(
            L(f"r = {r:+.2f}   傾き = {slope:.2f}", f"r = {r:+.2f}   slope = {slope:.2f}"),
            xy=(0.03, 0.95), xycoords="axes fraction", fontsize=10,
            color=C_INK, va="top", fontweight="bold",
        )

    ax3.set_title(
        L(f"直近{recent_days}日の日次リターン", f"Daily returns, last {recent_days}d"),
        fontsize=12, color=C_INK, loc="left", pad=10,
    )
    ax3.set_xlabel(L("KOSPI 日次リターン (%)", "KOSPI daily return (%)"),
                   fontsize=10, color=C_INK2)
    ax3.set_ylabel(L("日経225 日次リターン (%)", "Nikkei daily return (%)"),
                   fontsize=10, color=C_INK2)

    if not has_jp:
        fig.text(
            0.5, 0.005,
            "(Japanese font not found - labels shown in English. "
            "pip install japanize-matplotlib for Japanese)",
            ha="center", fontsize=8, color=C_MUTED,
        )

    path = OUTPUT_DIR / "korea_correlation.png"
    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor=C_SURFACE)
    if show:
        plt.show()
    plt.close(fig)
    return path


# ──────────────────────────────────────────
# CLI
# ──────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser(
        description="韓国指数（KOSPI）との連動チェック",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("--start", default="2015-01-01", help="分析開始日")
    p.add_argument("--end", default=None, help="分析終了日（省略時は今日）")
    p.add_argument("--window", type=int, default=60, help="ローリング相関の窓（営業日）")
    p.add_argument("--recent", type=int, default=60, help="「直近」とみなす営業日数")
    p.add_argument("--split-date", default=None,
                   help="この日以降を「直近」として過去と比較する（--recent より優先）。"
                        "例: 2026-05-27（韓国レバレッジETF上場日）")
    p.add_argument("--max-lag", type=int, default=5, help="リードラグで見るラグの範囲")
    p.add_argument("--us-lag", type=int, default=US_LAG_DEFAULT,
                   help="米国指数をずらす日数（既定1=前夜のセッション）")
    p.add_argument("--scan", action="store_true", help="銘柄別の対KOSPI感応度を計算する")
    p.add_argument("--trades", action="store_true",
                   help="バックテスト取引をKOSPIレジームで層別する")
    p.add_argument("--tickers", nargs="+", default=None, help="--scan/--trades の対象銘柄")
    p.add_argument("--no-cache", action="store_true", help="キャッシュを使わない")
    p.add_argument("--no-plot", action="store_true", help="グラフを出力しない")
    p.add_argument("--show", action="store_true", help="グラフを画面表示する")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    end = args.end or pd.Timestamp.today().strftime("%Y-%m-%d")
    use_cache = not args.no_cache

    print("\n" + "=" * 66)
    print("  韓国指数との連動チェック — データ取得中")
    print("=" * 66)

    all_tickers = {**MARKET_INDICES, **GLOBAL_FACTORS, **REFERENCE_SERIES}
    closes, failed = fetch_closes(all_tickers, args.start, end, use_cache=use_cache)

    if failed:
        print(f"  取得できなかった系列: {', '.join(failed)}")

    # KOSPIと日経は必須
    missing_required = [n for n in (KOSPI, NIKKEI) if n not in closes]
    if missing_required:
        print(f"\nエラー: 必須の系列を取得できませんでした: {', '.join(missing_required)}")
        print("  このツールは実データのみで動きます（合成データは使いません）。")
        print("  ネットワーク接続とyfinanceの状態を確認してください:")
        print("    python -c \"import yfinance as yf; print(yf.download('^KS11', period='5d'))\"")
        return 1

    print(f"  取得できた系列: {', '.join(closes.keys())}")

    global_names = [n for n in GLOBAL_FACTORS if n in closes]
    rets = build_return_frame(closes, args.us_lag, global_names)

    # 「直近」の定義: 日数指定か、構造変化の日付で分割するか
    recent_days = args.recent
    if args.split_date:
        try:
            split = pd.Timestamp(args.split_date)
        except ValueError:
            print(f"\nエラー: --split-date の日付を解釈できません: {args.split_date}")
            return 1
        recent_days = int((rets.index >= split).sum())
        if recent_days < 20:
            print(f"\nエラー: {split.date()} 以降の営業日が {recent_days} 日しかありません。"
                  "比較に足りないので --split-date を早めてください。")
            return 1
        print(f"  分割日 {split.date()} 以降を「直近」として扱います（{recent_days}営業日）")

    if len(rets) <= max(args.window, recent_days) + 10:
        print(f"\nエラー: 共通営業日が {len(rets)} 日しかありません。"
              f"--start を早めるか --window/--recent を小さくしてください。")
        return 1

    print(f"  日韓共通の営業日: {len(rets)}日"
          f"（{rets.index[0].date()} 〜 {rets.index[-1].date()}）")

    # --- 分析 ---
    comparison = analyze_recent_vs_history(rets, KOSPI, recent_days)
    roll, latest, pct = rolling_corr_history(rets, KOSPI, NIKKEI, args.window, recent_days)
    partial_info = analyze_partial(rets, KOSPI, NIKKEI, global_names, recent_days)
    lag_table = cs.lead_lag_table(
        rets[KOSPI], rets[NIKKEI], max_lag=args.max_lag
    )

    print_report(
        rets, comparison, roll, latest, pct, partial_info,
        lag_table, args.window, recent_days,
    )

    # --- 銘柄別 ---
    if args.scan:
        print("\n" + "-" * 66)
        print("【4】銘柄別の対KOSPI感応度（直近ウィンドウ）")
        print("-" * 66)
        tickers = args.tickers or DEFAULT_TICKERS
        scan_df = scan_ticker_sensitivity(
            rets[KOSPI], rets[NIKKEI], tickers,
            args.start, end, recent_days, use_cache=use_cache,
        )
        if scan_df.empty:
            print("  → 銘柄データを取得できませんでした")
        else:
            OUTPUT_DIR.mkdir(exist_ok=True)
            csv_path = OUTPUT_DIR / "korea_ticker_sensitivity.csv"
            scan_df.to_csv(csv_path, index=False, encoding="utf-8-sig")

            print("\n  ■ 韓国感応度の高い銘柄 上位15（日経の影響を除いた偏相関の順）")
            print(
                "  " + _pad("ticker", 10) + _pad("対KOSPI", 10, "right")
                + _pad("ベータ", 9, "right") + _pad("対日経", 9, "right")
                + _pad("日経除く偏相関", 16, "right")
            )
            for _, r in scan_df.head(15).iterrows():
                print(
                    "  " + _pad(str(r["ticker"]), 10)
                    + _pad(_fmt(r["対KOSPI相関"]), 10, "right")
                    + _pad(_fmt(r["対KOSPIベータ"], "+.2f"), 9, "right")
                    + _pad(_fmt(r["対日経相関"]), 9, "right")
                    + _pad(_fmt(r["日経除く偏相関"]), 16, "right")
                    + _p_mark(r["偏相関p値"])
                )
            sig = scan_df[(scan_df["偏相関p値"] < 0.05) & (scan_df["日経除く偏相関"] > 0)]
            print(f"\n  日経を除いても有意に韓国と連動: {len(sig)} / {len(scan_df)} 銘柄")
            print(f"  → 全銘柄の結果: {csv_path}")

    # --- 取引レジーム ---
    if args.trades:
        print("\n" + "-" * 66)
        print("【5】バックテスト取引のKOSPIレジーム別成績")
        print("-" * 66)
        print("  エントリー日より前のKOSPI終値だけで判定（そのままフィルターに使える形）")
        tickers = args.tickers or DEFAULT_TICKERS
        regime_df = analyze_trade_regimes(
            closes[KOSPI], tickers, args.start, end, use_cache=use_cache,
        )
        if regime_df is not None:
            print(
                "\n  " + _pad("レジーム", 22) + _pad("取引数", 8, "right")
                + _pad("勝率", 10, "right") + _pad("平均リターン", 14, "right")
                + _pad("PF", 8, "right")
            )
            for _, r in regime_df.iterrows():
                print(
                    "  " + _pad(str(r["レジーム"]), 22)
                    + _pad(str(int(r["取引数"])), 8, "right")
                    + _pad(f"{r['勝率']:.1%}", 10, "right")
                    + _pad(f"{r['平均リターン']:.2%}", 14, "right")
                    + _pad(f"{r['PF']:.2f}", 8, "right")
                )
            if len(regime_df) == 2:
                up = regime_df[regime_df["レジーム"].str.contains(">")]
                dn = regime_df[regime_df["レジーム"].str.contains("<=")]
                if not up.empty and not dn.empty:
                    d = up.iloc[0]["平均リターン"] - dn.iloc[0]["平均リターン"]
                    print(f"\n  平均リターンの差: {d:+.2%}")
                    if abs(d) < 0.01:
                        print("  → KOSPIレジームで成績は変わらない。フィルターにする根拠は薄い")
                    else:
                        better = "上" if d > 0 else "下"
                        print(f"  → KOSPIが25日MAより{better}のときのほうが成績が良い。"
                              "フィルター候補")

    print_conclusion(comparison, latest, pct, partial_info, lag_table)

    # --- グラフ ---
    if not args.no_plot:
        roll_partial = pd.Series(dtype=float)
        if global_names:
            print("\nローリング偏相関を計算中...")
            roll_partial = partial_rolling(rets, KOSPI, NIKKEI, global_names, args.window)
        path = plot_report(
            rets, roll, roll_partial, lag_table,
            args.window, recent_days, show=args.show,
        )
        print(f"\nグラフ: {path}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
