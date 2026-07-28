"""
korea_correlation.py のエンドツーエンド検証

⚠️ ここで使うのは「正解が分かっている人工データ」であって市場データではない。
   目的はツールが設計どおりに判定できるかの確認だけで、
   出てくる数字を相場の実態として読んではいけない。

3つのシナリオで、仕込んだ構造をツールが正しく検出できるか確かめる:

  A. 直近だけ日韓の連動が強まり、それが米国要因では説明できない
     → 「相関上昇を検出」かつ「偏相関が残る」と判定できるか

  B. 直近の連動は強いが、その全部が米国という共通要因由来
     → 「偏相関はほぼ消える」と判定できるか（見せかけの相関を見抜けるか）

  C. KOSPIが1日先行する構造
     → リードラグで lag=1 が最大になるか

    python test_korea_correlation.py
"""

from __future__ import annotations

import math
import sys
from pathlib import Path

import numpy as np
import pandas as pd

import corrstats as cs
import data as data_mod
import korea_correlation as kc

PASS = 0
FAIL = 0

RECENT = 60
HISTORY = 900


def check(name: str, ok: bool, detail: str = "") -> None:
    global PASS, FAIL
    if ok:
        PASS += 1
        print(f"  [OK]   {name}" + (f" — {detail}" if detail else ""))
    else:
        FAIL += 1
        print(f"  [FAIL] {name}" + (f" — {detail}" if detail else ""))


# ──────────────────────────────────────────
# 人工データの生成
# ──────────────────────────────────────────

def _to_ohlcv(returns: pd.Series, base: float = 1000.0) -> pd.DataFrame:
    """リターン系列からOHLCVのDataFrameを作る。"""
    r = returns.fillna(0.0)
    close = base * np.exp(r.cumsum())
    rng = np.random.default_rng(abs(hash(returns.name)) % (2**32))
    noise = rng.uniform(0.001, 0.006, len(close))
    return pd.DataFrame(
        {
            "Open": close * (1 - noise / 2),
            "High": close * (1 + noise),
            "Low": close * (1 - noise),
            "Close": close,
            "Volume": rng.integers(1_000_000, 5_000_000, len(close)).astype(float),
        },
        index=close.index,
    )


def make_scenario(
    shared_recent: float,
    shared_past: float,
    us_load: float,
    kospi_lead: int = 0,
    n: int = HISTORY,
    seed: int = 11,
) -> dict:
    """指定した構造を持つ指数データを作る。

    Args:
        shared_recent: 直近RECENT日の「日韓固有の共通因子」の強さ
        shared_past  : それ以前の同因子の強さ
        us_load      : 米国要因の効き具合（日韓ともに同じだけ効く）
        kospi_lead   : >0 なら日本がKOSPIのこの日数分の遅れで反応する
    """
    rng = np.random.default_rng(seed)
    idx = pd.bdate_range("2021-01-04", periods=n)

    us = pd.Series(rng.standard_normal(n) * 0.01, index=idx, name="us")

    # 米国は前夜のセッションが翌日のアジアに効く → t-1 の値を使う
    us_prev = us.shift(1).fillna(0.0)

    shared = pd.Series(rng.standard_normal(n) * 0.01, index=idx)
    weights = np.where(np.arange(n) >= n - RECENT, shared_recent, shared_past)
    shared_w = shared * weights

    eps_k = pd.Series(rng.standard_normal(n) * 0.008, index=idx)
    eps_n = pd.Series(rng.standard_normal(n) * 0.008, index=idx)

    kospi = us_load * us_prev + shared_w + eps_k
    if kospi_lead > 0:
        # 日本はKOSPIの kospi_lead 日前の動きに反応する
        nikkei = us_load * us_prev + 0.9 * kospi.shift(kospi_lead).fillna(0.0) + eps_n
    else:
        nikkei = us_load * us_prev + shared_w + eps_n

    kospi.name, nikkei.name, us.name = "kospi", "nikkei", "us"

    return {
        "^KS11": _to_ohlcv(kospi, 2500),
        "^N225": _to_ohlcv(nikkei, 38000),
        "^GSPC": _to_ohlcv(us, 5000),
        "^SOX": _to_ohlcv(us * 1.4 + pd.Series(rng.standard_normal(n) * 0.004, index=idx), 5500),
    }


def install_fake_data(frames: dict) -> None:
    """korea_correlation / data のダウンロード関数を人工データに差し替える。"""

    def fake_download_index(ticker, start, end, use_cache=True):
        return frames.get(ticker)

    def fake_download_stock(ticker, start, end, use_cache=True, use_synthetic_fallback=True):
        return frames.get(ticker)

    def fake_download_all(tickers, start, end, use_cache=True, use_synthetic_fallback=True):
        return {t: frames[t] for t in tickers if t in frames}

    kc.download_index = fake_download_index
    data_mod.download_stock_data = fake_download_stock
    data_mod.download_all_stocks = fake_download_all


def run_tool(argv: list) -> int:
    """main() を指定引数で実行する。"""
    old = sys.argv
    sys.argv = ["korea_correlation.py"] + argv
    try:
        return kc.main()
    finally:
        sys.argv = old


def analyze(frames: dict):
    """main() を通さず分析部分だけ呼んで結果を取り出す。"""
    closes = {
        kc.KOSPI: frames["^KS11"]["Close"],
        kc.NIKKEI: frames["^N225"]["Close"],
        "S&P500": frames["^GSPC"]["Close"],
        "SOX(米半導体)": frames["^SOX"]["Close"],
    }
    global_names = ["S&P500", "SOX(米半導体)"]
    rets = kc.build_return_frame(closes, kc.US_LAG_DEFAULT, global_names)
    comparison = kc.analyze_recent_vs_history(rets, kc.KOSPI, RECENT)
    roll, latest, pct = kc.rolling_corr_history(rets, kc.KOSPI, kc.NIKKEI, 60, RECENT)
    partial = kc.analyze_partial(rets, kc.KOSPI, kc.NIKKEI, global_names, RECENT)
    lag = cs.lead_lag_table(rets[kc.KOSPI], rets[kc.NIKKEI], max_lag=5)
    return rets, comparison, roll, latest, pct, partial, lag


# ──────────────────────────────────────────
# シナリオ A: 直近だけ本物の連動が強まる
# ──────────────────────────────────────────

def test_scenario_a() -> None:
    print("\n[A] 直近だけ日韓固有の連動が強まる（米国要因では説明できない）")
    frames = make_scenario(shared_recent=2.2, shared_past=0.1, us_load=0.6, seed=101)
    rets, comparison, roll, latest, pct, partial, lag = analyze(frames)

    row = comparison[comparison["対象"] == kc.NIKKEI].iloc[0]
    check("直近の相関が過去より高い", row["直近相関"] > row["過去相関"],
          f"直近 {row['直近相関']:+.3f} vs 過去 {row['過去相関']:+.3f}")
    check("相関上昇が有意 (p<0.05)", row["差のp値"] < 0.05, f"p={row['差のp値']:.5f}")
    check("ローリング相関が高パーセンタイル", pct > 90, f"{pct:.1f}パーセンタイル")
    check("米国要因を除いても偏相関が残る", partial["partial"] > 0.5,
          f"生 {partial['raw']:+.3f} → 偏 {partial['partial']:+.3f}")
    check("偏相関が有意", partial["partial_p"] < 0.001, f"p={partial['partial_p']:.2e}")
    check("同時刻の相関が最大", int(lag["corr"].idxmax()) == 0,
          f"最大は lag={int(lag['corr'].idxmax())}")
    check("先行lagには相関がない", abs(lag.loc[1, "corr"]) < 0.25,
          f"lag=1 で {lag.loc[1, 'corr']:+.3f}")


# ──────────────────────────────────────────
# シナリオ B: 見せかけの相関（全部が米国要因）
# ──────────────────────────────────────────

def test_scenario_b() -> None:
    print("\n[B] 連動は強いが全部が米国要因（見せかけの相関を見抜けるか）")
    # 日韓固有の共通因子をゼロにし、米国要因だけを非常に強くする
    frames = make_scenario(shared_recent=0.0, shared_past=0.0, us_load=4.0, seed=202)
    rets, comparison, roll, latest, pct, partial, lag = analyze(frames)

    check("生の相関は高く出る", partial["raw"] > 0.7, f"生 {partial['raw']:+.3f}")
    check("米国要因を除くと偏相関はほぼ消える", abs(partial["partial"]) < 0.3,
          f"偏 {partial['partial']:+.3f}")
    share = partial["explained_share"]
    check("大半が米国要因で説明されると判定", share > 0.6, f"説明割合 {share*100:.0f}%")


# ──────────────────────────────────────────
# シナリオ C: KOSPIが1日先行する
# ──────────────────────────────────────────

def test_scenario_c() -> None:
    print("\n[C] KOSPIが1日先行する構造")
    frames = make_scenario(
        shared_recent=0.5, shared_past=0.5, us_load=0.3, kospi_lead=1, seed=303
    )
    rets, comparison, roll, latest, pct, partial, lag = analyze(frames)

    best = int(lag["corr"].idxmax())
    check("最大相関のlagが1", best == 1, f"最大は lag={best}, r={lag.loc[best, 'corr']:+.3f}")
    check("lag=1 が同時刻より強い", lag.loc[1, "corr"] > lag.loc[0, "corr"],
          f"lag1 {lag.loc[1, 'corr']:+.3f} > lag0 {lag.loc[0, 'corr']:+.3f}")
    check("lag=1 が有意", lag.loc[1, "p_value"] < 0.001, f"p={lag.loc[1, 'p_value']:.2e}")


# ──────────────────────────────────────────
# CLI 一式が落ちずに動くか
# ──────────────────────────────────────────

def test_cli_end_to_end() -> None:
    print("\n[D] CLI・レポート・グラフ・銘柄スキャン・取引層別が通るか")
    frames = make_scenario(shared_recent=2.0, shared_past=0.2, us_load=0.6, seed=404)

    # 銘柄スキャン用に個別株を追加（KOSPI寄りとKOSPI無関係の2種類）
    rng = np.random.default_rng(9)
    kospi_ret = cs.log_returns(frames["^KS11"]["Close"])
    n = len(kospi_ret)
    idx = kospi_ret.index
    for i, load in enumerate([1.3, 0.9, 0.0, 0.0]):
        r = load * kospi_ret.fillna(0.0) + pd.Series(rng.standard_normal(n) * 0.012, index=idx)
        r.name = f"stock{i}"
        frames[f"999{i}.T"] = _to_ohlcv(r, 3000)

    install_fake_data(frames)
    fake_tickers = ["9990.T", "9991.T", "9992.T", "9993.T"]

    png = Path("output/korea_correlation.png")
    csv = Path("output/korea_ticker_sensitivity.csv")
    for f in (png, csv):
        if f.exists():
            f.unlink()

    rc = run_tool([
        "--start", "2021-01-04", "--end", "2026-07-27",
        "--window", "60", "--recent", str(RECENT),
        "--scan", "--trades", "--tickers", *fake_tickers,
    ])
    check("終了コード0", rc == 0, f"rc={rc}")
    check("グラフが出力される", png.exists() and png.stat().st_size > 20_000,
          f"{png} {png.stat().st_size if png.exists() else 0} bytes")
    check("銘柄別CSVが出力される", csv.exists(), str(csv))

    if csv.exists():
        df = pd.read_csv(csv)
        check("CSVに全銘柄", len(df) == 4, f"{len(df)}行")
        # KOSPI連動を仕込んだ銘柄が上位に来るか
        top2 = set(df.head(2)["ticker"])
        check("韓国連動銘柄が上位2に入る", top2 == {"9990.T", "9991.T"}, f"上位2 = {top2}")

    # --split-date で構造変化日を指定して分割できること
    # 人工データは末尾RECENT日で連動が変わるので、その境目の日付を渡す
    idx = cs.log_returns(frames["^KS11"]["Close"]).dropna().index
    split = idx[-RECENT].strftime("%Y-%m-%d")
    rc_split = run_tool([
        "--start", "2021-01-04", "--split-date", split, "--no-plot",
    ])
    check("--split-date で実行できる", rc_split == 0, f"split={split}, rc={rc_split}")

    # 分割日以降が短すぎる場合はエラーで止まる
    rc_short = run_tool([
        "--start", "2021-01-04", "--split-date", idx[-3].strftime("%Y-%m-%d"), "--no-plot",
    ])
    check("分割日以降が短すぎればエラー", rc_short == 1, f"rc={rc_short}")

    # 不正な日付を渡した場合もクラッシュせずエラー終了
    rc_bad = run_tool(["--start", "2021-01-04", "--split-date", "not-a-date", "--no-plot"])
    check("不正な日付はエラー終了", rc_bad == 1, f"rc={rc_bad}")

    # 必須データが欠けたらエラーで止まること（合成データに逃げない）
    kc.download_index = lambda ticker, start, end, use_cache=True: (
        frames.get(ticker) if ticker != "^KS11" else None
    )
    rc2 = run_tool(["--start", "2021-01-04", "--no-plot"])
    check("KOSPIが取れなければ終了コード1", rc2 == 1, f"rc={rc2}")


def test_no_lookahead() -> None:
    print("\n[E] 先読み防止（レジーム判定に当日以降を使っていないか）")
    idx = pd.bdate_range("2024-01-01", periods=10)
    s = pd.Series(range(10), index=idx, dtype=float)

    # 3番目の日付を渡したら、その「前日」の値2が返るべき（3ではない）
    v = kc._asof_value(s, idx[3])
    check("当日を含まず前日の値を返す", v == 2.0, f"got {v}, 当日値は3.0")

    # 最初の日付より前は情報がないのでNaN
    check("履歴がなければNaN", math.isnan(kc._asof_value(s, idx[0])), "")


def main() -> int:
    print("=" * 66)
    print("  korea_correlation.py エンドツーエンド検証")
    print("  ※ 人工データによる動作確認。相場の実態を示すものではない")
    print("=" * 66)

    test_scenario_a()
    test_scenario_b()
    test_scenario_c()
    test_no_lookahead()
    test_cli_end_to_end()

    print("\n" + "=" * 66)
    print(f"  成功 {PASS} / 失敗 {FAIL}")
    print("=" * 66)
    return 1 if FAIL else 0


if __name__ == "__main__":
    raise SystemExit(main())
