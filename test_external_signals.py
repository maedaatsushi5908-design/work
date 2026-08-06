"""
external_signals.py の自己検証テスト

このモジュールで一番怖いのは先読みバイアス（未来のデータを使ってしまうこと）。
バックテストだけ異常に良い結果が出て、実運用で全く再現しない事故につながる。
そこを機械的に潰すためのテスト。

実行:
    python test_external_signals.py
"""

from __future__ import annotations

import sys
import warnings

warnings.filterwarnings("ignore")
if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import numpy as np
import pandas as pd

import external_signals as es
from config import DEFAULT_TICKERS


FAILURES: list[str] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    mark = "✅" if ok else "❌"
    print(f"  {mark} {name}" + (f"  — {detail}" if detail else ""))
    if not ok:
        FAILURES.append(name)


def _fake_series(n: int = 1200, seed: int = 0) -> pd.Series:
    rng = np.random.default_rng(seed)
    idx = pd.bdate_range("2018-01-01", periods=n)
    vals = 100 * np.exp(np.cumsum(rng.normal(0.0003, 0.012, n)))
    return pd.Series(vals, index=idx)


# ──────────────────────────────────────────
# 1. 設定の整合性
# ──────────────────────────────────────────

def test_sector_coverage() -> None:
    print("\n[1] 業種マッピングの整合性")

    unmapped = [t for t in DEFAULT_TICKERS if t not in es.TICKER_SECTOR]
    check("全銘柄に業種が割り当てられている",
          not unmapped,
          f"未割当 {len(unmapped)}件: {unmapped[:5]}" if unmapped else
          f"{len(DEFAULT_TICKERS)}銘柄")

    unknown = {s for s in es.TICKER_SECTOR.values() if s not in es.SECTOR_DRIVERS}
    check("業種が全て SECTOR_DRIVERS に定義されている",
          not unknown, f"未定義: {unknown}" if unknown else "")

    bad = []
    for sector, drivers in es.SECTOR_DRIVERS.items():
        total = sum(d.weight for d in drivers)
        if abs(total - 1.0) > 1e-9:
            bad.append(f"{sector}={total:.3f}")
    check("各業種のドライバー重みの合計が1.0", not bad, ", ".join(bad))

    missing = []
    for sector, drivers in es.SECTOR_DRIVERS.items():
        for d in drivers:
            if d.series not in es.SERIES_LABELS:
                missing.append(f"{sector}:{d.series}")
    check("全ドライバーが SERIES_LABELS に定義済み", not missing, ", ".join(missing))

    unused = set(es.SERIES_LABELS) - {
        d.series for ds in es.SECTOR_DRIVERS.values() for d in ds
    }
    check("未使用の外部系列が無い", not unused, f"未使用: {unused}" if unused else "")

    bad_sign = [d.series for ds in es.SECTOR_DRIVERS.values() for d in ds if d.sign not in (1, -1)]
    check("符号が +1 / -1 のみ", not bad_sign, ", ".join(bad_sign))


# ──────────────────────────────────────────
# 2. 前年同期比の計算が正しいか
# ──────────────────────────────────────────

def test_yoy_series() -> None:
    print("\n[2] 前年同期比の計算")

    s = _fake_series(seed=1)
    panel = es.yoy_series(s)

    # ナイーブな実装と一致するか（最適化前の定義と同じ値になること）
    i = len(s) - 1
    recent = s.iloc[i - es.QUARTER_DAYS + 1: i + 1].mean()
    year_ago_end = i - es.YEAR_DAYS
    prior = s.iloc[year_ago_end - es.QUARTER_DAYS + 1: year_ago_end + 1].mean()
    expected = recent / prior - 1.0

    got = float(panel.iloc[-1])
    check("ローリング実装がナイーブ実装と一致",
          abs(got - expected) < 1e-9, f"got={got:.6f} expected={expected:.6f}")

    # 定数列なら前年同期比はゼロ
    flat = pd.Series(50.0, index=s.index)
    z = es.yoy_series(flat).dropna()
    check("定数列の前年同期比は0", bool(len(z)) and abs(float(z.iloc[-1])) < 1e-12)

    # 立ち上がり期間は NaN（履歴不足で値を作らない）
    head = es.yoy_series(s).iloc[: es.YEAR_DAYS + es.QUARTER_DAYS - 2]
    check("履歴不足の期間は NaN", bool(head.isna().all()))


# ──────────────────────────────────────────
# 3. 先読みバイアスが無いこと（最重要）
# ──────────────────────────────────────────

def test_no_lookahead() -> None:
    print("\n[3] 先読みバイアスの検証")

    full = {name: _fake_series(seed=i) for i, name in enumerate(es.SERIES_LABELS)}
    asof = full["JPY=X"].index[900]

    # 過去だけを渡した場合と、未来まで渡した場合でスコアが変わらないこと。
    # 変わるなら未来のデータを参照している。
    truncated = {k: v[v.index <= asof] for k, v in full.items()}

    ch_full = es.build_change_panel(full)
    ch_trunc = es.build_change_panel(truncated)

    a = es.score_ticker("8035.T", asof, ch_full)
    b = es.score_ticker("8035.T", asof, ch_trunc)
    ok = a is not None and b is not None and abs(a.score - b.score) < 1e-12
    check("fixed モード: 未来データを足してもスコアが不変",
          ok, f"{a.score:.8f} vs {b.score:.8f}" if a and b else "スコア計算失敗")

    px = _fake_series(seed=99)
    fa = es.score_ticker_fitted("8035.T", asof, ch_full, px)
    fb = es.score_ticker_fitted("8035.T", asof, ch_trunc, px[px.index <= asof])
    ok = fa is not None and fb is not None and abs(fa.score - fb.score) < 1e-9
    check("fitted モード: 未来データを足してもスコアが不変",
          ok, f"{fa.score:.8f} vs {fb.score:.8f}" if fa and fb else "スコア計算失敗")

    # 感応度推定の学習データが asof を超えないこと
    drivers = es.SECTOR_DRIVERS["semiconductor"]
    coefs = es.estimate_sensitivity(px, ch_full, drivers, asof)
    coefs_trunc = es.estimate_sensitivity(px[px.index <= asof], ch_trunc, drivers, asof)
    ok = bool(coefs) and bool(coefs_trunc) and all(
        abs(coefs[k] - coefs_trunc[k]) < 1e-9 for k in coefs
    )
    check("感応度の推定係数が未来データに依存しない", ok)

    # lookup_change は asof より後の値を絶対に返さない
    s = full["JPY=X"]
    ch = es.build_change_panel({"JPY=X": s})["JPY=X"]
    t = s.index[800]
    val = es.lookup_change({"JPY=X": ch}, "JPY=X", t)
    expected = float(ch[ch.index <= t].iloc[-1])
    check("lookup_change が asof 以前の値のみ返す",
          val is not None and abs(val - expected) < 1e-12)


# ──────────────────────────────────────────
# 4. スコアの性質
# ──────────────────────────────────────────

def test_score_behaviour() -> None:
    print("\n[4] スコアの挙動")

    series = {name: _fake_series(seed=i) for i, name in enumerate(es.SERIES_LABELS)}
    changes = es.build_change_panel(series)
    asof = series["JPY=X"].index[-1]

    sc = es.score_ticker("7203.T", asof, changes)
    check("自動車セクターのスコアが計算できる", sc is not None)

    if sc:
        # 加重合成が内訳の合計と一致する（正規化を考慮）
        total = sum(c for _, _, c in sc.breakdown)
        check("スコアが内訳の合計と整合",
              abs(sc.score - total / sc.coverage) < 1e-12)

        # 符号 -1 のドライバーは逆向きに寄与する
        cl = [(chg, c) for n, chg, c in sc.breakdown if n == "CL=F"]
        ok = bool(cl) and (cl[0][0] * cl[0][1] <= 0)
        check("原油（sign=-1）の寄与が前年同期比と逆符号", ok)

    # 未登録銘柄は domestic 扱いになる
    check("未登録銘柄は domestic にフォールバック",
          es.get_sector("9999.T") == "domestic")

    # 欠損系列があってもスコアが出る（正規化が効いている）
    partial = {k: v for k, v in changes.items() if k != "^SOX"}
    sc2 = es.score_ticker("6758.T", asof, partial)
    check("一部の外部系列が欠損してもスコアが計算できる",
          sc2 is not None and sc2.coverage < 1.0,
          f"coverage={sc2.coverage:.2f}" if sc2 else "")

    # 全系列欠損なら None を返す（無言でゼロを返さない）
    check("全ドライバー欠損なら None を返す",
          es.score_ticker("7203.T", asof, {}) is None)


# ──────────────────────────────────────────
# 5. ランキング
# ──────────────────────────────────────────

def test_ranking() -> None:
    print("\n[5] ランキング")

    series = {name: _fake_series(seed=i) for i, name in enumerate(es.SERIES_LABELS)}
    changes = es.build_change_panel(series)
    asof = series["JPY=X"].index[-1]

    df = es.rank_universe(DEFAULT_TICKERS, asof, changes, mode="fixed")
    check("全銘柄のスコアが計算できる",
          len(df) == len(DEFAULT_TICKERS), f"{len(df)}/{len(DEFAULT_TICKERS)}銘柄")
    check("スコア降順にソートされている",
          bool(df["score"].is_monotonic_decreasing))

    # fixed モードは業種内で同一スコアになる（既知の制約）
    per_sector = df.groupby("sector")["score"].nunique()
    check("fixed モードでは業種内のスコアが同一（既知の制約）",
          bool((per_sector == 1).all()))

    prices = {t: _fake_series(seed=1000 + i) for i, t in enumerate(DEFAULT_TICKERS[:20])}
    df2 = es.rank_universe(DEFAULT_TICKERS[:20], asof, changes, prices=prices, mode="fitted")
    if len(df2) > 1:
        check("fitted モードでは銘柄ごとにスコアが分かれる",
              len(df2["score"].unique()) > 1,
              f"{len(df2['score'].unique())}通り/{len(df2)}銘柄")
    else:
        check("fitted モードで十分な銘柄数が処理できる", False, f"{len(df2)}銘柄のみ")


# ──────────────────────────────────────────
# 6. e-Stat クライアント
# ──────────────────────────────────────────

def test_estat() -> None:
    print("\n[6] e-Stat クライアント")

    # APIキー未設定なら例外を投げずに None を返すこと
    # （設定していないユーザーの環境でスキャナーが落ちないため）
    import os
    saved = os.environ.pop("ESTAT_APP_ID", None)
    try:
        check("ESTAT_APP_ID 未設定なら例外を出さず None を返す",
              es.fetch_estat_series("dummy-id") is None)
    finally:
        if saved:
            os.environ["ESTAT_APP_ID"] = saved


def main() -> None:
    print("=" * 62)
    print("  external_signals.py 自己検証テスト")
    print("=" * 62)

    test_sector_coverage()
    test_yoy_series()
    test_no_lookahead()
    test_score_behaviour()
    test_ranking()
    test_estat()

    print("\n" + "=" * 62)
    if FAILURES:
        print(f"  ❌ {len(FAILURES)}件が失敗しました:")
        for f in FAILURES:
            print(f"     - {f}")
        sys.exit(1)
    print("  ✅ すべて成功しました")
    print("=" * 62)


if __name__ == "__main__":
    main()
