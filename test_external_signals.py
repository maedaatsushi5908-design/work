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

    import os

    # APIキー未設定なら例外を投げずに None を返すこと
    # （設定していないユーザーの環境でスキャナーが落ちないため）
    saved = os.environ.pop("ESTAT_APP_ID", None)
    try:
        check("ESTAT_APP_ID 未設定なら例外を出さず None を返す",
              es.fetch_estat_series("dummy-id") is None)
        check("ESTAT_APP_ID 未設定なら検索も None を返す",
              es.search_estat_tables("建設") is None)
        check("ESTAT_APP_ID 未設定でも load_estat_series は空辞書",
              es.load_estat_series() == {})
    finally:
        if saved:
            os.environ["ESTAT_APP_ID"] = saved

    # 時間コードのパース
    check("時間コード 2024000103 → 2024年3月",
          es._parse_estat_time("2024000103") == pd.Timestamp("2024-03-01"))
    check("時間コード 202403 → 2024年3月",
          es._parse_estat_time("202403") == pd.Timestamp("2024-03-01"))
    check("不正な時間コードは None",
          es._parse_estat_time("abc") is None and es._parse_estat_time("2024000199") is None)

    # 設定ファイル: stats_data_id が空のものは無効として除外されること
    import json
    import tempfile
    from pathlib import Path as _Path

    original = es.ESTAT_CONFIG_FILE
    with tempfile.TemporaryDirectory() as tmp:
        cfg = _Path(tmp) / "estat_series.json"
        cfg.write_text(json.dumps({
            "estat:construction_orders": {"stats_data_id": "0003XXXX", "lag_months": 2},
            "estat:retail_sales": {"stats_data_id": ""},
        }, ensure_ascii=False), encoding="utf-8")
        es.ESTAT_CONFIG_FILE = cfg
        try:
            loaded = es.load_estat_config()
            check("statsDataId が空の系列は無効化される",
                  set(loaded) == {"estat:construction_orders"}, f"読込: {set(loaded)}")
        finally:
            es.ESTAT_CONFIG_FILE = original

    # 設定ファイルが無い環境でも落ちない
    with tempfile.TemporaryDirectory() as tmp:
        es.ESTAT_CONFIG_FILE = _Path(tmp) / "missing.json"
        try:
            check("設定ファイルが無くても空辞書を返す", es.load_estat_config() == {})
        finally:
            es.ESTAT_CONFIG_FILE = original


def test_monthly_series() -> None:
    print("\n[7] 政府統計（月次）の扱い")

    idx = pd.date_range("2018-01-01", periods=60, freq="MS")
    vals = pd.Series(np.arange(60, dtype=float) + 100, index=idx)

    m = es.yoy_series_monthly(vals)
    # 直近3ヶ月平均 vs 12ヶ月前の3ヶ月平均
    recent = vals.iloc[-3:].mean()
    prior = vals.iloc[-15:-12].mean()
    check("月次の前年同期比がナイーブ実装と一致",
          abs(float(m.iloc[-1]) - (recent / prior - 1.0)) < 1e-12)

    check("最初の14ヶ月は NaN（履歴不足）", bool(m.iloc[:14].isna().all()))

    # build_change_panel が estat 系列を月次として扱うこと
    panel = es.build_change_panel({"estat:construction_orders": vals})
    direct = es.yoy_series_monthly(vals).dropna()
    ok = len(panel["estat:construction_orders"]) == len(direct)
    check("build_change_panel が estat 系列を月次計算に振り分ける", ok)

    # 日次計算を誤って適用していないこと（日次だと履歴不足で空になる）
    daily = es.yoy_series(vals).dropna()
    check("月次系列に日次計算を使うと空になる（振り分けの必要性）", len(daily) == 0)


def test_estat_publication_lag() -> None:
    print("\n[8] 政府統計の公表ラグ（先読み防止）")

    # 調査対象月 → 公表日 へのインデックスずらしが効いているか。
    # fetch_estat_series は通信するのでレスポンス相当のデータで直接検証する。
    idx = pd.date_range("2020-01-01", periods=36, freq="MS")
    raw = pd.Series(np.arange(36, dtype=float), index=idx)

    lag = 2
    shifted = raw.copy()
    shifted.index = shifted.index + pd.DateOffset(months=lag)

    check("公表ラグの分だけインデックスが後ろにずれる",
          shifted.index[0] == pd.Timestamp("2020-03-01"))

    # ずらした後は「調査対象月の時点」では値を引けないこと
    panel = {"estat:x": shifted}
    v_at_survey = es.lookup_change(panel, "estat:x", pd.Timestamp("2020-01-15"))
    check("調査対象月の時点ではまだ値を参照できない", v_at_survey is None)

    v_after = shifted.asof(pd.Timestamp("2020-03-15"))
    check("公表後は値を参照できる", pd.notna(v_after) and float(v_after) == 0.0)


def test_score_history_and_filter() -> None:
    print("\n[9] スコア時系列とバックテストフィルター")

    from backtest import Backtester
    from config import BacktestConfig, StrategyConfig
    from synthetic_data import generate_all_stocks

    series = {name: _fake_series(n=2000, seed=i) for i, name in enumerate(es.SERIES_LABELS)
              if not name.startswith(es.ESTAT_PREFIX)}
    changes = es.build_change_panel(series)

    tickers = ["8035.T", "7203.T", "6758.T", "5019.T", "8306.T"]
    data = generate_all_stocks(tickers, "2016-01-01", "2024-12-31")
    prices = {t: df["Close"].dropna() for t, df in data.items()}

    hist = es.build_score_history(tickers, prices, changes, mode="fixed")
    check("スコア時系列が構築できる", len(hist) == len(tickers), f"{len(hist)}/{len(tickers)}銘柄")

    if hist:
        s = next(iter(hist.values()))
        check("スコア時系列が時刻順にソートされている", bool(s.index.is_monotonic_increasing))

        # 先読み検証: 過去だけで作った時系列が、全期間で作ったものと前半一致する
        cut = prices["8035.T"].index[1500]
        trunc_prices = {t: p[p.index <= cut] for t, p in prices.items()}
        trunc_changes = es.build_change_panel({k: v[v.index <= cut] for k, v in series.items()})
        hist_trunc = es.build_score_history(tickers, trunc_prices, trunc_changes, mode="fixed")

        full = hist["8035.T"]
        part = hist_trunc.get("8035.T")
        ok = part is not None and len(part) > 0
        if ok:
            common = full.index.intersection(part.index)
            ok = len(common) > 0 and np.allclose(
                full.loc[common].values, part.loc[common].values, atol=1e-9
            )
        check("スコア時系列に先読みが無い（過去だけで作っても同値）", ok)

    # フィルターの動作: 閾値を極端に上げると全シグナルが弾かれること
    bcfg = BacktestConfig(start_date="2016-01-01", end_date="2024-12-31")

    base = Backtester(data, StrategyConfig(), bcfg).run()

    scfg_block = StrategyConfig(
        external_filter=True, external_min_score=9.99, external_allow_missing=False
    )
    bt_block = Backtester(data, scfg_block, bcfg, hist)
    blocked = bt_block.run()

    check("閾値を極端に上げると取引が発生しない",
          blocked.total_trades == 0 and bt_block.filtered_out > 0,
          f"取引{blocked.total_trades}件 / 見送り{bt_block.filtered_out}件")

    scfg_pass = StrategyConfig(
        external_filter=True, external_min_score=-9.99, external_allow_missing=True
    )
    bt_pass = Backtester(data, scfg_pass, bcfg, hist)
    passed = bt_pass.run()

    check("閾値を極端に下げるとフィルター無しと同じ結果になる",
          passed.total_trades == base.total_trades,
          f"{passed.total_trades} vs {base.total_trades}件")

    # スコアが無い銘柄の扱い
    scfg_missing = StrategyConfig(
        external_filter=True, external_min_score=-9.99, external_allow_missing=False
    )
    bt_missing = Backtester(data, scfg_missing, bcfg, {})
    missing = bt_missing.run()
    check("スコア未取得かつ allow_missing=False なら全て見送る",
          missing.total_trades == 0)

    bt_default = Backtester(data, StrategyConfig(), bcfg, None)
    default = bt_default.run()
    check("external_filter=False ならスコア未指定でも従来通り動く",
          default.total_trades == base.total_trades and bt_default.filtered_out == 0)


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
    test_monthly_series()
    test_estat_publication_lag()
    test_score_history_and_filter()

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
