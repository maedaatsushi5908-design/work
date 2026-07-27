"""
移動平均線フィルターの単体テスト

ネットワーク不要・決定論的な合成データで、フィルターの各条件が
意図どおりに効いているかを検証する。

実行:
    python test_ma_filter.py
"""

from __future__ import annotations

import sys

import numpy as np
import pandas as pd

from backtest import Backtester
from config import BacktestConfig, StrategyConfig
from strategy import (
    compute_ma_exit,
    compute_ma_filter,
    compute_market_regime,
    compute_moving_averages,
    compute_signals,
    is_market_ok,
)

# Windows stdout を UTF-8 に
if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


# ──────────────────────────────────────────
# テスト用データ生成
# ──────────────────────────────────────────

def make_df(closes: list[float], volumes: list[float] | None = None) -> pd.DataFrame:
    """終値の列から OHLCV の DataFrame を作る（High/Low は終値と同じ）。"""
    n = len(closes)
    dates = pd.bdate_range("2020-01-01", periods=n)
    if volumes is None:
        # 出来高フィルターに引っかからないよう単調増加させる
        volumes = list(np.arange(1, n + 1) * 1000)
    return pd.DataFrame(
        {
            "Open": closes,
            "High": closes,
            "Low": closes,
            "Close": closes,
            "Volume": volumes,
        },
        index=dates,
    )


def small_cfg(**overrides) -> StrategyConfig:
    """短い期間で検証できるようにした設定。"""
    base = dict(
        high_period=5,
        ma_short_period=3,
        ma_mid_period=5,
        ma_long_period=10,
        ma_slope_days=2,
        ma_exit_period=5,
        ma_exit_confirm_days=2,
        market_ma_period=5,
    )
    base.update(overrides)
    return StrategyConfig(**base)


# ──────────────────────────────────────────
# 個別テスト
# ──────────────────────────────────────────

def test_moving_averages_are_simple_means():
    """移動平均が単純平均として正しく計算されている。"""
    closes = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120]
    df = make_df(closes)
    cfg = small_cfg()
    out = compute_moving_averages(df, cfg)

    # 3日移動平均: (10+20+30)/3 = 20
    assert abs(out["ma_short"].iloc[2] - 20.0) < 1e-9, out["ma_short"].iloc[2]
    # 5日移動平均: (10+20+30+40+50)/5 = 30
    assert abs(out["ma_mid"].iloc[4] - 30.0) < 1e-9, out["ma_mid"].iloc[4]
    # 期間に満たない部分は NaN
    assert pd.isna(out["ma_long"].iloc[8])


def test_uptrend_passes_ma_filter():
    """一貫した上昇トレンドはMAフィルターを通過する。"""
    closes = [100 * (1.02 ** i) for i in range(60)]
    df = compute_moving_averages(make_df(closes), small_cfg())
    ok = compute_ma_filter(df, small_cfg())

    # 移動平均が出そろった後半はすべて通過するはず
    assert ok.iloc[-1], "上昇トレンドがMAフィルターを通過していない"
    assert ok.iloc[30:].all(), "上昇トレンド中に不通過の日がある"


def test_downtrend_blocked_by_ma_filter():
    """下降トレンドはMAフィルターで弾かれる。"""
    closes = [100 * (0.98 ** i) for i in range(60)]
    df = compute_moving_averages(make_df(closes), small_cfg())
    ok = compute_ma_filter(df, small_cfg())

    assert not ok.iloc[-1], "下降トレンドがMAフィルターを通過してしまった"
    assert not ok.iloc[20:].any(), "下降トレンド中に通過した日がある"


def test_ma_filter_disabled_passes_everything():
    """ma_filter=False なら常に通過する。"""
    closes = [100 * (0.98 ** i) for i in range(60)]
    cfg = small_cfg(ma_filter=False)
    df = compute_moving_averages(make_df(closes), cfg)
    ok = compute_ma_filter(df, cfg)

    assert ok.all(), "フィルター無効時に不通過の日がある"


def test_alignment_condition_blocks_crossed_mas():
    """
    パーフェクトオーダー条件のみを有効にしたとき、
    MAの並びが崩れている日は弾かれる。
    """
    # 上昇 → 直近で急落（短期MAが中期MAを下抜ける）
    closes = [100 + i for i in range(40)] + [140 - i * 3 for i in range(1, 11)]
    cfg = small_cfg(ma_require_above=False, ma_require_slope_up=False)
    df = compute_moving_averages(make_df(closes), cfg)
    ok = compute_ma_filter(df, cfg)

    assert ok.iloc[38], "上昇局面でパーフェクトオーダーが成立していない"
    assert not ok.iloc[-1], "急落後もパーフェクトオーダー扱いになっている"


def test_slope_condition_blocks_flat_market():
    """長期MAが上向きでない（横ばい）局面は弾かれる。"""
    closes = [100.0] * 60  # 完全な横ばい → 長期MAの傾きはゼロ
    cfg = small_cfg(ma_require_above=False, ma_require_alignment=False)
    df = compute_moving_averages(make_df(closes), cfg)
    ok = compute_ma_filter(df, cfg)

    assert not ok.iloc[-1], "横ばい相場が上向き判定になっている"


def test_ma_exit_needs_consecutive_breaks():
    """MA割れエグジットは指定日数の連続割れで初めて発火する。"""
    # 上昇後に1日だけ大きく下げ、翌日戻す（ダマシ）→ 発火しない
    closes = [100 + i for i in range(30)] + [100.0] + [130.0]
    cfg = small_cfg(ma_exit=True, ma_exit_period=5, ma_exit_confirm_days=2)
    df = compute_moving_averages(make_df(closes), cfg)
    exits = compute_ma_exit(df, cfg)

    assert not exits.iloc[-2], "1日だけの割れでエグジットが発火した"
    assert not exits.iloc[-1], "戻した日にエグジットが発火した"

    # 2日連続で割れたら発火する
    closes2 = [100 + i for i in range(30)] + [100.0, 99.0]
    df2 = compute_moving_averages(make_df(closes2), cfg)
    exits2 = compute_ma_exit(df2, cfg)

    assert exits2.iloc[-1], "2日連続の割れでエグジットが発火しない"


def test_ma_exit_disabled_never_fires():
    """ma_exit=False なら発火しない。"""
    closes = [100 + i for i in range(30)] + [50.0, 40.0, 30.0]
    cfg = small_cfg(ma_exit=False)
    df = compute_moving_averages(make_df(closes), cfg)
    exits = compute_ma_exit(df, cfg)

    assert not exits.any(), "エグジット無効時に発火した"


def test_market_regime_follows_index_ma():
    """指数が長期MAより上のときだけ地合いOKになる。"""
    up = [100 + i for i in range(30)]
    down = [130 - i * 2 for i in range(1, 21)]
    bench = make_df(up + down)
    cfg = small_cfg(market_filter=True, market_ma_period=10)

    regime = compute_market_regime(bench, cfg)
    assert regime is not None
    assert regime.iloc[29], "上昇局面で地合いNGになっている"
    assert not regime.iloc[-1], "下落局面で地合いOKになっている"

    # フィルター無効なら None（＝常に許可）
    cfg_off = small_cfg(market_filter=False)
    assert compute_market_regime(bench, cfg_off) is None
    assert is_market_ok(None, bench.index[-1]), "None のとき常に True にならない"


def test_market_regime_missing_date_uses_previous():
    """指数に該当日が無い場合は直前の判定を引き継ぐ。"""
    bench = make_df([100 + i for i in range(30)])
    cfg = small_cfg(market_ma_period=10)
    regime = compute_market_regime(bench, cfg)

    # データ最終日より後の日付 → 直近の判定（True）を引き継ぐ
    future = bench.index[-1] + pd.Timedelta(days=3)
    assert is_market_ok(regime, future), "将来日付で直近判定が引き継がれていない"

    # データ開始より前の日付 → 判定材料がないので False
    past = bench.index[0] - pd.Timedelta(days=10)
    assert not is_market_ok(regime, past), "判定材料がないのに True になっている"


def test_ma_filter_reduces_signal_count():
    """MAフィルターを入れるとシグナル数は必ず減る（増えることはない）。"""
    rng = np.random.default_rng(42)
    closes = list(100 * np.exp(np.cumsum(rng.normal(0.0003, 0.02, 500))))
    df = make_df(closes, volumes=list(rng.integers(1000, 5000, 500)))

    cfg_off = small_cfg(high_period=20, ma_filter=False)
    cfg_on = small_cfg(high_period=20, ma_filter=True)

    n_off = int(compute_signals(df, cfg_off)["signal"].sum())
    n_on = int(compute_signals(df, cfg_on)["signal"].sum())

    assert n_on <= n_off, f"MAフィルターでシグナルが増えた: {n_off} → {n_on}"
    assert n_off > 0, "検証用データでシグナルが1つも出ていない"


def test_signal_has_no_lookahead():
    """
    シグナルは当日までの情報だけで決まる（先読みしていない）。

    ある日以降のデータを削っても、その日までのシグナルは変わらないことを確認する。
    """
    rng = np.random.default_rng(7)
    closes = list(100 * np.exp(np.cumsum(rng.normal(0.0005, 0.015, 400))))
    df = make_df(closes, volumes=list(rng.integers(1000, 5000, 400)))
    cfg = small_cfg(high_period=20, ma_exit=True)

    full = compute_signals(df, cfg)
    truncated = compute_signals(df.iloc[:300], cfg)

    cols = ["signal", "ma_ok", "ma_exit_signal"]
    for col in cols:
        a = full[col].iloc[:300].astype(bool)
        b = truncated[col].astype(bool)
        assert a.equals(b), f"{col} が将来データの有無で変わっている（先読みの疑い）"


def test_backtest_market_filter_blocks_entries():
    """地合いフィルターONのとき、指数が下落基調なら新規エントリーが減る。"""
    # 個別銘柄は上昇トレンド、指数は下降トレンドという状況を作る
    stock = make_df([100 * (1.01 ** i) for i in range(300)])
    bench = make_df([300 * (0.995 ** i) for i in range(300)])

    bcfg = BacktestConfig(initial_capital=1_000_000)
    cfg_on = small_cfg(high_period=20, market_filter=True, market_ma_period=20)
    cfg_off = small_cfg(high_period=20, market_filter=False)

    r_on = Backtester({"TEST": stock}, cfg_on, bcfg, benchmark=bench).run()
    r_off = Backtester({"TEST": stock}, cfg_off, bcfg, benchmark=bench).run()

    assert len(r_off.trades) > 0, "地合いフィルターOFFで1件も取引していない"
    assert len(r_on.trades) == 0, "指数が下降中なのに新規エントリーしている"


def test_backtest_records_ma_exit_reason():
    """MA割れで手仕舞ったトレードに ma_exit の理由が記録される。"""
    # 上昇 → 明確な下落（損切り幅より先にMAを割る）
    closes = [100 + i * 0.5 for i in range(200)] + [199 - i * 0.4 for i in range(60)]
    stock = make_df(closes)

    cfg = small_cfg(
        high_period=20,
        ma_exit=True,
        ma_exit_period=10,
        ma_exit_confirm_days=2,
        market_filter=False,
        stop_loss_pct=-0.50,
        trailing_stop_pct=-0.50,
    )
    result = Backtester({"TEST": stock}, cfg, BacktestConfig()).run()

    reasons = {t.exit_reason for t in result.closed_trades}
    assert "ma_exit" in reasons, f"MA割れエグジットが記録されていない: {reasons}"


def test_no_exit_before_entry_date():
    """翌日始値エントリーのとき、エントリー日より前に手仕舞われない。"""
    rng = np.random.default_rng(3)
    closes = list(100 * np.exp(np.cumsum(rng.normal(0.001, 0.02, 400))))
    stock = make_df(closes, volumes=list(rng.integers(1000, 5000, 400)))

    cfg = small_cfg(high_period=20, market_filter=False, ma_exit=True)
    result = Backtester({"TEST": stock}, cfg, BacktestConfig()).run()

    for t in result.closed_trades:
        assert t.exit_date >= t.entry_date, (
            f"エントリー日({t.entry_date})より前にエグジット({t.exit_date})している"
        )


# ──────────────────────────────────────────
# 実行
# ──────────────────────────────────────────

def main() -> int:
    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    failures = []

    print(f"\n移動平均線フィルターのテストを実行します（{len(tests)}件）\n")
    for fn in tests:
        try:
            fn()
            print(f"  ✅ {fn.__name__}")
        except AssertionError as e:
            failures.append((fn.__name__, str(e)))
            print(f"  ❌ {fn.__name__}: {e}")
        except Exception as e:  # 想定外のエラーも失敗として扱う
            failures.append((fn.__name__, f"{type(e).__name__}: {e}"))
            print(f"  💥 {fn.__name__}: {type(e).__name__}: {e}")

    print()
    if failures:
        print(f"失敗: {len(failures)}/{len(tests)}")
        for name, msg in failures:
            print(f"  - {name}: {msg}")
        return 1

    print(f"すべて成功: {len(tests)}/{len(tests)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
