"""
corrstats.py の数値検証

既知の答えが出るデータを作って、相関・偏相関・Fisher z検定・
リードラグの計算が正しいことを確認する。

    python test_corrstats.py
"""

from __future__ import annotations

import math

import numpy as np
import pandas as pd

import corrstats as cs

PASS = 0
FAIL = 0


def check(name: str, got, want, tol: float = 1e-6) -> None:
    global PASS, FAIL
    if isinstance(want, bool):
        ok = bool(got) == want
    elif want is None:
        ok = got is None
    elif isinstance(want, float) and math.isnan(want):
        ok = math.isnan(got)
    else:
        ok = abs(float(got) - float(want)) <= tol
    if ok:
        PASS += 1
        print(f"  [OK]   {name}: {got}")
    else:
        FAIL += 1
        print(f"  [FAIL] {name}: got={got} want={want}")


def _corr_series(n: int, rho: float, seed: int = 0):
    """相関が理論上 rho になる2系列を作る。"""
    rng = np.random.default_rng(seed)
    z1 = rng.standard_normal(n)
    z2 = rng.standard_normal(n)
    x = z1
    y = rho * z1 + math.sqrt(1 - rho**2) * z2
    return x, y


def test_log_returns() -> None:
    print("\n[log_returns]")
    # 倍々に増える系列 → 対数リターンは常に log(2)
    s = pd.Series([1.0, 2.0, 4.0, 8.0])
    r = cs.log_returns(s)
    check("最初はNaN", math.isnan(r.iloc[0]), True)
    check("log(2)", r.iloc[1], math.log(2))
    check("log(2) 3本目", r.iloc[3], math.log(2))

    # 0以下は欠損扱い（対数が取れないため）
    r2 = cs.log_returns(pd.Series([1.0, 0.0, 4.0]))
    check("0はNaN", math.isnan(r2.iloc[1]), True)


def test_corr_with_p() -> None:
    print("\n[corr_with_p]")
    # 完全な線形関係 → r = 1
    x = np.arange(50, dtype=float)
    r, p, n = cs.corr_with_p(x, 3 * x + 7)
    check("完全相関 r", r, 1.0, tol=1e-9)
    check("n", n, 50)

    # 符号反転 → r = -1
    r, _, _ = cs.corr_with_p(x, -2 * x)
    check("完全逆相関 r", r, -1.0, tol=1e-9)

    # 手計算できる小標本: x=[1,2,3,4], y=[2,4,5,9]
    #   Sxy=11, Sxx=5, Syy=26 → r = 11/sqrt(130) = 0.9647638212
    r, _, _ = cs.corr_with_p([1, 2, 3, 4], [2, 4, 5, 9])
    check("手計算との一致", r, 11 / math.sqrt(130), tol=1e-12)

    # 分散ゼロ → NaN
    r, p, _ = cs.corr_with_p([1, 1, 1, 1], [1, 2, 3, 4])
    check("定数列はNaN", math.isnan(r), True)

    # NaN を含む行は除外される
    r, _, n = cs.corr_with_p([1, 2, 3, 4, np.nan], [2, 4, 5, 9, 100])
    check("NaN除外後のn", n, 4)
    check("NaN除外後のr", r, 11 / math.sqrt(130), tol=1e-12)

    # 大標本で理論値 rho=0.7 を回収できるか
    x, y = _corr_series(20000, 0.7, seed=42)
    r, p, _ = cs.corr_with_p(x, y)
    check("rho=0.7 の回収", r, 0.7, tol=0.02)
    check("有意 (p<0.001)", p < 0.001, True)

    # 無相関ならp値は大きい
    x, y = _corr_series(500, 0.0, seed=7)
    _, p, _ = cs.corr_with_p(x, y)
    check("無相関のp値は0.05超", p > 0.05, True)


def test_fisher_and_compare() -> None:
    print("\n[fisher_z / compare_correlations]")
    check("atanh(0)", cs.fisher_z(0.0), 0.0)
    check("atanh(0.5)", cs.fisher_z(0.5), math.atanh(0.5))
    check("r=1はクリップされ有限", np.isfinite(cs.fisher_z(1.0)), True)

    # 同じ相関同士 → 差はゼロ、p=1
    z, p = cs.compare_correlations(0.5, 100, 0.5, 100)
    check("差なしのz", z, 0.0)
    check("差なしのp", p, 1.0, tol=1e-9)

    # 手計算: r1=0.8,n1=60, r2=0.4,n2=200
    #  z1=atanh(.8)=1.0986123, z2=atanh(.4)=0.4236489
    #  se=sqrt(1/57+1/197)=sqrt(0.0175439+0.0050761)=sqrt(0.0226200)=0.1504
    #  z=(1.0986123-0.4236489)/0.1504=4.4879
    z, p = cs.compare_correlations(0.8, 60, 0.4, 200)
    se = math.sqrt(1 / 57 + 1 / 197)
    check("z統計量", z, (math.atanh(0.8) - math.atanh(0.4)) / se, tol=1e-9)
    check("z ≈ 4.49", z, 4.4879, tol=1e-3)
    check("有意 (p<0.001)", p < 0.001, True)

    # 相関が下がった場合は z が負になる
    z, _ = cs.compare_correlations(0.3, 60, 0.7, 200)
    check("低下時はz<0", z < 0, True)

    # 標本が小さすぎる場合はNaN
    z, p = cs.compare_correlations(0.5, 3, 0.5, 100)
    check("n<4はNaN", math.isnan(z), True)


def test_partial_correlation() -> None:
    print("\n[partial_correlation]")
    rng = np.random.default_rng(123)
    n = 5000

    # 共通因子 c が x と y の相関を全部作っているケース
    # → 生の相関は高いが、c を除いた偏相関はゼロになるべき
    c = rng.standard_normal(n)
    x = c + 0.5 * rng.standard_normal(n)
    y = c + 0.5 * rng.standard_normal(n)

    raw, _, _ = cs.corr_with_p(x, y)
    check("見せかけの生相関は高い", raw > 0.7, True)

    pr, p, _ = cs.partial_correlation(x, y, pd.DataFrame({"c": c}))
    check("共通因子を除くと偏相関≈0", abs(pr) < 0.05, True)
    check("偏相関は有意でない", p > 0.05, True)

    # x と y に c 経由でない固有の結びつきがある場合は偏相関が残る
    shared = rng.standard_normal(n)
    x2 = c + shared + 0.3 * rng.standard_normal(n)
    y2 = c + shared + 0.3 * rng.standard_normal(n)
    pr2, p2, _ = cs.partial_correlation(x2, y2, pd.DataFrame({"c": c}))
    check("固有の連動は残る", pr2 > 0.5, True)
    check("有意 (p<0.001)", p2 < 0.001, True)

    # controls が空なら通常の相関と一致する
    pr3, _, _ = cs.partial_correlation(x, y, None)
    check("controls=Noneは通常相関", pr3, raw, tol=1e-9)

    # 統制変数がターゲットと完全一致 → 残差ゼロでNaN（発散させない）
    pr4, _, _ = cs.partial_correlation(x, y, pd.DataFrame({"x": x}))
    check("完全多重共線でもNaNで返る", math.isnan(pr4), True)


def test_beta() -> None:
    print("\n[beta]")
    x = np.arange(100, dtype=float)
    check("傾き2.5", cs.beta(x, 2.5 * x + 10), 2.5, tol=1e-9)
    check("傾き-1.0", cs.beta(x, -x), -1.0, tol=1e-9)
    check("定数xはNaN", math.isnan(cs.beta([1, 1, 1], [1, 2, 3])), True)


def test_lead_lag() -> None:
    print("\n[lead_lag_table]")
    # target が driver の2日遅れのコピー → lag=2 で相関1になるべき
    rng = np.random.default_rng(5)
    idx = pd.bdate_range("2024-01-01", periods=400)
    driver = pd.Series(rng.standard_normal(400), index=idx)
    target = driver.shift(2)

    tbl = cs.lead_lag_table(driver, target, max_lag=5)
    check("lag=2でr≈1", tbl.loc[2, "corr"], 1.0, tol=1e-9)
    check("lag=0はほぼ無相関", abs(tbl.loc[0, "corr"]) < 0.15, True)
    check("最大相関のlagは2", int(tbl["corr"].idxmax()), 2)
    check("lag範囲は-5..5", list(tbl.index) == list(range(-5, 6)), True)

    # target が driver に先行するケース → 最大は負のlag
    target2 = driver.shift(-3)
    tbl2 = cs.lead_lag_table(driver, target2, max_lag=5)
    check("逆向きはlag=-3", int(tbl2["corr"].idxmax()), -3)


def test_percentile_rank() -> None:
    print("\n[percentile_rank]")
    hist = list(range(100))  # 0..99
    check("50は50パーセンタイル", cs.percentile_rank(50, hist), 50.0)
    check("最大超は100", cs.percentile_rank(1000, hist), 100.0)
    check("最小未満は0", cs.percentile_rank(-1, hist), 0.0)
    check("NaN混在を無視", cs.percentile_rank(50, hist + [float("nan")]), 50.0)
    check("空履歴はNaN", math.isnan(cs.percentile_rank(1, [])), True)


def test_align_returns() -> None:
    print("\n[align_returns]")
    a = pd.Series([1.0, 2.0, 3.0], index=pd.to_datetime(["2024-01-01", "2024-01-02", "2024-01-03"]))
    b = pd.Series([4.0, 5.0], index=pd.to_datetime(["2024-01-02", "2024-01-03"]))
    out = cs.align_returns({"a": a, "b": b})
    check("共通日付のみ2行", len(out), 2)
    check("列名", list(out.columns) == ["a", "b"], True)
    check("落ちた日付は1/1", "2024-01-01" not in out.index.astype(str), True)
    check("空dictは空DF", cs.align_returns({}).empty, True)


def main() -> int:
    print("=" * 60)
    print("  corrstats.py 数値検証")
    print("=" * 60)
    print(f"  scipy: {'あり' if cs.HAS_SCIPY else 'なし（正規近似にフォールバック）'}")

    test_log_returns()
    test_corr_with_p()
    test_fisher_and_compare()
    test_partial_correlation()
    test_beta()
    test_lead_lag()
    test_percentile_rank()
    test_align_returns()

    print("\n" + "=" * 60)
    print(f"  成功 {PASS} / 失敗 {FAIL}")
    print("=" * 60)
    return 1 if FAIL else 0


if __name__ == "__main__":
    raise SystemExit(main())
