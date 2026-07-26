"""
数値・統計基盤（標準ライブラリのみ）

外部依存を持たない方針のため、正規分布・Student-t分布・不完全ベータ関数などを
自前で実装している。精度は倍精度で実用上十分（norm_ppf は約1e-15、
incomplete beta は約1e-12）。
"""

from __future__ import annotations

import math
import random
from typing import Callable, Iterable, List, Sequence, Tuple

EULER_MASCHERONI = 0.5772156649015328606


# ────────────────────────────────────────────────────────────────
# 基本統計量
# ────────────────────────────────────────────────────────────────

def mean(xs: Sequence[float]) -> float:
    n = len(xs)
    if n == 0:
        raise ValueError("mean: 空の系列")
    return math.fsum(xs) / n


def variance(xs: Sequence[float], ddof: int = 1) -> float:
    """標本分散。ddof=1 で不偏分散、ddof=0 で母分散。"""
    n = len(xs)
    if n - ddof <= 0:
        raise ValueError(f"variance: 標本数{n}に対し ddof={ddof} は不正")
    m = mean(xs)
    return math.fsum((x - m) ** 2 for x in xs) / (n - ddof)


def stdev(xs: Sequence[float], ddof: int = 1) -> float:
    return math.sqrt(variance(xs, ddof))


def skewness(xs: Sequence[float]) -> float:
    """母集団歪度（3次モーメント / 母標準偏差^3）。分布が定義できない場合は0。"""
    n = len(xs)
    if n < 3:
        return 0.0
    m = mean(xs)
    s = math.sqrt(math.fsum((x - m) ** 2 for x in xs) / n)
    if s == 0.0:
        return 0.0
    return math.fsum(((x - m) / s) ** 3 for x in xs) / n


def kurtosis(xs: Sequence[float], excess: bool = False) -> float:
    """母集団尖度。excess=False なら正規分布で3を返す（非超過尖度）。"""
    n = len(xs)
    if n < 4:
        return 0.0 if excess else 3.0
    m = mean(xs)
    s = math.sqrt(math.fsum((x - m) ** 2 for x in xs) / n)
    if s == 0.0:
        return 0.0 if excess else 3.0
    k = math.fsum(((x - m) / s) ** 4 for x in xs) / n
    return k - 3.0 if excess else k


def quantile(xs: Sequence[float], q: float) -> float:
    """線形補間による分位点（numpy の既定と同じ挙動）。"""
    if not xs:
        raise ValueError("quantile: 空の系列")
    if not 0.0 <= q <= 1.0:
        raise ValueError(f"quantile: q={q} は [0,1] の外")
    ys = sorted(xs)
    if len(ys) == 1:
        return ys[0]
    pos = q * (len(ys) - 1)
    lo = int(math.floor(pos))
    hi = min(lo + 1, len(ys) - 1)
    frac = pos - lo
    return ys[lo] * (1.0 - frac) + ys[hi] * frac


# ────────────────────────────────────────────────────────────────
# 正規分布
# ────────────────────────────────────────────────────────────────

def norm_pdf(x: float) -> float:
    return math.exp(-0.5 * x * x) / math.sqrt(2.0 * math.pi)


def norm_cdf(x: float) -> float:
    return 0.5 * math.erfc(-x / math.sqrt(2.0))


_ACKLAM_A = (
    -3.969683028665376e01, 2.209460984245205e02, -2.759285104469687e02,
    1.383577518672690e02, -3.066479806614716e01, 2.506628277459239e00,
)
_ACKLAM_B = (
    -5.447609879822406e01, 1.615858368580409e02, -1.556989798598866e02,
    6.680131188771972e01, -1.328068155288572e01,
)
_ACKLAM_C = (
    -7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e00,
    -2.549732539343734e00, 4.374664141464968e00, 2.938163982698783e00,
)
_ACKLAM_D = (
    7.784695709041462e-03, 3.224671290700398e-01,
    2.445134137142996e00, 3.754408661907416e00,
)


def norm_ppf(p: float) -> float:
    """標準正規分布の逆累積分布関数（Acklam近似 + Halley法で1段refine）。"""
    if not 0.0 < p < 1.0:
        if p == 0.0:
            return -math.inf
        if p == 1.0:
            return math.inf
        raise ValueError(f"norm_ppf: p={p} は (0,1) の外")

    a, b, c, d = _ACKLAM_A, _ACKLAM_B, _ACKLAM_C, _ACKLAM_D
    p_low, p_high = 0.02425, 1.0 - 0.02425

    if p < p_low:
        q = math.sqrt(-2.0 * math.log(p))
        x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / \
            ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1.0)
    elif p <= p_high:
        q = p - 0.5
        r = q * q
        x = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q / \
            (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1.0)
    else:
        q = math.sqrt(-2.0 * math.log(1.0 - p))
        x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / \
            ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1.0)

    # Halley 法による精度改善
    e = norm_cdf(x) - p
    u = e * math.sqrt(2.0 * math.pi) * math.exp(x * x / 2.0)
    return x - u / (1.0 + x * u / 2.0)


# ────────────────────────────────────────────────────────────────
# 不完全ベータ関数と Student-t 分布
# ────────────────────────────────────────────────────────────────

def _betacf(a: float, b: float, x: float, max_iter: int = 300, eps: float = 3e-16) -> float:
    """連分数展開（modified Lentz法）。Numerical Recipes の betacf に対応。"""
    tiny = 1e-300
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c = 1.0
    d = 1.0 - qab * x / qap
    if abs(d) < tiny:
        d = tiny
    d = 1.0 / d
    h = d

    for m in range(1, max_iter + 1):
        m2 = 2 * m
        # 偶数項
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        if abs(d) < tiny:
            d = tiny
        c = 1.0 + aa / c
        if abs(c) < tiny:
            c = tiny
        d = 1.0 / d
        h *= d * c
        # 奇数項
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        if abs(d) < tiny:
            d = tiny
        c = 1.0 + aa / c
        if abs(c) < tiny:
            c = tiny
        d = 1.0 / d
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < eps:
            break
    return h


def betainc(a: float, b: float, x: float) -> float:
    """正則化不完全ベータ関数 I_x(a, b)。"""
    if a <= 0.0 or b <= 0.0:
        raise ValueError(f"betainc: a={a}, b={b} は正でなければならない")
    if x <= 0.0:
        return 0.0
    if x >= 1.0:
        return 1.0
    log_bt = (
        math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b)
        + a * math.log(x) + b * math.log1p(-x)
    )
    bt = math.exp(log_bt)
    if x < (a + 1.0) / (a + b + 2.0):
        return bt * _betacf(a, b, x) / a
    return 1.0 - bt * _betacf(b, a, 1.0 - x) / b


def t_sf(t: float, df: float) -> float:
    """Student-t 分布の上側確率 P(T > t)。"""
    if df <= 0:
        raise ValueError(f"t_sf: df={df} は正でなければならない")
    if math.isinf(t):
        return 0.0 if t > 0 else 1.0
    half = 0.5 * betainc(0.5 * df, 0.5, df / (df + t * t))
    return half if t > 0 else 1.0 - half


def t_two_sided_p(t: float, df: float) -> float:
    """Student-t 両側 p 値。"""
    if df <= 0:
        raise ValueError(f"t_two_sided_p: df={df} は正でなければならない")
    if math.isinf(t):
        return 0.0
    return betainc(0.5 * df, 0.5, df / (df + t * t))


def norm_two_sided_p(z: float) -> float:
    """正規分布の両側 p 値。"""
    return math.erfc(abs(z) / math.sqrt(2.0))


# ────────────────────────────────────────────────────────────────
# 二項比率の信頼区間
# ────────────────────────────────────────────────────────────────

def wilson_interval(successes: int, n: int, confidence: float = 0.95) -> Tuple[float, float]:
    """Wilson スコア区間。n が小さいときも破綻しないため、ベースレート推定に適する。"""
    if n <= 0:
        return (0.0, 1.0)
    if successes < 0 or successes > n:
        raise ValueError(f"wilson_interval: successes={successes} が n={n} と矛盾")
    z = norm_ppf(1.0 - (1.0 - confidence) / 2.0)
    p = successes / n
    denom = 1.0 + z * z / n
    center = (p + z * z / (2.0 * n)) / denom
    margin = z * math.sqrt(p * (1.0 - p) / n + z * z / (4.0 * n * n)) / denom
    return (max(0.0, center - margin), min(1.0, center + margin))


# ────────────────────────────────────────────────────────────────
# ブートストラップ
# ────────────────────────────────────────────────────────────────

def bootstrap_ci(
    xs: Sequence[float],
    statistic: Callable[[Sequence[float]], float],
    confidence: float = 0.95,
    n_resamples: int = 2000,
    seed: int = 12345,
) -> Tuple[float, float, float]:
    """
    パーセンタイル法によるブートストラップ信頼区間。

    Returns:
        (点推定, 下限, 上限)
    """
    if not xs:
        raise ValueError("bootstrap_ci: 空の系列")
    point = statistic(xs)
    if len(xs) == 1:
        return (point, point, point)

    rng = random.Random(seed)
    n = len(xs)
    stats: List[float] = []
    for _ in range(n_resamples):
        sample = [xs[rng.randrange(n)] for _ in range(n)]
        try:
            stats.append(statistic(sample))
        except (ValueError, ZeroDivisionError):
            continue
    if not stats:
        return (point, float("nan"), float("nan"))
    alpha = (1.0 - confidence) / 2.0
    return (point, quantile(stats, alpha), quantile(stats, 1.0 - alpha))


# ────────────────────────────────────────────────────────────────
# 極値の期待値（過剰最適化の評価に使用）
# ────────────────────────────────────────────────────────────────

def expected_max_of_normals(n_trials: int) -> float:
    """
    標準正規分布から独立に n_trials 回引いたときの最大値の期待値の近似。

    E[max] ≈ (1-γ)·Φ⁻¹(1 - 1/N) + γ·Φ⁻¹(1 - 1/(N·e))

    Bailey & López de Prado (2014) が Deflated Sharpe Ratio で用いる近似。
    「54通り試して一番良かった」という結果を、偶然の期待値と比較するために使う。
    """
    if n_trials < 1:
        raise ValueError(f"expected_max_of_normals: n_trials={n_trials} は1以上")
    if n_trials == 1:
        return 0.0
    g = EULER_MASCHERONI
    return (
        (1.0 - g) * norm_ppf(1.0 - 1.0 / n_trials)
        + g * norm_ppf(1.0 - 1.0 / (n_trials * math.e))
    )


# ────────────────────────────────────────────────────────────────
# 補助
# ────────────────────────────────────────────────────────────────

def clamp(x: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, x))


def logit(p: float, eps: float = 1e-12) -> float:
    p = clamp(p, eps, 1.0 - eps)
    return math.log(p / (1.0 - p))


def expit(x: float) -> float:
    if x >= 0:
        z = math.exp(-x)
        return 1.0 / (1.0 + z)
    z = math.exp(x)
    return z / (1.0 + z)


def zscores(xs: Sequence[float]) -> List[float]:
    """z化。全要素が同値なら全て0を返す。"""
    if not xs:
        return []
    m = mean(xs)
    s = stdev(xs, ddof=1) if len(xs) > 1 else 0.0
    if s == 0.0:
        return [0.0] * len(xs)
    return [(x - m) / s for x in xs]


def winsorize(xs: Sequence[float], limit: float = 0.02) -> List[float]:
    """裾を分位点でクリップする。ファクター値の外れ値対策。"""
    if not xs or limit <= 0.0:
        return list(xs)
    lo, hi = quantile(xs, limit), quantile(xs, 1.0 - limit)
    return [clamp(x, lo, hi) for x in xs]


def ranks(xs: Sequence[float]) -> List[float]:
    """昇順の平均順位（同値はタイ平均）。1始まり。"""
    n = len(xs)
    if n == 0:
        return []
    order = sorted(range(n), key=lambda i: xs[i])
    out = [0.0] * n
    i = 0
    while i < n:
        j = i
        while j + 1 < n and xs[order[j + 1]] == xs[order[i]]:
            j += 1
        avg = (i + j) / 2.0 + 1.0
        for k in range(i, j + 1):
            out[order[k]] = avg
        i = j + 1
    return out
