"""
相関分析の統計モジュール

korea_correlation.py から使う純粋な計算関数群。
外部データに依存しないので test_corrstats.py で数値検証できる。

scipy があれば t分布・正規分布の正確なp値を使い、
無ければ正規近似にフォールバックする（n が数十以上あれば実用上ほぼ同じ）。
"""

from __future__ import annotations

import math
from typing import Optional, Sequence, Tuple

import numpy as np
import pandas as pd

try:
    from scipy import stats as _scipy_stats
    HAS_SCIPY = True
except ImportError:  # scipy は任意依存
    _scipy_stats = None
    HAS_SCIPY = False


# ──────────────────────────────────────────
# p値のヘルパー
# ──────────────────────────────────────────

def _norm_sf(z: float) -> float:
    """標準正規分布の上側確率 P(Z > z)。"""
    return 0.5 * math.erfc(z / math.sqrt(2.0))


def p_from_z(z: float) -> float:
    """両側検定のp値（正規分布）。"""
    if not np.isfinite(z):
        return float("nan")
    return 2.0 * _norm_sf(abs(z))


def p_from_t(t: float, df: int) -> float:
    """両側検定のp値（t分布）。scipy が無ければ正規近似。"""
    if not np.isfinite(t) or df <= 0:
        return float("nan")
    if HAS_SCIPY:
        return float(2.0 * _scipy_stats.t.sf(abs(t), df))
    return p_from_z(t)


# ──────────────────────────────────────────
# リターン系列
# ──────────────────────────────────────────

def log_returns(close: pd.Series) -> pd.Series:
    """終値の対数リターン。0以下の値は欠損として落とす。"""
    s = pd.Series(close).astype(float)
    s = s.where(s > 0)
    return np.log(s).diff()


def align_returns(series_map: dict) -> pd.DataFrame:
    """複数のリターン系列を共通日付で内部結合する。

    日本と韓国は祝日が異なるため、両市場が開いていた日のみを使う。
    """
    if not series_map:
        return pd.DataFrame()
    df = pd.concat(series_map, axis=1, join="inner")
    return df.dropna()


# ──────────────────────────────────────────
# 相関とその検定
# ──────────────────────────────────────────

def corr_with_p(x: Sequence[float], y: Sequence[float]) -> Tuple[float, float, int]:
    """ピアソン相関と両側p値、有効サンプル数を返す。"""
    a = np.asarray(x, dtype=float)
    b = np.asarray(y, dtype=float)
    mask = np.isfinite(a) & np.isfinite(b)
    a, b = a[mask], b[mask]
    n = len(a)
    if n < 3 or a.std() == 0 or b.std() == 0:
        return float("nan"), float("nan"), n

    r = float(np.corrcoef(a, b)[0, 1])
    r_clipped = min(max(r, -0.999999999), 0.999999999)
    t = r_clipped * math.sqrt((n - 2) / (1 - r_clipped**2))
    return r, p_from_t(t, n - 2), n


def fisher_z(r: float) -> float:
    """相関係数のFisher z変換。"""
    r = min(max(float(r), -0.999999999), 0.999999999)
    return math.atanh(r)


def compare_correlations(r1: float, n1: int, r2: float, n2: int) -> Tuple[float, float]:
    """独立な2標本の相関の差を検定する（Fisher z）。

    Returns:
        (z統計量, 両側p値)。z > 0 なら r1 のほうが大きい。

    注意: 2つの標本が期間として重ならないこと。日次リターンには
    ボラティリティ・クラスタリングがあるため、p値はやや楽観的に出る。
    """
    if n1 < 4 or n2 < 4 or not np.isfinite(r1) or not np.isfinite(r2):
        return float("nan"), float("nan")
    se = math.sqrt(1.0 / (n1 - 3) + 1.0 / (n2 - 3))
    z = (fisher_z(r1) - fisher_z(r2)) / se
    return z, p_from_z(z)


def rolling_correlation(x: pd.Series, y: pd.Series, window: int) -> pd.Series:
    """ローリング相関。"""
    return x.rolling(window=window, min_periods=window).corr(y)


def percentile_rank(value: float, history: Sequence[float]) -> float:
    """history の中で value が何パーセンタイルに位置するかを返す（0〜100）。"""
    arr = np.asarray([v for v in history if np.isfinite(v)], dtype=float)
    if len(arr) == 0 or not np.isfinite(value):
        return float("nan")
    return float((arr < value).sum() / len(arr) * 100.0)


# ──────────────────────────────────────────
# 偏相関
# ──────────────────────────────────────────

def _residuals(y: np.ndarray, controls: np.ndarray) -> np.ndarray:
    """定数項付きOLSで y を controls に回帰した残差を返す。"""
    design = np.column_stack([np.ones(len(y)), controls])
    beta, *_ = np.linalg.lstsq(design, y, rcond=None)
    return y - design @ beta


def _is_degenerate(residual: np.ndarray, original: np.ndarray, rtol: float = 1e-8) -> bool:
    """残差が元の系列の分散に対して無視できるほど小さいか。"""
    base = float(np.std(original))
    if base == 0:
        return True
    return float(np.std(residual)) <= rtol * base


def partial_correlation(
    x: Sequence[float],
    y: Sequence[float],
    controls: Optional[pd.DataFrame],
) -> Tuple[float, float, int]:
    """controls の影響を除いた x と y の偏相関、p値、サンプル数を返す。

    x・y をそれぞれ controls に回帰し、残差同士の相関をとる。
    「KOSPIと日本株の相関が、実は米国株という共通要因で説明できるだけ
    ではないか」を判定するために使う。
    """
    a = np.asarray(x, dtype=float)
    b = np.asarray(y, dtype=float)

    if controls is None or controls.shape[1] == 0:
        return corr_with_p(a, b)

    c = np.asarray(controls, dtype=float)
    mask = np.isfinite(a) & np.isfinite(b) & np.isfinite(c).all(axis=1)
    a, b, c = a[mask], b[mask], c[mask]

    n = len(a)
    k = c.shape[1]
    if n < k + 4:
        return float("nan"), float("nan"), n

    ra = _residuals(a, c)
    rb = _residuals(b, c)

    # 統制変数がx（またはy）をほぼ完全に説明してしまう場合、残差は
    # 浮動小数点誤差だけになる。std==0 の判定では 1e-16 レベルの
    # ノイズをすり抜けて無意味な相関が出るため、元の分散との比で見る。
    if _is_degenerate(ra, a) or _is_degenerate(rb, b):
        return float("nan"), float("nan"), n

    r = float(np.corrcoef(ra, rb)[0, 1])
    r_clipped = min(max(r, -0.999999999), 0.999999999)
    df = n - 2 - k
    if df <= 0:
        return r, float("nan"), n
    t = r_clipped * math.sqrt(df / (1 - r_clipped**2))
    return r, p_from_t(t, df), n


def beta(x: Sequence[float], y: Sequence[float]) -> float:
    """y = a + b*x の傾き b（x に対する y の感応度）。"""
    a = np.asarray(x, dtype=float)
    b_ = np.asarray(y, dtype=float)
    mask = np.isfinite(a) & np.isfinite(b_)
    a, b_ = a[mask], b_[mask]
    if len(a) < 3 or a.std() == 0:
        return float("nan")
    return float(np.polyfit(a, b_, 1)[0])


# ──────────────────────────────────────────
# リードラグ
# ──────────────────────────────────────────

def lead_lag_table(driver: pd.Series, target: pd.Series, max_lag: int = 5) -> pd.DataFrame:
    """driver を lag 日ずらしたときの target との相関表を返す。

    lag > 0 : driver が lag 日「先行」する（driver の過去 → target の当日）
    lag = 0 : 同時刻
    lag < 0 : target のほうが先行する

    韓国(KRX)と日本(TSE)はどちらも UTC+9 の 09:00〜15:30 で立会時間が
    ほぼ同じなので、lag=0 の相関は同時刻であって売買には使えない。
    実際にトレードで使えるのは lag>=1（前日までの韓国の値動き）だけ。
    """
    rows = []
    for lag in range(-max_lag, max_lag + 1):
        shifted = driver.shift(lag)
        joined = pd.concat([shifted, target], axis=1, join="inner").dropna()
        if len(joined) < 3:
            rows.append({"lag": lag, "corr": float("nan"), "p_value": float("nan"), "n": len(joined)})
            continue
        r, p, n = corr_with_p(joined.iloc[:, 0], joined.iloc[:, 1])
        rows.append({"lag": lag, "corr": r, "p_value": p, "n": n})
    return pd.DataFrame(rows).set_index("lag")
