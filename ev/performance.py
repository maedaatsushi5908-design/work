"""
成績評価 — 自分のアルファはいつ検証できるのか

結論: 事実上できない。

年3%のアルファ、トラッキングエラー15%なら、t=2 に到達するのに100年かかる。
これは能力の問題ではなく標本数の問題である。競馬なら1シーズンで数千回の
独立試行が得られるが、株の集中投資では一生かけても数百回に届かない。

この事実から2つの帰結がある。
  1. リターンで自分を評価するのは諦める（測れないものは管理できない）
  2. 代わりにキャリブレーションで評価する（遥かに速く測れる）

このモジュールは1を定量的に示し、諦める根拠を与える。
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Tuple

from . import stats

TRADING_DAYS = 252


class PerformanceError(ValueError):
    pass


# ────────────────────────────────────────────────────────────────
# アルファの有意性
# ────────────────────────────────────────────────────────────────

@dataclass
class AlphaReport:
    n_obs: int
    periods_per_year: int
    alpha_annual: float
    tracking_error_annual: float
    information_ratio: float
    t_stat: float
    p_value: float
    years_observed: float
    years_for_t2: float
    years_for_t3: float
    beta: Optional[float] = None
    r_squared: Optional[float] = None

    @property
    def is_significant(self) -> bool:
        return self.p_value < 0.05 and self.alpha_annual > 0

    def summary(self) -> str:
        def yrs(v: float) -> str:
            return "到達しない" if math.isinf(v) else f"{v:,.0f}年"

        lines = [
            f"観測期間          : {self.years_observed:.2f}年 ({self.n_obs}観測)",
            f"年率アルファ      : {self.alpha_annual * 100:+.2f}%",
            f"トラッキングエラー: {self.tracking_error_annual * 100:.2f}%",
            f"情報レシオ        : {self.information_ratio:+.3f}",
            f"t値 / p値         : {self.t_stat:+.3f} / {self.p_value:.4f}",
        ]
        if self.beta is not None:
            lines.append(f"ベータ / R²        : {self.beta:.3f} / {self.r_squared:.3f}")
        lines += [
            "",
            f"t=2 (有意水準5%) に必要な期間: {yrs(self.years_for_t2)}",
            f"t=3 (多重検定を考慮) に必要な期間: {yrs(self.years_for_t3)}",
            "",
        ]
        if math.isinf(self.years_for_t2):
            lines.append(
                "判定: アルファが0以下。何年続けても有意にならない。"
            )
        elif self.years_for_t2 > 40:
            lines.append(
                f"判定: 実力の証明には{self.years_for_t2:,.0f}年必要。"
                f"人生の残り時間では検証不能。リターンで自分を評価するのは諦めて、"
                f"キャリブレーション（ev.calibration）で測ること。"
            )
        elif self.is_significant:
            lines.append("判定: 統計的に有意。ただし多重検定と生存者バイアスに注意。")
        else:
            lines.append(
                f"判定: まだ有意でない。あと{max(0.0, self.years_for_t2 - self.years_observed):,.0f}年必要。"
            )
        return "\n".join(lines)


def years_to_significance(
    alpha_annual: float,
    tracking_error_annual: float,
    t_target: float = 2.0,
) -> float:
    """
    アルファが t_target に到達するのに必要な年数。

        t = (α / TE) · √years   ⟹   years = (t · TE / α)²

    例: α=3%, TE=15%, t=2 → (2 × 15 / 3)² = 100年

    アルファが0以下なら inf を返す。
    """
    if tracking_error_annual < 0:
        raise PerformanceError(f"tracking_error_annual={tracking_error_annual} は0以上")
    if t_target <= 0:
        raise PerformanceError(f"t_target={t_target} は正")
    if alpha_annual <= 0:
        return math.inf
    if tracking_error_annual == 0.0:
        return 0.0
    return (t_target * tracking_error_annual / alpha_annual) ** 2


def evaluate_alpha(
    portfolio_returns: Sequence[float],
    benchmark_returns: Sequence[float],
    periods_per_year: int = TRADING_DAYS,
    use_regression: bool = True,
) -> AlphaReport:
    """
    ベンチマーク対比の成績を評価し、有意性到達に必要な年数を示す。

    Args:
        portfolio_returns: 自分のポートフォリオの期間リターン
        benchmark_returns: ベンチマークの同期間リターン
        periods_per_year: 年率化係数（日次なら252）
        use_regression: True なら回帰でベータ調整したアルファも算出

    Returns:
        AlphaReport
    """
    if len(portfolio_returns) != len(benchmark_returns):
        raise PerformanceError(
            f"長さ不一致: portfolio={len(portfolio_returns)}, benchmark={len(benchmark_returns)}"
        )
    n = len(portfolio_returns)
    if n < 3:
        raise PerformanceError(f"観測数{n}では評価できない（3以上必要）")

    active = [p - b for p, b in zip(portfolio_returns, benchmark_returns)]
    mean_active = stats.mean(active)
    te_period = stats.stdev(active, ddof=1)

    alpha_annual = mean_active * periods_per_year
    te_annual = te_period * math.sqrt(periods_per_year)

    if te_period == 0.0:
        t_stat = math.inf if mean_active > 0 else (-math.inf if mean_active < 0 else 0.0)
        p_value = 0.0 if mean_active != 0 else 1.0
    else:
        t_stat = mean_active / (te_period / math.sqrt(n))
        p_value = stats.t_two_sided_p(t_stat, n - 1)

    beta = r2 = None
    if use_regression:
        beta, alpha_period_reg, r2, resid_sd = _ols(portfolio_returns, benchmark_returns)
        if beta is not None:
            # 回帰ベースのアルファとその残差リスクで置き換える
            alpha_annual = alpha_period_reg * periods_per_year
            te_annual = resid_sd * math.sqrt(periods_per_year)
            if resid_sd > 0:
                se = resid_sd / math.sqrt(n)
                t_stat = alpha_period_reg / se
                p_value = stats.t_two_sided_p(t_stat, max(1, n - 2))

    ir = alpha_annual / te_annual if te_annual > 0 else 0.0

    return AlphaReport(
        n_obs=n,
        periods_per_year=periods_per_year,
        alpha_annual=alpha_annual,
        tracking_error_annual=te_annual,
        information_ratio=ir,
        t_stat=t_stat,
        p_value=p_value,
        years_observed=n / periods_per_year,
        years_for_t2=years_to_significance(alpha_annual, te_annual, 2.0),
        years_for_t3=years_to_significance(alpha_annual, te_annual, 3.0),
        beta=beta,
        r_squared=r2,
    )


def _ols(
    y: Sequence[float], x: Sequence[float]
) -> Tuple[Optional[float], float, Optional[float], float]:
    """
    単回帰 y = a + b·x。

    Returns:
        (傾き b, 切片 a, R², 残差標準偏差)
    """
    n = len(x)
    mx, my = stats.mean(x), stats.mean(y)
    sxx = math.fsum((xi - mx) ** 2 for xi in x)
    if sxx <= 0.0:
        return (None, my, None, stats.stdev(y, ddof=1) if n > 1 else 0.0)
    sxy = math.fsum((xi - mx) * (yi - my) for xi, yi in zip(x, y))
    b = sxy / sxx
    a = my - b * mx
    resid = [yi - (a + b * xi) for xi, yi in zip(x, y)]
    sse = math.fsum(r * r for r in resid)
    syy = math.fsum((yi - my) ** 2 for yi in y)
    r2 = 1.0 - sse / syy if syy > 0 else None
    df = max(1, n - 2)
    return (b, a, r2, math.sqrt(sse / df))


# ────────────────────────────────────────────────────────────────
# ボラティリティ・ドラッグ
# ────────────────────────────────────────────────────────────────

@dataclass
class GrowthReport:
    arithmetic_mean: float
    volatility: float
    geometric_approx: float
    drag: float

    def summary(self) -> str:
        return (
            f"算術平均リターン: {self.arithmetic_mean * 100:+.2f}%\n"
            f"ボラティリティ  : {self.volatility * 100:.2f}%\n"
            f"幾何平均(近似)  : {self.geometric_approx * 100:+.2f}%\n"
            f"ドラッグ        : {self.drag * 100:.2f}%ポイント\n"
            + (
                "→ 算術期待値はプラスだが複利成長率はマイナス。この配分では資産が減る。"
                if self.arithmetic_mean > 0 and self.geometric_approx < 0
                else ""
            )
        )


def volatility_drag(arithmetic_mean: float, volatility: float) -> GrowthReport:
    """
    算術平均と幾何平均の差（ボラティリティ・ドラッグ）。

        幾何平均 ≈ 算術平均 − σ²/2

    分散が大きいほど、同じ算術期待値でも実際の資産成長は小さくなる。
    分散を下げること（分散投資）は算術期待値を下げずに幾何平均を上げる、
    競馬には存在しないフリーランチである。
    """
    if volatility < 0:
        raise PerformanceError(f"volatility={volatility} は0以上")
    drag = volatility * volatility / 2.0
    return GrowthReport(
        arithmetic_mean=arithmetic_mean,
        volatility=volatility,
        geometric_approx=arithmetic_mean - drag,
        drag=drag,
    )


@dataclass
class TwoPointBet:
    win_return: float
    loss_return: float
    win_probability: float
    arithmetic_ev: float
    full_bet_growth: float          # 全額投入時の1回あたり複利成長率
    optimal_fraction: float         # ケリー比率
    optimal_growth: float           # ケリー比率での成長率

    def summary(self) -> str:
        return (
            f"勝ち {self.win_return * 100:+.0f}% / 負け {self.loss_return * 100:+.0f}% "
            f"(勝率 {self.win_probability * 100:.0f}%)\n"
            f"  算術期待値            : {self.arithmetic_ev * 100:+.2f}%\n"
            f"  全額投入時の成長率    : {self.full_bet_growth * 100:+.2f}% / 回\n"
            f"  ケリー比率            : {self.optimal_fraction * 100:.1f}%\n"
            f"  ケリー比率での成長率  : {self.optimal_growth * 100:+.2f}% / 回"
        )


def two_point_bet(
    win_return: float,
    loss_return: float,
    win_probability: float = 0.5,
) -> TwoPointBet:
    """
    2値の賭けについて、算術期待値・全額投入時の成長率・ケリー比率を計算する。

    既定に近い例（+50% / −40% / 勝率50%）では算術期待値は +5% だが、
    全額投入時の成長率は −5.1%。賭け金を25%に落とすと +0.62% になる。

    この符号反転がサイジングの全てである。
    """
    if not 0.0 < win_probability < 1.0:
        raise PerformanceError(f"win_probability={win_probability} は (0,1)")
    if win_return <= -1.0 or loss_return <= -1.0:
        raise PerformanceError("リターンが -100% 以下です")

    p, q = win_probability, 1.0 - win_probability
    ev = p * win_return + q * loss_return
    full = math.exp(p * math.log1p(win_return) + q * math.log1p(loss_return)) - 1.0

    # 遅延インポートで循環参照を避ける
    from .sizing import exact_kelly, log_growth

    f, _ = exact_kelly([win_return, loss_return], [p, q])
    g = log_growth(f, [win_return, loss_return], [p, q])
    opt_growth = math.exp(g) - 1.0 if g > -math.inf else -1.0

    return TwoPointBet(
        win_return=win_return,
        loss_return=loss_return,
        win_probability=p,
        arithmetic_ev=ev,
        full_bet_growth=full,
        optimal_fraction=f,
        optimal_growth=opt_growth,
    )


# ────────────────────────────────────────────────────────────────
# コスト
# ────────────────────────────────────────────────────────────────

def cost_drag_over_time(
    annual_cost: float,
    years: int,
    gross_return: float = 0.06,
) -> Tuple[float, float, float]:
    """
    コストが最終資産に与える影響。

    コスト削減はリターンと違って確実な期待値改善である。年1%の差は
    30年で最終資産の約26%（1 − 0.99^30 ではなく複利の差）に相当する。

    Returns:
        (コストなしの倍率, コストありの倍率, 失われた割合)
    """
    if years < 1:
        raise PerformanceError(f"years={years} は1以上")
    if annual_cost < 0:
        raise PerformanceError(f"annual_cost={annual_cost} は0以上")
    gross = (1.0 + gross_return) ** years
    net = (1.0 + gross_return - annual_cost) ** years
    lost = 1.0 - net / gross if gross > 0 else 0.0
    return (gross, net, lost)
