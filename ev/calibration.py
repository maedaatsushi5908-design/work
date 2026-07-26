"""
キャリブレーション評価 — 自分の確率が当たっているかを測る

株式投資におけるアルファは、事実上一生かかっても統計的に検証できない。
年3%のアルファ、トラッキングエラー15%なら t=2 に到達するのに100年かかる。

一方でキャリブレーション（確率の当て具合）は、年20〜30個の予測でも傾向が見える。
「70%と言った事象の実現率が70%に近いか」は、リターンより桁違いに速く測れる。

測るもの:
  Brierスコア     — 確率予測の二乗誤差。低いほど良い。厳密に proper なスコア
  Murphy分解      — Brier = 信頼度 − 解像度 + 不確実性
                     信頼度: 確率のズレ（低いほど良い）
                     解像度: 事象を区別できているか（高いほど良い）
                     不確実性: 事象そのものの難しさ（自分では変えられない）
  Brier Skill Score — 「常にベースレートと言う」戦略に対する改善度。0以下なら無価値
  較正傾き        — ロジット上の回帰係数。1未満は自信過剰、1超は自信不足

注意: 解像度が0でも信頼度は完璧になりうる。「常に基準率を答える」は
完全に較正されているが役に立たない。だから BSS で両方を同時に見る。
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Tuple

from . import stats

# これ未満のサンプル数ではキャリブレーションを判断材料にしない
MIN_SAMPLE_FOR_JUDGEMENT = 20


class CalibrationError(ValueError):
    pass


@dataclass
class ReliabilityBin:
    lower: float
    upper: float
    n: int
    mean_forecast: float
    observed_rate: float

    @property
    def gap(self) -> float:
        """予測確率 − 実現率。正なら自信過剰の方向。"""
        return self.mean_forecast - self.observed_rate


@dataclass
class CalibrationReport:
    n: int
    base_rate: float                    # 実現率（ō）
    brier: float
    brier_binned: float                 # Murphy分解の再構成値
    reliability: float                  # 低いほど良い
    resolution: float                   # 高いほど良い
    uncertainty: float
    brier_skill_score: float            # >0 でベースレート予測に勝っている
    expected_calibration_error: float
    log_score: float                    # 平均負の対数尤度。低いほど良い
    sharpness: float                    # 予測確率の分散。0なら常に同じ確率
    calibration_slope: float
    calibration_intercept: float
    slope_converged: bool
    bins: List[ReliabilityBin] = field(default_factory=list)
    bss_ci: Optional[Tuple[float, float]] = None

    # ── 判定 ────────────────────────────────────────────

    @property
    def has_enough_data(self) -> bool:
        return self.n >= MIN_SAMPLE_FOR_JUDGEMENT

    @property
    def is_overconfident(self) -> bool:
        """較正傾きが1を明確に下回る = 極端な確率を出しすぎている。"""
        return self.slope_converged and self.calibration_slope < 0.8

    @property
    def demonstrates_skill(self) -> bool:
        """
        サイジングを許可する条件。

        3つ全てを要求する:
          1. 判断に足るサンプル数
          2. BSS > 0（ベースレート予測に勝っている）
          3. BSS の信頼区間下限 > 0（偶然では説明できない）
        """
        if not self.has_enough_data:
            return False
        if self.brier_skill_score <= 0.0:
            return False
        if self.bss_ci is not None and self.bss_ci[0] <= 0.0:
            return False
        return True

    def verdict(self) -> str:
        if not self.has_enough_data:
            return (
                f"サンプル不足（{self.n}件 / 必要{MIN_SAMPLE_FOR_JUDGEMENT}件）。"
                f"判断材料にはならない。予測を書き続けること。"
            )
        if self.brier_skill_score <= 0.0:
            return (
                "ベースレート予測に負けている。個別の見立てが価値を生んでいない。"
                "この状態で個別銘柄に資金を傾けるのは、期待値がマイナスの賭け。"
            )
        if self.bss_ci is not None and self.bss_ci[0] <= 0.0:
            return (
                f"BSS={self.brier_skill_score:.3f} はプラスだが、信頼区間が0を含む"
                f"（下限 {self.bss_ci[0]:.3f}）。まだ運と区別できない。"
            )
        if self.is_overconfident:
            return (
                f"スキルは検出されたが自信過剰（較正傾き {self.calibration_slope:.2f}）。"
                f"確率を中央に寄せるだけで成績が改善する。"
            )
        return (
            f"スキルを検出（BSS={self.brier_skill_score:.3f}、"
            f"較正傾き {self.calibration_slope:.2f}）。"
        )

    def summary(self) -> str:
        lines = [
            f"予測数            : {self.n}",
            f"実現率(ベースレート): {self.base_rate:.3f}",
            "",
            f"Brierスコア       : {self.brier:.4f}  (低いほど良い)",
            f"  信頼度 (REL)    : {self.reliability:.4f}  (低いほど良い)",
            f"  解像度 (RES)    : {self.resolution:.4f}  (高いほど良い)",
            f"  不確実性 (UNC)  : {self.uncertainty:.4f}  (事象固有の難しさ)",
            f"Brier Skill Score : {self.brier_skill_score:+.4f}"
            + (
                f"  95%CI [{self.bss_ci[0]:+.3f}, {self.bss_ci[1]:+.3f}]"
                if self.bss_ci
                else ""
            ),
            f"平均較正誤差(ECE) : {self.expected_calibration_error:.4f}",
            f"対数スコア        : {self.log_score:.4f}",
            f"鋭さ(予測の分散)  : {self.sharpness:.4f}",
            f"較正傾き / 切片   : {self.calibration_slope:.3f} / {self.calibration_intercept:+.3f}"
            + ("" if self.slope_converged else "  (収束せず)"),
            "",
            "信頼性曲線:",
            f"  {'区間':<14}{'件数':>6}{'予測':>8}{'実現':>8}{'差':>8}",
        ]
        for b in self.bins:
            lines.append(
                f"  {f'{b.lower:.2f}-{b.upper:.2f}':<14}{b.n:>6}"
                f"{b.mean_forecast:>8.3f}{b.observed_rate:>8.3f}{b.gap:>+8.3f}"
            )
        lines += ["", f"判定: {self.verdict()}"]
        return "\n".join(lines)


def _brier(pairs: Sequence[Tuple[float, bool]]) -> float:
    return math.fsum((p - (1.0 if o else 0.0)) ** 2 for p, o in pairs) / len(pairs)


def _brier_skill_score(pairs: Sequence[Tuple[float, bool]]) -> float:
    """ブートストラップ用の軽量版 BSS。"""
    n = len(pairs)
    if n == 0:
        raise ValueError("空の系列")
    obar = sum(1 for _, o in pairs if o) / n
    unc = obar * (1.0 - obar)
    if unc <= 0.0:
        raise ValueError("不確実性が0（全て同一の結果）")
    return 1.0 - _brier(pairs) / unc


def _fit_calibration_slope(
    probabilities: Sequence[float],
    outcomes: Sequence[bool],
    max_iter: int = 100,
    tol: float = 1e-10,
) -> Tuple[float, float, bool]:
    """
    outcome ~ logistic(b0 + b1 · logit(p)) を Newton-Raphson で当てはめる。

    Returns:
        (傾き b1, 切片 b0, 収束したか)
    """
    xs = [stats.logit(p) for p in probabilities]
    ys = [1.0 if o else 0.0 for o in outcomes]
    n = len(xs)
    if n < 3 or len(set(ys)) < 2:
        # 結果が全て同じ、または標本が小さすぎる場合は当てはめ不能
        return (float("nan"), float("nan"), False)

    b0, b1 = 0.0, 1.0
    ridge = 1e-8  # 特異行列の回避

    for _ in range(max_iter):
        g0 = g1 = 0.0
        h00 = h01 = h11 = 0.0
        for x, y in zip(xs, ys):
            mu = stats.expit(b0 + b1 * x)
            r = y - mu
            w = max(mu * (1.0 - mu), 1e-12)
            g0 += r
            g1 += r * x
            h00 += w
            h01 += w * x
            h11 += w * x * x

        h00 += ridge
        h11 += ridge
        det = h00 * h11 - h01 * h01
        if abs(det) < 1e-300:
            return (b1, b0, False)

        # (X'WX)^{-1} g
        d0 = (h11 * g0 - h01 * g1) / det
        d1 = (h00 * g1 - h01 * g0) / det

        # 発散防止のためステップを制限
        step = max(abs(d0), abs(d1))
        if step > 4.0:
            scale = 4.0 / step
            d0 *= scale
            d1 *= scale

        b0 += d0
        b1 += d1

        if max(abs(d0), abs(d1)) < tol:
            return (b1, b0, True)

    return (b1, b0, False)


def evaluate(
    probabilities: Sequence[float],
    outcomes: Sequence[bool],
    n_bins: int = 5,
    bootstrap: bool = True,
    n_resamples: int = 2000,
    seed: int = 4242,
) -> CalibrationReport:
    """
    確率予測の集合を評価する。

    Args:
        probabilities: 事前に記録した確率（0 < p < 1）
        outcomes: 実際の結果（True/False）
        n_bins: 信頼性曲線の区間数
        bootstrap: BSS の信頼区間を計算するか
        n_resamples: ブートストラップ回数
        seed: 乱数シード

    Returns:
        CalibrationReport
    """
    if len(probabilities) != len(outcomes):
        raise CalibrationError(
            f"長さ不一致: probabilities={len(probabilities)}, outcomes={len(outcomes)}"
        )
    n = len(probabilities)
    if n == 0:
        raise CalibrationError("評価対象の予測がありません")
    for p in probabilities:
        if not 0.0 < p < 1.0:
            raise CalibrationError(f"確率 {p} が (0,1) の外にあります")
    if n_bins < 2:
        raise CalibrationError(f"n_bins={n_bins} は2以上")

    pairs: List[Tuple[float, bool]] = list(zip(probabilities, [bool(o) for o in outcomes]))
    ys = [1.0 if o else 0.0 for _, o in pairs]

    brier = _brier(pairs)
    obar = math.fsum(ys) / n
    uncertainty = obar * (1.0 - obar)

    # ── Murphy 分解（ビン化した予測確率に基づく） ──
    edges = [i / n_bins for i in range(n_bins + 1)]
    bins: List[ReliabilityBin] = []
    reliability = 0.0
    resolution = 0.0
    ece = 0.0
    brier_binned = 0.0

    for k in range(n_bins):
        lo, hi = edges[k], edges[k + 1]
        if k == n_bins - 1:
            members = [(p, o) for p, o in pairs if lo <= p <= hi]
        else:
            members = [(p, o) for p, o in pairs if lo <= p < hi]
        if not members:
            continue
        nk = len(members)
        pk = math.fsum(p for p, _ in members) / nk
        ok = sum(1 for _, o in members if o) / nk

        reliability += nk * (pk - ok) ** 2
        resolution += nk * (ok - obar) ** 2
        ece += nk * abs(pk - ok)
        brier_binned += nk * ((pk - ok) ** 2 + ok * (1.0 - ok))

        bins.append(ReliabilityBin(lo, hi, nk, pk, ok))

    reliability /= n
    resolution /= n
    ece /= n
    brier_binned /= n

    bss = 1.0 - brier / uncertainty if uncertainty > 0.0 else float("nan")

    # 対数スコア（低いほど良い）
    log_score = -math.fsum(
        (math.log(p) if o else math.log1p(-p)) for p, o in pairs
    ) / n

    sharpness = stats.variance(list(probabilities), ddof=0)

    slope, intercept, converged = _fit_calibration_slope(
        probabilities, [o for _, o in pairs]
    )

    bss_ci: Optional[Tuple[float, float]] = None
    if bootstrap and n >= 5 and uncertainty > 0.0:
        try:
            _, lo_ci, hi_ci = stats.bootstrap_ci(
                pairs, _brier_skill_score,
                confidence=0.95, n_resamples=n_resamples, seed=seed,
            )
            if not (math.isnan(lo_ci) or math.isnan(hi_ci)):
                bss_ci = (lo_ci, hi_ci)
        except ValueError:
            bss_ci = None

    return CalibrationReport(
        n=n,
        base_rate=obar,
        brier=brier,
        brier_binned=brier_binned,
        reliability=reliability,
        resolution=resolution,
        uncertainty=uncertainty,
        brier_skill_score=bss,
        expected_calibration_error=ece,
        log_score=log_score,
        sharpness=sharpness,
        calibration_slope=slope,
        calibration_intercept=intercept,
        slope_converged=converged,
        bins=bins,
        bss_ci=bss_ci,
    )


def recalibrate(probability: float, report: CalibrationReport) -> float:
    """
    実測の較正傾き・切片を使って将来の確率を補正する。

    自信過剰（傾き<1）が検出されているなら、生の確率は中央に寄せるべきである。
    これは無料の期待値改善になる。より良い予測をする必要はなく、
    既存の予測の目盛りを直すだけでよい。

    Args:
        probability: 生の予測確率
        report: 過去実績から得た CalibrationReport

    Returns:
        補正後の確率
    """
    if not 0.0 < probability < 1.0:
        raise CalibrationError(f"確率 {probability} が (0,1) の外")
    if not report.slope_converged or math.isnan(report.calibration_slope):
        return probability
    if not report.has_enough_data:
        return probability
    z = report.calibration_intercept + report.calibration_slope * stats.logit(probability)
    return stats.clamp(stats.expit(z), 1e-6, 1.0 - 1e-6)
