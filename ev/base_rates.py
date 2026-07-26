"""
経験的ベースレート — 物語より先に頻度を置く

「この会社は5年で売上2倍になる」と思ったとき、最初に調べるべきは
その会社の事業ではなく、「同じ規模・同じ業種の会社が5年で売上2倍に
なった歴史的頻度」である。

ベースレートを見ずに個別の物語から確率を組み立てると、確率は例外なく
過大になる。人間は説得力のある物語を頻度の代わりに使ってしまう。

このモジュールは履歴パネルからバケット別の実現頻度を計算し、
経験ベイズで全体平均に縮小する。標本が小さいバケットの推定値が
極端になるのを防ぐため。
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, Hashable, Iterable, List, Optional, Sequence, Tuple

from . import stats


class BaseRateError(ValueError):
    pass


@dataclass(frozen=True)
class Observation:
    """履歴パネルの1行。"""
    bucket: Hashable      # 例: ("小型", "建設") や "PBR<1"
    outcome: bool         # 対象の事象が起きたか


@dataclass
class BaseRate:
    bucket: Hashable
    n: int
    successes: int
    raw_rate: float                 # 単純頻度
    shrunk_rate: float              # 全体平均に縮小した推定値
    ci_low: float
    ci_high: float
    global_rate: float
    shrinkage_weight: float         # 自分のバケットをどれだけ信用したか

    @property
    def is_reliable(self) -> bool:
        """信頼区間の幅が0.2未満なら、確率の出発点として使える。"""
        return self.n >= 30 and (self.ci_high - self.ci_low) < 0.20

    def summary(self) -> str:
        flag = "" if self.is_reliable else "  ※標本不足"
        return (
            f"{str(self.bucket):<28} n={self.n:>5}  "
            f"単純={self.raw_rate:.3f}  縮小後={self.shrunk_rate:.3f}  "
            f"95%CI[{self.ci_low:.3f}, {self.ci_high:.3f}]{flag}"
        )


@dataclass
class BaseRateTable:
    rates: Dict[Hashable, BaseRate]
    global_rate: float
    total_n: int
    prior_strength: float

    def get(self, bucket: Hashable) -> BaseRate:
        """
        バケットのベースレートを返す。未知のバケットなら全体平均を返す。
        """
        if bucket in self.rates:
            return self.rates[bucket]
        lo, hi = stats.wilson_interval(
            round(self.global_rate * self.total_n), self.total_n
        )
        return BaseRate(
            bucket=bucket,
            n=0,
            successes=0,
            raw_rate=self.global_rate,
            shrunk_rate=self.global_rate,
            ci_low=lo,
            ci_high=hi,
            global_rate=self.global_rate,
            shrinkage_weight=0.0,
        )

    def summary(self) -> str:
        lines = [
            f"全体ベースレート: {self.global_rate:.4f}  (n={self.total_n})",
            f"事前の強さ      : {self.prior_strength:.1f} 件相当",
            "",
        ]
        for r in sorted(self.rates.values(), key=lambda x: -x.n):
            lines.append(r.summary())
        return "\n".join(lines)


def build_base_rates(
    observations: Iterable[Observation],
    prior_strength: Optional[float] = None,
    confidence: float = 0.95,
) -> BaseRateTable:
    """
    履歴観測からベースレート表を構築する。

    経験ベイズによる縮小:
        縮小後 = (成功数 + k·全体率) / (件数 + k)
    k は事前の強さ（何件ぶんの情報として全体平均を効かせるか）。
    省略時はモーメント法で推定する。

    Args:
        observations: 観測の列
        prior_strength: 事前の強さ k。省略時は自動推定
        confidence: Wilson区間の信頼水準

    Returns:
        BaseRateTable
    """
    obs = list(observations)
    if not obs:
        raise BaseRateError("観測が空です")

    groups: Dict[Hashable, List[bool]] = {}
    for o in obs:
        groups.setdefault(o.bucket, []).append(bool(o.outcome))

    total_n = len(obs)
    total_success = sum(1 for o in obs if o.outcome)
    global_rate = total_success / total_n

    if prior_strength is None:
        prior_strength = _estimate_prior_strength(groups, global_rate)

    rates: Dict[Hashable, BaseRate] = {}
    for bucket, outcomes in groups.items():
        n = len(outcomes)
        s = sum(1 for x in outcomes if x)
        raw = s / n
        shrunk = (s + prior_strength * global_rate) / (n + prior_strength)
        lo, hi = stats.wilson_interval(s, n, confidence)
        rates[bucket] = BaseRate(
            bucket=bucket,
            n=n,
            successes=s,
            raw_rate=raw,
            shrunk_rate=shrunk,
            ci_low=lo,
            ci_high=hi,
            global_rate=global_rate,
            shrinkage_weight=n / (n + prior_strength),
        )

    return BaseRateTable(
        rates=rates,
        global_rate=global_rate,
        total_n=total_n,
        prior_strength=prior_strength,
    )


def _estimate_prior_strength(
    groups: Dict[Hashable, List[bool]],
    global_rate: float,
) -> float:
    """
    モーメント法で Beta 事前分布の強さ k = α+β を推定する。

    バケット間の実現率のばらつきが二項サンプリングノイズで説明できる分だけ
    なら k は大きく（＝強く縮小）、真にばらついているなら k は小さくなる。
    """
    if len(groups) < 2:
        return 10.0

    p = global_rate
    if p <= 0.0 or p >= 1.0:
        return 10.0

    rates = []
    ns = []
    for outcomes in groups.values():
        n = len(outcomes)
        if n < 1:
            continue
        rates.append(sum(1 for x in outcomes if x) / n)
        ns.append(n)
    if len(rates) < 2:
        return 10.0

    observed_var = stats.variance(rates, ddof=1)
    harmonic_n = len(ns) / math.fsum(1.0 / n for n in ns)
    sampling_var = p * (1.0 - p) / harmonic_n

    true_var = observed_var - sampling_var
    if true_var <= 1e-12:
        # ばらつきは全てノイズ。強く縮小する。
        return 200.0
    k = p * (1.0 - p) / true_var - 1.0
    return stats.clamp(k, 1.0, 500.0)


# ────────────────────────────────────────────────────────────────
# 予測確率のベースレート整合チェック
# ────────────────────────────────────────────────────────────────

@dataclass
class BaseRateCheck:
    probability: float
    base_rate: float
    logit_gap: float
    implied_odds_multiple: float
    requires_rationale: bool
    message: str


def check_against_base_rate(
    probability: float,
    base_rate: float,
    divergence_limit: float = 1.0,
) -> BaseRateCheck:
    """
    予測確率がベースレートからどれだけ離れているかを評価する。

    ロジット差はオッズ倍率に等しい。ロジット差1.0 は「平均的な銘柄より
    オッズが約2.7倍良い」という主張になる。そう主張するには理由が必要。

    Args:
        probability: 自分が置いた確率
        base_rate: 参照ベースレート
        divergence_limit: これを超えたら根拠の記述を要求する

    Returns:
        BaseRateCheck
    """
    if not 0.0 < probability < 1.0:
        raise BaseRateError(f"probability={probability} は (0,1) の外")
    if not 0.0 < base_rate < 1.0:
        raise BaseRateError(f"base_rate={base_rate} は (0,1) の外")

    gap = stats.logit(probability) - stats.logit(base_rate)
    odds_multiple = math.exp(gap)
    requires = abs(gap) > divergence_limit

    if requires:
        direction = "強気" if gap > 0 else "弱気"
        msg = (
            f"ベースレート {base_rate:.2f} に対して {probability:.2f} は"
            f"オッズで {odds_multiple:.2f} 倍 {direction}。"
            f"「なぜこの銘柄は平均と違うのか」を書けないなら、"
            f"確率をベースレート寄りに戻すこと。"
        )
    else:
        msg = (
            f"ベースレート {base_rate:.2f} に対しオッズ {odds_multiple:.2f} 倍。"
            f"妥当な範囲。"
        )

    return BaseRateCheck(
        probability=probability,
        base_rate=base_rate,
        logit_gap=gap,
        implied_odds_multiple=odds_multiple,
        requires_rationale=requires,
        message=msg,
    )


def observations_from_rows(
    rows: Iterable[Dict[str, Any]],
    bucket_fn: Callable[[Dict[str, Any]], Hashable],
    outcome_fn: Callable[[Dict[str, Any]], bool],
) -> List[Observation]:
    """
    CSV等の行辞書から Observation を作る補助関数。

    使用例:
        obs = observations_from_rows(
            rows,
            bucket_fn=lambda r: (r["size_bucket"], r["sector"]),
            outcome_fn=lambda r: float(r["fwd_3y_return"]) > 0.0,
        )
    """
    out: List[Observation] = []
    for r in rows:
        try:
            out.append(Observation(bucket=bucket_fn(r), outcome=outcome_fn(r)))
        except (KeyError, TypeError, ValueError):
            continue
    return out
