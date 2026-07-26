"""
バックテスト検証 — 「良い成績」と「良く見えるだけの成績」を分離する

このモジュールが答える問いは1つだけ。
「その数字は、何通り試した中から選ばれた数字なのか？」

54通りのパラメータを総当たりして最良を採用したなら、その最良値には
偶然の上振れが必ず混入している。さらに悪いことに、銘柄ユニバースを
バックテスト期間の成績で選抜した場合、混入は桁違いに大きくなる。

提供する道具:
  sharpe_ratio                     — シャープレシオ（期間あたり / 年率）
  deflated_sharpe_ratio            — 試行回数で割り引いたシャープの有意性
  min_track_record_length          — 有意性到達に必要な観測数
  probability_of_backtest_overfitting — CSCV による過剰最適化確率(PBO)
  multiple_testing_haircut         — Bonferroni / Holm / BHY によるシャープ減額
  walk_forward_grid                — 「学習期間で最良を選ぶ」方針の正直な期待値
  selection_inflation              — 銘柄選抜バイアスだけで生じる見かけのエッジ
  deflate_for_selection            — 選抜込みで補正したシャープ

参考文献:
  Bailey & López de Prado (2014), "The Deflated Sharpe Ratio"
  Bailey, Borwein, López de Prado & Zhu (2015), "The Probability of Backtest Overfitting"
  Harvey & Liu (2015), "Backtesting" / "...and the Cross-Section of Expected Returns"
"""

from __future__ import annotations

import math
import random
from dataclasses import dataclass, field
from itertools import combinations
from typing import Dict, List, Optional, Sequence, Tuple

from . import stats

TRADING_DAYS = 252


# ────────────────────────────────────────────────────────────────
# シャープレシオ
# ────────────────────────────────────────────────────────────────

def sharpe_ratio(returns: Sequence[float], periods_per_year: Optional[int] = None) -> float:
    """
    シャープレシオ。periods_per_year を渡すと年率化する。

    注意: 本モジュールの他の関数（DSR等）は「期間あたり」の非年率シャープを
    前提とする。年率値を渡すと結果が壊れるため、単位を混ぜないこと。
    """
    if len(returns) < 2:
        return 0.0
    s = stats.stdev(returns, ddof=1)
    if s == 0.0:
        return 0.0
    sr = stats.mean(returns) / s
    if periods_per_year is not None:
        sr *= math.sqrt(periods_per_year)
    return sr


# ────────────────────────────────────────────────────────────────
# Deflated Sharpe Ratio
# ────────────────────────────────────────────────────────────────

@dataclass
class DeflatedSharpeResult:
    sharpe: float                 # 期間あたり実測シャープ
    sharpe_annual: float          # 年率換算（表示用）
    n_trials: int                 # 試した構成の数
    expected_max_sharpe: float    # エッジ0でも試行数だけで到達する期待最大シャープ
    deflated_p: float             # 真のシャープ>0 である確率
    skew: float
    kurtosis: float
    n_obs: int
    passes: bool                  # deflated_p >= 0.95 か

    def summary(self) -> str:
        verdict = "有意（試行数を考慮しても残る）" if self.passes else "有意でない（試行数で説明できる）"
        return (
            f"実測シャープ      : {self.sharpe:.4f}/期間 (年率 {self.sharpe_annual:.2f})\n"
            f"試行数            : {self.n_trials}\n"
            f"偶然の期待最大値  : {self.expected_max_sharpe:.4f}/期間 "
            f"(年率 {self.expected_max_sharpe * math.sqrt(TRADING_DAYS):.2f})\n"
            f"観測数            : {self.n_obs}\n"
            f"歪度 / 尖度       : {self.skew:.3f} / {self.kurtosis:.3f}\n"
            f"Deflated SR (確率): {self.deflated_p:.4f}\n"
            f"判定              : {verdict}"
        )


def deflated_sharpe_ratio(
    returns: Sequence[float],
    n_trials: int,
    trial_sharpes: Optional[Sequence[float]] = None,
    threshold: float = 0.95,
) -> DeflatedSharpeResult:
    """
    Deflated Sharpe Ratio。

    「N通り試して最良だったシャープ」が、エッジ0の世界でも到達しうる水準かを判定する。

    Args:
        returns: 採用した構成の期間別リターン（日次なら日次）
        n_trials: 実際に試した構成の総数。グリッドサーチなら組み合わせ数。
                  ここを正直に申告することが全て。試したのに数えないと結果は無意味。
        trial_sharpes: 全試行の期間あたりシャープ。渡すと帰無仮説下の分散を
                  実測の散らばりから推定するため精度が上がる。省略時は 1/T を使う。
        threshold: 合格とみなす確率。既定0.95。

    Returns:
        DeflatedSharpeResult
    """
    n = len(returns)
    if n < 4:
        raise ValueError(f"deflated_sharpe_ratio: 観測数{n}では評価できない（4以上必要）")
    if n_trials < 1:
        raise ValueError(f"deflated_sharpe_ratio: n_trials={n_trials} は1以上")

    sr = sharpe_ratio(returns)
    sk = stats.skewness(returns)
    ku = stats.kurtosis(returns, excess=False)

    if trial_sharpes is not None and len(trial_sharpes) > 1:
        var_sr = stats.variance(trial_sharpes, ddof=1)
    else:
        # 帰無仮説（iid正規・エッジ0）下でのシャープの漸近分散
        var_sr = 1.0 / n

    sr_star = math.sqrt(var_sr) * stats.expected_max_of_normals(n_trials)

    denom_sq = 1.0 - sk * sr + (ku - 1.0) / 4.0 * sr * sr
    if denom_sq <= 0.0:
        # 極端な高次モーメントで分散推定が破綻するケース。判定不能として最悪側に倒す。
        deflated_p = 0.0
    else:
        z = (sr - sr_star) * math.sqrt(n - 1) / math.sqrt(denom_sq)
        deflated_p = stats.norm_cdf(z)

    return DeflatedSharpeResult(
        sharpe=sr,
        sharpe_annual=sr * math.sqrt(TRADING_DAYS),
        n_trials=n_trials,
        expected_max_sharpe=sr_star,
        deflated_p=deflated_p,
        skew=sk,
        kurtosis=ku,
        n_obs=n,
        passes=deflated_p >= threshold,
    )


def min_track_record_length(
    returns: Sequence[float],
    benchmark_sharpe: float = 0.0,
    confidence: float = 0.95,
) -> float:
    """
    シャープが benchmark_sharpe を有意に超えると言うために必要な最小観測数。

    MinTRL = 1 + (1 - γ₃·SR + (γ₄-1)/4·SR²) · (z_α / (SR - SR*))²

    実測シャープが benchmark 以下なら inf を返す（何年やっても到達しない）。
    """
    if len(returns) < 4:
        raise ValueError("min_track_record_length: 観測数が不足（4以上必要）")
    sr = sharpe_ratio(returns)
    if sr <= benchmark_sharpe:
        return math.inf
    sk = stats.skewness(returns)
    ku = stats.kurtosis(returns, excess=False)
    z = stats.norm_ppf(confidence)
    denom_sq = 1.0 - sk * sr + (ku - 1.0) / 4.0 * sr * sr
    if denom_sq <= 0.0:
        return math.inf
    return 1.0 + denom_sq * (z / (sr - benchmark_sharpe)) ** 2


# ────────────────────────────────────────────────────────────────
# PBO（過剰最適化確率）— CSCV法
# ────────────────────────────────────────────────────────────────

@dataclass
class PBOResult:
    pbo: float                       # 過剰最適化確率
    n_partitions: int
    median_logit: float
    oos_sharpes_of_is_best: List[float]  # 学習期間で最良だった構成の検証期間シャープ
    is_sharpes_of_is_best: List[float]
    n_configs: int

    @property
    def performance_degradation(self) -> float:
        """学習→検証でのシャープ劣化幅（正なら劣化）。"""
        if not self.oos_sharpes_of_is_best:
            return float("nan")
        return stats.mean(self.is_sharpes_of_is_best) - stats.mean(self.oos_sharpes_of_is_best)

    def summary(self) -> str:
        if self.pbo < 0.25:
            verdict = "過剰最適化の兆候は小さい"
        elif self.pbo < 0.50:
            verdict = "過剰最適化の疑いあり"
        else:
            verdict = "過剰最適化。学習期間で選んだ構成は検証期間で中位以下に沈む"
        oos = stats.mean(self.oos_sharpes_of_is_best) if self.oos_sharpes_of_is_best else float("nan")
        return (
            f"構成数            : {self.n_configs}\n"
            f"分割数            : {self.n_partitions}\n"
            f"PBO               : {self.pbo:.3f}\n"
            f"学習期間シャープ  : {stats.mean(self.is_sharpes_of_is_best):.4f}/期間\n"
            f"検証期間シャープ  : {oos:.4f}/期間\n"
            f"劣化幅            : {self.performance_degradation:.4f}/期間\n"
            f"判定              : {verdict}"
        )


def probability_of_backtest_overfitting(
    trial_returns: Dict[str, Sequence[float]],
    n_blocks: int = 10,
    max_partitions: int = 1000,
) -> PBOResult:
    """
    CSCV (Combinatorially Symmetric Cross-Validation) による PBO 推定。

    手順:
      1. 期間を n_blocks 個のブロックに等分する
      2. 半分を学習、残り半分を検証とする全組み合わせを作る
      3. 各分割で「学習期間シャープが最大の構成」を選ぶ
      4. その構成が検証期間で全構成中どの順位に来るかを見る
      5. 中位より下に沈んだ割合が PBO

    PBOが0.5なら、パラメータ選択はコイン投げと同じ。0.5超なら選択が
    害になっている（学習期間で良く見える構成ほど検証期間で悪い）。

    注意: 分割どうしが同じ時系列を共有するため強く相関しており、単一の
    PBO推定値は標準偏差0.2程度でばらつく。0.6と0.4の差に意味を読まないこと。
    エッジ0のときの実現値は0.2〜0.9の範囲に散る（期待値は0.5）。

    Args:
        trial_returns: {構成名: 期間別リターン列}。全構成で長さを揃える。
        n_blocks: ブロック数（偶数）。組み合わせ数は C(n_blocks, n_blocks/2)。
        max_partitions: 評価する分割数の上限（超える場合は決定論的に間引く）

    Returns:
        PBOResult
    """
    if len(trial_returns) < 2:
        raise ValueError("probability_of_backtest_overfitting: 構成が2つ以上必要")
    if n_blocks < 4 or n_blocks % 2 != 0:
        raise ValueError(f"probability_of_backtest_overfitting: n_blocks={n_blocks} は4以上の偶数")

    names = sorted(trial_returns.keys())
    lengths = {len(trial_returns[k]) for k in names}
    if len(lengths) != 1:
        raise ValueError(f"probability_of_backtest_overfitting: リターン列の長さが不一致 {lengths}")
    total = lengths.pop()
    if total < n_blocks * 2:
        raise ValueError(
            f"probability_of_backtest_overfitting: 観測数{total}に対し n_blocks={n_blocks} は過大"
        )

    # ブロック境界（余りは末尾ブロックに寄せる）
    edges = [round(i * total / n_blocks) for i in range(n_blocks + 1)]
    blocks: List[Tuple[int, int]] = [(edges[i], edges[i + 1]) for i in range(n_blocks)]

    half = n_blocks // 2
    all_parts = list(combinations(range(n_blocks), half))
    if len(all_parts) > max_partitions:
        step = len(all_parts) / max_partitions
        all_parts = [all_parts[int(i * step)] for i in range(max_partitions)]

    logits: List[float] = []
    oos_best: List[float] = []
    is_best: List[float] = []

    for train_idx in all_parts:
        train_set = set(train_idx)
        test_idx = [i for i in range(n_blocks) if i not in train_set]

        def slice_for(name: str, idxs) -> List[float]:
            out: List[float] = []
            for b in idxs:
                lo, hi = blocks[b]
                out.extend(trial_returns[name][lo:hi])
            return out

        train_sr = {n: sharpe_ratio(slice_for(n, train_idx)) for n in names}
        test_sr = {n: sharpe_ratio(slice_for(n, test_idx)) for n in names}

        best = max(names, key=lambda n: train_sr[n])
        is_best.append(train_sr[best])
        oos_best.append(test_sr[best])

        # 検証期間における best の相対順位（1が最下位、n が最上位）
        r = stats.ranks([test_sr[n] for n in names])
        rank_of_best = r[names.index(best)]
        omega = rank_of_best / (len(names) + 1.0)  # (0,1) に収める
        logits.append(math.log(omega / (1.0 - omega)))

    pbo = sum(1 for x in logits if x < 0.0) / len(logits)

    return PBOResult(
        pbo=pbo,
        n_partitions=len(logits),
        median_logit=stats.quantile(logits, 0.5),
        oos_sharpes_of_is_best=oos_best,
        is_sharpes_of_is_best=is_best,
        n_configs=len(names),
    )


# ────────────────────────────────────────────────────────────────
# 多重検定補正
# ────────────────────────────────────────────────────────────────

def _bonferroni(pvals: Sequence[float]) -> List[float]:
    n = len(pvals)
    return [min(1.0, p * n) for p in pvals]


def _holm(pvals: Sequence[float]) -> List[float]:
    n = len(pvals)
    order = sorted(range(n), key=lambda i: pvals[i])
    adj = [0.0] * n
    running = 0.0
    for rank, i in enumerate(order):
        running = max(running, (n - rank) * pvals[i])
        adj[i] = min(1.0, running)
    return adj


def _bhy(pvals: Sequence[float]) -> List[float]:
    """
    Benjamini-Yekutieli の step-up 調整 p 値。

    検定間の任意の相関を許容する代償として c(N)=Σ1/i の係数が入る。
    そのため最小順位の p 値に対する乗数は N·c(N) となり、Bonferroni の N
    より大きくなる（N=54 なら 247 対 54）。最良戦略の p 値については
    BY のほうが Bonferroni より厳しい判定を出す。これは仕様である。
    """
    n = len(pvals)
    c = math.fsum(1.0 / i for i in range(1, n + 1))
    order = sorted(range(n), key=lambda i: pvals[i])
    adj = [0.0] * n
    running = 1.0
    for rank in range(n - 1, -1, -1):
        i = order[rank]
        val = pvals[i] * n * c / (rank + 1)
        running = min(running, val)
        adj[i] = min(1.0, running)
    return adj


@dataclass
class HaircutResult:
    original_sharpe_annual: float
    n_tests: int
    haircuts: Dict[str, float]           # 手法 -> 減額率 (0〜1)
    adjusted_sharpe_annual: Dict[str, float]
    adjusted_pvalues: Dict[str, float]

    def summary(self) -> str:
        lines = [
            f"申告シャープ(年率): {self.original_sharpe_annual:.3f}",
            f"検定数            : {self.n_tests}",
            "",
            f"{'手法':<12}{'調整後p値':>12}{'調整後SR':>12}{'減額率':>10}",
        ]
        for m in ("bonferroni", "holm", "bhy"):
            if m in self.haircuts:
                lines.append(
                    f"{m:<12}{self.adjusted_pvalues[m]:>12.4f}"
                    f"{self.adjusted_sharpe_annual[m]:>12.3f}"
                    f"{self.haircuts[m] * 100:>9.1f}%"
                )
        return "\n".join(lines)


def multiple_testing_haircut(
    sharpe_annual: float,
    n_obs: int,
    n_tests: int,
    periods_per_year: int = TRADING_DAYS,
    other_pvalues: Optional[Sequence[float]] = None,
) -> HaircutResult:
    """
    Harvey & Liu 流のシャープ減額。

    「1つの戦略を検定した」ときの p 値を、「N個試した」前提に補正し、
    その調整済み p 値に対応するシャープを逆算する。

    Args:
        sharpe_annual: 申告されている年率シャープ
        n_obs: 観測数（日次なら営業日数）
        n_tests: 試した戦略・構成の総数
        periods_per_year: 年率化係数
        other_pvalues: 他の試行の p 値。渡すと Holm / BHY が実測分布を使う。
                       省略時は最悪ケース（他は全て有意すれすれ）を仮定。

    Returns:
        HaircutResult
    """
    if n_obs < 2:
        raise ValueError(f"multiple_testing_haircut: n_obs={n_obs} が不足")
    if n_tests < 1:
        raise ValueError(f"multiple_testing_haircut: n_tests={n_tests} は1以上")

    sr_per_period = sharpe_annual / math.sqrt(periods_per_year)
    t_stat = sr_per_period * math.sqrt(n_obs)
    p_single = stats.norm_two_sided_p(t_stat)

    if other_pvalues is not None:
        pvals = [p_single] + list(other_pvalues)
    else:
        # 他の試行の p 値が不明な場合、最も甘い（減額が小さくなる）想定を避けるため
        # 一様分布を仮定する。これは Harvey-Liu の想定に近い。
        pvals = [p_single] + [
            (i + 0.5) / n_tests for i in range(max(0, n_tests - 1))
        ]

    methods = {
        "bonferroni": _bonferroni(pvals),
        "holm": _holm(pvals),
        "bhy": _bhy(pvals),
    }

    haircuts: Dict[str, float] = {}
    adj_sr: Dict[str, float] = {}
    adj_p: Dict[str, float] = {}

    for name, adjusted in methods.items():
        p_adj = adjusted[0]
        adj_p[name] = p_adj
        if p_adj >= 1.0:
            t_adj = 0.0
        else:
            t_adj = stats.norm_ppf(1.0 - p_adj / 2.0)
        sr_adj_annual = t_adj / math.sqrt(n_obs) * math.sqrt(periods_per_year)
        adj_sr[name] = sr_adj_annual
        haircuts[name] = (
            1.0 - sr_adj_annual / sharpe_annual if sharpe_annual > 0 else 1.0
        )

    return HaircutResult(
        original_sharpe_annual=sharpe_annual,
        n_tests=n_tests,
        haircuts=haircuts,
        adjusted_sharpe_annual=adj_sr,
        adjusted_pvalues=adj_p,
    )


# ────────────────────────────────────────────────────────────────
# ウォークフォワード
# ────────────────────────────────────────────────────────────────

@dataclass
class WalkForwardResult:
    n_folds: int
    oos_returns: List[float]              # 連結した検証期間リターン
    selected_per_fold: List[str]
    is_sharpe_per_fold: List[float]
    oos_sharpe_per_fold: List[float]

    @property
    def oos_sharpe(self) -> float:
        return sharpe_ratio(self.oos_returns)

    @property
    def oos_sharpe_annual(self) -> float:
        return sharpe_ratio(self.oos_returns, TRADING_DAYS)

    @property
    def parameter_stability(self) -> float:
        """各foldで選ばれた構成が一致した割合。低いと「最良パラメータ」は幻。"""
        if not self.selected_per_fold:
            return float("nan")
        most = max(set(self.selected_per_fold), key=self.selected_per_fold.count)
        return self.selected_per_fold.count(most) / len(self.selected_per_fold)

    def summary(self) -> str:
        return (
            f"fold数              : {self.n_folds}\n"
            f"検証期間シャープ    : {self.oos_sharpe:.4f}/期間 (年率 {self.oos_sharpe_annual:.2f})\n"
            f"学習期間シャープ平均: {stats.mean(self.is_sharpe_per_fold):.4f}/期間\n"
            f"パラメータ安定性    : {self.parameter_stability * 100:.0f}% "
            f"(選択された構成: {', '.join(dict.fromkeys(self.selected_per_fold))})"
        )


def walk_forward_grid(
    trial_returns: Dict[str, Sequence[float]],
    n_folds: int = 5,
    min_train: int = 250,
) -> WalkForwardResult:
    """
    「学習期間で最良の構成を選び、次の期間で使う」という運用の正直な成績。

    グリッドサーチの結果をそのまま採用した場合の期待値は、この関数が返す
    検証期間シャープであって、グリッドの最良値ではない。

    アンカード方式（学習期間は常に先頭から拡大）で先読みを排除する。

    Args:
        trial_returns: {構成名: 期間別リターン列}
        n_folds: 検証fold数
        min_train: 最初の学習期間の最小長

    Returns:
        WalkForwardResult
    """
    if len(trial_returns) < 2:
        raise ValueError("walk_forward_grid: 構成が2つ以上必要")
    names = sorted(trial_returns.keys())
    lengths = {len(trial_returns[k]) for k in names}
    if len(lengths) != 1:
        raise ValueError(f"walk_forward_grid: リターン列の長さが不一致 {lengths}")
    total = lengths.pop()
    if total <= min_train + n_folds:
        raise ValueError(
            f"walk_forward_grid: 観測数{total}では min_train={min_train}, n_folds={n_folds} を満たせない"
        )

    test_len = (total - min_train) // n_folds
    if test_len < 2:
        raise ValueError("walk_forward_grid: 検証期間が短すぎる。n_folds を減らすこと")

    oos: List[float] = []
    picked: List[str] = []
    is_sr: List[float] = []
    oos_sr: List[float] = []

    for f in range(n_folds):
        train_end = min_train + f * test_len
        test_end = train_end + test_len if f < n_folds - 1 else total

        train_sr = {n: sharpe_ratio(trial_returns[n][:train_end]) for n in names}
        best = max(names, key=lambda n: train_sr[n])
        seg = list(trial_returns[best][train_end:test_end])

        picked.append(best)
        is_sr.append(train_sr[best])
        oos_sr.append(sharpe_ratio(seg))
        oos.extend(seg)

    return WalkForwardResult(
        n_folds=n_folds,
        oos_returns=oos,
        selected_per_fold=picked,
        is_sharpe_per_fold=is_sr,
        oos_sharpe_per_fold=oos_sr,
    )


# ────────────────────────────────────────────────────────────────
# 銘柄選抜バイアス
# ────────────────────────────────────────────────────────────────

@dataclass
class SelectionInflationResult:
    n_candidates: int
    n_selected: int
    n_obs: int
    per_obs_vol: float
    expected_inflation_per_obs: float     # 選抜だけで生じる見かけの平均リターン
    expected_inflation_annual: float      # 年率換算（%）
    expected_sharpe_inflation_annual: float
    p95_inflation_annual: float
    n_simulations: int

    def summary(self) -> str:
        return (
            f"候補銘柄数        : {self.n_candidates}\n"
            f"選抜数            : {self.n_selected}\n"
            f"観測数            : {self.n_obs}\n"
            f"真のエッジ        : 0（帰無仮説）\n"
            f"----\n"
            f"選抜だけで生じる見かけの年率リターン: "
            f"{self.expected_inflation_annual * 100:+.2f}% "
            f"(95%点 {self.p95_inflation_annual * 100:+.2f}%)\n"
            f"選抜だけで生じる見かけの年率シャープ: "
            f"{self.expected_sharpe_inflation_annual:+.3f}\n"
            f"→ 実測値がこの水準以下なら、成績は選抜バイアスで全て説明できる"
        )


def selection_inflation(
    n_candidates: int,
    n_selected: int,
    n_obs: int,
    per_obs_vol: float = 0.02,
    periods_per_year: int = TRADING_DAYS,
    n_simulations: int = 2000,
    seed: int = 20240726,
) -> SelectionInflationResult:
    """
    「バックテスト期間の成績が良かった銘柄を残す」だけで生じる見かけのエッジを測る。

    真のエッジが完全に0の世界を作り、n_candidates 銘柄から成績上位 n_selected 銘柄を
    選抜して、その平均成績がどれだけプラスに見えるかをモンテカルロで求める。

    これが本ライブラリで最も重要な関数である。パラメータの過剰最適化より、
    ユニバースを結果で選ぶことのほうが圧倒的に大きなバイアスを生む。
    東証約4000銘柄から116銘柄を「期待値プラス」で抽出した場合、
    真のエッジが0でも極めて良好なバックテストが必ず得られる。

    Args:
        n_candidates: スクリーニング前の候補銘柄数
        n_selected: 選抜後の銘柄数
        n_obs: 各銘柄の観測数（バックテストの営業日数やトレード数）
        per_obs_vol: 1観測あたりのリターン標準偏差
        periods_per_year: 年率化係数
        n_simulations: シミュレーション回数
        seed: 乱数シード（再現性のため固定）

    Returns:
        SelectionInflationResult
    """
    if n_selected < 1 or n_candidates < n_selected:
        raise ValueError(
            f"selection_inflation: n_candidates={n_candidates} >= n_selected={n_selected} >= 1 が必要"
        )
    if n_obs < 2:
        raise ValueError(f"selection_inflation: n_obs={n_obs} が不足")
    if per_obs_vol <= 0:
        raise ValueError(f"selection_inflation: per_obs_vol={per_obs_vol} は正でなければならない")

    rng = random.Random(seed)
    se = per_obs_vol / math.sqrt(n_obs)   # 各銘柄の標本平均の標準誤差
    inflations: List[float] = []

    for _ in range(n_simulations):
        # 真の平均0の銘柄群。各銘柄の「バックテスト上の平均リターン」は N(0, se²)
        sample_means = [rng.gauss(0.0, se) for _ in range(n_candidates)]
        sample_means.sort(reverse=True)
        top = sample_means[:n_selected]
        inflations.append(stats.mean(top))

    exp_inf = stats.mean(inflations)
    p95 = stats.quantile(inflations, 0.95)

    return SelectionInflationResult(
        n_candidates=n_candidates,
        n_selected=n_selected,
        n_obs=n_obs,
        per_obs_vol=per_obs_vol,
        expected_inflation_per_obs=exp_inf,
        expected_inflation_annual=exp_inf * periods_per_year,
        expected_sharpe_inflation_annual=(exp_inf / per_obs_vol) * math.sqrt(periods_per_year),
        p95_inflation_annual=p95 * periods_per_year,
        n_simulations=n_simulations,
        )


def deflate_for_selection(
    observed_sharpe_annual: float,
    n_candidates: int,
    n_selected: int,
    n_obs: int,
    per_obs_vol: float = 0.02,
    periods_per_year: int = TRADING_DAYS,
    n_simulations: int = 2000,
    seed: int = 20240726,
) -> Tuple[float, SelectionInflationResult]:
    """
    実測シャープから銘柄選抜バイアス分を差し引く。

    Returns:
        (補正後の年率シャープ, 選抜バイアスの詳細)
    """
    inf = selection_inflation(
        n_candidates, n_selected, n_obs, per_obs_vol,
        periods_per_year, n_simulations, seed,
    )
    return observed_sharpe_annual - inf.expected_sharpe_inflation_annual, inf
