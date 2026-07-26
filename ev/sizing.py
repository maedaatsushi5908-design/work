"""
ポジションサイジング — 期待値がプラスでも、賭け金を間違えれば資産は減る

算術期待値がプラスなのに破産する例:
  勝てば +50%、負ければ -40%、確率は50/50
    算術期待値 = +5%     → 一見おいしい
    毎回全額投入時の複利成長率 = √(1.5 × 0.6) − 1 = −5.1%  → 100回で無一文
    毎回25%投入時         = +0.62% / 回                    → 増える

同じ賭けでも賭け金だけで符号が反転する。この事実がサイジングの全て。

このモジュールの処理順:
  1. シナリオ分布から厳密なケリー比率を数値解で求める（正規近似ではない）
  2. エッジ推定値をベイズ縮小する（推定誤差が大きいほど0に寄せる）
  3. モデルリスクに対する分数ケリーを適用する
  4. 実測キャリブレーションでゲートする ← ここが要点
  5. 流動性・集中度・最悪シナリオのキャップを適用する
  6. どの制約が効いたかを明示して返す

4が本ライブラリの中核である。「自分にエッジがある」という主張は、
実測されたキャリブレーションによって裏付けられるまで、サイジングに
反映されない。スキルが証明されていない状態でのサテライト配分は0になる。
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from . import stats
from .calibration import CalibrationReport
from .forecast import Forecast, Scenario

# 分数ケリーの既定値。1.0（フルケリー）は理論上の成長率最大だが
# 50%ドローダウンの発生確率が50%になるため実務では使えない。
DEFAULT_KELLY_FRACTION = 0.25

# BSS がこの水準なら「一人前」とみなしゲートを全開にする
BSS_REFERENCE = 0.15

# 事前分布は情報レシオで与える。エッジの事前標準偏差 τ = この値 × σ。
# 絶対値（「年率3%」等）で与えると保有期間や銘柄のボラティリティと
# 単位が噛み合わなくなるため、必ず σ に比例させる。
# 0.25 は個人投資家の実現IRとしてはかなり寛大な想定。
DEFAULT_PRIOR_INFORMATION_RATIO = 0.25

# 中心推定値の誤差 s の既定値 = この値 × σ。
# 重要: σ は「結果のばらつき」であり、s は「自分の中心推定値がどれだけ
# 外れうるか」であって別物。原理的な既定値は存在しないので、保守側に置く。
# この2つの既定値の組み合わせは縮小の重みを
#   w = 0.25² / (0.25² + 0.5²) = 0.2
# に固定する。つまり既定では主張エッジを2割に値引きして扱う。
DEFAULT_ESTIMATE_ERROR_RATIO = 0.5

# ケリー比率の数値解の上限（これを超える解はモデル誤りとして扱う）
MAX_KELLY_SEARCH = 5.0


class SizingError(ValueError):
    pass


# ────────────────────────────────────────────────────────────────
# ケリー比率
# ────────────────────────────────────────────────────────────────

def log_growth(weight: float, returns: Sequence[float], probs: Sequence[float]) -> float:
    """
    ウェイト weight を賭けたときの1期間あたり対数期待成長率 Σ p·ln(1 + w·r)。

    破産（1 + w·r <= 0）を含む場合は -inf。
    """
    total = 0.0
    for r, p in zip(returns, probs):
        if p <= 0.0:
            continue
        x = 1.0 + weight * r
        if x <= 0.0:
            return -math.inf
        total += p * math.log(x)
    return total


def _log_growth_derivative(weight: float, returns: Sequence[float], probs: Sequence[float]) -> float:
    """d/dw Σ p·ln(1 + w·r) = Σ p·r/(1 + w·r)"""
    total = 0.0
    for r, p in zip(returns, probs):
        if p <= 0.0:
            continue
        x = 1.0 + weight * r
        if x <= 1e-12:
            return -math.inf
        total += p * r / x
    return total


def exact_kelly(
    returns: Sequence[float],
    probs: Sequence[float],
    tol: float = 1e-10,
    max_iter: int = 200,
) -> Tuple[float, List[str]]:
    """
    離散シナリオ分布に対する厳密なケリー比率を二分法で求める。

    正規近似 f = μ/σ² は裾が厚い分布では過大なサイズを出す。シナリオが
    離散で与えられている場合、対数成長率を直接最大化するのが正しい。

    Returns:
        (ケリー比率, 警告メッセージのリスト)
    """
    if len(returns) != len(probs):
        raise SizingError(f"長さ不一致: returns={len(returns)}, probs={len(probs)}")
    if not returns:
        raise SizingError("シナリオが空です")
    total_p = math.fsum(probs)
    if abs(total_p - 1.0) > 1e-6:
        raise SizingError(f"確率の合計が {total_p:.6f} です。1.0 にしてください。")

    notes: List[str] = []
    active = [(r, p) for r, p in zip(returns, probs) if p > 0.0]
    if not active:
        raise SizingError("正の確率を持つシナリオがありません")

    rs = [r for r, _ in active]
    ps = [p for _, p in active]

    # 期待値が0以下ならケリーは0（賭けない）
    mu = math.fsum(p * r for r, p in active)
    if mu <= 0.0:
        return (0.0, ["算術期待リターンが0以下。この賭けは見送り。"])

    worst = min(rs)
    if worst >= 0.0:
        notes.append(
            "下方シナリオがありません。損失ケースを書いていない予測は必ず過大な"
            "サイズを出します。負けシナリオを追加してください。"
        )
        return (MAX_KELLY_SEARCH, notes)

    # 1 + w·worst > 0 となる上限
    upper_bound = -1.0 / worst if worst < 0.0 else MAX_KELLY_SEARCH
    hi = min(upper_bound * (1.0 - 1e-9), MAX_KELLY_SEARCH)

    if _log_growth_derivative(hi, rs, ps) > 0.0:
        notes.append(
            f"ケリー比率が探索上限 {MAX_KELLY_SEARCH:.1f} に達しました。"
            f"シナリオ設定が楽観的すぎる可能性があります。"
        )
        return (hi, notes)

    lo = 0.0
    for _ in range(max_iter):
        mid = 0.5 * (lo + hi)
        d = _log_growth_derivative(mid, rs, ps)
        if d > 0.0:
            lo = mid
        else:
            hi = mid
        if hi - lo < tol:
            break

    return (0.5 * (lo + hi), notes)


def gaussian_kelly(edge: float, volatility: float, edge_variance: float = 0.0) -> float:
    """
    正規近似によるケリー比率 f = μ / (σ² + Var[μ])。

    Var[μ] を分母に加えるのは推定誤差を織り込むため。推定が不確かなほど
    サイズが小さくなる。
    """
    denom = volatility * volatility + edge_variance
    if denom <= 0.0:
        raise SizingError(f"分散が0以下です: σ²={volatility ** 2}, Var[μ]={edge_variance}")
    return edge / denom


def drawdown_probability(kelly_fraction_used: float, drawdown_level: float) -> float:
    """
    分数ケリーで運用したときに、いずれかの時点で drawdown_level まで
    下落する確率。

    連続時間の幾何ブラウン運動近似で、フルケリーの c 倍を賭ける場合:
        P(残高が最高値の a 倍まで下落する) = a^(2/c − 1)

    フルケリー(c=1) では P(50%下落) = 50%。これがフルケリーが実務で
    使えない理由。c=0.25 なら P(50%下落) = 0.5^7 ≒ 0.8%。

    Args:
        kelly_fraction_used: フルケリーに対する倍率 c（0 < c <= 2）
        drawdown_level: 下落後の残高比率 a（0 < a < 1）。0.5 なら「半減」

    Returns:
        到達確率
    """
    if not 0.0 < kelly_fraction_used <= 2.0:
        raise SizingError(f"kelly_fraction_used={kelly_fraction_used} は (0, 2] の範囲")
    if not 0.0 < drawdown_level < 1.0:
        raise SizingError(f"drawdown_level={drawdown_level} は (0,1) の範囲")
    exponent = 2.0 / kelly_fraction_used - 1.0
    if exponent <= 0.0:
        return 1.0
    return drawdown_level ** exponent


# ────────────────────────────────────────────────────────────────
# ベイズ縮小
# ────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class ShrunkEdge:
    raw_edge: float
    raw_stderr: float
    prior_mean: float
    prior_stdev: float
    shrinkage_weight: float     # w。1に近いほど自分の推定を信用
    posterior_mean: float
    posterior_stdev: float

    def summary(self) -> str:
        return (
            f"生の推定エッジ  : {self.raw_edge:+.4f} (SE {self.raw_stderr:.4f})\n"
            f"事前分布        : 平均 {self.prior_mean:+.4f}, 標準偏差 {self.prior_stdev:.4f}\n"
            f"縮小の重み w    : {self.shrinkage_weight:.3f}\n"
            f"事後エッジ      : {self.posterior_mean:+.4f} (SD {self.posterior_stdev:.4f})"
        )


def shrink_edge(
    edge: float,
    stderr: float,
    prior_mean: float = 0.0,
    prior_stdev: float = 0.03,
) -> ShrunkEdge:
    """
    エッジ推定値をベイズ縮小する。

    事前分布 μ ~ N(μ₀, τ²)、尤度 μ̂|μ ~ N(μ, s²) のとき
        w = τ² / (τ² + s²)
        事後平均 = w·μ̂ + (1−w)·μ₀
        事後分散 = w·s²

    既定の事前平均は0、つまり「エッジはない」が出発点。推定誤差 s が
    大きいほど推定値は0に引き戻される。

    注意: τ と s は μ と同じ単位でなければならない。保有期間リターンに
    対して「年率3%」のような絶対値の事前分布を当てると単位が壊れる。
    size_position() は τ = 情報レシオ × σ として自動的に単位を揃える。

    Args:
        edge: 生のエッジ推定値 μ̂
        stderr: μ̂ の標準誤差 s（μ と同単位）
        prior_mean: 事前平均 μ₀
        prior_stdev: 事前標準偏差 τ（μ と同単位）

    Returns:
        ShrunkEdge
    """
    if stderr < 0.0:
        raise SizingError(f"stderr={stderr} は0以上")
    if prior_stdev <= 0.0:
        raise SizingError(f"prior_stdev={prior_stdev} は正")

    tau2 = prior_stdev * prior_stdev
    s2 = stderr * stderr
    if tau2 + s2 <= 0.0:
        w = 1.0
    else:
        w = tau2 / (tau2 + s2)

    post_mean = w * edge + (1.0 - w) * prior_mean
    post_var = w * s2
    return ShrunkEdge(
        raw_edge=edge,
        raw_stderr=stderr,
        prior_mean=prior_mean,
        prior_stdev=prior_stdev,
        shrinkage_weight=w,
        posterior_mean=post_mean,
        posterior_stdev=math.sqrt(max(0.0, post_var)),
    )


# ────────────────────────────────────────────────────────────────
# キャリブレーションゲート
# ────────────────────────────────────────────────────────────────

def calibration_gate(
    report: Optional[CalibrationReport],
    bss_reference: float = BSS_REFERENCE,
) -> Tuple[float, str]:
    """
    実測キャリブレーションからサイジング倍率（0〜1）を決める。

    スキルが実証されていなければ0を返す。これは意図的に厳しい。
    「エッジがあると思う」と「エッジがあると測れた」の差が、
    株式投資で最も金を失う勘違いだからである。

    Returns:
        (倍率, 理由)
    """
    if report is None:
        return (0.0, "キャリブレーション実績なし。個別の見立てにはまだ賭けられない。")
    if not report.demonstrates_skill:
        return (0.0, report.verdict())
    gate = stats.clamp(report.brier_skill_score / bss_reference, 0.0, 1.0)
    return (
        gate,
        f"BSS={report.brier_skill_score:.3f} / 基準{bss_reference:.2f} → 倍率{gate:.2f}",
    )


# ────────────────────────────────────────────────────────────────
# 制約
# ────────────────────────────────────────────────────────────────

@dataclass
class SizingConstraints:
    """サイジングに掛かる上限。全てポートフォリオに対するウェイト。"""

    max_weight: float = 0.10
    """1銘柄あたり上限ウェイト。"""

    kelly_fraction: float = DEFAULT_KELLY_FRACTION
    """フルケリーに対する倍率。"""

    max_worst_case_loss: float = 0.05
    """最悪シナリオが実現したときのポートフォリオ損失の上限。"""

    adv_participation: float = 0.10
    """1日の出来高に対する参加率の上限。"""

    exit_days: float = 5.0
    """何営業日で全量を処分できる必要があるか。"""

    portfolio_value: float = 10_000_000.0
    """ポートフォリオ総額（円）。流動性キャップの計算に使う。"""

    def validate(self) -> None:
        if not 0.0 < self.max_weight <= 1.0:
            raise SizingError(f"max_weight={self.max_weight} は (0,1] の範囲")
        if not 0.0 < self.kelly_fraction <= 2.0:
            raise SizingError(f"kelly_fraction={self.kelly_fraction} は (0,2] の範囲")
        if not 0.0 < self.max_worst_case_loss <= 1.0:
            raise SizingError(f"max_worst_case_loss={self.max_worst_case_loss} は (0,1] の範囲")
        if not 0.0 < self.adv_participation <= 1.0:
            raise SizingError(f"adv_participation={self.adv_participation} は (0,1] の範囲")
        if self.exit_days <= 0.0:
            raise SizingError(f"exit_days={self.exit_days} は正")
        if self.portfolio_value <= 0.0:
            raise SizingError(f"portfolio_value={self.portfolio_value} は正")


def liquidity_cap(
    adv_value: float,
    constraints: SizingConstraints,
) -> float:
    """
    出口から逆算した上限ウェイト。

    「exit_days 営業日以内に、1日の出来高の adv_participation 以下で
    全量を処分できる」を満たす最大ウェイト。

    個人投資家の構造的優位は小型株にあるが、優位の源泉である低流動性は
    そのまま出口リスクでもある。買える額ではなく売れる額で決める。

    Args:
        adv_value: 平均日次売買代金（円）
        constraints: 制約

    Returns:
        上限ウェイト
    """
    if adv_value < 0.0:
        raise SizingError(f"adv_value={adv_value} は0以上")
    max_position_value = adv_value * constraints.adv_participation * constraints.exit_days
    return min(1.0, max_position_value / constraints.portfolio_value)


# ────────────────────────────────────────────────────────────────
# 結果
# ────────────────────────────────────────────────────────────────

@dataclass
class SizingResult:
    weight: float
    kelly_raw: float
    kelly_shrunk: float
    after_fraction: float
    after_gate: float
    shrunk_edge: ShrunkEdge
    gate: float
    gate_reason: str
    caps: Dict[str, float] = field(default_factory=dict)
    binding_constraint: str = ""
    notes: List[str] = field(default_factory=list)
    log_growth_at_weight: float = 0.0
    drawdown_50_prob: float = float("nan")

    @property
    def is_zero(self) -> bool:
        return self.weight <= 1e-9

    def summary(self) -> str:
        lines = [
            "── サイジング内訳 ──",
            f"厳密ケリー(生)        : {self.kelly_raw * 100:7.2f}%",
            f"エッジ縮小後ケリー    : {self.kelly_shrunk * 100:7.2f}%",
            f"分数ケリー適用後      : {self.after_fraction * 100:7.2f}%",
            f"キャリブレーション後  : {self.after_gate * 100:7.2f}%   "
            f"(倍率 {self.gate:.2f})",
            f"最終ウェイト          : {self.weight * 100:7.2f}%",
            "",
            "── 各制約が許すウェイト ──",
        ]
        for name, val in sorted(self.caps.items(), key=lambda kv: kv[1]):
            mark = " ← 効いている" if name == self.binding_constraint else ""
            lines.append(f"  {name:<24}{val * 100:7.2f}%{mark}")
        lines += [
            "",
            f"対数期待成長率        : {self.log_growth_at_weight:+.5f} / 保有期間",
            f"半減到達確率(概算)    : {self.drawdown_50_prob * 100:.2f}%",
            "",
            f"ゲート理由: {self.gate_reason}",
        ]
        if self.notes:
            lines.append("")
            lines.append("警告:")
            for n in self.notes:
                lines.append(f"  - {n}")
        return "\n".join(lines)


# ────────────────────────────────────────────────────────────────
# 本体
# ────────────────────────────────────────────────────────────────

def size_position(
    forecast: Forecast,
    constraints: Optional[SizingConstraints] = None,
    calibration: Optional[CalibrationReport] = None,
    edge_stderr: Optional[float] = None,
    prior_information_ratio: float = DEFAULT_PRIOR_INFORMATION_RATIO,
    adv_value: Optional[float] = None,
) -> SizingResult:
    """
    予測1件に対する目標ウェイトを算出する。

    エッジの事前分布はシナリオの σ に比例させる（τ = 情報レシオ × σ）ため、
    保有期間の長さや銘柄のボラティリティが変わっても単位が壊れない。

    Args:
        forecast: シナリオ付きの Forecast
        constraints: 制約。省略時は既定値
        calibration: 過去実績の CalibrationReport。None なら倍率0（＝賭けない）
        edge_stderr: 自分の中心推定値の誤差。σ と同単位。省略時は
                     DEFAULT_ESTIMATE_ERROR_RATIO × σ。σ（結果のばらつき）
                     とは別概念であり、原理的な既定値は存在しない。
                     根拠があるなら明示的に渡すこと。
        prior_information_ratio: 事前分布の情報レシオ。τ = この値 × σ
        adv_value: 平均日次売買代金（円）。渡すと流動性キャップが効く

    Returns:
        SizingResult
    """
    cons = constraints or SizingConstraints()
    cons.validate()
    if not forecast.scenarios:
        raise SizingError("scenarios が未設定です。サイジングにはリターン分布が必要です。")

    returns = [s.total_return for s in forecast.scenarios]
    probs = [s.probability for s in forecast.scenarios]
    notes: List[str] = []

    # 1. 厳密ケリー
    kelly_raw, kelly_notes = exact_kelly(returns, probs)
    notes.extend(kelly_notes)

    # 2. エッジのベイズ縮小（事前分布は σ に比例させて単位を揃える）
    mu = forecast.expected_return()
    sigma = forecast.sigma()
    if sigma <= 0.0:
        raise SizingError(
            "シナリオの標準偏差が0です。全シナリオが同一リターンでは賭けになりません。"
        )
    if prior_information_ratio <= 0.0:
        raise SizingError(
            f"prior_information_ratio={prior_information_ratio} は正でなければならない"
        )
    se = edge_stderr if edge_stderr is not None else DEFAULT_ESTIMATE_ERROR_RATIO * sigma
    shrunk = shrink_edge(
        mu, se, prior_mean=0.0, prior_stdev=prior_information_ratio * sigma
    )

    # 縮小後の平均に合わせてシナリオ分布を平行移動し、厳密ケリーを再計算する。
    # 分布の形（裾の厚さ）を保ったまま期待値だけを事後平均に置き換える。
    shift = shrunk.posterior_mean - mu
    shifted = [r + shift for r in returns]
    # 平行移動で -100% を割り込む場合はクリップして破産の扱いを保つ
    shifted = [max(r, -1.0) for r in shifted]
    kelly_shrunk, shrunk_notes = exact_kelly(shifted, probs)
    for n in shrunk_notes:
        if n not in notes:
            notes.append(n)

    # 3. 分数ケリー
    after_fraction = kelly_shrunk * cons.kelly_fraction

    # 4. キャリブレーションゲート
    gate, gate_reason = calibration_gate(calibration)
    after_gate = after_fraction * gate

    # 5. キャップ
    caps: Dict[str, float] = {
        "分数ケリー": after_gate,
        "1銘柄上限": cons.max_weight,
    }

    worst = forecast.worst_case()
    if worst < 0.0:
        caps["最悪シナリオ損失上限"] = cons.max_worst_case_loss / abs(worst)
    if adv_value is not None:
        caps["流動性(出口)"] = liquidity_cap(adv_value, cons)

    binding = min(caps, key=lambda k: caps[k])
    weight = max(0.0, min(caps.values()))

    # 6. 対数成長率とドローダウン確率
    lg = log_growth(weight, returns, probs) if weight > 0 else 0.0
    if lg < 0.0 and weight > 0.0:
        notes.append(
            f"この賭けはウェイト{weight * 100:.1f}%で対数期待成長率がマイナス"
            f"({lg:+.5f})。算術期待値がプラスでも長期では資産を減らす。"
        )

    if kelly_shrunk > 1e-9 and weight > 1e-9:
        c = stats.clamp(weight / kelly_shrunk, 1e-9, 2.0)
        dd50 = drawdown_probability(c, 0.5)
    else:
        dd50 = 0.0

    if forecast.expected_log_return() == -math.inf:
        notes.append(
            "全損シナリオ（-100%）に正の確率があります。対数期待成長率は -inf。"
            "この銘柄に資産を集中させてはいけません。"
        )

    return SizingResult(
        weight=weight,
        kelly_raw=kelly_raw,
        kelly_shrunk=kelly_shrunk,
        after_fraction=after_fraction,
        after_gate=after_gate,
        shrunk_edge=shrunk,
        gate=gate,
        gate_reason=gate_reason,
        caps=caps,
        binding_constraint=binding,
        notes=notes,
        log_growth_at_weight=lg,
        drawdown_50_prob=dd50,
    )


# ────────────────────────────────────────────────────────────────
# ポートフォリオ配分
# ────────────────────────────────────────────────────────────────

@dataclass
class PortfolioAllocation:
    core_weight: float
    satellite_weights: Dict[str, float]
    cash_weight: float
    satellite_budget: float
    scaled_by: float
    notes: List[str] = field(default_factory=list)

    @property
    def total_satellite(self) -> float:
        return math.fsum(self.satellite_weights.values())

    def summary(self) -> str:
        lines = [
            f"コア（低コスト市場全体）: {self.core_weight * 100:6.2f}%",
            f"サテライト合計          : {self.total_satellite * 100:6.2f}%  "
            f"(予算 {self.satellite_budget * 100:.1f}%)",
            f"現金                    : {self.cash_weight * 100:6.2f}%",
        ]
        if self.scaled_by < 1.0:
            lines.append(f"※ 予算超過のため一律 {self.scaled_by:.3f} 倍に縮小")
        if self.satellite_weights:
            lines.append("")
            lines.append("サテライト内訳:")
            for t, w in sorted(self.satellite_weights.items(), key=lambda kv: -kv[1]):
                lines.append(f"  {t:<12}{w * 100:6.2f}%")
        for n in self.notes:
            lines.append(f"  - {n}")
        return "\n".join(lines)


def allocate_portfolio(
    sized: Dict[str, SizingResult],
    satellite_budget: float = 0.20,
    sector_of: Optional[Dict[str, str]] = None,
    max_sector_weight: float = 0.10,
) -> PortfolioAllocation:
    """
    個別サイジング結果をポートフォリオに組み上げる。

    コア・サテライト構成は妥協案ではない。自分のアルファ推定値自体に
    大きな推定誤差があるとき、期待値最大化の解がこの形になる。
    サテライトが全部外れてもコアの市場リターンは残る。

    Args:
        sized: {ティッカー: SizingResult}
        satellite_budget: サテライトに割ける上限（ポートフォリオ比）
        sector_of: {ティッカー: セクター名}
        max_sector_weight: 同一セクターの上限ウェイト

    Returns:
        PortfolioAllocation
    """
    if not 0.0 <= satellite_budget <= 1.0:
        raise SizingError(f"satellite_budget={satellite_budget} は [0,1] の範囲")

    weights = {t: r.weight for t, r in sized.items() if r.weight > 1e-9}
    notes: List[str] = []

    # セクターキャップ
    if sector_of:
        by_sector: Dict[str, List[str]] = {}
        for t in weights:
            by_sector.setdefault(sector_of.get(t, "未分類"), []).append(t)
        for sec, members in by_sector.items():
            total = math.fsum(weights[t] for t in members)
            if total > max_sector_weight:
                scale = max_sector_weight / total
                for t in members:
                    weights[t] *= scale
                notes.append(
                    f"セクター『{sec}』が上限{max_sector_weight * 100:.0f}%を超過 "
                    f"({total * 100:.1f}%) → {scale:.3f}倍に縮小"
                )

    # サテライト予算
    total = math.fsum(weights.values())
    scaled_by = 1.0
    if total > satellite_budget and total > 0.0:
        scaled_by = satellite_budget / total
        weights = {t: w * scaled_by for t, w in weights.items()}
        total = math.fsum(weights.values())

    core = 1.0 - total
    if not weights:
        notes.append(
            "サテライトが0件。キャリブレーションが実証されるまでは全額コアが正解。"
        )

    return PortfolioAllocation(
        core_weight=core,
        satellite_weights=weights,
        cash_weight=0.0,
        satellite_budget=satellite_budget,
        scaled_by=scaled_by,
        notes=notes,
    )
