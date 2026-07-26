"""
構造的ユニバース選別 — 大口資金が物理的に入れない場所を探す

歪みが残るのは「誰も気づいていない場所」ではない。気づいていても
入れない場所である。裁定の限界（limits to arbitrage）。

運用額1000億円のファンドは、意味のあるポジション（例えば運用額の0.5%
= 5億円）を作るために、5億円を数日で買って数日で売れる流動性を必要とする。
1日の売買代金が3000万円の会社では、それが物理的にできない。だから
どんなに割安でも買わない。買えないのだ。

このモジュールは各銘柄について「意味のあるポジションを作れる最大の
ファンド規模」を計算する。この数字が小さい銘柄こそ、個人が構造的に
優位に立てる場所である。

同時に、その優位の源泉である低流動性は自分にとっての出口リスクでもある。
だから「彼らには小さすぎるが、自分には十分」という帯域を探す。
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence

from . import stats

# 機関投資家が「意味のある」とみなすポジションの最小比率（運用額に対して）
INSTITUTIONAL_MIN_POSITION_PCT = 0.005   # 0.5%

# 機関投資家の想定参加率と建玉日数
INSTITUTIONAL_PARTICIPATION = 0.15
INSTITUTIONAL_BUILD_DAYS = 10.0


@dataclass
class Security:
    """選別に必要な最小限の銘柄情報。"""
    ticker: str
    name: str = ""
    market_cap: float = 0.0            # 時価総額（円）
    adv_value: float = 0.0             # 平均日次売買代金（円）
    price: float = 0.0
    analyst_count: int = 0             # アナリストカバレッジ数
    sector: str = ""
    net_cash: Optional[float] = None   # 純現金 = 現金等 − 有利子負債（円）
    pbr: Optional[float] = None
    free_float: Optional[float] = None # 浮動株比率 0〜1
    is_parent_subsidiary: bool = False # 親子上場・持株会社ディスカウント等
    forced_seller_flag: bool = False   # 指数除外・税金売り等の強制売却が観測されるか

    def validate(self) -> None:
        if not self.ticker.strip():
            raise ValueError("ticker が空です")
        if self.market_cap < 0 or self.adv_value < 0:
            raise ValueError(f"{self.ticker}: 時価総額・売買代金は0以上")
        if self.analyst_count < 0:
            raise ValueError(f"{self.ticker}: analyst_count は0以上")
        if self.free_float is not None and not 0.0 <= self.free_float <= 1.0:
            raise ValueError(f"{self.ticker}: free_float={self.free_float} は [0,1]")

    @property
    def net_cash_ratio(self) -> Optional[float]:
        """純現金 / 時価総額。1超なら現金だけで時価総額を上回る。"""
        if self.net_cash is None or self.market_cap <= 0:
            return None
        return self.net_cash / self.market_cap

    @property
    def turnover_ratio(self) -> Optional[float]:
        """年間回転率の目安。低いほど機関が動きにくい。"""
        if self.market_cap <= 0:
            return None
        return self.adv_value * 250.0 / self.market_cap


@dataclass
class InvestorProfile:
    """自分の側の制約。"""
    portfolio_value: float = 10_000_000.0
    target_position_weight: float = 0.05    # 1銘柄に入れたいウェイト
    adv_participation: float = 0.10         # 自分が許容する出来高参加率
    exit_days: float = 5.0                  # 何日で処分できる必要があるか
    min_price: float = 100.0                # 単元・呼値の実務上の下限

    def validate(self) -> None:
        if self.portfolio_value <= 0:
            raise ValueError("portfolio_value は正")
        if not 0.0 < self.target_position_weight <= 1.0:
            raise ValueError("target_position_weight は (0,1]")
        if not 0.0 < self.adv_participation <= 1.0:
            raise ValueError("adv_participation は (0,1]")
        if self.exit_days <= 0:
            raise ValueError("exit_days は正")

    @property
    def target_position_value(self) -> float:
        return self.portfolio_value * self.target_position_weight


def largest_participating_fund(
    adv_value: float,
    min_position_pct: float = INSTITUTIONAL_MIN_POSITION_PCT,
    participation: float = INSTITUTIONAL_PARTICIPATION,
    build_days: float = INSTITUTIONAL_BUILD_DAYS,
) -> float:
    """
    この銘柄に「意味のあるポジション」を作れる最大のファンド規模（円）。

    ファンドが build_days 営業日で、1日の出来高の participation 以下の
    参加率で建てられる最大額を、意味のある最小ポジション比率で割る。

    この値が小さい銘柄ほど、機関投資家が構造的に排除されている。

    Args:
        adv_value: 平均日次売買代金（円）
        min_position_pct: 意味のあるポジションの最小比率
        participation: 出来高参加率
        build_days: 建玉に許容する日数

    Returns:
        参加可能な最大ファンド規模（円）
    """
    if adv_value < 0:
        raise ValueError(f"adv_value={adv_value} は0以上")
    if not 0.0 < min_position_pct <= 1.0:
        raise ValueError(f"min_position_pct={min_position_pct} は (0,1]")
    max_position = adv_value * participation * build_days
    return max_position / min_position_pct


@dataclass
class ScreenResult:
    security: Security
    passed: bool
    reasons: List[str] = field(default_factory=list)      # 落選理由
    edge_flags: List[str] = field(default_factory=list)   # 構造的優位の根拠
    largest_fund: float = 0.0
    my_capacity_value: float = 0.0
    capacity_ratio: float = 0.0        # 自分の目標建玉に対する余裕度（1以上必要）
    structural_score: float = 0.0      # 0〜100

    def summary(self) -> str:
        status = "通過" if self.passed else "除外"
        lines = [
            f"[{status}] {self.security.ticker} {self.security.name}",
            f"  構造スコア      : {self.structural_score:.1f}/100",
            f"  参加可能最大ファンド: {self.largest_fund / 1e8:,.1f}億円",
            f"  自分の建玉余裕度: {self.capacity_ratio:.2f}倍 "
            f"(建てられる額 {self.my_capacity_value / 1e4:,.0f}万円)",
        ]
        if self.edge_flags:
            lines.append(f"  構造的優位      : {', '.join(self.edge_flags)}")
        if self.reasons:
            lines.append(f"  除外理由        : {', '.join(self.reasons)}")
        return "\n".join(lines)


@dataclass
class ScreenCriteria:
    """
    選別基準。

    既定値は「東証の小型株で、機関が入れず、自分は出口を確保できる」帯域。
    """
    max_participating_fund: float = 50e8      # 参加可能最大ファンドの上限（50億円）
    max_analyst_count: int = 2                # カバレッジ数の上限
    min_market_cap: float = 1e8               # 時価総額の下限（1億円。上場維持リスク回避）
    max_market_cap: float = 1000e8            # 時価総額の上限（1000億円）
    min_capacity_ratio: float = 1.0           # 自分の建玉が入りきるか
    require_positive_market_cap: bool = True

    def validate(self) -> None:
        if self.min_market_cap >= self.max_market_cap:
            raise ValueError("min_market_cap < max_market_cap が必要")
        if self.max_analyst_count < 0:
            raise ValueError("max_analyst_count は0以上")
        if self.min_capacity_ratio <= 0:
            raise ValueError("min_capacity_ratio は正")


def screen_security(
    sec: Security,
    profile: InvestorProfile,
    criteria: Optional[ScreenCriteria] = None,
) -> ScreenResult:
    """
    1銘柄を構造的観点で選別する。

    ファクター（バリュー・モメンタム等）の評価は行わない。ここで見るのは
    「そもそも大口が入れない場所か」「自分は出口を確保できるか」だけ。
    """
    crit = criteria or ScreenCriteria()
    crit.validate()
    profile.validate()
    sec.validate()

    reasons: List[str] = []
    flags: List[str] = []

    largest = largest_participating_fund(sec.adv_value)
    my_capacity = sec.adv_value * profile.adv_participation * profile.exit_days
    target = profile.target_position_value
    capacity_ratio = my_capacity / target if target > 0 else 0.0

    # ── 除外条件 ──
    if crit.require_positive_market_cap and sec.market_cap <= 0:
        reasons.append("時価総額が不明または0")
    if sec.market_cap > 0:
        if sec.market_cap < crit.min_market_cap:
            reasons.append(
                f"時価総額 {sec.market_cap / 1e8:.1f}億円 < 下限 {crit.min_market_cap / 1e8:.1f}億円"
            )
        if sec.market_cap > crit.max_market_cap:
            reasons.append(
                f"時価総額 {sec.market_cap / 1e8:.0f}億円 > 上限 {crit.max_market_cap / 1e8:.0f}億円"
                f"（機関が普通に買える規模）"
            )
    if sec.analyst_count > crit.max_analyst_count:
        reasons.append(
            f"カバレッジ {sec.analyst_count}社 > 上限 {crit.max_analyst_count}社"
            f"（既に情報が行き渡っている）"
        )
    if largest > crit.max_participating_fund:
        reasons.append(
            f"参加可能ファンド規模 {largest / 1e8:.0f}億円 > 上限 "
            f"{crit.max_participating_fund / 1e8:.0f}億円（機関が入れる）"
        )
    if capacity_ratio < crit.min_capacity_ratio:
        reasons.append(
            f"自分の建玉が入らない（余裕度 {capacity_ratio:.2f}倍 < "
            f"{crit.min_capacity_ratio:.2f}倍）。出口が確保できない"
        )
    if sec.price > 0 and sec.price < profile.min_price:
        reasons.append(f"株価 {sec.price:.0f}円 < 下限 {profile.min_price:.0f}円")

    # ── 構造的優位のフラグ ──
    if sec.analyst_count == 0:
        flags.append("カバレッジゼロ")
    if largest <= 20e8:
        flags.append(f"機関ほぼ参加不能（最大{largest / 1e8:.0f}億円）")
    ncr = sec.net_cash_ratio
    if ncr is not None and ncr >= 0.5:
        flags.append(f"純現金が時価総額の{ncr * 100:.0f}%")
    if sec.pbr is not None and sec.pbr < 1.0:
        flags.append(f"PBR {sec.pbr:.2f}（東証の改善要請対象）")
    if sec.is_parent_subsidiary:
        flags.append("親子上場・SOTPディスカウント")
    if sec.forced_seller_flag:
        flags.append("強制売却の観測")
    if sec.free_float is not None and sec.free_float < 0.3:
        flags.append(f"浮動株{sec.free_float * 100:.0f}%")

    score = _structural_score(sec, largest, capacity_ratio)

    return ScreenResult(
        security=sec,
        passed=not reasons,
        reasons=reasons,
        edge_flags=flags,
        largest_fund=largest,
        my_capacity_value=my_capacity,
        capacity_ratio=capacity_ratio,
        structural_score=score,
    )


def _structural_score(sec: Security, largest_fund: float, capacity_ratio: float) -> float:
    """
    構造的優位の総合スコア（0〜100）。内訳を説明できる形で加算する。

      機関排除度   最大40点  参加可能ファンドが小さいほど高い（対数スケール）
      カバレッジ   最大20点  0社で満点
      触媒         最大25点  純現金・PBR1倍割れ・親子上場・強制売却
      自分の余裕   最大15点  出口が確保できているか
    """
    # 機関排除度: 1000億円で0点、1億円で40点（対数線形）
    if largest_fund <= 0:
        exclusion = 40.0
    else:
        lo, hi = math.log10(1e8), math.log10(1000e8)
        x = stats.clamp(math.log10(max(largest_fund, 1.0)), lo, hi)
        exclusion = 40.0 * (hi - x) / (hi - lo)

    # カバレッジ
    coverage = max(0.0, 20.0 - 6.0 * sec.analyst_count)

    # 触媒
    catalyst = 0.0
    ncr = sec.net_cash_ratio
    if ncr is not None:
        catalyst += stats.clamp(ncr, 0.0, 1.0) * 10.0
    if sec.pbr is not None and sec.pbr < 1.0:
        catalyst += stats.clamp((1.0 - sec.pbr) * 10.0, 0.0, 6.0)
    if sec.is_parent_subsidiary:
        catalyst += 5.0
    if sec.forced_seller_flag:
        catalyst += 4.0
    catalyst = min(catalyst, 25.0)

    # 自分の余裕（1倍で0点、5倍以上で満点）
    capacity = stats.clamp((capacity_ratio - 1.0) / 4.0, 0.0, 1.0) * 15.0

    return stats.clamp(exclusion + coverage + catalyst + capacity, 0.0, 100.0)


def screen_universe(
    securities: Sequence[Security],
    profile: InvestorProfile,
    criteria: Optional[ScreenCriteria] = None,
) -> List[ScreenResult]:
    """
    銘柄群を選別し、構造スコアの降順で返す。

    重要: この関数はリターン実績を一切参照しない。過去の成績で銘柄を
    選ぶと選抜バイアスが入り、バックテストが無意味になる
    （ev.validation.selection_inflation を参照）。事前に観測可能な
    構造的特徴だけで絞ることが、バイアスを避ける唯一の方法である。
    """
    results = [screen_security(s, profile, criteria) for s in securities]
    return sorted(results, key=lambda r: (-int(r.passed), -r.structural_score))


def universe_summary(results: Sequence[ScreenResult]) -> str:
    passed = [r for r in results if r.passed]
    lines = [
        f"選別結果: {len(passed)} / {len(results)} 銘柄が通過",
        "",
    ]
    if passed:
        lines.append("通過銘柄（構造スコア順）:")
        for r in passed:
            lines.append(r.summary())
    rejected = [r for r in results if not r.passed]
    if rejected:
        lines += ["", f"除外 {len(rejected)}銘柄の主な理由:"]
        counts: Dict[str, int] = {}
        for r in rejected:
            for reason in r.reasons:
                key = reason.split("（")[0].split(" ")[0]
                counts[key] = counts.get(key, 0) + 1
        for k, v in sorted(counts.items(), key=lambda kv: -kv[1]):
            lines.append(f"  {k}: {v}件")
    return "\n".join(lines)
