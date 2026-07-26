"""
事前確約型の予測記録

買う前に、反証可能な形で確率を書く。買った後に理由を作らない。

このモジュールが強制するルール:
  1. 確率は 0 でも 1 でもない値でなければならない（確実は存在しない）
  2. 各主張には期限と判定方法が必須（「いつ・何を見れば外れたと分かるか」）
  3. 期限は記録時点より未来（後から日付を遡って書けない）
  4. ベースレートから大きく離れた確率を置くには根拠の記述が必須
  5. シナリオ確率の合計は1
  6. 記録は追記専用台帳へ。解決も追記。上書きは物理的に不可能

シナリオからは算術期待値と対数期待成長率の両方を計算する。
算術期待値がプラスでも対数期待成長率がマイナスなら、その賭けは
長期的に資産を減らす。サイジングは後者を見る。
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from datetime import date
from typing import Any, Dict, List, Optional, Sequence

from . import stats
from .ledger import Ledger

# 確率の許容範囲。0/1を許すとロジットが発散するだけでなく、
# 「絶対」という表現自体がキャリブレーションを壊す。
MIN_PROBABILITY = 0.02
MAX_PROBABILITY = 0.98

# ベースレートからこれ以上（ロジット差）離れるなら根拠の記述を要求する
BASE_RATE_DIVERGENCE_LIMIT = 1.0


class ForecastValidationError(ValueError):
    """予測の記録が規約を満たしていない。"""


# ────────────────────────────────────────────────────────────────
# 反証可能な主張
# ────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class Claim:
    """
    ひとつの反証可能な主張。

    良い例: 「2026年3月期の営業利益が会社予想の120億円を上回る」
            deadline='2026-05-15', resolution_source='決算短信の営業利益'
    悪い例: 「この会社は成長性がある」（判定できない = 記録する価値がない）
    """
    statement: str
    probability: float
    deadline: str                      # ISO形式の日付 'YYYY-MM-DD'
    resolution_source: str             # 何を見て判定するか
    claim_id: str = ""
    base_rate: Optional[float] = None  # 参照した経験的ベースレート
    divergence_rationale: str = ""     # ベースレートから離れる理由

    def validate(self, as_of: date) -> None:
        if not self.statement or len(self.statement.strip()) < 5:
            raise ForecastValidationError(
                f"claim '{self.claim_id}': statement が短すぎます。"
                f"何が起きたら外れなのかを具体的に書いてください。"
            )
        if not (MIN_PROBABILITY <= self.probability <= MAX_PROBABILITY):
            raise ForecastValidationError(
                f"claim '{self.claim_id}': probability={self.probability} は "
                f"[{MIN_PROBABILITY}, {MAX_PROBABILITY}] の範囲外です。"
                f"確実だと思っているなら、それがまさに測るべき思い込みです。"
            )
        if not self.resolution_source.strip():
            raise ForecastValidationError(
                f"claim '{self.claim_id}': resolution_source が空です。"
                f"判定方法を決めずに書いた予測は、後から都合よく解釈されます。"
            )
        try:
            dl = date.fromisoformat(self.deadline)
        except ValueError as e:
            raise ForecastValidationError(
                f"claim '{self.claim_id}': deadline='{self.deadline}' が日付として不正: {e}"
            ) from e
        if dl <= as_of:
            raise ForecastValidationError(
                f"claim '{self.claim_id}': deadline={self.deadline} が記録日 {as_of} 以前です。"
                f"結果が既に判明している事象は予測になりません。"
            )
        if self.base_rate is not None:
            if not 0.0 < self.base_rate < 1.0:
                raise ForecastValidationError(
                    f"claim '{self.claim_id}': base_rate={self.base_rate} は (0,1) の範囲外です。"
                )
            gap = abs(stats.logit(self.probability) - stats.logit(self.base_rate))
            if gap > BASE_RATE_DIVERGENCE_LIMIT and not self.divergence_rationale.strip():
                raise ForecastValidationError(
                    f"claim '{self.claim_id}': 確率 {self.probability:.2f} が "
                    f"ベースレート {self.base_rate:.2f} から大きく離れています "
                    f"(ロジット差 {gap:.2f})。divergence_rationale に "
                    f"「なぜこの銘柄は平均と違うのか」を書いてください。"
                )

    def to_dict(self) -> Dict[str, Any]:
        return {
            "claim_id": self.claim_id,
            "statement": self.statement,
            "probability": self.probability,
            "deadline": self.deadline,
            "resolution_source": self.resolution_source,
            "base_rate": self.base_rate,
            "divergence_rationale": self.divergence_rationale,
        }

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "Claim":
        return Claim(
            statement=d["statement"],
            probability=d["probability"],
            deadline=d["deadline"],
            resolution_source=d["resolution_source"],
            claim_id=d.get("claim_id", ""),
            base_rate=d.get("base_rate"),
            divergence_rationale=d.get("divergence_rationale", ""),
        )


# ────────────────────────────────────────────────────────────────
# シナリオ（サイジング用のリターン分布）
# ────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class Scenario:
    """保有期間トータルのリターンシナリオ。total_return=0.5 で +50%。"""
    name: str
    probability: float
    total_return: float

    def validate(self) -> None:
        if not 0.0 <= self.probability <= 1.0:
            raise ForecastValidationError(
                f"scenario '{self.name}': probability={self.probability} が [0,1] の外"
            )
        if self.total_return < -1.0:
            raise ForecastValidationError(
                f"scenario '{self.name}': total_return={self.total_return} は -1 未満（-100%超の損失）"
            )

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "probability": self.probability,
            "total_return": self.total_return,
        }


# ────────────────────────────────────────────────────────────────
# 予測
# ────────────────────────────────────────────────────────────────

@dataclass
class Forecast:
    ticker: str
    thesis: str
    horizon_months: int
    claims: List[Claim] = field(default_factory=list)
    scenarios: List[Scenario] = field(default_factory=list)
    structural_reason: str = ""     # なぜ大口が買えない/見ていないのか
    forecast_id: str = ""
    created_at: str = ""

    # ── 検証 ────────────────────────────────────────────

    def validate(self, as_of: Optional[date] = None) -> None:
        as_of = as_of or date.today()
        if not self.ticker.strip():
            raise ForecastValidationError("ticker が空です")
        if len(self.thesis.strip()) < 20:
            raise ForecastValidationError(
                "thesis が短すぎます。何を賭けているのかを20文字以上で書いてください。"
            )
        if self.horizon_months < 1:
            raise ForecastValidationError(
                f"horizon_months={self.horizon_months} は1以上にしてください"
            )
        if not self.claims:
            raise ForecastValidationError(
                "claims が空です。確率を書かない投資判断は、後から検証できません。"
            )

        seen = set()
        for c in self.claims:
            if c.claim_id in seen:
                raise ForecastValidationError(f"claim_id '{c.claim_id}' が重複しています")
            seen.add(c.claim_id)
            c.validate(as_of)

        if self.scenarios:
            for s in self.scenarios:
                s.validate()
            total = math.fsum(s.probability for s in self.scenarios)
            if abs(total - 1.0) > 1e-6:
                raise ForecastValidationError(
                    f"シナリオ確率の合計が {total:.6f} です。1.0 にしてください。"
                )

    # ── 期待値の計算 ────────────────────────────────────

    def expected_return(self) -> float:
        """算術期待リターン（保有期間トータル）。"""
        self._require_scenarios()
        return math.fsum(s.probability * s.total_return for s in self.scenarios)

    def expected_log_return(self) -> float:
        """
        対数期待成長率 E[ln(1+r)]。長期の資産成長を支配するのはこちら。

        全損シナリオ（total_return = -1）に正の確率があれば -inf を返す。
        これは数学的に正しく、「その賭けに資産を全額投じてはならない」を意味する。
        """
        self._require_scenarios()
        total = 0.0
        for s in self.scenarios:
            if s.probability <= 0.0:
                continue
            if s.total_return <= -1.0:
                return -math.inf
            total += s.probability * math.log1p(s.total_return)
        return total

    def sigma(self) -> float:
        """シナリオ分布の標準偏差（保有期間トータル）。"""
        self._require_scenarios()
        mu = self.expected_return()
        var = math.fsum(s.probability * (s.total_return - mu) ** 2 for s in self.scenarios)
        return math.sqrt(max(0.0, var))

    def probability_of_loss(self) -> float:
        self._require_scenarios()
        return math.fsum(s.probability for s in self.scenarios if s.total_return < 0.0)

    def worst_case(self) -> float:
        self._require_scenarios()
        return min(s.total_return for s in self.scenarios if s.probability > 0.0)

    def annualized_expected_return(self) -> float:
        """算術期待リターンを年率換算（複利ベース）。"""
        mu = self.expected_return()
        if mu <= -1.0:
            return -1.0
        years = self.horizon_months / 12.0
        if years <= 0:
            raise ForecastValidationError("horizon_months が不正です")
        return (1.0 + mu) ** (1.0 / years) - 1.0

    def _require_scenarios(self) -> None:
        if not self.scenarios:
            raise ForecastValidationError(
                "scenarios が未設定です。サイジングにはリターン分布が必要です。"
            )

    # ── 直列化 ──────────────────────────────────────────

    def to_dict(self) -> Dict[str, Any]:
        d: Dict[str, Any] = {
            "forecast_id": self.forecast_id,
            "ticker": self.ticker,
            "thesis": self.thesis,
            "horizon_months": self.horizon_months,
            "structural_reason": self.structural_reason,
            "created_at": self.created_at,
            "claims": [c.to_dict() for c in self.claims],
            "scenarios": [s.to_dict() for s in self.scenarios],
        }
        if self.scenarios:
            d["derived"] = {
                "expected_return": self.expected_return(),
                "expected_log_return": self.expected_log_return(),
                "sigma": self.sigma(),
                "probability_of_loss": self.probability_of_loss(),
                "worst_case": self.worst_case(),
            }
        return d

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "Forecast":
        return Forecast(
            ticker=d["ticker"],
            thesis=d["thesis"],
            horizon_months=d["horizon_months"],
            claims=[Claim.from_dict(c) for c in d.get("claims", [])],
            scenarios=[
                Scenario(s["name"], s["probability"], s["total_return"])
                for s in d.get("scenarios", [])
            ],
            structural_reason=d.get("structural_reason", ""),
            forecast_id=d.get("forecast_id", ""),
            created_at=d.get("created_at", ""),
        )


# ────────────────────────────────────────────────────────────────
# 解決済みの主張
# ────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class ResolvedClaim:
    forecast_id: str
    claim_id: str
    ticker: str
    statement: str
    probability: float
    base_rate: Optional[float]
    outcome: bool
    resolved_at: str
    note: str


# ────────────────────────────────────────────────────────────────
# 予測ジャーナル
# ────────────────────────────────────────────────────────────────

class ForecastJournal:
    """
    台帳の上に構築した予測ジャーナル。

    使用例:
        with ForecastJournal("journal.db") as j:
            fid = j.record(forecast)
            j.resolve(fid, "c1", outcome=True, note="決算短信で確認")
            pairs = j.resolved_claims()
    """

    KIND_FORECAST = "forecast"
    KIND_RESOLUTION = "resolution"
    KIND_TRADE = "trade"

    def __init__(self, path: str, clock: Optional[Any] = None):
        self.ledger = Ledger(path, clock=clock)

    def __enter__(self) -> "ForecastJournal":
        return self

    def __exit__(self, *exc: Any) -> None:
        self.close()

    def close(self) -> None:
        self.ledger.close()

    # ── 記録 ────────────────────────────────────────────

    def record(self, forecast: Forecast, as_of: Optional[date] = None) -> str:
        """
        予測を記録し forecast_id を返す。検証に落ちた予測は記録されない。
        """
        as_of = as_of or date.today()

        n_existing = len(self.ledger.find(self.KIND_FORECAST))
        if not forecast.forecast_id:
            forecast.forecast_id = f"FC{n_existing + 1:05d}"
        if not forecast.created_at:
            forecast.created_at = as_of.isoformat()

        # claim_id の自動採番
        for i, c in enumerate(forecast.claims, start=1):
            if not c.claim_id:
                forecast.claims[i - 1] = Claim(
                    statement=c.statement,
                    probability=c.probability,
                    deadline=c.deadline,
                    resolution_source=c.resolution_source,
                    claim_id=f"c{i}",
                    base_rate=c.base_rate,
                    divergence_rationale=c.divergence_rationale,
                )

        forecast.validate(as_of)

        if self.ledger.find(self.KIND_FORECAST, forecast_id=forecast.forecast_id):
            raise ForecastValidationError(
                f"forecast_id '{forecast.forecast_id}' は既に存在します"
            )

        self.ledger.append(self.KIND_FORECAST, forecast.to_dict())
        return forecast.forecast_id

    def resolve(
        self,
        forecast_id: str,
        claim_id: str,
        outcome: bool,
        note: str = "",
        resolved_at: Optional[date] = None,
    ) -> None:
        """
        主張の結果を記録する。既に解決済みなら拒否する（付け替え防止）。
        """
        fc = self.get(forecast_id)
        if fc is None:
            raise ForecastValidationError(f"forecast_id '{forecast_id}' が見つかりません")
        if not any(c.claim_id == claim_id for c in fc.claims):
            raise ForecastValidationError(
                f"forecast '{forecast_id}' に claim_id '{claim_id}' はありません"
            )
        already = self.ledger.find(
            self.KIND_RESOLUTION, forecast_id=forecast_id, claim_id=claim_id
        )
        if already:
            raise ForecastValidationError(
                f"claim '{forecast_id}/{claim_id}' は既に "
                f"{already[0].payload.get('outcome')} として解決済みです。"
                f"解決の書き換えはできません。"
            )
        self.ledger.append(
            self.KIND_RESOLUTION,
            {
                "forecast_id": forecast_id,
                "claim_id": claim_id,
                "outcome": bool(outcome),
                "note": note,
                "resolved_at": (resolved_at or date.today()).isoformat(),
            },
        )

    def record_trade(
        self,
        forecast_id: str,
        action: str,
        price: float,
        shares: float,
        weight: float,
        note: str = "",
        traded_at: Optional[date] = None,
    ) -> None:
        """建玉・処分の記録。予測と実行を紐付ける。"""
        if action not in ("buy", "sell"):
            raise ValueError(f"action は 'buy' か 'sell'（受領: {action}）")
        if self.get(forecast_id) is None:
            raise ForecastValidationError(f"forecast_id '{forecast_id}' が見つかりません")
        self.ledger.append(
            self.KIND_TRADE,
            {
                "forecast_id": forecast_id,
                "action": action,
                "price": price,
                "shares": shares,
                "weight": weight,
                "note": note,
                "traded_at": (traded_at or date.today()).isoformat(),
            },
        )

    # ── 参照 ────────────────────────────────────────────

    def get(self, forecast_id: str) -> Optional[Forecast]:
        found = self.ledger.find(self.KIND_FORECAST, forecast_id=forecast_id)
        if not found:
            return None
        return Forecast.from_dict(found[0].payload)

    def all_forecasts(self) -> List[Forecast]:
        return [Forecast.from_dict(e.payload) for e in self.ledger.entries(self.KIND_FORECAST)]

    def resolved_claims(self) -> List[ResolvedClaim]:
        """キャリブレーション評価の入力となる (確率, 結果) の組。"""
        by_id = {f.forecast_id: f for f in self.all_forecasts()}
        out: List[ResolvedClaim] = []
        for e in self.ledger.entries(self.KIND_RESOLUTION):
            p = e.payload
            fc = by_id.get(p["forecast_id"])
            if fc is None:
                continue
            claim = next((c for c in fc.claims if c.claim_id == p["claim_id"]), None)
            if claim is None:
                continue
            out.append(
                ResolvedClaim(
                    forecast_id=fc.forecast_id,
                    claim_id=claim.claim_id,
                    ticker=fc.ticker,
                    statement=claim.statement,
                    probability=claim.probability,
                    base_rate=claim.base_rate,
                    outcome=bool(p["outcome"]),
                    resolved_at=p["resolved_at"],
                    note=p.get("note", ""),
                )
            )
        return out

    def overdue_claims(self, as_of: Optional[date] = None) -> List[tuple[str, Claim]]:
        """
        期限を過ぎているのに未解決の主張。

        ここが溜まっているとキャリブレーションは測れない。都合の悪い予測を
        放置するのが最も一般的な自己欺瞞の形なので、明示的に取り出す。
        """
        as_of = as_of or date.today()
        resolved = {
            (e.payload["forecast_id"], e.payload["claim_id"])
            for e in self.ledger.entries(self.KIND_RESOLUTION)
        }
        out: List[tuple[str, Claim]] = []
        for fc in self.all_forecasts():
            for c in fc.claims:
                if (fc.forecast_id, c.claim_id) in resolved:
                    continue
                try:
                    if date.fromisoformat(c.deadline) < as_of:
                        out.append((fc.forecast_id, c))
                except ValueError:
                    continue
        return out

    def verify(self) -> tuple[bool, str]:
        return self.ledger.verify()
