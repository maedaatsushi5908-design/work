import math
import tempfile
import unittest
from datetime import date
from pathlib import Path

from ev.forecast import (
    Claim,
    Forecast,
    ForecastJournal,
    ForecastValidationError,
    Scenario,
)

AS_OF = date(2026, 1, 15)


def good_claim(**kw):
    base = dict(
        statement="2027年3月期の営業利益が会社予想の120億円を上回る",
        probability=0.60,
        deadline="2027-05-15",
        resolution_source="決算短信の営業利益",
        claim_id="c1",
    )
    base.update(kw)
    return Claim(**base)


def good_scenarios():
    return [
        Scenario("強気", 0.25, 0.80),
        Scenario("基本", 0.45, 0.15),
        Scenario("弱気", 0.30, -0.35),
    ]


def good_forecast(**kw):
    base = dict(
        ticker="1234.T",
        thesis="カバレッジゼロの小型建設。純現金が時価総額の6割。受注残が過去最高。",
        horizon_months=24,
        claims=[good_claim()],
        scenarios=good_scenarios(),
        structural_reason="日次売買代金2000万円。機関が意味のある建玉を作れない。",
    )
    base.update(kw)
    return Forecast(**base)


class TestClaimValidation(unittest.TestCase):
    def test_valid_claim_passes(self):
        good_claim().validate(AS_OF)

    def test_rejects_certainty(self):
        for p in (0.0, 1.0, 0.99, 0.005):
            with self.assertRaises(ForecastValidationError):
                good_claim(probability=p).validate(AS_OF)

    def test_accepts_boundary_probabilities(self):
        good_claim(probability=0.02).validate(AS_OF)
        good_claim(probability=0.98).validate(AS_OF)

    def test_rejects_short_statement(self):
        with self.assertRaises(ForecastValidationError):
            good_claim(statement="上がる").validate(AS_OF)

    def test_rejects_missing_resolution_source(self):
        with self.assertRaises(ForecastValidationError):
            good_claim(resolution_source="   ").validate(AS_OF)

    def test_rejects_past_deadline(self):
        with self.assertRaises(ForecastValidationError) as cm:
            good_claim(deadline="2025-01-01").validate(AS_OF)
        self.assertIn("予測になりません", str(cm.exception))

    def test_rejects_deadline_equal_to_as_of(self):
        with self.assertRaises(ForecastValidationError):
            good_claim(deadline=AS_OF.isoformat()).validate(AS_OF)

    def test_rejects_malformed_deadline(self):
        with self.assertRaises(ForecastValidationError):
            good_claim(deadline="2027/05/15").validate(AS_OF)

    def test_base_rate_divergence_requires_rationale(self):
        with self.assertRaises(ForecastValidationError) as cm:
            good_claim(probability=0.80, base_rate=0.30).validate(AS_OF)
        self.assertIn("なぜこの銘柄は平均と違うのか", str(cm.exception))

    def test_base_rate_divergence_allowed_with_rationale(self):
        good_claim(
            probability=0.80,
            base_rate=0.30,
            divergence_rationale="受注残が既に来期計画の8割を確保済み（四半期開示で確認）",
        ).validate(AS_OF)

    def test_small_divergence_needs_no_rationale(self):
        good_claim(probability=0.42, base_rate=0.35).validate(AS_OF)

    def test_rejects_base_rate_out_of_range(self):
        with self.assertRaises(ForecastValidationError):
            good_claim(base_rate=0.0).validate(AS_OF)

    def test_roundtrip_dict(self):
        c = good_claim(base_rate=0.4)
        self.assertEqual(Claim.from_dict(c.to_dict()), c)


class TestScenarioValidation(unittest.TestCase):
    def test_rejects_below_total_loss(self):
        with self.assertRaises(ForecastValidationError):
            Scenario("破綻", 0.1, -1.5).validate()

    def test_allows_exact_total_loss(self):
        Scenario("破綻", 0.1, -1.0).validate()

    def test_rejects_bad_probability(self):
        with self.assertRaises(ForecastValidationError):
            Scenario("x", 1.5, 0.1).validate()


class TestForecastValidation(unittest.TestCase):
    def test_valid_forecast_passes(self):
        good_forecast().validate(AS_OF)

    def test_rejects_empty_claims(self):
        with self.assertRaises(ForecastValidationError) as cm:
            good_forecast(claims=[]).validate(AS_OF)
        self.assertIn("後から検証できません", str(cm.exception))

    def test_rejects_short_thesis(self):
        with self.assertRaises(ForecastValidationError):
            good_forecast(thesis="割安").validate(AS_OF)

    def test_rejects_duplicate_claim_ids(self):
        with self.assertRaises(ForecastValidationError):
            good_forecast(
                claims=[good_claim(claim_id="c1"), good_claim(claim_id="c1")]
            ).validate(AS_OF)

    def test_rejects_scenarios_not_summing_to_one(self):
        bad = [Scenario("a", 0.5, 0.2), Scenario("b", 0.3, -0.1)]
        with self.assertRaises(ForecastValidationError) as cm:
            good_forecast(scenarios=bad).validate(AS_OF)
        self.assertIn("1.0 にしてください", str(cm.exception))

    def test_rejects_bad_horizon(self):
        with self.assertRaises(ForecastValidationError):
            good_forecast(horizon_months=0).validate(AS_OF)


class TestExpectedValue(unittest.TestCase):
    def setUp(self):
        self.f = good_forecast()

    def test_arithmetic_expected_return(self):
        # 0.25*0.80 + 0.45*0.15 + 0.30*(-0.35) = 0.2 + 0.0675 - 0.105 = 0.1625
        self.assertAlmostEqual(self.f.expected_return(), 0.1625, places=12)

    def test_expected_log_return(self):
        expected = (
            0.25 * math.log(1.80) + 0.45 * math.log(1.15) + 0.30 * math.log(0.65)
        )
        self.assertAlmostEqual(self.f.expected_log_return(), expected, places=12)

    def test_log_return_below_arithmetic(self):
        self.assertLess(self.f.expected_log_return(), self.f.expected_return())

    def test_total_loss_makes_log_return_negative_infinity(self):
        f = good_forecast(
            scenarios=[Scenario("好調", 0.9, 0.5), Scenario("破綻", 0.1, -1.0)]
        )
        self.assertEqual(f.expected_log_return(), -math.inf)
        self.assertGreater(f.expected_return(), 0.0)

    def test_sigma(self):
        mu = 0.1625
        var = (
            0.25 * (0.80 - mu) ** 2 + 0.45 * (0.15 - mu) ** 2 + 0.30 * (-0.35 - mu) ** 2
        )
        self.assertAlmostEqual(self.f.sigma(), math.sqrt(var), places=12)

    def test_probability_of_loss_and_worst_case(self):
        self.assertAlmostEqual(self.f.probability_of_loss(), 0.30, places=12)
        self.assertAlmostEqual(self.f.worst_case(), -0.35, places=12)

    def test_annualized_expected_return(self):
        # 24ヶ月で +16.25% → 年率 (1.1625)^(1/2) - 1
        self.assertAlmostEqual(
            self.f.annualized_expected_return(), 1.1625 ** 0.5 - 1.0, places=12
        )

    def test_requires_scenarios(self):
        f = good_forecast(scenarios=[])
        with self.assertRaises(ForecastValidationError):
            f.expected_return()

    def test_zero_probability_scenario_ignored_in_log(self):
        f = good_forecast(
            scenarios=[
                Scenario("好調", 1.0, 0.2),
                Scenario("あり得ない破綻", 0.0, -1.0),
            ]
        )
        self.assertAlmostEqual(f.expected_log_return(), math.log(1.2), places=12)

    def test_roundtrip_dict(self):
        d = good_forecast().to_dict()
        f2 = Forecast.from_dict(d)
        self.assertAlmostEqual(f2.expected_return(), 0.1625, places=12)
        self.assertIn("derived", d)


class JournalTestCase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = str(Path(self.tmp.name) / "j.db")

    def tearDown(self):
        self.tmp.cleanup()


class TestJournal(JournalTestCase):
    def test_record_and_get(self):
        with ForecastJournal(self.path) as j:
            fid = j.record(good_forecast(), as_of=AS_OF)
            self.assertEqual(fid, "FC00001")
            fc = j.get(fid)
            self.assertEqual(fc.ticker, "1234.T")
            self.assertEqual(fc.created_at, AS_OF.isoformat())

    def test_ids_increment(self):
        with ForecastJournal(self.path) as j:
            a = j.record(good_forecast(), as_of=AS_OF)
            b = j.record(good_forecast(ticker="5678.T"), as_of=AS_OF)
            self.assertEqual((a, b), ("FC00001", "FC00002"))

    def test_claim_ids_auto_assigned(self):
        with ForecastJournal(self.path) as j:
            f = good_forecast(
                claims=[good_claim(claim_id=""), good_claim(claim_id="", probability=0.4)]
            )
            fid = j.record(f, as_of=AS_OF)
            ids = [c.claim_id for c in j.get(fid).claims]
            self.assertEqual(ids, ["c1", "c2"])

    def test_invalid_forecast_is_not_recorded(self):
        with ForecastJournal(self.path) as j:
            with self.assertRaises(ForecastValidationError):
                j.record(good_forecast(claims=[]), as_of=AS_OF)
            self.assertEqual(len(j.all_forecasts()), 0)

    def test_duplicate_id_rejected(self):
        with ForecastJournal(self.path) as j:
            j.record(good_forecast(forecast_id="X1"), as_of=AS_OF)
            with self.assertRaises(ForecastValidationError):
                j.record(good_forecast(forecast_id="X1"), as_of=AS_OF)

    def test_resolve_and_read_back(self):
        with ForecastJournal(self.path) as j:
            fid = j.record(good_forecast(), as_of=AS_OF)
            j.resolve(fid, "c1", outcome=True, note="決算短信で確認",
                      resolved_at=date(2027, 5, 16))
            rs = j.resolved_claims()
            self.assertEqual(len(rs), 1)
            self.assertTrue(rs[0].outcome)
            self.assertAlmostEqual(rs[0].probability, 0.60)
            self.assertEqual(rs[0].ticker, "1234.T")

    def test_cannot_resolve_twice(self):
        with ForecastJournal(self.path) as j:
            fid = j.record(good_forecast(), as_of=AS_OF)
            j.resolve(fid, "c1", outcome=False)
            with self.assertRaises(ForecastValidationError) as cm:
                j.resolve(fid, "c1", outcome=True)
            self.assertIn("書き換えはできません", str(cm.exception))

    def test_resolve_unknown_forecast_or_claim(self):
        with ForecastJournal(self.path) as j:
            fid = j.record(good_forecast(), as_of=AS_OF)
            with self.assertRaises(ForecastValidationError):
                j.resolve("NOPE", "c1", True)
            with self.assertRaises(ForecastValidationError):
                j.resolve(fid, "c99", True)

    def test_overdue_claims(self):
        with ForecastJournal(self.path) as j:
            fid = j.record(good_forecast(), as_of=AS_OF)
            self.assertEqual(j.overdue_claims(as_of=date(2027, 1, 1)), [])
            overdue = j.overdue_claims(as_of=date(2027, 6, 1))
            self.assertEqual(len(overdue), 1)
            self.assertEqual(overdue[0][0], fid)

    def test_resolved_claims_excluded_from_overdue(self):
        with ForecastJournal(self.path) as j:
            fid = j.record(good_forecast(), as_of=AS_OF)
            j.resolve(fid, "c1", True)
            self.assertEqual(j.overdue_claims(as_of=date(2027, 6, 1)), [])

    def test_record_trade_requires_known_forecast(self):
        with ForecastJournal(self.path) as j:
            fid = j.record(good_forecast(), as_of=AS_OF)
            j.record_trade(fid, "buy", 1500.0, 100, 0.03)
            with self.assertRaises(ForecastValidationError):
                j.record_trade("NOPE", "buy", 1500.0, 100, 0.03)
            with self.assertRaises(ValueError):
                j.record_trade(fid, "hold", 1500.0, 100, 0.03)

    def test_ledger_stays_verifiable(self):
        with ForecastJournal(self.path) as j:
            fid = j.record(good_forecast(), as_of=AS_OF)
            j.resolve(fid, "c1", True)
            j.record_trade(fid, "buy", 1000.0, 100, 0.05)
            ok, _ = j.verify()
            self.assertTrue(ok)

    def test_persists_across_reopen(self):
        with ForecastJournal(self.path) as j:
            fid = j.record(good_forecast(), as_of=AS_OF)
        with ForecastJournal(self.path) as j:
            self.assertIsNotNone(j.get(fid))
            self.assertEqual(len(j.all_forecasts()), 1)


if __name__ == "__main__":
    unittest.main()
