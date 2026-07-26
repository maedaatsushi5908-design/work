import unittest

from ev import base_rates
from ev.base_rates import (
    BaseRateError,
    Observation,
    build_base_rates,
    check_against_base_rate,
    observations_from_rows,
)


def obs(bucket, n, successes):
    return [Observation(bucket, i < successes) for i in range(n)]


class TestBuildBaseRates(unittest.TestCase):
    def setUp(self):
        self.table = build_base_rates(
            obs("小型建設", 200, 60) + obs("小型IT", 200, 100) + obs("大型", 200, 80)
        )

    def test_global_rate(self):
        self.assertAlmostEqual(self.table.global_rate, 240 / 600, places=12)
        self.assertEqual(self.table.total_n, 600)

    def test_raw_rates(self):
        self.assertAlmostEqual(self.table.get("小型建設").raw_rate, 0.30, places=12)
        self.assertAlmostEqual(self.table.get("小型IT").raw_rate, 0.50, places=12)

    def test_shrunk_rates_pulled_toward_global(self):
        low = self.table.get("小型建設")
        self.assertGreater(low.shrunk_rate, low.raw_rate)
        high = self.table.get("小型IT")
        self.assertLess(high.shrunk_rate, high.raw_rate)

    def test_shrinkage_formula(self):
        r = self.table.get("小型建設")
        k = self.table.prior_strength
        expected = (r.successes + k * self.table.global_rate) / (r.n + k)
        self.assertAlmostEqual(r.shrunk_rate, expected, places=12)

    def test_confidence_interval_brackets_raw(self):
        r = self.table.get("小型建設")
        self.assertLess(r.ci_low, r.raw_rate)
        self.assertGreater(r.ci_high, r.raw_rate)

    def test_unknown_bucket_falls_back_to_global(self):
        r = self.table.get("知らないバケット")
        self.assertEqual(r.n, 0)
        self.assertAlmostEqual(r.shrunk_rate, self.table.global_rate, places=12)
        self.assertEqual(r.shrinkage_weight, 0.0)

    def test_summary_renders(self):
        s = self.table.summary()
        self.assertIn("全体ベースレート", s)
        self.assertIn("小型建設", s)

    def test_rejects_empty(self):
        with self.assertRaises(BaseRateError):
            build_base_rates([])


class TestShrinkageStrength(unittest.TestCase):
    def test_small_bucket_shrinks_more(self):
        table = build_base_rates(
            obs("大", 500, 250) + obs("小", 5, 5) + obs("中", 100, 40)
        )
        small = table.get("小")
        big = table.get("大")
        self.assertLess(small.shrinkage_weight, big.shrinkage_weight)
        # 5戦5勝でも縮小後は1.0にならない
        self.assertEqual(small.raw_rate, 1.0)
        self.assertLess(small.shrunk_rate, 1.0)

    def test_explicit_prior_strength_is_used(self):
        data = obs("a", 10, 8) + obs("b", 10, 2)
        weak = build_base_rates(data, prior_strength=1.0)
        strong = build_base_rates(data, prior_strength=1000.0)
        self.assertGreater(
            abs(weak.get("a").shrunk_rate - weak.global_rate),
            abs(strong.get("a").shrunk_rate - strong.global_rate),
        )

    def test_pure_noise_buckets_shrink_hard(self):
        """
        バケット間の差が二項ノイズだけなら、事前の強さは大きくなる（強く縮小）。
        """
        import random
        rng = random.Random(0)
        data = []
        for b in range(20):
            n = 100
            s = sum(1 for _ in range(n) if rng.random() < 0.4)
            data.extend(obs(f"b{b}", n, s))
        table = build_base_rates(data)
        self.assertGreater(table.prior_strength, 50.0)

    def test_genuinely_different_buckets_shrink_less(self):
        data = []
        for b, rate in enumerate((0.1, 0.3, 0.5, 0.7, 0.9)):
            data.extend(obs(f"b{b}", 200, int(200 * rate)))
        table = build_base_rates(data)
        self.assertLess(table.prior_strength, 50.0)

    def test_single_bucket_uses_default(self):
        table = build_base_rates(obs("only", 50, 25))
        self.assertEqual(table.prior_strength, 10.0)

    def test_reliability_flag(self):
        table = build_base_rates(obs("big", 2000, 800) + obs("tiny", 6, 3))
        self.assertTrue(table.get("big").is_reliable)
        self.assertFalse(table.get("tiny").is_reliable)


class TestTupleBuckets(unittest.TestCase):
    def test_composite_bucket_keys(self):
        data = obs(("小型", "建設"), 100, 30) + obs(("小型", "IT"), 100, 55)
        table = build_base_rates(data)
        self.assertAlmostEqual(table.get(("小型", "建設")).raw_rate, 0.30, places=12)
        self.assertAlmostEqual(table.get(("小型", "IT")).raw_rate, 0.55, places=12)


class TestCheckAgainstBaseRate(unittest.TestCase):
    def test_matching_probability_needs_no_rationale(self):
        c = check_against_base_rate(0.30, 0.30)
        self.assertAlmostEqual(c.logit_gap, 0.0, places=12)
        self.assertAlmostEqual(c.implied_odds_multiple, 1.0, places=12)
        self.assertFalse(c.requires_rationale)

    def test_large_divergence_requires_rationale(self):
        c = check_against_base_rate(0.80, 0.30)
        self.assertGreater(c.logit_gap, 1.0)
        self.assertTrue(c.requires_rationale)
        self.assertIn("なぜこの銘柄は平均と違うのか", c.message)

    def test_odds_multiple_is_exponential_of_gap(self):
        import math
        c = check_against_base_rate(0.70, 0.40)
        self.assertAlmostEqual(c.implied_odds_multiple, math.exp(c.logit_gap), places=12)

    def test_bearish_divergence_also_flagged(self):
        c = check_against_base_rate(0.10, 0.50)
        self.assertLess(c.logit_gap, -1.0)
        self.assertTrue(c.requires_rationale)
        self.assertIn("弱気", c.message)

    def test_limit_is_configurable(self):
        self.assertFalse(
            check_against_base_rate(0.80, 0.30, divergence_limit=5.0).requires_rationale
        )

    def test_rejects_out_of_range(self):
        for p, b in ((0.0, 0.3), (1.0, 0.3), (0.3, 0.0), (0.3, 1.0)):
            with self.assertRaises(BaseRateError):
                check_against_base_rate(p, b)


class TestObservationsFromRows(unittest.TestCase):
    def test_builds_from_dict_rows(self):
        rows = [
            {"size": "小型", "sector": "建設", "fwd": "0.15"},
            {"size": "小型", "sector": "建設", "fwd": "-0.05"},
            {"size": "大型", "sector": "IT", "fwd": "0.30"},
        ]
        got = observations_from_rows(
            rows,
            bucket_fn=lambda r: (r["size"], r["sector"]),
            outcome_fn=lambda r: float(r["fwd"]) > 0.0,
        )
        self.assertEqual(len(got), 3)
        self.assertEqual(got[0].bucket, ("小型", "建設"))
        self.assertTrue(got[0].outcome)
        self.assertFalse(got[1].outcome)

    def test_skips_malformed_rows(self):
        rows = [{"a": "1"}, {"nope": "x"}, {"a": "bad"}]
        got = observations_from_rows(
            rows, bucket_fn=lambda r: "b", outcome_fn=lambda r: float(r["a"]) > 0
        )
        self.assertEqual(len(got), 1)


if __name__ == "__main__":
    unittest.main()
