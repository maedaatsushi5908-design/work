import math
import random
import unittest

from ev import performance


def noise(n, seed, mu=0.0, sd=0.01):
    rng = random.Random(seed)
    return [rng.gauss(mu, sd) for _ in range(n)]


def active_with_exact_stats(n, seed, alpha_annual, te_annual, periods=252):
    """
    標本平均と標本標準偏差が指定値にぴったり一致する超過リターン列を作る。

    乱数の引き方に成績を左右させないため、中心化してから目標値に合わせる。
    """
    raw = noise(n, seed, 0.0, 1.0)
    m = sum(raw) / n
    centered = [x - m for x in raw]
    sd = (sum(x * x for x in centered) / (n - 1)) ** 0.5
    scale = (te_annual / math.sqrt(periods)) / sd
    per_period_alpha = alpha_annual / periods
    return [per_period_alpha + x * scale for x in centered]


class TestYearsToSignificance(unittest.TestCase):
    def test_canonical_hundred_years(self):
        """α=3%, TE=15%, t=2 → (2×15/3)² = 100年。"""
        self.assertAlmostEqual(
            performance.years_to_significance(0.03, 0.15, 2.0), 100.0, places=9
        )

    def test_concentrated_skilled_case(self):
        """α=5%, TE=10% でも16年かかる。"""
        self.assertAlmostEqual(
            performance.years_to_significance(0.05, 0.10, 2.0), 16.0, places=9
        )

    def test_t3_needs_more_than_t2(self):
        t2 = performance.years_to_significance(0.03, 0.15, 2.0)
        t3 = performance.years_to_significance(0.03, 0.15, 3.0)
        self.assertAlmostEqual(t3 / t2, 2.25, places=9)

    def test_zero_or_negative_alpha_never_arrives(self):
        self.assertEqual(performance.years_to_significance(0.0, 0.15), math.inf)
        self.assertEqual(performance.years_to_significance(-0.02, 0.15), math.inf)

    def test_zero_tracking_error_is_instant(self):
        self.assertEqual(performance.years_to_significance(0.03, 0.0), 0.0)

    def test_scales_quadratically_with_te(self):
        a = performance.years_to_significance(0.03, 0.10)
        b = performance.years_to_significance(0.03, 0.20)
        self.assertAlmostEqual(b / a, 4.0, places=9)

    def test_rejects_bad_inputs(self):
        with self.assertRaises(performance.PerformanceError):
            performance.years_to_significance(0.03, -0.1)
        with self.assertRaises(performance.PerformanceError):
            performance.years_to_significance(0.03, 0.15, t_target=0.0)


class TestEvaluateAlpha(unittest.TestCase):
    def test_no_skill_gives_insignificant_result(self):
        bench = noise(2520, 1, mu=0.0003, sd=0.011)
        port = [b + n for b, n in zip(bench, noise(2520, 2, mu=0.0, sd=0.005))]
        r = performance.evaluate_alpha(port, bench)
        self.assertGreater(r.p_value, 0.05)
        self.assertFalse(r.is_significant)

    def test_identical_series_has_zero_alpha(self):
        bench = noise(500, 3, mu=0.0004)
        r = performance.evaluate_alpha(bench, bench, use_regression=False)
        self.assertAlmostEqual(r.alpha_annual, 0.0, places=12)
        self.assertAlmostEqual(r.tracking_error_annual, 0.0, places=12)

    def test_beta_recovered_by_regression(self):
        bench = noise(2000, 4, mu=0.0003, sd=0.01)
        port = [1.5 * b for b in bench]
        r = performance.evaluate_alpha(port, bench)
        self.assertAlmostEqual(r.beta, 1.5, places=6)
        self.assertAlmostEqual(r.r_squared, 1.0, places=6)

    def test_years_observed(self):
        bench = noise(504, 5)
        port = noise(504, 6)
        r = performance.evaluate_alpha(port, bench)
        self.assertAlmostEqual(r.years_observed, 2.0, places=6)

    def test_strong_alpha_is_significant(self):
        bench = noise(2520, 7, mu=0.0003, sd=0.01)
        port = [b + 0.0008 for b in bench]
        r = performance.evaluate_alpha(port, bench)
        self.assertLess(r.p_value, 0.05)
        self.assertTrue(r.is_significant)

    def test_information_ratio_sign_follows_alpha(self):
        bench = noise(1000, 8, mu=0.0003)
        worse = [b - 0.0005 for b in bench]
        r = performance.evaluate_alpha(worse, bench, use_regression=False)
        self.assertLess(r.information_ratio, 0.0)

    def test_rejects_mismatched_lengths(self):
        with self.assertRaises(performance.PerformanceError):
            performance.evaluate_alpha([0.01] * 10, [0.01] * 9)

    def test_rejects_short_series(self):
        with self.assertRaises(performance.PerformanceError):
            performance.evaluate_alpha([0.01, 0.02], [0.01, 0.02])

    def test_recovers_exact_alpha_and_tracking_error(self):
        bench = noise(2520, 20, mu=0.0003, sd=0.01)
        active = active_with_exact_stats(2520, 21, alpha_annual=0.03, te_annual=0.15)
        port = [b + a for b, a in zip(bench, active)]
        r = performance.evaluate_alpha(port, bench, use_regression=False)
        self.assertAlmostEqual(r.alpha_annual, 0.03, places=10)
        self.assertAlmostEqual(r.tracking_error_annual, 0.15, places=10)
        self.assertAlmostEqual(r.information_ratio, 0.2, places=10)

    def test_summary_mentions_calibration_when_horizon_is_absurd(self):
        """
        年率アルファ0.5%、TE16% → 必要年数 (2×0.16/0.005)² = 4096年。
        リターンでの自己評価を諦めるべき典型例。
        """
        bench = noise(756, 9, mu=0.0003, sd=0.01)
        active = active_with_exact_stats(756, 10, alpha_annual=0.005, te_annual=0.16)
        port = [b + a for b, a in zip(bench, active)]
        r = performance.evaluate_alpha(port, bench, use_regression=False)
        self.assertAlmostEqual(r.years_for_t2, 4096.0, delta=1.0)
        self.assertGreater(r.years_for_t2, 40)
        self.assertIn("キャリブレーション", r.summary())

    def test_realistic_three_percent_alpha_needs_a_century(self):
        bench = noise(2520, 22, mu=0.0003, sd=0.01)
        active = active_with_exact_stats(2520, 23, alpha_annual=0.03, te_annual=0.15)
        port = [b + a for b, a in zip(bench, active)]
        r = performance.evaluate_alpha(port, bench, use_regression=False)
        self.assertAlmostEqual(r.years_for_t2, 100.0, delta=0.1)
        self.assertAlmostEqual(r.years_observed, 10.0, places=6)
        # 10年観測しても t 値は 0.63 程度にしか届かない
        self.assertAlmostEqual(r.t_stat, 0.632, delta=0.01)
        self.assertGreater(r.p_value, 0.05)

    def test_summary_renders(self):
        bench = noise(500, 11, mu=0.0003)
        port = noise(500, 12, mu=0.0004)
        self.assertIn("年率アルファ", performance.evaluate_alpha(port, bench).summary())


class TestVolatilityDrag(unittest.TestCase):
    def test_drag_formula(self):
        r = performance.volatility_drag(0.10, 0.20)
        self.assertAlmostEqual(r.drag, 0.02, places=12)
        self.assertAlmostEqual(r.geometric_approx, 0.08, places=12)

    def test_high_vol_flips_sign(self):
        r = performance.volatility_drag(0.05, 0.40)
        self.assertGreater(r.arithmetic_mean, 0.0)
        self.assertLess(r.geometric_approx, 0.0)
        self.assertIn("複利成長率はマイナス", r.summary())

    def test_zero_vol_has_no_drag(self):
        r = performance.volatility_drag(0.07, 0.0)
        self.assertEqual(r.drag, 0.0)
        self.assertEqual(r.geometric_approx, 0.07)

    def test_rejects_negative_vol(self):
        with self.assertRaises(performance.PerformanceError):
            performance.volatility_drag(0.05, -0.1)


class TestTwoPointBet(unittest.TestCase):
    """会話で提示した数値例が実装と一致するかを固定する。"""

    def setUp(self):
        self.b = performance.two_point_bet(0.50, -0.40, 0.5)

    def test_arithmetic_ev_is_five_percent(self):
        self.assertAlmostEqual(self.b.arithmetic_ev, 0.05, places=12)

    def test_full_bet_growth_is_negative_five_point_one(self):
        self.assertAlmostEqual(self.b.full_bet_growth, math.sqrt(0.9) - 1.0, places=12)
        self.assertAlmostEqual(self.b.full_bet_growth, -0.0513167, places=6)

    def test_kelly_is_one_quarter(self):
        self.assertAlmostEqual(self.b.optimal_fraction, 0.25, places=9)

    def test_kelly_growth_is_positive_point_six_two(self):
        self.assertAlmostEqual(self.b.optimal_growth, 0.0062306, places=6)
        self.assertGreater(self.b.optimal_growth, 0.0)

    def test_sign_flip_between_full_and_kelly(self):
        self.assertLess(self.b.full_bet_growth, 0.0)
        self.assertGreater(self.b.optimal_growth, 0.0)

    def test_negative_ev_bet_has_zero_kelly(self):
        b = performance.two_point_bet(0.20, -0.30, 0.5)
        self.assertLess(b.arithmetic_ev, 0.0)
        self.assertEqual(b.optimal_fraction, 0.0)

    def test_rejects_bad_inputs(self):
        with self.assertRaises(performance.PerformanceError):
            performance.two_point_bet(0.5, -0.4, 0.0)
        with self.assertRaises(performance.PerformanceError):
            performance.two_point_bet(0.5, -1.2, 0.5)

    def test_summary_renders(self):
        self.assertIn("ケリー比率", self.b.summary())


class TestCostDrag(unittest.TestCase):
    def test_one_percent_over_thirty_years(self):
        gross, net, lost = performance.cost_drag_over_time(0.01, 30, gross_return=0.06)
        self.assertAlmostEqual(gross, 1.06 ** 30, places=12)
        self.assertAlmostEqual(net, 1.05 ** 30, places=12)
        self.assertGreater(lost, 0.20)
        self.assertLess(lost, 0.30)

    def test_zero_cost_loses_nothing(self):
        _, _, lost = performance.cost_drag_over_time(0.0, 20)
        self.assertAlmostEqual(lost, 0.0, places=12)

    def test_longer_horizon_loses_more(self):
        _, _, short = performance.cost_drag_over_time(0.01, 10)
        _, _, long = performance.cost_drag_over_time(0.01, 40)
        self.assertGreater(long, short)

    def test_rejects_bad_inputs(self):
        with self.assertRaises(performance.PerformanceError):
            performance.cost_drag_over_time(0.01, 0)
        with self.assertRaises(performance.PerformanceError):
            performance.cost_drag_over_time(-0.01, 10)


if __name__ == "__main__":
    unittest.main()
