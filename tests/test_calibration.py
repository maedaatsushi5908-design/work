import math
import unittest

from ev import calibration


def build(pairs):
    """[(確率, 件数, 成功数), ...] から (確率列, 結果列) を作る。"""
    ps, os_ = [], []
    for p, n, s in pairs:
        for i in range(n):
            ps.append(p)
            os_.append(i < s)
    return ps, os_


# 完全に較正された予測者: 各確率帯の実現率が予測確率に一致
CALIBRATED = [(0.1, 20, 2), (0.3, 20, 6), (0.5, 20, 10), (0.7, 20, 14), (0.9, 20, 18)]


class TestPerfectlyCalibrated(unittest.TestCase):
    """
    手計算で検算できるケース。
      ō = 50/100 = 0.5、UNC = 0.25
      Brier = 0.17、REL = 0、RES = 0.08
      恒等式 BS = REL − RES + UNC = 0 − 0.08 + 0.25 = 0.17
      BSS = 1 − 0.17/0.25 = 0.32
    """

    def setUp(self):
        ps, os_ = build(CALIBRATED)
        self.r = calibration.evaluate(ps, os_, n_bins=5, bootstrap=False)

    def test_base_rate(self):
        self.assertAlmostEqual(self.r.base_rate, 0.5, places=12)

    def test_brier(self):
        self.assertAlmostEqual(self.r.brier, 0.17, places=12)

    def test_uncertainty(self):
        self.assertAlmostEqual(self.r.uncertainty, 0.25, places=12)

    def test_reliability_is_zero(self):
        self.assertAlmostEqual(self.r.reliability, 0.0, places=12)

    def test_resolution(self):
        self.assertAlmostEqual(self.r.resolution, 0.08, places=12)

    def test_murphy_identity(self):
        self.assertAlmostEqual(
            self.r.brier_binned,
            self.r.reliability - self.r.resolution + self.r.uncertainty,
            places=12,
        )
        self.assertAlmostEqual(self.r.brier_binned, self.r.brier, places=12)

    def test_brier_skill_score(self):
        self.assertAlmostEqual(self.r.brier_skill_score, 0.32, places=12)

    def test_calibration_slope_near_one(self):
        self.assertTrue(self.r.slope_converged)
        self.assertAlmostEqual(self.r.calibration_slope, 1.0, delta=0.25)
        self.assertAlmostEqual(self.r.calibration_intercept, 0.0, delta=0.25)

    def test_not_flagged_overconfident(self):
        self.assertFalse(self.r.is_overconfident)

    def test_ece_is_zero(self):
        self.assertAlmostEqual(self.r.expected_calibration_error, 0.0, places=12)

    def test_bins_populated(self):
        self.assertEqual(len(self.r.bins), 5)
        self.assertEqual([b.n for b in self.r.bins], [20] * 5)


class TestAlwaysBaseRate(unittest.TestCase):
    """常にベースレートを答える予測者。較正は完璧だが役に立たない。"""

    def setUp(self):
        ps, os_ = build([(0.5, 100, 50)])
        self.r = calibration.evaluate(ps, os_, n_bins=5, bootstrap=False)

    def test_brier_equals_uncertainty(self):
        self.assertAlmostEqual(self.r.brier, 0.25, places=12)

    def test_resolution_is_zero(self):
        self.assertAlmostEqual(self.r.resolution, 0.0, places=12)

    def test_bss_is_zero(self):
        self.assertAlmostEqual(self.r.brier_skill_score, 0.0, places=12)

    def test_sharpness_is_zero(self):
        self.assertAlmostEqual(self.r.sharpness, 0.0, places=12)

    def test_does_not_demonstrate_skill(self):
        self.assertFalse(self.r.demonstrates_skill)
        self.assertIn("ベースレート予測に負けている", self.r.verdict())


class TestOverconfident(unittest.TestCase):
    """0.95/0.05 と言うが実際は 0.7/0.3 しか当たらない予測者。"""

    def setUp(self):
        ps, os_ = build([(0.95, 50, 35), (0.05, 50, 15)])
        self.r = calibration.evaluate(ps, os_, n_bins=5, bootstrap=False)

    def test_slope_well_below_one(self):
        self.assertTrue(self.r.slope_converged)
        self.assertLess(self.r.calibration_slope, 0.5)

    def test_flagged_overconfident(self):
        self.assertTrue(self.r.is_overconfident)

    def test_reliability_is_large(self):
        self.assertGreater(self.r.reliability, 0.05)

    def test_recalibration_pulls_toward_center(self):
        adjusted = calibration.recalibrate(0.95, self.r)
        self.assertLess(adjusted, 0.95)
        self.assertGreater(adjusted, 0.5)

    def test_recalibration_is_monotone(self):
        a = calibration.recalibrate(0.60, self.r)
        b = calibration.recalibrate(0.80, self.r)
        self.assertLess(a, b)

    def test_recalibration_improves_brier(self):
        ps, os_ = build([(0.95, 50, 35), (0.05, 50, 15)])
        adj = [calibration.recalibrate(p, self.r) for p in ps]
        improved = calibration.evaluate(adj, os_, n_bins=5, bootstrap=False)
        self.assertLess(improved.brier, self.r.brier)


class TestUnderconfident(unittest.TestCase):
    """0.6/0.4 と言うが実際は 0.9/0.1 当たる予測者。"""

    def setUp(self):
        ps, os_ = build([(0.6, 50, 45), (0.4, 50, 5)])
        self.r = calibration.evaluate(ps, os_, n_bins=5, bootstrap=False)

    def test_slope_above_one(self):
        self.assertGreater(self.r.calibration_slope, 1.0)

    def test_not_flagged_overconfident(self):
        self.assertFalse(self.r.is_overconfident)


class TestSkillGating(unittest.TestCase):
    def test_small_sample_blocks_judgement(self):
        ps, os_ = build([(0.7, 5, 4), (0.3, 5, 1)])
        r = calibration.evaluate(ps, os_, n_bins=5, bootstrap=False)
        self.assertFalse(r.has_enough_data)
        self.assertFalse(r.demonstrates_skill)
        self.assertIn("サンプル不足", r.verdict())

    def test_wide_ci_blocks_skill_claim(self):
        """
        サンプル数は足りていて BSS もプラスだが、信頼区間が0を含むケース。
        「たまたま当たっていただけ」と区別できないので、サイジングは許可しない。
        """
        ps, os_ = build([(0.7, 12, 9), (0.3, 12, 4)])
        r = calibration.evaluate(ps, os_, n_bins=5, n_resamples=2000, seed=4242)
        self.assertTrue(r.has_enough_data)
        self.assertGreater(r.brier_skill_score, 0.0)
        self.assertLess(r.bss_ci[0], 0.0)
        self.assertFalse(r.demonstrates_skill)
        self.assertIn("運と区別できない", r.verdict())

    def test_strong_record_demonstrates_skill(self):
        ps, os_ = build([(0.85, 60, 51), (0.15, 60, 9)])
        r = calibration.evaluate(ps, os_, n_bins=5, n_resamples=1000)
        self.assertTrue(r.has_enough_data)
        self.assertGreater(r.brier_skill_score, 0.0)
        self.assertIsNotNone(r.bss_ci)
        self.assertGreater(r.bss_ci[0], 0.0)
        self.assertTrue(r.demonstrates_skill)

    def test_bootstrap_ci_brackets_point_estimate(self):
        ps, os_ = build(CALIBRATED)
        r = calibration.evaluate(ps, os_, n_bins=5, n_resamples=1000)
        self.assertLessEqual(r.bss_ci[0], r.brier_skill_score)
        self.assertGreaterEqual(r.bss_ci[1], r.brier_skill_score)


class TestProperScoring(unittest.TestCase):
    def test_brier_is_minimized_by_truth(self):
        """真の確率を答えるのが最適（proper score であることの確認）。"""
        n, true_p = 4000, 0.3
        os_ = [i < int(n * true_p) for i in range(n)]
        best = calibration.evaluate([true_p] * n, os_, bootstrap=False).brier
        for wrong in (0.1, 0.2, 0.4, 0.5, 0.7):
            worse = calibration.evaluate([wrong] * n, os_, bootstrap=False).brier
            self.assertGreater(worse, best, f"p={wrong}")

    def test_log_score_is_minimized_by_truth(self):
        n, true_p = 4000, 0.3
        os_ = [i < int(n * true_p) for i in range(n)]
        best = calibration.evaluate([true_p] * n, os_, bootstrap=False).log_score
        for wrong in (0.15, 0.45, 0.6):
            worse = calibration.evaluate([wrong] * n, os_, bootstrap=False).log_score
            self.assertGreater(worse, best, f"p={wrong}")


class TestValidation(unittest.TestCase):
    def test_rejects_length_mismatch(self):
        with self.assertRaises(calibration.CalibrationError):
            calibration.evaluate([0.5, 0.5], [True])

    def test_rejects_empty(self):
        with self.assertRaises(calibration.CalibrationError):
            calibration.evaluate([], [])

    def test_rejects_probability_out_of_range(self):
        for p in (0.0, 1.0, -0.1, 1.2):
            with self.assertRaises(calibration.CalibrationError):
                calibration.evaluate([p], [True])

    def test_rejects_too_few_bins(self):
        with self.assertRaises(calibration.CalibrationError):
            calibration.evaluate([0.5], [True], n_bins=1)

    def test_all_same_outcome_gives_nan_bss(self):
        r = calibration.evaluate([0.7] * 10, [True] * 10, bootstrap=False)
        self.assertTrue(math.isnan(r.brier_skill_score))
        self.assertFalse(r.demonstrates_skill)

    def test_recalibrate_rejects_bad_probability(self):
        ps, os_ = build(CALIBRATED)
        r = calibration.evaluate(ps, os_, bootstrap=False)
        with self.assertRaises(calibration.CalibrationError):
            calibration.recalibrate(1.0, r)

    def test_recalibrate_passthrough_when_insufficient_data(self):
        ps, os_ = build([(0.7, 5, 4), (0.3, 5, 1)])
        r = calibration.evaluate(ps, os_, bootstrap=False)
        self.assertEqual(calibration.recalibrate(0.7, r), 0.7)

    def test_summary_renders(self):
        ps, os_ = build(CALIBRATED)
        s = calibration.evaluate(ps, os_, n_resamples=200).summary()
        self.assertIn("Brier Skill Score", s)
        self.assertIn("信頼性曲線", s)


if __name__ == "__main__":
    unittest.main()
