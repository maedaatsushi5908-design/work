import math
import unittest

from ev import stats


class TestNormal(unittest.TestCase):
    def test_norm_cdf_known_values(self):
        self.assertAlmostEqual(stats.norm_cdf(0.0), 0.5, places=12)
        self.assertAlmostEqual(stats.norm_cdf(1.0), 0.8413447460685429, places=12)
        self.assertAlmostEqual(stats.norm_cdf(-1.96), 0.024997895148220435, places=12)
        self.assertAlmostEqual(stats.norm_cdf(2.5758293035489004), 0.995, places=10)

    def test_norm_ppf_known_values(self):
        self.assertAlmostEqual(stats.norm_ppf(0.5), 0.0, places=12)
        self.assertAlmostEqual(stats.norm_ppf(0.975), 1.959963984540054, places=9)
        self.assertAlmostEqual(stats.norm_ppf(0.95), 1.6448536269514722, places=9)
        self.assertAlmostEqual(stats.norm_ppf(0.005), -2.5758293035489004, places=9)

    def test_norm_ppf_roundtrip(self):
        for p in (1e-8, 0.001, 0.02, 0.2, 0.5, 0.8, 0.98, 0.999, 1 - 1e-8):
            self.assertAlmostEqual(stats.norm_cdf(stats.norm_ppf(p)), p, places=10)

    def test_norm_ppf_rejects_out_of_range(self):
        with self.assertRaises(ValueError):
            stats.norm_ppf(-0.1)
        self.assertEqual(stats.norm_ppf(0.0), -math.inf)
        self.assertEqual(stats.norm_ppf(1.0), math.inf)


class TestIncompleteBeta(unittest.TestCase):
    def test_symmetry(self):
        # I_x(a,b) = 1 - I_{1-x}(b,a)
        for a, b, x in ((2.0, 3.0, 0.4), (0.5, 0.5, 0.7), (10.0, 1.0, 0.2)):
            self.assertAlmostEqual(
                stats.betainc(a, b, x), 1.0 - stats.betainc(b, a, 1.0 - x), places=12
            )

    def test_known_values(self):
        # I_x(1,1) = x
        self.assertAlmostEqual(stats.betainc(1.0, 1.0, 0.37), 0.37, places=12)
        # I_x(1,2) = 1-(1-x)^2
        self.assertAlmostEqual(stats.betainc(1.0, 2.0, 0.5), 0.75, places=12)
        # I_x(2,1) = x^2
        self.assertAlmostEqual(stats.betainc(2.0, 1.0, 0.5), 0.25, places=12)

    def test_boundaries(self):
        self.assertEqual(stats.betainc(2.0, 2.0, 0.0), 0.0)
        self.assertEqual(stats.betainc(2.0, 2.0, 1.0), 1.0)


class TestStudentT(unittest.TestCase):
    def test_two_sided_p_known(self):
        # t=2.0, df=10 -> p = 0.0733880348 (下の数値積分テストで独立に検算)
        self.assertAlmostEqual(stats.t_two_sided_p(2.0, 10), 0.0733880348, places=9)
        # t=1.96, df=1e6 は正規分布に一致
        self.assertAlmostEqual(
            stats.t_two_sided_p(1.959963984540054, 1_000_000), 0.05, places=4
        )

    def test_two_sided_p_matches_quadrature(self):
        """不完全ベータ経由の値を、t分布pdfの台形積分と突き合わせる。"""

        def t_pdf(x: float, df: float) -> float:
            return (
                math.gamma((df + 1) / 2)
                / (math.sqrt(df * math.pi) * math.gamma(df / 2))
                * (1 + x * x / df) ** (-(df + 1) / 2)
            )

        def quadrature(t: float, df: float, n: int = 200_000) -> float:
            a, b = -t, t
            h = (b - a) / n
            s = 0.5 * (t_pdf(a, df) + t_pdf(b, df))
            for i in range(1, n):
                s += t_pdf(a + i * h, df)
            return 1.0 - s * h

        for t, df in ((2.0, 10), (2.5, 20), (1.0, 5), (3.0, 100)):
            self.assertAlmostEqual(
                stats.t_two_sided_p(t, df), quadrature(t, df), places=10,
                msg=f"t={t}, df={df}",
            )

    def test_symmetry(self):
        self.assertAlmostEqual(
            stats.t_two_sided_p(2.5, 20), stats.t_two_sided_p(-2.5, 20), places=14
        )

    def test_sf_complements(self):
        self.assertAlmostEqual(stats.t_sf(0.0, 5), 0.5, places=12)
        self.assertAlmostEqual(stats.t_sf(1.5, 8) + stats.t_sf(-1.5, 8), 1.0, places=12)

    def test_norm_two_sided_p(self):
        self.assertAlmostEqual(stats.norm_two_sided_p(1.959963984540054), 0.05, places=12)


class TestMoments(unittest.TestCase):
    def test_mean_variance(self):
        xs = [1.0, 2.0, 3.0, 4.0]
        self.assertAlmostEqual(stats.mean(xs), 2.5)
        self.assertAlmostEqual(stats.variance(xs, ddof=1), 5.0 / 3.0)
        self.assertAlmostEqual(stats.variance(xs, ddof=0), 1.25)

    def test_skew_kurtosis_of_symmetric(self):
        xs = [-2.0, -1.0, 0.0, 1.0, 2.0]
        self.assertAlmostEqual(stats.skewness(xs), 0.0, places=12)
        self.assertGreater(stats.kurtosis(xs), 0.0)

    def test_kurtosis_normal_is_about_three(self):
        import random
        rng = random.Random(1)
        xs = [rng.gauss(0, 1) for _ in range(20000)]
        self.assertAlmostEqual(stats.kurtosis(xs), 3.0, delta=0.2)

    def test_degenerate_inputs(self):
        self.assertEqual(stats.skewness([1.0, 1.0, 1.0]), 0.0)
        self.assertEqual(stats.kurtosis([1.0]), 3.0)
        with self.assertRaises(ValueError):
            stats.mean([])
        with self.assertRaises(ValueError):
            stats.variance([1.0], ddof=1)


class TestQuantile(unittest.TestCase):
    def test_linear_interpolation(self):
        xs = [1.0, 2.0, 3.0, 4.0]
        self.assertAlmostEqual(stats.quantile(xs, 0.0), 1.0)
        self.assertAlmostEqual(stats.quantile(xs, 1.0), 4.0)
        self.assertAlmostEqual(stats.quantile(xs, 0.5), 2.5)
        self.assertAlmostEqual(stats.quantile(xs, 0.25), 1.75)

    def test_rejects_bad_q(self):
        with self.assertRaises(ValueError):
            stats.quantile([1.0], 1.5)


class TestWilson(unittest.TestCase):
    def test_contains_point_estimate(self):
        lo, hi = stats.wilson_interval(30, 100)
        self.assertLess(lo, 0.30)
        self.assertGreater(hi, 0.30)

    def test_narrows_with_n(self):
        w_small = stats.wilson_interval(3, 10)
        w_big = stats.wilson_interval(300, 1000)
        self.assertGreater(w_small[1] - w_small[0], w_big[1] - w_big[0])

    def test_extremes_stay_in_unit_interval(self):
        lo, hi = stats.wilson_interval(0, 5)
        self.assertGreaterEqual(lo, 0.0)
        self.assertLessEqual(hi, 1.0)
        lo, hi = stats.wilson_interval(5, 5)
        self.assertLessEqual(hi, 1.0)

    def test_zero_n_returns_full_range(self):
        self.assertEqual(stats.wilson_interval(0, 0), (0.0, 1.0))

    def test_rejects_inconsistent(self):
        with self.assertRaises(ValueError):
            stats.wilson_interval(11, 10)


class TestExpectedMax(unittest.TestCase):
    def test_single_trial_is_zero(self):
        self.assertEqual(stats.expected_max_of_normals(1), 0.0)

    def test_increases_with_trials(self):
        vals = [stats.expected_max_of_normals(n) for n in (2, 10, 100, 1000, 10000)]
        self.assertEqual(vals, sorted(vals))

    def test_magnitude_is_plausible(self):
        # 100回引いた最大値はおよそ 2.5σ 前後
        self.assertAlmostEqual(stats.expected_max_of_normals(100), 2.5, delta=0.3)
        # 10000回ならおよそ 3.85σ
        self.assertAlmostEqual(stats.expected_max_of_normals(10000), 3.85, delta=0.25)

    def test_monte_carlo_agreement(self):
        import random
        rng = random.Random(7)
        n_trials = 50
        sims = [max(rng.gauss(0, 1) for _ in range(n_trials)) for _ in range(4000)]
        self.assertAlmostEqual(
            stats.expected_max_of_normals(n_trials), stats.mean(sims), delta=0.1
        )


class TestBootstrap(unittest.TestCase):
    def test_brackets_the_mean(self):
        xs = [float(i) for i in range(100)]
        point, lo, hi = stats.bootstrap_ci(xs, stats.mean, n_resamples=500, seed=3)
        self.assertAlmostEqual(point, 49.5)
        self.assertLess(lo, 49.5)
        self.assertGreater(hi, 49.5)

    def test_deterministic_given_seed(self):
        xs = [float(i) for i in range(50)]
        a = stats.bootstrap_ci(xs, stats.mean, n_resamples=200, seed=11)
        b = stats.bootstrap_ci(xs, stats.mean, n_resamples=200, seed=11)
        self.assertEqual(a, b)

    def test_single_observation(self):
        self.assertEqual(stats.bootstrap_ci([5.0], stats.mean), (5.0, 5.0, 5.0))


class TestHelpers(unittest.TestCase):
    def test_logit_expit_roundtrip(self):
        for p in (0.01, 0.3, 0.5, 0.77, 0.99):
            self.assertAlmostEqual(stats.expit(stats.logit(p)), p, places=12)

    def test_expit_extremes_do_not_overflow(self):
        self.assertAlmostEqual(stats.expit(-800.0), 0.0, places=12)
        self.assertAlmostEqual(stats.expit(800.0), 1.0, places=12)

    def test_clamp(self):
        self.assertEqual(stats.clamp(5.0, 0.0, 1.0), 1.0)
        self.assertEqual(stats.clamp(-5.0, 0.0, 1.0), 0.0)
        self.assertEqual(stats.clamp(0.5, 0.0, 1.0), 0.5)

    def test_zscores(self):
        z = stats.zscores([1.0, 2.0, 3.0])
        self.assertAlmostEqual(stats.mean(z), 0.0, places=12)
        self.assertEqual(stats.zscores([7.0, 7.0]), [0.0, 0.0])

    def test_ranks_with_ties(self):
        self.assertEqual(stats.ranks([10.0, 20.0, 30.0]), [1.0, 2.0, 3.0])
        self.assertEqual(stats.ranks([5.0, 5.0, 9.0]), [1.5, 1.5, 3.0])

    def test_winsorize_clips_tails(self):
        xs = [float(i) for i in range(100)] + [10000.0]
        w = stats.winsorize(xs, limit=0.05)
        self.assertLess(max(w), 10000.0)


if __name__ == "__main__":
    unittest.main()
