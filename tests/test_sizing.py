import math
import unittest

from ev import sizing
from ev.calibration import evaluate
from ev.forecast import Forecast, Scenario


def calibrated_report(demonstrates=True):
    """サイジングを通す/通さない CalibrationReport を作る。"""
    if demonstrates:
        pairs = [(0.85, 60, 51), (0.15, 60, 9)]
    else:
        pairs = [(0.5, 100, 50)]
    ps, os_ = [], []
    for p, n, s in pairs:
        for i in range(n):
            ps.append(p)
            os_.append(i < s)
    return evaluate(ps, os_, n_bins=5, n_resamples=1000)


def make_forecast(scenarios, horizon_months=12):
    return Forecast(
        ticker="1234.T",
        thesis="カバレッジゼロの小型株。純現金が時価総額の6割で受注残は最高水準。",
        horizon_months=horizon_months,
        scenarios=scenarios,
    )


# 検算用の教科書的な賭け: +50% / -40% を 50/50
COIN = [Scenario("勝ち", 0.5, 0.50), Scenario("負け", 0.5, -0.40)]

# キャップが効く水準まで大きいエッジ: 生ケリー110%、既定処理後4.51%
STRONG = [Scenario("成功", 0.6, 1.00), Scenario("失敗", 0.4, -0.40)]


class TestExactKelly(unittest.TestCase):
    def test_textbook_case_is_exactly_quarter(self):
        """
        0.25/(1+0.5f) = 0.2/(1-0.4f) を解くと f = 0.25。
        数値解が解析解に一致するか。
        """
        f, notes = sizing.exact_kelly([0.50, -0.40], [0.5, 0.5])
        self.assertAlmostEqual(f, 0.25, places=9)
        self.assertEqual(notes, [])

    def test_log_growth_signs_around_kelly(self):
        rs, ps = [0.50, -0.40], [0.5, 0.5]
        # 全額投入は成長率マイナス、ケリー比率はプラス
        self.assertLess(sizing.log_growth(1.0, rs, ps), 0.0)
        self.assertGreater(sizing.log_growth(0.25, rs, ps), 0.0)

    def test_full_bet_growth_matches_closed_form(self):
        g = sizing.log_growth(1.0, [0.50, -0.40], [0.5, 0.5])
        self.assertAlmostEqual(math.exp(g) - 1.0, math.sqrt(0.9) - 1.0, places=12)

    def test_kelly_growth_matches_closed_form(self):
        g = sizing.log_growth(0.25, [0.50, -0.40], [0.5, 0.5])
        expected = 0.5 * math.log(1.125) + 0.5 * math.log(0.9)
        self.assertAlmostEqual(g, expected, places=12)
        self.assertAlmostEqual(math.exp(g) - 1.0, 0.0062306, places=6)

    def test_kelly_maximizes_log_growth(self):
        rs, ps = [0.50, -0.40], [0.5, 0.5]
        f, _ = sizing.exact_kelly(rs, ps)
        best = sizing.log_growth(f, rs, ps)
        for other in (0.05, 0.15, 0.20, 0.30, 0.40, 0.60):
            self.assertLessEqual(sizing.log_growth(other, rs, ps), best + 1e-12)

    def test_negative_ev_gives_zero(self):
        f, notes = sizing.exact_kelly([0.20, -0.30], [0.5, 0.5])
        self.assertEqual(f, 0.0)
        self.assertIn("見送り", notes[0])

    def test_no_downside_is_flagged(self):
        f, notes = sizing.exact_kelly([0.20, 0.05], [0.5, 0.5])
        self.assertEqual(f, sizing.MAX_KELLY_SEARCH)
        self.assertIn("下方シナリオがありません", notes[0])

    def test_total_loss_bounds_kelly_below_one(self):
        f, _ = sizing.exact_kelly([1.00, -1.00], [0.6, 0.4])
        self.assertLess(f, 1.0)
        self.assertGreater(f, 0.0)
        self.assertEqual(sizing.log_growth(1.0, [1.00, -1.00], [0.6, 0.4]), -math.inf)

    def test_bigger_edge_gives_bigger_kelly(self):
        small, _ = sizing.exact_kelly([0.50, -0.40], [0.50, 0.50])
        big, _ = sizing.exact_kelly([0.50, -0.40], [0.65, 0.35])
        self.assertGreater(big, small)

    def test_fatter_tail_gives_smaller_kelly(self):
        """
        同じ算術期待値(+5%)でも下方の裾が厚いほどケリーは小さくなる。
        正規近似 μ/σ² では捉えきれない差を、厳密解は捉える。
        """
        thin_scen, fat_scen = [0.30, -0.20], [1.05, -0.95]
        self.assertAlmostEqual(
            0.5 * thin_scen[0] + 0.5 * thin_scen[1],
            0.5 * fat_scen[0] + 0.5 * fat_scen[1],
            places=12,
        )
        thin, _ = sizing.exact_kelly(thin_scen, [0.5, 0.5])
        fat, _ = sizing.exact_kelly(fat_scen, [0.5, 0.5])
        self.assertAlmostEqual(thin, 0.05 / 0.06, places=9)
        self.assertAlmostEqual(fat, 0.05 / 0.9975, places=9)
        self.assertLess(fat, thin / 10.0)

    def test_rejects_bad_probabilities(self):
        with self.assertRaises(sizing.SizingError):
            sizing.exact_kelly([0.1, -0.1], [0.5, 0.4])
        with self.assertRaises(sizing.SizingError):
            sizing.exact_kelly([0.1], [0.5, 0.5])
        with self.assertRaises(sizing.SizingError):
            sizing.exact_kelly([], [])


class TestGaussianKelly(unittest.TestCase):
    def test_basic_formula(self):
        self.assertAlmostEqual(sizing.gaussian_kelly(0.05, 0.20), 0.05 / 0.04, places=12)

    def test_edge_uncertainty_reduces_size(self):
        certain = sizing.gaussian_kelly(0.05, 0.20, edge_variance=0.0)
        uncertain = sizing.gaussian_kelly(0.05, 0.20, edge_variance=0.02)
        self.assertLess(uncertain, certain)

    def test_rejects_zero_variance(self):
        with self.assertRaises(sizing.SizingError):
            sizing.gaussian_kelly(0.05, 0.0)


class TestDrawdownProbability(unittest.TestCase):
    def test_full_kelly_halves_capital_half_the_time(self):
        """P(a) = a^(2/c − 1)。c=1 なら P(0.5) = 0.5。"""
        self.assertAlmostEqual(sizing.drawdown_probability(1.0, 0.5), 0.5, places=12)

    def test_quarter_kelly_is_far_safer(self):
        self.assertAlmostEqual(
            sizing.drawdown_probability(0.25, 0.5), 0.5 ** 7, places=12
        )
        self.assertAlmostEqual(sizing.drawdown_probability(0.25, 0.5), 0.0078125, places=9)

    def test_half_kelly(self):
        self.assertAlmostEqual(
            sizing.drawdown_probability(0.5, 0.5), 0.5 ** 3, places=12
        )

    def test_monotone_in_fraction(self):
        vals = [sizing.drawdown_probability(c, 0.5) for c in (0.1, 0.25, 0.5, 1.0)]
        self.assertEqual(vals, sorted(vals))

    def test_deeper_drawdown_is_less_likely(self):
        self.assertGreater(
            sizing.drawdown_probability(0.25, 0.8), sizing.drawdown_probability(0.25, 0.2)
        )

    def test_rejects_bad_inputs(self):
        with self.assertRaises(sizing.SizingError):
            sizing.drawdown_probability(0.0, 0.5)
        with self.assertRaises(sizing.SizingError):
            sizing.drawdown_probability(3.0, 0.5)
        with self.assertRaises(sizing.SizingError):
            sizing.drawdown_probability(0.25, 1.0)


class TestShrinkage(unittest.TestCase):
    def test_no_uncertainty_keeps_estimate(self):
        s = sizing.shrink_edge(0.10, stderr=0.0, prior_stdev=0.03)
        self.assertAlmostEqual(s.shrinkage_weight, 1.0, places=12)
        self.assertAlmostEqual(s.posterior_mean, 0.10, places=12)

    def test_huge_uncertainty_collapses_to_prior(self):
        s = sizing.shrink_edge(0.50, stderr=100.0, prior_stdev=0.03)
        self.assertLess(s.shrinkage_weight, 1e-5)
        self.assertAlmostEqual(s.posterior_mean, 0.0, places=6)

    def test_weight_formula(self):
        tau, se = 0.03, 0.05
        s = sizing.shrink_edge(0.20, stderr=se, prior_stdev=tau)
        expected_w = tau ** 2 / (tau ** 2 + se ** 2)
        self.assertAlmostEqual(s.shrinkage_weight, expected_w, places=12)
        self.assertAlmostEqual(s.posterior_mean, expected_w * 0.20, places=12)

    def test_posterior_variance_formula(self):
        tau, se = 0.03, 0.05
        s = sizing.shrink_edge(0.20, stderr=se, prior_stdev=tau)
        w = tau ** 2 / (tau ** 2 + se ** 2)
        self.assertAlmostEqual(s.posterior_stdev, math.sqrt(w * se ** 2), places=12)

    def test_posterior_is_between_estimate_and_prior(self):
        s = sizing.shrink_edge(0.20, stderr=0.05, prior_mean=0.0, prior_stdev=0.03)
        self.assertGreater(s.posterior_mean, 0.0)
        self.assertLess(s.posterior_mean, 0.20)

    def test_default_shrinkage_weight_is_one_fifth(self):
        """既定の情報レシオ0.25と推定誤差比0.5から w=0.2 が決まる。"""
        expected = (
            sizing.DEFAULT_PRIOR_INFORMATION_RATIO ** 2
            / (
                sizing.DEFAULT_PRIOR_INFORMATION_RATIO ** 2
                + sizing.DEFAULT_ESTIMATE_ERROR_RATIO ** 2
            )
        )
        self.assertAlmostEqual(expected, 0.2, places=12)
        r = sizing.size_position(
            make_forecast(COIN), calibration=calibrated_report(True)
        )
        self.assertAlmostEqual(r.shrunk_edge.shrinkage_weight, 0.2, places=12)


class TestPriorUnitConsistency(unittest.TestCase):
    """
    回帰テスト: 事前分布を絶対値で固定すると、保有期間やボラティリティが
    変わったときに単位が噛み合わず、エッジが不当に潰れる。
    τ = 情報レシオ × σ とすることで、σ の大小に関わらず縮小の重みは一定になる。
    """

    def test_shrinkage_weight_invariant_to_volatility(self):
        low_vol = [Scenario("勝ち", 0.5, 0.10), Scenario("負け", 0.5, -0.08)]
        high_vol = [Scenario("勝ち", 0.5, 1.00), Scenario("負け", 0.5, -0.80)]
        rep = calibrated_report(True)
        a = sizing.size_position(make_forecast(low_vol), calibration=rep)
        b = sizing.size_position(make_forecast(high_vol), calibration=rep)
        self.assertGreater(b.shrunk_edge.raw_stderr, a.shrunk_edge.raw_stderr * 5)
        self.assertAlmostEqual(
            a.shrunk_edge.shrinkage_weight, b.shrunk_edge.shrinkage_weight, places=12
        )

    def test_posterior_edge_stays_a_fixed_share_of_claim(self):
        rep = calibrated_report(True)
        for scen in ([Scenario("w", 0.5, 0.10), Scenario("l", 0.5, -0.08)], COIN, STRONG):
            f = make_forecast(scen)
            r = sizing.size_position(f, calibration=rep)
            self.assertAlmostEqual(
                r.shrunk_edge.posterior_mean, 0.2 * f.expected_return(), places=12
            )

    def test_explicit_stderr_overrides_default(self):
        rep = calibrated_report(True)
        loose = sizing.size_position(make_forecast(COIN), calibration=rep)
        tight = sizing.size_position(
            make_forecast(COIN), calibration=rep, edge_stderr=0.05
        )
        self.assertGreater(
            tight.shrunk_edge.shrinkage_weight, loose.shrunk_edge.shrinkage_weight
        )
        self.assertGreater(tight.weight, loose.weight)

    def test_rejects_zero_sigma(self):
        flat = [Scenario("a", 0.5, 0.10), Scenario("b", 0.5, 0.10)]
        with self.assertRaises(sizing.SizingError):
            sizing.size_position(
                make_forecast(flat), calibration=calibrated_report(True)
            )

    def test_rejects_bad_prior_ratio(self):
        with self.assertRaises(sizing.SizingError):
            sizing.size_position(
                make_forecast(COIN),
                calibration=calibrated_report(True),
                prior_information_ratio=0.0,
            )

    def test_rejects_bad_inputs(self):
        with self.assertRaises(sizing.SizingError):
            sizing.shrink_edge(0.1, stderr=-1.0)
        with self.assertRaises(sizing.SizingError):
            sizing.shrink_edge(0.1, stderr=0.1, prior_stdev=0.0)


class TestCalibrationGate(unittest.TestCase):
    def test_no_record_gives_zero(self):
        gate, reason = sizing.calibration_gate(None)
        self.assertEqual(gate, 0.0)
        self.assertIn("実績なし", reason)

    def test_no_skill_gives_zero(self):
        gate, _ = sizing.calibration_gate(calibrated_report(demonstrates=False))
        self.assertEqual(gate, 0.0)

    def test_demonstrated_skill_opens_gate(self):
        gate, _ = sizing.calibration_gate(calibrated_report(demonstrates=True))
        self.assertGreater(gate, 0.0)
        self.assertLessEqual(gate, 1.0)

    def test_gate_is_capped_at_one(self):
        r = calibrated_report(demonstrates=True)
        gate, _ = sizing.calibration_gate(r, bss_reference=0.001)
        self.assertEqual(gate, 1.0)


class TestLiquidityCap(unittest.TestCase):
    def test_formula(self):
        c = sizing.SizingConstraints(
            portfolio_value=10_000_000, adv_participation=0.10, exit_days=5.0
        )
        # 売買代金2000万円 × 10% × 5日 = 1000万円 = 100%
        self.assertAlmostEqual(sizing.liquidity_cap(20_000_000, c), 1.0, places=12)
        # 売買代金200万円 → 100万円 = 10%
        self.assertAlmostEqual(sizing.liquidity_cap(2_000_000, c), 0.10, places=12)

    def test_zero_volume_gives_zero(self):
        self.assertEqual(sizing.liquidity_cap(0.0, sizing.SizingConstraints()), 0.0)

    def test_capped_at_one(self):
        self.assertEqual(sizing.liquidity_cap(1e12, sizing.SizingConstraints()), 1.0)

    def test_rejects_negative(self):
        with self.assertRaises(sizing.SizingError):
            sizing.liquidity_cap(-1.0, sizing.SizingConstraints())


class TestSizePosition(unittest.TestCase):
    def test_no_calibration_means_no_position(self):
        """本ライブラリの中核。実績がなければサテライトは0。"""
        r = sizing.size_position(make_forecast(COIN), calibration=None)
        self.assertTrue(r.is_zero)
        self.assertGreater(r.kelly_raw, 0.0)
        self.assertEqual(r.gate, 0.0)

    def test_pipeline_is_monotonically_decreasing(self):
        r = sizing.size_position(
            make_forecast(COIN), calibration=calibrated_report(True)
        )
        self.assertGreaterEqual(r.kelly_raw, r.kelly_shrunk)
        self.assertGreaterEqual(r.kelly_shrunk, r.after_fraction)
        self.assertGreaterEqual(r.after_fraction, r.after_gate)
        self.assertGreaterEqual(r.after_gate, r.weight)

    def test_shrinkage_reduces_kelly(self):
        r = sizing.size_position(
            make_forecast(COIN), calibration=calibrated_report(True)
        )
        self.assertLess(r.kelly_shrunk, r.kelly_raw)
        self.assertLess(r.shrunk_edge.posterior_mean, r.shrunk_edge.raw_edge)

    def test_max_weight_can_bind(self):
        cons = sizing.SizingConstraints(max_weight=0.01)
        r = sizing.size_position(
            make_forecast(STRONG), constraints=cons, calibration=calibrated_report(True)
        )
        self.assertAlmostEqual(r.weight, 0.01, places=12)
        self.assertEqual(r.binding_constraint, "1銘柄上限")

    def test_strong_edge_sizes_larger_than_weak(self):
        rep = calibrated_report(True)
        weak = sizing.size_position(make_forecast(COIN), calibration=rep)
        strong = sizing.size_position(make_forecast(STRONG), calibration=rep)
        self.assertGreater(strong.weight, weak.weight)
        self.assertAlmostEqual(strong.weight, 0.0451, delta=0.0005)
        self.assertAlmostEqual(weak.weight, 0.0124, delta=0.0005)

    def test_worst_case_cap_can_bind(self):
        cons = sizing.SizingConstraints(max_weight=1.0, max_worst_case_loss=0.02)
        scen = [Scenario("勝ち", 0.6, 0.5), Scenario("負け", 0.4, -0.5)]
        r = sizing.size_position(
            make_forecast(scen), constraints=cons, calibration=calibrated_report(True)
        )
        self.assertAlmostEqual(r.caps["最悪シナリオ損失上限"], 0.04, places=12)

    def test_liquidity_cap_can_bind(self):
        cons = sizing.SizingConstraints(
            max_weight=1.0, portfolio_value=10_000_000,
            adv_participation=0.10, exit_days=5.0,
        )
        # 売買代金20万円 → 出口から逆算した上限は 20万×10%×5日/1000万 = 1%
        r = sizing.size_position(
            make_forecast(STRONG), constraints=cons,
            calibration=calibrated_report(True), adv_value=200_000,
        )
        self.assertEqual(r.binding_constraint, "流動性(出口)")
        self.assertAlmostEqual(r.weight, 0.01, places=12)

    def test_ample_liquidity_does_not_bind(self):
        cons = sizing.SizingConstraints(max_weight=1.0, portfolio_value=10_000_000)
        r = sizing.size_position(
            make_forecast(STRONG), constraints=cons,
            calibration=calibrated_report(True), adv_value=5_000_000_000,
        )
        self.assertNotEqual(r.binding_constraint, "流動性(出口)")

    def test_total_loss_scenario_is_warned(self):
        scen = [Scenario("成功", 0.7, 1.5), Scenario("破綻", 0.3, -1.0)]
        r = sizing.size_position(
            make_forecast(scen), calibration=calibrated_report(True)
        )
        self.assertTrue(any("全損" in n for n in r.notes))

    def test_no_downside_scenario_is_warned(self):
        scen = [Scenario("好調", 0.5, 0.5), Scenario("横ばい", 0.5, 0.02)]
        r = sizing.size_position(
            make_forecast(scen), calibration=calibrated_report(True)
        )
        self.assertTrue(any("下方シナリオ" in n for n in r.notes))

    def test_negative_ev_gives_zero_weight(self):
        scen = [Scenario("勝ち", 0.4, 0.3), Scenario("負け", 0.6, -0.3)]
        r = sizing.size_position(
            make_forecast(scen), calibration=calibrated_report(True)
        )
        self.assertTrue(r.is_zero)

    def test_log_growth_positive_at_final_weight(self):
        r = sizing.size_position(
            make_forecast(COIN), calibration=calibrated_report(True)
        )
        self.assertGreater(r.weight, 0.0)
        self.assertGreater(r.log_growth_at_weight, 0.0)

    def test_drawdown_probability_is_small(self):
        r = sizing.size_position(
            make_forecast(COIN), calibration=calibrated_report(True)
        )
        self.assertLess(r.drawdown_50_prob, 0.10)

    def test_requires_scenarios(self):
        f = Forecast(ticker="X", thesis="a" * 30, horizon_months=12)
        with self.assertRaises(sizing.SizingError):
            sizing.size_position(f, calibration=calibrated_report(True))

    def test_rejects_bad_constraints(self):
        with self.assertRaises(sizing.SizingError):
            sizing.size_position(
                make_forecast(COIN),
                constraints=sizing.SizingConstraints(max_weight=0.0),
                calibration=calibrated_report(True),
            )

    def test_summary_renders(self):
        s = sizing.size_position(
            make_forecast(COIN), calibration=calibrated_report(True)
        ).summary()
        self.assertIn("サイジング内訳", s)
        self.assertIn("← 効いている", s)


class TestPortfolioAllocation(unittest.TestCase):
    def _sized(self, n, weight=0.05):
        out = {}
        rep = calibrated_report(True)
        for i in range(n):
            r = sizing.size_position(
                make_forecast(STRONG),
                constraints=sizing.SizingConstraints(max_weight=weight),
                calibration=rep,
            )
            out[f"T{i}"] = r
        return out

    def test_core_fills_remainder(self):
        alloc = sizing.allocate_portfolio(self._sized(2), satellite_budget=0.30)
        self.assertAlmostEqual(
            alloc.core_weight + alloc.total_satellite, 1.0, places=12
        )

    def test_budget_is_enforced(self):
        alloc = sizing.allocate_portfolio(self._sized(10), satellite_budget=0.20)
        self.assertLessEqual(alloc.total_satellite, 0.20 + 1e-12)
        self.assertLess(alloc.scaled_by, 1.0)

    def test_under_budget_is_not_scaled(self):
        alloc = sizing.allocate_portfolio(self._sized(2), satellite_budget=0.50)
        self.assertEqual(alloc.scaled_by, 1.0)

    def test_sector_cap_is_enforced(self):
        sized = self._sized(4)
        sectors = {t: "建設" for t in sized}
        alloc = sizing.allocate_portfolio(
            sized, satellite_budget=1.0, sector_of=sectors, max_sector_weight=0.08
        )
        self.assertLessEqual(alloc.total_satellite, 0.08 + 1e-12)
        self.assertTrue(any("建設" in n for n in alloc.notes))

    def test_empty_gives_all_core(self):
        alloc = sizing.allocate_portfolio({})
        self.assertEqual(alloc.core_weight, 1.0)
        self.assertEqual(alloc.satellite_weights, {})
        self.assertTrue(any("全額コア" in n for n in alloc.notes))

    def test_zero_weight_positions_excluded(self):
        sized = {"A": sizing.size_position(make_forecast(COIN), calibration=None)}
        alloc = sizing.allocate_portfolio(sized)
        self.assertEqual(alloc.satellite_weights, {})

    def test_rejects_bad_budget(self):
        with self.assertRaises(sizing.SizingError):
            sizing.allocate_portfolio({}, satellite_budget=1.5)

    def test_summary_renders(self):
        s = sizing.allocate_portfolio(self._sized(3), satellite_budget=0.2).summary()
        self.assertIn("コア", s)


if __name__ == "__main__":
    unittest.main()
