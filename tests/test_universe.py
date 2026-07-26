import unittest

from ev import universe
from ev.universe import (
    InvestorProfile,
    ScreenCriteria,
    Security,
    largest_participating_fund,
    screen_security,
    screen_universe,
)


def micro_cap(**kw):
    """
    機関が入れず、個人には十分な小型株。

    売買代金1000万円/日は次の帯域に収まる:
      参加可能最大ファンド = 1000万 × 15% × 10日 / 0.5% = 30億円（上限50億以下）
      自分の建玉余裕度     = 1000万 × 10% × 5日 = 500万 ÷ 目標50万 = 10倍
    """
    base = dict(
        ticker="1234.T",
        name="小型建設",
        market_cap=40e8,
        adv_value=10_000_000,
        price=1500.0,
        analyst_count=0,
        sector="建設",
        net_cash=24e8,
        pbr=0.6,
    )
    base.update(kw)
    return Security(**base)


def large_cap(**kw):
    base = dict(
        ticker="7203.T",
        name="大型",
        market_cap=400_000e8,
        adv_value=50_000_000_000,
        price=3000.0,
        analyst_count=22,
        sector="自動車",
    )
    base.update(kw)
    return Security(**base)


PROFILE = InvestorProfile(
    portfolio_value=10_000_000,
    target_position_weight=0.05,
    adv_participation=0.10,
    exit_days=5.0,
)


class TestLargestParticipatingFund(unittest.TestCase):
    def test_formula(self):
        # 売買代金2000万 × 15% × 10日 = 3000万。0.5%で意味を持つ運用額は60億
        self.assertAlmostEqual(largest_participating_fund(20_000_000), 60e8, places=2)

    def test_scales_linearly_with_volume(self):
        a = largest_participating_fund(10_000_000)
        b = largest_participating_fund(20_000_000)
        self.assertAlmostEqual(b / a, 2.0, places=9)

    def test_zero_volume_excludes_everyone(self):
        self.assertEqual(largest_participating_fund(0.0), 0.0)

    def test_rejects_bad_inputs(self):
        with self.assertRaises(ValueError):
            largest_participating_fund(-1.0)
        with self.assertRaises(ValueError):
            largest_participating_fund(1e7, min_position_pct=0.0)


class TestSecurity(unittest.TestCase):
    def test_net_cash_ratio(self):
        self.assertAlmostEqual(micro_cap().net_cash_ratio, 0.6, places=12)

    def test_net_cash_ratio_none_when_unknown(self):
        self.assertIsNone(micro_cap(net_cash=None).net_cash_ratio)

    def test_turnover_ratio(self):
        s = micro_cap(market_cap=40e8, adv_value=10_000_000)
        self.assertAlmostEqual(s.turnover_ratio, 10_000_000 * 250 / 40e8, places=12)

    def test_validate_rejects_bad_values(self):
        with self.assertRaises(ValueError):
            Security(ticker="").validate()
        with self.assertRaises(ValueError):
            micro_cap(market_cap=-1).validate()
        with self.assertRaises(ValueError):
            micro_cap(analyst_count=-1).validate()
        with self.assertRaises(ValueError):
            micro_cap(free_float=1.5).validate()


class TestScreenMicroCap(unittest.TestCase):
    def setUp(self):
        self.r = screen_security(micro_cap(), PROFILE)

    def test_passes(self):
        self.assertTrue(self.r.passed, self.r.reasons)

    def test_capacity_is_sufficient(self):
        # 1000万 × 10% × 5日 = 500万、目標建玉は50万 → 10倍
        self.assertAlmostEqual(self.r.my_capacity_value, 5_000_000, places=2)
        self.assertAlmostEqual(self.r.capacity_ratio, 10.0, places=9)

    def test_institutions_are_excluded(self):
        # 1000万 × 15% × 10日 / 0.5% = 30億円
        self.assertAlmostEqual(self.r.largest_fund, 30e8, places=2)
        self.assertLess(self.r.largest_fund, ScreenCriteria().max_participating_fund)

    def test_edge_flags_detected(self):
        joined = " ".join(self.r.edge_flags)
        self.assertIn("カバレッジゼロ", joined)
        self.assertIn("純現金", joined)
        self.assertIn("PBR", joined)

    def test_structural_score_is_high(self):
        self.assertGreater(self.r.structural_score, 60.0)

    def test_summary_renders(self):
        self.assertIn("通過", self.r.summary())


class TestScreenLargeCap(unittest.TestCase):
    def setUp(self):
        self.r = screen_security(large_cap(), PROFILE)

    def test_rejected(self):
        self.assertFalse(self.r.passed)

    def test_reasons_cite_size_and_coverage(self):
        joined = " ".join(self.r.reasons)
        self.assertIn("時価総額", joined)
        self.assertIn("カバレッジ", joined)

    def test_structural_score_is_low(self):
        self.assertLess(self.r.structural_score, 20.0)

    def test_institutions_can_participate(self):
        self.assertGreater(self.r.largest_fund, 1000e8)


class TestScreenRejections(unittest.TestCase):
    def test_too_illiquid_for_own_position(self):
        # 売買代金10万円 → 建てられるのは5万円、目標は50万円
        r = screen_security(micro_cap(adv_value=100_000), PROFILE)
        self.assertFalse(r.passed)
        self.assertTrue(any("出口が確保できない" in x for x in r.reasons))

    def test_too_much_coverage(self):
        r = screen_security(micro_cap(analyst_count=8), PROFILE)
        self.assertFalse(r.passed)
        self.assertTrue(any("カバレッジ" in x for x in r.reasons))

    def test_market_cap_floor(self):
        r = screen_security(micro_cap(market_cap=0.5e8), PROFILE)
        self.assertFalse(r.passed)

    def test_zero_market_cap(self):
        r = screen_security(micro_cap(market_cap=0.0), PROFILE)
        self.assertFalse(r.passed)
        self.assertTrue(any("時価総額が不明" in x for x in r.reasons))

    def test_price_floor(self):
        r = screen_security(micro_cap(price=50.0), PROFILE)
        self.assertFalse(r.passed)
        self.assertTrue(any("株価" in x for x in r.reasons))

    def test_criteria_can_be_relaxed(self):
        loose = ScreenCriteria(max_analyst_count=20, max_participating_fund=1e15,
                               max_market_cap=1e15)
        r = screen_security(micro_cap(analyst_count=8), PROFILE, loose)
        self.assertTrue(r.passed, r.reasons)

    def test_bigger_portfolio_loses_access(self):
        """
        資金が増えると小型株に入れなくなる。これは機関が入れない理由と
        同じ構造であり、個人の優位が資金量に依存することを意味する。
        """
        rich = InvestorProfile(portfolio_value=1_000_000_000,
                               target_position_weight=0.05)
        r = screen_security(micro_cap(), rich)
        self.assertFalse(r.passed)
        self.assertTrue(any("出口が確保できない" in x for x in r.reasons))


class TestStructuralScore(unittest.TestCase):
    def test_coverage_lowers_score(self):
        a = screen_security(micro_cap(analyst_count=0), PROFILE).structural_score
        b = screen_security(micro_cap(analyst_count=2), PROFILE).structural_score
        self.assertGreater(a, b)

    def test_catalysts_raise_score(self):
        plain = screen_security(
            micro_cap(net_cash=None, pbr=None), PROFILE
        ).structural_score
        rich = screen_security(
            micro_cap(net_cash=40e8, pbr=0.4, is_parent_subsidiary=True,
                      forced_seller_flag=True),
            PROFILE,
        ).structural_score
        self.assertGreater(rich, plain)

    def test_score_bounded(self):
        for sec in (micro_cap(), large_cap(),
                    micro_cap(net_cash=100e8, pbr=0.1, is_parent_subsidiary=True,
                              forced_seller_flag=True, analyst_count=0)):
            s = screen_security(sec, PROFILE).structural_score
            self.assertGreaterEqual(s, 0.0)
            self.assertLessEqual(s, 100.0)

    def test_less_liquid_scores_higher_when_still_investable(self):
        a = screen_security(micro_cap(adv_value=5_000_000), PROFILE).structural_score
        b = screen_security(micro_cap(adv_value=100_000_000), PROFILE).structural_score
        self.assertGreater(a, b)


class TestScreenUniverse(unittest.TestCase):
    def test_sorts_passers_first_then_by_score(self):
        secs = [
            large_cap(),
            micro_cap(ticker="A", analyst_count=1, net_cash=None, pbr=None),
            micro_cap(ticker="B", analyst_count=0, net_cash=30e8, pbr=0.5),
        ]
        results = screen_universe(secs, PROFILE)
        self.assertEqual([r.security.ticker for r in results[:2]], ["B", "A"])
        self.assertFalse(results[-1].passed)

    def test_summary_counts(self):
        secs = [large_cap(), micro_cap()]
        s = universe.universe_summary(screen_universe(secs, PROFILE))
        self.assertIn("1 / 2 銘柄が通過", s)

    def test_empty_universe(self):
        self.assertEqual(screen_universe([], PROFILE), [])
        self.assertIn("0 / 0", universe.universe_summary([]))


class TestInvestorProfile(unittest.TestCase):
    def test_target_position_value(self):
        self.assertAlmostEqual(PROFILE.target_position_value, 500_000, places=2)

    def test_validate_rejects_bad_values(self):
        for kw in (
            dict(portfolio_value=0),
            dict(target_position_weight=0.0),
            dict(target_position_weight=1.5),
            dict(adv_participation=0.0),
            dict(exit_days=0.0),
        ):
            with self.assertRaises(ValueError):
                InvestorProfile(**kw).validate()


class TestScreenCriteria(unittest.TestCase):
    def test_validate_rejects_inverted_bounds(self):
        with self.assertRaises(ValueError):
            ScreenCriteria(min_market_cap=100e8, max_market_cap=10e8).validate()
        with self.assertRaises(ValueError):
            ScreenCriteria(max_analyst_count=-1).validate()
        with self.assertRaises(ValueError):
            ScreenCriteria(min_capacity_ratio=0.0).validate()


if __name__ == "__main__":
    unittest.main()
