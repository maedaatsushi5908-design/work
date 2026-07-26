import math
import random
import unittest

from ev import stats, validation


def noise(n, seed, mu=0.0, sd=0.01):
    rng = random.Random(seed)
    return [rng.gauss(mu, sd) for _ in range(n)]


class TestSharpe(unittest.TestCase):
    def test_per_period_and_annual(self):
        r = [0.001] * 100
        # 分散0なら0を返す（無限大にしない）
        self.assertEqual(validation.sharpe_ratio(r), 0.0)

    def test_annualization(self):
        r = noise(1000, 1, mu=0.001, sd=0.01)
        per = validation.sharpe_ratio(r)
        ann = validation.sharpe_ratio(r, validation.TRADING_DAYS)
        self.assertAlmostEqual(ann, per * math.sqrt(252), places=12)

    def test_too_short(self):
        self.assertEqual(validation.sharpe_ratio([0.01]), 0.0)


class TestDeflatedSharpe(unittest.TestCase):
    """
    同一のシャープが、試行数によって「有意」から「無意味」に変わることを示す。
    これが optimize.py の54通り総当たりに対する評価そのものである。
    """

    def setUp(self):
        # 日次シャープ 0.1（年率約1.59）を500日
        self.returns = noise(500, 42, mu=0.001, sd=0.01)

    def test_single_trial_is_significant(self):
        r = validation.deflated_sharpe_ratio(self.returns, n_trials=1)
        self.assertGreater(r.deflated_p, 0.95)
        self.assertTrue(r.passes)
        self.assertEqual(r.expected_max_sharpe, 0.0)

    def test_many_trials_destroys_significance(self):
        r = validation.deflated_sharpe_ratio(self.returns, n_trials=10_000)
        self.assertLess(r.deflated_p, 0.5)
        self.assertFalse(r.passes)
        self.assertGreater(r.expected_max_sharpe, r.sharpe)

    def test_expected_max_grows_with_trials(self):
        prev = -1.0
        for n in (1, 10, 54, 1000):
            r = validation.deflated_sharpe_ratio(self.returns, n_trials=n)
            self.assertGreaterEqual(r.expected_max_sharpe, prev)
            prev = r.expected_max_sharpe

    def test_deflated_p_decreases_with_trials(self):
        ps = [
            validation.deflated_sharpe_ratio(self.returns, n_trials=n).deflated_p
            for n in (1, 10, 100, 1000)
        ]
        self.assertEqual(ps, sorted(ps, reverse=True))

    def test_trial_sharpes_used_for_null_variance(self):
        # 試行間のシャープの散らばりが大きいなら、期待最大値も大きくなる
        wide = validation.deflated_sharpe_ratio(
            self.returns, n_trials=54, trial_sharpes=[0.0, 0.1, -0.1, 0.2, -0.2]
        )
        narrow = validation.deflated_sharpe_ratio(
            self.returns, n_trials=54, trial_sharpes=[0.0, 0.001, -0.001, 0.002, -0.002]
        )
        self.assertGreater(wide.expected_max_sharpe, narrow.expected_max_sharpe)

    def test_zero_edge_fails_even_at_one_trial(self):
        r = validation.deflated_sharpe_ratio(noise(500, 7, mu=0.0), n_trials=1)
        self.assertFalse(r.passes)

    def test_rejects_short_series(self):
        with self.assertRaises(ValueError):
            validation.deflated_sharpe_ratio([0.01, 0.02], n_trials=1)

    def test_rejects_bad_trials(self):
        with self.assertRaises(ValueError):
            validation.deflated_sharpe_ratio(self.returns, n_trials=0)

    def test_summary_renders(self):
        s = validation.deflated_sharpe_ratio(self.returns, n_trials=54).summary()
        self.assertIn("Deflated SR", s)


class TestMinTrackRecord(unittest.TestCase):
    def test_infinite_when_no_edge(self):
        self.assertEqual(
            validation.min_track_record_length(noise(500, 3, mu=-0.001)), math.inf
        )

    def test_finite_and_positive_with_edge(self):
        n = validation.min_track_record_length(noise(2000, 5, mu=0.001, sd=0.01))
        self.assertTrue(math.isfinite(n))
        self.assertGreater(n, 1.0)

    def test_higher_bar_needs_more_data(self):
        r = noise(2000, 5, mu=0.001, sd=0.01)
        low = validation.min_track_record_length(r, benchmark_sharpe=0.0)
        high = validation.min_track_record_length(r, benchmark_sharpe=0.05)
        self.assertGreater(high, low)


class TestPBO(unittest.TestCase):
    def test_pure_noise_gives_pbo_near_half(self):
        """
        全構成が真のエッジ0なら、学習期間で選んだ構成の検証期間順位は一様。
        つまりPBOの期待値は0.5。パラメータ選択がコイン投げであることを意味する。

        1回の推定値は標準偏差0.2程度と大きくばらつくため（分割が互いに強く
        相関しているため）、複数実現の平均で検証する。実務でも単一のPBO値を
        鵜呑みにしてはいけない。
        """
        pbos = []
        for s in range(20):
            trials = {f"cfg{i}": noise(1000, 10_000 * s + i) for i in range(10)}
            pbos.append(
                validation.probability_of_backtest_overfitting(trials, n_blocks=10).pbo
            )
        self.assertAlmostEqual(stats.mean(pbos), 0.5, delta=0.08)

    def test_genuine_edge_gives_low_pbo(self):
        trials = {f"cfg{i}": noise(2000, 200 + i) for i in range(9)}
        trials["winner"] = noise(2000, 999, mu=0.0025, sd=0.01)
        r = validation.probability_of_backtest_overfitting(trials, n_blocks=10)
        self.assertLess(r.pbo, 0.25)

    def test_degradation_positive_under_noise(self):
        trials = {f"cfg{i}": noise(1000, 300 + i) for i in range(12)}
        r = validation.probability_of_backtest_overfitting(trials, n_blocks=8)
        # ノイズでは学習期間の見かけ上の優位が検証期間で消える
        self.assertGreater(r.performance_degradation, 0.0)

    def test_rejects_odd_blocks(self):
        trials = {"a": noise(500, 1), "b": noise(500, 2)}
        with self.assertRaises(ValueError):
            validation.probability_of_backtest_overfitting(trials, n_blocks=7)

    def test_rejects_mismatched_lengths(self):
        trials = {"a": noise(500, 1), "b": noise(400, 2)}
        with self.assertRaises(ValueError):
            validation.probability_of_backtest_overfitting(trials, n_blocks=10)

    def test_rejects_single_config(self):
        with self.assertRaises(ValueError):
            validation.probability_of_backtest_overfitting({"a": noise(500, 1)})

    def test_partition_cap_respected(self):
        trials = {f"cfg{i}": noise(600, 400 + i) for i in range(4)}
        r = validation.probability_of_backtest_overfitting(
            trials, n_blocks=10, max_partitions=50
        )
        self.assertLessEqual(r.n_partitions, 50)

    def test_summary_renders(self):
        trials = {f"cfg{i}": noise(600, 500 + i) for i in range(4)}
        s = validation.probability_of_backtest_overfitting(trials, n_blocks=6).summary()
        self.assertIn("PBO", s)


class TestHaircut(unittest.TestCase):
    def test_haircut_grows_with_tests(self):
        few = validation.multiple_testing_haircut(1.5, n_obs=2500, n_tests=1)
        many = validation.multiple_testing_haircut(1.5, n_obs=2500, n_tests=500)
        self.assertLess(
            few.haircuts["bonferroni"], many.haircuts["bonferroni"]
        )

    def test_single_test_has_no_haircut(self):
        r = validation.multiple_testing_haircut(1.5, n_obs=2500, n_tests=1)
        self.assertAlmostEqual(r.haircuts["bonferroni"], 0.0, places=9)

    def test_bonferroni_dominates_holm(self):
        """Holm は Bonferroni を一律に緩める手続きなので、減額は常に Bonferroni 以下。"""
        r = validation.multiple_testing_haircut(1.0, n_obs=2500, n_tests=54)
        self.assertGreaterEqual(r.haircuts["bonferroni"], r.haircuts["holm"] - 1e-12)

    def test_holm_never_exceeds_bonferroni_on_random_pvalues(self):
        import random
        rng = random.Random(0)
        for _ in range(300):
            k = rng.randint(2, 40)
            ps = [rng.random() for _ in range(k)]
            b = validation._bonferroni(ps)
            h = validation._holm(ps)
            for bb, hh in zip(b, h):
                self.assertLessEqual(hh, bb + 1e-12)

    def test_bhy_is_stricter_than_bonferroni_at_top_rank(self):
        """
        BY(Benjamini-Yekutieli) は任意の相関を許容する代償として
        c(N)=Σ1/i の係数を持つ。最小順位の p 値に対する乗数は N·c(N) となり、
        Bonferroni の N より大きい。つまり最良戦略の p 値に対しては
        BY のほうが厳しい。これは実装の誤りではなく BY の性質である。
        """
        n_tests = 54
        r = validation.multiple_testing_haircut(1.0, n_obs=2500, n_tests=n_tests)
        c = math.fsum(1.0 / i for i in range(1, n_tests + 1))
        self.assertGreater(n_tests * c, n_tests)
        self.assertGreater(r.haircuts["bhy"], r.haircuts["bonferroni"])

    def test_haircuts_in_unit_range(self):
        r = validation.multiple_testing_haircut(0.6, n_obs=1000, n_tests=200)
        for m, h in r.haircuts.items():
            self.assertGreaterEqual(h, 0.0, m)
            self.assertLessEqual(h, 1.0, m)

    def test_rejects_bad_inputs(self):
        with self.assertRaises(ValueError):
            validation.multiple_testing_haircut(1.0, n_obs=1, n_tests=10)
        with self.assertRaises(ValueError):
            validation.multiple_testing_haircut(1.0, n_obs=100, n_tests=0)

    def test_summary_renders(self):
        s = validation.multiple_testing_haircut(1.2, n_obs=2500, n_tests=54).summary()
        self.assertIn("bonferroni", s)


class TestWalkForward(unittest.TestCase):
    def test_noise_gives_near_zero_oos(self):
        """
        真のエッジがなければ、学習期間で最良を選び続けても検証期間の
        シャープは0近傍。グリッドの最良値が期待値ではないことを示す。
        """
        trials = {f"cfg{i}": noise(2000, 600 + i) for i in range(8)}
        r = validation.walk_forward_grid(trials, n_folds=5, min_train=500)
        self.assertLess(abs(r.oos_sharpe_annual), 1.0)

    def test_genuine_edge_survives(self):
        trials = {f"cfg{i}": noise(2500, 700 + i) for i in range(6)}
        trials["winner"] = noise(2500, 1234, mu=0.002, sd=0.01)
        r = validation.walk_forward_grid(trials, n_folds=5, min_train=500)
        self.assertGreater(r.oos_sharpe_annual, 1.0)
        self.assertEqual(r.parameter_stability, 1.0)
        self.assertEqual(set(r.selected_per_fold), {"winner"})

    def test_instability_detected_under_noise(self):
        trials = {f"cfg{i}": noise(3000, 800 + i) for i in range(15)}
        r = validation.walk_forward_grid(trials, n_folds=6, min_train=600)
        # ノイズでは選ばれる構成がfoldごとに変わる
        self.assertLess(r.parameter_stability, 1.0)

    def test_oos_covers_expected_length(self):
        trials = {f"cfg{i}": noise(1200, 900 + i) for i in range(4)}
        r = validation.walk_forward_grid(trials, n_folds=4, min_train=400)
        self.assertEqual(len(r.oos_returns), 1200 - 400)

    def test_rejects_insufficient_data(self):
        trials = {"a": noise(100, 1), "b": noise(100, 2)}
        with self.assertRaises(ValueError):
            validation.walk_forward_grid(trials, n_folds=5, min_train=250)

    def test_summary_renders(self):
        trials = {f"cfg{i}": noise(1200, 950 + i) for i in range(3)}
        s = validation.walk_forward_grid(trials, n_folds=3, min_train=400).summary()
        self.assertIn("パラメータ安定性", s)


class TestSelectionInflation(unittest.TestCase):
    """
    config.py の 116/4000銘柄選抜が生む見かけのエッジを定量化する。
    真のエッジが完全に0でも、選抜だけで大きなプラスが現れる。
    """

    def test_zero_edge_produces_positive_apparent_edge(self):
        r = validation.selection_inflation(
            n_candidates=4000, n_selected=116, n_obs=2500,
            per_obs_vol=0.02, n_simulations=800,
        )
        self.assertGreater(r.expected_inflation_annual, 0.10)
        self.assertGreater(r.expected_sharpe_inflation_annual, 0.4)

    def test_inflation_grows_with_candidate_pool(self):
        small = validation.selection_inflation(
            200, 50, 2500, n_simulations=600, seed=1
        )
        large = validation.selection_inflation(
            8000, 50, 2500, n_simulations=600, seed=1
        )
        self.assertGreater(
            large.expected_inflation_annual, small.expected_inflation_annual
        )

    def test_inflation_shrinks_when_selecting_almost_everything(self):
        few = validation.selection_inflation(1000, 10, 2500, n_simulations=600, seed=2)
        most = validation.selection_inflation(1000, 990, 2500, n_simulations=600, seed=2)
        self.assertGreater(few.expected_inflation_annual, most.expected_inflation_annual)

    def test_selecting_all_gives_no_inflation(self):
        r = validation.selection_inflation(500, 500, 2500, n_simulations=400, seed=3)
        self.assertAlmostEqual(r.expected_inflation_per_obs, 0.0, delta=2e-5)

    def test_inflation_shrinks_with_longer_history(self):
        short = validation.selection_inflation(4000, 116, 250, n_simulations=600, seed=4)
        long = validation.selection_inflation(4000, 116, 5000, n_simulations=600, seed=4)
        self.assertGreater(
            short.expected_inflation_annual, long.expected_inflation_annual
        )

    def test_matches_order_statistic_theory(self):
        """
        上位 α 分位の標準正規の平均は φ(z_α)/α。理論値と一致するか。
        """
        n_cand, n_sel, n_obs, vol = 4000, 116, 2500, 0.02
        r = validation.selection_inflation(
            n_cand, n_sel, n_obs, per_obs_vol=vol, n_simulations=3000, seed=5
        )
        alpha = n_sel / n_cand
        z = stats.norm_ppf(1.0 - alpha)
        theoretical = stats.norm_pdf(z) / alpha * (vol / math.sqrt(n_obs))
        self.assertAlmostEqual(
            r.expected_inflation_per_obs, theoretical, delta=theoretical * 0.06
        )

    def test_deterministic_given_seed(self):
        a = validation.selection_inflation(1000, 100, 1000, n_simulations=300, seed=9)
        b = validation.selection_inflation(1000, 100, 1000, n_simulations=300, seed=9)
        self.assertEqual(
            a.expected_inflation_per_obs, b.expected_inflation_per_obs
        )

    def test_rejects_bad_inputs(self):
        with self.assertRaises(ValueError):
            validation.selection_inflation(100, 200, 1000)
        with self.assertRaises(ValueError):
            validation.selection_inflation(100, 10, 1)
        with self.assertRaises(ValueError):
            validation.selection_inflation(100, 10, 1000, per_obs_vol=0.0)

    def test_deflate_subtracts_inflation(self):
        adj, inf = validation.deflate_for_selection(
            observed_sharpe_annual=0.88,
            n_candidates=4000, n_selected=116, n_obs=2500,
            per_obs_vol=0.02, n_simulations=800,
        )
        self.assertLess(adj, 0.88)
        self.assertAlmostEqual(
            adj, 0.88 - inf.expected_sharpe_inflation_annual, places=12
        )

    def test_summary_renders(self):
        s = validation.selection_inflation(
            4000, 116, 2500, n_simulations=200
        ).summary()
        self.assertIn("選抜だけで生じる", s)


if __name__ == "__main__":
    unittest.main()
