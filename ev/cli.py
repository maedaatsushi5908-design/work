"""
ev のコマンドラインインタフェース

    python3 -m ev.cli <サブコマンド> [オプション]

サブコマンド:
    diagnose     既存バックテストの過剰最適化・選抜バイアスを診断する
    bet          2値の賭けの期待値・ケリー比率・破産リスクを計算する
    years        アルファが有意になるまでの年数を計算する
    screen       CSVから構造的ユニバース選別を行う
    forecast     予測をJSONから台帳に登録する
    resolve      予測の結果を台帳に記録する
    journal      台帳の内容と未解決の予測を表示する
    calibration  台帳の実績からキャリブレーションを評価する
    size         登録済み予測のポジションサイズを算出する
    demo         全工程を合成データで一気に実行する

依存パッケージはない（標準ライブラリのみ）。
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import random
import sys
from datetime import date
from pathlib import Path
from typing import List, Optional, Sequence

from . import base_rates, calibration, performance, sizing, stats, universe, validation
from .forecast import Claim, Forecast, ForecastJournal, Scenario


def _rule(title: str) -> str:
    return f"\n{'═' * 68}\n  {title}\n{'═' * 68}"


def _section(title: str) -> str:
    return f"\n── {title} " + "─" * max(0, 64 - len(title))


# ────────────────────────────────────────────────────────────────
# diagnose
# ────────────────────────────────────────────────────────────────

def cmd_diagnose(args: argparse.Namespace) -> int:
    print(_rule("バックテスト診断"))
    print(
        "申告された成績が、試行回数と銘柄選抜だけで説明できるかを検定する。\n"
        "この2つを申告しない成績は、成績ではない。"
    )

    print(_section("1. パラメータ探索の試行回数による割引"))
    hc = validation.multiple_testing_haircut(
        args.sharpe, n_obs=args.n_obs, n_tests=args.trials
    )
    print(hc.summary())

    print(_section("2. 銘柄選抜バイアス"))
    if args.candidates and args.selected:
        adj, inf = validation.deflate_for_selection(
            observed_sharpe_annual=args.sharpe,
            n_candidates=args.candidates,
            n_selected=args.selected,
            n_obs=args.n_obs,
            per_obs_vol=args.vol,
            n_simulations=args.sims,
        )
        print(inf.summary())
        print(f"\n選抜バイアス控除後の年率シャープ: {adj:+.3f}")
        if adj <= 0.0:
            print(
                "→ 実測シャープは選抜バイアスの期待値以下。真のエッジの証拠はない。"
            )
    else:
        print("--candidates と --selected を指定すると選抜バイアスを評価します。")

    print(_section("3. 総合（逐次控除）"))
    print(
        "2つの控除は独立ではないため、順に適用する。\n"
        "  段階A: 選抜バイアスを引き、戦略の真のシャープの点推定を得る\n"
        "  段階B: その点推定が試行数を考慮しても有意かを問う\n"
    )

    stage_a = args.sharpe
    if args.candidates and args.selected:
        stage_a = args.sharpe - inf.expected_sharpe_inflation_annual
        print(f"段階A: {args.sharpe:+.3f} − {inf.expected_sharpe_inflation_annual:.3f}"
              f" = {stage_a:+.3f}")
    else:
        print(f"段階A: 選抜情報が未指定のため {stage_a:+.3f} のまま")

    if stage_a <= 0.0:
        print("段階B: 段階Aで既に0以下。試行数を考慮する前に真のエッジの証拠がない。")
        final = stage_a
    else:
        hc2 = validation.multiple_testing_haircut(
            stage_a, n_obs=args.n_obs, n_tests=args.trials
        )
        final = hc2.adjusted_sharpe_annual["holm"]
        print(
            f"段階B: {stage_a:+.3f} に Holm 補正 "
            f"(減額 {hc2.haircuts['holm'] * 100:.1f}%) → {final:+.3f}"
            f"   [BHYはより厳しく {hc2.adjusted_sharpe_annual['bhy']:+.3f}]"
        )

    print(f"\n残る年率シャープ: {final:+.3f}")
    if final <= 0.2:
        print(
            "\n判定: 申告成績はほぼ全て探索と選抜で説明できる。\n"
            "      対策は次の3つ。\n"
            "        (1) ユニバースを成績で選ばない（構造的特徴だけで絞る: ev screen）\n"
            "        (2) パラメータは学習期間だけで決め、検証期間に触らない\n"
            "        (3) 試した構成数を必ず記録し、DSRで割り引く"
        )
    else:
        print("\n判定: 控除後もシャープが残っている。ウォークフォワードで再確認すること。")
    return 0


# ────────────────────────────────────────────────────────────────
# bet
# ────────────────────────────────────────────────────────────────

def cmd_bet(args: argparse.Namespace) -> int:
    b = performance.two_point_bet(args.win, args.loss, args.p)
    print(_rule("2値の賭け"))
    print(b.summary())

    print(_section("賭け金による成長率の変化"))
    print(
        f"  {'投入比率':>9}{'ケリー倍率':>12}{'対数成長率':>13}"
        f"{'1回あたり':>11}{'半減確率':>11}"
    )
    rows = sorted({0.05, 0.10, 0.25, 0.50, 0.75, 1.00} | {
        round(b.optimal_fraction, 4)
    } if b.optimal_fraction > 0 else {0.05, 0.10, 0.25, 0.50, 0.75, 1.00})
    for f in rows:
        if f <= 0.0:
            continue
        g = sizing.log_growth(f, [args.win, args.loss], [args.p, 1.0 - args.p])
        per = math.exp(g) - 1.0 if g > -math.inf else -1.0
        if b.optimal_fraction > 0:
            ratio = f / b.optimal_fraction
            c = stats.clamp(ratio, 1e-9, 2.0)
            dd = sizing.drawdown_probability(c, 0.5)
            ratio_s = f"{ratio:>11.2f}"
            dd_s = f"{dd * 100:>10.2f}%"
        else:
            ratio_s = f"{'—':>11}"
            dd_s = f"{'—':>11}"
        mark = "  ← フルケリー" if (
            b.optimal_fraction > 0 and abs(f - b.optimal_fraction) < 1e-9
        ) else ""
        print(
            f"  {f * 100:>8.1f}%{ratio_s}{g:>+13.5f}{per * 100:>+10.2f}%{dd_s}{mark}"
        )

    print(
        "\n算術期待値が同じでも、投入比率だけで成長率の符号が変わる。\n"
        "フルケリーは成長率が最大だが、半減確率がちょうど50%になるため実務では使えない。\n"
        "ケリー倍率2.0では成長率が0になり、それ以上は必ず破産する。"
    )
    return 0


# ────────────────────────────────────────────────────────────────
# years
# ────────────────────────────────────────────────────────────────

def cmd_years(args: argparse.Namespace) -> int:
    print(_rule("アルファ検証に必要な年数"))
    print("t = (α / TE) · √years  ⟹  years = (t · TE / α)²\n")
    print(f"  {'年率α':>8}{'TE=10%':>12}{'TE=15%':>12}{'TE=20%':>12}{'TE=25%':>12}")
    for alpha in (0.01, 0.02, 0.03, 0.05, 0.08, 0.12):
        cells = []
        for te in (0.10, 0.15, 0.20, 0.25):
            y = performance.years_to_significance(alpha, te, args.t)
            cells.append("∞" if math.isinf(y) else f"{y:,.0f}年")
        print(f"  {alpha * 100:>7.0f}%" + "".join(f"{c:>12}" for c in cells))

    y = performance.years_to_significance(args.alpha, args.te, args.t)
    print(
        f"\n指定値 (α={args.alpha * 100:.1f}%, TE={args.te * 100:.1f}%, t={args.t}): "
        + ("到達しない" if math.isinf(y) else f"{y:,.1f}年")
    )
    print(
        "\nリターンで自分の実力を検証するのは、人生の長さでは不可能である。\n"
        "だから代わりにキャリブレーションを測る（ev calibration）。\n"
        "確率予測なら年20〜30件でも傾向が見える。"
    )
    return 0


# ────────────────────────────────────────────────────────────────
# screen
# ────────────────────────────────────────────────────────────────

def _load_securities(path: Path) -> List[universe.Security]:
    def num(row, key, default=0.0):
        v = (row.get(key) or "").strip()
        if not v:
            return default
        return float(v)

    out: List[universe.Security] = []
    with path.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            ticker = (row.get("ticker") or "").strip()
            if not ticker:
                continue
            net_cash = (row.get("net_cash") or "").strip()
            pbr = (row.get("pbr") or "").strip()
            out.append(
                universe.Security(
                    ticker=ticker,
                    name=(row.get("name") or "").strip(),
                    market_cap=num(row, "market_cap"),
                    adv_value=num(row, "adv_value"),
                    price=num(row, "price"),
                    analyst_count=int(num(row, "analyst_count")),
                    sector=(row.get("sector") or "").strip(),
                    net_cash=float(net_cash) if net_cash else None,
                    pbr=float(pbr) if pbr else None,
                )
            )
    return out


def cmd_screen(args: argparse.Namespace) -> int:
    path = Path(args.csv)
    if not path.exists():
        print(f"エラー: {path} が見つかりません", file=sys.stderr)
        return 1
    secs = _load_securities(path)
    if not secs:
        print(f"エラー: {path} から銘柄を読み込めませんでした", file=sys.stderr)
        return 1

    profile = universe.InvestorProfile(
        portfolio_value=args.portfolio,
        target_position_weight=args.target_weight,
    )
    print(_rule("構造的ユニバース選別"))
    print(
        f"運用額 {args.portfolio / 1e4:,.0f}万円 / "
        f"1銘柄目標 {args.target_weight * 100:.1f}%"
        f" (= {profile.target_position_value / 1e4:,.0f}万円)\n"
    )
    print(universe.universe_summary(universe.screen_universe(secs, profile)))
    print(
        "\n注意: この選別は過去リターンを一切参照していない。\n"
        "成績で銘柄を選ぶと ev diagnose が示す選抜バイアスが入る。"
    )
    return 0


# ────────────────────────────────────────────────────────────────
# forecast / resolve / journal
# ────────────────────────────────────────────────────────────────

def cmd_forecast(args: argparse.Namespace) -> int:
    path = Path(args.file)
    if not path.exists():
        print(f"エラー: {path} が見つかりません", file=sys.stderr)
        return 1
    data = json.loads(path.read_text(encoding="utf-8"))
    fc = Forecast(
        ticker=data["ticker"],
        thesis=data["thesis"],
        horizon_months=int(data["horizon_months"]),
        structural_reason=data.get("structural_reason", ""),
        claims=[Claim.from_dict(c) for c in data.get("claims", [])],
        scenarios=[
            Scenario(s["name"], float(s["probability"]), float(s["total_return"]))
            for s in data.get("scenarios", [])
        ],
    )
    with ForecastJournal(args.journal) as j:
        fid = j.record(fc)
    print(f"登録しました: {fid} ({fc.ticker})")
    if fc.scenarios:
        print(f"  算術期待リターン: {fc.expected_return() * 100:+.2f}%")
        lg = fc.expected_log_return()
        print(
            f"  対数期待成長率  : "
            + ("-∞（全損シナリオあり）" if lg == -math.inf else f"{lg:+.5f}")
        )
    return 0


def cmd_resolve(args: argparse.Namespace) -> int:
    outcome = args.outcome.lower() in ("true", "yes", "1", "y", "はい")
    with ForecastJournal(args.journal) as j:
        j.resolve(args.id, args.claim, outcome=outcome, note=args.note or "")
    print(f"記録しました: {args.id}/{args.claim} = {outcome}")
    return 0


def cmd_journal(args: argparse.Namespace) -> int:
    with ForecastJournal(args.journal) as j:
        ok, msg = j.verify()
        print(_rule("予測台帳"))
        print(f"整合性: {'OK' if ok else '異常'} — {msg}")

        forecasts = j.all_forecasts()
        print(f"\n登録済み予測: {len(forecasts)}件")
        for fc in forecasts:
            print(f"  {fc.forecast_id}  {fc.ticker:<10} {fc.thesis[:40]}")

        resolved = j.resolved_claims()
        print(f"\n解決済みの主張: {len(resolved)}件")

        overdue = j.overdue_claims()
        if overdue:
            print(f"\n期限超過で未解決: {len(overdue)}件 ← ここを放置すると実力は測れない")
            for fid, c in overdue:
                print(f"  {fid}/{c.claim_id} (期限 {c.deadline}) {c.statement[:44]}")
        else:
            print("\n期限超過の未解決なし。")
    return 0 if ok else 1


def cmd_calibration(args: argparse.Namespace) -> int:
    with ForecastJournal(args.journal) as j:
        resolved = j.resolved_claims()
    if not resolved:
        print("解決済みの予測がありません。まず予測を記録し、期限が来たら結果を入れること。")
        return 1
    report = calibration.evaluate(
        [r.probability for r in resolved],
        [r.outcome for r in resolved],
        n_bins=args.bins,
    )
    print(_rule("キャリブレーション評価"))
    print(report.summary())

    gate, reason = sizing.calibration_gate(report)
    print(f"\nサイジング倍率: {gate:.2f}  ({reason})")
    return 0


def cmd_size(args: argparse.Namespace) -> int:
    with ForecastJournal(args.journal) as j:
        fc = j.get(args.id)
        if fc is None:
            print(f"エラー: {args.id} が見つかりません", file=sys.stderr)
            return 1
        resolved = j.resolved_claims()

    report = None
    if resolved:
        report = calibration.evaluate(
            [r.probability for r in resolved], [r.outcome for r in resolved]
        )

    cons = sizing.SizingConstraints(
        portfolio_value=args.portfolio,
        max_weight=args.max_weight,
        kelly_fraction=args.kelly_fraction,
    )
    result = sizing.size_position(
        fc, constraints=cons, calibration=report, adv_value=args.adv
    )
    print(_rule(f"サイジング: {fc.ticker} ({fc.forecast_id})"))
    print(result.summary())
    print(
        f"\n金額換算: {result.weight * args.portfolio / 1e4:,.1f}万円 "
        f"（運用額 {args.portfolio / 1e4:,.0f}万円）"
    )
    return 0


# ────────────────────────────────────────────────────────────────
# demo
# ────────────────────────────────────────────────────────────────

def cmd_demo(args: argparse.Namespace) -> int:
    rng = random.Random(20260726)

    print(_rule("ev デモ: 期待値ベースの投資プロセス全工程"))

    # ── 0. サイジングの直観 ──
    print(_section("0. 期待値がプラスでも資産が減る例"))
    b = performance.two_point_bet(0.50, -0.40, 0.5)
    print(b.summary())

    # ── 1. 既存バックテストの診断 ──
    print(_section("1. 成績で銘柄を選ぶと何が起きるか"))
    inf = validation.selection_inflation(
        n_candidates=4000, n_selected=116, n_obs=2500,
        per_obs_vol=0.02, n_simulations=1500,
    )
    print(inf.summary())

    # ── 2. 構造的ユニバース選別 ──
    print(_section("2. 大口が入れない領域を探す"))
    secs = [
        universe.Security("1234.T", "小型建設", 40e8, 10_000_000, 1500, 0, "建設",
                          net_cash=24e8, pbr=0.6),
        universe.Security("5678.T", "地方部品", 80e8, 12_000_000, 900, 1, "機械",
                          net_cash=20e8, pbr=0.8),
        universe.Security("9012.T", "中堅IT", 600e8, 400_000_000, 2200, 6, "情報通信",
                          net_cash=50e8, pbr=1.8),
        universe.Security("7203.T", "大型製造", 400_000e8, 50_000_000_000, 3000, 22,
                          "自動車", pbr=1.1),
        universe.Security("3333.T", "超薄商い", 15e8, 300_000, 400, 0, "小売",
                          net_cash=9e8, pbr=0.4),
    ]
    profile = universe.InvestorProfile(portfolio_value=10_000_000,
                                       target_position_weight=0.05)
    results = universe.screen_universe(secs, profile)
    print(universe.universe_summary(results))

    # ── 3. ベースレート ──
    print(_section("3. 物語より先にベースレートを置く"))
    observations = []
    for bucket, rate, n in (
        (("小型", "建設"), 0.34, 420),
        (("小型", "機械"), 0.31, 380),
        (("中型", "情報通信"), 0.46, 260),
        (("大型", "自動車"), 0.41, 180),
    ):
        for _ in range(n):
            observations.append(
                base_rates.Observation(bucket, rng.random() < rate)
            )
    table = base_rates.build_base_rates(observations)
    print("事象: 3年後に営業利益が5割以上増加")
    print(table.summary())

    br = table.get(("小型", "建設")).shrunk_rate
    check = base_rates.check_against_base_rate(0.62, br)
    print(f"\n自分の確率 0.62 に対する検定:\n  {check.message}")

    # ── 4. 予測の記録とキャリブレーション ──
    print(_section("4. 買う前に確率を書き、後で当て具合を測る"))
    journal_path = Path(args.journal)
    if journal_path.exists() and args.reset:
        journal_path.unlink()

    with ForecastJournal(str(journal_path)) as j:
        if not j.all_forecasts():
            # 過去数年ぶんの予測実績を投入（やや自信過剰な予測者を模擬）
            n_hist = 150
            for i in range(n_hist):
                stated = [0.15, 0.35, 0.55, 0.75, 0.92][i % 5]
                fid = j.record(
                    Forecast(
                        ticker=f"H{i:03d}.T",
                        thesis=f"過去の実績記録 #{i}。小型株の構造的歪みに対する見立て。",
                        horizon_months=12,
                        claims=[
                            Claim(
                                statement=f"検証可能な事象 #{i} が期限内に実現する",
                                probability=stated,
                                deadline="2027-12-31",
                                resolution_source="決算短信および適時開示",
                                claim_id="c1",
                            )
                        ],
                    ),
                    as_of=date(2026, 1, 5),
                )
                # 実際の実現率は述べた確率より中央寄り（＝やや自信過剰）
                true_p = 0.5 + (stated - 0.5) * 0.85
                j.resolve(fid, "c1", outcome=rng.random() < true_p,
                          resolved_at=date(2027, 1, 5))

        resolved = j.resolved_claims()
        report = calibration.evaluate(
            [r.probability for r in resolved], [r.outcome for r in resolved], n_bins=5
        )
        print(report.summary())

        raw_p = 0.88
        adjusted = calibration.recalibrate(raw_p, report)
        if report.is_overconfident:
            why = "自信過剰（傾き<0.8）が検出されたので中央に寄せる"
        elif report.calibration_slope > 1.2:
            why = "自信不足が検出されたので確率を外側に伸ばす"
        else:
            why = f"傾き {report.calibration_slope:.2f} でほぼ適正。切片ぶんだけ補正"
        print(f"\n較正補正: 生の確率 {raw_p:.2f} → {adjusted:.3f}  （{why}）")

        # 記録が少ない段階ではゲートが閉じることを示す
        print("\n記録件数によるサイジング倍率の変化:")
        for k in (10, 20, 40, len(resolved)):
            if k > len(resolved):
                continue
            sub = resolved[:k]
            rep_k = calibration.evaluate(
                [r.probability for r in sub], [r.outcome for r in sub],
                n_bins=5, n_resamples=800,
            )
            gate_k, reason_k = sizing.calibration_gate(rep_k)
            print(f"  {k:>4}件: 倍率 {gate_k:.2f}  — {reason_k}")

        # ── 5. サイジング ──
        print(_section("5. 実測キャリブレーションでゲートしたサイジング"))
        target = None
        for r in results:
            if r.passed:
                target = r
                break
        if target is None:
            print("通過銘柄がないためサイジングは省略")
            return 0

        fc = Forecast(
            ticker=target.security.ticker,
            thesis=(
                "カバレッジゼロの小型建設。純現金が時価総額の6割、PBR0.6。"
                "受注残は過去最高で、東証の資本効率改善要請の対象。"
            ),
            horizon_months=24,
            structural_reason=(
                f"参加可能な最大ファンドは{target.largest_fund / 1e8:.0f}億円。"
                f"機関投資家が意味のある建玉を作れない。"
            ),
            claims=[
                Claim(
                    statement="2028年3月期までに自己株買いまたは増配が公表される",
                    probability=0.55,
                    deadline="2028-06-30",
                    resolution_source="適時開示（自己株取得・配当予想の修正）",
                    claim_id="c1",
                    base_rate=round(br, 3),
                )
            ],
            scenarios=[
                Scenario("資本政策が動く", 0.30, 0.90),
                Scenario("業績横ばい", 0.45, 0.10),
                Scenario("受注減速", 0.25, -0.35),
            ],
        )
        fid = j.record(fc, as_of=date(2026, 7, 26))

        result = sizing.size_position(
            fc,
            constraints=sizing.SizingConstraints(
                portfolio_value=10_000_000, max_weight=0.10, kelly_fraction=0.25
            ),
            calibration=report,
            adv_value=target.security.adv_value,
        )
        print(result.summary())

        alloc = sizing.allocate_portfolio({fc.ticker: result}, satellite_budget=0.20)
        print(_section("6. ポートフォリオ"))
        print(alloc.summary())

        # ── 7. 検証可能性 ──
        print(_section("7. この成績はいつ検証できるか"))
        need = performance.years_to_significance(0.03, 0.15, 2.0)
        print(f"年率アルファ3%、TE15%、t=2 に必要な観測期間: {need:,.0f}年")
        print(
            "→ リターンでは検証不能。だから工程4のキャリブレーションが実力の測定器になる。\n"
            "  確率予測なら150件で有意差が出た。同じ実力をリターンで示すには一世紀かかる。"
        )

        ok, msg = j.verify()
        print(f"\n台帳の整合性: {'OK' if ok else '異常'} — {msg}")
    return 0


# ────────────────────────────────────────────────────────────────
# エントリポイント
# ────────────────────────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="python3 -m ev.cli",
        description="期待値ベースの株式投資プロセス支援ツール（依存パッケージなし）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = p.add_subparsers(dest="command", required=True)

    d = sub.add_parser("diagnose", help="バックテストの過剰最適化・選抜バイアスを診断")
    d.add_argument("--sharpe", type=float, required=True, help="申告された年率シャープ")
    d.add_argument("--n-obs", type=int, required=True, help="観測数（営業日数）")
    d.add_argument("--trials", type=int, required=True, help="試したパラメータ構成の数")
    d.add_argument("--candidates", type=int, help="スクリーニング前の候補銘柄数")
    d.add_argument("--selected", type=int, help="選抜後の銘柄数")
    d.add_argument("--vol", type=float, default=0.02, help="1観測あたりのリターン標準偏差")
    d.add_argument("--sims", type=int, default=2000, help="モンテカルロ回数")
    d.set_defaults(func=cmd_diagnose)

    b = sub.add_parser("bet", help="2値の賭けの期待値とケリー比率")
    b.add_argument("--win", type=float, default=0.50, help="勝ったときのリターン")
    b.add_argument("--loss", type=float, default=-0.40, help="負けたときのリターン")
    b.add_argument("--p", type=float, default=0.50, help="勝率")
    b.set_defaults(func=cmd_bet)

    y = sub.add_parser("years", help="アルファ検証に必要な年数")
    y.add_argument("--alpha", type=float, default=0.03, help="年率アルファ")
    y.add_argument("--te", type=float, default=0.15, help="年率トラッキングエラー")
    y.add_argument("--t", type=float, default=2.0, help="目標t値")
    y.set_defaults(func=cmd_years)

    s = sub.add_parser("screen", help="CSVから構造的ユニバース選別")
    s.add_argument("--csv", required=True, help="銘柄情報CSV")
    s.add_argument("--portfolio", type=float, default=10_000_000, help="運用額（円）")
    s.add_argument("--target-weight", type=float, default=0.05, help="1銘柄目標ウェイト")
    s.set_defaults(func=cmd_screen)

    f = sub.add_parser("forecast", help="予測をJSONから台帳に登録")
    f.add_argument("--journal", default="journal.db", help="台帳ファイル")
    f.add_argument("--file", required=True, help="予測JSON")
    f.set_defaults(func=cmd_forecast)

    r = sub.add_parser("resolve", help="予測の結果を記録")
    r.add_argument("--journal", default="journal.db")
    r.add_argument("--id", required=True, help="forecast_id")
    r.add_argument("--claim", required=True, help="claim_id")
    r.add_argument("--outcome", required=True, help="true / false")
    r.add_argument("--note", help="判定の根拠")
    r.set_defaults(func=cmd_resolve)

    j = sub.add_parser("journal", help="台帳の内容と未解決の予測を表示")
    j.add_argument("--journal", default="journal.db")
    j.set_defaults(func=cmd_journal)

    c = sub.add_parser("calibration", help="台帳の実績からキャリブレーションを評価")
    c.add_argument("--journal", default="journal.db")
    c.add_argument("--bins", type=int, default=5)
    c.set_defaults(func=cmd_calibration)

    z = sub.add_parser("size", help="登録済み予測のポジションサイズを算出")
    z.add_argument("--journal", default="journal.db")
    z.add_argument("--id", required=True, help="forecast_id")
    z.add_argument("--portfolio", type=float, default=10_000_000)
    z.add_argument("--max-weight", type=float, default=0.10)
    z.add_argument("--kelly-fraction", type=float, default=0.25)
    z.add_argument("--adv", type=float, help="平均日次売買代金（円）")
    z.set_defaults(func=cmd_size)

    m = sub.add_parser("demo", help="全工程を合成データで実行")
    m.add_argument("--journal", default="demo_journal.db")
    m.add_argument("--reset", action="store_true", help="既存の台帳を削除してやり直す")
    m.set_defaults(func=cmd_demo)

    return p


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return int(args.func(args))
    except (ValueError, KeyError, OSError) as e:
        print(f"エラー: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
