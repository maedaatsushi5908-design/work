"""
新高値ブレイク × 移動平均線フィルター 投資システム エントリーポイント

使い方:
    python main.py [オプション]

基本オプション:
    --start       バックテスト開始日 (例: 2015-01-01)
    --end         バックテスト終了日 (例: 2024-12-31)
    --capital     初期資金 (例: 1000000)
    --high-period 新高値判定期間 (営業日, 例: 195 = 39週)
    --stop-loss   損切りライン (例: -0.10 = -10%)
    --trailing    トレーリングストップ (例: -0.20 = -20%)
    --max-pos     最大保有銘柄数 (例: 5)
    --no-cache    キャッシュを使わず再取得
    --show        グラフを画面表示する
    --tickers     対象ティッカー (スペース区切り, 省略時はデフォルト銘柄リスト)

移動平均線フィルター:
    --ma-short        短期移動平均 (営業日, デフォルト 25)
    --ma-mid          中期移動平均 (営業日, デフォルト 75)
    --ma-long         長期移動平均 (営業日, デフォルト 200)
    --ma-slope-days   長期移動平均の傾きを見る期間 (デフォルト 20)
    --no-ma-filter    移動平均フィルターを無効にする
    --no-alignment    パーフェクトオーダー条件を外す
    --ma-exit         MA割れエグジットを有効にする (デフォルトは無効)
    --ma-exit-period  MA割れエグジットに使う移動平均 (デフォルト 75)
    --ma-exit-days    何日連続で割れたら手仕舞うか (デフォルト 2)
    --market-ma       地合い判定の移動平均 (デフォルト 200)
    --no-market-filter 地合いフィルターを無効にする

比較モード:
    --compare     フィルターを1段ずつ追加した4パターンを比較表示する
"""

import argparse
import sys
from dataclasses import replace

from config import BacktestConfig, StrategyConfig, DEFAULT_TICKERS
from data import download_all_stocks, download_benchmark
from backtest import Backtester
from report import print_summary, plot_equity_curve, plot_trade_distribution, plot_monthly_returns, save_trade_log

# Windows stdout を UTF-8 に
if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def parse_args():
    _defaults = StrategyConfig()
    _bt_defaults = BacktestConfig()
    parser = argparse.ArgumentParser(
        description="新高値ブレイク × 移動平均線フィルター バックテスト",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--start", default=_bt_defaults.start_date, help="バックテスト開始日")
    parser.add_argument("--end", default=_bt_defaults.end_date, help="バックテスト終了日")
    parser.add_argument("--capital", type=float, default=_bt_defaults.initial_capital, help="初期資金（円）")
    parser.add_argument("--high-period", type=int, default=_defaults.high_period, help="新高値判定期間（営業日）")
    parser.add_argument("--stop-loss", type=float, default=_defaults.stop_loss_pct, help="損切りライン（例: -0.10）")
    parser.add_argument("--trailing", type=float, default=_defaults.trailing_stop_pct, help="トレーリングストップ（例: -0.20）")
    parser.add_argument("--max-pos", type=int, default=_defaults.max_positions, help="最大保有銘柄数")
    parser.add_argument("--no-cache", action="store_true", help="キャッシュを使わない")
    parser.add_argument("--show", action="store_true", help="グラフを画面表示する")
    parser.add_argument("--tickers", nargs="+", default=None, help="対象ティッカー")

    # ── 移動平均線フィルター ──────────────────────────
    ma = parser.add_argument_group("移動平均線フィルター")
    ma.add_argument("--ma-short", type=int, default=_defaults.ma_short_period, help="短期移動平均（営業日）")
    ma.add_argument("--ma-mid", type=int, default=_defaults.ma_mid_period, help="中期移動平均（営業日）")
    ma.add_argument("--ma-long", type=int, default=_defaults.ma_long_period, help="長期移動平均（営業日）")
    ma.add_argument("--ma-slope-days", type=int, default=_defaults.ma_slope_days, help="長期移動平均の傾き判定期間")
    ma.add_argument("--no-ma-filter", action="store_true", help="移動平均フィルターを無効にする")
    ma.add_argument("--no-alignment", action="store_true", help="パーフェクトオーダー条件を外す")
    ma.add_argument("--ma-exit", action="store_true", help="MA割れエグジットを有効にする")
    ma.add_argument("--ma-exit-period", type=int, default=_defaults.ma_exit_period, help="MA割れエグジットの移動平均")
    ma.add_argument("--ma-exit-days", type=int, default=_defaults.ma_exit_confirm_days, help="MA割れ確認日数")
    ma.add_argument("--market-ma", type=int, default=_defaults.market_ma_period, help="地合い判定の移動平均")
    ma.add_argument("--no-market-filter", action="store_true", help="地合いフィルターを無効にする")

    parser.add_argument("--compare", action="store_true", help="フィルター有無を比較する")
    return parser.parse_args()


def build_strategy_config(args) -> StrategyConfig:
    return StrategyConfig(
        high_period=args.high_period,
        stop_loss_pct=args.stop_loss,
        trailing_stop_pct=args.trailing,
        max_positions=args.max_pos,
        position_size_pct=1.0 / args.max_pos,
        ma_filter=not args.no_ma_filter,
        ma_short_period=args.ma_short,
        ma_mid_period=args.ma_mid,
        ma_long_period=args.ma_long,
        ma_slope_days=args.ma_slope_days,
        ma_require_alignment=not args.no_alignment,
        ma_exit=args.ma_exit,
        ma_exit_period=args.ma_exit_period,
        ma_exit_confirm_days=args.ma_exit_days,
        market_filter=not args.no_market_filter,
        market_ma_period=args.market_ma,
    )


def print_config(args, cfg: StrategyConfig, ticker_count: int) -> None:
    print("\n" + "=" * 64)
    print("  新高値ブレイク × 移動平均線フィルター バックテスト")
    print("=" * 64)
    print(f"  期間            : {args.start} 〜 {args.end}")
    print(f"  初期資金        : {args.capital:,.0f} 円")
    print(f"  新高値判定期間  : {cfg.high_period} 営業日")
    print(f"  損切りライン    : {cfg.stop_loss_pct:.1%}")
    print(f"  トレーリング    : {cfg.trailing_stop_pct:.1%}")
    print(f"  最大保有銘柄数  : {cfg.max_positions}")
    print(f"  対象銘柄数      : {ticker_count}")
    print("-" * 64)
    if cfg.ma_filter:
        conditions = []
        if cfg.ma_require_above:
            conditions.append("終値>全MA")
        if cfg.ma_require_alignment:
            conditions.append("パーフェクトオーダー")
        if cfg.ma_require_slope_up:
            conditions.append(f"長期MA上向き({cfg.ma_slope_days}日)")
        print(f"  MAフィルター    : {cfg.ma_short_period}/{cfg.ma_mid_period}/{cfg.ma_long_period}日  "
              f"[{' + '.join(conditions)}]")
    else:
        print("  MAフィルター    : なし")
    if cfg.ma_exit:
        print(f"  MA割れエグジット: {cfg.ma_exit_period}日MAを{cfg.ma_exit_confirm_days}日連続で下回ったら手仕舞い")
    else:
        print("  MA割れエグジット: なし")
    if cfg.market_filter:
        print(f"  地合いフィルター: {cfg.market_ticker} が {cfg.market_ma_period}日MAより上のときだけ新規建て")
    else:
        print("  地合いフィルター: なし")
    print()


def run_compare(stock_data, benchmark, base_cfg: StrategyConfig, bcfg: BacktestConfig) -> None:
    """フィルターを1段ずつ追加して効果を比較する。"""
    variants = [
        ("① ベースライン（新高値ブレイクのみ）",
         replace(base_cfg, ma_filter=False, ma_exit=False, market_filter=False)),
        ("② ＋ MAフィルター（エントリー）",
         replace(base_cfg, ma_filter=True, ma_exit=False, market_filter=False)),
        ("③ ＋ MA割れエグジット",
         replace(base_cfg, ma_filter=True, ma_exit=True, market_filter=False)),
        ("④ ＋ 地合いフィルター（フル）",
         replace(base_cfg, ma_filter=True, ma_exit=True, market_filter=True)),
    ]

    print("\nフィルター比較を実行中...")
    rows = []
    for label, cfg in variants:
        bt = Backtester(stock_data, cfg, bcfg, benchmark=benchmark)
        r = bt.run()
        rows.append((label, r))
        print(f"  完了: {label}")

    print("\n" + "=" * 96)
    print("  移動平均線フィルターの効果比較")
    print("=" * 96)
    header = (f"  {'パターン':<34} {'総リターン':>10} {'シャープ':>8} {'最大DD':>9} "
              f"{'勝率':>6} {'PF':>6} {'取引数':>7} {'平均保有':>8}")
    print(header)
    print("  " + "-" * 92)
    for label, r in rows:
        print(f"  {label:<34} {r.total_return:>9.1%} {r.sharpe_ratio:>8.2f} "
              f"{r.max_drawdown:>8.1%} {r.win_rate:>6.0%} {r.profit_factor:>6.2f} "
              f"{r.total_trades:>7} {r.avg_holding_days:>7.0f}日")
    print("=" * 96)

    base_r = rows[0][1]
    full_r = rows[-1][1]
    print("\n  【ベースライン → フル装備の変化】")
    print(f"    総リターン      : {base_r.total_return:>8.1%} → {full_r.total_return:>8.1%}")
    print(f"    シャープレシオ  : {base_r.sharpe_ratio:>8.2f} → {full_r.sharpe_ratio:>8.2f}")
    print(f"    最大ドローダウン: {base_r.max_drawdown:>8.1%} → {full_r.max_drawdown:>8.1%}")
    print(f"    勝率            : {base_r.win_rate:>8.0%} → {full_r.win_rate:>8.0%}")
    print(f"    取引数          : {base_r.total_trades:>8} → {full_r.total_trades:>8}")
    print()


def main():
    args = parse_args()

    tickers = args.tickers if args.tickers else DEFAULT_TICKERS
    strategy_cfg = build_strategy_config(args)
    backtest_cfg = BacktestConfig(
        start_date=args.start,
        end_date=args.end,
        initial_capital=args.capital,
    )

    print_config(args, strategy_cfg, len(tickers))

    use_cache = not args.no_cache

    # データ取得
    stock_data = download_all_stocks(tickers, args.start, args.end, use_cache=use_cache)
    if not stock_data:
        print("エラー: 株価データを取得できませんでした。")
        sys.exit(1)

    benchmark = download_benchmark(args.start, args.end, use_cache=use_cache)
    if benchmark is None and strategy_cfg.market_filter:
        print("警告: 指数データを取得できないため、地合いフィルターは無効になります。")

    if args.compare:
        run_compare(stock_data, benchmark, strategy_cfg, backtest_cfg)
        return

    # バックテスト実行
    print("\nバックテスト実行中...")
    backtester = Backtester(stock_data, strategy_cfg, backtest_cfg, benchmark=benchmark)
    result = backtester.run()

    # 結果表示
    print_summary(result, benchmark)

    # グラフ・ログ出力
    print("結果を出力中...")
    plot_equity_curve(result, benchmark, save=True, show=args.show)
    plot_trade_distribution(result, save=True, show=args.show)
    plot_monthly_returns(result, save=True, show=args.show)
    save_trade_log(result)

    print("\n完了。output/ フォルダに結果が保存されました。")


if __name__ == "__main__":
    main()
