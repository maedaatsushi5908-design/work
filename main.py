"""
kenmo氏の新高値ブレイク投資法 バックテスト エントリーポイント

使い方:
    python main.py [オプション]

オプション:
    --start       バックテスト開始日 (例: 2015-01-01)
    --end         バックテスト終了日 (例: 2024-12-31)
    --capital     初期資金 (例: 1000000)
    --high-period 新高値判定期間 (営業日, 例: 260 = 52週)
    --stop-loss   損切りライン (例: -0.08 = -8%)
    --trailing    トレーリングストップ (例: -0.15 = -15%)
    --max-pos     最大保有銘柄数 (例: 5)
    --no-cache    キャッシュを使わず再取得
    --show        グラフを画面表示する
    --tickers     対象ティッカー (スペース区切り, 省略時はデフォルト銘柄リスト)
"""

import argparse
import sys

from config import BacktestConfig, StrategyConfig, DEFAULT_TICKERS
from data import download_all_stocks, download_benchmark
from backtest import Backtester
from report import print_summary, plot_equity_curve, plot_trade_distribution, plot_monthly_returns, save_trade_log


def parse_args():
    _defaults = StrategyConfig()
    _bt_defaults = BacktestConfig()
    parser = argparse.ArgumentParser(
        description="kenmo氏の新高値ブレイク投資法 バックテスト",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--start", default=_bt_defaults.start_date, help="バックテスト開始日")
    parser.add_argument("--end", default=_bt_defaults.end_date, help="バックテスト終了日")
    parser.add_argument("--capital", type=float, default=_bt_defaults.initial_capital, help="初期資金（円）")
    parser.add_argument("--high-period", type=int, default=_defaults.high_period, help="新高値判定期間（営業日）")
    parser.add_argument("--stop-loss", type=float, default=_defaults.stop_loss_pct, help="損切りライン（例: -0.08）")
    parser.add_argument("--trailing", type=float, default=_defaults.trailing_stop_pct, help="トレーリングストップ（例: -0.15）")
    parser.add_argument("--max-pos", type=int, default=_defaults.max_positions, help="最大保有銘柄数")
    parser.add_argument("--no-cache", action="store_true", help="キャッシュを使わない")
    parser.add_argument("--show", action="store_true", help="グラフを画面表示する")
    parser.add_argument("--tickers", nargs="+", default=None, help="対象ティッカー")
    return parser.parse_args()


def main():
    args = parse_args()

    tickers = args.tickers if args.tickers else DEFAULT_TICKERS

    print("\n" + "=" * 60)
    print("  kenmo 新高値ブレイク投資法 バックテスト")
    print("=" * 60)
    print(f"  期間          : {args.start} 〜 {args.end}")
    print(f"  初期資金      : {args.capital:,.0f} 円")
    print(f"  新高値判定期間: {args.high_period} 営業日")
    print(f"  損切りライン  : {args.stop_loss:.1%}")
    print(f"  トレーリング  : {args.trailing:.1%}")
    print(f"  最大保有銘柄数: {args.max_pos}")
    print(f"  対象銘柄数    : {len(tickers)}")
    print()

    # 設定
    strategy_cfg = StrategyConfig(
        high_period=args.high_period,
        stop_loss_pct=args.stop_loss,
        trailing_stop_pct=args.trailing,
        max_positions=args.max_pos,
        position_size_pct=1.0 / args.max_pos,
    )
    backtest_cfg = BacktestConfig(
        start_date=args.start,
        end_date=args.end,
        initial_capital=args.capital,
    )

    use_cache = not args.no_cache

    # データ取得
    stock_data = download_all_stocks(tickers, args.start, args.end, use_cache=use_cache)
    if not stock_data:
        print("エラー: 株価データを取得できませんでした。")
        sys.exit(1)

    benchmark = download_benchmark(args.start, args.end, use_cache=use_cache)

    # バックテスト実行
    print("\nバックテスト実行中...")
    backtester = Backtester(stock_data, strategy_cfg, backtest_cfg)
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
