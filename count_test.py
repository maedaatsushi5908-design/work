import sys, warnings
warnings.filterwarnings("ignore")
if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from backtest import Backtester
from config import BacktestConfig, StrategyConfig, TSE_PRIME_TOP150
from data import download_all_stocks

BCFG = BacktestConfig()
SCFG = StrategyConfig()

print("全銘柄データを取得中...")
all_data = download_all_stocks(TSE_PRIME_TOP150, BCFG.start_date, BCFG.end_date)
tickers_available = list(all_data.keys())
print(f"取得完了: {len(tickers_available)} 銘柄\n")

header = f"{'銘柄数':>6}  {'総リターン':>10}  {'勝率':>6}  {'シャープ':>8}  {'最大DD':>8}  {'PF':>5}  {'取引数':>6}"
print(header)
print("-" * 65)

for n in [20, 30, 40, 50, 60, 70, 80, 100, len(tickers_available)]:
    tickers = tickers_available[:n]
    data = {t: all_data[t] for t in tickers}
    bt = Backtester(data, SCFG, BCFG)
    r = bt.run()
    row = f"{n:>6}  {r.total_return*100:>9.1f}%  {r.win_rate*100:>5.0f}%  {r.sharpe_ratio:>8.2f}  {r.max_drawdown*100:>7.1f}%  {r.profit_factor:>5.2f}  {r.total_trades:>6}"
    print(row)
