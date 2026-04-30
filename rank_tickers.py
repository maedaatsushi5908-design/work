"""各銘柄を個別にバックテストして成績でランキングする"""
import sys, warnings
warnings.filterwarnings("ignore")
if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import pandas as pd
from backtest import Backtester
from config import BacktestConfig, StrategyConfig, TSE_PRIME_TOP150
from data import download_all_stocks

BCFG = BacktestConfig()
SCFG = StrategyConfig()

print("全銘柄データを取得中...")
all_data = download_all_stocks(TSE_PRIME_TOP150, BCFG.start_date, BCFG.end_date)
print(f"取得完了: {len(all_data)} 銘柄\n")

rows = []
for ticker, df in all_data.items():
    bt = Backtester({ticker: df}, SCFG, BCFG)
    r = bt.run()
    if r.total_trades < 2:
        continue
    rows.append({
        "ticker":      ticker,
        "総リターン":  round(r.total_return * 100, 1),
        "勝率":        round(r.win_rate * 100, 0),
        "シャープ":    round(r.sharpe_ratio, 2),
        "PF":          round(r.profit_factor, 2),
        "最大DD":      round(r.max_drawdown * 100, 1),
        "取引数":      r.total_trades,
    })

df_rank = pd.DataFrame(rows)

# スコア: シャープ × PF × (1 + リターン) / (1 + |DD|)
df_rank["スコア"] = (
    df_rank["シャープ"]
    * df_rank["PF"]
    * (1 + df_rank["総リターン"] / 100)
    / (1 + df_rank["最大DD"].abs() / 100)
).round(3)

df_rank = df_rank.sort_values("スコア", ascending=False).reset_index(drop=True)

print("=" * 75)
print("  銘柄別スコアランキング（上位60件）")
print("=" * 75)
print(df_rank.head(60).to_string(index=True))

print("\n上位50銘柄のティッカーリスト:")
top50 = df_rank.head(50)["ticker"].tolist()
print(top50)
