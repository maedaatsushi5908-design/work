"""
kenmo氏の新高値ブレイク投資法 バックテストエンジン
"""

from __future__ import annotations

import warnings
from dataclasses import dataclass, field
from typing import Dict, List, Optional

import numpy as np
import pandas as pd

from config import BacktestConfig, StrategyConfig
from strategy import compute_signals, get_entry_date, get_entry_price

warnings.filterwarnings("ignore")


@dataclass
class Trade:
    ticker: str
    entry_date: pd.Timestamp
    entry_price: float
    shares: float
    capital_used: float

    exit_date: Optional[pd.Timestamp] = None
    exit_price: Optional[float] = None
    exit_reason: Optional[str] = None  # "stop_loss" | "trailing_stop" | "end_of_period"

    @property
    def is_open(self) -> bool:
        return self.exit_date is None

    @property
    def pnl(self) -> float:
        if self.exit_price is None:
            return 0.0
        return (self.exit_price - self.entry_price) * self.shares

    @property
    def return_pct(self) -> float:
        if self.exit_price is None or self.entry_price == 0:
            return 0.0
        return (self.exit_price - self.entry_price) / self.entry_price

    @property
    def holding_days(self) -> int:
        if self.exit_date is None:
            return 0
        return (self.exit_date - self.entry_date).days


@dataclass
class BacktestResult:
    trades: List[Trade]
    equity_curve: pd.Series  # 日次資産推移
    initial_capital: float

    @property
    def closed_trades(self) -> List[Trade]:
        return [t for t in self.trades if not t.is_open]

    @property
    def total_trades(self) -> int:
        return len(self.closed_trades)

    @property
    def win_trades(self) -> int:
        return sum(1 for t in self.closed_trades if t.pnl > 0)

    @property
    def win_rate(self) -> float:
        if self.total_trades == 0:
            return 0.0
        return self.win_trades / self.total_trades

    @property
    def total_return(self) -> float:
        if self.initial_capital == 0:
            return 0.0
        return (self.equity_curve.iloc[-1] - self.initial_capital) / self.initial_capital

    @property
    def avg_return_per_trade(self) -> float:
        if not self.closed_trades:
            return 0.0
        return float(np.mean([t.return_pct for t in self.closed_trades]))

    @property
    def avg_win(self) -> float:
        wins = [t.return_pct for t in self.closed_trades if t.pnl > 0]
        return float(np.mean(wins)) if wins else 0.0

    @property
    def avg_loss(self) -> float:
        losses = [t.return_pct for t in self.closed_trades if t.pnl <= 0]
        return float(np.mean(losses)) if losses else 0.0

    @property
    def profit_factor(self) -> float:
        gross_profit = sum(t.pnl for t in self.closed_trades if t.pnl > 0)
        gross_loss = abs(sum(t.pnl for t in self.closed_trades if t.pnl < 0))
        if gross_loss == 0:
            return float("inf")
        return gross_profit / gross_loss

    @property
    def max_drawdown(self) -> float:
        equity = self.equity_curve
        peak = equity.cummax()
        drawdown = (equity - peak) / peak
        return float(drawdown.min())

    @property
    def sharpe_ratio(self) -> float:
        daily_returns = self.equity_curve.pct_change().dropna()
        if daily_returns.std() == 0:
            return 0.0
        return float(daily_returns.mean() / daily_returns.std() * np.sqrt(252))

    @property
    def avg_holding_days(self) -> float:
        if not self.closed_trades:
            return 0.0
        return float(np.mean([t.holding_days for t in self.closed_trades]))


class Backtester:
    def __init__(
        self,
        stock_data: Dict[str, pd.DataFrame],
        strategy_config: StrategyConfig,
        backtest_config: BacktestConfig,
    ):
        self.stock_data = stock_data
        self.scfg = strategy_config
        self.bcfg = backtest_config

        # 全取引日の生成（データが存在する全日付）
        all_dates: set = set()
        for df in stock_data.values():
            all_dates.update(df.index.tolist())
        self.all_dates = sorted(all_dates)

        # シグナルを事前計算
        self.signals: Dict[str, pd.DataFrame] = {}
        for ticker, df in stock_data.items():
            self.signals[ticker] = compute_signals(df, strategy_config)

    def run(self) -> BacktestResult:
        capital = self.bcfg.initial_capital
        cash = capital
        positions: Dict[str, Trade] = {}  # ticker -> Trade（1銘柄1ポジション想定）
        trades: List[Trade] = []
        equity_history: Dict[pd.Timestamp, float] = {}

        for date in self.all_dates:
            # --- 1. 既存ポジションのエグジット判定 ---
            for ticker in list(positions.keys()):
                trade = positions[ticker]
                df = self.stock_data[ticker]

                if date not in df.index:
                    continue

                current_low = float(df.loc[date, "Low"])
                current_close = float(df.loc[date, "Close"])

                # 損切りライン
                stop_loss_price = trade.entry_price * (1 + self.scfg.stop_loss_pct)

                # トレーリングストップ: 保有開始から当日までの最高値を追跡
                entry_idx = df.index.get_loc(trade.entry_date)
                date_idx = df.index.get_loc(date)
                if date_idx >= entry_idx:
                    peak_price = float(df.iloc[entry_idx : date_idx + 1]["High"].max())
                else:
                    peak_price = trade.entry_price
                trailing_stop_price = peak_price * (1 + self.scfg.trailing_stop_pct)

                exit_price = None
                exit_reason = None

                # 損切り判定（当日安値がストップロスを下回った）
                if current_low <= stop_loss_price:
                    exit_price = max(stop_loss_price, current_low)
                    exit_reason = "stop_loss"
                # トレーリングストップ判定
                elif current_low <= trailing_stop_price:
                    exit_price = max(trailing_stop_price, current_low)
                    exit_reason = "trailing_stop"

                if exit_price is not None:
                    # スリッページ・手数料を考慮
                    exit_price *= 1 - self.bcfg.slippage_pct
                    exit_price *= 1 - self.bcfg.commission_pct

                    trade.exit_date = date
                    trade.exit_price = exit_price
                    trade.exit_reason = exit_reason

                    cash += exit_price * trade.shares
                    trades.append(trade)
                    del positions[ticker]

            # --- 2. エントリーシグナルの確認 ---
            for ticker, sig_df in self.signals.items():
                # すでにポジション保有中 or 最大ポジション数に達している
                if ticker in positions:
                    continue
                if len(positions) >= self.scfg.max_positions:
                    break

                if date not in sig_df.index:
                    continue

                # 当日がシグナル日かどうか確認
                if not sig_df.loc[date, "signal"]:
                    continue

                # エントリー価格・日付を決定
                df = self.stock_data[ticker]
                entry_date = get_entry_date(df, date, self.scfg)
                if pd.isna(entry_date) or entry_date not in df.index:
                    continue
                entry_price = get_entry_price(df, date, self.scfg)
                if np.isnan(entry_price) or entry_price <= 0:
                    continue

                # スリッページ・手数料を考慮
                entry_price *= 1 + self.bcfg.slippage_pct
                entry_price *= 1 + self.bcfg.commission_pct

                # 購入資金の計算
                capital_for_trade = capital * self.scfg.position_size_pct
                if capital_for_trade > cash:
                    capital_for_trade = cash
                if capital_for_trade <= 0:
                    continue

                shares = capital_for_trade / entry_price

                trade = Trade(
                    ticker=ticker,
                    entry_date=entry_date,
                    entry_price=entry_price,
                    shares=shares,
                    capital_used=capital_for_trade,
                )
                positions[ticker] = trade
                cash -= capital_for_trade

            # --- 3. 当日の資産評価 ---
            position_value = 0.0
            for ticker, trade in positions.items():
                df = self.stock_data[ticker]
                if date in df.index:
                    current_price = float(df.loc[date, "Close"])
                    position_value += current_price * trade.shares
                else:
                    position_value += trade.entry_price * trade.shares

            equity = cash + position_value
            equity_history[date] = equity
            capital = equity  # 複利計算のため資産総額を更新

        # --- 4. 期末に残っているポジションを強制クローズ ---
        last_date = self.all_dates[-1]
        for ticker, trade in list(positions.items()):
            df = self.stock_data[ticker]
            close_date = last_date
            if close_date not in df.index:
                available = df.index[df.index <= close_date]
                if len(available) == 0:
                    continue
                close_date = available[-1]

            exit_price = float(df.loc[close_date, "Close"])
            exit_price *= 1 - self.bcfg.commission_pct

            trade.exit_date = close_date
            trade.exit_price = exit_price
            trade.exit_reason = "end_of_period"
            trades.append(trade)

        equity_series = pd.Series(equity_history).sort_index()

        return BacktestResult(
            trades=trades,
            equity_curve=equity_series,
            initial_capital=self.bcfg.initial_capital,
        )
