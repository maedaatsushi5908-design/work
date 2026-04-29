"""
kenmo 新高値ブレイク投資法 デイリースキャナー

毎朝実行することで、その日の「エントリー候補」「損切り対象」を通知する。

使い方:
    # 朝のスキャン（メインの使い方）
    python scanner.py

    # ポジションを追加（買った後に記録）
    python scanner.py add --ticker 7203.T --price 2500 --shares 100

    # ポジションを削除（売り切った後に記録）
    python scanner.py remove --ticker 7203.T

    # 保有中ポジション一覧
    python scanner.py list

    # スキャン対象銘柄を変える
    python scanner.py --tickers 7203.T 6758.T 9984.T

オプション（scanコマンド）:
    --high-period  新高値判定期間（営業日, デフォルト: 260）
    --stop-loss    損切りライン（デフォルト: -0.08）
    --trailing     トレーリングストップ（デフォルト: -0.15）
    --warn-trailing  トレーリングストップへの接近警告ライン（デフォルト: -0.10）
"""

from __future__ import annotations

import argparse
import json
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

from config import DEFAULT_TICKERS, StrategyConfig

POSITIONS_FILE = Path("positions.json")
DATA_LOOKBACK_DAYS = 400  # データ取得期間（新高値判定 + バッファ）


# ──────────────────────────────────────────
# ポジション管理
# ──────────────────────────────────────────

def load_positions() -> List[Dict]:
    if not POSITIONS_FILE.exists():
        return []
    with open(POSITIONS_FILE, "r", encoding="utf-8") as f:
        return json.load(f).get("positions", [])


def save_positions(positions: List[Dict]) -> None:
    with open(POSITIONS_FILE, "w", encoding="utf-8") as f:
        json.dump({"positions": positions}, f, ensure_ascii=False, indent=2)


def add_position(ticker: str, price: float, shares: float, note: str = "") -> None:
    positions = load_positions()
    for p in positions:
        if p["ticker"] == ticker:
            print(f"  既に {ticker} のポジションが存在します。先に remove してください。")
            return
    positions.append({
        "ticker": ticker,
        "entry_date": date.today().isoformat(),
        "entry_price": price,
        "shares": shares,
        "note": note,
    })
    save_positions(positions)
    print(f"  ✅ ポジション追加: {ticker}  {price:,.0f}円 × {shares:,.0f}株")


def remove_position(ticker: str) -> None:
    positions = load_positions()
    before = len(positions)
    positions = [p for p in positions if p["ticker"] != ticker]
    if len(positions) == before:
        print(f"  {ticker} のポジションは見つかりませんでした。")
        return
    save_positions(positions)
    print(f"  🗑️  ポジション削除: {ticker}")


def list_positions() -> None:
    positions = load_positions()
    if not positions:
        print("  保有中のポジションはありません。")
        return
    print(f"\n  {'銘柄':<12} {'エントリー日':<14} {'取得価格':>10} {'株数':>8}  メモ")
    print("  " + "-" * 60)
    for p in positions:
        print(f"  {p['ticker']:<12} {p['entry_date']:<14} {p['entry_price']:>10,.0f} {p['shares']:>8,.0f}  {p.get('note', '')}")


# ──────────────────────────────────────────
# データ取得
# ──────────────────────────────────────────

def _fetch_data(tickers: List[str], high_period: int) -> Dict[str, pd.DataFrame]:
    """最近のデータを取得する（オフライン時は合成データ）。"""
    end = date.today()
    # 新高値判定に必要な期間 + バッファ（土日・祝日を考慮して1.5倍）
    start = end - timedelta(days=int(high_period * 1.5) + 30)

    # yfinance を試みる
    import warnings, contextlib, io
    import yfinance as yf
    warnings.filterwarnings("ignore")

    def try_yf(ticker: str) -> Optional[pd.DataFrame]:
        try:
            with contextlib.redirect_stderr(io.StringIO()):
                df = yf.download(ticker, start=start.isoformat(), end=end.isoformat(),
                                 auto_adjust=True, progress=False)
            if df.empty:
                return None
            if isinstance(df.columns, pd.MultiIndex):
                df.columns = df.columns.get_level_values(0)
            df = df[["Open", "High", "Low", "Close", "Volume"]].copy()
            df.index = pd.to_datetime(df.index)
            return df.sort_index().dropna(subset=["Close"])
        except Exception:
            return None

    # ネットワーク確認
    test = try_yf(tickers[0])
    if test is None:
        print("  ⚠️  ネットワーク接続なし → 合成データで動作します（デモモード）")
        from synthetic_data import generate_stock_data
        result = {}
        for t in tickers:
            df = generate_stock_data(t, start.isoformat(), end.isoformat())
            # 「今日」を最新日として切り出す
            result[t] = df
        return result

    result = {tickers[0]: test}
    for t in tickers[1:]:
        df = try_yf(t)
        if df is not None:
            result[t] = df
    return result


# ──────────────────────────────────────────
# スキャン本体
# ──────────────────────────────────────────

def run_scan(
    scan_tickers: List[str],
    high_period: int,
    stop_loss_pct: float,
    trailing_stop_pct: float,
    warn_trailing_pct: float,
) -> None:
    today = date.today()
    print(f"\n{'='*62}")
    print(f"  kenmo 新高値ブレイク デイリースキャン  {today.strftime('%Y年%m月%d日')}")
    print(f"{'='*62}")

    positions = load_positions()
    held_tickers = [p["ticker"] for p in positions]

    # 全ティッカー（スキャン対象 + 保有中）
    all_tickers = list(dict.fromkeys(scan_tickers + held_tickers))

    print(f"\n  データ取得中... ({len(all_tickers)}銘柄)")
    data = _fetch_data(all_tickers, high_period)
    if not data:
        print("  データを取得できませんでした。")
        return

    # ── 1. 保有ポジションのチェック ──────────────────────────
    print(f"\n{'─'*62}")
    print("  【保有ポジション状況】")
    print(f"{'─'*62}")

    exit_alerts: List[Dict] = []
    warn_alerts: List[Dict] = []
    hold_ok: List[Dict] = []

    if not positions:
        print("  保有中のポジションはありません。")
    else:
        for p in positions:
            ticker = p["ticker"]
            entry_price = p["entry_price"]
            shares = p["shares"]
            entry_date = p["entry_date"]

            df = data.get(ticker)
            if df is None or df.empty:
                print(f"  ⚠️  {ticker}: データなし")
                continue

            current_price = float(df["Close"].iloc[-1])
            latest_date = df.index[-1].strftime("%Y/%m/%d")

            # エントリー以降の最高値
            entry_dt = pd.Timestamp(entry_date)
            df_since_entry = df[df.index >= entry_dt]
            peak_price = float(df_since_entry["High"].max()) if not df_since_entry.empty else entry_price

            pnl_pct = (current_price - entry_price) / entry_price
            stop_loss_price = entry_price * (1 + stop_loss_pct)
            trailing_stop_price = peak_price * (1 + trailing_stop_pct)
            warn_trailing_price = peak_price * (1 + warn_trailing_pct)
            current_value = current_price * shares
            pnl_yen = (current_price - entry_price) * shares

            info = {
                "ticker": ticker,
                "entry_price": entry_price,
                "current_price": current_price,
                "peak_price": peak_price,
                "stop_loss_price": stop_loss_price,
                "trailing_stop_price": trailing_stop_price,
                "warn_trailing_price": warn_trailing_price,
                "pnl_pct": pnl_pct,
                "pnl_yen": pnl_yen,
                "current_value": current_value,
                "shares": shares,
                "entry_date": entry_date,
                "latest_date": latest_date,
            }

            if current_price <= stop_loss_price:
                exit_alerts.append({**info, "reason": "stop_loss"})
            elif current_price <= trailing_stop_price:
                exit_alerts.append({**info, "reason": "trailing_stop"})
            elif current_price <= warn_trailing_price:
                warn_alerts.append(info)
            else:
                hold_ok.append(info)

        # 損切り・トレーリングストップ
        if exit_alerts:
            print()
            for info in exit_alerts:
                reason_label = "損切りライン到達" if info["reason"] == "stop_loss" else "トレーリングストップ到達"
                print(f"  🔴 【売り推奨】{info['ticker']}  {reason_label}")
                print(f"       取得価格: {info['entry_price']:>10,.0f}円  現在値: {info['current_price']:>10,.0f}円")
                print(f"       損益: {info['pnl_pct']:+.1%}  ({info['pnl_yen']:+,.0f}円)  [{info['latest_date']}]")
                print(f"       損切りライン: {info['stop_loss_price']:,.0f}円  "
                      f"トレーリング: {info['trailing_stop_price']:,.0f}円（高値{info['peak_price']:,.0f}円から）")
                print()

        # 接近警告
        if warn_alerts:
            for info in warn_alerts:
                print(f"  🟡 【警告】{info['ticker']}  トレーリングストップに接近中")
                print(f"       取得価格: {info['entry_price']:>10,.0f}円  現在値: {info['current_price']:>10,.0f}円")
                print(f"       損益: {info['pnl_pct']:+.1%}  ({info['pnl_yen']:+,.0f}円)  [{info['latest_date']}]")
                print(f"       トレーリングストップ: {info['trailing_stop_price']:,.0f}円（高値{info['peak_price']:,.0f}円から）")
                print()

        # 継続保有
        if hold_ok:
            print(f"  {'銘柄':<10} {'取得価格':>10} {'現在値':>10} {'損益%':>8} {'損益(円)':>12}  {'最終更新'}")
            print("  " + "-" * 65)
            for info in hold_ok:
                print(f"  🟢 {info['ticker']:<8} {info['entry_price']:>10,.0f} {info['current_price']:>10,.0f} "
                      f"{info['pnl_pct']:>+8.1%} {info['pnl_yen']:>+12,.0f}  {info['latest_date']}")

    # ── 2. エントリー候補スキャン ──────────────────────────────
    print(f"\n{'─'*62}")
    print(f"  【新規エントリー候補】（直近営業日に新{high_period}日高値を更新した銘柄）")
    print(f"{'─'*62}\n")

    entry_candidates: List[Dict] = []

    for ticker in scan_tickers:
        if ticker in held_tickers:
            continue  # 既に保有中はスキップ

        df = data.get(ticker)
        if df is None or len(df) < high_period + 1:
            continue

        # 最終営業日のシグナル確認
        latest_close = float(df["Close"].iloc[-1])
        latest_date = df.index[-1]

        # 前N日間（当日を除く）の最高値
        close_series = df["Close"]
        idx = len(close_series) - 1
        if idx < high_period:
            continue
        rolling_high = float(close_series.iloc[idx - high_period: idx].max())

        # 出来高フィルター
        vol_ok = float(df["Volume"].iloc[-1]) > float(df["Volume"].iloc[-2])

        if latest_close > rolling_high and vol_ok:
            prev_high = rolling_high
            breakout_pct = (latest_close - prev_high) / prev_high
            stop_price = latest_close * (1 + stop_loss_pct)

            entry_candidates.append({
                "ticker": ticker,
                "signal_date": latest_date.strftime("%Y/%m/%d"),
                "close": latest_close,
                "prev_high": prev_high,
                "breakout_pct": breakout_pct,
                "stop_price": stop_price,
            })

    if entry_candidates:
        # ブレイク率の高い順にソート
        entry_candidates.sort(key=lambda x: x["breakout_pct"], reverse=True)
        print(f"  {'銘柄':<10} {'終値':>10} {'前高値':>10} {'ブレイク率':>10} {'損切り目安':>12}  {'シグナル日'}")
        print("  " + "-" * 68)
        for c in entry_candidates:
            print(f"  📈 {c['ticker']:<8} {c['close']:>10,.0f} {c['prev_high']:>10,.0f} "
                  f"{c['breakout_pct']:>+10.2%} {c['stop_price']:>12,.0f}  {c['signal_date']}")
        print()
        print("  ※ エントリーは翌営業日の始値が目安です。")
        print(f"  ※ 損切り目安 = 翌日始値の約 {stop_loss_pct:.0%}（実際の取得価格で再計算してください）")
    else:
        print("  本日のエントリー候補はありません。")

    # ── 3. アクションサマリー ──────────────────────────────────
    print(f"\n{'─'*62}")
    print("  【本日のアクションまとめ】")
    print(f"{'─'*62}")

    if exit_alerts:
        for info in exit_alerts:
            label = "損切り" if info["reason"] == "stop_loss" else "トレーリングストップ"
            print(f"  🔴 {info['ticker']}  → 売り（{label}）  目安価格: {info['current_price']:,.0f}円")
    if entry_candidates:
        for c in entry_candidates[:3]:  # 上位3件
            print(f"  📈 {c['ticker']}  → 買い検討（翌日始値）  損切り目安: {c['stop_price']:,.0f}円")
    if not exit_alerts and not entry_candidates:
        print("  本日は特になし。引き続き保有継続。")

    print(f"\n  ポジション追加: python scanner.py add --ticker <コード> --price <取得価格> --shares <株数>")
    print(f"  ポジション削除: python scanner.py remove --ticker <コード>")
    print(f"{'='*62}\n")


# ──────────────────────────────────────────
# CLI
# ──────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="kenmo 新高値ブレイク デイリースキャナー",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    subparsers = parser.add_subparsers(dest="command")

    # scan（デフォルト）
    scan_p = subparsers.add_parser("scan", help="デイリースキャンを実行（デフォルト）")
    scan_p.add_argument("--tickers", nargs="+", default=None)
    scan_p.add_argument("--high-period", type=int, default=260)
    scan_p.add_argument("--stop-loss", type=float, default=-0.08)
    scan_p.add_argument("--trailing", type=float, default=-0.15)
    scan_p.add_argument("--warn-trailing", type=float, default=-0.10)

    # add
    add_p = subparsers.add_parser("add", help="ポジションを追加")
    add_p.add_argument("--ticker", required=True)
    add_p.add_argument("--price", type=float, required=True)
    add_p.add_argument("--shares", type=float, required=True)
    add_p.add_argument("--note", default="")

    # remove
    rm_p = subparsers.add_parser("remove", help="ポジションを削除")
    rm_p.add_argument("--ticker", required=True)

    # list
    subparsers.add_parser("list", help="保有ポジション一覧")

    # トップレベルのオプション（サブコマンドなしでも使える）
    parser.add_argument("--tickers", nargs="+", default=None)
    parser.add_argument("--high-period", type=int, default=260)
    parser.add_argument("--stop-loss", type=float, default=-0.08)
    parser.add_argument("--trailing", type=float, default=-0.15)
    parser.add_argument("--warn-trailing", type=float, default=-0.10)

    args = parser.parse_args()

    if args.command == "add":
        add_position(args.ticker, args.price, args.shares, args.note)
    elif args.command == "remove":
        remove_position(args.ticker)
    elif args.command == "list":
        list_positions()
    else:
        # scan（デフォルト動作）
        tickers = getattr(args, "tickers", None) or DEFAULT_TICKERS
        run_scan(
            scan_tickers=tickers,
            high_period=getattr(args, "high_period", 260),
            stop_loss_pct=getattr(args, "stop_loss", -0.08),
            trailing_stop_pct=getattr(args, "trailing", -0.15),
            warn_trailing_pct=getattr(args, "warn_trailing", -0.10),
        )


if __name__ == "__main__":
    main()
