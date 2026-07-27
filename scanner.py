"""
新高値ブレイク × 移動平均線フィルター デイリースキャナー

毎朝実行することで、その日の「エントリー候補」「損切り対象」を通知する。
エントリー候補はバックテストと同じ移動平均線フィルターを通したものだけが並ぶ。

使い方:
    # 朝のスキャン
    python scanner.py

    # 移動平均フィルターを外して素の新高値ブレイクを見る
    python scanner.py --no-ma-filter

    # スキャン結果を LINE / Discord / メールに通知
    python scanner.py --notify

    # ポジションを追加（買った後に記録）
    python scanner.py add --ticker 7203.T --price 2500 --shares 100

    # ポジションを削除（売った後に記録）
    python scanner.py remove --ticker 7203.T

    # 保有中ポジション一覧
    python scanner.py list
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path
from typing import Dict, List, Optional

# Windows では stdout を UTF-8 に切り替えて絵文字を表示できるようにする
if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import pandas as pd

from config import DEFAULT_TICKERS, StrategyConfig
from strategy import compute_market_regime, compute_signals

POSITIONS_FILE = Path("positions.json")


# ──────────────────────────────────────────
# .env 読み込み（python-dotenv がなくても動く）
# ──────────────────────────────────────────

def _load_dotenv() -> None:
    env_file = Path(".env")
    if not env_file.exists():
        return
    for line in env_file.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        key = key.strip()
        val = val.strip().strip('"').strip("'")
        if key and val and key not in os.environ:
            os.environ[key] = val


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
        print(f"  {p['ticker']:<12} {p['entry_date']:<14} "
              f"{p['entry_price']:>10,.0f} {p['shares']:>8,.0f}  {p.get('note', '')}")


# ──────────────────────────────────────────
# データ取得
# ──────────────────────────────────────────

def _fetch_data(tickers: List[str], min_bars: int) -> Dict[str, pd.DataFrame]:
    """
    直近データを取得する。

    min_bars は判定に必要な営業日数。営業日 → 暦日の換算（約1.5倍）に
    余裕を足して取得期間を決める。
    """
    end = date.today()
    start = end - timedelta(days=int(min_bars * 1.5) + 60)

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

    test = try_yf(tickers[0])
    if test is None:
        print("  ⚠️  ネットワーク接続なし → 合成データで動作します（デモモード）")
        from synthetic_data import generate_benchmark, generate_stock_data
        return {
            t: (generate_benchmark(start.isoformat(), end.isoformat())
                if t.startswith("^")
                else generate_stock_data(t, start.isoformat(), end.isoformat()))
            for t in tickers
        }

    result = {tickers[0]: test}
    for t in tickers[1:]:
        df = try_yf(t)
        if df is not None:
            result[t] = df
    return result


# ──────────────────────────────────────────
# スキャン結果データクラス
# ──────────────────────────────────────────

EXIT_REASON_LABELS = {
    "stop_loss": "損切りライン到達",
    "trailing_stop": "トレーリングストップ到達",
    "ma_exit": "移動平均線割れ",
}

EXIT_REASON_SHORT = {
    "stop_loss": "損切り",
    "trailing_stop": "トレーリングST",
    "ma_exit": "MA割れ",
}


@dataclass
class ScanResult:
    today: date
    exit_alerts: List[Dict] = field(default_factory=list)   # 売り推奨
    warn_alerts: List[Dict] = field(default_factory=list)   # トレーリング接近警告
    hold_ok: List[Dict] = field(default_factory=list)       # 継続保有
    entry_candidates: List[Dict] = field(default_factory=list)  # エントリー候補
    filtered_out: List[Dict] = field(default_factory=list)  # 新高値だがMAフィルター不通過
    market_ok: Optional[bool] = None                        # 地合い（None = 判定なし）
    market_note: str = ""                                   # 地合いの説明

    def has_action(self) -> bool:
        return bool(self.exit_alerts or self.entry_candidates)

    def to_console(self) -> str:
        """コンソール表示用テキスト（絵文字・罫線あり）"""
        lines = []
        w = 62
        lines.append(f"\n{'='*w}")
        lines.append(f"  新高値ブレイク × MAフィルター デイリースキャン  {self.today.strftime('%Y年%m月%d日')}")
        lines.append(f"{'='*w}")

        # 地合い
        if self.market_ok is not None:
            mark = "🟢 良好" if self.market_ok else "🔴 悪化"
            lines.append(f"\n  【地合い】{mark}  {self.market_note}")
            if not self.market_ok:
                lines.append("  ※ 指数が長期移動平均を下回っています。新規エントリーは見送り推奨。")

        # 保有ポジション
        lines.append(f"\n{'─'*w}")
        lines.append("  【保有ポジション状況】")
        lines.append(f"{'─'*w}")

        if not (self.exit_alerts or self.warn_alerts or self.hold_ok):
            lines.append("  保有中のポジションはありません。")
        else:
            for info in self.exit_alerts:
                reason_label = EXIT_REASON_LABELS.get(info["reason"], info["reason"])
                lines.append(f"\n  🔴 【売り推奨】{info['ticker']}  {reason_label}")
                lines.append(f"       取得価格: {info['entry_price']:>10,.0f}円  現在値: {info['current_price']:>10,.0f}円")
                lines.append(f"       損益: {info['pnl_pct']:+.1%}  ({info['pnl_yen']:+,.0f}円)  [{info['latest_date']}]")
                lines.append(f"       損切りライン: {info['stop_loss_price']:,.0f}円  "
                             f"トレーリング: {info['trailing_stop_price']:,.0f}円")
                if info.get("ma_exit_line"):
                    lines.append(f"       移動平均線: {info['ma_exit_line']:,.0f}円")

            for info in self.warn_alerts:
                lines.append(f"\n  🟡 【警告】{info['ticker']}  トレーリングストップに接近中")
                lines.append(f"       取得価格: {info['entry_price']:>10,.0f}円  現在値: {info['current_price']:>10,.0f}円")
                lines.append(f"       損益: {info['pnl_pct']:+.1%}  ({info['pnl_yen']:+,.0f}円)  [{info['latest_date']}]")
                lines.append(f"       トレーリングストップ: {info['trailing_stop_price']:,.0f}円（高値 {info['peak_price']:,.0f}円から）")

            if self.hold_ok:
                lines.append(f"\n  {'銘柄':<10} {'取得価格':>10} {'現在値':>10} {'損益%':>8} {'損益(円)':>12}  最終更新")
                lines.append("  " + "-" * 65)
                for info in self.hold_ok:
                    lines.append(f"  🟢 {info['ticker']:<8} {info['entry_price']:>10,.0f} "
                                 f"{info['current_price']:>10,.0f} {info['pnl_pct']:>+8.1%} "
                                 f"{info['pnl_yen']:>+12,.0f}  {info['latest_date']}")

        # エントリー候補
        lines.append(f"\n{'─'*w}")
        lines.append(f"  【新規エントリー候補】")
        lines.append(f"{'─'*w}\n")

        if self.entry_candidates:
            lines.append(f"  {'銘柄':<10} {'終値':>10} {'前高値':>10} {'ブレイク率':>10} {'損切り目安':>12}  シグナル日")
            lines.append("  " + "-" * 68)
            for c in self.entry_candidates:
                lines.append(f"  📈 {c['ticker']:<8} {c['close']:>10,.0f} {c['prev_high']:>10,.0f} "
                             f"{c['breakout_pct']:>+10.2%} {c['stop_price']:>12,.0f}  {c['signal_date']}")
                if c.get("ma_short"):
                    lines.append(f"       MA: 短期{c['ma_short']:,.0f} > 中期{c['ma_mid']:,.0f} "
                                 f"> 長期{c['ma_long']:,.0f}  ✓トレンド良好")
            lines.append("\n  ※ エントリーは翌営業日の始値が目安です。")
        else:
            lines.append("  本日のエントリー候補はありません。")

        # 移動平均フィルターで除外された銘柄（判断材料として表示）
        if self.filtered_out:
            lines.append(f"\n  【参考】新高値だがMAフィルター不通過（見送り）: {len(self.filtered_out)}銘柄")
            for c in self.filtered_out[:5]:
                lines.append(f"     ⚪ {c['ticker']}  終値{c['close']:,.0f}円  理由: {c['reason']}")

        # アクションまとめ
        lines.append(f"\n{'─'*w}")
        lines.append("  【本日のアクションまとめ】")
        lines.append(f"{'─'*w}")
        if self.exit_alerts:
            for info in self.exit_alerts:
                label = EXIT_REASON_SHORT.get(info["reason"], info["reason"])
                lines.append(f"  🔴 {info['ticker']} → 売り（{label}）  目安: {info['current_price']:,.0f}円")
        if self.entry_candidates:
            for c in self.entry_candidates[:3]:
                lines.append(f"  📈 {c['ticker']} → 買い検討（翌日始値）  損切り目安: {c['stop_price']:,.0f}円")
        if not self.exit_alerts and not self.entry_candidates:
            lines.append("  本日は特になし。引き続き保有継続。")

        lines.append(f"\n{'='*w}\n")
        return "\n".join(lines)

    def to_notify_message(self) -> str:
        """通知用の簡潔なテキスト（LINEなど向け）"""
        today_str = self.today.strftime("%Y/%m/%d")
        lines = [f"【新高値ブレイク×MAフィルター】{today_str}"]

        if self.market_ok is not None:
            lines.append(f"地合い: {'🟢 良好' if self.market_ok else '🔴 悪化（新規見送り推奨）'}")

        if self.exit_alerts:
            lines.append("\n🔴 売り推奨")
            for info in self.exit_alerts:
                label = EXIT_REASON_SHORT.get(info["reason"], info["reason"])
                lines.append(f"  {info['ticker']}  {label}  {info['pnl_pct']:+.1%}"
                             f"  ({info['pnl_yen']:+,.0f}円)")

        if self.warn_alerts:
            lines.append("\n🟡 トレーリングST接近")
            for info in self.warn_alerts:
                lines.append(f"  {info['ticker']}  損益{info['pnl_pct']:+.1%}")

        if self.hold_ok:
            lines.append("\n🟢 保有継続")
            for info in self.hold_ok:
                lines.append(f"  {info['ticker']}  {info['pnl_pct']:+.1%}  ({info['pnl_yen']:+,.0f}円)")

        if self.entry_candidates:
            lines.append("\n📈 エントリー候補（翌日始値）")
            for c in self.entry_candidates:
                lines.append(f"  {c['ticker']}  終値{c['close']:,.0f}円"
                             f"  ブレイク{c['breakout_pct']:+.2%}"
                             f"  損切り目安{c['stop_price']:,.0f}円")

        if not self.exit_alerts and not self.entry_candidates and not self.warn_alerts:
            lines.append("\n本日は特になし。")

        return "\n".join(lines)


# ──────────────────────────────────────────
# スキャン本体
# ──────────────────────────────────────────

def _ma_reject_reason(row, cfg: StrategyConfig) -> str:
    """新高値ブレイクがMAフィルターで弾かれた理由を短い日本語で返す。"""
    if not cfg.ma_filter:
        return "フィルター条件不一致"

    close = row["Close"]
    ma_short, ma_mid, ma_long = row["ma_short"], row["ma_mid"], row["ma_long"]

    if pd.isna(ma_long) or pd.isna(ma_mid) or pd.isna(ma_short):
        return "移動平均の計算に必要なデータが不足"
    if cfg.ma_require_above and not (close > ma_short and close > ma_mid and close > ma_long):
        return "終値が移動平均線を下回っている"
    if cfg.ma_require_alignment and not (ma_short > ma_mid > ma_long):
        return "パーフェクトオーダーでない（MAの並びが崩れている）"
    if cfg.ma_require_slope_up:
        return f"長期MA（{cfg.ma_long_period}日）が上向きでない"
    return "MAフィルター不通過"


def run_scan(
    scan_tickers: List[str],
    cfg: StrategyConfig,
    warn_trailing_pct: float,
    notify: bool = False,
) -> ScanResult:
    stop_loss_pct = cfg.stop_loss_pct
    trailing_stop_pct = cfg.trailing_stop_pct

    result = ScanResult(today=date.today())
    positions = load_positions()
    held_tickers = [p["ticker"] for p in positions]
    all_tickers = list(dict.fromkeys(scan_tickers + held_tickers))

    # 地合い判定に使う指数もまとめて取得する
    fetch_tickers = list(all_tickers)
    if cfg.market_filter and cfg.market_ticker not in fetch_tickers:
        fetch_tickers.append(cfg.market_ticker)

    print(f"  データ取得中... ({len(fetch_tickers)}銘柄)")
    data = _fetch_data(fetch_tickers, cfg.min_bars_required)
    if not data:
        print("  データを取得できませんでした。")
        return result

    # ── 地合い判定（指数が長期移動平均より上か）─────────────
    if cfg.market_filter:
        index_df = data.get(cfg.market_ticker)
        regime = compute_market_regime(index_df, cfg)
        if regime is not None and len(regime) > 0:
            result.market_ok = bool(regime.iloc[-1])
            index_close = float(index_df["Close"].iloc[-1])
            index_ma = float(index_df["Close"].rolling(cfg.market_ma_period).mean().iloc[-1])
            result.market_note = (
                f"{cfg.market_ticker} {index_close:,.0f} / "
                f"{cfg.market_ma_period}日MA {index_ma:,.0f}"
            )
        else:
            print("  ⚠️  指数データが不足しているため地合い判定はスキップします。")

    # ── 保有ポジションのチェック ─────────────────────────────
    for p in positions:
        ticker = p["ticker"]
        entry_price = p["entry_price"]
        shares = p["shares"]
        entry_date = p["entry_date"]

        df = data.get(ticker)
        if df is None or df.empty:
            continue

        current_price = float(df["Close"].iloc[-1])
        latest_date = df.index[-1].strftime("%Y/%m/%d")

        entry_dt = pd.Timestamp(entry_date)
        df_since_entry = df[df.index >= entry_dt]
        peak_price = float(df_since_entry["High"].max()) if not df_since_entry.empty else entry_price

        pnl_pct = (current_price - entry_price) / entry_price
        stop_loss_price = entry_price * (1 + stop_loss_pct)
        trailing_stop_price = peak_price * (1 + trailing_stop_pct)
        warn_trailing_price = peak_price * (1 + warn_trailing_pct)

        # 移動平均割れの判定（バックテストと同じロジック）
        ma_exit_hit = False
        ma_exit_line = None
        if cfg.ma_exit and len(df) >= cfg.ma_exit_period + cfg.ma_exit_confirm_days:
            sig = compute_signals(df, cfg)
            ma_exit_hit = bool(sig["ma_exit_signal"].iloc[-1])
            ma_line = sig["ma_exit_line"].iloc[-1]
            ma_exit_line = None if pd.isna(ma_line) else float(ma_line)

        info = {
            "ticker": ticker,
            "entry_price": entry_price,
            "current_price": current_price,
            "peak_price": peak_price,
            "stop_loss_price": stop_loss_price,
            "trailing_stop_price": trailing_stop_price,
            "warn_trailing_price": warn_trailing_price,
            "pnl_pct": pnl_pct,
            "pnl_yen": (current_price - entry_price) * shares,
            "current_value": current_price * shares,
            "shares": shares,
            "entry_date": entry_date,
            "latest_date": latest_date,
            "ma_exit_line": ma_exit_line,
        }

        if current_price <= stop_loss_price:
            result.exit_alerts.append({**info, "reason": "stop_loss"})
        elif current_price <= trailing_stop_price:
            result.exit_alerts.append({**info, "reason": "trailing_stop"})
        elif ma_exit_hit:
            result.exit_alerts.append({**info, "reason": "ma_exit"})
        elif current_price <= warn_trailing_price:
            result.warn_alerts.append(info)
        else:
            result.hold_ok.append(info)

    # ── エントリー候補スキャン ────────────────────────────────
    # バックテストと同じ compute_signals を使うことで、
    # 「バックテストでは買っていたのにスキャナーには出ない」というズレを防ぐ。
    for ticker in scan_tickers:
        if ticker in held_tickers:
            continue

        df = data.get(ticker)
        if df is None or len(df) < cfg.high_period + 1:
            continue

        sig = compute_signals(df, cfg)
        last = sig.iloc[-1]

        if not bool(last["new_high"]) or not bool(last["volume_ok"]):
            continue

        latest_close = float(last["Close"])
        rolling_high = float(last["rolling_high"])
        breakout_pct = (latest_close - rolling_high) / rolling_high

        candidate = {
            "ticker": ticker,
            "signal_date": df.index[-1].strftime("%Y/%m/%d"),
            "close": latest_close,
            "prev_high": rolling_high,
            "breakout_pct": breakout_pct,
            "stop_price": latest_close * (1 + stop_loss_pct),
        }

        # 移動平均フィルターを通過したものだけをエントリー候補にする
        if bool(last["signal"]):
            if cfg.ma_filter:
                candidate.update({
                    "ma_short": float(last["ma_short"]),
                    "ma_mid": float(last["ma_mid"]),
                    "ma_long": float(last["ma_long"]),
                })
            result.entry_candidates.append(candidate)
        else:
            candidate["reason"] = _ma_reject_reason(last, cfg)
            result.filtered_out.append(candidate)

    result.entry_candidates.sort(key=lambda x: x["breakout_pct"], reverse=True)

    # ── 出力 ──────────────────────────────────────────────────
    print(result.to_console())

    if notify:
        from notify import send_all
        msg = result.to_notify_message()
        print("  通知を送信中...")
        send_all(msg)

    return result


# ──────────────────────────────────────────
# CLI
# ──────────────────────────────────────────

def main():
    _load_dotenv()

    defaults = StrategyConfig()

    parser = argparse.ArgumentParser(
        description="新高値ブレイク × 移動平均線フィルター デイリースキャナー",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    subparsers = parser.add_subparsers(dest="command")

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

    # トップレベル（サブコマンドなし = scan）
    # デフォルト値は config.py の StrategyConfig と揃えている
    parser.add_argument("--tickers", nargs="+", default=None)
    parser.add_argument("--high-period", type=int, default=defaults.high_period)
    parser.add_argument("--stop-loss", type=float, default=defaults.stop_loss_pct)
    parser.add_argument("--trailing", type=float, default=defaults.trailing_stop_pct)
    parser.add_argument("--warn-trailing", type=float, default=-0.10)
    parser.add_argument("--notify", action="store_true", help="スキャン後に通知を送る")

    ma = parser.add_argument_group("移動平均線フィルター")
    ma.add_argument("--ma-short", type=int, default=defaults.ma_short_period, help="短期移動平均（営業日）")
    ma.add_argument("--ma-mid", type=int, default=defaults.ma_mid_period, help="中期移動平均（営業日）")
    ma.add_argument("--ma-long", type=int, default=defaults.ma_long_period, help="長期移動平均（営業日）")
    ma.add_argument("--no-ma-filter", action="store_true", help="移動平均フィルターを無効にする")
    ma.add_argument("--no-alignment", action="store_true", help="パーフェクトオーダー条件を外す")
    ma.add_argument("--ma-exit", action="store_true", help="MA割れを売り推奨に含める")
    ma.add_argument("--ma-exit-period", type=int, default=defaults.ma_exit_period, help="MA割れ判定の移動平均")
    ma.add_argument("--ma-exit-days", type=int, default=defaults.ma_exit_confirm_days, help="MA割れ確認日数")
    ma.add_argument("--no-market-filter", action="store_true", help="地合い判定をしない")
    ma.add_argument("--market-ma", type=int, default=defaults.market_ma_period, help="地合い判定の移動平均")

    args = parser.parse_args()

    if args.command == "add":
        add_position(args.ticker, args.price, args.shares, args.note)
    elif args.command == "remove":
        remove_position(args.ticker)
    elif args.command == "list":
        list_positions()
    else:
        cfg = StrategyConfig(
            high_period=args.high_period,
            stop_loss_pct=args.stop_loss,
            trailing_stop_pct=args.trailing,
            ma_filter=not args.no_ma_filter,
            ma_short_period=args.ma_short,
            ma_mid_period=args.ma_mid,
            ma_long_period=args.ma_long,
            ma_require_alignment=not args.no_alignment,
            ma_exit=args.ma_exit,
            ma_exit_period=args.ma_exit_period,
            ma_exit_confirm_days=args.ma_exit_days,
            market_filter=not args.no_market_filter,
            market_ma_period=args.market_ma,
        )
        run_scan(
            scan_tickers=args.tickers or DEFAULT_TICKERS,
            cfg=cfg,
            warn_trailing_pct=args.warn_trailing,
            notify=args.notify,
        )


if __name__ == "__main__":
    main()
