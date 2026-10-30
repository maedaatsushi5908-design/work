"""週の 2 営業日目の朝（通常は 月曜引け→火曜寄り）の夜間リターンを記録する。

バックテスト (f:/vscode/jp_anomalies/tuesday_overnight.py) で 2020-2026 に
+22bp/回 (t=4.4) だったが、2019 年以前は逆にマイナス。今後も続くかを前向きに確認する。

毎回 2020 年以降を yfinance から再取得して log.csv を作り直し（冪等）、
FORWARD_START 以降だけの成績を summary.md に書く。
"""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np
import pandas as pd
import yfinance as yf

if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")

HERE = Path(__file__).parent
TICKERS = {"1321.T": "日経225", "1308.T": "TOPIX"}
FORWARD_START = "2026-11-02"   # 検証を始めた後の最初の週
COST_BP = 4                    # 往復コスト想定 (bp)


def overnight(ticker: str) -> pd.DataFrame:
    d = yf.Ticker(ticker).history(start="2020-01-01", auto_adjust=False)
    d.index = d.index.tz_localize(None).normalize()
    d = d[d["Volume"] > 0]
    prev_close = d["Close"].shift(1)
    ret = np.log((d["Open"] + d["Dividends"]) / prev_close)
    nth = pd.Series(1, index=d.index).groupby(d.index.to_period("W-SUN")).cumsum()
    out = pd.DataFrame({
        "date": d.index.strftime("%Y-%m-%d"),
        "ticker": ticker,
        "prev_date": pd.Series(d.index, index=d.index).shift(1).dt.strftime("%Y-%m-%d"),
        "prev_close": prev_close.round(2),
        "open": d["Open"].round(2),
        "ret_bp": (np.expm1(ret) * 1e4).round(1),
        "nth_day": nth,
    })
    out = out[(out["nth_day"] == 2) & out["ret_bp"].notna()]
    # 分割などデータ誤りの除外
    return out[out["ret_bp"].abs() < 1500]


def stats(r: pd.Series) -> dict:
    net = r - COST_BP
    n = len(net)
    if n == 0:
        return {"回数": 0, "平均(bp,コスト後)": "-", "t値": "-", "勝率": "-", "累積(%)": "-"}
    t = net.mean() / net.std(ddof=1) * np.sqrt(n) if n > 2 else float("nan")
    return {"回数": n, "平均(bp,コスト後)": round(net.mean(), 1) if n else float("nan"),
            "t値": round(t, 2), "勝率": f"{(net > 0).mean():.0%}" if n else "-",
            "累積(%)": round((np.prod(1 + net / 1e4) - 1) * 100, 2) if n else 0.0}


def main():
    log = pd.concat([overnight(t) for t in TICKERS], ignore_index=True)
    log.to_csv(HERE / "log.csv", index=False)

    lines = ["# 週明け夜間リターン 前向き検証", "",
             f"週の2営業日目の朝（通常は月曜引け→火曜寄り）に ETF を持った場合の成績。往復コスト {COST_BP}bp 控除後。",
             "",
             "バックテスト（2020-2026）: 1321 で平均 +22bp/回、t=4.4、勝率61%。2019年以前は平均 -12bp（逆効果）。",
             ""]
    for label, start in [(f"前向き検証（{FORWARD_START}〜）", FORWARD_START), ("参考: 2020年〜", "2020-01-01")]:
        lines += [f"## {label}", "", "| ETF | 回数 | 平均(bp) | t値 | 勝率 | 累積(%) |", "|---|---|---|---|---|---|"]
        for t, name in TICKERS.items():
            s = stats(log[(log["ticker"] == t) & (log["date"] >= start)]["ret_bp"])
            lines.append(f"| {name} ({t}) | {s['回数']} | {s['平均(bp,コスト後)']} | {s['t値']} | {s['勝率']} | {s['累積(%)']} |")
        lines.append("")
    fwd = log[(log["ticker"] == "1321.T") & (log["date"] >= FORWARD_START)]
    if len(fwd):
        lines += ["## 前向き検証の各回 (1321)", "", "| 日付 | 前日終値 | 始値 | リターン(bp) |", "|---|---|---|---|"]
        for r in fwd.itertuples():
            lines.append(f"| {r.date} | {r.prev_close} | {r.open} | {r.ret_bp:+.1f} |")
    (HERE / "summary.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("\n".join(lines))


if __name__ == "__main__":
    main()
