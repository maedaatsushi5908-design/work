"""
「おにや(o-228)の売買と逆をやれば儲かるのでは？」仮説の検証

結論: 成立しない。理由を数値で確認するためのスクリプト。

前提となる事実（いずれも公開情報）:
  - おにやの主戦場はキオクシアHD (285A) の「信用買い（ロング）」
  - 彼が配信で見せている「爆損」は、値下がりで損した記録ではなく
    “10倍に上がった銘柄をフルレバで持ち続けたことによるドローダウン”
  - つまり方向（ロング）は当たっている。逆張り＝ショートは全損コース。

実行:
    python oniya_inverse_analysis.py
"""

from dataclasses import dataclass
from typing import List, Optional


# ---------------------------------------------------------------------------
# キオクシアHD (285A) の公開株価マイルストーン
# 出所: 各社マーケット情報 / 報道。日次全データは取得できないため節目のみ。
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class PricePoint:
    date: str
    price: float
    note: str


KIOXIA_MILESTONES: List[PricePoint] = [
    PricePoint("2024-12-18", 1_440, "東証プライム上場 初値（公開価格 1,455円）"),
    PricePoint("2026-01-06", 10_945, "2026年 年初来安値"),
    PricePoint("2026-06-19", 103_000, "上場来高値（終値）時価総額56兆円超・国内1位"),
    PricePoint("2026-06-22", 112_700, "2026年 年初来高値（ザラ場）"),
    PricePoint("2026-07-24", 56_010, "直近終値（前日比 -9.49%）"),
]

# おにやが公開している自身のポジション（X / 配信より）
ONIYA_STATED_ENTRY = 53_259.0   # キオクシア 信用ロング 建値
ONIYA_STATED_SHARES = 300       # 300株 ≒ 1,598万円相当

# 日本株の制度信用取引の一般的な条件
MARGIN_RATE = 0.30              # 委託保証金率 30% → 約3.3倍
MAINTENANCE_RATE = 0.20         # 維持率20%割れで追証
LEVERAGE = 1.0 / MARGIN_RATE    # ≒ 3.33倍


def leveraged_return(entry: float, exit_: float, leverage: float, is_long: bool) -> float:
    """レバレッジ込みの証拠金に対するリターン（率）。-1.0 = 証拠金全損。"""
    raw = (exit_ - entry) / entry
    if not is_long:
        raw = -raw
    return raw * leverage


def wipeout_price(entry: float, leverage: float, is_long: bool) -> float:
    """証拠金が全額吹き飛ぶ株価。ショートは上昇方向、ロングは下落方向。"""
    move = 1.0 / leverage
    return entry * (1 - move) if is_long else entry * (1 + move)


def margin_call_price(entry: float, leverage: float, is_long: bool) -> float:
    """追証が発生する株価（維持率がMAINTENANCE_RATEを割る水準）。

    保証金維持率 = (預託保証金 + 評価損益) / 建玉金額
                 = MARGIN_RATE - |値動き率|
    これが MAINTENANCE_RATE を割る = 値動き率が (0.30 - 0.20) = 10% に達した時。
    """
    move = MARGIN_RATE - MAINTENANCE_RATE
    return entry * (1 - move) if is_long else entry * (1 + move)


def fmt_yen(v: float) -> str:
    return f"{v:>15,.0f}円"


def fmt_pct(v: float) -> str:
    return f"{v * 100:>+10.1f}%"


def section(title: str) -> None:
    print()
    print("=" * 78)
    print(f"  {title}")
    print("=" * 78)


def show_price_history() -> None:
    section("1. キオクシアHD (285A) の値動き — おにやの主戦場")
    base = KIOXIA_MILESTONES[0].price
    print(f"{'日付':<12} {'株価':>12} {'上場初値比':>12}   備考")
    print("-" * 78)
    for p in KIOXIA_MILESTONES:
        mult = p.price / base
        print(f"{p.date:<12} {p.price:>10,.0f}円 {mult:>10.1f}倍   {p.note}")

    low = next(p for p in KIOXIA_MILESTONES if p.date == "2026-01-06")
    high = next(p for p in KIOXIA_MILESTONES if p.date == "2026-06-22")
    last = KIOXIA_MILESTONES[-1]

    print()
    print(f"  2026年 安値→高値 : {low.price:,.0f}円 → {high.price:,.0f}円 "
          f"= {high.price / low.price:.1f}倍（約半年）")
    print(f"  高値→直近       : {high.price:,.0f}円 → {last.price:,.0f}円 "
          f"= {(last.price / high.price - 1) * 100:.1f}%")
    print(f"  安値→直近       : {low.price:,.0f}円 → {last.price:,.0f}円 "
          f"= {(last.price / low.price - 1) * 100:+.1f}%  ← 年初来では依然として大幅高")
    print()
    print("  ※ ここが最重要: この銘柄は『上がった』銘柄。彼はそれをロングしている。")


def compare_long_vs_short() -> None:
    section("2. 「おにやの逆」＝ キオクシアのショート をやっていたら？")

    low = next(p for p in KIOXIA_MILESTONES if p.date == "2026-01-06")
    high = next(p for p in KIOXIA_MILESTONES if p.date == "2026-06-22")
    last = KIOXIA_MILESTONES[-1]

    # (ラベル, 建値, 決済値, 保有期間中の最高値, 備考)
    # 保有期間中の最高値を持つのが重要: ショートは途中の高値で強制退場するため、
    # 「建値と決済値だけ」で計算すると損失を大幅に過小評価する。
    cases = [
        ("2026年安値で参入", low.price, high.price, high.price, "年初来高値まで持った場合"),
        ("2026年安値で参入", low.price, last.price, high.price, "直近まで持った場合"),
        ("おにや建値で参入", ONIYA_STATED_ENTRY, high.price, high.price, "年初来高値まで持った場合"),
        ("おにや建値で参入", ONIYA_STATED_ENTRY, last.price, high.price, "直近まで持った場合"),
    ]

    print(f"信用取引 保証金率{MARGIN_RATE:.0%} → レバレッジ約{LEVERAGE:.1f}倍 で計算")
    print("（ショートは保有期間中の高値で強制決済されるものとして経路を考慮）")
    print()
    print(f"{'ケース':<18} {'建値':>10} {'決済':>10} {'おにや(買)':>12} {'逆張り(売)':>12}")
    print("-" * 78)
    for label, entry, exit_, path_high, note in cases:
        long_r = leveraged_return(entry, exit_, LEVERAGE, is_long=True)

        # ショート: 途中の高値が全損ラインを超えていたら、そこで退場
        wo = wipeout_price(entry, LEVERAGE, is_long=False)
        if path_high >= wo:
            short_disp = "全 損"
        else:
            short_disp = fmt_pct(leveraged_return(entry, exit_, LEVERAGE, is_long=False))

        print(f"{label:<18} {entry:>8,.0f}円 {exit_:>8,.0f}円 "
              f"{fmt_pct(long_r):>12} {short_disp:>12}   {note}")

    print()
    print("  → 逆張り（ショート）は全ケースで証拠金が消し飛ぶ。")
    print("    ※ 4行目は建値と決済値だけ見ると -17% で済んだように見えるが、")
    print("      途中で112,700円を通過するため、そこに辿り着く前に強制退場している。")

    print()
    print("  【ショートがいつ死ぬか】")
    for label, entry in [("2026年安値 10,945円で売建", low.price),
                         ("おにや建値 53,259円で売建", ONIYA_STATED_ENTRY)]:
        mc = margin_call_price(entry, LEVERAGE, is_long=False)
        wo = wipeout_price(entry, LEVERAGE, is_long=False)
        print(f"    {label}")
        print(f"      追証発生   : 株価 {mc:>9,.0f}円 到達時")
        print(f"      証拠金全損 : 株価 {wo:>9,.0f}円 到達時")
        print(f"      （実際の年初来高値 {high.price:,.0f}円 は、いずれも遥かに上）")


def show_oniya_actual_position() -> None:
    section("3. おにや本人のポジションは、実は勝っている")

    cost = ONIYA_STATED_ENTRY * ONIYA_STATED_SHARES
    high = next(p for p in KIOXIA_MILESTONES if p.date == "2026-06-22")
    last = KIOXIA_MILESTONES[-1]

    print(f"  公開ポジション : キオクシア 信用ロング {ONIYA_STATED_SHARES}株 "
          f"@ {ONIYA_STATED_ENTRY:,.0f}円")
    print(f"  建玉金額       : {fmt_yen(cost)}")
    print()

    for p in (high, last):
        pnl = (p.price - ONIYA_STATED_ENTRY) * ONIYA_STATED_SHARES
        pct = (p.price / ONIYA_STATED_ENTRY - 1)
        state = "含み益" if pnl >= 0 else "含み損"
        print(f"  {p.date} 株価{p.price:>8,.0f}円 → {state} {fmt_yen(pnl)} "
              f"({fmt_pct(pct)} / レバ込み {fmt_pct(pct * LEVERAGE)})")

    print()
    print("  彼の配信タイトルに並ぶ『-1660万』『-2300万』は、")
    print("  6月の高値から下げた分の“ドローダウン”であって、")
    print("  『安く売って高く買い戻した』類の方向を外した損失ではない。")
    print()
    print("  実際、同じチャンネルには以下の動画も存在する:")
    print("    ・『キオクシア大暴騰により地獄の-800万からプラ転』")
    print("    ・『含み益が360万を超えた件について』")
    print("    ・『株価8万円を突破し国内時価総額1位』【日給500万】")
    print("    ・『株価10万を軽々と突破し、含み益が1000万になっている件』")
    print()
    print("  → 『損している人』ではなく『振れ幅が極端な人』。")
    print("     負けの動画だけが切り取られて拡散しているだけ。")


def show_structural_problems() -> None:
    section("4. 仮に方向が逆でも、この手法が機能しない構造的理由")

    problems = [
        ("情報の遅延",
         "配信で建玉を語る時点で株価は既に動いた後。彼のスタイルは"
         "モメンタム順張りで、1日で±10%動く銘柄。約定時には別の価格。"),
        ("ストップ高/安",
         "キオクシアは値幅制限に張り付く日がある。逆をやりたい方向は"
         "まさに板が消える方向で、そもそも約定しない。"),
        ("売買コストが両側にかかる",
         "手数料・スプレッド・スリッページは買いでも売りでも同じだけ引かれる。"
         "『期待値ゼロ - コスト』の逆は『期待値ゼロ - コスト』であって、"
         "プラスにはならない。損の原因がコストとレバレッジなら、"
         "反転させてもコストは戻ってこない。"),
        ("逆日歩・貸株規制",
         "人気銘柄の空売りは逆日歩（品貸料）が高騰し、"
         "日々コストを垂れ流す。増担保規制・売建停止も入る。"),
        ("サンプル数",
         "株を始めて1〜2年、実質1銘柄への集中投資。独立した試行回数が"
         "統計的に全く足りず、『彼が負ける』という性質自体を検証できない。"),
        ("含み損は決済ではない",
         "逆をやるには建値と決済の“時刻”が必要だが、"
         "彼が公開しているのは含み損益のスナップショットだけ。"),
    ]

    for i, (title, body) in enumerate(problems, 1):
        print(f"  ({i}) {title}")
        for line in _wrap(body, 68):
            print(f"      {line}")
        print()


def _wrap(text: str, width: int) -> List[str]:
    lines, cur = [], ""
    for ch in text:
        cur += ch
        if len(cur) >= width and ch in "。、":
            lines.append(cur)
            cur = ""
    if cur:
        lines.append(cur)
    return lines


def conclusion() -> None:
    section("結論")
    print("""
  「おにやの逆をやれば儲かる」は成立しない。

  最大の理由は単純で、彼は方向を外していないから。
  キオクシアは2026年に 10,945円 → 112,700円 と約10倍になった銘柄で、
  彼はそれを一貫して買っている。その逆＝空売りは、
  どの時点で入っても証拠金が全額吹き飛ぶ。

  彼の「爆損」は、
      方向を外した損失  ではなく
      フルレバレッジ × 一点集中 × 握力 が生む 上下の振れ幅
  である。つまり真似すべきでないのは『銘柄の選択』ではなく
  『ポジションサイズ』の方。そして、サイズの取りすぎは
  反転させても直らない（逆側で同じように吹き飛ぶだけ）。

  一般論として、負けトレーダーの逆張りが機能するのは
  「シグナルが系統的にマイナスの方向を向いている」場合に限られる。
  レバレッジ・手数料・建玉管理で溶かしている人の逆をやっても、
  手に入るのは『期待値ゼロ - 同じコスト』でしかない。
""")


def main() -> None:
    print()
    print("#" * 78)
    print("#  検証: おにや(o-228)の売買と逆をやれば儲かるか？")
    print("#" * 78)

    show_price_history()
    compare_long_vs_short()
    show_oniya_actual_position()
    show_structural_problems()
    conclusion()


if __name__ == "__main__":
    main()
