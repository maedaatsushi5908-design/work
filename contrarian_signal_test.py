"""
「逆神」仮説の検証フレームワーク

ある人物（例: 岐阜暴威）の相場コールを集め、その【逆】を張った場合に
統計的に有意なプラス期待値があるかを検定する。

■ このプログラムの立ち位置
   「逆神の逆張りbot」を作る前に、そもそもエッジが存在するかを確かめる。
   逆張りbotの実装は簡単だが、エッジの検証を飛ばすと
   「ランダムな売買 - 手数料」を高速で実行する機械にしかならない。

■ 使い方
    # 1) 逆神が偶然生まれる確率のシミュレーション（データ不要・まず これを見る）
    python contrarian_signal_test.py --demo

    # 2) 実際のコール履歴を検定する
    python contrarian_signal_test.py --analyze calls.csv

    # 3) エッジ検出に必要なサンプル数を計算
    python contrarian_signal_test.py --power

■ calls.csv のスキーマ
    date,ticker,direction,source,note
    2024-03-15,^N225,long,https://x.com/...,「日経は上がる」と発言
      date      : コール日（この日の終値時点で発言があったとみなす）
      ticker    : 対象銘柄（yfinance形式。日経225は ^N225）
      direction : long / short（本人がどちらを向いていたか）
      source    : 検証可能なURL（必須。これが無い記録は後から捏造と区別できない）
      note      : 発言内容の要約
"""

import argparse
import csv
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd


# ---------------------------------------------------------------------------
# 検証パラメータ
# ---------------------------------------------------------------------------

# 評価する保有期間（営業日）
HORIZONS: List[int] = [1, 5, 10, 20, 60]

# 往復の売買コスト（手数料 + スプレッド + スリッページ）
# 信用売りの場合はここに貸株料・逆日歩が上乗せされる
ROUND_TRIP_COST = 0.002  # 0.2%

# 空売りの継続コスト（貸株料の年率。逆日歩は含まず、発生すればこれ以上）
SHORT_BORROW_ANNUAL = 0.0115  # 1.15%
TRADING_DAYS_PER_YEAR = 245

# ブートストラップ反復回数
N_BOOTSTRAP = 10_000

# 有意水準
ALPHA = 0.05


@dataclass
class Call:
    """検証対象となる1件のコール（発言）。"""
    date: pd.Timestamp
    ticker: str
    direction: str  # "long" or "short"
    source: str
    note: str = ""

    @property
    def sign(self) -> int:
        """本人の向き。long=+1 / short=-1"""
        return 1 if self.direction.lower() in ("long", "buy", "bull", "bullish") else -1


# ---------------------------------------------------------------------------
# コール履歴の読み込み
# ---------------------------------------------------------------------------

def load_calls(path: Path) -> List[Call]:
    """calls.csv を読み込む。出典URLが無い行は棄却する。"""
    calls: List[Call] = []
    rejected = 0

    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(f):
            if not row.get("date", "").strip() or row["date"].lstrip().startswith("#"):
                continue

            source = row.get("source", "").strip()
            if not source or not source.startswith("http"):
                rejected += 1
                continue

            calls.append(Call(
                date=pd.Timestamp(row["date"].strip()),
                ticker=row["ticker"].strip(),
                direction=row["direction"].strip(),
                source=source,
                note=row.get("note", "").strip(),
            ))

    if rejected:
        print(f"  ※ 出典URLが無い {rejected} 行を除外しました "
              f"（検証不能な記録は分析に含めない）")

    return sorted(calls, key=lambda c: c.date)


# ---------------------------------------------------------------------------
# フォワードリターンの計算
# ---------------------------------------------------------------------------

def forward_returns(
    calls: List[Call],
    horizon: int,
    price_data: Dict[str, pd.DataFrame],
) -> Tuple[np.ndarray, List[Call]]:
    """各コールの「翌営業日始値 → horizon営業日後の終値」のリターンを返す。

    発言当日の終値では約定できない（発言を見て動く時点で引け後）ため、
    翌営業日の始値をエントリーとする。
    """
    rets: List[float] = []
    used: List[Call] = []

    for call in calls:
        df = price_data.get(call.ticker)
        if df is None or df.empty:
            continue

        after = df.index[df.index > call.date]
        if len(after) < horizon + 1:
            continue

        entry = float(df.loc[after[0], "Open"])
        exit_ = float(df.loc[after[horizon], "Close"])
        if entry <= 0:
            continue

        rets.append((exit_ - entry) / entry)
        used.append(call)

    return np.array(rets), used


def inverse_pnl(raw_returns: np.ndarray, calls: List[Call], horizon: int) -> np.ndarray:
    """「本人の逆」を張った場合の、コスト控除後のリターン列。"""
    signs = np.array([c.sign for c in calls])

    # 逆を張る = 本人の向きの反対
    gross = -signs * raw_returns

    # 往復コストは方向に関係なく必ず引かれる
    net = gross - ROUND_TRIP_COST

    # 本人がlong → こちらはshort → 貸株料が保有日数分かかる
    is_short_side = signs > 0
    borrow = SHORT_BORROW_ANNUAL * horizon / TRADING_DAYS_PER_YEAR
    net = net - np.where(is_short_side, borrow, 0.0)

    return net


# ---------------------------------------------------------------------------
# 統計（scipy非依存・ブートストラップ）
# ---------------------------------------------------------------------------

def bootstrap_test(x: np.ndarray, n_boot: int = N_BOOTSTRAP,
                   seed: int = 42) -> Dict[str, float]:
    """平均がゼロと異なるかのブートストラップ検定 + 信頼区間。

    リターン分布は正規から大きく外れる（ファットテール）ため、
    t検定よりブートストラップの方が妥当。
    """
    n = len(x)
    if n < 2:
        return {"mean": float("nan"), "p_value": float("nan"),
                "ci_low": float("nan"), "ci_high": float("nan"), "n": n}

    rng = np.random.default_rng(seed)
    observed = float(np.mean(x))

    # 信頼区間: 素のブートストラップ
    idx = rng.integers(0, n, size=(n_boot, n))
    boot_means = x[idx].mean(axis=1)
    ci_low, ci_high = np.percentile(boot_means, [100 * ALPHA / 2, 100 * (1 - ALPHA / 2)])

    # p値: 平均0の帰無仮説へ中心化してから、観測値以上の乖離が出る頻度
    centered = x - observed
    null_means = centered[rng.integers(0, n, size=(n_boot, n))].mean(axis=1)
    p_value = float(np.mean(np.abs(null_means) >= abs(observed)))

    return {
        "mean": observed,
        "p_value": p_value,
        "ci_low": float(ci_low),
        "ci_high": float(ci_high),
        "n": n,
    }


def required_sample_size(effect_size: float, power: float = 0.80) -> int:
    """効果量 d（= 平均/標準偏差）を検出するのに必要なコール数。

    両側 alpha=0.05, 指定power の正規近似。
    """
    z_alpha = 1.959964  # 両側 5%
    z_beta = {0.80: 0.841621, 0.90: 1.281552, 0.95: 1.644854}.get(power, 0.841621)
    if effect_size <= 0:
        return sys.maxsize
    return int(np.ceil(((z_alpha + z_beta) / effect_size) ** 2))


# ---------------------------------------------------------------------------
# 出力
# ---------------------------------------------------------------------------

def section(title: str) -> None:
    print()
    print("=" * 78)
    print(f"  {title}")
    print("=" * 78)


def analyze(calls_path: Path) -> None:
    """実際のコール履歴を検定する。"""
    section("逆神仮説の検定")

    if not calls_path.exists():
        print(f"  コール履歴が見つかりません: {calls_path}")
        print("  calls_template.csv をコピーして実際の発言記録を入れてください。")
        return

    calls = load_calls(calls_path)
    print(f"  読み込んだコール数: {len(calls)}件")

    if len(calls) < 30:
        print()
        print(f"  ⚠  {len(calls)}件では統計的な結論は出せません。")
        print("     後述の必要サンプル数（--power）を確認してください。")

    if not calls:
        return

    print(f"  期間: {calls[0].date.date()} 〜 {calls[-1].date.date()}")

    tickers = sorted({c.ticker for c in calls})
    print(f"  対象銘柄: {len(tickers)}種")

    # 価格データ取得
    from data import download_stock_data, _download_via_yfinance
    start = (calls[0].date - pd.Timedelta(days=30)).strftime("%Y-%m-%d")
    end = (calls[-1].date + pd.Timedelta(days=180)).strftime("%Y-%m-%d")

    # data.py は取得失敗時に無警告で合成データへフォールバックする。
    # 合成データで検定を回すと「それらしい数字」が出てしまい、
    # 実データの結果と見分けがつかなくなるため、ここで明示的に止める。
    if _download_via_yfinance(tickers[0], start, end) is None:
        print()
        print("  " + "!" * 70)
        print("  ! 株価データを取得できません（ネットワーク制限）。")
        print("  !")
        print("  ! data.py は失敗時に合成データへフォールバックしますが、")
        print("  ! 合成データで検定しても出てくるのは乱数の性質であって、")
        print("  ! 本人のコールの性質ではありません。")
        print("  ! 誤った結論を出さないため、ここで中断します。")
        print("  !")
        print("  ! ネットワークが通る環境で再実行してください。")
        print("  " + "!" * 70)
        return

    price_data: Dict[str, pd.DataFrame] = {}
    for t in tickers:
        df = download_stock_data(t, start, end, use_synthetic_fallback=False)
        if df is not None:
            price_data[t] = df

    missing = [t for t in tickers if t not in price_data]
    if missing:
        print(f"  ※ 価格データを取得できなかった銘柄を除外: {', '.join(missing)}")

    print()
    print(f"  往復コスト {ROUND_TRIP_COST:.2%} / 貸株料 年{SHORT_BORROW_ANNUAL:.2%} を控除")
    print()
    print(f"{'保有期間':>8} {'N':>5} {'平均リターン':>12} {'勝率':>8} "
          f"{'95%信頼区間':>22} {'p値':>8} {'判定':>8}")
    print("-" * 78)

    any_significant = False

    for h in HORIZONS:
        raw, used = forward_returns(calls, h, price_data)
        if len(raw) < 2:
            print(f"{h:>6}日 {len(raw):>5}   データ不足")
            continue

        pnl = inverse_pnl(raw, used, h)
        stats = bootstrap_test(pnl)
        win_rate = float(np.mean(pnl > 0))

        significant = stats["p_value"] < ALPHA and stats["mean"] > 0
        if significant:
            any_significant = True

        verdict = "有意" if significant else "－"
        print(f"{h:>6}日 {stats['n']:>5} {stats['mean']:>11.2%} {win_rate:>7.1%} "
              f"[{stats['ci_low']:>+7.2%}, {stats['ci_high']:>+7.2%}] "
              f"{stats['p_value']:>8.3f} {verdict:>8}")

    # 多重検定の補正
    print()
    bonferroni = ALPHA / len(HORIZONS)
    print(f"  ※ {len(HORIZONS)}個の保有期間を同時に試しているため、"
          f"多重検定の補正が必要。")
    print(f"     Bonferroni補正後の有意水準: p < {bonferroni:.3f}")
    print(f"     『どれか1つの期間で p<0.05 が出た』は、ほぼ何も意味しない。")

    print()
    if any_significant:
        print("  → 一部の期間で有意。ただし上記の多重検定補正と、")
        print("     アウトオブサンプル検証（検定に使っていない期間での再現）を")
        print("     通過するまでは実弾を入れないこと。")
    else:
        print("  → 統計的に有意な逆張りエッジは検出されず。")


def demo() -> None:
    """「逆神」が偶然どれだけ生まれるかのシミュレーション。"""
    section("シミュレーション: 「逆神」は偶然どれくらい生まれるか")

    print("""
  設定: 実力ゼロ（期待値0）のトレーダーを1,000人用意し、
        全員にランダムな方向のコールを50回ずつさせる。
        誰も予測能力を持っていない。純粋なコイン投げ。
""")

    rng = np.random.default_rng(2026)
    n_traders, n_calls = 1_000, 50

    # 全員、的中率はちょうど50%（エッジゼロ）
    hits = rng.binomial(n_calls, 0.5, size=n_traders)
    hit_rate = hits / n_calls

    print(f"  最も『当たらなかった』人の的中率 : {hit_rate.min():.1%} "
          f"({hits.min()}/{n_calls}勝)")
    print(f"  最も『当てた』人の的中率         : {hit_rate.max():.1%} "
          f"({hits.max()}/{n_calls}勝)")
    print()

    for threshold in (0.40, 0.35, 0.30):
        n_gods = int(np.sum(hit_rate <= threshold))
        print(f"  的中率{threshold:.0%}以下の『逆神』   : {n_gods:>3}人 "
              f"/ 1,000人 ({n_gods / 10:.1f}%)")

    print(f"""
  → エッジがゼロの集団からでも、これだけ『逆神』が量産される。
     日本の投資クラスタには数万人が発信しており、
     その中で最も外した人が『逆神』として有名になる。
     つまり《逆神という評判は、それ自体では何の証拠にもならない》。

  さらに悪いことに、この選ばれ方は完全に事後的である。
  評判が確立した時点で偶然の外れは使い切られており、
  そこから先の的中率は50%に回帰する（平均への回帰）。
""")

    # 実際に「有名になった逆神」をその後追いかけるとどうなるか
    print("-" * 78)
    print("  では、こうして選ばれた『逆神』の“その後”を追うとどうなるか？")
    print()

    worst_idx = np.argsort(hit_rate)[:20]  # 下位20人を「逆神」として選抜
    future_hits = rng.binomial(n_calls, 0.5, size=len(worst_idx))
    future_rate = future_hits / n_calls

    print(f"  選抜時の的中率（過去）: {hit_rate[worst_idx].mean():.1%}  ← 逆神らしい")
    print(f"  その後の的中率（未来）: {future_rate.mean():.1%}  ← ただの50%")
    print()
    print("  逆張りしても、コストを引いた分だけ負ける。")


def cost_explains_loss() -> None:
    """「生涯収支マイナス」は、方向を外している証拠になるか？"""
    section("生涯収支マイナスは『逆神』の証拠になるか")

    print("""
  岐阜暴威の公開情報: 入金 約5,000万円 / 通算 約-2,500万円（約20年）

  これは確かに実際の負けだが、《方向を外している証拠にはならない》。
  予測能力がちょうどゼロ（コイン投げ）のトレーダーでも、
  売買コストの分だけ、時間をかけて確実に資金を失うからである。

  では、エッジゼロのまま-2,500万円に到達するには何回売買すればよいか:
""")

    target_loss = 25_000_000
    round_trip = ROUND_TRIP_COST
    years = 20

    print(f"{'平均建玉':>12} {'1回あたりコスト':>16} {'必要売買回数':>14} {'年間':>10} {'頻度':>16}")
    print("-" * 78)

    for position in (1_000_000, 3_000_000, 5_000_000, 10_000_000):
        cost_per_trade = position * round_trip
        n_trades = target_loss / cost_per_trade
        per_year = n_trades / years
        per_week = per_year / 52
        if per_week >= 1:
            freq = f"週{per_week:.1f}回"
        else:
            freq = f"月{per_year / 12:.1f}回"
        print(f"{position:>11,}円 {cost_per_trade:>15,.0f}円 "
              f"{n_trades:>12,.0f}回 {per_year:>9,.0f}回 {freq:>16}")

    print(f"""
  → 平均建玉500万円なら、20年で2,500回（週2.4回）売買するだけで
     -2,500万円に到達する。エッジは一切必要ない。
     デイトレ・スイングを続ける個人としてはごく普通の頻度である。

  つまり彼の通算成績は、
      「系統的に逆を向いている」                 ← 逆張りが効く
      「方向はランダムだがコストで削られている」  ← 逆張りは効かない
  のどちらとも整合する。通算損益だけでは 両者を区別できない。

  区別するには、個々のコールの《方向》と《その後の値動き》を
  突き合わせるしかない。それが --analyze の役割。
""")


def power_analysis() -> None:
    """必要サンプル数の計算。"""
    section("エッジ検出に必要なコール数")

    print("""
  「逆を張れば勝てる」と統計的に言うには、何回分のコールが必要か。
  効果量 d = 1コールあたりの平均リターン ÷ その標準偏差。
""")

    print(f"{'1回の平均リターン':>18} {'リターンの標準偏差':>20} {'効果量d':>10} {'必要コール数':>14}")
    print("-" * 78)

    scenarios = [
        (0.005, 0.05, "1回0.5%勝てる（ボラ5%）"),
        (0.010, 0.05, "1回1.0%勝てる（ボラ5%）"),
        (0.020, 0.05, "1回2.0%勝てる（ボラ5%）"),
        (0.010, 0.10, "1回1.0%勝てる（ボラ10%）"),
        (0.030, 0.10, "1回3.0%勝てる（ボラ10%）"),
    ]

    for mean_r, sd_r, label in scenarios:
        d = mean_r / sd_r
        n = required_sample_size(d)
        print(f"{mean_r:>17.1%} {sd_r:>19.1%} {d:>10.2f} {n:>12,}回   {label}")

    print(f"""
  ※ これは「コスト控除後」の数字である点に注意。
     往復コスト {ROUND_TRIP_COST:.1%} を上回るエッジが無ければ、
     どれだけサンプルを集めても勝てない。

  現実的な水準（1回1%、ボラ5%）でも {required_sample_size(0.01 / 0.05):,}回のコールが必要。
  1日1回コールしても約{required_sample_size(0.01 / 0.05) / 245:.1f}年分。
  しかもその間、本人の手法が一定であることが前提になる。
""")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="「逆神」仮説の検証フレームワーク",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--analyze", type=Path, metavar="CSV",
                        help="コール履歴CSVを検定する")
    parser.add_argument("--demo", action="store_true",
                        help="逆神が偶然生まれる確率のシミュレーション")
    parser.add_argument("--power", action="store_true",
                        help="必要サンプル数の計算")
    parser.add_argument("--cost", action="store_true",
                        help="生涯収支マイナスがコストだけで説明できるかの試算")
    args = parser.parse_args()

    if not any([args.analyze, args.demo, args.power, args.cost]):
        args.demo = True
        args.cost = True
        args.power = True

    if args.demo:
        demo()
    if args.cost:
        cost_explains_loss()
    if args.power:
        power_analysis()
    if args.analyze:
        analyze(args.analyze)


if __name__ == "__main__":
    main()
