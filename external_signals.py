"""
外部データによる業績サプライズ推定モジュール

会社が公表する決算情報（短信・会社予想）を一切使わず、
企業の「周り」で観測できる公開データだけから業績の上振れ／下振れを推定する。

【考え方】
企業の売上・利益は必ず外界に痕跡を残す。物が売れれば運ばれ、材料を買い、
市況が動く。その痕跡は決算発表よりも先に観測できる。

  ドル円が前年同期より15%円安  → 輸出企業の営業利益は機械的に上振れる
  SOX指数・ASMLの受注が急回復  → 半導体製造装置の受注は数ヶ月遅れて追随する
  銅・金価格が高騰             → 非鉄の採掘マージンは直接的に改善する
  コンテナ運賃が2倍           → 海運の利益はほぼ運賃で決まる

これらは全て公開・観測可能な情報なので、利用は合法である。
（違法になるのは企業関係者から「未公表の決算数値そのもの」を得る場合のみ）

【重要な注意】
本モジュールが出すスコアは「業績が上振れる方向に外部環境が動いたか」を測るもので、
株価がそれを既に織り込んでいるかどうかは測っていない。実際に儲かるかどうかは
「推定精度 × 市場の織り込み度合い」で決まる。必ず validate モードで
情報係数（IC）を確認してから使うこと。

使い方:
    # 現時点の外部環境スコアで全銘柄をランキング
    python external_signals.py rank
    python external_signals.py rank --mode fitted   # 銘柄固有の感応度を使う

    # スコアの内訳を銘柄ごとに表示
    python external_signals.py explain --ticker 8035.T --mode fitted

    # この手法が実際に効くかを過去データで検証（情報係数を測る）
    python external_signals.py validate --start 2016-01-01 --end 2024-12-31

    # 利用可能な外部データ系列とセクター対応表を表示
    python external_signals.py sources

    # 政府統計（e-Stat）の接続設定
    python external_signals.py estat-init
    python external_signals.py estat-search --keyword 建設工事受注動態統計調査

新高値ブレイク戦略のフィルターとして使う場合は main.py 側から:
    python main.py --external-filter --external-min-score 0.0
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import numpy as np
import pandas as pd

from config import DEFAULT_TICKERS
from data import download_stock_data


def _load_dotenv() -> None:
    """.env を環境変数に読み込む（python-dotenv がなくても動く）。"""
    env_file = Path(".env")
    if not env_file.exists():
        return
    for line in env_file.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        key, val = key.strip(), val.strip().strip('"').strip("'")
        if key and val and key not in os.environ:
            os.environ[key] = val


_load_dotenv()


# ──────────────────────────────────────────
# 外部データ系列の定義
#
# すべて yfinance で取得できる = APIキー不要・無料。
# e-Stat（政府統計）系は ESTAT_APP_ID が設定されている場合のみ有効になる。
# ──────────────────────────────────────────

SERIES_LABELS: Dict[str, str] = {
    "JPY=X": "ドル円（上昇=円安=輸出企業にプラス）",
    "^SOX": "フィラデルフィア半導体指数",
    "ASML": "ASML（半導体露光装置・世界シェア首位）",
    "TSM": "TSMC（半導体受託製造・設備投資の源流）",
    "AMAT": "アプライドマテリアルズ（前工程装置）",
    "CL=F": "WTI原油先物",
    "HG=F": "銅先物（世界景気・非鉄マージン）",
    "GC=F": "金先物",
    "ZIM": "ZIM（コンテナ海運・運賃の代理変数）",
    "^TNX": "米10年国債利回り（銀行の利ざや代理変数）",
    "^N225": "日経225（市場全体の水準）",
    # 以下は e-Stat 接続時のみ有効（estat_series.json に statsDataId を設定する）
    "estat:construction_orders": "建設工事受注動態統計（国交省・月次）",
    "estat:retail_sales": "商業動態統計 小売業販売額（経産省・月次）",
    "estat:machinery_orders": "機械受注統計（内閣府・月次）",
}

# e-Stat 由来（月次）の系列かどうか。前年同期比の計算方法が日次と異なる。
ESTAT_PREFIX = "estat:"


@dataclass(frozen=True)
class Driver:
    """ある業種の業績を説明する外部変数ひとつ分の定義。"""

    series: str      # SERIES_LABELS のキー
    sign: int        # +1 = その系列が上がると業績にプラス / -1 = マイナス
    weight: float    # セクター内での重み（合計1.0になるようにする）


# 業種 → 外部ドライバーの対応
#
# 重みは「その業種の利益がどれだけその変数で説明されるか」の主観的な事前分布。
# validate モードで IC を測り、効いていない変数は重みを下げるか外すこと。
SECTOR_DRIVERS: Dict[str, List[Driver]] = {
    # 半導体・製造装置: 海外の設備投資サイクルがそのまま受注になる
    "semiconductor": [
        Driver("^SOX", +1, 0.35),
        Driver("ASML", +1, 0.20),
        Driver("TSM", +1, 0.15),
        Driver("AMAT", +1, 0.10),
        Driver("JPY=X", +1, 0.20),
    ],
    # 電機・重電: 為替感応度が高く、半導体サイクルの影響も受ける
    "electronics": [
        Driver("JPY=X", +1, 0.40),
        Driver("^SOX", +1, 0.30),
        Driver("^N225", +1, 0.30),
    ],
    # 自動車・部品: 為替がほぼ全て。原油高はコスト要因
    "auto": [
        Driver("JPY=X", +1, 0.60),
        Driver("^N225", +1, 0.20),
        Driver("CL=F", -1, 0.20),
    ],
    # 精密・光学: 輸出比率が高い
    "precision": [
        Driver("JPY=X", +1, 0.50),
        Driver("^N225", +1, 0.30),
        Driver("^SOX", +1, 0.20),
    ],
    # 機械: 設備投資サイクル。機械受注統計が本命
    "machinery": [
        Driver("estat:machinery_orders", +1, 0.35),
        Driver("^N225", +1, 0.25),
        Driver("JPY=X", +1, 0.20),
        Driver("HG=F", +1, 0.20),
    ],
    # エネルギー: 原油価格が利益をほぼ決める
    "energy": [
        Driver("CL=F", +1, 0.70),
        Driver("JPY=X", +1, 0.30),
    ],
    # 非鉄・鉄鋼: 市況がそのまま採掘・製錬マージンになる
    "materials": [
        Driver("HG=F", +1, 0.45),
        Driver("GC=F", +1, 0.25),
        Driver("JPY=X", +1, 0.30),
    ],
    # 海運: 運賃指数で利益がほぼ決まる。燃料費はコスト
    "shipping": [
        Driver("ZIM", +1, 0.55),
        Driver("JPY=X", +1, 0.25),
        Driver("CL=F", -1, 0.20),
    ],
    # 銀行・証券: 金利と市場水準
    "financial": [
        Driver("^TNX", +1, 0.50),
        Driver("^N225", +1, 0.50),
    ],
    # 商社: 資源市況と為替の合成
    "trading": [
        Driver("CL=F", +1, 0.30),
        Driver("HG=F", +1, 0.30),
        Driver("JPY=X", +1, 0.40),
    ],
    # 建設: 公共工事の受注統計が本命。未接続時は市場全体で代用（説明力は低い）
    "construction": [
        Driver("estat:construction_orders", +1, 0.70),
        Driver("^N225", +1, 0.30),
    ],
    # 小売・消費: 商業動態統計が本命。円安は仕入コスト要因なので符号はマイナス
    "retail": [
        Driver("estat:retail_sales", +1, 0.55),
        Driver("^N225", +1, 0.25),
        Driver("JPY=X", -1, 0.20),
    ],
    # 内需サービス・IT: 外部市況との連動が弱い（＝この手法が効きにくい）
    "domestic": [
        Driver("^N225", +1, 1.00),
    ],
}


# 銘柄 → 業種の対応（config.CURATED_TICKERS の116銘柄をカバー）
TICKER_SECTOR: Dict[str, str] = {}


def _assign(sector: str, tickers: List[str]) -> None:
    for t in tickers:
        TICKER_SECTOR[t] = sector


_assign("semiconductor", [
    "8035.T", "7735.T", "6963.T", "6762.T", "6981.T", "6971.T", "6298.T",
])
_assign("electronics", [
    "6501.T", "6702.T", "6701.T", "6645.T", "6758.T", "6503.T", "6861.T",
    "6507.T", "6504.T", "6919.T", "1948.T", "1992.T",
])
_assign("auto", ["7203.T", "6479.T", "5989.T", "8117.T"])
_assign("precision", ["7741.T", "7751.T", "7729.T", "7726.T", "7743.T"])
_assign("machinery", [
    "6361.T", "6367.T", "6339.T", "6151.T", "6335.T", "7841.T", "1964.T",
])
_assign("energy", ["5019.T"])
_assign("materials", ["5411.T", "5713.T", "7305.T"])
_assign("shipping", ["9104.T"])
_assign("financial", [
    "8306.T", "8316.T", "8411.T", "8511.T", "8708.T", "8707.T", "8591.T",
])
_assign("trading", ["8002.T", "8130.T"])
_assign("retail", [
    "8218.T", "8242.T", "8267.T", "7565.T", "7483.T", "9983.T", "7846.T",
    "7974.T", "2802.T", "2009.T", "2002.T", "2060.T", "7989.T",
])
_assign("domestic", [
    "6098.T", "6134.T", "6533.T", "6539.T", "6549.T", "9433.T", "9413.T",
    "2763.T", "8830.T", "8935.T", "1766.T", "1776.T",
])
_assign("construction", [
    "1965.T", "1966.T", "1952.T", "1898.T", "1879.T", "1828.T", "1959.T",
    "1770.T", "1938.T", "1975.T", "1719.T", "1812.T", "1944.T", "1981.T",
    "1928.T", "1967.T", "1997.T", "1802.T", "1888.T", "1736.T", "1807.T",
    "1801.T", "1758.T", "1787.T", "1911.T", "1808.T", "1870.T", "1942.T",
    "1803.T", "1979.T", "1994.T", "1982.T", "1721.T", "1941.T", "1968.T",
    "1963.T", "1813.T", "1946.T", "1925.T", "1926.T", "1893.T", "1847.T",
])

DEFAULT_SECTOR = "domestic"

# 市場データ（yfinance系列）だけでは説明力が乏しい業種。
# 建設・小売・機械は e-Stat の政府統計を繋ぐと信頼できるようになる。
WEAK_COVERAGE_SECTORS = {"construction", "retail", "domestic"}


def get_sector(ticker: str) -> str:
    """銘柄の業種を返す（未登録は domestic 扱い）。"""
    return TICKER_SECTOR.get(ticker, DEFAULT_SECTOR)


def _is_reliable(sector: str, used_series: List[str]) -> bool:
    """スコアが信頼できる水準かを判定する。

    構造的に弱い業種でも、政府統計（e-Stat）が実際に効いていれば信頼できる。
    """
    if sector not in WEAK_COVERAGE_SECTORS:
        return True
    return any(name.startswith(ESTAT_PREFIX) for name in used_series)


# ──────────────────────────────────────────
# 外部データの取得
# ──────────────────────────────────────────

def load_series(
    name: str,
    start_date: str,
    end_date: str,
    use_cache: bool = True,
) -> Optional[pd.Series]:
    """外部データ系列の終値を取得する。取得できない場合は None。

    data.download_stock_data を経由するので、キャッシュとオフライン時の
    合成データフォールバックは既存の仕組みをそのまま使う。
    """
    df = download_stock_data(name, start_date, end_date, use_cache=use_cache)
    if df is None or df.empty:
        return None
    return df["Close"].dropna()


def load_all_series(
    start_date: str,
    end_date: str,
    use_cache: bool = True,
    include_estat: bool = True,
) -> Dict[str, pd.Series]:
    """必要な外部データ系列をまとめて取得する。

    e-Stat 系列は設定済みかつAPIキーがある場合のみ含まれる。無ければ
    単に欠損として扱われ、スコアは残りのドライバーで正規化される。
    """
    out: Dict[str, pd.Series] = {}
    for name in SERIES_LABELS:
        if name.startswith(ESTAT_PREFIX):
            continue
        s = load_series(name, start_date, end_date, use_cache=use_cache)
        if s is not None and len(s) > 0:
            out[name] = s

    if include_estat:
        out.update(load_estat_series())

    return out


# ──────────────────────────────────────────
# 政府統計（e-Stat）— 建設・小売の本命データ
# ──────────────────────────────────────────

ESTAT_DATA_ENDPOINT = "https://api.e-stat.go.jp/rest/3.0/app/json/getStatsData"
ESTAT_LIST_ENDPOINT = "https://api.e-stat.go.jp/rest/3.0/app/json/getStatsList"

# e-Stat の接続設定ファイル。statsDataId は統計表ごとに異なり、政府統計は
# 改定されると ID が変わるため、コードに埋め込まず設定として持つ。
# ID は `python external_signals.py estat-search --keyword 建設工事受注` で探せる。
ESTAT_CONFIG_FILE = Path("estat_series.json")

# 設定ファイルのテンプレート。statsDataId を埋めると自動的に有効になる。
ESTAT_TEMPLATE: Dict[str, Dict] = {
    "estat:construction_orders": {
        "label": "建設工事受注動態統計（国交省・月次）",
        "stats_data_id": "",
        "cd_cat01": "",
        "lag_months": 2,
        "_hint": "estat-search --keyword 建設工事受注動態統計調査",
    },
    "estat:retail_sales": {
        "label": "商業動態統計 小売業販売額（経産省・月次）",
        "stats_data_id": "",
        "cd_cat01": "",
        "lag_months": 2,
        "_hint": "estat-search --keyword 商業動態統計",
    },
    "estat:machinery_orders": {
        "label": "機械受注統計（内閣府・月次）",
        "stats_data_id": "",
        "cd_cat01": "",
        "lag_months": 2,
        "_hint": "estat-search --keyword 機械受注統計",
    },
}


def load_estat_config() -> Dict[str, Dict]:
    """e-Stat 接続設定を読み込む。statsDataId が空のものは無効として除外する。"""
    if not ESTAT_CONFIG_FILE.exists():
        return {}
    try:
        raw = json.loads(ESTAT_CONFIG_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}
    return {
        k: v for k, v in raw.items()
        if isinstance(v, dict) and str(v.get("stats_data_id", "")).strip()
    }


def write_estat_template() -> Path:
    """設定ファイルのテンプレートを書き出す（既存ファイルは上書きしない）。"""
    if not ESTAT_CONFIG_FILE.exists():
        ESTAT_CONFIG_FILE.write_text(
            json.dumps(ESTAT_TEMPLATE, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    return ESTAT_CONFIG_FILE


def search_estat_tables(keyword: str, app_id: Optional[str] = None, limit: int = 20):
    """キーワードから statsDataId を検索する。

    統計表IDは e-Stat のサイトを手で辿らないと分からないうえ、統計の改定で
    変わることがある。手作業を避けるための検索ヘルパー。

    Returns:
        [(statsDataId, 統計名, 調査年月), ...]。取得失敗時は None。
    """
    app_id = app_id or os.environ.get("ESTAT_APP_ID")
    if not app_id:
        return None

    try:
        import requests

        resp = requests.get(
            ESTAT_LIST_ENDPOINT,
            params={"appId": app_id, "searchWord": keyword, "limit": limit},
            timeout=30,
        )
        resp.raise_for_status()
        payload = resp.json()

        tables = (
            payload.get("GET_STATS_LIST", {})
            .get("DATALIST_INF", {})
            .get("TABLE_INF", [])
        )
        if isinstance(tables, dict):
            tables = [tables]

        out = []
        for t in tables:
            def _text(node) -> str:
                if isinstance(node, dict):
                    return str(node.get("$", ""))
                return str(node or "")

            out.append((
                str(t.get("@id", "")),
                f"{_text(t.get('STAT_NAME'))} / {_text(t.get('TITLE'))}".strip(" /"),
                _text(t.get("SURVEY_DATE")),
            ))
        return out

    except Exception:
        return None


def _parse_estat_time(raw_time: str) -> Optional[pd.Timestamp]:
    """e-Stat の時間コードを月初の Timestamp に変換する。

    月次は "2024000101"（2024年 月 01月）のような10桁コードで返ってくる。
    """
    raw_time = str(raw_time)
    try:
        if len(raw_time) == 10 and raw_time[4:6] == "00":
            year, month = int(raw_time[:4]), int(raw_time[8:10])
        elif len(raw_time) >= 6:
            year, month = int(raw_time[:4]), int(raw_time[4:6])
        else:
            return None
        if not 1 <= month <= 12:
            return None
        return pd.Timestamp(year=year, month=month, day=1)
    except ValueError:
        return None


def fetch_estat_series(
    stats_data_id: str,
    app_id: Optional[str] = None,
    cd_cat01: Optional[str] = None,
    lag_months: int = 2,
) -> Optional[pd.Series]:
    """e-Stat（政府統計の総合窓口）から月次系列を取得する。

    建設工事受注動態統計調査や商業動態統計はここから取れる。無料だが
    アプリケーションIDの登録が必要（https://www.e-stat.go.jp/api/ で即日発行）。
    取得した ID を環境変数 ESTAT_APP_ID または .env に設定する。

    【公表ラグの扱い】
    政府統計は調査対象月から1〜2ヶ月遅れて公表される。2024年1月分の統計を
    2024年1月時点で使ってしまうと重大な先読みバイアスになるため、
    インデックスを lag_months ヶ月ずらして「実際に入手できた日」に置き直す。

    Returns:
        公表日（概算）をインデックスとする float の系列。取得失敗時は None。
    """
    app_id = app_id or os.environ.get("ESTAT_APP_ID")
    if not app_id or not stats_data_id:
        return None

    try:
        import requests

        params = {
            "appId": app_id,
            "statsDataId": stats_data_id,
            "metaGetFlg": "N",
            "cntGetFlg": "N",
        }
        if cd_cat01:
            params["cdCat01"] = cd_cat01

        resp = requests.get(ESTAT_DATA_ENDPOINT, params=params, timeout=30)
        resp.raise_for_status()
        payload = resp.json()

        values = (
            payload.get("GET_STATS_DATA", {})
            .get("STATISTICAL_DATA", {})
            .get("DATA_INF", {})
            .get("VALUE", [])
        )
        if isinstance(values, dict):
            values = [values]
        if not values:
            return None

        records: Dict[pd.Timestamp, float] = {}
        for v in values:
            ts = _parse_estat_time(v.get("@time", ""))
            if ts is None:
                continue
            try:
                val = float(str(v.get("$", "")).replace(",", ""))
            except (TypeError, ValueError):
                continue
            records[ts] = val

        if len(records) < 24:  # 前年同期比を取るには最低2年必要
            return None

        s = pd.Series(records).sort_index()
        # 調査対象月 → 公表日（概算）へインデックスをずらす
        s.index = s.index + pd.DateOffset(months=lag_months)
        return s

    except Exception:
        return None


def load_estat_series() -> Dict[str, pd.Series]:
    """設定済みの e-Stat 系列をまとめて取得する。

    APIキー未設定・設定ファイル未作成・取得失敗のいずれでも空辞書を返す。
    政府統計が無くてもスコアリング自体は動く設計にしてある。
    """
    config = load_estat_config()
    if not config:
        return {}

    out: Dict[str, pd.Series] = {}
    for key, spec in config.items():
        s = fetch_estat_series(
            stats_data_id=str(spec.get("stats_data_id", "")),
            cd_cat01=str(spec.get("cd_cat01", "")) or None,
            lag_months=int(spec.get("lag_months", 2)),
        )
        if s is not None and len(s) > 0:
            out[key] = s
            SERIES_LABELS.setdefault(key, str(spec.get("label", key)))
    return out


# ──────────────────────────────────────────
# スコア計算
# ──────────────────────────────────────────

# 四半期＝約63営業日
QUARTER_DAYS = 63
YEAR_DAYS = 252


def yoy_series(series: pd.Series, window: int = QUARTER_DAYS) -> pd.Series:
    """各時点における「直近1四半期平均 ÷ 1年前の同四半期平均 − 1」の系列を返す。

    決算は「前年同期比」で語られるので、外部変数も前年同期比で揃える。
    スポット値ではなく期間平均を使うのは、決算に効くのが期中平均だから。
    （例：ドル円は決算日の値ではなく、四半期の平均レートで換算される）

    時点ごとに毎回スライスすると重いので、全時点分を一括で計算しておく。
    """
    recent = series.rolling(window).mean()
    year_ago = recent.shift(YEAR_DAYS)
    out = recent / year_ago.where(year_ago != 0) - 1.0
    return out.replace([np.inf, -np.inf], np.nan)


def yoy_series_monthly(series: pd.Series, window: int = 3) -> pd.Series:
    """月次系列の前年同期比（直近3ヶ月平均 ÷ 前年同期3ヶ月平均 − 1）。

    政府統計は月次なので、営業日ベースの yoy_series は使えない。
    """
    recent = series.rolling(window).mean()
    year_ago = recent.shift(12)
    out = recent / year_ago.where(year_ago != 0) - 1.0
    return out.replace([np.inf, -np.inf], np.nan)


def build_change_panel(
    series: Dict[str, pd.Series],
    window: int = QUARTER_DAYS,
) -> Dict[str, pd.Series]:
    """全外部系列の前年同期比を前計算する。

    e-Stat 由来の系列は月次なので、月次用の計算に振り分ける。
    """
    out: Dict[str, pd.Series] = {}
    for name, s in series.items():
        if name.startswith(ESTAT_PREFIX):
            out[name] = yoy_series_monthly(s).dropna()
        else:
            out[name] = yoy_series(s, window).dropna()
    return out


def lookup_change(changes: Dict[str, pd.Series], name: str, asof: pd.Timestamp) -> Optional[float]:
    """asof 時点で判明している最新の前年同期比を返す（先読みしない）。

    外部系列は米国休場などで日本のカレンダーと一致しないため、
    asof 以前の直近値を拾う asof() を使う。
    """
    s = changes.get(name)
    if s is None or len(s) == 0:
        return None
    val = s.asof(asof)
    if val is None or pd.isna(val):
        return None
    return float(val)


@dataclass
class SignalScore:
    ticker: str
    sector: str
    score: float                       # 加重合成スコア（正=業績上振れ方向）
    breakdown: List[Tuple[str, float, float]]  # (系列名, 前年同期比, 寄与度)
    coverage: float                    # 実際に計算できたドライバーの重み合計
    reliable: bool                     # 外部指標で説明できる業種か
    mode: str = "fixed"                # fixed = 業種共通の重み / fitted = 銘柄固有の感応度


def score_ticker(
    ticker: str,
    asof: pd.Timestamp,
    changes: Dict[str, pd.Series],
) -> Optional[SignalScore]:
    """業種共通の重みで外部環境スコアを計算する（ベースライン）。

    注意: この方式では同一業種の銘柄が全て同じスコアになる。業種の
    ローテーション判断には使えるが、業種内の銘柄選択には使えない。
    銘柄を選び分けたい場合は score_ticker_fitted を使うこと。
    """
    sector = get_sector(ticker)
    drivers = SECTOR_DRIVERS.get(sector, [])
    if not drivers:
        return None

    breakdown: List[Tuple[str, float, float]] = []
    total = 0.0
    coverage = 0.0

    for d in drivers:
        chg = lookup_change(changes, d.series, asof)
        if chg is None:
            continue
        contrib = d.sign * d.weight * chg
        total += contrib
        coverage += d.weight
        breakdown.append((d.series, chg, contrib))

    if coverage == 0:
        return None

    # 取得できたドライバーの重みで正規化（欠損があっても比較可能にする）
    normalized = total / coverage

    return SignalScore(
        ticker=ticker,
        sector=sector,
        score=normalized,
        breakdown=breakdown,
        coverage=coverage,
        reliable=_is_reliable(sector, [n for n, _, _ in breakdown]),
        mode="fixed",
    )


# ──────────────────────────────────────────
# 銘柄固有の感応度推定
#
# 同じ「輸出企業」でも為替感応度は企業ごとに全く違う。海外売上比率、
# 現地生産比率、為替予約の方針で決まるからだ。業種共通の重みを使うと
# その差が消えてしまうので、過去データから銘柄ごとの感応度を推定する。
# ──────────────────────────────────────────

RIDGE_LAMBDA = 0.05  # 感応度推定の正則化強度（サンプル数が少ないため必須）


def estimate_sensitivity(
    price: pd.Series,
    changes: Dict[str, pd.Series],
    drivers: List[Driver],
    asof: pd.Timestamp,
    forward_days: int = QUARTER_DAYS,
    lookback_years: int = 6,
    step_days: int = 21,
) -> Optional[Dict[str, float]]:
    """銘柄ごとの外部変数に対する感応度をリッジ回帰で推定する。

    asof 時点より前のデータだけを使って、
        「各ドライバーの前年同期比」→「その後 forward_days 営業日のリターン」
    を回帰する。先読みバイアスを避けるため、学習に使う観測は
    リターン測定期間の終端が asof 以前に収まるものに限定している。

    サンプル数が少ない（月次で数十点）うえに観測期間が重複するため、
    素の OLS だと係数が暴れる。リッジ回帰で縮小をかけている。

    Returns:
        {系列名: 係数} の辞書。推定できない場合は None。
    """
    train_start = asof - pd.Timedelta(days=365 * lookback_years)
    hist = price[(price.index >= train_start) & (price.index <= asof)]
    if len(hist) < YEAR_DAYS + QUARTER_DAYS + forward_days:
        return None

    names = [d.series for d in drivers if d.series in changes]
    if not names:
        return None

    rows_x: List[List[float]] = []
    rows_y: List[float] = []

    dates = hist.index
    # リターンの終端が asof を超えないところまでで学習を打ち切る
    last = len(dates) - forward_days
    for i in range(0, last, step_days):
        t0 = dates[i]

        feats = []
        ok = True
        for name in names:
            chg = lookup_change(changes, name, t0)
            if chg is None:
                ok = False
                break
            feats.append(chg)
        if not ok:
            continue

        p0, p1 = float(hist.iloc[i]), float(hist.iloc[i + forward_days])
        if p0 <= 0:
            continue

        rows_x.append(feats)
        rows_y.append(p1 / p0 - 1.0)

    # 特徴量の数に対して最低限のサンプルを要求する
    if len(rows_y) < max(12, 3 * len(names)):
        return None

    X = np.asarray(rows_x, dtype=float)
    y = np.asarray(rows_y, dtype=float)

    # 標準化してからリッジ回帰（スケールの違うドライバーを公平に扱う）
    mu, sd = X.mean(axis=0), X.std(axis=0)
    sd = np.where(sd > 1e-12, sd, 1.0)
    Xs = (X - mu) / sd
    ys = y - y.mean()

    k = Xs.shape[1]
    try:
        coef_s = np.linalg.solve(Xs.T @ Xs + RIDGE_LAMBDA * len(ys) * np.eye(k), Xs.T @ ys)
    except np.linalg.LinAlgError:
        return None

    if not np.all(np.isfinite(coef_s)):
        return None

    # 標準化前のスケールに戻す
    coef = coef_s / sd
    return {name: float(c) for name, c in zip(names, coef)}


def score_ticker_fitted(
    ticker: str,
    asof: pd.Timestamp,
    changes: Dict[str, pd.Series],
    price: pd.Series,
    forward_days: int = QUARTER_DAYS,
) -> Optional[SignalScore]:
    """銘柄固有の感応度を使って外部環境スコアを計算する。

    スコアの単位は「予測される forward_days 営業日先のリターン」になる。
    """
    sector = get_sector(ticker)
    drivers = SECTOR_DRIVERS.get(sector, [])
    if not drivers:
        return None

    coefs = estimate_sensitivity(price, changes, drivers, asof, forward_days=forward_days)
    if not coefs:
        return None

    breakdown: List[Tuple[str, float, float]] = []
    total = 0.0
    for name, coef in coefs.items():
        chg = lookup_change(changes, name, asof)
        if chg is None:
            continue
        contrib = coef * chg
        total += contrib
        breakdown.append((name, chg, contrib))

    if not breakdown:
        return None

    return SignalScore(
        ticker=ticker,
        sector=sector,
        score=total,
        breakdown=breakdown,
        coverage=1.0,
        reliable=_is_reliable(sector, [n for n, _, _ in breakdown]),
        mode="fitted",
    )


def rank_universe(
    tickers: List[str],
    asof: pd.Timestamp,
    changes: Dict[str, pd.Series],
    prices: Optional[Dict[str, pd.Series]] = None,
    mode: str = "fixed",
) -> pd.DataFrame:
    """全銘柄を外部環境スコアでランキングする。

    mode="fitted" の場合は prices（銘柄→終値系列）が必要。
    """
    rows = []
    for t in tickers:
        if mode == "fitted":
            px = (prices or {}).get(t)
            if px is None:
                continue
            sc = score_ticker_fitted(t, asof, changes, px)
        else:
            sc = score_ticker(t, asof, changes)

        if sc is None:
            continue
        rows.append({
            "ticker": sc.ticker,
            "sector": sc.sector,
            "score": sc.score,
            "reliable": sc.reliable,
        })

    if not rows:
        return pd.DataFrame(columns=["ticker", "sector", "score", "reliable"])

    df = pd.DataFrame(rows).sort_values("score", ascending=False)
    return df.reset_index(drop=True)


# ──────────────────────────────────────────
# 検証: この手法は本当に効くのか
# ──────────────────────────────────────────

# ──────────────────────────────────────────
# バックテスト用: スコアの時系列化
# ──────────────────────────────────────────

def build_score_history(
    tickers: List[str],
    prices: Dict[str, pd.Series],
    changes: Dict[str, pd.Series],
    step_days: int = 21,
    mode: str = "fixed",
    forward_days: int = QUARTER_DAYS,
) -> Dict[str, pd.Series]:
    """各銘柄について、時点ごとの外部環境スコアの時系列を作る。

    バックテストのフィルターとして使うためのもの。毎営業日計算すると重いので
    step_days ごとに計算し、利用側は asof() で直近値を引く。

    先読みバイアス対策: 各時点のスコアはその時点までのデータのみで計算される
    （fitted モードの感応度推定も含む）。
    """
    out: Dict[str, pd.Series] = {}

    for t in tickers:
        px = prices.get(t)
        if px is None or len(px) == 0:
            continue

        # スコア計算に必要な履歴が溜まるまでは値を作らない
        first = YEAR_DAYS + QUARTER_DAYS
        if mode == "fitted":
            first += YEAR_DAYS + forward_days
        if len(px) <= first:
            continue

        points: Dict[pd.Timestamp, float] = {}
        for i in range(first, len(px), step_days):
            asof = px.index[i]
            if mode == "fitted":
                sc = score_ticker_fitted(t, asof, changes, px, forward_days=forward_days)
            else:
                sc = score_ticker(t, asof, changes)
            if sc is not None:
                points[asof] = sc.score

        if points:
            out[t] = pd.Series(points).sort_index()

    return out


def _spearman(a: List[float], b: List[float]) -> Optional[float]:
    """スピアマン順位相関。順位に直してからピアソン相関を取る。

    pandas の method="spearman" は scipy を要求するので、依存を増やさない
    ために自前で計算する（同順位は平均順位で処理される）。
    """
    sa, sb = pd.Series(a).rank(), pd.Series(b).rank()
    if sa.std() == 0 or sb.std() == 0:
        return None
    r = sa.corr(sb)
    return float(r) if pd.notna(r) else None


def information_coefficient(
    tickers: List[str],
    start_date: str,
    end_date: str,
    forward_days: int = QUARTER_DAYS,
    step_days: int = 21,
    use_cache: bool = True,
    mode: str = "fixed",
) -> Optional[pd.DataFrame]:
    """外部環境スコアと将来リターンの順位相関（情報係数 IC）を測る。

    月次で全銘柄にスコアを付け、その後 forward_days 営業日のリターンとの
    スピアマン順位相関を計算する。IC の平均が 0 から有意に離れていなければ、
    この手法にエッジは無いということ。

    一般に、月次 IC の平均が 0.03〜0.05 あれば実運用に足る水準とされる。

    先読みバイアス対策:
      - スコアは asof 時点までのデータのみで計算する
      - fitted モードの感応度も asof 以前のデータのみで推定する（expanding window）
      - リターンは asof より後の期間でのみ測る

    注意: 観測期間が重複している（step_days < forward_days）ため、
    t値は実際より大きく出る傾向がある。目安として読むこと。
    """
    series = load_all_series(start_date, end_date, use_cache=use_cache)
    if not series:
        print("  外部データを取得できませんでした。")
        return None
    changes = build_change_panel(series)

    prices: Dict[str, pd.Series] = {}
    for t in tickers:
        df = download_stock_data(t, start_date, end_date, use_cache=use_cache)
        if df is not None and not df.empty:
            prices[t] = df["Close"].dropna()

    if not prices:
        print("  株価データを取得できませんでした。")
        return None

    calendar = sorted(set().union(*[set(s.index) for s in prices.values()]))
    calendar = pd.DatetimeIndex(calendar)

    results = []
    # スコア計算に1年+1四半期、fitted は感応度推定にさらに履歴が要る
    first = YEAR_DAYS + QUARTER_DAYS
    if mode == "fitted":
        first += YEAR_DAYS + forward_days

    for i in range(first, len(calendar) - forward_days, step_days):
        asof = calendar[i]
        future = calendar[i + forward_days]

        scores, returns = [], []
        for t, px in prices.items():
            if mode == "fitted":
                sc = score_ticker_fitted(t, asof, changes, px, forward_days=forward_days)
            else:
                sc = score_ticker(t, asof, changes)
            if sc is None:
                continue

            hist = px[px.index <= asof]
            fut = px[px.index <= future]
            if len(hist) == 0 or len(fut) == 0:
                continue
            p0, p1 = float(hist.iloc[-1]), float(fut.iloc[-1])
            if p0 <= 0:
                continue

            scores.append(sc.score)
            returns.append(p1 / p0 - 1.0)

        if len(scores) < 10:
            continue

        ic = _spearman(scores, returns)
        if ic is not None:
            results.append({"date": asof, "ic": ic, "n": len(scores)})

    if not results:
        return None
    return pd.DataFrame(results)


# ──────────────────────────────────────────
# CLI
# ──────────────────────────────────────────

def _default_range(years: int = 4) -> Tuple[str, str]:
    end = date.today()
    start = end - timedelta(days=365 * years + 120)
    return start.isoformat(), end.isoformat()


def _load_prices(tickers: List[str], start: str, end: str, use_cache: bool) -> Dict[str, pd.Series]:
    prices: Dict[str, pd.Series] = {}
    for t in tickers:
        df = download_stock_data(t, start, end, use_cache=use_cache)
        if df is not None and not df.empty:
            prices[t] = df["Close"].dropna()
    return prices


def cmd_rank(args: argparse.Namespace) -> None:
    # fitted モードは感応度推定に長い履歴が必要
    years = 9 if args.mode == "fitted" else 4
    start, end = _default_range(years)
    series = load_all_series(start, end, use_cache=not args.no_cache)

    if not series:
        print("  外部データを取得できませんでした。")
        return

    missing = [k for k in SERIES_LABELS if k not in series]
    if missing:
        print(f"  ⚠️  取得できなかった系列: {', '.join(missing)}\n")

    changes = build_change_panel(series)

    prices = None
    if args.mode == "fitted":
        prices = _load_prices(DEFAULT_TICKERS, start, end, not args.no_cache)

    asof = pd.Timestamp(end)
    df = rank_universe(DEFAULT_TICKERS, asof, changes, prices=prices, mode=args.mode)
    if df.empty:
        print("  スコアを計算できませんでした（履歴が不足しています）。")
        return

    unit = "予測四半期リターン" if args.mode == "fitted" else "外部環境スコア"
    top = df.head(args.top)
    print(f"\n  {unit} 上位{len(top)}銘柄  （基準日: {asof.date()} / mode={args.mode}）")
    print(f"  {'順位':<5}{'銘柄':<11}{'業種':<15}{'スコア':>9}   信頼度")
    print("  " + "-" * 56)
    for i, r in top.iterrows():
        mark = "○" if r["reliable"] else "△ 外部指標弱"
        print(f"  {i + 1:<5}{r['ticker']:<11}{r['sector']:<15}{r['score']:>+8.2%}   {mark}")

    weak = int((~df["reliable"]).sum())
    print(f"\n  △ = 外部市況との連動が構造的に弱い業種（{weak}/{len(df)}銘柄）。")
    print("     建設・小売はe-Stat政府統計を接続すると精度が上がります。")
    if args.mode == "fixed":
        print("     fixed モードは業種共通の重みなので、同一業種内の順位は付きません。")
        print("     銘柄を選び分けるには --mode fitted を使ってください。")
    print("     実際に使う前に validate モードでICを確認してください。")


def cmd_explain(args: argparse.Namespace) -> None:
    years = 9 if args.mode == "fitted" else 4
    start, end = _default_range(years)
    series = load_all_series(start, end, use_cache=not args.no_cache)
    if not series:
        print("  外部データを取得できませんでした。")
        return
    changes = build_change_panel(series)
    asof = pd.Timestamp(end)

    if args.mode == "fitted":
        df = download_stock_data(args.ticker, start, end, use_cache=not args.no_cache)
        if df is None or df.empty:
            print(f"  {args.ticker} の株価データを取得できませんでした。")
            return
        sc = score_ticker_fitted(args.ticker, asof, changes, df["Close"].dropna())
    else:
        sc = score_ticker(args.ticker, asof, changes)

    if sc is None:
        print(f"  {args.ticker} のスコアを計算できませんでした。")
        return

    print(f"\n  {sc.ticker}  業種: {sc.sector}  基準日: {asof.date()}  mode: {sc.mode}")
    print(f"  合成スコア: {sc.score:+.2%}"
          f"{'' if sc.reliable else '  ⚠️ この業種は外部指標での説明力が低い'}")
    print(f"\n  {'外部変数':<38}{'前年同期比':>11}{'寄与':>10}")
    print("  " + "-" * 60)
    for name, chg, contrib in sc.breakdown:
        label = SERIES_LABELS.get(name, name)
        print(f"  {label:<38}{chg:>+10.1%}{contrib:>+10.2%}")
    print("\n  ※ 前年同期比は直近1四半期の平均値と1年前同期の平均値の比較")
    if sc.mode == "fitted":
        print("  ※ 寄与 = 銘柄固有の感応度（過去データから推定）× 前年同期比")


def cmd_validate(args: argparse.Namespace) -> None:
    print(f"  検証期間: {args.start} 〜 {args.end}")
    print(f"  対象: {len(DEFAULT_TICKERS)}銘柄 / 予測期間: {args.forward}営業日先\n")

    ic_df = information_coefficient(
        DEFAULT_TICKERS,
        args.start,
        args.end,
        forward_days=args.forward,
        use_cache=not args.no_cache,
        mode=args.mode,
    )
    if ic_df is None or ic_df.empty:
        print("  ICを計算できませんでした（データ不足）。")
        return

    mean_ic = ic_df["ic"].mean()
    std_ic = ic_df["ic"].std()
    n = len(ic_df)
    # IR = 平均IC / ICの標準偏差、t値 = IR * sqrt(観測数)
    ir = mean_ic / std_ic if std_ic > 0 else float("nan")
    t_stat = ir * np.sqrt(n) if np.isfinite(ir) else float("nan")
    hit = float((ic_df["ic"] > 0).mean())

    print(f"  観測回数        : {n}")
    print(f"  平均IC          : {mean_ic:+.4f}")
    print(f"  IC標準偏差      : {std_ic:.4f}")
    print(f"  情報比 (IR)     : {ir:+.3f}")
    print(f"  t値             : {t_stat:+.2f}")
    print(f"  ICがプラスの割合: {hit:.1%}")

    print()
    if abs(t_stat) < 2:
        print("  判定: 統計的に有意なエッジは検出されませんでした（|t| < 2）。")
        print("        重み付けの見直し、または業種別の検証が必要です。")
    elif mean_ic > 0:
        print("  判定: プラスのエッジが検出されました。")
        print("        ただし取引コストとスリッページを引いた後で残るかは別途要検証。")
    else:
        print("  判定: 有意にマイナスのICです。符号が逆の可能性があります。")

    if args.output:
        ic_df.to_csv(args.output, index=False)
        print(f"\n  IC時系列を保存: {args.output}")


def cmd_sources(args: argparse.Namespace) -> None:
    print("\n  【市場データ】yfinance経由・APIキー不要")
    print("  " + "-" * 62)
    for name, label in SERIES_LABELS.items():
        if not name.startswith(ESTAT_PREFIX):
            print(f"  {name:<10}{label}")

    has_key = bool(os.environ.get("ESTAT_APP_ID"))
    estat_config = load_estat_config()

    print("\n  【政府統計】e-Stat API")
    print("  " + "-" * 62)
    print(f"  ESTAT_APP_ID      : {'設定済み' if has_key else '未設定'}")
    print(f"  {ESTAT_CONFIG_FILE.name:<18}: "
          f"{f'{len(estat_config)}系列を設定済み' if estat_config else '未設定'}")

    for name, label in SERIES_LABELS.items():
        if name.startswith(ESTAT_PREFIX):
            mark = "✓" if name in estat_config else "×"
            print(f"  {mark} {name:<28}{label}")

    if not (has_key and estat_config):
        print("\n  接続手順:")
        print("    1. https://www.e-stat.go.jp/api/ で登録しアプリケーションIDを取得")
        print("    2. .env に ESTAT_APP_ID=取得したID を記入")
        print("    3. python external_signals.py estat-init   ← 設定ファイルを生成")
        print("    4. python external_signals.py estat-search --keyword 建設工事受注動態統計")
        print(f"    5. 出てきた統計表IDを {ESTAT_CONFIG_FILE.name} の stats_data_id に記入")

    print("\n  【業種別カバレッジ】")
    print("  " + "-" * 62)
    counts: Dict[str, int] = {}
    for t in DEFAULT_TICKERS:
        s = get_sector(t)
        counts[s] = counts.get(s, 0) + 1

    reliable_n = 0
    for sector, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        drivers = SECTOR_DRIVERS.get(sector, [])
        used = [d.series for d in drivers
                if not d.series.startswith(ESTAT_PREFIX) or d.series in estat_config]
        ok = _is_reliable(sector, used)
        reliable_n += cnt if ok else 0
        names = ", ".join(used)
        pending = [d.series for d in drivers
                   if d.series.startswith(ESTAT_PREFIX) and d.series not in estat_config]
        suffix = f"  （未接続: {', '.join(pending)}）" if pending else ""
        print(f"  {'○' if ok else '△'} {sector:<15}{cnt:>4}銘柄  ← {names}{suffix}")

    print(f"\n  信頼できるスコアが出る銘柄: {reliable_n}/{len(DEFAULT_TICKERS)}")


def cmd_estat_init(args: argparse.Namespace) -> None:
    existed = ESTAT_CONFIG_FILE.exists()
    path = write_estat_template()
    if existed:
        print(f"\n  {path} は既に存在するため、上書きしませんでした。")
    else:
        print(f"\n  設定ファイルを作成しました: {path}")
    print("  各系列の stats_data_id を埋めると自動的に有効になります。")
    print("  IDの探し方: python external_signals.py estat-search --keyword 建設工事受注動態統計")


def cmd_estat_search(args: argparse.Namespace) -> None:
    if not os.environ.get("ESTAT_APP_ID"):
        print("\n  ESTAT_APP_ID が設定されていません。")
        print("  https://www.e-stat.go.jp/api/ で登録し .env に記入してください。")
        return

    results = search_estat_tables(args.keyword, limit=args.limit)
    if results is None:
        print("\n  e-Stat への接続に失敗しました（APIキー・ネットワークを確認してください）。")
        return
    if not results:
        print(f"\n  「{args.keyword}」に一致する統計表は見つかりませんでした。")
        return

    print(f"\n  「{args.keyword}」の検索結果 {len(results)}件")
    print("  " + "-" * 76)
    for sid, title, survey in results:
        print(f"  {sid:<14}{title[:52]:<54}{survey}")
    print(f"\n  使いたい統計表のIDを {ESTAT_CONFIG_FILE.name} の stats_data_id に記入してください。")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="外部データによる業績サプライズ推定",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--no-cache", action="store_true", help="キャッシュを使わず再取得")
    sub = parser.add_subparsers(dest="command")

    mode_help = ("fixed = 業種共通の重み（業種ローテーション向け） / "
                 "fitted = 銘柄固有の感応度を過去データから推定（銘柄選択向け）")

    p_rank = sub.add_parser("rank", help="外部環境スコアで全銘柄をランキング")
    p_rank.add_argument("--top", type=int, default=20, help="表示件数")
    p_rank.add_argument("--mode", choices=["fixed", "fitted"], default="fixed", help=mode_help)
    p_rank.set_defaults(func=cmd_rank)

    p_exp = sub.add_parser("explain", help="スコアの内訳を表示")
    p_exp.add_argument("--ticker", required=True, help="例: 8035.T")
    p_exp.add_argument("--mode", choices=["fixed", "fitted"], default="fixed", help=mode_help)
    p_exp.set_defaults(func=cmd_explain)

    p_val = sub.add_parser("validate", help="情報係数(IC)で有効性を検証")
    p_val.add_argument("--start", default="2016-01-01")
    p_val.add_argument("--end", default="2024-12-31")
    p_val.add_argument("--forward", type=int, default=QUARTER_DAYS,
                       help="何営業日先のリターンを予測対象にするか")
    p_val.add_argument("--mode", choices=["fixed", "fitted"], default="fixed", help=mode_help)
    p_val.add_argument("--output", default="", help="IC時系列のCSV出力先")
    p_val.set_defaults(func=cmd_validate)

    p_src = sub.add_parser("sources", help="外部データ系列とカバレッジを表示")
    p_src.set_defaults(func=cmd_sources)

    p_init = sub.add_parser("estat-init", help="e-Stat 接続設定ファイルを生成")
    p_init.set_defaults(func=cmd_estat_init)

    p_search = sub.add_parser("estat-search", help="e-Stat の統計表IDをキーワード検索")
    p_search.add_argument("--keyword", required=True, help="例: 建設工事受注動態統計調査")
    p_search.add_argument("--limit", type=int, default=20, help="表示件数")
    p_search.set_defaults(func=cmd_estat_search)

    args = parser.parse_args()
    if not getattr(args, "command", None):
        parser.print_help()
        return
    args.func(args)


if __name__ == "__main__":
    main()
