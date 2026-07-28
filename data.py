"""
日本株データ取得モジュール

yfinance でのダウンロードが失敗した場合（ネットワーク制限など）は
synthetic_data モジュールによる合成データに自動フォールバックする。
"""

import pickle
import warnings
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import yfinance as yf
from tqdm import tqdm

warnings.filterwarnings("ignore")

CACHE_DIR = Path("cache")


def _cache_path(ticker: str, start: str, end: str) -> Path:
    CACHE_DIR.mkdir(exist_ok=True)
    key = f"{ticker}_{start}_{end}.pkl"
    return CACHE_DIR / key


def _download_via_yfinance(
    ticker: str,
    start_date: str,
    end_date: str,
) -> Optional[pd.DataFrame]:
    import contextlib, io
    try:
        with contextlib.redirect_stderr(io.StringIO()):
            df = yf.download(
                ticker,
                start=start_date,
                end=end_date,
                auto_adjust=True,
                progress=False,
            )
        if df.empty:
            return None

        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)

        df = df[["Open", "High", "Low", "Close", "Volume"]].copy()
        df.index = pd.to_datetime(df.index)
        df = df.sort_index()
        df.dropna(subset=["Close"], inplace=True)
        return df if not df.empty else None

    except Exception:
        return None


def download_stock_data(
    ticker: str,
    start_date: str,
    end_date: str,
    use_cache: bool = True,
    use_synthetic_fallback: bool = True,
) -> Optional[pd.DataFrame]:
    """単一銘柄のOHLCVデータを取得する。"""
    cache_file = _cache_path(ticker, start_date, end_date)

    if use_cache and cache_file.exists():
        with open(cache_file, "rb") as f:
            return pickle.load(f)

    df = _download_via_yfinance(ticker, start_date, end_date)

    if df is None and use_synthetic_fallback:
        from synthetic_data import generate_stock_data
        df = generate_stock_data(ticker, start_date, end_date)

    if df is not None and use_cache:
        with open(cache_file, "wb") as f:
            pickle.dump(df, f)

    return df


def download_all_stocks(
    tickers: List[str],
    start_date: str,
    end_date: str,
    use_cache: bool = True,
    use_synthetic_fallback: bool = True,
) -> Dict[str, pd.DataFrame]:
    """複数銘柄のデータを一括取得する。"""
    # まず yfinance を試みる（1件でも成功するか確認）
    test_df = _download_via_yfinance(tickers[0], start_date, end_date)
    network_ok = test_df is not None

    if not network_ok and use_synthetic_fallback:
        print("ネットワーク接続なし → 合成データモードで実行します")
        from synthetic_data import generate_all_stocks
        return generate_all_stocks(tickers, start_date, end_date)

    print(f"株価データを取得中... ({len(tickers)}銘柄)")
    result: Dict[str, pd.DataFrame] = {}

    for ticker in tqdm(tickers, desc="データ取得"):
        df = download_stock_data(ticker, start_date, end_date, use_cache, use_synthetic_fallback)
        if df is not None and len(df) > 0:
            result[ticker] = df

    print(f"取得完了: {len(result)}/{len(tickers)} 銘柄")
    return result


def download_index(
    ticker: str,
    start_date: str,
    end_date: str,
    use_cache: bool = True,
) -> Optional[pd.DataFrame]:
    """指数・為替の実データのみを取得する（合成データにフォールバックしない）。

    相関分析では合成データを1本でも混ぜると結果が無意味になるため、
    取得できなかった場合は None を返し、呼び出し側で明示的に扱う。
    """
    return download_stock_data(
        ticker,
        start_date,
        end_date,
        use_cache=use_cache,
        use_synthetic_fallback=False,
    )


def download_benchmark(
    start_date: str,
    end_date: str,
    use_cache: bool = True,
    use_synthetic_fallback: bool = True,
) -> Optional[pd.DataFrame]:
    """ベンチマーク（日経225: ^N225）を取得する。"""
    df = download_stock_data("^N225", start_date, end_date, use_cache, use_synthetic_fallback=False)
    if df is None and use_synthetic_fallback:
        from synthetic_data import generate_benchmark
        df = generate_benchmark(start_date, end_date)
    return df
