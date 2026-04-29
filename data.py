"""
日本株データ取得モジュール
"""

import os
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


def download_stock_data(
    ticker: str,
    start_date: str,
    end_date: str,
    use_cache: bool = True,
) -> Optional[pd.DataFrame]:
    """単一銘柄のOHLCVデータを取得する。"""
    cache_file = _cache_path(ticker, start_date, end_date)

    if use_cache and cache_file.exists():
        with open(cache_file, "rb") as f:
            return pickle.load(f)

    try:
        df = yf.download(
            ticker,
            start=start_date,
            end=end_date,
            auto_adjust=True,
            progress=False,
        )
        if df.empty:
            return None

        # マルチインデックスの場合はフラット化
        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)

        df = df[["Open", "High", "Low", "Close", "Volume"]].copy()
        df.index = pd.to_datetime(df.index)
        df = df.sort_index()
        df.dropna(subset=["Close"], inplace=True)

        if use_cache:
            with open(cache_file, "wb") as f:
                pickle.dump(df, f)

        return df

    except Exception as e:
        print(f"  警告: {ticker} のデータ取得に失敗しました: {e}")
        return None


def download_all_stocks(
    tickers: List[str],
    start_date: str,
    end_date: str,
    use_cache: bool = True,
) -> Dict[str, pd.DataFrame]:
    """複数銘柄のデータを一括取得する。"""
    print(f"株価データを取得中... ({len(tickers)}銘柄)")
    result: Dict[str, pd.DataFrame] = {}

    for ticker in tqdm(tickers, desc="データ取得"):
        df = download_stock_data(ticker, start_date, end_date, use_cache)
        if df is not None and len(df) > 0:
            result[ticker] = df

    print(f"取得完了: {len(result)}/{len(tickers)} 銘柄")
    return result


def download_benchmark(
    start_date: str,
    end_date: str,
    use_cache: bool = True,
) -> Optional[pd.DataFrame]:
    """ベンチマーク（日経225: ^N225）を取得する。"""
    return download_stock_data("^N225", start_date, end_date, use_cache)
