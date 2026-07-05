# -*- coding: utf-8 -*-
"""日資料快取：避免每次分析重抓 100+ 天歷史。"""

from __future__ import annotations

import json
import threading
from datetime import datetime, timezone

import pandas as pd
from sqlalchemy.orm import Session

from config import settings
from database import SessionLocal
from models import DailyMarketCache

_write_lock = threading.Lock()
_mem_cache: dict[tuple[str, str], pd.DataFrame] = {}


def _now_utc():
    return datetime.now(timezone.utc)


def load_daily_cache(db: Session, trade_date: str, market: str) -> pd.DataFrame | None:
    row = (
        db.query(DailyMarketCache)
        .filter(DailyMarketCache.trade_date == trade_date, DailyMarketCache.market == market)
        .first()
    )
    if not row or not row.data_json:
        return None
    try:
        records = json.loads(row.data_json)
        if not records:
            return None
        return pd.DataFrame(records)
    except Exception:
        return None


def save_daily_cache(db: Session, trade_date: str, market: str, df: pd.DataFrame) -> None:
    if df is None or df.empty:
        return
    payload = json.dumps(df.to_dict(orient="records"), ensure_ascii=False)
    with _write_lock:
        row = (
            db.query(DailyMarketCache)
            .filter(DailyMarketCache.trade_date == trade_date, DailyMarketCache.market == market)
            .first()
        )
        if row is None:
            row = DailyMarketCache(trade_date=trade_date, market=market, data_json=payload, updated_at=_now_utc())
            db.add(row)
        else:
            row.data_json = payload
            row.updated_at = _now_utc()
        db.commit()
    _mem_cache[(trade_date, market)] = df


def preload_cached_history_frames(dates: list) -> tuple[list[pd.DataFrame], list[tuple], list[tuple]]:
    """單次 DB 連線載入已快取日資料，回傳 (frames, 上市待抓, 上櫃待抓)。"""
    frames: list[pd.DataFrame] = []
    listed_missing: list[tuple] = []
    otc_missing: list[tuple] = []
    db = SessionLocal()
    try:
        for day in dates:
            trade_date = day.strftime("%Y-%m-%d")
            listed = load_daily_cache(db, trade_date, "上市")
            if listed is not None and not listed.empty:
                frames.append(listed)
                _mem_cache[(trade_date, "上市")] = listed
            else:
                listed_missing.append(day)
            otc = load_daily_cache(db, trade_date, "上櫃")
            if otc is not None and not otc.empty:
                frames.append(otc)
                _mem_cache[(trade_date, "上櫃")] = otc
            else:
                otc_missing.append(day)
    finally:
        db.close()
    return frames, listed_missing, otc_missing


def count_cached_trading_days() -> int:
    db = SessionLocal()
    try:
        return int(db.query(DailyMarketCache.trade_date).distinct().count())
    except Exception:
        return 0
    finally:
        db.close()


def fetch_market_day_cached(trade_date: str, market: str, fetch_fn):
    mem_key = (trade_date, market)
    cached_mem = _mem_cache.get(mem_key)
    if cached_mem is not None and not cached_mem.empty:
        return cached_mem

    db = SessionLocal()
    try:
        cached = load_daily_cache(db, trade_date, market)
        if cached is not None and not cached.empty:
            _mem_cache[mem_key] = cached
            return cached
        day_obj = datetime.strptime(trade_date, "%Y-%m-%d")
        df = fetch_fn(day_obj)
        if isinstance(df, pd.DataFrame) and not df.empty:
            save_daily_cache(db, trade_date, market, df)
            return df
        return None
    finally:
        db.close()


def history_fetch_workers() -> tuple[int, int]:
    """SQLite 只能單寫入，Render 上並行過高會拖垮 API。"""
    url = (settings.DATABASE_URL or "").lower()
    if "sqlite" in url:
        return 3, 3
    return 8, 10
