# -*- coding: utf-8 -*-
"""日資料快取：避免每次分析重抓 100+ 天歷史。"""

from __future__ import annotations

import json
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
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
            return pd.DataFrame()
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


def _parse_cache_row(trade_date: str, market: str, data_json: str | None) -> pd.DataFrame:
    df = pd.DataFrame()
    if data_json:
        try:
            records = json.loads(data_json)
            if records:
                df = pd.DataFrame(records)
        except Exception:
            df = pd.DataFrame()
    _mem_cache[(trade_date, market)] = df
    return df


def preload_cached_history_frames(dates: list, progress_callback=None):
    """
    兩階段：
    1) 只查 key（不 parse JSON）→ 立刻回報覆蓋率
    2) 合併 parse 成單一 DataFrame（避免 600+ 次 concat 卡死）
    """
    listed_missing: list = []
    otc_missing: list = []
    if not dates:
        return [], listed_missing, otc_missing

    wanted = {day.strftime("%Y-%m-%d") for day in dates}
    start = min(wanted)
    end = max(wanted)
    total_tasks = len(dates) * 2

    db = SessionLocal()
    cached_rows: list[tuple[str, str, str | None]] = []
    cached_keys: set[tuple[str, str]] = set()
    try:
        rows = (
            db.query(DailyMarketCache.trade_date, DailyMarketCache.market, DailyMarketCache.data_json)
            .filter(DailyMarketCache.trade_date >= start, DailyMarketCache.trade_date <= end)
            .all()
        )
        for trade_date, market, data_json in rows:
            if trade_date not in wanted:
                continue
            cached_keys.add((trade_date, market))
            cached_rows.append((trade_date, market, data_json))

        for day in dates:
            trade_date = day.strftime("%Y-%m-%d")
            if (trade_date, "上市") not in cached_keys:
                listed_missing.append(day)
            if (trade_date, "上櫃") not in cached_keys:
                otc_missing.append(day)
    finally:
        db.close()

    cached_done = total_tasks - len(listed_missing) - len(otc_missing)
    if progress_callback:
        progress_callback(cached_done, total_tasks)

    all_records: list[dict] = []
    parse_total = len(cached_rows)

    def _parse_row(row):
        trade_date, market, data_json = row
        if not data_json or data_json == "[]":
            return trade_date, market, []
        try:
            records = json.loads(data_json)
            return trade_date, market, records if isinstance(records, list) else []
        except Exception:
            return trade_date, market, []

    workers = 4
    done_parse = 0
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(_parse_row, row) for row in cached_rows]
        for future in as_completed(futures):
            trade_date, market, records = future.result()
            if records:
                all_records.extend(records)
            _mem_cache[(trade_date, market)] = pd.DataFrame(records) if records else pd.DataFrame()
            done_parse += 1
            if progress_callback and (done_parse % 40 == 0 or done_parse == parse_total):
                mapped = min(total_tasks, max(cached_done, done_parse))
                progress_callback(mapped, total_tasks)

    frames: list[pd.DataFrame] = []
    if all_records:
        frames.append(pd.DataFrame(all_records))

    if progress_callback:
        progress_callback(total_tasks, total_tasks)

    return frames, listed_missing, otc_missing


def count_cached_trading_days() -> int:
    db = SessionLocal()
    try:
        return int(db.query(DailyMarketCache.trade_date).distinct().count())
    except Exception:
        return 0
    finally:
        db.close()


def save_empty_daily_cache(db: Session, trade_date: str, market: str) -> None:
    """標記無資料日，避免每次分析重試卡死。"""
    payload = "[]"
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
    _mem_cache[(trade_date, market)] = pd.DataFrame()


def fetch_market_day_cached(trade_date: str, market: str, fetch_fn):
    mem_key = (trade_date, market)
    cached_mem = _mem_cache.get(mem_key)
    if cached_mem is not None:
        return cached_mem if not cached_mem.empty else None

    db = SessionLocal()
    try:
        cached = load_daily_cache(db, trade_date, market)
        if cached is not None:
            _mem_cache[mem_key] = cached
            return cached if not cached.empty else None
        day_obj = datetime.strptime(trade_date, "%Y-%m-%d")
        df = fetch_fn(day_obj)
        if isinstance(df, pd.DataFrame) and not df.empty:
            save_daily_cache(db, trade_date, market, df)
            return df
        save_empty_daily_cache(db, trade_date, market)
        return None
    finally:
        db.close()


def history_fetch_workers() -> tuple[int, int]:
    """SQLite 只能單寫入，Render 上並行過高會拖垮 API。"""
    url = (settings.DATABASE_URL or "").lower()
    if "sqlite" in url:
        return 3, 3
    return 8, 10
