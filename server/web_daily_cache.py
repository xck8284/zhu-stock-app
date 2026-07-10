# -*- coding: utf-8 -*-
"""日資料快取：避免每次分析重抓 100+ 天歷史。"""

from __future__ import annotations

import json
import threading
from concurrent.futures import ThreadPoolExecutor
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
    1) 只查 key（不載入 JSON）→ 回報快取覆蓋率
    2) 分批載入 JSON 並 parse（避免一次把 600+ 天 JSON 全塞進記憶體）
    progress_callback(done, total, message=None)
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
    cached_keys: set[tuple[str, str]] = set()
    try:
        rows = (
            db.query(DailyMarketCache.trade_date, DailyMarketCache.market)
            .filter(DailyMarketCache.trade_date >= start, DailyMarketCache.trade_date <= end)
            .all()
        )
        for trade_date, market in rows:
            if trade_date in wanted:
                cached_keys.add((trade_date, market))

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
        progress_callback(
            cached_done,
            total_tasks,
            f"快取覆蓋 {cached_done}/{total_tasks}，開始解析歷史資料…",
        )

    keys_to_parse: list[tuple[str, str]] = []
    for day in dates:
        trade_date = day.strftime("%Y-%m-%d")
        for market in ("上市", "上櫃"):
            if (trade_date, market) in cached_keys:
                keys_to_parse.append((trade_date, market))

    parse_total = len(keys_to_parse)
    frames: list[pd.DataFrame] = []
    batch_size = 40

    for batch_start in range(0, parse_total, batch_size):
        batch_keys = keys_to_parse[batch_start : batch_start + batch_size]
        batch_dates = list({trade_date for trade_date, _ in batch_keys})

        db = SessionLocal()
        try:
            rows = (
                db.query(DailyMarketCache.trade_date, DailyMarketCache.market, DailyMarketCache.data_json)
                .filter(
                    DailyMarketCache.trade_date.in_(batch_dates),
                    DailyMarketCache.market.in_(["上市", "上櫃"]),
                )
                .all()
            )
            row_map = {(trade_date, market): data_json for trade_date, market, data_json in rows}
        finally:
            db.close()

        batch_records: list[dict] = []

        def _parse_key(key):
            trade_date, market = key
            data_json = row_map.get(key)
            records: list = []
            if data_json and data_json != "[]":
                try:
                    parsed = json.loads(data_json)
                    if isinstance(parsed, list):
                        records = parsed
                except Exception:
                    records = []
            return trade_date, market, records

        with ThreadPoolExecutor(max_workers=4) as pool:
            for trade_date, market, records in pool.map(_parse_key, batch_keys):
                if records:
                    batch_records.extend(records)

        if batch_records:
            frames.append(pd.DataFrame(batch_records))

        parsed = min(parse_total, batch_start + len(batch_keys))
        if progress_callback:
            progress_callback(
                parsed,
                parse_total,
                f"解析歷史快取 {parsed}/{parse_total}（完成後開始選股）",
            )

    if progress_callback:
        progress_callback(total_tasks, total_tasks, "歷史快取完成，準備選股…")

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
