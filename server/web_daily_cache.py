# -*- coding: utf-8 -*-
"""日資料快取：避免每次分析重抓 100+ 天歷史。"""

from __future__ import annotations

import json
from datetime import datetime, timezone

import pandas as pd
from sqlalchemy.orm import Session

from database import SessionLocal
from models import DailyMarketCache


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


def fetch_market_day_cached(trade_date: str, market: str, fetch_fn):
    db = SessionLocal()
    try:
        cached = load_daily_cache(db, trade_date, market)
        if cached is not None and not cached.empty:
            return cached
        day_obj = datetime.strptime(trade_date, "%Y-%m-%d")
        df = fetch_fn(day_obj)
        if isinstance(df, pd.DataFrame) and not df.empty:
            save_daily_cache(db, trade_date, market, df)
            return df
        return None
    finally:
        db.close()
