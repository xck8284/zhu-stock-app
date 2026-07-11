from datetime import datetime, timezone

from sqlalchemy import Column, DateTime, Integer, String, Text, UniqueConstraint

from cache_database import CacheBase


class DailyMarketCache(CacheBase):
    __tablename__ = "daily_market_cache"
    __table_args__ = (
        UniqueConstraint("trade_date", "market", name="uq_cache_trade_date_market"),
    )

    id = Column(Integer, primary_key=True, index=True)
    trade_date = Column(String(10), nullable=False, index=True)
    market = Column(String(16), nullable=False, index=True)
    data_json = Column(Text, nullable=False)
    updated_at = Column(DateTime, default=lambda: datetime.now(timezone.utc), nullable=False)
