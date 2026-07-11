from sqlalchemy import create_engine
from sqlalchemy.orm import declarative_base, sessionmaker

from config import settings


cache_engine = create_engine(
    settings.CACHE_DATABASE_URL,
    connect_args={"check_same_thread": False}
    if settings.CACHE_DATABASE_URL.startswith("sqlite")
    else {},
)
CacheSessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=cache_engine)
CacheBase = declarative_base()
