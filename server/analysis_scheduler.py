# -*- coding: utf-8 -*-
"""交易日收盤後自動執行網頁版分析。"""

from __future__ import annotations

import logging
import threading
from datetime import datetime
from zoneinfo import ZoneInfo

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

from analysis import run_analysis
from config import settings
from web_analysis_store import load_web_analysis_result, save_web_analysis_result

logger = logging.getLogger("zhu.analysis")
TW = ZoneInfo("Asia/Taipei")

_scheduler: BackgroundScheduler | None = None
_job_lock = threading.Lock()
_job_running = False


def tw_now() -> datetime:
    return datetime.now(TW)


def is_trading_weekday(dt: datetime | None = None) -> bool:
    dt = dt or tw_now()
    return dt.weekday() < 5


def should_refresh_analysis(meta: dict | None, dt: datetime | None = None) -> bool:
    """收盤資料穩定後（預設 16:05 台北時間）且當日尚未更新時需要重跑。"""
    dt = dt or tw_now()
    if not is_trading_weekday(dt):
        return False

    hour, minute = settings.AUTO_ANALYSIS_HOUR, settings.AUTO_ANALYSIS_MINUTE
    if (dt.hour, dt.minute) < (hour, minute):
        return False

    meta = meta or {}
    settle_date = str(meta.get("settle_date") or "")
    today = dt.strftime("%Y-%m-%d")
    if settle_date == today and meta.get("job_status") != "failed":
        return False
    return True


def _set_job_meta(result: dict, status: str, error: str = "") -> dict:
    result = dict(result or {})
    result["job_status"] = status
    result["job_error"] = error
    return result


def run_analysis_and_persist(trigger: str = "manual") -> dict:
    global _job_running

    with _job_lock:
        if _job_running:
            cached = load_web_analysis_result() or {}
            cached["job_status"] = "running"
            cached["message"] = "分析進行中，請稍後再試"
            return cached
        _job_running = True

    try:
        logger.info("[analysis] start trigger=%s", trigger)
        running_meta = _set_job_meta(load_web_analysis_result() or {}, "running")
        save_web_analysis_result(running_meta)

        result = run_analysis()
        result = _set_job_meta(result, "idle")
        save_web_analysis_result(result)
        logger.info(
            "[analysis] done trigger=%s bullish=%s bearish=%s warrants=%s settle=%s",
            trigger,
            result.get("bullish_count"),
            result.get("bearish_count"),
            result.get("warrant_count"),
            result.get("settle_date"),
        )
        return result
    except Exception as exc:
        logger.exception("[analysis] failed trigger=%s", trigger)
        failed = _set_job_meta(load_web_analysis_result() or {}, "failed", str(exc))
        save_web_analysis_result(failed)
        raise
    finally:
        with _job_lock:
            _job_running = False


def run_analysis_in_background(trigger: str = "auto") -> bool:
    if _job_running:
        return False

    def _worker():
        try:
            run_analysis_and_persist(trigger=trigger)
        except Exception:
            pass

    threading.Thread(target=_worker, daemon=True, name=f"analysis-{trigger}").start()
    return True


def start_analysis_scheduler() -> None:
    global _scheduler
    if _scheduler is not None:
        return

    _scheduler = BackgroundScheduler(timezone=TW)
    _scheduler.add_job(
        lambda: run_analysis_in_background(trigger="cron"),
        CronTrigger(
            day_of_week="mon-fri",
            hour=settings.AUTO_ANALYSIS_HOUR,
            minute=settings.AUTO_ANALYSIS_MINUTE,
            timezone=TW,
        ),
        id="daily_web_analysis",
        replace_existing=True,
        max_instances=1,
        coalesce=True,
    )
    _scheduler.start()
    logger.info(
        "[analysis] scheduler started at %02d:%02d Asia/Taipei weekdays",
        settings.AUTO_ANALYSIS_HOUR,
        settings.AUTO_ANALYSIS_MINUTE,
    )


def stop_analysis_scheduler() -> None:
    global _scheduler
    if _scheduler is not None:
        _scheduler.shutdown(wait=False)
        _scheduler = None


def maybe_refresh_on_startup() -> None:
    cached = load_web_analysis_result()
    if should_refresh_analysis(cached):
        logger.info("[analysis] startup refresh scheduled")
        run_analysis_in_background(trigger="startup")
