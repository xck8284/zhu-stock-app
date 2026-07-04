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

STALE_RUNNING_MINUTES = 12

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
    if status == "running":
        result["job_started_at"] = tw_now().strftime("%Y-%m-%d %H:%M:%S")
        result["job_progress"] = 5
        result["job_message"] = "正在抓取台股歷史資料…"
    elif status == "idle":
        result["job_progress"] = 100
        result["job_message"] = "分析完成"
    elif status == "failed":
        result["job_progress"] = 0
        if not result.get("job_message"):
            result["job_message"] = error or "分析失敗"
    return result


def _parse_job_started_at(meta: dict):
    raw = str(meta.get("job_started_at") or "").strip()
    if not raw:
        return None
    for fmt in ["%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"]:
        try:
            return datetime.strptime(raw, fmt).replace(tzinfo=TW)
        except Exception:
            continue
    return None


def recover_stale_running_job(meta: dict | None = None) -> dict:
    """若 running 超過時限，視為 Render 逾時中斷並重置，避免永遠卡住。"""
    meta = dict(meta or load_web_analysis_result() or {})
    if meta.get("job_status") != "running":
        return meta

    started = _parse_job_started_at(meta)
    if started is None:
        meta = _set_job_meta(meta, "failed", "分析中斷（可重新啟動）")
        save_web_analysis_result(meta)
        return meta

    elapsed_min = (tw_now() - started).total_seconds() / 60.0
    meta["job_elapsed_sec"] = int(elapsed_min * 60)
    if elapsed_min >= STALE_RUNNING_MINUTES:
        meta = _set_job_meta(
            meta,
            "failed",
            f"分析逾時（>{STALE_RUNNING_MINUTES} 分鐘），請重新啟動",
        )
        save_web_analysis_result(meta)
    return meta


def get_running_progress(meta: dict | None = None) -> dict:
    meta = dict(meta or load_web_analysis_result() or {})
    if meta.get("job_status") != "running":
        return meta

    started = _parse_job_started_at(meta)
    elapsed_sec = 0
    if started is not None:
        elapsed_sec = max(0, int((tw_now() - started).total_seconds()))

    # 前端進度條用：依已耗時估算（完整分析約 5～10 分鐘）
    estimated_total = 8 * 60
    progress = min(95, max(5, int(elapsed_sec / estimated_total * 100)))
    if elapsed_sec < 60:
        message = "正在抓取台股歷史資料…"
    elif elapsed_sec < 180:
        message = "正在計算週K 與趨勢線…"
    elif elapsed_sec < 360:
        message = "正在套用策略篩選與權證…"
    else:
        message = "即將完成，請稍候…"

    meta["job_elapsed_sec"] = elapsed_sec
    meta["job_progress"] = progress
    meta["job_message"] = message
    return meta


_on_complete_callbacks = []


def register_analysis_complete_callback(callback) -> None:
    if callback not in _on_complete_callbacks:
        _on_complete_callbacks.append(callback)


def _notify_analysis_complete(result: dict) -> None:
    for callback in _on_complete_callbacks:
        try:
            callback(result)
        except Exception:
            logger.exception("[analysis] callback failed")


def run_analysis_and_persist(trigger: str = "manual") -> dict:
    global _job_running

    recover_stale_running_job()

    with _job_lock:
        if _job_running:
            cached = get_running_progress(load_web_analysis_result() or {})
            cached["message"] = "分析進行中，請稍後再試"
            return cached
        _job_running = True

    try:
        logger.info("[analysis] start trigger=%s", trigger)
        running_meta = _set_job_meta(load_web_analysis_result() or {}, "running")
        save_web_analysis_result(running_meta)
        _notify_analysis_complete(running_meta)

        result = run_analysis()
        result = _set_job_meta(result, "idle")
        save_web_analysis_result(result)
        _notify_analysis_complete(result)
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
        _notify_analysis_complete(failed)
        raise
    finally:
        with _job_lock:
            _job_running = False


def run_analysis_in_background(trigger: str = "auto") -> bool:
    recover_stale_running_job()
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
