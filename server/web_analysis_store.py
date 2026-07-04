# -*- coding: utf-8 -*-
"""網頁版分析結果持久化（PostgreSQL / SQLite）。"""

from __future__ import annotations

import json
from datetime import datetime, timezone

from sqlalchemy.orm import Session

from database import SessionLocal
from models import WebAnalysisSnapshot


def _utc_now():
    return datetime.now(timezone.utc)


def save_web_analysis_result(result: dict) -> dict:
    db: Session = SessionLocal()
    try:
        payload = {
            "bullish": result.get("bullish") or [],
            "bearish": result.get("bearish") or [],
            "bullish_keyk": result.get("bullish_keyk") or [],
            "bearish_keyk": result.get("bearish_keyk") or [],
            "warrants": result.get("warrants") or [],
            "meta": {
                "updated_at": result.get("updated_at", ""),
                "source": result.get("source", "web-strategy"),
                "market": result.get("market", ""),
                "settle_date": result.get("settle_date", ""),
                "bullish_count": result.get("bullish_count", 0),
                "bearish_count": result.get("bearish_count", 0),
                "bullish_keyk_count": result.get("bullish_keyk_count", 0),
                "bearish_keyk_count": result.get("bearish_keyk_count", 0),
                "warrant_count": result.get("warrant_count", 0),
                "strategy": result.get("strategy") or {},
                "job_status": result.get("job_status", "idle"),
                "job_error": result.get("job_error", ""),
                "job_started_at": result.get("job_started_at", ""),
                "job_progress": result.get("job_progress", 0),
                "job_message": result.get("job_message", ""),
                "job_elapsed_sec": result.get("job_elapsed_sec", 0),
                "analysis_data_ready": result.get("analysis_data_ready", False),
                "data_stats": result.get("data_stats") or {},
            },
        }

        row = (
            db.query(WebAnalysisSnapshot)
            .order_by(WebAnalysisSnapshot.id.desc())
            .first()
        )
        if row is None:
            row = WebAnalysisSnapshot()
            db.add(row)

        row.bullish_json = json.dumps(payload["bullish"], ensure_ascii=False)
        row.bearish_json = json.dumps(payload["bearish"], ensure_ascii=False)
        row.warrants_json = json.dumps(
            {
                "warrants": payload["warrants"],
                "bullish_keyk": payload["bullish_keyk"],
                "bearish_keyk": payload["bearish_keyk"],
            },
            ensure_ascii=False,
        )
        row.meta_json = json.dumps(payload["meta"], ensure_ascii=False)
        row.settle_date = str(payload["meta"].get("settle_date") or "")
        row.updated_at = _utc_now()
        db.commit()
        db.refresh(row)
        return payload
    finally:
        db.close()


def load_web_analysis_result() -> dict | None:
    db: Session = SessionLocal()
    try:
        row = (
            db.query(WebAnalysisSnapshot)
            .order_by(WebAnalysisSnapshot.id.desc())
            .first()
        )
        if row is None:
            return None

        meta = json.loads(row.meta_json or "{}")
        extra = {}
        try:
            extra = json.loads(row.warrants_json or "{}")
            if not isinstance(extra, dict):
                extra = {"warrants": extra}
        except Exception:
            extra = {}
        warrants = extra.get("warrants") if isinstance(extra, dict) else []
        if warrants is None:
            warrants = extra if isinstance(extra, list) else []
        return {
            "bullish": json.loads(row.bullish_json or "[]"),
            "bearish": json.loads(row.bearish_json or "[]"),
            "bullish_keyk": extra.get("bullish_keyk", []) if isinstance(extra, dict) else [],
            "bearish_keyk": extra.get("bearish_keyk", []) if isinstance(extra, dict) else [],
            "warrants": warrants or [],
            "updated_at": meta.get("updated_at", ""),
            "source": meta.get("source", "web-strategy"),
            "market": meta.get("market", ""),
            "settle_date": meta.get("settle_date", row.settle_date or ""),
            "bullish_count": meta.get("bullish_count", 0),
            "bearish_count": meta.get("bearish_count", 0),
            "bullish_keyk_count": meta.get("bullish_keyk_count", 0),
            "bearish_keyk_count": meta.get("bearish_keyk_count", 0),
            "warrant_count": meta.get("warrant_count", 0),
            "strategy": meta.get("strategy") or {},
            "job_status": meta.get("job_status", "idle"),
            "job_error": meta.get("job_error", ""),
            "job_started_at": meta.get("job_started_at", ""),
            "job_progress": meta.get("job_progress", 0),
            "job_message": meta.get("job_message", ""),
            "job_elapsed_sec": meta.get("job_elapsed_sec", 0),
            "analysis_data_ready": meta.get("analysis_data_ready", False),
            "data_stats": meta.get("data_stats") or {},
        }
    except Exception:
        return None
    finally:
        db.close()
