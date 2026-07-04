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
            "warrants": result.get("warrants") or [],
            "meta": {
                "updated_at": result.get("updated_at", ""),
                "source": result.get("source", "web-strategy"),
                "market": result.get("market", ""),
                "settle_date": result.get("settle_date", ""),
                "bullish_count": result.get("bullish_count", 0),
                "bearish_count": result.get("bearish_count", 0),
                "warrant_count": result.get("warrant_count", 0),
                "strategy": result.get("strategy") or {},
                "job_status": result.get("job_status", "idle"),
                "job_error": result.get("job_error", ""),
                "job_started_at": result.get("job_started_at", ""),
                "job_progress": result.get("job_progress", 0),
                "job_message": result.get("job_message", ""),
                "job_elapsed_sec": result.get("job_elapsed_sec", 0),
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
        row.warrants_json = json.dumps(payload["warrants"], ensure_ascii=False)
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
        return {
            "bullish": json.loads(row.bullish_json or "[]"),
            "bearish": json.loads(row.bearish_json or "[]"),
            "warrants": json.loads(row.warrants_json or "[]"),
            "updated_at": meta.get("updated_at", ""),
            "source": meta.get("source", "web-strategy"),
            "market": meta.get("market", ""),
            "settle_date": meta.get("settle_date", row.settle_date or ""),
            "bullish_count": meta.get("bullish_count", 0),
            "bearish_count": meta.get("bearish_count", 0),
            "warrant_count": meta.get("warrant_count", 0),
            "strategy": meta.get("strategy") or {},
            "job_status": meta.get("job_status", "idle"),
            "job_error": meta.get("job_error", ""),
            "job_started_at": meta.get("job_started_at", ""),
            "job_progress": meta.get("job_progress", 0),
            "job_message": meta.get("job_message", ""),
            "job_elapsed_sec": meta.get("job_elapsed_sec", 0),
        }
    except Exception:
        return None
    finally:
        db.close()
