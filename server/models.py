
from datetime import datetime, timezone, timedelta

from sqlalchemy import Boolean, Column, DateTime, Integer, String, Text, Float, ForeignKey

from database import Base
from config import settings


def now_utc():
    return datetime.now(timezone.utc)


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(100), unique=True, index=True, nullable=True)
    full_name = Column(String(100), nullable=False, default="")
    gender = Column(String(20), nullable=True, default="")
    phone = Column(String(30), nullable=True, default="")
    email = Column(String(255), unique=True, index=True, nullable=False)
    password_hash = Column(String(255), nullable=False)

    is_email_verified = Column(Boolean, default=False)

    device_id = Column(String(255), nullable=True)
    device_name = Column(String(255), nullable=True)

    subscription_status = Column(String(30), default="trial")
    plan_type = Column(String(30), default="trial")
    payment_status = Column(String(30), default="unpaid")

    trial_end_at = Column(DateTime, default=lambda: now_utc() + timedelta(days=settings.TRIAL_DAYS))
    subscription_start_at = Column(DateTime, nullable=True)
    subscription_end_at = Column(DateTime, nullable=True)
    free_reason = Column(String(255), nullable=True)

    is_active = Column(Boolean, default=True)
    is_creator = Column(Boolean, default=False)

    created_at = Column(DateTime, default=now_utc)
    updated_at = Column(DateTime, default=now_utc)


class VerificationCode(Base):
    __tablename__ = "verification_codes"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String(255), index=True, nullable=False)
    purpose = Column(String(50), index=True, nullable=False)  # register / reset_password
    code = Column(String(20), nullable=False)
    expires_at = Column(DateTime, nullable=False)
    is_used = Column(Boolean, default=False)
    created_at = Column(DateTime, default=now_utc)


class PaymentReport(Base):
    __tablename__ = "payment_reports"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=False)
    username = Column(String(100), nullable=False)
    email = Column(String(255), nullable=False)

    plan_type = Column(String(30), nullable=False)  # monthly / quarterly / yearly
    amount = Column(Float, nullable=True)
    transfer_last5 = Column(String(10), nullable=True)
    transfer_time = Column(DateTime, nullable=True)
    payer_name = Column(String(100), nullable=True)
    note = Column(Text, nullable=True)

    status = Column(String(30), default="pending")  # pending / approved / rejected
    review_note = Column(Text, nullable=True)
    reviewed_at = Column(DateTime, nullable=True)
    created_at = Column(DateTime, default=now_utc)


class WebAnalysisSnapshot(Base):
    __tablename__ = "web_analysis_snapshots"

    id = Column(Integer, primary_key=True, index=True)
    bullish_json = Column(Text, nullable=False, default="[]")
    bearish_json = Column(Text, nullable=False, default="[]")
    warrants_json = Column(Text, nullable=False, default="[]")
    meta_json = Column(Text, nullable=False, default="{}")
    settle_date = Column(String(20), nullable=True, default="")
    updated_at = Column(DateTime, default=now_utc)


class AbuseLog(Base):
    __tablename__ = "abuse_logs"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String(255), nullable=True)
    ip = Column(String(100), nullable=True)
    country = Column(String(100), nullable=True)
    region = Column(String(100), nullable=True)
    city = Column(String(100), nullable=True)
    location = Column(String(255), nullable=True)
    org = Column(String(255), nullable=True)
    timezone = Column(String(100), nullable=True)
    event_type = Column(String(100), nullable=False)
    detail = Column(Text, nullable=True)
    created_at = Column(DateTime, default=now_utc)


class FeedbackReport(Base):
    __tablename__ = "feedback_reports"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"), nullable=True)
    feedback_id = Column(String(64), unique=True, index=True, nullable=False)
    account = Column(String(100), nullable=True, default="")
    email = Column(String(255), nullable=True, default="")
    topic = Column(String(100), nullable=False, default="")
    content = Column(Text, nullable=False, default="")
    app_version = Column(String(50), nullable=True, default="web")
    device_info = Column(Text, nullable=True, default="")
    status = Column(String(30), default="new")  # new / read / archived
    created_at = Column(DateTime, default=now_utc)
