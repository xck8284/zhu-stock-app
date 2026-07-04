import certifi
import requests
import os
import smtplib
import logging
from email.mime.text import MIMEText

from contextlib import asynccontextmanager
from datetime import datetime, timedelta, timezone
import random
import re
import uuid
from typing import Optional

from fastapi import FastAPI, Depends, HTTPException, Request, Header
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.orm import Session
from sqlalchemy import text, func

import sib_api_v3_sdk
from sib_api_v3_sdk.rest import ApiException

from config import settings
from database import SessionLocal, engine, Base
from models import User, AbuseLog, VerificationCode, PaymentReport, WebAnalysisSnapshot, FeedbackReport
from schemas import (
    MessageResponse,
    SendRegisterCodeRequest, VerifyRegisterCodeRequest,
    LoginByAccountRequest, ForgotPasswordRequest, ResetPasswordRequest,
    AdminGrantFreeRequest, AdminSetPlanRequest, AdminRebindDeviceRequest,
    AdminDeactivateUserRequest, PaymentReportCreateRequest, FeedbackSubmitRequest,
)
from security import hash_password, verify_password, create_access_token, decode_access_token

Base.metadata.create_all(bind=engine)

logger = logging.getLogger("zhu.app")


@asynccontextmanager
async def lifespan(_app: FastAPI):
    from analysis_scheduler import (
        get_running_progress,
        maybe_refresh_on_startup,
        recover_stale_running_job,
        register_analysis_complete_callback,
        start_analysis_scheduler,
        stop_analysis_scheduler,
    )
    from web_analysis_store import load_web_analysis_result

    register_analysis_complete_callback(_apply_web_analysis_result)

    cached = load_web_analysis_result()
    if cached:
        cached = recover_stale_running_job(cached)
        if cached.get("job_status") == "running":
            cached = get_running_progress(cached)
        _apply_web_analysis_result(cached)
        logger.info(
            "loaded web analysis cache settle=%s updated=%s status=%s",
            cached.get("settle_date"),
            cached.get("updated_at"),
            cached.get("job_status"),
        )

    start_analysis_scheduler()
    maybe_refresh_on_startup()
    yield
    stop_analysis_scheduler()


app = FastAPI(title="ZHU STOCK PLATFORM - COMPLETE UPGRADE", version="2.0.0", lifespan=lifespan)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


PLAN_DAYS = {
    "monthly": 30,
    "halfyear": 180,
    "quarterly": 90,
    "yearly": 365,
    "trial": 30,
    "free_grant": 30,
    "none": 0,
}

VALID_PAID_PLANS = ("monthly", "halfyear", "quarterly", "yearly")


def normalize_plan_type(plan_type: str) -> str:
    """對齊桌面版 halfyear；舊版 web 曾用 quarterly 表示半年。"""
    plan = (plan_type or "").strip().lower()
    if plan == "quarterly":
        return "halfyear"
    return plan


def now_utc():
    return datetime.now(timezone.utc)


def normalize_email(email: str) -> str:
    return email.strip().lower()


def normalize_username(username: str) -> str:
    return username.strip()


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def get_client_ip(request: Optional[Request] = None) -> str:
    if request is None:
        return "127.0.0.1"
    forwarded = request.headers.get("x-forwarded-for")
    if forwarded:
        return forwarded.split(",")[0].strip()
    real_ip = request.headers.get("x-real-ip")
    if real_ip:
        return real_ip.strip()
    if request.client:
        return request.client.host
    return "127.0.0.1"


def send_email_code(to_email: str, subject: str, body: str):
    if settings.EMAIL_DEV_MODE:
        return True

    if not settings.BREVO_API_KEY:
        raise HTTPException(status_code=500, detail="BREVO_API_KEY 未設定")

    if not settings.BREVO_FROM_EMAIL:
        raise HTTPException(status_code=500, detail="BREVO_FROM_EMAIL 未設定")

    configuration = sib_api_v3_sdk.Configuration()
    configuration.api_key["api-key"] = settings.BREVO_API_KEY

    api_instance = sib_api_v3_sdk.TransactionalEmailsApi(
        sib_api_v3_sdk.ApiClient(configuration)
    )

    html_body = body.replace("\n", "<br>")

    send_smtp_email = sib_api_v3_sdk.SendSmtpEmail(
        sender={
            "name": "ZHU STOCK",
            "email": settings.BREVO_FROM_EMAIL,
        },
        to=[
            {
                "email": to_email,
            }
        ],
        subject=subject,
        html_content=html_body,
    )

    try:
        api_instance.send_transac_email(send_smtp_email)
        return True
    except ApiException as e:
        raise HTTPException(status_code=500, detail=f"Brevo 寄信失敗: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Brevo 寄信失敗: {str(e)}")


def create_numeric_code(length: int = 6) -> str:
    return "".join(str(random.randint(0, 9)) for _ in range(length))


def deactivate_old_codes(db: Session, email: str, purpose: str):
    db.query(VerificationCode).filter(
        VerificationCode.email == normalize_email(email),
        VerificationCode.purpose == purpose,
        VerificationCode.is_used == False,
    ).update({"is_used": True}, synchronize_session=False)
    db.commit()


def save_verification_code(db: Session, email: str, purpose: str, code: str, minutes: int = 10):
    deactivate_old_codes(db, email, purpose)
    item = VerificationCode(
        email=normalize_email(email),
        purpose=purpose,
        code=code,
        expires_at=now_utc() + timedelta(minutes=minutes),
        is_used=False,
    )
    db.add(item)
    db.commit()
    db.refresh(item)
    return item


def check_verification_code(db: Session, email: str, purpose: str, code: str) -> VerificationCode:
    item = db.query(VerificationCode).filter(
        VerificationCode.email == normalize_email(email),
        VerificationCode.purpose == purpose,
        VerificationCode.code == str(code).strip(),
        VerificationCode.is_used == False,
    ).order_by(VerificationCode.created_at.desc()).first()

    if not item:
        raise HTTPException(status_code=400, detail="驗證碼錯誤或不存在")

    expires_at = _to_utc_naive(item.expires_at)
    now_val = _to_utc_naive(now_utc())

    if expires_at and expires_at < now_val:
        raise HTTPException(status_code=400, detail="驗證碼已過期")

    return item


def mark_code_used(db: Session, item: VerificationCode):
    item.is_used = True
    db.add(item)
    db.commit()


def _to_utc_naive(dt):
    if dt is None:
        return None
    if getattr(dt, "tzinfo", None) is not None:
        return dt.astimezone(timezone.utc).replace(tzinfo=None)
    return dt


def compute_license_status(user: User):
    now = _to_utc_naive(now_utc())
    allowed = False
    status = user.subscription_status or "none"
    plan_type = user.plan_type or "none"
    days_left = 0
    end_at = None
    label = "未授權"

    if not user.is_active:
        return {
            "allowed": False,
            "subscription_status": "inactive",
            "plan_type": plan_type,
            "days_left": 0,
            "end_at": None,
            "label": "帳號已停用",
        }

    trial_end_at = _to_utc_naive(getattr(user, "trial_end_at", None))
    subscription_end_at = _to_utc_naive(getattr(user, "subscription_end_at", None))

    if status == "trial":
        end_at = trial_end_at
        if end_at and end_at > now:
            allowed = True
            days_left = max((end_at - now).days, 0)
            label = f"免費試用 30 天（剩餘 {days_left} 天）"
        else:
            allowed = False
            status = "expired"
            label = "免費試用已到期"
    elif status in ("active", "free_grant"):
        end_at = subscription_end_at
        if end_at and end_at > now:
            allowed = True
            days_left = max((end_at - now).days, 0)
            if plan_type == "free_grant":
                label = f"活動贈送（剩餘 {days_left} 天）"
            elif plan_type == "monthly":
                label = f"月訂閱（剩餘 {days_left} 天）"
            elif plan_type in ("halfyear", "quarterly"):
                label = f"半年訂閱（剩餘 {days_left} 天）"
            elif plan_type == "yearly":
                label = f"年訂閱（剩餘 {days_left} 天）"
            else:
                label = f"授權有效（剩餘 {days_left} 天）"
        else:
            allowed = False
            status = "expired"
            label = "授權已到期"
    else:
        allowed = False
        label = "未授權"

    return {
        "allowed": allowed,
        "subscription_status": status,
        "plan_type": plan_type,
        "days_left": days_left,
        "end_at": end_at.isoformat() if hasattr(end_at, "isoformat") and end_at else None,
        "label": label,
    }


def auto_bind_or_validate_device(user: User, device_id: Optional[str], device_name: Optional[str], db: Session):
    device_id = (device_id or "").strip()
    device_name = (device_name or "").strip()

    if not device_id:
        return

    if not user.device_id:
        user.device_id = device_id
        if device_name:
            user.device_name = device_name
        db.add(user)
        db.commit()
        return

    if user.device_id != device_id:
        raise HTTPException(status_code=403, detail="此帳號目前綁定於其他裝置，請聯絡管理員協助轉移授權")


def get_bearer_token(authorization: Optional[str]) -> str:
    if not authorization:
        raise HTTPException(status_code=401, detail="缺少 token")
    if not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="token 格式錯誤")
    return authorization.replace("Bearer ", "", 1).strip()


def get_current_user(authorization: Optional[str], db: Session) -> User:
    token = get_bearer_token(authorization)
    payload = decode_access_token(token)
    if not payload:
        raise HTTPException(status_code=401, detail="token 無效")
    sub = payload.get("sub")
    if not sub:
        raise HTTPException(status_code=401, detail="token 無效")
    user = db.query(User).filter(User.email == sub).first()
    if not user:
        raise HTTPException(status_code=401, detail="找不到使用者")
    return user


def get_current_creator(authorization: Optional[str], db: Session) -> User:
    user = get_current_user(authorization, db)
    if not user.is_creator:
        raise HTTPException(status_code=403, detail="非創作者權限")
    return user


def require_active_license(user: User) -> None:
    if user.is_creator:
        return
    lic = compute_license_status(user)
    if not lic.get("allowed"):
        raise HTTPException(status_code=403, detail=lic.get("label") or "會員資格已到期")


def ensure_default_creator(db: Session):
    if not settings.CREATOR_ADMIN_EMAIL:
        return
    email = normalize_email(settings.CREATOR_ADMIN_EMAIL)
    existed = db.query(User).filter(User.email == email).first()
    if existed:
        if not existed.is_creator:
            existed.is_creator = True
            db.add(existed)
            db.commit()
        return

    username = settings.CREATOR_ADMIN_USERNAME.strip() or "admin"
    base_username = username
    idx = 1
    while db.query(User).filter(User.username == username).first():
        idx += 1
        username = f"{base_username}{idx}"

    user = User(
        username=username,
        full_name="Creator Admin",
        gender="",
        phone="",
        email=email,
        password_hash=hash_password(settings.CREATOR_ADMIN_PASSWORD),
        is_email_verified=True,
        subscription_status="active",
        plan_type="yearly",
        payment_status="approved",
        subscription_start_at=now_utc(),
        subscription_end_at=now_utc() + timedelta(days=3650),
        is_active=True,
        is_creator=True,
    )
    db.add(user)
    db.commit()


def ensure_db_columns():
    dialect = engine.dialect.name
    with engine.begin() as conn:
        if dialect == "sqlite":
            existing = {row[1] for row in conn.execute(text("PRAGMA table_info(users)")).fetchall()}
            needed = {
                "username": "ALTER TABLE users ADD COLUMN username VARCHAR(100)",
                "is_email_verified": "ALTER TABLE users ADD COLUMN is_email_verified BOOLEAN DEFAULT 0",
                "plan_type": "ALTER TABLE users ADD COLUMN plan_type VARCHAR(30) DEFAULT 'trial'",
                "subscription_start_at": "ALTER TABLE users ADD COLUMN subscription_start_at DATETIME",
                "subscription_end_at": "ALTER TABLE users ADD COLUMN subscription_end_at DATETIME",
                "free_reason": "ALTER TABLE users ADD COLUMN free_reason VARCHAR(255)",
                "updated_at": "ALTER TABLE users ADD COLUMN updated_at DATETIME",
            }
            for col, sql in needed.items():
                if col not in existing:
                    conn.execute(text(sql))

        Base.metadata.create_all(bind=engine)


ensure_db_columns()
with SessionLocal() as db:
    ensure_default_creator(db)


# =========================
# 後台 Email API：供桌面版 APP 測試信、到期提醒、付款/註冊通知使用
# 同時支援 Render 既有 SMTP_* 與前台新版 ZHU_SMTP_* 變數
# =========================
def _env_first(*names, default=""):
    for name in names:
        val = os.getenv(name)
        if val is not None and str(val).strip() != "":
            return str(val).strip()
    return default


def send_backend_smtp_email(to_email: str, subject: str, body: str):
    """
    Render Free 會封鎖 SMTP 25/465/587，所以後台通知信改回 Email API（Brevo）。
    需要 Render Environment：
    - BREVO_API_KEY
    - BREVO_FROM_EMAIL
    """
    to_email = (to_email or "").strip()
    subject = (subject or "ZHU STOCK 通知").strip()
    body = body or ""

    if not to_email:
        raise HTTPException(status_code=400, detail="缺少收件人 Email")

    if not settings.BREVO_API_KEY:
        raise HTTPException(status_code=500, detail="BREVO_API_KEY 未設定，請到 Render Environment 設定")

    if not settings.BREVO_FROM_EMAIL:
        raise HTTPException(status_code=500, detail="BREVO_FROM_EMAIL 未設定，請到 Render Environment 設定")

    configuration = sib_api_v3_sdk.Configuration()
    configuration.api_key["api-key"] = settings.BREVO_API_KEY

    api_instance = sib_api_v3_sdk.TransactionalEmailsApi(
        sib_api_v3_sdk.ApiClient(configuration)
    )

    html_body = str(body).replace("\n", "<br>")

    send_smtp_email = sib_api_v3_sdk.SendSmtpEmail(
        sender={
            "name": "ZHU STOCK",
            "email": settings.BREVO_FROM_EMAIL,
        },
        to=[
            {
                "email": to_email,
            }
        ],
        subject=subject,
        html_content=html_body,
    )

    try:
        api_instance.send_transac_email(send_smtp_email)
        print(f"[EMAIL_OK] Brevo API 已寄出 to={to_email} subject={subject}", flush=True)
        return True
    except ApiException as e:
        print(f"[EMAIL_ERROR] Brevo ApiException: {e}", flush=True)
        raise HTTPException(status_code=500, detail=f"Brevo API 寄信失敗：{str(e)}")
    except Exception as e:
        print(f"[EMAIL_ERROR] {type(e).__name__}: {e}", flush=True)
        raise HTTPException(status_code=500, detail=f"Brevo API 寄信失敗：{type(e).__name__}: {e}")

def _extract_email_payload(data: dict):
    to_email = (
        data.get("to")
        or data.get("email")
        or data.get("to_email")
        or data.get("recipient")
        or data.get("receiver")
        or ""
    )
    subject = data.get("subject") or data.get("title") or "ZHU STOCK 通知"
    body = data.get("body") or data.get("message") or data.get("content") or ""
    return str(to_email).strip(), str(subject).strip(), str(body)


def _require_creator_if_token(authorization: Optional[str], db: Session):
    # 桌面版管理員通常會帶 Bearer token；有 token 就驗證創作者權限。
    if authorization:
        return get_current_creator(authorization, db)
    return None


@app.post("/admin/send-email")
def admin_send_email_api(
    data: dict,
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    _require_creator_if_token(authorization, db)
    to_email, subject, body = _extract_email_payload(data)
    send_backend_smtp_email(to_email, subject, body)
    return {"success": True, "message": "Email 已送出", "to": to_email}


@app.post("/admin/notify-email")
def admin_notify_email_api(
    data: dict,
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    _require_creator_if_token(authorization, db)
    to_email, subject, body = _extract_email_payload(data)
    send_backend_smtp_email(to_email, subject, body)
    return {"success": True, "message": "通知信已送出", "to": to_email}


@app.post("/api/admin/notify-email")
def api_admin_notify_email_api(
    data: dict,
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    _require_creator_if_token(authorization, db)
    to_email, subject, body = _extract_email_payload(data)
    send_backend_smtp_email(to_email, subject, body)
    return {"success": True, "message": "通知信已送出", "to": to_email}


@app.post("/notify-email")
def notify_email_api(
    data: dict,
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    _require_creator_if_token(authorization, db)
    to_email, subject, body = _extract_email_payload(data)
    send_backend_smtp_email(to_email, subject, body)
    return {"success": True, "message": "通知信已送出", "to": to_email}


@app.get("/")
def root():
    return {"success": True, "message": "ZHU STOCK PLATFORM API is running", "version": "2.0.0"}


@app.post("/auth/send-register-code")
def auth_send_register_code(data: SendRegisterCodeRequest, db: Session = Depends(get_db)):
    email = normalize_email(data.email)
    username = normalize_username(data.username)

    if data.password != data.confirm_password:
        raise HTTPException(status_code=400, detail="密碼與確認密碼不一致")

    if len(data.password) < 6:
        raise HTTPException(status_code=400, detail="密碼長度至少 6 碼")

    if not re.fullmatch(r"[A-Za-z0-9_.\-@]+", username):
        raise HTTPException(status_code=400, detail="帳號只能使用英數與 . _ - @")

    if db.query(User).filter(User.email == email).first():
        raise HTTPException(status_code=400, detail="此 Email 已註冊")

    if db.query(User).filter(User.username == username).first():
        raise HTTPException(status_code=400, detail="此帳號已存在")

    code = create_numeric_code(6)
    save_verification_code(db, email, "register", code, minutes=10)

    body = (
        f"您好，您的 ZHU STOCK 註冊驗證碼為：{code}\n\n"
        f"此驗證碼 10 分鐘內有效。若非本人操作，請忽略此信件。"
    )
    send_email_code(email, "ZHU STOCK 註冊驗證碼", body)

    response = {"success": True, "message": "驗證碼已寄出", "email": email}
    if settings.EMAIL_DEV_MODE:
        response["dev_code"] = code
    return response


@app.post("/auth/verify-register-code")
def auth_verify_register_code(
    data: VerifyRegisterCodeRequest,
    request: Request,
    db: Session = Depends(get_db),
):
    email = normalize_email(data.email)
    username = normalize_username(data.username)

    if data.password != data.confirm_password:
        raise HTTPException(status_code=400, detail="密碼與確認密碼不一致")

    if db.query(User).filter(User.email == email).first():
        raise HTTPException(status_code=400, detail="此 Email 已註冊")

    if db.query(User).filter(User.username == username).first():
        raise HTTPException(status_code=400, detail="此帳號已存在")

    item = check_verification_code(db, email, "register", data.code)
    mark_code_used(db, item)

    user = User(
        username=username,
        full_name=username,
        gender="",
        phone=(data.phone or "").strip(),
        email=email,
        password_hash=hash_password(data.password),
        device_id=(data.device_id or "").strip() or None,
        device_name=(data.device_name or "").strip() or None,
        is_email_verified=True,
        subscription_status="trial",
        plan_type="trial",
        payment_status="unpaid",
        trial_end_at=now_utc() + timedelta(days=settings.TRIAL_DAYS),
        subscription_start_at=None,
        subscription_end_at=None,
        is_active=True,
        is_creator=False,
        updated_at=now_utc(),
    )
    db.add(user)
    db.commit()
    db.refresh(user)

    token = create_access_token(
        {"sub": user.email, "username": user.username, "is_creator": user.is_creator},
        expires_delta=timedelta(hours=settings.JWT_EXPIRE_HOURS),
    )
    lic = compute_license_status(user)

    db.add(AbuseLog(
        email=user.email,
        ip=get_client_ip(request),
        event_type="register_success",
        detail=f"username={user.username}",
    ))
    db.commit()

    return {
        "success": True,
        "message": "註冊成功，已開通免費試用 30 天",
        "access_token": token,
        "token_type": "bearer",
        "subscription_status": lic["subscription_status"],
        "plan_type": lic["plan_type"],
        "trial_end_at": user.trial_end_at,
        "username": user.username,
        "email": user.email,
    }


@app.get("/api/bullish")
def get_bullish():
    return [
        {
            "symbol": "2330",
            "name": "台積電",
            "score": 128,
            "star": "★★★★★",
            "bias": 12.5
        },
        {
            "symbol": "3017",
            "name": "奇鋐",
            "score": 115,
            "star": "★★★★",
            "bias": 9.2
        }
    ]


@app.post("/auth/login")
def auth_login(
    data: LoginByAccountRequest,
    request: Request,
    db: Session = Depends(get_db),
):
    account = data.account.strip()
    user = db.query(User).filter((User.username == account) | (User.email == normalize_email(account))).first()
    if not user:
        raise HTTPException(status_code=404, detail="帳號不存在")

    if not verify_password(data.password, user.password_hash):
        raise HTTPException(status_code=400, detail="密碼錯誤")

    if not user.is_active:
        raise HTTPException(status_code=403, detail="帳號已停用")

    token = create_access_token(
        {"sub": user.email, "username": user.username, "is_creator": user.is_creator},
        expires_delta=timedelta(hours=settings.JWT_EXPIRE_HOURS),
    )
    lic = compute_license_status(user)

    db.add(AbuseLog(
        email=user.email,
        ip=get_client_ip(request),
        event_type="login_success",
        detail=f"username={user.username}, device_id={data.device_id or ''}",
    ))
    db.commit()

    return {
        "success": True,
        "message": "登入成功",
        "access_token": token,
        "token_type": "bearer",
        "user": {
            "username": user.username,
            "email": user.email,
            "phone": user.phone,
        },
        "subscription_status": lic["subscription_status"],
        "plan_type": lic["plan_type"],
        "trial_end_at": user.trial_end_at,
        "subscription_end_at": user.subscription_end_at,
        "license_label": lic["label"],
        "allowed": lic["allowed"],
        "days_left": lic["days_left"],
    }


@app.post("/auth/forgot-password")
def auth_forgot_password(data: ForgotPasswordRequest, db: Session = Depends(get_db)):
    account = data.account.strip()
    user = db.query(User).filter((User.username == account) | (User.email == normalize_email(account))).first()
    if not user:
        raise HTTPException(status_code=404, detail="找不到帳號")

    code = create_numeric_code(6)
    save_verification_code(db, user.email, "reset_password", code, minutes=10)

    body = (
        f"您好，您的 ZHU STOCK 重設密碼驗證碼為：{code}\n\n"
        f"此驗證碼 10 分鐘內有效。若非本人操作，請忽略此信件。"
    )
    send_email_code(user.email, "ZHU STOCK 重設密碼驗證碼", body)

    response = {"success": True, "message": "重設驗證碼已寄出", "email": user.email}
    if settings.EMAIL_DEV_MODE:
        response["dev_code"] = code
    return response


@app.post("/auth/reset-password")
def auth_reset_password(data: ResetPasswordRequest, db: Session = Depends(get_db)):
    account = data.account.strip()
    user = db.query(User).filter((User.username == account) | (User.email == normalize_email(account))).first()
    if not user:
        raise HTTPException(status_code=404, detail="找不到帳號")

    if data.new_password != data.confirm_new_password:
        raise HTTPException(status_code=400, detail="新密碼與確認新密碼不一致")

    item = check_verification_code(db, user.email, "reset_password", data.code)
    mark_code_used(db, item)

    user.password_hash = hash_password(data.new_password)
    user.updated_at = now_utc()
    db.add(user)
    db.commit()

    return {"success": True, "message": "密碼已重設，請重新登入"}


@app.get("/auth/me")
def auth_me(
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    user = get_current_user(authorization, db)
    lic = compute_license_status(user)
    return {
        "success": True,
        "user": {
            "username": user.username,
            "email": user.email,
            "phone": user.phone,
            "is_creator": user.is_creator,
        },
        "license": lic,
    }


@app.get("/license/status")
def license_status(
    account: Optional[str] = None,
    device_id: Optional[str] = None,
    device_name: Optional[str] = None,
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    """
    相容兩種呼叫方式：
    1. 桌面端：GET /license/status?account=...&device_id=...
    2. 已登入 API：Header Authorization: Bearer xxx
    """
    user = None

    if account:
        account_norm = account.strip()
        user = db.query(User).filter(
            (User.username == account_norm) | (User.email == normalize_email(account_norm))
        ).first()

        if not user:
            raise HTTPException(status_code=404, detail="帳號不存在")

        if not user.is_active:
            raise HTTPException(status_code=401, detail="帳號已停用")

        if device_id:
            if hasattr(user, "device_id"):
                user.device_id = device_id
            if hasattr(user, "device_name") and device_name:
                user.device_name = device_name
            if hasattr(user, "updated_at"):
                user.updated_at = now_utc()
            db.add(user)
            db.commit()
            db.refresh(user)

    else:
        user = get_current_user(authorization, db)

    lic = compute_license_status(user)
    return {
        "success": True,
        "allowed": lic["allowed"],
        "subscription_status": lic["subscription_status"],
        "plan_type": lic["plan_type"],
        "days_left": lic["days_left"],
        "end_at": lic["end_at"],
        "label": lic["label"],
        "trial_end_at": user.trial_end_at,
        "subscription_end_at": user.subscription_end_at,
        "account": user.username,
        "email": user.email,
        "device_id": getattr(user, "device_id", None),
    }


@app.post("/payments/report")
def payments_report(
    data: PaymentReportCreateRequest,
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    user = get_current_user(authorization, db)
    plan_type = normalize_plan_type(data.plan_type)
    if plan_type not in VALID_PAID_PLANS:
        raise HTTPException(status_code=400, detail="方案類型錯誤")

    item = PaymentReport(
        user_id=user.id,
        username=user.username,
        email=user.email,
        plan_type=plan_type,
        amount=data.amount,
        transfer_last5=data.transfer_last5,
        transfer_time=data.transfer_time,
        payer_name=(data.payer_name or "").strip(),
        note=(data.note or "").strip(),
        status="pending",
        created_at=now_utc(),
    )
    db.add(item)
    db.commit()
    db.refresh(item)
    return {"success": True, "message": "付款回報已送出，待管理員審核", "report_id": item.id}


@app.post("/feedback/submit")
def feedback_submit(
    data: FeedbackSubmitRequest,
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    user = get_current_user(authorization, db)
    feedback_id = (data.feedback_id or "").strip() or str(uuid.uuid4())
    existing = db.query(FeedbackReport).filter(FeedbackReport.feedback_id == feedback_id).first()
    if existing:
        return {"success": True, "message": "回饋已收到", "feedback_id": feedback_id}

    item = FeedbackReport(
        user_id=user.id,
        feedback_id=feedback_id,
        account=user.username or "",
        email=user.email or "",
        topic=(data.topic or "").strip(),
        content=(data.content or "").strip(),
        app_version=(data.app_version or "web").strip(),
        device_info=(data.device_info or "").strip(),
        status="new",
        created_at=now_utc(),
    )
    db.add(item)
    db.commit()
    return {"success": True, "message": "回饋已送出，感謝您的意見", "feedback_id": feedback_id}


@app.get("/admin/feedback")
def admin_feedback_list(
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    _ = get_current_creator(authorization, db)
    rows = db.query(FeedbackReport).order_by(FeedbackReport.created_at.desc()).limit(200).all()
    items = []
    for r in rows:
        items.append({
            "id": r.id,
            "feedback_id": r.feedback_id,
            "account": r.account,
            "email": r.email,
            "topic": r.topic,
            "content": r.content,
            "app_version": r.app_version,
            "status": r.status,
            "created_at": r.created_at,
        })
    return {"success": True, "count": len(items), "items": items}


@app.get("/admin/users")
def admin_users(
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    _ = get_current_creator(authorization, db)
    users = db.query(User).order_by(User.created_at.desc()).all()
    pending_rows = (
        db.query(PaymentReport.user_id, func.count(PaymentReport.id))
        .filter(PaymentReport.status == "pending")
        .group_by(PaymentReport.user_id)
        .all()
    )
    pending_map = {uid: cnt for uid, cnt in pending_rows}
    rows = []
    for u in users:
        lic = compute_license_status(u)
        rows.append({
            "id": u.id,
            "username": u.username,
            "email": u.email,
            "phone": u.phone,
            "subscription_status": lic["subscription_status"],
            "plan_type": lic["plan_type"],
            "days_left": lic["days_left"],
            "trial_end_at": u.trial_end_at,
            "subscription_end_at": u.subscription_end_at,
            "free_reason": u.free_reason,
            "is_active": u.is_active,
            "is_creator": u.is_creator,
            "device_id": u.device_id,
            "device_name": u.device_name,
            "created_at": u.created_at,
            "pending_review": pending_map.get(u.id, 0),
        })
    return {"success": True, "count": len(rows), "items": rows}


@app.post("/admin/grant-free")
def admin_grant_free(
    data: AdminGrantFreeRequest,
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    admin = get_current_creator(authorization, db)
    user = db.query(User).filter((User.username == data.account) | (User.email == normalize_email(data.account))).first()
    if not user:
        raise HTTPException(status_code=404, detail="找不到使用者")

    start = now_utc()
    end = start + timedelta(days=data.free_days)

    user.subscription_status = "free_grant"
    user.plan_type = "free_grant"
    user.subscription_start_at = start
    user.subscription_end_at = end
    user.payment_status = "free"
    user.free_reason = (data.reason or "").strip()
    user.updated_at = now_utc()
    db.add(user)
    db.add(AbuseLog(
        email=user.email,
        ip="admin",
        event_type="grant_free",
        detail=f"admin={admin.email}, days={data.free_days}, reason={data.reason}",
    ))
    db.commit()

    return {
        "success": True,
        "message": f"已開通免費資格 {data.free_days} 天",
        "account": data.account,
        "plan_type": "free_grant",
        "subscription_end_at": end,
    }


@app.post("/admin/set-plan")
def admin_set_plan(
    data: AdminSetPlanRequest,
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    admin = get_current_creator(authorization, db)
    user = db.query(User).filter((User.username == data.account) | (User.email == normalize_email(data.account))).first()
    if not user:
        raise HTTPException(status_code=404, detail="找不到使用者")

    if data.plan_type not in (
        "monthly",
        "halfyear",
        "quarterly",
        "yearly",
        "trial",
        "free_grant",
        "none"
    ):
        raise HTTPException(
            status_code=400,
            detail="方案類型錯誤"
        )

    plan_type = normalize_plan_type(data.plan_type) if data.plan_type in VALID_PAID_PLANS else data.plan_type
    start = now_utc()
    end = start + timedelta(days=PLAN_DAYS[plan_type])

    user.subscription_status = "active"
    user.plan_type = plan_type
    user.subscription_start_at = start
    user.subscription_end_at = end
    user.payment_status = "approved"
    user.free_reason = None
    user.updated_at = now_utc()
    db.add(user)
    db.add(AbuseLog(
        email=user.email,
        ip="admin",
        event_type="set_plan",
        detail=f"admin={admin.email}, plan={data.plan_type}",
    ))
    db.commit()

    return {
        "success": True,
        "message": f"已開通 {data.plan_type}",
        "account": data.account,
        "plan_type": data.plan_type,
        "subscription_end_at": end,
    }


@app.post("/admin/rebind-device")
def admin_rebind_device(
    data: AdminRebindDeviceRequest,
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    _ = get_current_creator(authorization, db)
    user = db.query(User).filter((User.username == data.account) | (User.email == normalize_email(data.account))).first()
    if not user:
        raise HTTPException(status_code=404, detail="找不到使用者")

    user.device_id = data.device_id.strip()
    user.device_name = (data.device_name or "").strip()
    user.updated_at = now_utc()
    db.add(user)
    db.commit()

    return {"success": True, "message": "裝置已重綁", "account": data.account}


@app.post("/admin/deactivate-user")
def admin_deactivate_user(
    data: AdminDeactivateUserRequest,
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    _ = get_current_creator(authorization, db)
    user = db.query(User).filter((User.username == data.account) | (User.email == normalize_email(data.account))).first()
    if not user:
        raise HTTPException(status_code=404, detail="找不到使用者")

    user.is_active = data.is_active
    user.updated_at = now_utc()
    db.add(user)
    db.commit()

    return {"success": True, "message": "帳號狀態已更新", "account": data.account, "is_active": data.is_active}


@app.get("/admin/payment-reports")
def admin_payment_reports(
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    _ = get_current_creator(authorization, db)
    rows = db.query(PaymentReport).order_by(PaymentReport.created_at.desc()).all()
    items = []
    for r in rows:
        items.append({
            "id": r.id,
            "username": r.username,
            "email": r.email,
            "plan_type": r.plan_type,
            "amount": r.amount,
            "transfer_last5": r.transfer_last5,
            "transfer_time": r.transfer_time,
            "payer_name": r.payer_name,
            "note": r.note,
            "status": r.status,
            "review_note": r.review_note,
            "created_at": r.created_at,
            "reviewed_at": r.reviewed_at,
        })
    return {"success": True, "count": len(items), "items": items}


@app.post("/admin/approve-payment-report/{report_id}")
def admin_approve_payment_report(
    report_id: int,
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    admin = get_current_creator(authorization, db)
    report = db.query(PaymentReport).filter(PaymentReport.id == report_id).first()
    if not report:
        raise HTTPException(status_code=404, detail="找不到付款回報")

    user = db.query(User).filter(User.id == report.user_id).first()
    if not user:
        raise HTTPException(status_code=404, detail="找不到使用者")

    if report.plan_type not in VALID_PAID_PLANS:
        raise HTTPException(status_code=400, detail="付款方案錯誤")

    plan_type = normalize_plan_type(report.plan_type)
    start = now_utc()
    end = start + timedelta(days=PLAN_DAYS[plan_type])

    report.status = "approved"
    report.review_note = f"approved by {admin.email}"
    report.reviewed_at = now_utc()
    db.add(report)

    user.subscription_status = "active"
    user.plan_type = plan_type
    user.subscription_start_at = start
    user.subscription_end_at = end
    user.payment_status = "approved"
    user.free_reason = None
    user.updated_at = now_utc()
    db.add(user)
    db.commit()

    return {"success": True, "message": "付款已核准並開通方案", "plan_type": report.plan_type, "subscription_end_at": end}


@app.post("/admin/reject-payment-report/{report_id}")
def admin_reject_payment_report(
    report_id: int,
    note: str = "",
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    admin = get_current_creator(authorization, db)
    report = db.query(PaymentReport).filter(PaymentReport.id == report_id).first()
    if not report:
        raise HTTPException(status_code=404, detail="找不到付款回報")

    report.status = "rejected"
    report.review_note = note or f"rejected by {admin.email}"
    report.reviewed_at = now_utc()
    db.add(report)
    db.commit()
    return {"success": True, "message": "付款回報已駁回"}


@app.post("/register")
def register_legacy(data: VerifyRegisterCodeRequest, request: Request, db: Session = Depends(get_db)):
    email = normalize_email(data.email)
    if db.query(User).filter(User.email == email).first():
        return {"success": False, "message": "此 Email 已註冊"}
    user = User(
        username=normalize_username(data.username or email.split("@")[0]),
        full_name=data.username,
        gender="",
        phone=data.phone,
        email=email,
        password_hash=hash_password(data.password),
        device_id=(data.device_id or "").strip() or None,
        device_name=(data.device_name or "").strip() or None,
        is_email_verified=True,
        subscription_status="trial",
        plan_type="trial",
        payment_status="unpaid",
        trial_end_at=now_utc() + timedelta(days=settings.TRIAL_DAYS),
        is_active=True,
        is_creator=False,
        updated_at=now_utc(),
    )
    db.add(user)
    db.commit()
    return {"success": True, "message": "註冊成功，已開通免費試用 30 天"}


@app.post("/login")
def login_legacy(data: LoginByAccountRequest, request: Request, db: Session = Depends(get_db)):
    return auth_login(data, request, db)


@app.post("/license/check")
def license_check_legacy(
    data: LoginByAccountRequest,
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    user = db.query(User).filter((User.username == data.account) | (User.email == normalize_email(data.account))).first()
    if not user:
        return {"success": False, "allowed": False, "message": "找不到使用者", "subscription_status": "none", "payment_status": "unpaid"}
    lic = compute_license_status(user)
    return {
        "success": True,
        "allowed": lic["allowed"],
        "message": lic["label"],
        "subscription_status": lic["subscription_status"],
        "plan_type": lic["plan_type"],
        "payment_status": user.payment_status,
        "trial_end_at": user.trial_end_at,
        "subscription_end_at": user.subscription_end_at,
        "days_left": lic["days_left"],
    }

# =========================
# 管理員取得付款回報列表
# =========================
@app.get("/mobile/admin/payment-reports")
def admin_payment_reports(
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    user = get_current_user(authorization, db)

    if not user.is_creator:
        raise HTTPException(status_code=403, detail="沒有權限")

    reports = (
        db.query(PaymentReport)
        .order_by(PaymentReport.created_at.desc())
        .all()
    )

    result = []

    for r in reports:
        result.append({
            "id": r.id,
            "username": r.username,
            "email": r.email,
            "plan_type": r.plan_type,
            "amount": r.amount,
            "transfer_last5": r.transfer_last5,
            "transfer_time": r.transfer_time,
            "payer_name": r.payer_name,
            "note": r.note,
            "status": r.status,
            "created_at": r.created_at,
        })

    return {
        "success": True,
        "reports": result
    }


# =========================
# 管理員審核付款
# =========================
@app.post("/admin/approve-payment/{report_id}")
def approve_payment(
    report_id: int,
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    admin = get_current_user(authorization, db)

    if not admin.is_creator:
        raise HTTPException(status_code=403, detail="沒有權限")

    report = (
        db.query(PaymentReport)
        .filter(PaymentReport.id == report_id)
        .first()
    )

    if not report:
        raise HTTPException(status_code=404, detail="找不到付款資料")

    report.status = "approved"

    user = (
        db.query(User)
        .filter(User.id == report.user_id)
        .first()
    )

    if user:

        now = datetime.utcnow()

        if report.plan_type == "month":
            days = 30

        elif report.plan_type == "half_year":
            days = 180

        elif report.plan_type == "year":
            days = 365

        else:
            days = 30

        if user.vip_expire_at and user.vip_expire_at > now:
            user.vip_expire_at = user.vip_expire_at + timedelta(days=days)
        else:
            user.vip_expire_at = now + timedelta(days=days)

        user.is_vip = True

    db.commit()

    return {
        "success": True,
        "message": "已完成審核與開通"
    }

FALLBACK_BULLISH = [
    {
        "stock_id": "2330",
        "name": "台積電",
        "stars": "★★★★★",
        "strong_score": 128,
        "bias": "12%",
    },
    {
        "stock_id": "3017",
        "name": "奇鋐",
        "stars": "★★★★☆",
        "strong_score": 115,
        "bias": "18%",
    },
]

BULLISH_DATA = []
BEARISH_DATA = []
WARRANT_DATA = []

WEB_BULLISH_DATA = []
WEB_BEARISH_DATA = []
WEB_WARRANT_DATA = []
WEB_ANALYSIS_META = {"updated_at": "", "source": "web"}


def _format_bias(value):
    if value is None or value == "":
        return ""
    text = str(value)
    return text if "%" in text else f"{text}%"


def _normalize_stock_items(items):
    normalized = []
    for item in items or []:
        if not isinstance(item, dict):
            continue
        stock_id = str(item.get("stock_id") or item.get("code") or item.get("symbol") or "").strip()
        if not stock_id:
            continue
        normalized.append({
            "stock_id": stock_id,
            "name": str(item.get("name") or "").strip(),
            "stars": str(item.get("stars") or item.get("star") or "").strip(),
            "strong_score": item.get("strong_score", item.get("score", 0)),
            "bias": _format_bias(item.get("bias")),
        })
    return normalized


def _normalize_warrant_items(items):
    normalized = []
    for item in items or []:
        if not isinstance(item, dict):
            continue
        code = str(item.get("code") or item.get("stock_id") or "").strip()
        if not code:
            continue
        normalized.append({
            "stock_id": code,
            "code": code,
            "name": str(item.get("name") or "").strip(),
            "type": str(item.get("type") or "").strip(),
            "issuer": str(item.get("issuer") or item.get("broker") or "").strip(),
            "broker": str(item.get("issuer") or item.get("broker") or "").strip(),
            "strike": item.get("strike", ""),
            "price": item.get("price_text") or item.get("price", ""),
        })
    return normalized


@app.post("/admin/upload-stock-results")
def upload_stock_results(data: dict):
    global BULLISH_DATA, BEARISH_DATA, WARRANT_DATA

    BULLISH_DATA = _normalize_stock_items(data.get("bullish", []))
    BEARISH_DATA = _normalize_stock_items(data.get("bearish", []))
    WARRANT_DATA = _normalize_warrant_items(data.get("warrants", []))

    return {
        "success": True,
        "bullish_count": len(BULLISH_DATA),
        "bearish_count": len(BEARISH_DATA),
        "warrant_count": len(WARRANT_DATA),
        "updated_at": data.get("updated_at", ""),
        "settle_date": data.get("settle_date", ""),
    }


@app.post("/upload/bullish")
def upload_bullish(data: dict):

    global BULLISH_DATA

    BULLISH_DATA = data.get("items", [])

    return {
        "success": True,
        "count": len(BULLISH_DATA)
    }


@app.get("/stocks/bullish")
def get_bullish_stocks():
    return {
        "items": BULLISH_DATA if BULLISH_DATA else FALLBACK_BULLISH
    }

from analysis_scheduler import run_analysis_in_background


def _apply_web_analysis_result(result):
    global WEB_BULLISH_DATA, WEB_BEARISH_DATA, WEB_WARRANT_DATA, WEB_ANALYSIS_META

    WEB_BULLISH_DATA = _normalize_stock_items(result.get("bullish", []))
    WEB_BEARISH_DATA = _normalize_stock_items(result.get("bearish", []))
    WEB_WARRANT_DATA = _normalize_warrant_items(result.get("warrants", []))
    WEB_ANALYSIS_META = {
        "updated_at": result.get("updated_at", ""),
        "settle_date": result.get("settle_date", ""),
        "source": result.get("source", "web-strategy"),
        "market": result.get("market", ""),
        "bullish_count": len(WEB_BULLISH_DATA),
        "bearish_count": len(WEB_BEARISH_DATA),
        "warrant_count": len(WEB_WARRANT_DATA),
        "job_status": result.get("job_status", "idle"),
        "job_error": result.get("job_error", ""),
        "job_started_at": result.get("job_started_at", ""),
        "job_progress": result.get("job_progress", 0),
        "job_message": result.get("job_message", ""),
        "job_elapsed_sec": result.get("job_elapsed_sec", 0),
        "auto_refresh": "weekday 16:05 Asia/Taipei",
    }
    return WEB_ANALYSIS_META


def _reload_web_analysis_from_db():
    from analysis_scheduler import get_running_progress, recover_stale_running_job
    from web_analysis_store import load_web_analysis_result

    cached = load_web_analysis_result()
    if not cached:
        return WEB_ANALYSIS_META
    cached = recover_stale_running_job(cached)
    if cached.get("job_status") == "running":
        cached = get_running_progress(cached)
    return _apply_web_analysis_result(cached)


def get_effective_cron_secret() -> str:
    explicit = (settings.CRON_SECRET or "").strip()
    if explicit:
        return explicit
    return (settings.SECRET_KEY or "").strip()


def _verify_cron_secret(x_cron_secret: Optional[str]) -> None:
    expected = get_effective_cron_secret()
    if not expected:
        raise HTTPException(status_code=503, detail="CRON_SECRET / SECRET_KEY 未設定，無法觸發自動分析")
    if (x_cron_secret or "").strip() != expected:
        raise HTTPException(status_code=403, detail="排程密鑰錯誤")


@app.get("/web/public/status")
def web_public_status():
    """公開狀態（不含個股清單），供監控與前端顯示更新時間。"""
    meta = _reload_web_analysis_from_db()
    return {
        "success": True,
        **meta,
        "has_data": bool(WEB_BULLISH_DATA or WEB_BEARISH_DATA or WEB_WARRANT_DATA),
        "cron_ready": bool(get_effective_cron_secret()),
    }


@app.get("/web/analysis-status")
def web_analysis_status(
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    _ = get_current_user(authorization, db)
    meta = _reload_web_analysis_from_db()
    return {
        "success": True,
        **meta,
        "has_data": bool(WEB_BULLISH_DATA or WEB_BEARISH_DATA or WEB_WARRANT_DATA),
    }


@app.post("/web/cron/daily-analysis")
def web_cron_daily_analysis(
    x_cron_secret: Optional[str] = Header(default=None, alias="X-Cron-Secret"),
):
    """Render Cron / 外部排程用。非同步執行，避免 HTTP timeout。"""
    _verify_cron_secret(x_cron_secret)
    started = run_analysis_in_background(trigger="cron-http")
    return {
        "success": True,
        "started": started,
        "message": "已開始背景分析" if started else "分析已在進行中",
        **WEB_ANALYSIS_META,
    }


@app.post("/web/run-analysis")
@app.get("/web/run-analysis")
def web_run_analysis(
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
    force: bool = False,
):
    user = get_current_user(authorization, db)

    if not user.is_active:
        raise HTTPException(status_code=403, detail="帳號已停用")
    require_active_license(user)

    _reload_web_analysis_from_db()

    if force and WEB_ANALYSIS_META.get("job_status") == "running":
        from analysis_scheduler import _set_job_meta
        from web_analysis_store import load_web_analysis_result, save_web_analysis_result

        cached = load_web_analysis_result() or {}
        cached = _set_job_meta(cached, "failed", "使用者強制重新啟動")
        save_web_analysis_result(cached)
        _apply_web_analysis_result(cached)

    if WEB_ANALYSIS_META.get("job_status") == "running":
        return {
            "success": True,
            "message": WEB_ANALYSIS_META.get("job_message") or "分析進行中，請稍候",
            **WEB_ANALYSIS_META,
        }

    started = run_analysis_in_background(trigger="manual")
    _reload_web_analysis_from_db()

    return {
        "success": True,
        "started": started,
        "message": "已開始背景分析，請稍候" if started else "分析已在進行中",
        **WEB_ANALYSIS_META,
    }


@app.get("/web/bullish")
def web_get_bullish(
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    user = get_current_user(authorization, db)
    require_active_license(user)
    return {"items": WEB_BULLISH_DATA}


@app.get("/web/bearish")
def web_get_bearish(
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    user = get_current_user(authorization, db)
    require_active_license(user)
    return {"items": WEB_BEARISH_DATA}


@app.get("/web/warrants")
def web_get_warrants(
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    user = get_current_user(authorization, db)
    require_active_license(user)
    return {"items": WEB_WARRANT_DATA}


@app.get("/run-analysis")
def run_analysis_legacy():

    global BULLISH_DATA

    BULLISH_DATA = [
        {
            "stock_id": "2330",
            "name": "台積電",
            "stars": "★★★★★",
            "strong_score": 128,
            "bias": "12%"
        },
        {
            "stock_id": "3017",
            "name": "奇鋐",
            "stars": "★★★★☆",
            "strong_score": 115,
            "bias": "18%"
        },
        {
            "stock_id": "3661",
            "name": "世芯",
            "stars": "★★★★★",
            "strong_score": 135,
            "bias": "9%"
        }
    ]

    return {
        "success": True,
        "count": len(BULLISH_DATA)
    }


@app.get("/bullish")
def get_bullish():

    return {
        "items": BULLISH_DATA
    }

@app.get("/stocks/bearish")
def get_bearish_stocks():
    return {
        "items": BEARISH_DATA
    }


@app.get("/run-bearish-analysis")
def run_bearish_analysis():

    global BEARISH_DATA

    BEARISH_DATA = [
        {
            "stock_id": "2409",
            "name": "友達",
            "stars": "★★★★★",
            "strong_score": 132,
            "bias": "-8%"
        },
        {
            "stock_id": "3481",
            "name": "群創",
            "stars": "★★★★☆",
            "strong_score": 118,
            "bias": "-12%"
        },
        {
            "stock_id": "2618",
            "name": "長航",
            "stars": "★★★★★",
            "strong_score": 126,
            "bias": "-10%"
        }
    ]

    return {
        "success": True,
        "count": len(BEARISH_DATA)
    }


@app.get("/bearish")
def get_bearish():

    return {
        "items": BEARISH_DATA
    }

@app.get("/run-warrant-analysis")
def run_warrant_analysis():

    global WARRANT_DATA

    WARRANT_DATA = [
        {
            "stock_id": "2330C",
            "name": "台積電認購",
            "broker": "元大",
            "price": "1.25",
            "strike": "1180",
            "premium": "12%"
        },
        {
            "stock_id": "3017C",
            "name": "奇鋐認購",
            "broker": "凱基",
            "price": "0.88",
            "strike": "820",
            "premium": "10%"
        },
        {
            "stock_id": "2409P",
            "name": "友達認售",
            "broker": "群益",
            "price": "0.72",
            "strike": "13.5",
            "premium": "9%"
        }
    ]

    return {
        "success": True,
        "count": len(WARRANT_DATA)
    }


@app.get("/warrants")
def get_warrants():

    return {
        "items": WARRANT_DATA
    }

HOLDING_DATA = []

@app.get("/run-holding-analysis")
def run_holding_analysis():

    global HOLDING_DATA

    HOLDING_DATA = [
        {
            "stock_id": "2330",
            "name": "台積電",
            "price": "1125",
            "profit": "+12%",
            "stars": "★★★★★"
        },
        {
            "stock_id": "3017",
            "name": "奇鋐",
            "price": "785",
            "profit": "+8%",
            "stars": "★★★★☆"
        },
        {
            "stock_id": "3661",
            "name": "世芯",
            "price": "2520",
            "profit": "+15%",
            "stars": "★★★★★"
        }
    ]

    return {
        "success": True,
        "count": len(HOLDING_DATA)
    }


@app.get("/holding")
def get_holding():

    return {
        "items": HOLDING_DATA
    }
