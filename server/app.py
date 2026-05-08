import certifi
import requests

from datetime import datetime, timedelta, timezone
import random
import re
from typing import Optional

from fastapi import FastAPI, Depends, HTTPException, Request, Header
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.orm import Session
from sqlalchemy import text

import sib_api_v3_sdk
from sib_api_v3_sdk.rest import ApiException

from config import settings
from database import SessionLocal, engine, Base
from models import User, AbuseLog, VerificationCode, PaymentReport
from schemas import (
    MessageResponse,
    SendRegisterCodeRequest, VerifyRegisterCodeRequest,
    LoginByAccountRequest, ForgotPasswordRequest, ResetPasswordRequest,
    AdminGrantFreeRequest, AdminSetPlanRequest, AdminRebindDeviceRequest,
    AdminDeactivateUserRequest, PaymentReportCreateRequest,
)
from security import hash_password, verify_password, create_access_token, decode_access_token

Base.metadata.create_all(bind=engine)

app = FastAPI(title="ZHU STOCK PLATFORM - COMPLETE UPGRADE", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


PLAN_DAYS = {
    "trial": settings.TRIAL_DAYS,
    "monthly": 30,
    "quarterly": 90,
    "yearly": 365,
}


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
            elif plan_type == "quarterly":
                label = f"季訂閱（剩餘 {days_left} 天）"
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
    item = PaymentReport(
        user_id=user.id,
        username=user.username,
        email=user.email,
        plan_type=data.plan_type,
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


@app.get("/admin/users")
def admin_users(
    authorization: Optional[str] = Header(default=None),
    db: Session = Depends(get_db),
):
    _ = get_current_creator(authorization, db)
    users = db.query(User).order_by(User.created_at.desc()).all()
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

    if data.plan_type not in ("monthly", "quarterly", "yearly"):
        raise HTTPException(status_code=400, detail="方案類型錯誤")

    start = now_utc()
    end = start + timedelta(days=PLAN_DAYS[data.plan_type])

    user.subscription_status = "active"
    user.plan_type = data.plan_type
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

    if report.plan_type not in ("monthly", "quarterly", "yearly"):
        raise HTTPException(status_code=400, detail="付款方案錯誤")

    start = now_utc()
    end = start + timedelta(days=PLAN_DAYS[report.plan_type])

    report.status = "approved"
    report.review_note = f"approved by {admin.email}"
    report.reviewed_at = now_utc()
    db.add(report)

    user.subscription_status = "active"
    user.plan_type = report.plan_type
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
@app.get("/admin/payment-reports")
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

@app.get("/stocks/bullish")
def get_bullish_stocks():
    return {
        "items":[
            {
                "stock_id":"2330",
                "name":"台積電",
                "stars":"★★★★★",
                "strong_score":128,
                "bias":"12%"
            },
            {
                "stock_id":"3017",
                "name":"奇鋐",
                "stars":"★★★★☆",
                "strong_score":115,
                "bias":"18%"
            }
        ]
    }

BULLISH_DATA = []

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
        "items": BULLISH_DATA
    }

from analysis import run_analysis

@app.get("/run-analysis")
def trigger_analysis():

    global BULLISH_DATA

    BULLISH_DATA = run_analysis()

    return {
        "success": True,
        "count": len(BULLISH_DATA)
    }
def run_analysis():

    global BULLISH_DATA

    BULLISH_DATA = [
        {
            "stock_id":"2330",
            "name":"台積電",
            "stars":"★★★★★",
            "strong_score":128,
            "bias":"12%"
        },
        {
            "stock_id":"3017",
            "name":"奇鋐",
            "stars":"★★★★☆",
            "strong_score":115,
            "bias":"18%"
        },
        {
            "stock_id":"3661",
            "name":"世芯",
            "stars":"★★★★★",
            "strong_score":135,
            "bias":"9%"
        }
    ]

    return {
        "success": True,
        "count": len(BULLISH_DATA)
    }
