# -*- coding: utf-8 -*-
"""
ZHU STOCK APP - server/app.py restore version
功能：恢復登入/註冊/授權/API Docs + 保留手機版資料同步與獨立分析。
"""

import os
import re
import json
import random
from datetime import datetime, timedelta, timezone
from typing import Optional, Dict, Any, List

from fastapi import FastAPI, HTTPException, Request, Header, Depends
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, EmailStr
from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker

try:
    import jwt
except Exception:
    jwt = None

try:
    from passlib.context import CryptContext
    pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")
except Exception:
    pwd_context = None

try:
    from analysis_core import run_analysis_core
except Exception:
    run_analysis_core = None


# =========================
# 基本設定
# =========================
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./zhu_stock.db")
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)

JWT_SECRET = os.getenv("JWT_SECRET", os.getenv("SECRET_KEY", "zhu-stock-secret-change-me"))
JWT_ALGORITHM = "HS256"
JWT_EXPIRE_HOURS = int(os.getenv("JWT_EXPIRE_HOURS", "720"))
TRIAL_DAYS = int(os.getenv("TRIAL_DAYS", "30"))

CREATOR_ADMIN_USERNAME = os.getenv("CREATOR_ADMIN_USERNAME", "admin")
CREATOR_ADMIN_EMAIL = os.getenv("CREATOR_ADMIN_EMAIL", "admin@zhustock.local")
CREATOR_ADMIN_PASSWORD = os.getenv("CREATOR_ADMIN_PASSWORD", "admin123456")

engine_kwargs = {}
if DATABASE_URL.startswith("sqlite"):
    engine_kwargs["connect_args"] = {"check_same_thread": False}

engine = create_engine(DATABASE_URL, pool_pre_ping=True, **engine_kwargs)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)

app = FastAPI(title="ZHU STOCK PLATFORM API", version="2.1.0-restore")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

STOCK_DATA: Dict[str, Any] = {"bullish": [], "bearish": [], "warrants": []}


# =========================
# Schema
# =========================
class LoginByAccountRequest(BaseModel):
    account: str
    password: str
    device_id: Optional[str] = ""
    device_name: Optional[str] = ""


class SendRegisterCodeRequest(BaseModel):
    username: str
    password: str
    confirm_password: str
    phone: str
    email: EmailStr
    invite_code: Optional[str] = ""
    device_id: Optional[str] = ""
    device_name: Optional[str] = ""


class VerifyRegisterCodeRequest(BaseModel):
    username: str
    password: str
    confirm_password: Optional[str] = ""
    phone: str
    email: EmailStr
    code: str
    invite_code: Optional[str] = ""
    device_id: Optional[str] = ""
    device_name: Optional[str] = ""


class ForgotPasswordRequest(BaseModel):
    account: str


class ResetPasswordRequest(BaseModel):
    account: str
    code: str
    new_password: str
    confirm_password: Optional[str] = ""


class PaymentReportCreateRequest(BaseModel):
    account: Optional[str] = ""
    email: Optional[str] = ""
    plan_type: Optional[str] = ""
    bank: Optional[str] = ""
    last5: Optional[str] = ""
    amount: Optional[float] = 0
    phone: Optional[str] = ""
    note: Optional[str] = ""


class AdminSetPlanRequest(BaseModel):
    account: str
    plan_type: str = "yearly"
    days: Optional[int] = None
    note: Optional[str] = ""


class AdminGrantFreeRequest(BaseModel):
    account: str
    days: int = 30
    reason: Optional[str] = ""


class AdminDeactivateUserRequest(BaseModel):
    account: str
    is_active: bool = False


# =========================
# 工具函式
# =========================
def now_utc() -> datetime:
    return datetime.now(timezone.utc)


def to_iso(dt):
    if not dt:
        return ""
    if isinstance(dt, str):
        return dt
    try:
        return dt.isoformat()
    except Exception:
        return str(dt)


def parse_dt(value):
    if not value:
        return None
    if isinstance(value, datetime):
        return value
    s = str(value)
    try:
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except Exception:
        try:
            return datetime.strptime(s[:19], "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc)
        except Exception:
            return None


def normalize_email(email: str) -> str:
    return (email or "").strip().lower()


def normalize_username(username: str) -> str:
    return (username or "").strip()


def create_numeric_code(length: int = 6) -> str:
    return "".join(str(random.randint(0, 9)) for _ in range(length))


def hash_password(password: str) -> str:
    if pwd_context:
        return pwd_context.hash(password)
    return "plain$" + password


def verify_password(password: str, password_hash: str) -> bool:
    if not password_hash:
        return False
    if password_hash.startswith("plain$"):
        return password_hash == "plain$" + password
    if pwd_context:
        try:
            return pwd_context.verify(password, password_hash)
        except Exception:
            pass
    return password == password_hash


def create_access_token(payload: dict) -> str:
    data = dict(payload)
    data["exp"] = now_utc() + timedelta(hours=JWT_EXPIRE_HOURS)
    if jwt:
        return jwt.encode(data, JWT_SECRET, algorithm=JWT_ALGORITHM)
    return json.dumps({k: str(v) for k, v in data.items()}, ensure_ascii=False)


def decode_access_token(token: str) -> Optional[dict]:
    if not token:
        return None
    if jwt:
        try:
            return jwt.decode(token, JWT_SECRET, algorithms=[JWT_ALGORITHM])
        except Exception:
            return None
    try:
        return json.loads(token)
    except Exception:
        return None


def get_client_ip(request: Optional[Request]) -> str:
    if not request:
        return ""
    xff = request.headers.get("x-forwarded-for")
    if xff:
        return xff.split(",")[0].strip()
    if request.client:
        return request.client.host
    return ""


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def db_execute(sql: str, params: Optional[dict] = None):
    with engine.begin() as conn:
        return conn.execute(text(sql), params or {})


def table_columns(table_name: str) -> set:
    dialect = engine.dialect.name
    with engine.begin() as conn:
        if dialect == "sqlite":
            rows = conn.execute(text(f"PRAGMA table_info({table_name})")).fetchall()
            return {r[1] for r in rows}
        rows = conn.execute(text("""
            SELECT column_name FROM information_schema.columns
            WHERE table_name = :table_name
        """), {"table_name": table_name}).fetchall()
        return {r[0] for r in rows}


def add_column_if_missing(table: str, col: str, sql_type: str):
    try:
        cols = table_columns(table)
        if col not in cols:
            db_execute(f"ALTER TABLE {table} ADD COLUMN {col} {sql_type}")
    except Exception:
        pass


def init_db():
    db_execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username VARCHAR(100) UNIQUE,
        email VARCHAR(255) UNIQUE,
        phone VARCHAR(50),
        password_hash VARCHAR(255),
        is_email_verified BOOLEAN DEFAULT 0,
        subscription_status VARCHAR(50) DEFAULT 'trial',
        plan_type VARCHAR(50) DEFAULT 'trial',
        payment_status VARCHAR(50) DEFAULT '',
        trial_end_at TIMESTAMP,
        subscription_start_at TIMESTAMP,
        subscription_end_at TIMESTAMP,
        is_active BOOLEAN DEFAULT 1,
        is_creator BOOLEAN DEFAULT 0,
        device_id VARCHAR(255),
        device_name VARCHAR(255),
        invite_code VARCHAR(50),
        referred_by VARCHAR(50),
        free_reason VARCHAR(255),
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    # 舊 DB 若缺欄位，補上
    for col, typ in {
        "username": "VARCHAR(100)",
        "email": "VARCHAR(255)",
        "phone": "VARCHAR(50)",
        "password_hash": "VARCHAR(255)",
        "is_email_verified": "BOOLEAN DEFAULT 0",
        "subscription_status": "VARCHAR(50) DEFAULT 'trial'",
        "plan_type": "VARCHAR(50) DEFAULT 'trial'",
        "payment_status": "VARCHAR(50) DEFAULT ''",
        "trial_end_at": "TIMESTAMP",
        "subscription_start_at": "TIMESTAMP",
        "subscription_end_at": "TIMESTAMP",
        "is_active": "BOOLEAN DEFAULT 1",
        "is_creator": "BOOLEAN DEFAULT 0",
        "device_id": "VARCHAR(255)",
        "device_name": "VARCHAR(255)",
        "invite_code": "VARCHAR(50)",
        "referred_by": "VARCHAR(50)",
        "free_reason": "VARCHAR(255)",
        "created_at": "TIMESTAMP",
        "updated_at": "TIMESTAMP",
    }.items():
        add_column_if_missing("users", col, typ)

    db_execute("""
    CREATE TABLE IF NOT EXISTS verification_codes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email VARCHAR(255),
        purpose VARCHAR(50),
        code VARCHAR(20),
        expires_at TIMESTAMP,
        is_used BOOLEAN DEFAULT 0,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    db_execute("""
    CREATE TABLE IF NOT EXISTS payment_reports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        account VARCHAR(100),
        email VARCHAR(255),
        plan_type VARCHAR(50),
        bank VARCHAR(255),
        last5 VARCHAR(20),
        amount FLOAT DEFAULT 0,
        phone VARCHAR(50),
        note TEXT,
        status VARCHAR(50) DEFAULT 'pending',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    db_execute("""
    CREATE TABLE IF NOT EXISTS abuse_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        email VARCHAR(255),
        ip VARCHAR(100),
        event_type VARCHAR(100),
        detail TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)


def get_user_by_account(account: str) -> Optional[dict]:
    account = (account or "").strip()
    email = normalize_email(account)
    with engine.begin() as conn:
        row = conn.execute(text("""
            SELECT * FROM users
            WHERE username = :account OR email = :email
            LIMIT 1
        """), {"account": account, "email": email}).mappings().first()
        return dict(row) if row else None


def get_user_by_email(email: str) -> Optional[dict]:
    with engine.begin() as conn:
        row = conn.execute(text("SELECT * FROM users WHERE email=:email LIMIT 1"), {"email": normalize_email(email)}).mappings().first()
        return dict(row) if row else None


def get_user_by_username(username: str) -> Optional[dict]:
    with engine.begin() as conn:
        row = conn.execute(text("SELECT * FROM users WHERE username=:username LIMIT 1"), {"username": normalize_username(username)}).mappings().first()
        return dict(row) if row else None


def compute_license_status(user: dict) -> dict:
    if not user:
        return {"allowed": False, "subscription_status": "none", "plan_type": "none", "days_left": 0, "label": "未授權"}
    if not bool(user.get("is_active", True)):
        return {"allowed": False, "subscription_status": "inactive", "plan_type": user.get("plan_type") or "none", "days_left": 0, "label": "帳號已停用"}

    status = user.get("subscription_status") or "trial"
    plan_type = user.get("plan_type") or "trial"
    now = now_utc()
    end_at = parse_dt(user.get("subscription_end_at")) if status in ("active", "free_grant") else parse_dt(user.get("trial_end_at"))

    if end_at and end_at.tzinfo is None:
        end_at = end_at.replace(tzinfo=timezone.utc)

    allowed = bool(end_at and end_at > now)
    days_left = max((end_at - now).days, 0) if allowed else 0
    if not allowed:
        status = "expired"
        label = "授權已到期"
    elif status == "trial":
        label = f"免費試用中（剩餘 {days_left} 天）"
    elif status == "free_grant":
        label = f"活動贈送（剩餘 {days_left} 天）"
    else:
        label = f"授權有效（剩餘 {days_left} 天）"
    return {"allowed": allowed, "subscription_status": status, "plan_type": plan_type, "days_left": days_left, "end_at": to_iso(end_at), "label": label}


def user_payload(user: dict) -> dict:
    lic = compute_license_status(user)
    return {
        "username": user.get("username") or "",
        "email": user.get("email") or "",
        "phone": user.get("phone") or "",
        "subscription_status": lic["subscription_status"],
        "plan_type": lic["plan_type"],
        "trial_end_at": to_iso(user.get("trial_end_at")),
        "subscription_end_at": to_iso(user.get("subscription_end_at")),
        "role": "creator" if bool(user.get("is_creator")) else "user",
        "is_creator": bool(user.get("is_creator")),
        "is_admin": bool(user.get("is_creator")),
        "allowed": lic["allowed"],
        "days_left": lic["days_left"],
        "license_label": lic["label"],
    }


def ensure_default_creator():
    init_db()
    admin = get_user_by_account(CREATOR_ADMIN_USERNAME) or get_user_by_email(CREATOR_ADMIN_EMAIL)
    if admin:
        db_execute("""
            UPDATE users
            SET is_creator=1, is_active=1,
                subscription_status='active', plan_type='yearly',
                subscription_end_at=COALESCE(subscription_end_at, :end_at),
                updated_at=CURRENT_TIMESTAMP
            WHERE id=:id
        """, {"id": admin["id"], "end_at": now_utc() + timedelta(days=3650)})
        return

    db_execute("""
        INSERT INTO users
        (username, email, phone, password_hash, is_email_verified, subscription_status, plan_type,
         subscription_start_at, subscription_end_at, is_active, is_creator, created_at, updated_at)
        VALUES
        (:username, :email, '', :password_hash, 1, 'active', 'yearly', :start_at, :end_at, 1, 1, :now, :now)
    """, {
        "username": CREATOR_ADMIN_USERNAME,
        "email": normalize_email(CREATOR_ADMIN_EMAIL),
        "password_hash": hash_password(CREATOR_ADMIN_PASSWORD),
        "start_at": now_utc(),
        "end_at": now_utc() + timedelta(days=3650),
        "now": now_utc(),
    })


def save_verification_code(email: str, purpose: str, code: str):
    email = normalize_email(email)
    db_execute("UPDATE verification_codes SET is_used=1 WHERE email=:email AND purpose=:purpose AND is_used=0", {"email": email, "purpose": purpose})
    db_execute("""
        INSERT INTO verification_codes (email, purpose, code, expires_at, is_used, created_at)
        VALUES (:email, :purpose, :code, :expires_at, 0, :created_at)
    """, {"email": email, "purpose": purpose, "code": code, "expires_at": now_utc() + timedelta(minutes=10), "created_at": now_utc()})


def check_verification_code(email: str, purpose: str, code: str):
    email = normalize_email(email)
    with engine.begin() as conn:
        row = conn.execute(text("""
            SELECT * FROM verification_codes
            WHERE email=:email AND purpose=:purpose AND code=:code AND is_used=0
            ORDER BY id DESC LIMIT 1
        """), {"email": email, "purpose": purpose, "code": str(code).strip()}).mappings().first()
        if not row:
            raise HTTPException(status_code=400, detail="驗證碼錯誤或不存在")
        exp = parse_dt(row["expires_at"])
        if exp and exp.tzinfo is None:
            exp = exp.replace(tzinfo=timezone.utc)
        if exp and exp < now_utc():
            raise HTTPException(status_code=400, detail="驗證碼已過期")
        conn.execute(text("UPDATE verification_codes SET is_used=1 WHERE id=:id"), {"id": row["id"]})


def current_user_from_auth(authorization: Optional[str]) -> dict:
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="缺少 token")
    token = authorization.replace("Bearer ", "", 1).strip()
    payload = decode_access_token(token)
    if not payload:
        raise HTTPException(status_code=401, detail="token 無效")
    sub = payload.get("sub") or payload.get("email")
    user = get_user_by_email(sub)
    if not user:
        raise HTTPException(status_code=401, detail="找不到使用者")
    return user


def require_creator(authorization: Optional[str]) -> dict:
    user = current_user_from_auth(authorization)
    if not bool(user.get("is_creator")):
        raise HTTPException(status_code=403, detail="非管理員權限")
    return user


@app.on_event("startup")
def startup():
    ensure_default_creator()


# =========================
# 基本與 Auth API
# =========================
@app.get("/")
def root():
    return {"success": True, "message": "ZHU STOCK PLATFORM API is running", "version": "2.1.0-restore"}


@app.post("/auth/send-register-code")
def auth_send_register_code(data: SendRegisterCodeRequest):
    username = normalize_username(data.username)
    email = normalize_email(str(data.email))
    if data.password != data.confirm_password:
        raise HTTPException(status_code=400, detail="密碼與確認密碼不一致")
    if len(data.password) < 6:
        raise HTTPException(status_code=400, detail="密碼長度至少 6 碼")
    if not re.fullmatch(r"[A-Za-z0-9_.\-@]+", username):
        raise HTTPException(status_code=400, detail="帳號只能使用英數與 . _ - @")
    if get_user_by_username(username):
        raise HTTPException(status_code=400, detail="此帳號已存在")
    if get_user_by_email(email):
        raise HTTPException(status_code=400, detail="此 Email 已註冊")
    code = create_numeric_code(6)
    save_verification_code(email, "register", code)
    return {"success": True, "message": "驗證碼已產生", "dev_code": code}


@app.post("/auth/verify-register-code")
def auth_verify_register_code(data: VerifyRegisterCodeRequest, request: Request):
    username = normalize_username(data.username)
    email = normalize_email(str(data.email))
    if data.confirm_password and data.password != data.confirm_password:
        raise HTTPException(status_code=400, detail="密碼與確認密碼不一致")
    check_verification_code(email, "register", data.code)
    if get_user_by_username(username):
        raise HTTPException(status_code=400, detail="此帳號已存在")
    if get_user_by_email(email):
        raise HTTPException(status_code=400, detail="此 Email 已註冊")
    trial_end = now_utc() + timedelta(days=TRIAL_DAYS)
    db_execute("""
        INSERT INTO users
        (username, email, phone, password_hash, is_email_verified, subscription_status, plan_type,
         trial_end_at, is_active, is_creator, device_id, device_name, invite_code, referred_by, created_at, updated_at)
        VALUES
        (:username, :email, :phone, :password_hash, 1, 'trial', 'trial', :trial_end,
         1, 0, :device_id, :device_name, :invite_code, :referred_by, :now, :now)
    """, {
        "username": username,
        "email": email,
        "phone": data.phone,
        "password_hash": hash_password(data.password),
        "trial_end": trial_end,
        "device_id": data.device_id or "",
        "device_name": data.device_name or "",
        "invite_code": username.upper()[:8],
        "referred_by": (data.invite_code or "").upper(),
        "now": now_utc(),
    })
    user = get_user_by_username(username)
    token = create_access_token({"sub": user["email"], "username": user["username"], "is_creator": bool(user.get("is_creator"))})
    payload = user_payload(user)
    return {"success": True, "message": "註冊成功，已開通免費試用 30 天", "access_token": token, "token_type": "bearer", **payload, "user": payload}


@app.post("/auth/login")
def auth_login(data: LoginByAccountRequest, request: Request):
    account = normalize_username(data.account)
    user = get_user_by_account(account)
    if not user:
        raise HTTPException(status_code=404, detail="帳號不存在")
    if not verify_password(data.password, user.get("password_hash") or ""):
        raise HTTPException(status_code=400, detail="密碼錯誤")
    if not bool(user.get("is_active", True)):
        raise HTTPException(status_code=403, detail="帳號已停用")

    if data.device_id and not user.get("device_id"):
        db_execute("UPDATE users SET device_id=:device_id, device_name=:device_name, updated_at=CURRENT_TIMESTAMP WHERE id=:id", {
            "device_id": data.device_id,
            "device_name": data.device_name or "",
            "id": user["id"],
        })
        user = get_user_by_account(account)

    token = create_access_token({"sub": user["email"], "username": user["username"], "is_creator": bool(user.get("is_creator"))})
    payload = user_payload(user)
    return {
        "success": True,
        "message": "登入成功",
        "access_token": token,
        "token_type": "bearer",
        **payload,
        "user": payload,
    }


@app.get("/auth/me")
def auth_me(authorization: Optional[str] = Header(None)):
    user = current_user_from_auth(authorization)
    payload = user_payload(user)
    return {"success": True, **payload, "user": payload}


@app.post("/auth/forgot-password")
def auth_forgot_password(data: ForgotPasswordRequest):
    user = get_user_by_account(data.account)
    if not user:
        raise HTTPException(status_code=404, detail="找不到帳號")
    code = create_numeric_code(6)
    save_verification_code(user["email"], "reset_password", code)
    return {"success": True, "message": "重設驗證碼已產生", "dev_code": code}


@app.post("/auth/reset-password")
def auth_reset_password(data: ResetPasswordRequest):
    user = get_user_by_account(data.account)
    if not user:
        raise HTTPException(status_code=404, detail="找不到帳號")
    if data.confirm_password and data.new_password != data.confirm_password:
        raise HTTPException(status_code=400, detail="密碼與確認密碼不一致")
    if len(data.new_password) < 6:
        raise HTTPException(status_code=400, detail="密碼長度至少 6 碼")
    check_verification_code(user["email"], "reset_password", data.code)
    db_execute("UPDATE users SET password_hash=:password_hash, updated_at=CURRENT_TIMESTAMP WHERE id=:id", {
        "password_hash": hash_password(data.new_password),
        "id": user["id"],
    })
    return {"success": True, "message": "密碼已更新"}


@app.post("/license/status")
def license_status(data: LoginByAccountRequest):
    user = get_user_by_account(data.account)
    if not user:
        raise HTTPException(status_code=404, detail="帳號不存在")
    payload = user_payload(user)
    return {"success": True, "message": payload["license_label"], **payload}


# =========================
# 付款與管理 API
# =========================
@app.post("/payment/report")
def payment_report(data: PaymentReportCreateRequest):
    db_execute("""
        INSERT INTO payment_reports
        (account, email, plan_type, bank, last5, amount, phone, note, status, created_at, updated_at)
        VALUES (:account, :email, :plan_type, :bank, :last5, :amount, :phone, :note, 'pending', :now, :now)
    """, {**data.dict(), "now": now_utc()})
    return {"success": True, "message": "付款回報已送出，待管理員審核"}


@app.get("/admin/overview")
def admin_overview(authorization: Optional[str] = Header(None)):
    require_creator(authorization)
    with engine.begin() as conn:
        users = [dict(r) for r in conn.execute(text("SELECT * FROM users ORDER BY id DESC")).mappings().all()]
        payments = [dict(r) for r in conn.execute(text("SELECT * FROM payment_reports ORDER BY id DESC")).mappings().all()]
    user_rows = [user_payload(u) | {"id": u.get("id"), "created_at": to_iso(u.get("created_at")), "is_active": bool(u.get("is_active", True))} for u in users]
    return {"success": True, "users": user_rows, "payments": payments, "payment_reports": payments}


@app.post("/admin/set-plan")
def admin_set_plan(data: AdminSetPlanRequest, authorization: Optional[str] = Header(None)):
    require_creator(authorization)
    user = get_user_by_account(data.account)
    if not user:
        raise HTTPException(status_code=404, detail="帳號不存在")
    plan = data.plan_type
    days_map = {"monthly": 30, "halfyear": 183, "yearly": 365, "quarterly": 90}
    days = data.days or days_map.get(plan, 365)
    db_execute("""
        UPDATE users
        SET subscription_status='active', plan_type=:plan_type,
            subscription_start_at=:start_at, subscription_end_at=:end_at,
            is_active=1, updated_at=CURRENT_TIMESTAMP
        WHERE id=:id
    """, {"plan_type": plan, "start_at": now_utc(), "end_at": now_utc() + timedelta(days=days), "id": user["id"]})
    return {"success": True, "message": "方案已更新"}


@app.post("/admin/grant-free")
def admin_grant_free(data: AdminGrantFreeRequest, authorization: Optional[str] = Header(None)):
    require_creator(authorization)
    user = get_user_by_account(data.account)
    if not user:
        raise HTTPException(status_code=404, detail="帳號不存在")
    db_execute("""
        UPDATE users
        SET subscription_status='free_grant', plan_type='free_grant', free_reason=:reason,
            subscription_start_at=:start_at, subscription_end_at=:end_at,
            is_active=1, updated_at=CURRENT_TIMESTAMP
        WHERE id=:id
    """, {"reason": data.reason or "", "start_at": now_utc(), "end_at": now_utc() + timedelta(days=data.days), "id": user["id"]})
    return {"success": True, "message": "已贈送使用天數"}


@app.post("/admin/deactivate-user")
def admin_deactivate_user(data: AdminDeactivateUserRequest, authorization: Optional[str] = Header(None)):
    require_creator(authorization)
    user = get_user_by_account(data.account)
    if not user:
        raise HTTPException(status_code=404, detail="帳號不存在")
    db_execute("UPDATE users SET is_active=:is_active, updated_at=CURRENT_TIMESTAMP WHERE id=:id", {"is_active": 1 if data.is_active else 0, "id": user["id"]})
    return {"success": True, "message": "會員狀態已更新"}


# =========================
# 手機版資料 API
# =========================
@app.post("/admin/upload-stock-results")
def upload_stock_results(payload: Dict[str, Any]):
    global STOCK_DATA
    STOCK_DATA = {
        "bullish": payload.get("bullish", []) or [],
        "bearish": payload.get("bearish", []) or [],
        "warrants": payload.get("warrants", []) or [],
        "updated_at": payload.get("updated_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        "settle_date": payload.get("settle_date", ""),
    }
    return {"status": "success", "msg": "資料已更新", "count": {"bullish": len(STOCK_DATA["bullish"]), "bearish": len(STOCK_DATA["bearish"]), "warrants": len(STOCK_DATA["warrants"])}}


@app.get("/mobile/stock-pools")
def get_stock_pools():
    return {"bullish": STOCK_DATA.get("bullish", []), "bearish": STOCK_DATA.get("bearish", [])}


@app.get("/mobile/warrants")
def get_warrants():
    return {"warrants": STOCK_DATA.get("warrants", [])}


@app.post("/mobile/run-analysis")
def mobile_run_analysis():
    global STOCK_DATA
    try:
        if run_analysis_core is None:
            raise RuntimeError("analysis_core.run_analysis_core 尚未載入")
        result = run_analysis_core()
        bullish = result.get("bullish", []) or []
        bearish = result.get("bearish", []) or []
        warrants = result.get("warrants", []) or []
        STOCK_DATA = {"bullish": bullish, "bearish": bearish, "warrants": warrants, "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
        return {"status": "success", "msg": "分析完成", "data": STOCK_DATA, "count": {"bullish": len(bullish), "bearish": len(bearish), "warrants": len(warrants)}, "error": result.get("error", "")}
    except Exception as e:
        return {"status": "error", "msg": "分析失敗", "data": {"bullish": [], "bearish": [], "warrants": []}, "count": {"bullish": 0, "bearish": 0, "warrants": 0}, "error": str(e)}
