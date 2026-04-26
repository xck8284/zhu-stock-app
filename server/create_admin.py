from datetime import datetime, timedelta

from sqlalchemy.orm import Session

from database import SessionLocal
from models import User
import security as sec


ADMIN_USERNAME = "admin"
ADMIN_EMAIL = "admin@zhustock.local"
ADMIN_PASSWORD = "Admin123456!"
ADMIN_PHONE = "0000000000"


def resolve_password_hash(password: str) -> str:
    """
    自動相容不同 security.py 的密碼加密函式名稱
    """
    if hasattr(sec, "get_password_hash"):
        return sec.get_password_hash(password)

    if hasattr(sec, "hash_password"):
        return sec.hash_password(password)

    if hasattr(sec, "create_password_hash"):
        return sec.create_password_hash(password)

    if hasattr(sec, "pwd_context"):
        return sec.pwd_context.hash(password)

    if hasattr(sec, "pwd_ctx"):
        return sec.pwd_ctx.hash(password)

    raise RuntimeError("找不到可用的密碼雜湊函式，請檢查 security.py")


def build_user_kwargs() -> dict:
    """
    只填目前 User model 真正存在的欄位，避免欄位不相容炸掉
    """
    now = datetime.utcnow()
    column_names = set(User.__table__.columns.keys())

    data = {
        "username": ADMIN_USERNAME,
        "full_name": "System Admin",
        "gender": "未設定",
        "phone": ADMIN_PHONE,
        "email": ADMIN_EMAIL,
        "password_hash": resolve_password_hash(ADMIN_PASSWORD),
        "is_active": True,
        "is_email_verified": True,
        "is_admin": True,
        "subscription_status": "active",
        "plan_type": "yearly",
        "trial_end_at": now + timedelta(days=3650),
        "subscription_start_at": now,
        "subscription_end_at": now + timedelta(days=3650),
        "payment_status": "paid",
        "device_id": None,
        "device_name": "ADMIN-CONSOLE",
        "created_at": now,
        "updated_at": now,
    }

    return {k: v for k, v in data.items() if k in column_names}


def main():
    db: Session = SessionLocal()
    try:
        # 先找舊帳號
        user = db.query(User).filter(
            (User.username == ADMIN_USERNAME) | (User.email == ADMIN_EMAIL)
        ).first()

        if user:
            print("管理員帳號已存在，不重複建立。")
            print(f"username={getattr(user, 'username', '')}")
            print(f"email={getattr(user, 'email', '')}")
            return

        kwargs = build_user_kwargs()
        admin = User(**kwargs)

        db.add(admin)
        db.commit()
        db.refresh(admin)

        print("管理員帳號建立成功")
        print(f"username: {ADMIN_USERNAME}")
        print(f"password: {ADMIN_PASSWORD}")

    except Exception as e:
        db.rollback()
        print("建立失敗：", repr(e))
    finally:
        db.close()


if __name__ == "__main__":
    main()