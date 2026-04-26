from database import SessionLocal
from models import User

ADMIN_ACCOUNT = "admin"
ADMIN_EMAIL = "admin@zhustock.local"

def main():
    db = SessionLocal()
    try:
        user = db.query(User).filter(
            (User.username == ADMIN_ACCOUNT) | (User.email == ADMIN_EMAIL)
        ).first()

        if not user:
            print("找不到 admin 帳號")
            return

        changed = []

        if hasattr(user, "is_creator"):
            setattr(user, "is_creator", True)
            changed.append("is_creator=True")

        if hasattr(user, "is_admin"):
            setattr(user, "is_admin", True)
            changed.append("is_admin=True")

        if hasattr(user, "is_active"):
            setattr(user, "is_active", True)
            changed.append("is_active=True")

        db.add(user)
        db.commit()
        db.refresh(user)

        print("admin 權限修正成功")
        print("已更新：", ", ".join(changed) if changed else "沒有可更新欄位")
        print("username =", getattr(user, "username", None))
        print("email    =", getattr(user, "email", None))
        print("is_creator =", getattr(user, "is_creator", None))
        print("is_admin   =", getattr(user, "is_admin", None))

    except Exception as e:
        db.rollback()
        print("修正失敗：", repr(e))
    finally:
        db.close()

if __name__ == "__main__":
    main()