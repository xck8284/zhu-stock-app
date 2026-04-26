from database import SessionLocal
from models import User

db = SessionLocal()

email = "xck8284@gmail.com"
user = db.query(User).filter(User.email == email).first()

if user:
    user.is_creator = True
    db.commit()
    print(f"{email} 已設為創作者")
else:
    print("查無此使用者")

db.close()