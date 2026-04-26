
ZHU STOCK server complete upgrade v2

覆蓋建議檔案：
- app.py
- models.py
- schemas.py
- config.py

保留不動：
- database.py
- security.py
- .env
- zhu_stock.db
- geoip.py
- mailer.py

啟動：
cd /d C:\Users\user\Desktop\zhustock_app\server
uvicorn app:app --host 127.0.0.1 --port 8000 --reload

新功能：
1. 註冊驗證碼
- POST /auth/send-register-code
- POST /auth/verify-register-code

2. 忘記密碼
- POST /auth/forgot-password
- POST /auth/reset-password

3. 授權查詢
- GET /auth/me
- GET /license/status

4. 後台會員管理
- GET /admin/users
- POST /admin/grant-free
- POST /admin/set-plan
- POST /admin/rebind-device
- POST /admin/deactivate-user

5. 付款回報
- POST /payments/report
- GET /admin/payment-reports
- POST /admin/approve-payment-report/{report_id}
- POST /admin/reject-payment-report/{report_id}

重要：
- 免費試用固定 30 天
- free_grant 表示活動贈送或人工免費資格
- monthly / quarterly / yearly 代表月 / 季 / 年方案
- EMAIL_DEV_MODE=True 時，驗證碼會在 API response 中直接回 dev_code，方便測試

管理員帳號：
- account: admin
- email: admin@zhustock.local
- password: Admin123456!

管理員登入：
POST /auth/login
{
  "account": "admin",
  "password": "Admin123456!"
}

之後把 access_token 複製到 Swagger 右上角 Authorize
格式：
Bearer 你的token

建議明天先測：
1. 用 admin 登入
2. /admin/users 看用戶列表
3. /admin/grant-free 測活動贈送 7 天或 30 天
4. /admin/set-plan 測 monthly / quarterly / yearly
