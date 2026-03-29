# -*- coding: utf-8 -*-
"""
ZHU STOCK 啟動器
先做授權驗證，通過後再啟動真正主程式
"""

import json
import os
import uuid
import platform
import subprocess
import sys
import tkinter as tk
from tkinter import ttk, messagebox

import requests

# =========================
# 使用者設定
# =========================
API_BASE = "http://127.0.0.1:8000"
LOCAL_SAVE_FILE = "license_local.json"

# 這裡改成你真正的股票主程式路徑
TARGET_APP = r"C:\Users\user\Desktop\zhustock\zhustock_app.py"

REQUEST_TIMEOUT = 10


# =========================
# 工具函式
# =========================
def get_device_id() -> str:
    return str(uuid.getnode())


def get_device_name() -> str:
    return platform.node() or "Unknown-PC"


def load_local_account() -> dict:
    if os.path.exists(LOCAL_SAVE_FILE):
        try:
            with open(LOCAL_SAVE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_local_account(data: dict):
    with open(LOCAL_SAVE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def post_json(endpoint: str, payload: dict) -> dict:
    url = f"{API_BASE}{endpoint}"
    r = requests.post(url, json=payload, timeout=REQUEST_TIMEOUT)
    r.raise_for_status()
    return r.json()


def launch_target_app():
    if not os.path.exists(TARGET_APP):
        messagebox.showerror("錯誤", f"找不到主程式：\n{TARGET_APP}")
        return False

    try:
        subprocess.Popen([sys.executable, TARGET_APP], cwd=os.path.dirname(TARGET_APP))
        return True
    except Exception as e:
        messagebox.showerror("錯誤", f"啟動主程式失敗：\n{e}")
        return False


# =========================
# GUI
# =========================
class LicenseLauncherApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("ZHU STOCK｜授權啟動器")
        self.root.geometry("560x540")
        self.root.resizable(False, False)

        self.device_id = get_device_id()
        self.device_name = get_device_name()
        self.local_data = load_local_account()

        self.build_ui()
        self.fill_saved_data()

    def build_ui(self):
        main = ttk.Frame(self.root, padding=16)
        main.pack(fill="both", expand=True)

        ttk.Label(
            main,
            text="ZHU STOCK 授權啟動器",
            font=("Microsoft JhengHei UI", 16, "bold")
        ).pack(pady=(0, 12))

        info_frame = ttk.LabelFrame(main, text="本機裝置資訊", padding=12)
        info_frame.pack(fill="x", pady=(0, 12))

        ttk.Label(info_frame, text=f"Device ID：{self.device_id}").pack(anchor="w", pady=2)
        ttk.Label(info_frame, text=f"Device Name：{self.device_name}").pack(anchor="w", pady=2)
        ttk.Label(info_frame, text=f"API：{API_BASE}").pack(anchor="w", pady=2)

        form_frame = ttk.LabelFrame(main, text="帳號資訊", padding=12)
        form_frame.pack(fill="x", pady=(0, 12))

        self.full_name_var = tk.StringVar()
        self.gender_var = tk.StringVar(value="男")
        self.phone_var = tk.StringVar()
        self.email_var = tk.StringVar()
        self.password_var = tk.StringVar()

        self._add_labeled_entry(form_frame, "姓名", self.full_name_var)
        self._add_labeled_combo(form_frame, "性別", self.gender_var, ["男", "女", "其他"])
        self._add_labeled_entry(form_frame, "電話", self.phone_var)
        self._add_labeled_entry(form_frame, "Email", self.email_var)
        self._add_labeled_entry(form_frame, "密碼", self.password_var, show="*")

        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill="x", pady=(0, 12))

        ttk.Button(btn_frame, text="註冊", command=self.on_register).pack(side="left", padx=4, fill="x", expand=True)
        ttk.Button(btn_frame, text="登入並啟動主程式", command=self.on_login_and_launch).pack(side="left", padx=4, fill="x", expand=True)
        ttk.Button(btn_frame, text="直接驗證本機帳號並啟動", command=self.on_check_saved_and_launch).pack(side="left", padx=4, fill="x", expand=True)

        result_frame = ttk.LabelFrame(main, text="執行結果", padding=12)
        result_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(result_frame, height=14, wrap="word", font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)

        self.write_log("啟動器已開啟")
        self.write_log(f"Device ID：{self.device_id}")
        self.write_log(f"Device Name：{self.device_name}")
        self.write_log(f"主程式路徑：{TARGET_APP}")

    def _add_labeled_entry(self, parent, label, text_var, show=None):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=4)

        ttk.Label(row, text=label, width=10).pack(side="left")
        ent = ttk.Entry(row, textvariable=text_var, show=show)
        ent.pack(side="left", fill="x", expand=True)

    def _add_labeled_combo(self, parent, label, text_var, values):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=4)

        ttk.Label(row, text=label, width=10).pack(side="left")
        combo = ttk.Combobox(row, textvariable=text_var, values=values, state="readonly")
        combo.pack(side="left", fill="x", expand=True)

    def write_log(self, msg: str):
        self.log_text.insert("end", f"{msg}\n")
        self.log_text.see("end")

    def fill_saved_data(self):
        if not self.local_data:
            return
        self.full_name_var.set(self.local_data.get("full_name", ""))
        self.gender_var.set(self.local_data.get("gender", "男"))
        self.phone_var.set(self.local_data.get("phone", ""))
        self.email_var.set(self.local_data.get("email", ""))
        self.password_var.set(self.local_data.get("password", ""))

    def get_form_data(self) -> dict:
        return {
            "full_name": self.full_name_var.get().strip(),
            "gender": self.gender_var.get().strip(),
            "phone": self.phone_var.get().strip(),
            "email": self.email_var.get().strip(),
            "password": self.password_var.get().strip(),
            "device_id": self.device_id,
            "device_name": self.device_name,
        }

    def validate_basic(self, require_full=True) -> bool:
        data = self.get_form_data()

        if require_full and not data["full_name"]:
            messagebox.showwarning("提醒", "請輸入姓名")
            return False

        if not data["email"]:
            messagebox.showwarning("提醒", "請輸入 Email")
            return False

        if not data["password"]:
            messagebox.showwarning("提醒", "請輸入密碼")
            return False

        return True

    def do_license_check(self, email: str) -> tuple[bool, dict]:
        check_payload = {
            "email": email,
            "device_id": self.device_id,
        }

        check_result = post_json("/license/check", check_payload)
        self.write_log(f"/license/check 回傳：{check_result}")

        if check_result.get("success") and check_result.get("allowed"):
            return True, check_result
        return False, check_result

    def on_register(self):
        if not self.validate_basic(require_full=True):
            return

        data = self.get_form_data()

        try:
            self.write_log("開始註冊...")
            result = post_json("/register", data)
            self.write_log(f"/register 回傳：{result}")

            if result.get("success"):
                save_local_account(data)
                self.write_log("本機帳號資料已保存")
                messagebox.showinfo("成功", result.get("message", "註冊成功"))
            else:
                messagebox.showwarning("提醒", result.get("message", "註冊失敗"))

        except requests.exceptions.RequestException as e:
            self.write_log(f"註冊連線失敗：{e}")
            messagebox.showerror("錯誤", f"無法連線到後台：\n{e}")
        except Exception as e:
            self.write_log(f"註冊發生錯誤：{e}")
            messagebox.showerror("錯誤", str(e))

    def on_login_and_launch(self):
        if not self.validate_basic(require_full=False):
            return

        payload = {
            "email": self.email_var.get().strip(),
            "password": self.password_var.get().strip(),
            "device_id": self.device_id,
            "device_name": self.device_name,
        }

        try:
            self.write_log("開始登入...")
            login_result = post_json("/login", payload)
            self.write_log(f"/login 回傳：{login_result}")

            if not login_result.get("success"):
                messagebox.showwarning("登入失敗", login_result.get("message", "登入失敗"))
                return

            save_local_account(self.get_form_data())
            self.write_log("登入成功，本機帳號資料已保存")

            ok, check_result = self.do_license_check(payload["email"])
            if not ok:
                messagebox.showwarning("授權失敗", check_result.get("message", "不可使用"))
                return

            self.write_log("授權驗證成功，準備啟動主程式...")
            if launch_target_app():
                messagebox.showinfo(
                    "成功",
                    f"授權驗證通過\n\n"
                    f"狀態：{check_result.get('subscription_status')}\n"
                    f"付款：{check_result.get('payment_status')}\n"
                    f"即將啟動主程式"
                )
                self.root.destroy()

        except requests.exceptions.RequestException as e:
            self.write_log(f"登入/驗證連線失敗：{e}")
            messagebox.showerror("錯誤", f"無法連線到後台：\n{e}")
        except Exception as e:
            self.write_log(f"登入/驗證發生錯誤：{e}")
            messagebox.showerror("錯誤", str(e))

    def on_check_saved_and_launch(self):
        saved = load_local_account()
        if not saved:
            messagebox.showwarning("提醒", "本機沒有保存帳號資料")
            return

        email = saved.get("email", "").strip()
        if not email:
            messagebox.showwarning("提醒", "本機保存資料不完整")
            return

        try:
            self.write_log("使用本機保存帳號進行授權驗證...")
            ok, check_result = self.do_license_check(email)

            if not ok:
                messagebox.showwarning("授權失敗", check_result.get("message", "不可使用"))
                return

            self.write_log("本機授權有效，準備啟動主程式...")
            if launch_target_app():
                messagebox.showinfo(
                    "成功",
                    f"本機授權有效\n\n"
                    f"狀態：{check_result.get('subscription_status')}\n"
                    f"付款：{check_result.get('payment_status')}\n"
                    f"即將啟動主程式"
                )
                self.root.destroy()

        except requests.exceptions.RequestException as e:
            self.write_log(f"驗證連線失敗：{e}")
            messagebox.showerror("錯誤", f"無法連線到後台：\n{e}")
        except Exception as e:
            self.write_log(f"驗證發生錯誤：{e}")
            messagebox.showerror("錯誤", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = LicenseLauncherApp(root)
    root.mainloop()