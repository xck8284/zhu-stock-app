from fastapi import FastAPI, Body
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

# 允許所有來源（手機 / 本地 / Web 都能打 API）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =========================
# 全域暫存資料
# =========================
STOCK_DATA = {
    "bullish": [],
    "bearish": [],
    "warrants": []
}


# =========================
# 測試 API（確認服務正常）
# =========================
@app.get("/")
def root():
    return {"status": "ok"}


# =========================
# 上傳選股結果（電腦版）
# =========================
@app.post("/admin/upload-stock-results")
def upload_stock_results(data: dict = Body(...)):
    global STOCK_DATA

    STOCK_DATA["bullish"] = data.get("bullish", [])
    STOCK_DATA["bearish"] = data.get("bearish", [])
    STOCK_DATA["warrants"] = data.get("warrants", [])

    return {
        "status": "success",
        "msg": "資料已更新",
        "count": {
            "bullish": len(STOCK_DATA["bullish"]),
            "bearish": len(STOCK_DATA["bearish"]),
            "warrants": len(STOCK_DATA["warrants"]),
        }
    }


# =========================
# 手機讀取：多 / 空
# =========================
@app.get("/mobile/stock-pools")
def get_stock_pools():
    return {
        "bullish": STOCK_DATA["bullish"],
        "bearish": STOCK_DATA["bearish"],
    }


# =========================
# 手機讀取：權證
# =========================
@app.get("/mobile/warrants")
def get_warrants():
    return {
        "warrants": STOCK_DATA["warrants"],
    }


# =========================
# 手機獨立分析（重點）
# =========================
@app.post("/mobile/run-analysis")
def mobile_run_analysis():
    global STOCK_DATA

    try:
        # 🔥 這行很重要（避免 import 問題）
        from analysis_core import run_analysis_core

        result = run_analysis_core()

        STOCK_DATA["bullish"] = result.get("bullish", [])
        STOCK_DATA["bearish"] = result.get("bearish", [])
        STOCK_DATA["warrants"] = result.get("warrants", [])

        return {
            "status": "success",
            "msg": "分析完成",
            "data": STOCK_DATA,
            "count": {
                "bullish": len(STOCK_DATA["bullish"]),
                "bearish": len(STOCK_DATA["bearish"]),
                "warrants": len(STOCK_DATA["warrants"]),
            },
            "error": result.get("error", "")
        }

    except Exception as e:
        return {
            "status": "error",
            "msg": str(e),
            "data": {
                "bullish": [],
                "bearish": [],
                "warrants": [],
            },
            "count": {
                "bullish": 0,
                "bearish": 0,
                "warrants": 0,
            }
        }
