from fastapi import Body

# 暫存選股資料
STOCK_DATA = {
    "bullish": [],
    "bearish": [],
    "warrants": []
}


# 🔥 上傳選股結果（電腦版用）
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


# 📱 手機讀取（看多 / 看空）
@app.get("/mobile/stock-pools")
def get_stock_pools():
    return {
        "bullish": STOCK_DATA["bullish"],
        "bearish": STOCK_DATA["bearish"],
    }


# 📱 手機讀取（權證）
@app.get("/mobile/warrants")
def get_warrants():
    return {
        "warrants": STOCK_DATA["warrants"],
    }


from analysis_core import run_analysis_core


# 📱 手機獨立分析
@app.post("/mobile/run-analysis")
def mobile_run_analysis():
    global STOCK_DATA

    try:
        result = run_analysis_core()

        STOCK_DATA["bullish"] = result.get("bullish", [])
        STOCK_DATA["bearish"] = result.get("bearish", [])
        STOCK_DATA["warrants"] = result.get("warrants", [])

        return {
            "status": "success",
            "msg": "分析完成",
            "raw_result": result,
            "data": STOCK_DATA,
            "count": {
                "bullish": len(STOCK_DATA["bullish"]),
                "bearish": len(STOCK_DATA["bearish"]),
                "warrants": len(STOCK_DATA["warrants"]),
            }
        }

    except Exception as e:
        return {
            "status": "error",
            "msg": str(e),
            "raw_result": {},
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
