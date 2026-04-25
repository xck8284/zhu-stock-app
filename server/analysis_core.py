import pandas as pd
import requests


def safe_pct(a, b):
    try:
        if b is None or b == 0 or pd.isna(a) or pd.isna(b):
            return None
        return round((a - b) / b * 100, 2)
    except Exception:
        return None


def calc_training_score(grp, trend_info=None):
    latest = grp.iloc[-1]
    score = 0
    tags = []

    close_ = latest["close"]
    ma20 = latest["ma20"]

    if pd.notna(ma20) and close_ >= ma20:
        score += 20

    return score, ""


# 🔥 手機分析主入口（重點）
def run_analysis_core():
    url = "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL"

    try:
        res = requests.get(url, timeout=10)
        data = res.json()
    except:
        return {"bullish": [], "bearish": [], "warrants": []}

    results = []

    for row in data:
        try:
            code = row["Code"]
            name = row["Name"]
            close = float(row["ClosingPrice"])

            # 👉 簡單先給假 MA20（之後再優化）
            ma20 = close * 0.97

            df = pd.DataFrame([{
                "close": close,
                "ma20": ma20
            }])

            score, _ = calc_training_score(df)

            results.append({
                "code": code,
                "name": name,
                "industry": "",
                "settle_date": "",
                "stars": "★★★★★" if score > 10 else "★★★",
                "strong_score": score,
                "bias": round((close - ma20) / ma20 * 100, 2),
                "short_alarm": "否",
                "long_alarm": "否"
            })

        except:
            continue

    # 👉 排序
    results = sorted(results, key=lambda x: x["strong_score"], reverse=True)

    return {
        "bullish": results[:80],
        "bearish": [],
        "warrants": []
    }
