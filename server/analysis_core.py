import requests
import pandas as pd


def fetch_twse_data():
    url = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&type=ALLBUT0999"
    headers = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/json",
}

res = requests.get(url, headers=headers, timeout=10)
    data = res.json()

    rows = data["data9"]
    columns = data["fields9"]

    df = pd.DataFrame(rows, columns=columns)

    df = df[["證券代號", "證券名稱", "收盤價", "漲跌價差", "成交股數"]]
    df.columns = ["code", "name", "close", "change", "volume"]

    df["close"] = df["close"].str.replace(",", "").astype(float)
    df["volume"] = df["volume"].str.replace(",", "").astype(float)

    return df


def calc_score(row):
    score = 0

    # 價格動能
    if row["change"] != "--":
        try:
            change = float(row["change"])
            if change > 0:
                score += 20
            elif change < 0:
                score -= 10
        except:
            pass

    # 量能
    if row["volume"] > 10000000:
        score += 20
    elif row["volume"] > 5000000:
        score += 10
    elif row["volume"] > 1000000:
        score += 5

    # 價格大小
    if row["close"] < 50:
        score += 5
    elif row["close"] > 200:
        score -= 5

    return score


def run_analysis_core():
    try:
        df = fetch_twse_data()
    except Exception as e:
        return {
            "bullish": [],
            "bearish": [],
            "warrants": []
        }

    result = []

    for _, row in df.iterrows():
        try:
            score = calc_score(row)

            item = {
                "code": row["code"],
                "name": row["name"],
                "industry": "",
                "settle_date": "",
                "stars": "★" * max(1, min(5, score // 20)),
                "strong_score": round(score, 2),
                "bias": 0,
                "short_alarm": "否",
                "long_alarm": "否"
            }

            result.append(item)

        except:
            continue

    result = sorted(result, key=lambda x: x["strong_score"], reverse=True)

    return {
        "bullish": result[:50],
        "bearish": [],
        "warrants": []
    }
