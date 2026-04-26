import requests
import pandas as pd


def fetch_twse_data():
    url = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&type=ALLBUT0999"

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json",
    }

    res = requests.get(url, headers=headers, timeout=10)
    res.raise_for_status()
    data = res.json()

    rows = data.get("data9") or data.get("data") or []
    columns = data.get("fields9") or data.get("fields") or []

    if not rows or not columns:
        return pd.DataFrame(columns=["code", "name", "close", "change", "volume"])

    df = pd.DataFrame(rows, columns=columns)

    need_cols = ["證券代號", "證券名稱", "收盤價", "漲跌價差", "成交股數"]
    for c in need_cols:
        if c not in df.columns:
            return pd.DataFrame(columns=["code", "name", "close", "change", "volume"])

    df = df[need_cols].copy()
    df.columns = ["code", "name", "close", "change", "volume"]

    df["close"] = (
        df["close"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("--", "0", regex=False)
    )
    df["volume"] = (
        df["volume"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("--", "0", regex=False)
    )

    df["close"] = pd.to_numeric(df["close"], errors="coerce").fillna(0)
    df["volume"] = pd.to_numeric(df["volume"], errors="coerce").fillna(0)

    df = df[df["close"] > 0]

    return df


def calc_score(row):
    score = 0

    try:
        change = float(str(row["change"]).replace(",", "").replace("+", ""))
        if change > 0:
            score += 20
        elif change < 0:
            score -= 10
    except Exception:
        pass

    volume = float(row["volume"])
    close = float(row["close"])

    if volume > 10000000:
        score += 20
    elif volume > 5000000:
        score += 10
    elif volume > 1000000:
        score += 5

    if close < 50:
        score += 5
    elif close > 200:
        score -= 5

    return score


def stars_from_score(score):
    if score >= 35:
        return "★★★★★"
    if score >= 25:
        return "★★★★☆"
    if score >= 15:
        return "★★★☆☆"
    if score >= 5:
        return "★★☆☆☆"
    return "★☆☆☆☆"


def run_analysis_core():
    try:
        df = fetch_twse_data()
    except Exception:
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
                "code": str(row["code"]),
                "name": str(row["name"]),
                "industry": "",
                "settle_date": "",
                "stars": stars_from_score(score),
                "strong_score": round(score, 2),
                "bias": 0,
                "short_alarm": "否",
                "long_alarm": "否"
            }

            result.append(item)
        except Exception:
            continue

    result = sorted(result, key=lambda x: x["strong_score"], reverse=True)

    return {
        "bullish": result[:50],
        "bearish": [],
        "warrants": []
    }
