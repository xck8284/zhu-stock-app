import requests
import pandas as pd


def fetch_twse_data():
    url = "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL"

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json",
    }

    res = requests.get(url, headers=headers, timeout=20)
    res.raise_for_status()
    data = res.json()

    out = []
    for r in data:
        try:
            code = str(r.get("Code", "")).strip()
            name = str(r.get("Name", "")).strip()
            close = pd.to_numeric(str(r.get("ClosingPrice", "0")).replace(",", ""), errors="coerce")
            volume = pd.to_numeric(str(r.get("TradeVolume", "0")).replace(",", ""), errors="coerce")

            if not code or not name or pd.isna(close) or close <= 0:
                continue

            out.append({
                "code": code,
                "name": name,
                "close": close,
                "volume": 0 if pd.isna(volume) else volume,
                "change": 0,
            })
        except Exception:
            continue

    return pd.DataFrame(out)


def calc_score(row):
    score = 0
    volume = float(row["volume"])
    close = float(row["close"])

    if volume > 10000000:
        score += 20
    elif volume > 5000000:
        score += 10
    elif volume > 1000000:
        score += 5

    if 50 <= close <= 200:
        score += 10
    elif close < 50:
        score += 5
    elif close > 200:
        score -= 5

    return score


def stars_from_score(score):
    if score >= 30:
        return "★★★★★"
    if score >= 20:
        return "★★★★☆"
    if score >= 10:
        return "★★★☆☆"
    if score >= 5:
        return "★★☆☆☆"
    return "★☆☆☆☆"


def run_analysis_core():
    try:
        df = fetch_twse_data()
    except Exception as e:
        return {
            "bullish": [],
            "bearish": [],
            "warrants": [],
            "error": str(e),
        }

    result = []

    for _, row in df.iterrows():
        try:
            score = calc_score(row)

            result.append({
                "code": str(row["code"]),
                "name": str(row["name"]),
                "industry": "",
                "settle_date": "",
                "stars": stars_from_score(score),
                "strong_score": round(score, 2),
                "bias": 0,
                "short_alarm": "否",
                "long_alarm": "否",
            })
        except Exception:
            continue

    result = sorted(result, key=lambda x: x["strong_score"], reverse=True)

    return {
        "bullish": result[:50],
        "bearish": [],
        "warrants": [],
    }
