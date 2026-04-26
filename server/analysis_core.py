import requests
import pandas as pd


def fetch_twse_data():
    url = "https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&type=ALLBUT0999"

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json,text/plain,*/*",
        "Referer": "https://www.twse.com.tw/",
    }

    res = requests.get(url, headers=headers, timeout=15)
    res.raise_for_status()
    data = res.json()

    rows = data.get("data9") or data.get("data") or []
    fields = data.get("fields9") or data.get("fields") or []

    if not rows:
        return pd.DataFrame(columns=["code", "name", "close", "change", "volume"])

    # 優先用欄位名稱抓；如果欄位名稱不同，就用固定位置抓
    out = []

    code_i = fields.index("證券代號") if "證券代號" in fields else 0
    name_i = fields.index("證券名稱") if "證券名稱" in fields else 1
    close_i = fields.index("收盤價") if "收盤價" in fields else 8
    change_i = fields.index("漲跌價差") if "漲跌價差" in fields else 10
    volume_i = fields.index("成交股數") if "成交股數" in fields else 2

    for r in rows:
        try:
            code = str(r[code_i]).strip()
            name = str(r[name_i]).strip()
            close = str(r[close_i]).replace(",", "").replace("--", "0").strip()
            change = str(r[change_i]).replace(",", "").replace("+", "").replace("--", "0").strip()
            volume = str(r[volume_i]).replace(",", "").replace("--", "0").strip()

            out.append({
                "code": code,
                "name": name,
                "close": pd.to_numeric(close, errors="coerce"),
                "change": pd.to_numeric(change, errors="coerce"),
                "volume": pd.to_numeric(volume, errors="coerce"),
            })
        except Exception:
            continue

    df = pd.DataFrame(out)
    if df.empty:
        return pd.DataFrame(columns=["code", "name", "close", "change", "volume"])

    df["close"] = df["close"].fillna(0)
    df["change"] = df["change"].fillna(0)
    df["volume"] = df["volume"].fillna(0)

    df = df[df["close"] > 0]

    return df


def calc_score(row):
    score = 0

    change = float(row["change"])
    volume = float(row["volume"])
    close = float(row["close"])

    if change > 0:
        score += 20
    elif change < 0:
        score -= 10

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

            item = {
                "code": str(row["code"]),
                "name": str(row["name"]),
                "industry": "",
                "settle_date": "",
                "stars": stars_from_score(score),
                "strong_score": round(score, 2),
                "bias": 0,
                "short_alarm": "否",
                "long_alarm": "否",
            }

            result.append(item)
        except Exception:
            continue

    result = sorted(result, key=lambda x: x["strong_score"], reverse=True)

    return {
        "bullish": result[:50],
        "bearish": [],
        "warrants": [],
    }
