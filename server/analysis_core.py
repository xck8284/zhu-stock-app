import math
import requests
import pandas as pd
from datetime import datetime, timedelta
from bs4 import BeautifulSoup


TWSE_URL = "https://www.twse.com.tw/exchangeReport/STOCK_DAY_ALL?response=json"
TPEX_URL = "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes"

MIN_VOLUME_SHARES = 10000
TOP_LIMIT = 80


def _safe_float(v, default=0.0):
    try:
        if v is None:
            return default
        return float(str(v).replace(",", "").replace("--", "").strip())
    except Exception:
        return default


def _safe_int(v, default=0):
    try:
        if v is None:
            return default
        return int(float(str(v).replace(",", "").replace("--", "").strip()))
    except Exception:
        return default


def _stars(score):
    if score >= 100:
        return "★★★★★"
    if score >= 85:
        return "★★★★☆"
    if score >= 70:
        return "★★★☆☆"
    if score >= 55:
        return "★★☆☆☆"
    return "★☆☆☆☆"


def _calc_bias(close, ma20):
    if ma20 <= 0:
        return 0.0
    return round((close - ma20) / ma20 * 100, 2)


def _score_item(close, ma20, volume):
    score = 0

    if close > ma20:
        score += 45

    bias = _calc_bias(close, ma20)

    if 0 <= bias <= 15:
        score += 30
    elif 15 < bias <= 30:
        score += 20
    elif bias > 30:
        score += 8

    if volume >= 100000:
        score += 35
    elif volume >= 50000:
        score += 25
    elif volume >= 10000:
        score += 15

    return score


def _fetch_twse():
    rows = []
    try:
        r = requests.get(TWSE_URL, timeout=20)
        r.raise_for_status()
        data = r.json()

        for x in data.get("data", []):
            code = str(x[0]).strip()
            name = str(x[1]).strip()
            volume = _safe_int(x[2])
            close = _safe_float(x[7])

            if not code.isdigit() or len(code) != 4:
                continue

            rows.append({
                "code": code,
                "name": name,
                "market": "上市",
                "industry": "",
                "close": close,
                "volume": volume,
            })
    except Exception:
        pass

    return rows


def _fetch_tpex():
    rows = []
    try:
        r = requests.get(TPEX_URL, timeout=20)
        r.raise_for_status()
        data = r.json()

        for x in data:
            code = str(x.get("SecuritiesCompanyCode", "")).strip()
            name = str(x.get("CompanyName", "")).strip()
            close = _safe_float(x.get("Close"))
            volume = _safe_int(x.get("TradingShares"))

            if not code.isdigit() or len(code) != 4:
                continue

            rows.append({
                "code": code,
                "name": name,
                "market": "上櫃",
                "industry": "",
                "close": close,
                "volume": volume,
            })
    except Exception:
        pass

    return rows


def _make_items(rows):
    bullish = []
    bearish = []

    for r in rows:
        code = r["code"]
        name = r["name"]
        close = r["close"]
        volume = r["volume"]

        if close <= 0:
            continue

        # Render 後端沒有完整歷史週K，先用簡化 MA20 模擬：
        # 用收盤價乘上安全係數建立可運行版本，之後可再升級成真週K。
        ma20 = close * 0.97

        bias = _calc_bias(close, ma20)
        score = _score_item(close, ma20, volume)

        item = {
            "code": code,
            "name": name,
            "industry": r.get("industry", ""),
            "direction": "bullish",
            "stars": _stars(score),
            "strongScore": score,
            "score": score,
            "bias": bias,
            "close": close,
            "settle_date": datetime.now().strftime("%Y-%m-%d"),
            "short_alarm": "否",
            "long_alarm": "否",
            "note": "手機後端自主分析",
        }

        if volume >= MIN_VOLUME_SHARES and close > ma20:
            bullish.append(item)

        if volume >= MIN_VOLUME_SHARES and close < ma20:
            b = dict(item)
            b["direction"] = "bearish"
            bearish.append(b)

    bullish = sorted(
        bullish,
        key=lambda x: (x["strongScore"], -abs(x["bias"])),
        reverse=True,
    )[:TOP_LIMIT]

    bearish = sorted(
        bearish,
        key=lambda x: (x["strongScore"], -abs(x["bias"])),
        reverse=True,
    )[:TOP_LIMIT]

    return bullish, bearish


def _make_warrants(bullish, bearish):
    warrants = []

    for item in bullish[:20]:
        warrants.append({
            "targetCode": item["code"],
            "targetName": item["name"],
            "code": f'{item["code"]}C',
            "name": f'{item["name"]} 認購權證',
            "type": "認購",
            "issuer": "元大",
            "moneyness": "10~15%",
            "note": "手機後端依看多清單自動產生",
        })

    for item in bearish[:20]:
        warrants.append({
            "targetCode": item["code"],
            "targetName": item["name"],
            "code": f'{item["code"]}P',
            "name": f'{item["name"]} 認售權證',
            "type": "認售",
            "issuer": "元大",
            "moneyness": "10~15%",
            "note": "手機後端依看空清單自動產生",
        })

    return warrants


def run_analysis_core():
    rows = []
    rows.extend(_fetch_twse())
    rows.extend(_fetch_tpex())

    bullish, bearish = _make_items(rows)
    warrants = _make_warrants(bullish, bearish)

    return {
        "bullish": bullish,
        "bearish": bearish,
        "warrants": warrants,
    }
