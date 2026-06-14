# -*- coding: utf-8 -*-
import random
from datetime import datetime


def run_analysis():
    """
    網頁版專用分析入口。
    目前先建立正式資料格式，之後再把假資料換成真實台股分析。
    不影響電腦版 zhustock_app.py。
    """

    bullish = [
        make_stock("2330", "台積電", "半導體", "看多"),
        make_stock("3017", "奇鋐", "散熱", "看多"),
        make_stock("3661", "世芯-KY", "IC設計", "看多"),
        make_stock("2454", "聯發科", "IC設計", "看多"),
        make_stock("2382", "廣達", "AI伺服器", "看多"),
    ]

    bearish = [
        make_stock("2603", "長榮", "航運", "看空"),
        make_stock("2615", "萬海", "航運", "看空"),
        make_stock("3481", "群創", "面板", "看空"),
    ]

    warrants = [
        {
            "stock_id": "2330",
            "stock_name": "台積電",
            "warrant_id": "088888",
            "warrant_name": "台積電元大購01",
            "issuer": "元大",
            "type": "認購",
            "days_left": 120,
            "moneyness": "12%",
        }
    ]

    return {
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "bullish": bullish,
        "bearish": bearish,
        "warrants": warrants,
    }


def make_stock(stock_id, name, industry, direction):
    score = random.randint(90, 140)

    return {
        "stock_id": stock_id,
        "name": name,
        "industry": industry,
        "direction": direction,
        "stars": score_to_stars(score),
        "strong_score": score,
        "bias": f"{random.randint(5, 20)}%",
        "short_alarm": "否",
        "long_alarm": "否",
    }


def score_to_stars(score):
    if score >= 120:
        return "★★★★★"
    if score >= 100:
        return "★★★★☆"
    if score >= 80:
        return "★★★☆☆"
    return "★★☆☆☆"
