import random

def run_analysis():

    sample_stocks = [
        ("2330","台積電"),
        ("3017","奇鋐"),
        ("3661","世芯"),
        ("2454","聯發科"),
        ("2382","廣達"),
    ]

    results = []

    for stock_id, name in sample_stocks:

        results.append({
            "stock_id": stock_id,
            "name": name,
            "stars": "★★★★★",
            "strong_score": random.randint(90,140),
            "bias": str(random.randint(5,20)) + "%"
        })

    return results
