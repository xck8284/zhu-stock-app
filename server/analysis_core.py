from zhustock_app import build_all_excel

def run_analysis_core():
    try:
        result = build_all_excel()

        # 👉 這裡你要依你實際result結構調整
        bullish = result.get("bullish", [])
        bearish = result.get("bearish", [])
        warrants = result.get("warrants", [])

        return {
            "bullish": bullish,
            "bearish": bearish,
            "warrants": warrants
        }

    except Exception as e:
        return {
            "bullish": [],
            "bearish": [],
            "warrants": [],
            "error": str(e)
        }
