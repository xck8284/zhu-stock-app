import sys
sys.path.append(".")

from zhustock_app import build_all_excel

def run_analysis_core():
    try:
        result = build_all_excel(logger=print)

        # 👉 這裡你要確保有這三個key
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
