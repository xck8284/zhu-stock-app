def run_analysis_core():
    from zhustock_app import build_all_excel

    result = build_all_excel(logger=print)

    def df_to_list(df):
        if df is None:
            return []
        try:
            return df.to_dict(orient="records")
        except Exception:
            return []

    bullish = df_to_list(result.get("CLIENT_BULLISH"))
    bearish = df_to_list(result.get("CLIENT_BEARISH"))

    return {
        "bullish": bullish,
        "bearish": bearish,
        "warrants": []
    }
