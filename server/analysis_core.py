import pandas as pd


def safe_pct(a, b):
    try:
        if b is None or b == 0 or pd.isna(a) or pd.isna(b):
            return None
        return round((a - b) / b * 100, 2)
    except Exception:
        return None


def calc_training_score(grp, trend_info):
    latest = grp.iloc[-1]
    score = 0
    tags = []

    close_ = latest["週收盤價"]
    ma3 = latest["週3MA"]
    ma5 = latest["週5MA"]
    ma10 = latest["週10MA"]
    ma20 = latest["週20MA"]
    vol = latest["週成交量(張)"]
    vol5 = latest["量5MA"]
    vol20 = latest["量20MA"]
    high13 = latest["近13週最高"]
    high26 = latest["近26週最高"]
    high52 = latest["近52週最高"]
    slope20 = latest["20MA斜率"]
    open_ = latest["週開盤價"]

    if pd.notna(ma20) and close_ >= ma20:
        score += 18
        tags.append("站上20MA")
    elif pd.notna(ma20) and close_ >= ma20 * 0.97:
        score += 8
        tags.append("接近20MA")

    if pd.notna(ma3) and pd.notna(ma5) and pd.notna(ma10) and pd.notna(ma20):
        if ma3 >= ma5 >= ma10 >= ma20:
            score += 18
            tags.append("均線多頭")
        elif ma3 >= ma5 and ma5 >= ma20:
            score += 10
            tags.append("均線轉強")

    if pd.notna(slope20):
        if slope20 > 0:
            score += 8
            tags.append("20MA上彎")
        elif slope20 > -0.05:
            score += 3
            tags.append("20MA走平")

    if pd.notna(vol) and pd.notna(vol20) and vol20 > 0:
        vr20 = vol / vol20
        if vr20 >= 2.0:
            score += 16
            tags.append("量能爆發")
        elif vr20 >= 1.4:
            score += 10
            tags.append("量能放大")
        elif vr20 >= 1.0:
            score += 5
            tags.append("量能不弱")

    if pd.notna(vol) and pd.notna(vol5) and vol5 > 0:
        vr5 = vol / vol5
        if vr5 >= 1.5:
            score += 6
            tags.append("短線量增")

    if pd.notna(high13) and high13 > 0:
        d13 = safe_pct(close_, high13)
        if d13 is not None:
            if d13 >= -2:
                score += 10
                tags.append("近13週高")
            elif d13 >= -5:
                score += 5
                tags.append("接近13週高")

    if pd.notna(high26) and high26 > 0:
        d26 = safe_pct(close_, high26)
        if d26 is not None:
            if d26 >= -3:
                score += 8
                tags.append("近26週高")
            elif d26 >= -8:
                score += 4
                tags.append("接近26週高")

    if pd.notna(high52) and high52 > 0:
        d52 = safe_pct(close_, high52)
        if d52 is not None:
            if d52 >= -5:
                score += 8
                tags.append("接近52週高")
            elif d52 >= -12:
                score += 4
                tags.append("中長期強")

    if trend_info is not None:
        line_dist = trend_info.get("line_distance_pct")
        box_dist = trend_info.get("box_distance_pct")
        strict_ok = trend_info.get("strict_ok", False)
        training_hold_ok = trend_info.get("training_hold_ok", False)

        if strict_ok:
            score += 20
            tags.append("最新正式突破")

        if training_hold_ok:
            score += 24
            tags.append("突破後守穩")
        else:
            if line_dist is not None:
                if line_dist >= -1.5:
                    score += 8
                    tags.append("接近趨勢線")
                elif line_dist >= -4:
                    score += 4
                    tags.append("逼近趨勢線")

            if box_dist is not None:
                if box_dist >= -1.5:
                    score += 8
                    tags.append("接近盤整突破")
                elif box_dist >= -4:
                    score += 4
                    tags.append("逼近盤整突破")

        mem_bonus = trend_info.get("memory_bonus", 0)
        if mem_bonus > 0:
            score += mem_bonus
            tags.append("經驗加分")
        elif mem_bonus < 0:
            score += mem_bonus
            tags.append("經驗扣分")

    if pd.notna(open_) and pd.notna(close_) and open_ > 0:
        body_pct_val = safe_pct(close_, open_)
        if body_pct_val is not None:
            if body_pct_val >= 8:
                score += 10
                tags.append("長紅強攻")
            elif body_pct_val >= 4:
                score += 6
                tags.append("中紅K")

    if pd.notna(vol):
        if vol >= 100000:
            score += 8
            tags.append("週量極大")
        elif vol >= 80000:
            score += 5
            tags.append("週量大")
        elif vol >= 50000:
            score += 3
            tags.append("週量達標")

    uniq_tags = []
    for t in tags:
        if t not in uniq_tags:
            uniq_tags.append(t)

    return round(score, 2), "、".join(uniq_tags[:12])


def run_analysis_core():
    return {
        "bullish": [],
        "bearish": [],
        "warrants": []
    }
