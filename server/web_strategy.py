# -*- coding: utf-8 -*-
"""
網頁版策略分析（對齊桌面版 zhustock_app 核心邏輯，不含記憶檔）。

使用者策略：
- 看多：StrongScore>=100、星等>=5、週量>=10000、站上週20MA+趨勢線突破守穩、乖離率無上限
- 看空：桌面版 BEARISH_TRAINING_POOL 邏輯
- 權證：剩餘天數 90~120、全部發行券商
"""

from __future__ import annotations

import json
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta

import pandas as pd
import requests

SESSION = requests.Session()
SESSION.headers.update(
    {
        "User-Agent": "Mozilla/5.0 (ZHU-STOCK-WEB/1.0)",
        "Accept": "application/json,text/plain,*/*",
    }
)


def _http_get(url, params=None, timeout=20):
    return requests.get(url, params=params, timeout=timeout, headers=SESSION.headers)

# --- 策略常數（對齊桌面版） ---
MIN_WEEKLY_VOLUME = 10000
LOOKBACK_WEEKS = 60
MIN_PEAK_GAP = 4
MAX_PEAK_GAP = 40
PEAK_WINDOW = 2
BREAKOUT_BUFFER_PCT = 0.003
BOX_BREAKOUT_BUFFER_PCT = 0.003
HOLD_BUFFER_PCT = 0.0
MIN_DESCENT_PCT = 0.01
BREAKDOWN_BUFFER_PCT = 0.003
HOLD_ABOVE_BUFFER_PCT = 0.0
MIN_ASCENT_PCT = 0.01
BOX_LOOKBACK_WEEKS = 10
BOX_MIN_WEEKS = 4
TRAINING_SCORE_THRESHOLD = 55
BEARISH_TRAINING_SCORE_THRESHOLD = 55

# 網頁展示篩選（使用者指定）
DISPLAY_MIN_SCORE = 100
DISPLAY_MIN_STARS = 5
WARRANT_MIN_DAYS = 90
WARRANT_MAX_DAYS = 120

MAX_RESULTS = 100
MAX_WARRANT_STOCKS = 8
# 週K 20MA + 趨勢線需約 24 週以上；並行抓取仍可在 2～3 分鐘內完成
MIN_TRADING_DAYS_TARGET = 165
HISTORY_CALENDAR_DAYS = 240
FETCH_WORKERS = 12

# 權證專用（桌面版 build_warrant_fastscan 同款）
WARRANT_MIN_SCORE = 100
WARRANT_MIN_STARS = 5


def clean_text(value):
    if value is None:
        return ""
    return str(value).strip().replace("\u3000", " ").strip()


def parse_number(value, is_int=False):
    text = clean_text(value).replace(",", "").replace("--", "").replace("X", "")
    if not text:
        return None
    try:
        num = float(text)
        return int(num) if is_int else num
    except Exception:
        return None


def normalize_code(code):
    code = clean_text(code)
    match = re.match(r"^(\d{4,6})", code)
    return match.group(1) if match else code


def is_valid_stock_code(code):
    return bool(re.fullmatch(r"\d{4}", str(code)))


def parse_date_any(value):
    text = clean_text(value)
    if not text:
        return None
    text = text.replace(".", "/").replace("-", "/")
    for fmt in ["%Y/%m/%d", "%Y%m%d"]:
        try:
            return datetime.strptime(text[:10] if fmt == "%Y/%m/%d" else text[:8], fmt).date()
        except Exception:
            pass
    match = re.search(r"(\d{2,3})/(\d{1,2})/(\d{1,2})", text)
    if match:
        try:
            year = int(match.group(1))
            if year < 1911:
                year += 1911
            return datetime(year, int(match.group(2)), int(match.group(3))).date()
        except Exception:
            return None
    return None


def safe_pct(a, b):
    if b in [0, None] or pd.isna(a) or pd.isna(b):
        return None
    return (a / b - 1.0) * 100.0


def clamp(x, low, high):
    return max(low, min(high, x))


def body_top(open_price, close_price):
    if pd.isna(open_price) or pd.isna(close_price):
        return None
    return max(open_price, close_price)


def body_bottom(open_price, close_price):
    if pd.isna(open_price) or pd.isna(close_price):
        return None
    return min(open_price, close_price)


def is_red_k(open_price, close_price):
    if pd.isna(open_price) or pd.isna(close_price):
        return False
    return close_price > open_price


def is_black_k(open_price, close_price):
    if pd.isna(open_price) or pd.isna(close_price):
        return False
    return close_price < open_price


def is_real_body_breakout(open_price, close_price, line_y, breakout_buffer_pct=0.0):
    if pd.isna(open_price) or pd.isna(close_price) or pd.isna(line_y) or line_y <= 0:
        return False
    if not is_red_k(open_price, close_price):
        return False
    top = body_top(open_price, close_price)
    if pd.isna(top):
        return False
    if close_price <= line_y * (1 + breakout_buffer_pct):
        return False
    if top <= line_y:
        return False
    return True


def is_real_body_breakdown(open_price, close_price, line_y, breakdown_buffer_pct=0.0):
    if pd.isna(open_price) or pd.isna(close_price) or pd.isna(line_y) or line_y <= 0:
        return False
    if not is_black_k(open_price, close_price):
        return False
    bottom = body_bottom(open_price, close_price)
    if pd.isna(bottom):
        return False
    if close_price >= line_y * (1 - breakdown_buffer_pct):
        return False
    if bottom >= line_y:
        return False
    return True


def star_text(n):
    n = int(clamp(n, 1, 5))
    return "★" * n + "☆" * (5 - n)


def score_to_star(score):
    if score >= 90:
        return 5
    if score >= 75:
        return 4
    if score >= 60:
        return 3
    if score >= 45:
        return 2
    return 1


def calc_bias_pct(close_, ma20):
    if pd.isna(close_) or pd.isna(ma20) or ma20 in [0, None]:
        return None
    return (close_ - ma20) / ma20 * 100.0


def get_memory_bonus(_direction, _code, _trend_info):
    return 0, ""


def line_value(x1, y1, x2, y2, x):
    if x2 == x1:
        return y1
    return y1 + (y2 - y1) * ((x - x1) / (x2 - x1))


def find_pivot_highs_above_20ma(sub_df, window=2):
    pivots = []
    n = len(sub_df)
    for i in range(window, n - window):
        cur_high = sub_df.iloc[i]["週最高價"]
        cur_ma20 = sub_df.iloc[i]["週20MA"]
        if pd.isna(cur_high) or pd.isna(cur_ma20) or cur_high <= cur_ma20:
            continue
        left = sub_df.iloc[i - window : i]["週最高價"].tolist()
        right = sub_df.iloc[i + 1 : i + 1 + window]["週最高價"].tolist()
        if any(pd.isna(x) for x in left) or any(pd.isna(x) for x in right):
            continue
        if all(cur_high >= x for x in left) and all(cur_high >= x for x in right):
            pivots.append(i)
    return pivots


def find_pivot_lows_below_20ma(sub_df, window=2):
    pivots = []
    n = len(sub_df)
    for i in range(window, n - window):
        cur_low = sub_df.iloc[i]["週最低價"]
        cur_ma20 = sub_df.iloc[i]["週20MA"]
        if pd.isna(cur_low) or pd.isna(cur_ma20) or cur_low >= cur_ma20:
            continue
        left = sub_df.iloc[i - window : i]["週最低價"].tolist()
        right = sub_df.iloc[i + 1 : i + 1 + window]["週最低價"].tolist()
        if any(pd.isna(x) for x in left) or any(pd.isna(x) for x in right):
            continue
        if all(cur_low <= x for x in left) and all(cur_low <= x for x in right):
            pivots.append(i)
    return pivots


def get_recent_box_info(work_df, latest_idx, lookback_weeks=10):
    if latest_idx < 2:
        return None
    start_idx = max(0, latest_idx - lookback_weeks)
    end_idx = latest_idx - 1
    box_df = work_df.iloc[start_idx : end_idx + 1].copy()
    if len(box_df) < BOX_MIN_WEEKS:
        return None
    box_high = box_df["週最高價"].max()
    box_low = box_df["週最低價"].min()
    if pd.isna(box_high) or pd.isna(box_low) or box_high <= 0 or box_low <= 0:
        return None
    return {
        "box_start_idx": start_idx,
        "box_end_idx": end_idx,
        "box_high": float(box_high),
        "box_low": float(box_low),
    }


def no_close_back_below_trendline_after_breakout(work, i, j, breakout_idx, buffer_pct=0.0):
    y1 = work.iloc[i]["週最高價"]
    y2 = work.iloc[j]["週最高價"]
    if pd.isna(y1) or pd.isna(y2):
        return False
    for k in range(breakout_idx + 1, len(work)):
        line_y = line_value(i, y1, j, y2, k)
        close_k = work.iloc[k]["週收盤價"]
        if pd.isna(line_y) or pd.isna(close_k):
            return False
        if close_k < line_y * (1 - buffer_pct):
            return False
    return True


def no_close_back_above_trendline_after_breakdown(work, i, j, breakdown_idx, buffer_pct=0.0):
    y1 = work.iloc[i]["週最低價"]
    y2 = work.iloc[j]["週最低價"]
    if pd.isna(y1) or pd.isna(y2):
        return False
    for k in range(breakdown_idx + 1, len(work)):
        line_y = line_value(i, y1, j, y2, k)
        close_k = work.iloc[k]["週收盤價"]
        if pd.isna(line_y) or pd.isna(close_k):
            return False
        if close_k > line_y * (1 + buffer_pct):
            return False
    return True


def find_first_valid_breakout_and_hold(work, i, j, breakout_buffer_pct=0.003, hold_buffer_pct=0.0):
    y1 = work.iloc[i]["週最高價"]
    y2 = work.iloc[j]["週最高價"]
    if pd.isna(y1) or pd.isna(y2):
        return None
    for k in range(j + 1, len(work)):
        line_y = line_value(i, y1, j, y2, k)
        open_k = work.iloc[k]["週開盤價"]
        close_k = work.iloc[k]["週收盤價"]
        if pd.isna(line_y) or pd.isna(open_k) or pd.isna(close_k):
            continue
        if is_real_body_breakout(open_k, close_k, line_y, breakout_buffer_pct=breakout_buffer_pct):
            if no_close_back_below_trendline_after_breakout(work, i, j, k, buffer_pct=hold_buffer_pct):
                return {
                    "breakout_idx": k,
                    "breakout_date": work.iloc[k]["週結算日期"],
                    "breakout_open": open_k,
                    "breakout_close": close_k,
                    "breakout_line": line_y,
                }
    return None


def find_first_valid_breakdown_and_hold(work, i, j, breakdown_buffer_pct=0.003, hold_buffer_pct=0.0):
    y1 = work.iloc[i]["週最低價"]
    y2 = work.iloc[j]["週最低價"]
    if pd.isna(y1) or pd.isna(y2):
        return None
    for k in range(j + 1, len(work)):
        line_y = line_value(i, y1, j, y2, k)
        open_k = work.iloc[k]["週開盤價"]
        close_k = work.iloc[k]["週收盤價"]
        if pd.isna(line_y) or pd.isna(open_k) or pd.isna(close_k):
            continue
        if is_real_body_breakdown(open_k, close_k, line_y, breakdown_buffer_pct=breakdown_buffer_pct):
            if no_close_back_above_trendline_after_breakdown(work, i, j, k, buffer_pct=hold_buffer_pct):
                return {
                    "breakdown_idx": k,
                    "breakdown_date": work.iloc[k]["週結算日期"],
                    "breakdown_open": open_k,
                    "breakdown_close": close_k,
                    "breakdown_line": line_y,
                }
    return None


def analyze_best_descending_trendline(sub_df):
    if len(sub_df) < 24:
        return None

    code = str(sub_df.iloc[-1]["股票代號"])
    work = sub_df.tail(LOOKBACK_WEEKS).copy().reset_index(drop=True)
    pivots = find_pivot_highs_above_20ma(work, window=PEAK_WINDOW)
    if len(pivots) < 2:
        return None

    candidates = []
    latest_idx = len(work) - 1
    prev_idx = latest_idx - 1
    if prev_idx < 0:
        return None

    latest_open = work.iloc[latest_idx]["週開盤價"]
    latest_close = work.iloc[latest_idx]["週收盤價"]
    prev_close = work.iloc[prev_idx]["週收盤價"]

    for a in range(len(pivots) - 1):
        for b in range(a + 1, len(pivots)):
            i = pivots[a]
            j = pivots[b]
            gap = j - i
            if gap < MIN_PEAK_GAP or gap > MAX_PEAK_GAP:
                continue

            y1 = work.iloc[i]["週最高價"]
            y2 = work.iloc[j]["週最高價"]
            ma1 = work.iloc[i]["週20MA"]
            ma2 = work.iloc[j]["週20MA"]
            if any(pd.isna(x) for x in [y1, y2, ma1, ma2, latest_open, latest_close, prev_close]):
                continue
            if y1 <= ma1 or y2 <= ma2 or y2 >= y1:
                continue

            descent_pct = (y1 - y2) / y1 if y1 not in [0, None] else 0
            if descent_pct < MIN_DESCENT_PCT:
                continue

            latest_line = line_value(i, y1, j, y2, latest_idx)
            prev_line = line_value(i, y1, j, y2, prev_idx)
            if pd.isna(latest_line) or pd.isna(prev_line) or latest_line <= 0 or prev_line <= 0:
                continue

            line_distance_pct = safe_pct(latest_close, latest_line)
            prev_line_distance_pct = safe_pct(prev_close, prev_line)
            line_break_now = is_real_body_breakout(
                latest_open, latest_close, latest_line, breakout_buffer_pct=BREAKOUT_BUFFER_PCT
            )
            line_break_prev = prev_close > prev_line * (1 + BREAKOUT_BUFFER_PCT)

            box_info = get_recent_box_info(work, latest_idx, BOX_LOOKBACK_WEEKS)
            if box_info is None:
                continue

            box_high = box_info["box_high"]
            box_low = box_info["box_low"]
            box_distance_pct = safe_pct(latest_close, box_high)
            prev_box_distance_pct = safe_pct(prev_close, box_high)
            box_break_now = is_red_k(latest_open, latest_close) and latest_close > box_high * (
                1 + BOX_BREAKOUT_BUFFER_PCT
            )
            box_break_prev = prev_close > box_high * (1 + BOX_BREAKOUT_BUFFER_PCT)
            strict_ok = line_break_now and (not line_break_prev) and box_break_now and (not box_break_prev)

            hold_info = find_first_valid_breakout_and_hold(
                work=work,
                i=i,
                j=j,
                breakout_buffer_pct=BREAKOUT_BUFFER_PCT,
                hold_buffer_pct=HOLD_BUFFER_PCT,
            )
            training_hold_ok = hold_info is not None

            base_score = (
                j * 1000
                + clamp((line_distance_pct if line_distance_pct is not None else -999), -20, 20) * 60
                + clamp((box_distance_pct if box_distance_pct is not None else -999), -20, 20) * 70
                + descent_pct * 1000
                + (250 if strict_ok else 0)
                + (180 if training_hold_ok else 0)
            )

            memory_bonus, memory_text = get_memory_bonus("bullish", code, {"i": i, "j": j})
            candidates.append(
                {
                    "i": i,
                    "j": j,
                    "latest_line": latest_line,
                    "line_distance_pct": line_distance_pct,
                    "box_high": box_high,
                    "box_low": box_low,
                    "box_distance_pct": box_distance_pct,
                    "strict_ok": strict_ok,
                    "training_hold_ok": training_hold_ok,
                    "hold_info": hold_info,
                    "work_df": work,
                    "memory_bonus": memory_bonus,
                    "memory_text": memory_text,
                    "score": base_score + memory_bonus,
                    "line_break_now": line_break_now,
                    "line_break_prev": line_break_prev,
                }
            )

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: (
            x["score"],
            x["j"],
            x.get("descent_pct", 0),
            x["line_distance_pct"] if x["line_distance_pct"] is not None else -999,
        ),
        reverse=True,
    )
    return candidates[0]


def analyze_best_ascending_trendline(sub_df):
    if len(sub_df) < 24:
        return None

    code = str(sub_df.iloc[-1]["股票代號"])
    work = sub_df.tail(LOOKBACK_WEEKS).copy().reset_index(drop=True)
    pivots = find_pivot_lows_below_20ma(work, window=PEAK_WINDOW)
    if len(pivots) < 2:
        return None

    candidates = []
    latest_idx = len(work) - 1
    prev_idx = latest_idx - 1
    if prev_idx < 0:
        return None

    latest_open = work.iloc[latest_idx]["週開盤價"]
    latest_close = work.iloc[latest_idx]["週收盤價"]
    prev_close = work.iloc[prev_idx]["週收盤價"]

    for a in range(len(pivots) - 1):
        for b in range(a + 1, len(pivots)):
            i = pivots[a]
            j = pivots[b]
            gap = j - i
            if gap < MIN_PEAK_GAP or gap > MAX_PEAK_GAP:
                continue

            y1 = work.iloc[i]["週最低價"]
            y2 = work.iloc[j]["週最低價"]
            ma1 = work.iloc[i]["週20MA"]
            ma2 = work.iloc[j]["週20MA"]
            if any(pd.isna(x) for x in [y1, y2, ma1, ma2, latest_open, latest_close, prev_close]):
                continue
            if y1 >= ma1 or y2 >= ma2 or y2 <= y1:
                continue

            ascent_pct = (y2 - y1) / y1 if y1 not in [0, None] else 0
            if ascent_pct < MIN_ASCENT_PCT:
                continue

            latest_line = line_value(i, y1, j, y2, latest_idx)
            prev_line = line_value(i, y1, j, y2, prev_idx)
            if pd.isna(latest_line) or pd.isna(prev_line) or latest_line <= 0 or prev_line <= 0:
                continue

            line_distance_pct = safe_pct(latest_close, latest_line)
            line_break_now = is_real_body_breakdown(
                latest_open, latest_close, latest_line, breakdown_buffer_pct=BREAKDOWN_BUFFER_PCT
            )
            line_break_prev = prev_close < prev_line * (1 - BREAKDOWN_BUFFER_PCT)

            hold_info = find_first_valid_breakdown_and_hold(
                work=work,
                i=i,
                j=j,
                breakdown_buffer_pct=BREAKDOWN_BUFFER_PCT,
                hold_buffer_pct=HOLD_ABOVE_BUFFER_PCT,
            )
            training_hold_ok = hold_info is not None

            base_score = (
                j * 1000
                + clamp(((-line_distance_pct) if line_distance_pct is not None else -999), -20, 20) * 70
                + ascent_pct * 1000
                + (220 if line_break_now and (not line_break_prev) else 0)
                + (180 if training_hold_ok else 0)
            )

            memory_bonus, memory_text = get_memory_bonus("bearish", code, {"i": i, "j": j})
            candidates.append(
                {
                    "i": i,
                    "j": j,
                    "latest_line": latest_line,
                    "line_distance_pct": line_distance_pct,
                    "training_hold_ok": training_hold_ok,
                    "hold_info": hold_info,
                    "work_df": work,
                    "memory_bonus": memory_bonus,
                    "memory_text": memory_text,
                    "score": base_score + memory_bonus,
                    "line_break_now": line_break_now,
                    "line_break_prev": line_break_prev,
                }
            )

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: (
            x["score"],
            x["j"],
            x.get("ascent_pct", 0),
            -(x["line_distance_pct"] if x["line_distance_pct"] is not None else 999),
        ),
        reverse=True,
    )
    return candidates[0]


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

    if pd.notna(vol) and pd.notna(vol5) and vol5 > 0 and vol / vol5 >= 1.5:
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
        if trend_info.get("strict_ok"):
            score += 20
            tags.append("最新正式突破")
        if trend_info.get("training_hold_ok"):
            score += 24
            tags.append("突破後守穩")
        else:
            if line_dist is not None and line_dist >= -1.5:
                score += 8
                tags.append("接近趨勢線")
            elif line_dist is not None and line_dist >= -4:
                score += 4
                tags.append("逼近趨勢線")
            if box_dist is not None and box_dist >= -1.5:
                score += 8
                tags.append("接近盤整突破")
            elif box_dist is not None and box_dist >= -4:
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
        elif vol >= 60000:
            score += 5
            tags.append("週量大")
        elif vol >= 10000:
            score += 3
            tags.append("週量達標")

    uniq_tags = []
    for tag in tags:
        if tag not in uniq_tags:
            uniq_tags.append(tag)
    return round(score, 2), "、".join(uniq_tags[:12])


def calc_bearish_training_score(grp, trend_info):
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
    low13 = latest["近13週最低"]
    low26 = latest["近26週最低"]
    low52 = latest["近52週最低"]
    slope20 = latest["20MA斜率"]
    open_ = latest["週開盤價"]

    if pd.notna(ma20) and close_ < ma20:
        score += 18
        tags.append("跌破20MA")
    elif pd.notna(ma20) and close_ <= ma20 * 1.03:
        score += 8
        tags.append("接近20MA下")

    if pd.notna(ma3) and pd.notna(ma5) and pd.notna(ma10) and pd.notna(ma20):
        if ma3 <= ma5 <= ma10 <= ma20:
            score += 18
            tags.append("均線空頭")
        elif ma3 <= ma5 and ma5 <= ma20:
            score += 10
            tags.append("均線轉弱")

    if pd.notna(slope20):
        if slope20 < 0:
            score += 8
            tags.append("20MA下彎")
        elif slope20 < 0.05:
            score += 3
            tags.append("20MA走平偏弱")

    if pd.notna(vol) and pd.notna(vol20) and vol20 > 0:
        vr20 = vol / vol20
        if vr20 >= 2.0:
            score += 16
            tags.append("量能放大")
        elif vr20 >= 1.4:
            score += 10
            tags.append("量能增強")
        elif vr20 >= 1.0:
            score += 5
            tags.append("量能不弱")

    if pd.notna(vol) and pd.notna(vol5) and vol5 > 0 and vol / vol5 >= 1.5:
        score += 6
        tags.append("短線量增")

    if pd.notna(low13) and low13 > 0:
        d13 = safe_pct(close_, low13)
        if d13 is not None:
            if d13 <= 2:
                score += 10
                tags.append("近13週低")
            elif d13 <= 5:
                score += 5
                tags.append("接近13週低")

    if trend_info is not None:
        line_dist = trend_info.get("line_distance_pct")
        if trend_info.get("line_break_now"):
            score += 20
            tags.append("最新正式跌破")
        if trend_info.get("training_hold_ok"):
            score += 24
            tags.append("跌破後守弱")
        if line_dist is not None and line_dist <= 1.5:
            score += 8
            tags.append("接近上升趨勢線")
        elif line_dist is not None and line_dist <= 4:
            score += 4
            tags.append("逼近上升趨勢線")

    if pd.notna(open_) and pd.notna(close_) and open_ > 0:
        body_pct_val = safe_pct(close_, open_)
        if body_pct_val is not None:
            if body_pct_val <= -8:
                score += 10
                tags.append("長黑轉弱")
            elif body_pct_val <= -4:
                score += 6
                tags.append("中黑K")

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
    for tag in tags:
        if tag not in uniq_tags:
            uniq_tags.append(tag)
    return round(score, 2), "、".join(uniq_tags[:12])


def calc_bullish_star_and_alarm_by_score(score, grp):
    latest = grp.iloc[-1]
    prev = grp.iloc[-2] if len(grp) >= 2 else latest
    close_ = latest["週收盤價"]
    ma3 = latest["週3MA"]
    ma5 = latest["週5MA"]
    ma20 = latest["週20MA"]
    prev_ma3 = prev["週3MA"]
    prev_ma5 = prev["週5MA"]
    bias = calc_bias_pct(close_, ma20)
    star = score_to_star(score)
    short_alarm = "否"
    long_alarm = "否"
    if pd.notna(prev_ma3) and pd.notna(prev_ma5) and pd.notna(ma3) and pd.notna(ma5):
        if prev_ma3 >= prev_ma5 and ma3 < ma5:
            short_alarm = "是"
    if pd.notna(ma20) and pd.notna(close_) and close_ < ma20:
        long_alarm = "是"
    return {
        "星等數值": star,
        "星等": star_text(star),
        "乖離率(%)": round(bias, 2) if bias is not None else None,
        "短線停利Alarm": short_alarm,
        "長線停利Alarm": long_alarm,
    }


def calc_bearish_star_and_alarm_by_score(score, grp):
    latest = grp.iloc[-1]
    prev = grp.iloc[-2] if len(grp) >= 2 else latest
    close_ = latest["週收盤價"]
    ma3 = latest["週3MA"]
    ma5 = latest["週5MA"]
    ma20 = latest["週20MA"]
    prev_ma3 = prev["週3MA"]
    prev_ma5 = prev["週5MA"]
    bias = calc_bias_pct(close_, ma20)
    star = score_to_star(score)
    short_alarm = "否"
    long_alarm = "否"
    if pd.notna(prev_ma3) and pd.notna(prev_ma5) and pd.notna(ma3) and pd.notna(ma5):
        if prev_ma3 <= prev_ma5 and ma3 > ma5:
            short_alarm = "是"
    if pd.notna(ma20) and pd.notna(close_) and close_ > ma20:
        long_alarm = "是"
    return {
        "星等數值": star,
        "星等": star_text(star),
        "乖離率(%)": round(bias, 2) if bias is not None else None,
        "短線回補Alarm": short_alarm,
        "長線回補Alarm": long_alarm,
    }


def build_weekly_k_from_daily(daily_df):
    if daily_df.empty:
        return pd.DataFrame()

    df = daily_df.copy()
    df["日期"] = pd.to_datetime(df["日期"])
    df = df.sort_values(["股票代號", "日期"]).reset_index(drop=True)
    df["週別"] = df["日期"].dt.to_period("W-FRI")

    rows = []
    for (code, week_period), grp in df.groupby(["股票代號", "週別"]):
        grp = grp.sort_values("日期")
        first_row = grp.iloc[0]
        last_row = grp.iloc[-1]
        rows.append(
            {
                "股票代號": code,
                "股票名稱": last_row["股票名稱"],
                "市場別": last_row["市場別"],
                "週別": str(week_period),
                "週結算日期": last_row["日期"].strftime("%Y-%m-%d"),
                "週開盤價": first_row["開盤價"],
                "週最高價": grp["最高價"].max(),
                "週最低價": grp["最低價"].min(),
                "週收盤價": last_row["收盤價"],
                "週成交量(張)": int(round(grp["成交股數"].fillna(0).sum() / 1000)),
            }
        )
    return pd.DataFrame(rows).sort_values(["股票代號", "週結算日期"]).reset_index(drop=True)


def calculate_weekly_indicators(weekly_df):
    if weekly_df.empty:
        return pd.DataFrame()

    df = weekly_df.copy()
    df["週結算日期"] = pd.to_datetime(df["週結算日期"])
    df = df.sort_values(["股票代號", "週結算日期"]).reset_index(drop=True)
    g = df.groupby("股票代號")

    df["週3MA"] = g["週收盤價"].transform(lambda s: s.rolling(3, min_periods=3).mean())
    df["週5MA"] = g["週收盤價"].transform(lambda s: s.rolling(5, min_periods=5).mean())
    df["週10MA"] = g["週收盤價"].transform(lambda s: s.rolling(10, min_periods=10).mean())
    df["週20MA"] = g["週收盤價"].transform(lambda s: s.rolling(20, min_periods=20).mean())
    df["量5MA"] = g["週成交量(張)"].transform(lambda s: s.rolling(5, min_periods=5).mean())
    df["量10MA"] = g["週成交量(張)"].transform(lambda s: s.rolling(10, min_periods=10).mean())
    df["量20MA"] = g["週成交量(張)"].transform(lambda s: s.rolling(20, min_periods=20).mean())
    df["20MA斜率"] = g["週20MA"].transform(lambda s: s.diff())
    df["5MA斜率"] = g["週5MA"].transform(lambda s: s.diff())
    df["近13週最高"] = g["週最高價"].transform(lambda s: s.rolling(13, min_periods=13).max())
    df["近26週最高"] = g["週最高價"].transform(lambda s: s.rolling(26, min_periods=26).max())
    df["近52週最高"] = g["週最高價"].transform(lambda s: s.rolling(52, min_periods=20).max())
    df["近13週最低"] = g["週最低價"].transform(lambda s: s.rolling(13, min_periods=13).min())
    df["近26週最低"] = g["週最低價"].transform(lambda s: s.rolling(26, min_periods=26).min())
    df["近52週最低"] = g["週最低價"].transform(lambda s: s.rolling(52, min_periods=20).min())
    df["是否站上週20MA"] = df["週收盤價"] >= df["週20MA"]
    df["是否跌破週20MA"] = df["週收盤價"] < df["週20MA"]
    df["均線多頭排列"] = (
        (df["週3MA"] >= df["週5MA"])
        & (df["週5MA"] >= df["週10MA"])
        & (df["週10MA"] >= df["週20MA"])
    )
    df["均線空頭排列"] = (
        (df["週3MA"] <= df["週5MA"])
        & (df["週5MA"] <= df["週10MA"])
        & (df["週10MA"] <= df["週20MA"])
    )
    return df


def build_master_df(daily_df):
    if daily_df.empty:
        return pd.DataFrame(columns=["股票代號", "產業別"])
    master = daily_df[["股票代號"]].drop_duplicates().copy()
    master["產業別"] = "未分類"
    return master


def build_training_pool(weekly_ma_df, master_df):
    rows = []
    industry_map = master_df[["股票代號", "產業別"]].drop_duplicates(subset=["股票代號"]).copy()
    df = weekly_ma_df.copy()
    df["週結算日期"] = pd.to_datetime(df["週結算日期"])
    df = df.sort_values(["股票代號", "週結算日期"]).reset_index(drop=True)

    for code, grp in df.groupby("股票代號"):
        grp = grp.sort_values("週結算日期").reset_index(drop=True).copy()
        latest = grp.iloc[-1]
        if pd.isna(latest["週20MA"]) or pd.isna(latest["週成交量(張)"]):
            continue
        if latest["週成交量(張)"] < MIN_WEEKLY_VOLUME:
            continue

        trend = analyze_best_descending_trendline(grp)
        score, tags = calc_training_score(grp, trend)
        base_ok = (
            latest["週收盤價"] >= latest["週20MA"] * 0.95 if pd.notna(latest["週20MA"]) else False
        ) or (latest["均線多頭排列"] if pd.notna(latest["均線多頭排列"]) else False)

        if not base_ok or trend is None or not trend["training_hold_ok"] or score < TRAINING_SCORE_THRESHOLD:
            continue

        industry = industry_map.loc[industry_map["股票代號"] == code, "產業別"]
        industry_val = industry.iloc[0] if len(industry) > 0 else "未分類"
        extra = calc_bullish_star_and_alarm_by_score(score, grp)

        rows.append(
            {
                "股票代號": str(code),
                "股票名稱": str(latest["股票名稱"]),
                "市場別": str(latest["市場別"]),
                "產業別": str(industry_val),
                "週結算日期": latest["週結算日期"].strftime("%Y-%m-%d"),
                "StrongScore": score,
                "最新一週成交量(張)": int(latest["週成交量(張)"]),
                "是否突破後守穩趨勢線": "是",
                **extra,
            }
        )

    if not rows:
        return pd.DataFrame()
    out = pd.DataFrame(rows)
    out = out.sort_values(["StrongScore", "星等數值", "最新一週成交量(張)", "股票代號"], ascending=[False, False, False, True])
    return out.reset_index(drop=True)


def build_bearish_pool(weekly_ma_df, master_df):
    rows = []
    industry_map = master_df[["股票代號", "產業別"]].drop_duplicates(subset=["股票代號"]).copy()
    df = weekly_ma_df.copy()
    df["週結算日期"] = pd.to_datetime(df["週結算日期"])
    df = df.sort_values(["股票代號", "週結算日期"]).reset_index(drop=True)

    for code, grp in df.groupby("股票代號"):
        grp = grp.sort_values("週結算日期").reset_index(drop=True).copy()
        latest = grp.iloc[-1]
        if pd.isna(latest["週20MA"]) or pd.isna(latest["週成交量(張)"]):
            continue
        if latest["週成交量(張)"] < MIN_WEEKLY_VOLUME or latest["週收盤價"] >= latest["週20MA"]:
            continue

        trend = analyze_best_ascending_trendline(grp)
        score, _tags = calc_bearish_training_score(grp, trend)
        base_ok = (
            latest["週收盤價"] <= latest["週20MA"] * 1.05 if pd.notna(latest["週20MA"]) else False
        ) or (latest["均線空頭排列"] if pd.notna(latest["均線空頭排列"]) else False)

        if not base_ok or trend is None or not trend["training_hold_ok"] or score < BEARISH_TRAINING_SCORE_THRESHOLD:
            continue

        industry = industry_map.loc[industry_map["股票代號"] == code, "產業別"]
        industry_val = industry.iloc[0] if len(industry) > 0 else "未分類"
        extra = calc_bearish_star_and_alarm_by_score(score, grp)

        rows.append(
            {
                "股票代號": str(code),
                "股票名稱": str(latest["股票名稱"]),
                "市場別": str(latest["市場別"]),
                "產業別": str(industry_val),
                "週結算日期": latest["週結算日期"].strftime("%Y-%m-%d"),
                "BearishScore": score,
                "最新一週成交量(張)": int(latest["週成交量(張)"]),
                **extra,
            }
        )

    if not rows:
        return pd.DataFrame()
    out = pd.DataFrame(rows)
    out = out.sort_values(["BearishScore", "星等數值", "最新一週成交量(張)", "股票代號"], ascending=[False, False, False, True])
    return out.reset_index(drop=True)


def filter_bullish_for_display(pool_df):
    """看多清單：對齊桌面版 CLIENT_BULLISH，顯示完整 training pool（趨勢突破守穩）。"""
    if pool_df is None or pool_df.empty:
        return pool_df
    work = pool_df.copy()
    work["StrongScore"] = pd.to_numeric(work["StrongScore"], errors="coerce")
    work["星等數值"] = pd.to_numeric(work["星等數值"], errors="coerce")
    work["最新一週成交量(張)"] = pd.to_numeric(work["最新一週成交量(張)"], errors="coerce")
    work = work[work["最新一週成交量(張)"] >= MIN_WEEKLY_VOLUME].copy()
    work = work.sort_values(
        ["StrongScore", "星等數值", "最新一週成交量(張)", "股票代號"],
        ascending=[False, False, False, True],
    )
    return work.head(MAX_RESULTS)


def filter_warrant_candidates(pool_df):
    """權證：StrongScore≥100 且 5 星（桌面版權證快篩）。"""
    if pool_df is None or pool_df.empty:
        return pool_df
    work = pool_df.copy()
    work["StrongScore"] = pd.to_numeric(work["StrongScore"], errors="coerce")
    work["星等數值"] = pd.to_numeric(work["星等數值"], errors="coerce")
    filtered = work[
        (work["StrongScore"] >= WARRANT_MIN_SCORE) & (work["星等數值"] >= WARRANT_MIN_STARS)
    ].copy()
    return filtered.sort_values(["StrongScore", "乖離率(%)", "股票代號"], ascending=[False, True, True])


def pool_to_api_items(pool_df, score_col="StrongScore"):
    items = []
    for _, row in pool_df.iterrows():
        bias = row.get("乖離率(%)")
        bias_text = f"{bias}%" if bias is not None and str(bias) != "" else ""
        items.append(
            {
                "stock_id": str(row["股票代號"]),
                "code": str(row["股票代號"]),
                "name": str(row.get("股票名稱", "")),
                "industry": str(row.get("產業別", "")),
                "market": str(row.get("市場別", "")),
                "direction": "看多" if score_col == "StrongScore" else "看空",
                "strong_score": float(row.get(score_col, 0)),
                "stars": str(row.get("星等", "")),
                "bias": bias_text,
                "settle_date": str(row.get("週結算日期", "")),
            }
        )
    return items


def extract_twse_table(tables):
    for table in tables or []:
        fields = [clean_text(x) for x in table.get("fields", [])]
        joined = "|".join(fields)
        need = ["證券代號", "證券名稱", "成交股數", "開盤價", "最高價", "最低價", "收盤價"]
        if all(x in joined for x in need):
            return table
    return None


def fetch_twse_daily_all(date_obj):
    url = "https://www.twse.com.tw/exchangeReport/MI_INDEX"
    params = {"response": "json", "date": date_obj.strftime("%Y%m%d"), "type": "ALLBUT0999"}
    try:
        response = _http_get(url, params=params, timeout=20)
        response.raise_for_status()
        payload = response.json()
    except Exception:
        return None

    stat = clean_text(payload.get("stat", ""))
    if any(x in stat for x in ["沒有符合條件的資料", "很抱歉", "查詢日期大於今日"]):
        return None

    table = extract_twse_table(payload.get("tables", []))
    if table is None:
        return None

    fields = [clean_text(x) for x in table.get("fields", [])]
    raw_df = pd.DataFrame(table.get("data", []), columns=fields)
    keep_cols = ["證券代號", "證券名稱", "成交股數", "開盤價", "最高價", "最低價", "收盤價"]
    if not all(c in raw_df.columns for c in keep_cols):
        return None

    out = raw_df[keep_cols].copy()
    out["股票代號"] = out["證券代號"].map(normalize_code)
    out["股票名稱"] = out["證券名稱"].map(clean_text)
    out["成交股數"] = out["成交股數"].map(lambda x: parse_number(x, is_int=True))
    out["開盤價"] = out["開盤價"].map(parse_number)
    out["最高價"] = out["最高價"].map(parse_number)
    out["最低價"] = out["最低價"].map(parse_number)
    out["收盤價"] = out["收盤價"].map(parse_number)
    out["日期"] = date_obj.strftime("%Y-%m-%d")
    out = out[out["股票代號"].map(is_valid_stock_code)].copy()
    out["市場別"] = "上市"
    return out[
        ["日期", "股票代號", "股票名稱", "開盤價", "最高價", "最低價", "收盤價", "成交股數", "市場別"]
    ].reset_index(drop=True)


def fetch_tpex_daily_all(date_obj):
    url = "https://www.tpex.org.tw/www/zh-tw/afterTrading/dailyQuotes"
    params = {"response": "json", "date": date_obj.strftime("%Y/%m/%d")}
    try:
        response = _http_get(url, params=params, timeout=15)
        payload = response.json()
    except Exception:
        return None

    rows = []
    tables = payload.get("tables") or payload.get("table") or []
    if isinstance(tables, dict):
        tables = [tables]
    for table in tables:
        data = table.get("data") or table.get("aaData") or []
        fields = [clean_text(x if not isinstance(x, dict) else x.get("title", "")) for x in (table.get("fields") or [])]
        for raw in data:
            if isinstance(raw, dict):
                code = normalize_code(raw.get("SecuritiesCompanyCode") or raw.get("代號") or raw.get("Code") or "")
                name = clean_text(raw.get("CompanyName") or raw.get("名稱") or raw.get("Name") or "")
                close_price = parse_number(raw.get("Close") or raw.get("收盤價"))
                open_price = parse_number(raw.get("Open") or raw.get("開盤價"))
                high_price = parse_number(raw.get("High") or raw.get("最高價"))
                low_price = parse_number(raw.get("Low") or raw.get("最低價"))
                volume_shares = parse_number(raw.get("TradingShares") or raw.get("成交股數"), is_int=True)
            else:
                if not fields:
                    continue
                row = {fields[i]: raw[i] if i < len(raw) else "" for i in range(len(fields))}
                code = normalize_code(row.get("代號") or row.get("證券代號") or "")
                name = clean_text(row.get("名稱") or row.get("證券名稱") or "")
                close_price = parse_number(row.get("收盤價"))
                open_price = parse_number(row.get("開盤價"))
                high_price = parse_number(row.get("最高價"))
                low_price = parse_number(row.get("最低價"))
                volume_shares = parse_number(row.get("成交股數"), is_int=True)

            if not is_valid_stock_code(code) or close_price is None:
                continue
            rows.append(
                {
                    "日期": date_obj.strftime("%Y-%m-%d"),
                    "股票代號": code,
                    "股票名稱": name,
                    "開盤價": open_price,
                    "最高價": high_price,
                    "最低價": low_price,
                    "收盤價": close_price,
                    "成交股數": int(volume_shares or 0),
                    "市場別": "上櫃",
                }
            )
    if not rows:
        return None
    return pd.DataFrame(rows).reset_index(drop=True)


def collect_daily_history(max_calendar_days=HISTORY_CALENDAR_DAYS):
    """並行抓取上市/上櫃日資料（策略明確，瓶頸在資料下載而非運算）。"""
    cursor = datetime.now()
    attempts = 0
    dates = []

    while attempts < max_calendar_days and len(dates) < MIN_TRADING_DAYS_TARGET:
        if cursor.weekday() < 5:
            dates.append(cursor.date())
        cursor -= timedelta(days=1)
        attempts += 1

    if not dates:
        return pd.DataFrame()

    def _fetch_one(day):
        day_obj = datetime.combine(day, datetime.min.time())
        parts = []
        twse = fetch_twse_daily_all(day_obj)
        if isinstance(twse, pd.DataFrame) and not twse.empty:
            parts.append(twse)
        tpex = fetch_tpex_daily_all(day_obj)
        if isinstance(tpex, pd.DataFrame) and not tpex.empty:
            parts.append(tpex)
        return parts

    frames = []
    workers = min(FETCH_WORKERS, max(1, len(dates)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(_fetch_one, day) for day in dates]
        for future in as_completed(futures):
            try:
                frames.extend(future.result())
            except Exception:
                continue

    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def _pick_col(cols, keywords):
    for i, col in enumerate(cols):
        if all(k in col for k in keywords):
            return col
    return None


def _flatten_json_records(obj):
    if isinstance(obj, list):
        return [x for x in obj if isinstance(x, dict)]
    if isinstance(obj, dict):
        for key in ["data", "records", "result", "items", "aaData"]:
            val = obj.get(key)
            if isinstance(val, list):
                return [x for x in val if isinstance(x, dict)]
    return []


def infer_issuer_from_warrant_name(name):
    name = clean_text(name)
    issuers = [
        "元大", "元富", "永豐金", "永豐", "統一", "國票", "凱基", "群益金鼎", "群益",
        "國泰", "富邦", "台新", "兆豐", "中國信託", "中信", "玉山", "第一", "華南",
    ]
    for issuer in sorted(issuers, key=len, reverse=True):
        if issuer in name:
            return issuer
    return ""


WARRANT_MARKET_URLS = [
    "https://openapi.twse.com.tw/v1/opendata/t187ap47_L",
    "https://www.tpex.org.tw/openapi/v1/t187ap47_O",
    "https://openapi.tpex.org.tw/v1/t187ap47_O",
]

_warrant_market_batches_cache = None


def fetch_warrant_market_data(force=False):
    """全市場權證資料只抓一次（3 來源並行），避免每檔標的重複下載。"""
    global _warrant_market_batches_cache
    if not force and _warrant_market_batches_cache is not None:
        return _warrant_market_batches_cache

    def _fetch_url(url):
        try:
            response = _http_get(url, timeout=25)
            return _flatten_json_records(response.json())
        except Exception:
            return []

    batches = []
    workers = min(3, len(WARRANT_MARKET_URLS))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        for records in pool.map(_fetch_url, WARRANT_MARKET_URLS):
            if records:
                batches.append(records)

    _warrant_market_batches_cache = batches
    return batches


def _warrants_from_record_batches(record_batches, target_code, stock_price=None):
    target_code = normalize_code(target_code)
    today = datetime.now().date()
    rows = []

    for records in record_batches:
        if not records:
            continue

        df = pd.DataFrame(records)
        if df.empty:
            continue
        df.columns = [clean_text(c) for c in df.columns]
        cols = list(df.columns)
        code_col = _pick_col(cols, ["權證", "代號"]) or _pick_col(cols, ["證券", "代號"])
        name_col = _pick_col(cols, ["權證", "名稱"]) or _pick_col(cols, ["證券", "名稱"])
        issuer_col = _pick_col(cols, ["發行"]) or _pick_col(cols, ["券商"])
        underlying_col = _pick_col(cols, ["標的", "代號"]) or _pick_col(cols, ["標的"])
        strike_col = _pick_col(cols, ["履約", "價"])
        expiry_col = _pick_col(cols, ["到期"])
        type_col = _pick_col(cols, ["認購"]) or _pick_col(cols, ["權證", "類型"])
        price_col = _pick_col(cols, ["收盤"]) or _pick_col(cols, ["成交"])
        ratio_col = _pick_col(cols, ["行使", "比例"])
        days_col = _pick_col(cols, ["剩餘", "天"])

        for _, record in df.iterrows():
            row_text = " ".join(clean_text(v) for v in record.astype(str).tolist())
            under = clean_text(record.get(underlying_col, "")) if underlying_col else ""
            if target_code not in under and target_code not in row_text:
                continue

            wcode = clean_text(record.get(code_col, "")) if code_col else ""
            wname = clean_text(record.get(name_col, "")) if name_col else ""
            if not wcode and not wname:
                continue

            expiry = parse_date_any(record.get(expiry_col, "")) if expiry_col else None
            days_left = (expiry - today).days if expiry else None
            if days_col:
                dval = parse_number(record.get(days_col, ""), is_int=True)
                if dval is not None:
                    days_left = dval

            if days_left is None or days_left < WARRANT_MIN_DAYS or days_left > WARRANT_MAX_DAYS:
                continue

            strike = parse_number(record.get(strike_col, "")) if strike_col else None
            raw_type = clean_text(record.get(type_col, "")) if type_col else wname
            if "售" in raw_type:
                wtype = "認售"
            elif "購" in raw_type:
                wtype = "認購"
            else:
                wtype = "認購/認售"

            issuer = clean_text(record.get(issuer_col, "")) if issuer_col else infer_issuer_from_warrant_name(wname)
            price_text = clean_text(record.get(price_col, "")) if price_col else ""
            ratio_text = clean_text(record.get(ratio_col, "")) if ratio_col else ""

            rows.append(
                {
                    "code": wcode,
                    "stock_id": wcode,
                    "name": wname,
                    "type": wtype,
                    "issuer": issuer,
                    "broker": issuer,
                    "stock_code": target_code,
                    "strike": strike if strike is not None else "",
                    "days_left": int(days_left),
                    "price": price_text,
                    "ratio": ratio_text,
                    "underlying_price": stock_price if stock_price is not None else "",
                }
            )

    if not rows:
        return []
    out_df = pd.DataFrame(rows).drop_duplicates(subset=["code", "name"], keep="first")
    return out_df.to_dict("records")


def fetch_warrants_for_stock(target_code, stock_price=None, market_batches=None):
    batches = market_batches if market_batches is not None else fetch_warrant_market_data()
    return _warrants_from_record_batches(batches, target_code, stock_price=stock_price)


def build_warrants_from_bullish(bullish_items):
    warrants = []
    seen = set()
    market_batches = fetch_warrant_market_data(force=True)
    for stock in bullish_items[:MAX_WARRANT_STOCKS]:
        code = str(stock.get("stock_id") or stock.get("code") or "")
        if not code:
            continue
        for item in fetch_warrants_for_stock(code, market_batches=market_batches):
            key = (item.get("code"), item.get("name"))
            if key in seen:
                continue
            seen.add(key)
            warrants.append(item)
            if len(warrants) >= MAX_RESULTS:
                return warrants
    return warrants


def run_web_strategy_analysis():
    daily_all = collect_daily_history()
    if daily_all.empty:
        return {
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "settle_date": "",
            "source": "web-strategy",
            "market": "上市+上櫃（週K策略）",
            "bullish": [],
            "bearish": [],
            "warrants": [],
            "bullish_count": 0,
            "bearish_count": 0,
            "warrant_count": 0,
            "strategy": {
                "min_score": DISPLAY_MIN_SCORE,
                "min_stars": DISPLAY_MIN_STARS,
                "min_weekly_volume": MIN_WEEKLY_VOLUME,
                "warrant_days": f"{WARRANT_MIN_DAYS}-{WARRANT_MAX_DAYS}",
            },
        }

    weekly_df = build_weekly_k_from_daily(daily_all)
    weekly_ma_df = calculate_weekly_indicators(weekly_df)
    master_df = build_master_df(daily_all)

    training_pool = build_training_pool(weekly_ma_df, master_df)
    bearish_pool = build_bearish_pool(weekly_ma_df, master_df)

    bullish_display = filter_bullish_for_display(training_pool)
    bullish_items = pool_to_api_items(bullish_display, score_col="StrongScore")
    bearish_items = pool_to_api_items(bearish_pool.head(MAX_RESULTS), score_col="BearishScore")

    warrant_pool = filter_warrant_candidates(training_pool)
    warrant_items = pool_to_api_items(warrant_pool, score_col="StrongScore")
    warrants = build_warrants_from_bullish(warrant_items)

    settle_date = ""
    if "日期" in daily_all.columns:
        try:
            settle_date = pd.to_datetime(daily_all["日期"]).max().strftime("%Y-%m-%d")
        except Exception:
            settle_date = ""

    return {
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "settle_date": settle_date,
        "source": "web-strategy-weekly",
        "market": "上市+上櫃（週K｜趨勢突破守穩｜週量≥1萬）",
        "bullish": bullish_items,
        "bearish": bearish_items,
        "warrants": warrants,
        "bullish_count": len(bullish_items),
        "bearish_count": len(bearish_items),
        "warrant_count": len(warrants),
        "pool_count": len(training_pool),
        "warrant_candidate_count": len(warrant_pool),
        "strategy": {
            "bullish": "週20MA+趨勢線突破守穩+週量≥1萬（同桌面看多）",
            "warrant_min_score": WARRANT_MIN_SCORE,
            "warrant_min_stars": WARRANT_MIN_STARS,
            "min_weekly_volume": MIN_WEEKLY_VOLUME,
            "bias_limit": None,
            "require_weekly_20ma_breakout": True,
            "warrant_days": f"{WARRANT_MIN_DAYS}-{WARRANT_MAX_DAYS}",
            "warrant_issuers": "all",
        },
    }
