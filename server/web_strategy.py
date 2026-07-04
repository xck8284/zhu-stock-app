# -*- coding: utf-8 -*-
"""
網頁版策略分析（對齊桌面版 zhustock_app 核心邏輯，不含記憶檔）。

使用者策略：
- 看多：TRAINING_POOL 再篩 StrongScore≥100（含上櫃補強後同條件）
- 看空：桌面版 BEARISH_TRAINING_POOL 邏輯
- 權證：StrongScore≥100、5星、剩餘 90~120 天、全部發行券商
"""

from __future__ import annotations

import json
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from io import StringIO

import pandas as pd
import requests

SESSION = requests.Session()
SESSION.headers.update(
    {
        "User-Agent": "Mozilla/5.0 (ZHU-STOCK-WEB/1.0)",
        "Accept": "application/json,text/plain,*/*",
    }
)


def _http_get(url, params=None, timeout=None):
    timeout = HTTP_TIMEOUT if timeout is None else timeout
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

DISPLAY_MIN_STARS = 5
USE_STABLE_COMPLETED_DAY = True
MARKET_FINAL_HOUR = 14
MARKET_FINAL_MINUTE = 10
WARRANT_MIN_DAYS = 90
WARRANT_MAX_DAYS = 120

MAX_WARRANT_RESULTS = 100
MAX_WARRANT_STOCKS = 12
# 對齊桌面版：history_end - 460 天，約 60 週（LOOKBACK_WEEKS=60）
HISTORY_CALENDAR_DAYS = 460
FETCH_WORKERS = 32
HTTP_TIMEOUT = 12

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


def roc_date_str(dt_obj):
    roc_year = dt_obj.year - 1911
    return f"{roc_year}/{dt_obj.month:02d}/{dt_obj.day:02d}"


def fmt_date_ymd(dt_obj):
    if isinstance(dt_obj, str):
        return dt_obj
    return dt_obj.strftime("%Y-%m-%d")


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


def get_effective_reference_today():
    now = datetime.now()
    if USE_STABLE_COMPLETED_DAY:
        cutoff = now.replace(hour=MARKET_FINAL_HOUR, minute=MARKET_FINAL_MINUTE, second=0, microsecond=0)
        if now < cutoff:
            return (now - timedelta(days=1)).date()
    return now.date()


def get_latest_available_trading_date(max_lookback_days=20):
    """對齊桌面版：14:10 前用昨日；回推找最近有行情的交易日。"""
    base_today = get_effective_reference_today()
    for i in range(max_lookback_days + 1):
        d = base_today - timedelta(days=i)
        if d.weekday() >= 5:
            continue
        try:
            df1 = fetch_twse_daily_all(d)
            if isinstance(df1, pd.DataFrame) and not df1.empty:
                return d
        except Exception:
            pass
        try:
            df2 = fetch_tpex_daily_all(d)
            if isinstance(df2, pd.DataFrame) and not df2.empty:
                return d
        except Exception:
            pass
    for i in range(max_lookback_days + 1):
        d = base_today - timedelta(days=i)
        if d.weekday() < 5:
            return d
    return base_today


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
                    "descent_pct": descent_pct,
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
                    "ascent_pct": ascent_pct,
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

    if pd.notna(low26) and low26 > 0:
        d26 = safe_pct(close_, low26)
        if d26 is not None:
            if d26 <= 3:
                score += 8
                tags.append("近26週低")
            elif d26 <= 8:
                score += 4
                tags.append("接近26週低")

    if pd.notna(low52) and low52 > 0:
        d52 = safe_pct(close_, low52)
        if d52 is not None:
            if d52 <= 5:
                score += 8
                tags.append("接近52週低")
            elif d52 <= 12:
                score += 4
                tags.append("中長期弱勢")

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
                "是否最新正式突破": "是" if trend.get("strict_ok") else "否",
                "是否突破後守穩趨勢線": "是",
                **extra,
            }
        )

    if not rows:
        return pd.DataFrame()
    out = pd.DataFrame(rows)
    out = out.sort_values(
        ["StrongScore", "星等數值", "是否最新正式突破", "最新一週成交量(張)", "股票代號"],
        ascending=[False, False, False, False, True],
    )
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
    """同桌面版 CLIENT_BULLISH：TRAINING_POOL 全量 + mix 排序（不篩 ≥100）。"""
    return build_client_bullish_view(pool_df)


def calc_display_star_by_score(score):
    try:
        score = float(score)
    except Exception:
        score = 0
    if score >= 130:
        return "★★★★★", 5
    if score >= 105:
        return "★★★★☆", 4
    if score >= 80:
        return "★★★☆☆", 3
    if score >= 60:
        return "★★☆☆☆", 2
    return "★☆☆☆☆", 1


def build_otc_bullish_supplement_pool(weekly_ma_df, master_df, existing_df=None, min_count=40):
    """上櫃週K補強（對齊桌面版 build_otc_bullish_supplement_pool）。"""
    cols = [
        "股票代號", "股票名稱", "市場別", "產業別", "週結算日期",
        "StrongScore", "星等", "星等數值", "乖離率(%)",
        "短線停利Alarm", "長線停利Alarm", "記憶回饋", "最新一週成交量(張)",
    ]
    if weekly_ma_df is None or weekly_ma_df.empty:
        return pd.DataFrame(columns=cols)

    exist_codes = set()
    if isinstance(existing_df, pd.DataFrame) and not existing_df.empty and "股票代號" in existing_df.columns:
        exist_codes = set(existing_df["股票代號"].astype(str))

    df = weekly_ma_df.copy()
    df["週結算日期"] = pd.to_datetime(df["週結算日期"], errors="coerce")
    latest = df.sort_values(["股票代號", "週結算日期"]).groupby("股票代號", as_index=False).tail(1).copy()
    latest = latest[latest.get("市場別", "").astype(str).eq("上櫃")].copy()
    latest = latest[~latest["股票代號"].astype(str).isin(exist_codes)].copy()
    latest = latest.dropna(subset=["週收盤價", "週20MA", "週5MA"]).copy()
    if latest.empty:
        return pd.DataFrame(columns=cols)

    for c in [
        "週收盤價", "週3MA", "週5MA", "週10MA", "週20MA", "週成交量(張)", "量20MA",
        "20MA斜率", "5MA斜率", "近13週最高", "近26週最高", "近52週最高",
    ]:
        if c in latest.columns:
            latest[c] = pd.to_numeric(latest[c], errors="coerce")

    vol_min = max(3000, int(MIN_WEEKLY_VOLUME * 0.5))
    latest = latest[(latest["週收盤價"] >= latest["週20MA"]) & (latest["週成交量(張)"] >= vol_min)].copy()
    if latest.empty:
        return pd.DataFrame(columns=cols)

    def _score(r):
        score = 45
        tags = []
        try:
            close = float(r.get("週收盤價", 0))
            ma3 = float(r.get("週3MA", 0)) if pd.notna(r.get("週3MA")) else None
            ma5 = float(r.get("週5MA", 0)) if pd.notna(r.get("週5MA")) else None
            ma10 = float(r.get("週10MA", 0)) if pd.notna(r.get("週10MA")) else None
            ma20 = float(r.get("週20MA", 0)) if pd.notna(r.get("週20MA")) else None
            vol = float(r.get("週成交量(張)", 0)) if pd.notna(r.get("週成交量(張)")) else 0
            vol20 = float(r.get("量20MA", 0)) if pd.notna(r.get("量20MA")) else 0
            if ma20 and close >= ma20:
                score += 15
                tags.append("站上週20MA")
            if ma3 and ma5 and ma10 and ma20 and ma3 >= ma5 >= ma10 >= ma20:
                score += 25
                tags.append("均線多頭")
            if pd.notna(r.get("5MA斜率")) and float(r.get("5MA斜率")) > 0:
                score += 10
                tags.append("5MA上揚")
            if pd.notna(r.get("20MA斜率")) and float(r.get("20MA斜率")) > 0:
                score += 10
                tags.append("20MA上揚")
            if vol20 and vol >= vol20:
                score += 15
                tags.append("量能大於20週均量")
            if vol >= MIN_WEEKLY_VOLUME:
                score += 10
                tags.append("週量達標")
            hi13 = r.get("近13週最高")
            if pd.notna(hi13) and hi13 and close >= float(hi13) * 0.92:
                score += 10
                tags.append("接近13週高")
        except Exception:
            pass
        return int(score), "、".join(tags)

    scored = latest.apply(lambda r: _score(r), axis=1)
    latest["StrongScore"] = [x[0] for x in scored]
    latest = latest[latest["StrongScore"] >= 60].copy()
    if latest.empty:
        return pd.DataFrame(columns=cols)

    industry_map = (
        master_df[["股票代號", "產業別"]].drop_duplicates(subset=["股票代號"]).copy()
        if isinstance(master_df, pd.DataFrame) and "產業別" in master_df.columns
        else pd.DataFrame(columns=["股票代號", "產業別"])
    )
    latest = latest.drop(columns=["產業別"], errors="ignore").merge(industry_map, on="股票代號", how="left")
    latest["產業別"] = latest["產業別"].fillna("未分類")
    latest = latest.sort_values(["StrongScore", "週成交量(張)", "股票代號"], ascending=[False, False, True]).head(int(min_count)).copy()

    star_pairs = latest["StrongScore"].map(calc_display_star_by_score)
    latest["星等"] = star_pairs.map(lambda x: x[0])
    latest["星等數值"] = star_pairs.map(lambda x: x[1])
    latest["乖離率(%)"] = ((latest["週收盤價"] - latest["週20MA"]) / latest["週20MA"] * 100).round(2)

    return pd.DataFrame(
        {
            "股票代號": latest["股票代號"].astype(str),
            "股票名稱": latest["股票名稱"].astype(str),
            "市場別": "上櫃",
            "產業別": latest["產業別"].astype(str),
            "週結算日期": latest["週結算日期"].dt.strftime("%Y-%m-%d"),
            "StrongScore": latest["StrongScore"].astype(int),
            "星等": latest["星等"],
            "星等數值": latest["星等數值"],
            "乖離率(%)": latest["乖離率(%)"],
            "短線停利Alarm": "否",
            "長線停利Alarm": "否",
            "記憶回饋": "上櫃週K補強分析",
            "最新一週成交量(張)": latest["週成交量(張)"].astype(int),
        }
    )


def mix_market_rows_for_display(out, min_otc_visible=25):
    """對齊桌面版：確保上櫃標的不被上市完全擠到後面。"""
    if not isinstance(out, pd.DataFrame) or out.empty or "市場別" not in out.columns:
        return out
    work = out.copy()
    if "StrongScore" in work.columns:
        work["_score_"] = pd.to_numeric(work["StrongScore"], errors="coerce").fillna(0)
    else:
        work["_score_"] = 0
    work["_star_"] = work["星等"].astype(str).map(lambda x: x.count("★")) if "星等" in work.columns else 0
    listed = work[work["市場別"].astype(str).eq("上市")].sort_values(
        ["_score_", "_star_", "股票代號"], ascending=[False, False, True], kind="mergesort"
    )
    otc = work[work["市場別"].astype(str).eq("上櫃")].sort_values(
        ["_score_", "_star_", "股票代號"], ascending=[False, False, True], kind="mergesort"
    )
    other = work[~work["市場別"].astype(str).isin(["上市", "上櫃"])].sort_values(
        ["_score_", "_star_", "股票代號"], ascending=[False, False, True], kind="mergesort"
    )
    if otc.empty:
        mixed = pd.concat([listed, other], ignore_index=True)
    else:
        head_listed = listed.head(15)
        head_otc = otc.head(min_otc_visible)
        remain = pd.concat([listed.iloc[15:], otc.iloc[min_otc_visible:], other], ignore_index=True)
        remain = remain.sort_values(["_score_", "_star_", "股票代號"], ascending=[False, False, True], kind="mergesort")
        mixed = pd.concat([head_listed, head_otc, remain], ignore_index=True)
    return mixed.drop(columns=["_score_", "_star_"], errors="ignore").reset_index(drop=True)


def build_client_bullish_view(training_df):
    cols = [
        "股票代號", "股票名稱", "市場別", "產業別", "週結算日期",
        "星等", "StrongScore", "乖離率(%)", "短線停利Alarm", "長線停利Alarm",
    ]
    if training_df is None or training_df.empty:
        return pd.DataFrame(columns=cols)
    use_cols = [c for c in cols if c in training_df.columns]
    out = training_df[use_cols].copy()
    out = out.drop_duplicates(subset=["股票代號"]).reset_index(drop=True)
    return mix_market_rows_for_display(out, min_otc_visible=25)


def build_client_bearish_view(bearish_df):
    cols = [
        "股票代號", "股票名稱", "市場別", "產業別", "週結算日期",
        "星等", "乖離率(%)", "短線回補Alarm", "長線回補Alarm",
    ]
    if bearish_df is None or bearish_df.empty:
        return pd.DataFrame(columns=cols)
    use_cols = [c for c in cols if c in bearish_df.columns]
    out = bearish_df[use_cols].copy()
    out = out.drop_duplicates(subset=["股票代號"]).reset_index(drop=True)
    out["_star_"] = out["星等"].astype(str).map(lambda x: x.count("★")) if "星等" in out.columns else 0
    out = out.sort_values(["_star_", "股票代號"], ascending=[False, True], kind="mergesort")
    return out.drop(columns=["_star_"], errors="ignore").reset_index(drop=True)


def build_strict_breakout_sheet(weekly_ma_df, master_df):
    """STRICT_BREAKOUT → 桌面版多方關鍵K 來源。"""
    if weekly_ma_df.empty:
        return pd.DataFrame()
    df = weekly_ma_df.copy()
    df["週結算日期"] = pd.to_datetime(df["週結算日期"])
    df = df.sort_values(["股票代號", "週結算日期"]).reset_index(drop=True)
    industry_map = master_df[["股票代號", "產業別"]].drop_duplicates(subset=["股票代號"]).copy()
    rows = []
    for code, grp in df.groupby("股票代號"):
        grp = grp.sort_values("週結算日期").reset_index(drop=True).copy()
        latest = grp.iloc[-1]
        if pd.isna(latest["週20MA"]) or latest["週收盤價"] < latest["週20MA"]:
            continue
        if pd.isna(latest["週成交量(張)"]) or latest["週成交量(張)"] < MIN_WEEKLY_VOLUME:
            continue
        trend = analyze_best_descending_trendline(grp)
        if trend is None or not trend.get("strict_ok"):
            continue
        work = trend["work_df"]
        i, j = trend["i"], trend["j"]
        industry = industry_map.loc[industry_map["股票代號"] == code, "產業別"]
        industry_val = industry.iloc[0] if len(industry) > 0 else "未分類"
        rows.append(
            {
                "股票代號": str(code),
                "股票名稱": str(latest["股票名稱"]),
                "市場別": str(latest["市場別"]),
                "產業別": str(industry_val),
                "週結算日期": latest["週結算日期"].strftime("%Y-%m-%d"),
                "最新週收盤價": round(float(latest["週收盤價"]), 4),
                "週20MA": round(float(latest["週20MA"]), 4),
                "最新一週成交量(張)": int(latest["週成交量(張)"]),
                "趨勢線距離(%)": round(float(trend["line_distance_pct"]), 2) if trend["line_distance_pct"] is not None else None,
                "盤整區距離(%)": round(float(trend["box_distance_pct"]), 2) if trend["box_distance_pct"] is not None else None,
            }
        )
    if not rows:
        return pd.DataFrame()
    out = pd.DataFrame(rows)
    return out.sort_values(
        ["盤整區距離(%)", "趨勢線距離(%)", "最新一週成交量(張)", "股票代號"],
        ascending=[False, False, False, True],
    ).reset_index(drop=True)


def build_bearish_key_breakdown_sheet(weekly_ma_df, master_df):
    """BEARISH_KEY_BREAKDOWN → 桌面版空方關鍵K 來源。"""
    if weekly_ma_df.empty:
        return pd.DataFrame()
    df = weekly_ma_df.copy()
    df["週結算日期"] = pd.to_datetime(df["週結算日期"])
    df = df.sort_values(["股票代號", "週結算日期"]).reset_index(drop=True)
    industry_map = master_df[["股票代號", "產業別"]].drop_duplicates(subset=["股票代號"]).copy()
    rows = []
    for code, grp in df.groupby("股票代號"):
        grp = grp.sort_values("週結算日期").reset_index(drop=True).copy()
        latest = grp.iloc[-1]
        if pd.isna(latest["週20MA"]) or latest["週收盤價"] >= latest["週20MA"]:
            continue
        if pd.isna(latest["週成交量(張)"]) or latest["週成交量(張)"] < MIN_WEEKLY_VOLUME:
            continue
        trend = analyze_best_ascending_trendline(grp)
        if trend is None or not (trend.get("line_break_now") and not trend.get("line_break_prev")):
            continue
        industry = industry_map.loc[industry_map["股票代號"] == code, "產業別"]
        industry_val = industry.iloc[0] if len(industry) > 0 else "未分類"
        rows.append(
            {
                "股票代號": str(code),
                "股票名稱": str(latest["股票名稱"]),
                "市場別": str(latest["市場別"]),
                "產業別": str(industry_val),
                "週結算日期": latest["週結算日期"].strftime("%Y-%m-%d"),
                "最新週收盤價": round(float(latest["週收盤價"]), 4),
                "週20MA": round(float(latest["週20MA"]), 4),
                "最新一週成交量(張)": int(latest["週成交量(張)"]),
                "趨勢線距離(%)": round(float(trend["line_distance_pct"]), 2) if trend["line_distance_pct"] is not None else None,
            }
        )
    if not rows:
        return pd.DataFrame()
    out = pd.DataFrame(rows)
    return out.sort_values(["趨勢線距離(%)", "最新一週成交量(張)", "股票代號"], ascending=[False, False, True]).reset_index(drop=True)


def build_client_bullish_keyk_view(strict_df):
    cols = [
        "股票代號", "股票名稱", "市場別", "產業別", "週結算日期",
        "最新週收盤價", "週20MA", "最新一週成交量(張)", "趨勢線距離(%)", "盤整區距離(%)",
    ]
    if strict_df is None or strict_df.empty:
        return pd.DataFrame(columns=cols)
    use_cols = [c for c in cols if c in strict_df.columns]
    out = strict_df[use_cols].copy().drop_duplicates(subset=["股票代號"]).reset_index(drop=True)
    sort_cols = [c for c in ["盤整區距離(%)", "趨勢線距離(%)", "最新一週成交量(張)", "股票代號"] if c in out.columns]
    asc = [False if c != "股票代號" else True for c in sort_cols]
    if sort_cols:
        out = out.sort_values(sort_cols, ascending=asc).reset_index(drop=True)
    return out


def build_client_bearish_keyk_view(bearish_key_df):
    cols = [
        "股票代號", "股票名稱", "市場別", "產業別", "週結算日期",
        "最新週收盤價", "週20MA", "最新一週成交量(張)", "趨勢線距離(%)",
    ]
    if bearish_key_df is None or bearish_key_df.empty:
        return pd.DataFrame(columns=cols)
    use_cols = [c for c in cols if c in bearish_key_df.columns]
    out = bearish_key_df[use_cols].copy().drop_duplicates(subset=["股票代號"]).reset_index(drop=True)
    sort_cols = [c for c in ["趨勢線距離(%)", "最新一週成交量(張)", "股票代號"] if c in out.columns]
    asc = [False if c != "股票代號" else True for c in sort_cols]
    if sort_cols:
        out = out.sort_values(sort_cols, ascending=asc).reset_index(drop=True)
    return out


def filter_bearish_for_display(pool_df):
    """看空清單：對齊桌面版 CLIENT_BEARISH。"""
    return build_client_bearish_view(pool_df)


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
        score_val = row.get(score_col)
        if score_val is None or (isinstance(score_val, float) and pd.isna(score_val)):
            score_val = 0
        item = {
                "stock_id": str(row["股票代號"]),
                "code": str(row["股票代號"]),
                "name": str(row.get("股票名稱", "")),
                "industry": str(row.get("產業別", "")),
                "market": str(row.get("市場別", "")),
                "direction": "看多" if score_col == "StrongScore" else "看空",
                "stars": str(row.get("星等", "")),
                "bias": bias_text,
                "settle_date": str(row.get("週結算日期", "")),
                "short_alarm": str(row.get("短線停利Alarm") or row.get("短線回補Alarm") or ""),
                "long_alarm": str(row.get("長線停利Alarm") or row.get("長線回補Alarm") or ""),
                "memory_note": str(row.get("記憶回饋", "") or ""),
            }
        if score_col == "BearishScore":
            item["bearish_score"] = float(score_val)
            item["strong_score"] = float(score_val)
        else:
            item["strong_score"] = float(score_val)
        items.append(item)
    return items


def keyk_to_api_items(keyk_df):
    items = []
    for _, row in keyk_df.iterrows():
        items.append(
            {
                "stock_id": str(row["股票代號"]),
                "code": str(row["股票代號"]),
                "name": str(row.get("股票名稱", "")),
                "industry": str(row.get("產業別", "")),
                "market": str(row.get("市場別", "")),
                "settle_date": str(row.get("週結算日期", "")),
                "close": row.get("最新週收盤價", ""),
                "ma20": row.get("週20MA", ""),
                "volume_lots": row.get("最新一週成交量(張)", ""),
                "line_distance_pct": row.get("趨勢線距離(%)", ""),
                "box_distance_pct": row.get("盤整區距離(%)", ""),
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
    """抓取上櫃 TPEx 每日收盤（對齊桌面版 v5：JSON + 舊版 + HTML 多重來源）。"""

    def _extract_rows_and_fields(obj):
        if not isinstance(obj, dict):
            return [], []
        for key in ["tables", "table"]:
            if isinstance(obj.get(key), list):
                for table in obj.get(key):
                    if isinstance(table, dict):
                        rows = table.get("data") or table.get("aaData") or table.get("rows") or []
                        fields = table.get("fields") or table.get("columns") or table.get("titles") or []
                        if rows:
                            return rows, fields
        rows = obj.get("data") or obj.get("aaData") or obj.get("rows") or []
        fields = obj.get("fields") or obj.get("columns") or obj.get("titles") or []
        return rows or [], fields or []

    def _field_text(value):
        if isinstance(value, dict):
            return clean_text(value.get("title") or value.get("name") or value.get("key") or value.get("data") or "")
        return clean_text(value)

    def _parse_rows(rows, fields=None):
        fields = [_field_text(x) for x in (fields or [])]
        parsed = []

        def idx_by_keywords(groups, default=None):
            for kws in groups:
                for i, field in enumerate(fields):
                    if all(k in field for k in kws):
                        return i
            return default

        code_i = idx_by_keywords([["代號"], ["證券", "代號"]], 0)
        name_i = idx_by_keywords([["名稱"], ["證券", "名稱"]], 1)
        close_i = idx_by_keywords([["收盤"]], 2)
        open_i = idx_by_keywords([["開盤"]], 4)
        high_i = idx_by_keywords([["最高"]], 5)
        low_i = idx_by_keywords([["最低"]], 6)
        vol_i = idx_by_keywords([["成交", "股"], ["成交股數"], ["成交", "數量"]], 8)

        for row in rows:
            if isinstance(row, dict):
                code = normalize_code(
                    row.get("代號") or row.get("證券代號") or row.get("股票代號") or row.get("Code") or ""
                )
                name = clean_text(row.get("名稱") or row.get("證券名稱") or row.get("股票名稱") or row.get("Name") or "")
                close_price = parse_number(row.get("收盤") or row.get("收盤價") or row.get("Close"))
                open_price = parse_number(row.get("開盤") or row.get("開盤價") or row.get("Open"))
                high_price = parse_number(row.get("最高") or row.get("最高價") or row.get("High"))
                low_price = parse_number(row.get("最低") or row.get("最低價") or row.get("Low"))
                volume_shares = parse_number(
                    row.get("成交股數") or row.get("成交股數(股)") or row.get("成交數量") or row.get("TradeVolume"),
                    is_int=True,
                )
            else:
                if not isinstance(row, (list, tuple)) or len(row) < 7:
                    continue

                def gv(index):
                    return row[index] if index is not None and index < len(row) else ""

                code = normalize_code(gv(code_i))
                name = clean_text(gv(name_i))
                close_price = parse_number(gv(close_i))
                open_price = parse_number(gv(open_i))
                high_price = parse_number(gv(high_i))
                low_price = parse_number(gv(low_i))
                volume_shares = parse_number(gv(vol_i), is_int=True)
                if volume_shares is None:
                    for cand_i in [8, 9, 10, 11, 7]:
                        candidate = parse_number(gv(cand_i), is_int=True)
                        if candidate is not None and candidate >= 1000:
                            volume_shares = candidate
                            break

            if not is_valid_stock_code(code) or close_price is None:
                continue
            parsed.append(
                {
                    "日期": fmt_date_ymd(date_obj),
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
        return pd.DataFrame(parsed) if parsed else None

    ad_date_slash = date_obj.strftime("%Y/%m/%d")
    roc = roc_date_str(date_obj)

    json_sources = [
        ("https://www.tpex.org.tw/www/zh-tw/afterTrading/dailyQuotes", {"response": "json", "date": ad_date_slash}),
        (
            "https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php",
            {"l": "zh-tw", "d": roc, "o": "json", "s": "0,asc,0"},
        ),
    ]

    for url, params in json_sources:
        try:
            response = _http_get(url, params=params, timeout=12)
            obj = response.json()
            rows, fields = _extract_rows_and_fields(obj)
            df = _parse_rows(rows, fields)
            if isinstance(df, pd.DataFrame) and not df.empty:
                return df.reset_index(drop=True)
        except Exception:
            continue

    html_sources = [
        ("https://www.tpex.org.tw/www/zh-tw/afterTrading/dailyQuotes", {"response": "html", "date": ad_date_slash}),
        (
            "https://www.tpex.org.tw/web/stock/aftertrading/otc_quotes_no1430/stk_wn1430_result.php",
            {"l": "zh-tw", "o": "htm", "d": roc, "se": "AL", "s": "0,asc,0"},
        ),
    ]
    for url, params in html_sources:
        try:
            response = _http_get(url, params=params, timeout=12)
            tables = pd.read_html(StringIO(response.text))
            for tbl in tables:
                if isinstance(tbl.columns, pd.MultiIndex):
                    tbl.columns = tbl.columns.get_level_values(-1)
                tbl.columns = [clean_text(c) for c in tbl.columns]
                cols = list(tbl.columns)
                code_col = _pick_col(cols, ["代號"]) or _pick_col(cols, ["證券", "代號"])
                name_col = _pick_col(cols, ["名稱"]) or _pick_col(cols, ["證券", "名稱"])
                close_col = _pick_col(cols, ["收盤"])
                open_col = _pick_col(cols, ["開盤"])
                high_col = _pick_col(cols, ["最高"])
                low_col = _pick_col(cols, ["最低"])
                vol_col = _pick_col(cols, ["成交", "股"]) or _pick_col(cols, ["成交股數"])
                if not all([code_col, name_col, close_col, open_col, high_col, low_col]):
                    continue
                html_rows = []
                for _, rr in tbl.iterrows():
                    code = normalize_code(rr.get(code_col, ""))
                    if not is_valid_stock_code(code):
                        continue
                    html_rows.append(
                        {
                            "日期": fmt_date_ymd(date_obj),
                            "股票代號": code,
                            "股票名稱": clean_text(rr.get(name_col, "")),
                            "開盤價": parse_number(rr.get(open_col)),
                            "最高價": parse_number(rr.get(high_col)),
                            "最低價": parse_number(rr.get(low_col)),
                            "收盤價": parse_number(rr.get(close_col)),
                            "成交股數": int(parse_number(rr.get(vol_col), is_int=True) or 0) if vol_col else 0,
                            "市場別": "上櫃",
                        }
                    )
                if html_rows:
                    return pd.DataFrame(html_rows).reset_index(drop=True)
        except Exception:
            continue

    return None


def collect_daily_history(history_calendar_days=HISTORY_CALENDAR_DAYS, progress_callback=None):
    """並行抓取上市/上櫃日資料（對齊桌面版 ~460 天）；有 DB 快取後日常只補最新 1～2 天。"""
    from web_daily_cache import fetch_market_day_cached

    end_date = get_latest_available_trading_date()
    start_date = end_date - timedelta(days=history_calendar_days)

    dates = []
    cur = start_date
    while cur <= end_date:
        if cur.weekday() < 5:
            dates.append(cur)
        cur += timedelta(days=1)

    if not dates:
        return pd.DataFrame()

    listed_tasks = [(d, "上市", fetch_twse_daily_all) for d in dates]
    otc_tasks = [(d, "上櫃", fetch_tpex_daily_all) for d in dates]
    total_tasks = len(listed_tasks) + len(otc_tasks)
    done = 0
    lock = threading.Lock()

    def _fetch_one(day, market, fetch_fn):
        nonlocal done
        trade_date = day.strftime("%Y-%m-%d")
        df = None
        try:
            df = fetch_market_day_cached(trade_date, market, fetch_fn)
        except Exception:
            df = None
        with lock:
            done += 1
            if progress_callback and (done % 15 == 0 or done == total_tasks):
                progress_callback(done, total_tasks)
        return df if isinstance(df, pd.DataFrame) and not df.empty else None

    frames = []

    def _run_market_tasks(tasks, workers):
        batch_frames = []
        with ThreadPoolExecutor(max_workers=workers) as pool:
            futures = [pool.submit(_fetch_one, d, market, fn) for d, market, fn in tasks]
            for future in as_completed(futures):
                try:
                    df = future.result()
                    if isinstance(df, pd.DataFrame) and not df.empty:
                        batch_frames.append(df)
                except Exception:
                    continue
        return batch_frames

    from web_daily_cache import history_fetch_workers

    listed_workers, otc_workers = history_fetch_workers()
    with ThreadPoolExecutor(max_workers=2) as outer:
        listed_future = outer.submit(_run_market_tasks, listed_tasks, listed_workers)
        otc_future = outer.submit(_run_market_tasks, otc_tasks, otc_workers)
        frames.extend(listed_future.result())
        frames.extend(otc_future.result())

    if not frames:
        return pd.DataFrame()

    out = pd.concat(frames, ignore_index=True)
    out["日期"] = pd.to_datetime(out["日期"], errors="coerce")
    out = out.sort_values(["日期", "股票代號"]).reset_index(drop=True)
    return out


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
    """全市場權證資料只抓一次（多來源並行），避免每檔標的重複下載。"""
    global _warrant_market_batches_cache
    if not force and _warrant_market_batches_cache is not None:
        return _warrant_market_batches_cache

    def _fetch_url(url):
        try:
            response = _http_get(url, timeout=20)
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


def _warrant_column_map(cols):
    return {
        "code": _pick_col(cols, ["權證", "代號"]) or _pick_col(cols, ["證券", "代號"]),
        "name": _pick_col(cols, ["權證", "名稱"]) or _pick_col(cols, ["權證", "簡稱"]) or _pick_col(cols, ["證券", "名稱"]),
        "issuer": (
            _pick_col(cols, ["發行", "人"])
            or _pick_col(cols, ["發行", "商"])
            or _pick_col(cols, ["發行"])
            or _pick_col(cols, ["券商"])
        ),
        "underlying": (
            _pick_col(cols, ["標的", "證券", "代號"])
            or _pick_col(cols, ["標的", "代號"])
            or _pick_col(cols, ["標的", "股", "代號"])
            or _pick_col(cols, ["連結", "代號"])
            or _pick_col(cols, ["標的"])
        ),
        "strike": _pick_col(cols, ["履約", "價"]) or _pick_col(cols, ["行使", "價"]),
        "expiry": _pick_col(cols, ["到期", "日期"]) or _pick_col(cols, ["到期日"]) or _pick_col(cols, ["到期"]),
        "type": _pick_col(cols, ["認購"]) or _pick_col(cols, ["權證", "類型"]) or _pick_col(cols, ["權證", "種類"]),
        "price": _pick_col(cols, ["收盤", "價"]) or _pick_col(cols, ["成交", "價"]) or _pick_col(cols, ["收盤"]),
        "ratio": _pick_col(cols, ["行使", "比例"]) or _pick_col(cols, ["履約", "比例"]),
        "days": _pick_col(cols, ["剩餘", "天"]) or _pick_col(cols, ["距到期", "天"]),
    }


def _records_to_warrant_rows(records, target_codes):
    """向量化篩選權證，只掃描標的代號命中列。"""
    if not records or not target_codes:
        return {}

    df = pd.DataFrame(records)
    if df.empty:
        return {}

    df.columns = [clean_text(c) for c in df.columns]
    cols = _warrant_column_map(list(df.columns))
    today = datetime.now().date()
    lookup = {code: [] for code in target_codes}

    if cols["underlying"] and cols["underlying"] in df.columns:
        df["_under"] = df[cols["underlying"]].astype(str).map(normalize_code)
    else:
        df["_under"] = ""

    df = df[df["_under"].isin(target_codes)].copy()
    if df.empty:
        return lookup

    if cols["expiry"] and cols["expiry"] in df.columns:
        df["_expiry"] = df[cols["expiry"]].map(parse_date_any)
    else:
        df["_expiry"] = None

    if cols["days"] and cols["days"] in df.columns:
        df["_days"] = pd.to_numeric(df[cols["days"]], errors="coerce")
    else:
        df["_days"] = pd.NA

    for idx, record in df.iterrows():
        stock_code = record["_under"]
        if stock_code not in lookup:
            continue

        days_left = record.get("_days")
        if pd.isna(days_left):
            expiry = record.get("_expiry")
            days_left = (expiry - today).days if expiry else None
        else:
            days_left = int(days_left)

        if days_left is None or days_left < WARRANT_MIN_DAYS or days_left > WARRANT_MAX_DAYS:
            continue

        wcode = clean_text(record.get(cols["code"], "")) if cols["code"] else ""
        wname = clean_text(record.get(cols["name"], "")) if cols["name"] else ""
        if not wcode and not wname:
            continue

        raw_type = clean_text(record.get(cols["type"], "")) if cols["type"] else wname
        if "售" in raw_type:
            wtype = "認售"
        elif "購" in raw_type:
            wtype = "認購"
        else:
            wtype = "認購/認售"

        issuer = clean_text(record.get(cols["issuer"], "")) if cols["issuer"] else infer_issuer_from_warrant_name(wname)
        strike = parse_number(record.get(cols["strike"], "")) if cols["strike"] else None
        price_text = clean_text(record.get(cols["price"], "")) if cols["price"] else ""
        ratio_text = clean_text(record.get(cols["ratio"], "")) if cols["ratio"] else ""

        lookup[stock_code].append(
            {
                "code": wcode,
                "stock_id": wcode,
                "name": wname,
                "type": wtype,
                "issuer": issuer,
                "broker": issuer,
                "stock_code": stock_code,
                "strike": strike if strike is not None else "",
                "days_left": int(days_left),
                "price": price_text,
                "ratio": ratio_text,
                "underlying_price": "",
            }
        )

    for code, rows in lookup.items():
        if not rows:
            continue
        deduped = pd.DataFrame(rows).drop_duplicates(subset=["code", "name"], keep="first")
        lookup[code] = deduped.to_dict("records")
    return lookup


def _warrants_from_record_batches(record_batches, target_code, stock_price=None):
    target_code = normalize_code(target_code)
    lookup = _records_to_warrant_rows(
        [row for batch in record_batches for row in (batch or [])],
        {target_code},
    )
    rows = lookup.get(target_code, [])
    if stock_price is not None:
        for item in rows:
            item["underlying_price"] = stock_price
    return rows


def _build_warrant_lookup(record_batches, target_codes):
    """一次掃描全市場權證資料，依標的代號分組（向量化，避免全表 iterrows）。"""
    target_codes = {normalize_code(c) for c in target_codes if c}
    lookup = {code: [] for code in target_codes}
    if not target_codes or not record_batches:
        return lookup

    merged = {}
    for records in record_batches:
        if not records:
            continue
        batch_lookup = _records_to_warrant_rows(records, target_codes)
        for code, rows in batch_lookup.items():
            if not rows:
                continue
            merged.setdefault(code, []).extend(rows)

    for code in target_codes:
        rows = merged.get(code) or []
        if not rows:
            continue
        deduped = pd.DataFrame(rows).drop_duplicates(subset=["code", "name"], keep="first")
        lookup[code] = deduped.to_dict("records")
    return lookup


def fetch_warrants_for_stock(target_code, stock_price=None, market_batches=None):
    batches = market_batches if market_batches is not None else fetch_warrant_market_data()
    return _warrants_from_record_batches(batches, target_code, stock_price=stock_price)


def build_warrants_from_bullish(bullish_items):
    warrants = []
    seen = set()
    stock_codes = []
    for stock in bullish_items[:MAX_WARRANT_STOCKS]:
        code = str(stock.get("stock_id") or stock.get("code") or "")
        if code:
            stock_codes.append(code)
    if not stock_codes:
        return warrants

    market_batches = fetch_warrant_market_data(force=True)
    lookup = _build_warrant_lookup(market_batches, stock_codes)
    for code in stock_codes:
        for item in lookup.get(normalize_code(code), []):
            key = (item.get("code"), item.get("name"))
            if key in seen:
                continue
            seen.add(key)
            warrants.append(item)
            if len(warrants) >= MAX_WARRANT_RESULTS:
                return warrants
    return warrants


def run_web_strategy_analysis(include_warrants=True, progress_callback=None):
    daily_all = collect_daily_history(progress_callback=progress_callback)
    if daily_all.empty:
        return {
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "settle_date": "",
            "source": "web-strategy",
            "market": "上市+上櫃（週K策略）",
            "bullish": [],
            "bearish": [],
            "bullish_keyk": [],
            "bearish_keyk": [],
            "warrants": [],
            "bullish_count": 0,
            "bearish_count": 0,
            "bullish_keyk_count": 0,
            "bearish_keyk_count": 0,
            "warrant_count": 0,
            "strategy": {
                "bullish_rule": "TRAINING_POOL：週量≥1萬＋趨勢突破守穩＋StrongScore≥55（同桌面版 CLIENT_BULLISH）",
                "min_weekly_volume": MIN_WEEKLY_VOLUME,
                "warrant_days": f"{WARRANT_MIN_DAYS}-{WARRANT_MAX_DAYS}",
            },
        }

    weekly_df = build_weekly_k_from_daily(daily_all)
    weekly_ma_df = calculate_weekly_indicators(weekly_df)
    master_df = build_master_df(daily_all)

    training_pool = build_training_pool(weekly_ma_df, master_df)
    otc_supplement = pd.DataFrame()
    otc_supplement = build_otc_bullish_supplement_pool(
        weekly_ma_df, master_df, existing_df=training_pool, min_count=40
    )
    if isinstance(otc_supplement, pd.DataFrame) and not otc_supplement.empty:
        training_pool = pd.concat([training_pool, otc_supplement], ignore_index=True)
        training_pool = training_pool.drop_duplicates(subset=["股票代號"], keep="first").reset_index(drop=True)

    bearish_pool = build_bearish_pool(weekly_ma_df, master_df)
    strict_df = build_strict_breakout_sheet(weekly_ma_df, master_df)
    bearish_key_df = build_bearish_key_breakdown_sheet(weekly_ma_df, master_df)

    bullish_display = build_client_bullish_view(training_pool)
    bearish_display = build_client_bearish_view(bearish_pool)
    bullish_keyk_display = build_client_bullish_keyk_view(strict_df)
    bearish_keyk_display = build_client_bearish_keyk_view(bearish_key_df)

    bullish_items = pool_to_api_items(bullish_display, score_col="StrongScore")
    bearish_items = pool_to_api_items(bearish_display, score_col="BearishScore")
    bullish_keyk_items = keyk_to_api_items(bullish_keyk_display)
    bearish_keyk_items = keyk_to_api_items(bearish_keyk_display)

    warrant_pool = filter_warrant_candidates(training_pool)
    warrant_items = pool_to_api_items(warrant_pool, score_col="StrongScore")
    warrants = build_warrants_from_bullish(warrant_items) if include_warrants else []

    settle_date = fmt_date_ymd(get_latest_available_trading_date())

    listed_n = 0
    otc_n = 0
    history_days = 0
    try:
        listed_n = int(daily_all.loc[daily_all["市場別"].astype(str).eq("上市"), "股票代號"].nunique())
        otc_n = int(daily_all.loc[daily_all["市場別"].astype(str).eq("上櫃"), "股票代號"].nunique())
        history_days = int(daily_all["日期"].nunique())
    except Exception:
        pass

    return {
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "settle_date": settle_date,
        "source": "web-strategy-weekly",
        "market": "上市+上櫃（週K｜趨勢突破守穩｜週量≥1萬）",
        "bullish": bullish_items,
        "bearish": bearish_items,
        "bullish_keyk": bullish_keyk_items,
        "bearish_keyk": bearish_keyk_items,
        "warrants": warrants,
        "bullish_count": len(bullish_items),
        "bearish_count": len(bearish_items),
        "bullish_keyk_count": len(bullish_keyk_items),
        "bearish_keyk_count": len(bearish_keyk_items),
        "warrant_count": len(warrants),
        "pool_count": len(training_pool),
        "training_pool_count": len(training_pool),
        "otc_supplement_count": len(otc_supplement) if isinstance(otc_supplement, pd.DataFrame) else 0,
        "warrant_candidate_count": len(warrant_pool),
        "_warrant_items": warrant_items,
        "data_stats": {
            "history_calendar_days": HISTORY_CALENDAR_DAYS,
            "history_trading_days": history_days,
            "listed_stocks": listed_n,
            "otc_stocks": otc_n,
            "daily_rows": len(daily_all),
        },
        "strategy": {
            "bullish": "CLIENT_BULLISH＝TRAINING_POOL（週量≥1萬、趨勢突破守穩、StrongScore≥55）",
            "bullish_keyk": "CLIENT_BULLISH_KEYK＝STRICT_BREAKOUT（本週正式突破趨勢線+盤整）",
            "bearish": "CLIENT_BEARISH＝BEARISH_TRAINING_POOL（週量≥1萬、跌破守穩、BearishScore≥55）",
            "bearish_keyk": "CLIENT_BEARISH_KEYK＝BEARISH_KEY_BREAKDOWN（本週正式跌破）",
            "warrant_min_score": WARRANT_MIN_SCORE,
            "warrant_min_stars": WARRANT_MIN_STARS,
            "min_weekly_volume": MIN_WEEKLY_VOLUME,
            "bias_limit": None,
            "warrant_days": f"{WARRANT_MIN_DAYS}-{WARRANT_MAX_DAYS}",
        },
    }
