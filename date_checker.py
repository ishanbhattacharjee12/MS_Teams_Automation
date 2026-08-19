from datetime import datetime, timedelta
import pandas as pd
try:
    from zoneinfo import ZoneInfo
except ImportError:
    pass # Assume Python 3.9+
import config

def parse_date(date_val):
    if pd.isna(date_val) or str(date_val).strip() == "":
        return None
    try:
        dt = pd.to_datetime(date_val).to_pydatetime()
        return dt.date()
    except Exception:
        return None

import sys

def get_today():
    try:
        tz = ZoneInfo(config.TIMEZONE)
    except Exception as e:
        print(f"[ERROR] Invalid TIMEZONE configured: '{config.TIMEZONE}'. Error: {e}")
        sys.exit(1)
    return datetime.now(tz).date()

def is_in_reminder_window(target_date, days_before):
    if target_date is None:
        return False
    
    today = get_today()
    window_start = target_date - timedelta(days=days_before)
    
    return window_start <= today <= target_date
