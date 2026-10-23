"""Helpers to turn raw spreadsheet values into the display strings used in the deck."""
from __future__ import annotations

import math

import pandas as pd


def _is_blank(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    if isinstance(value, str) and not value.strip():
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return False


def format_text(value) -> str:
    return "" if _is_blank(value) else str(value).strip()


def format_date(value) -> str:
    if _is_blank(value):
        return ""
    if isinstance(value, str) and value.strip() and not value.strip()[0].isdigit():
        return value.strip()
    try:
        dt = pd.to_datetime(value)
    except (ValueError, TypeError):
        return str(value).strip()
    if isinstance(dt, pd.Timestamp):
        return dt.strftime("%b %Y")
    return str(value).strip()


def format_percent(value) -> str:
    if _is_blank(value):
        return ""
    if isinstance(value, str) and "%" in value:
        return value.strip()
    try:
        num = float(str(value).replace("%", ""))
    except ValueError:
        return str(value).strip()
    if abs(num) <= 1:
        num *= 100
    return f"{num:.0f}%"
