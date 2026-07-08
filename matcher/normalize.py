"""Normalization helpers for GST invoice matching."""

from __future__ import annotations

import re
from datetime import date, datetime
from typing import Any

import pandas as pd


def clean_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def normalize_gstin(value: Any) -> str:
    text = clean_text(value).upper().replace(" ", "")
    return re.sub(r"[^0-9A-Z]", "", text)


def normalize_invoice_no(value: Any) -> str:
    text = clean_text(value).upper()
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[^A-Z0-9/\-_.]", "", text)
    return text.lstrip("0") or text or ""


def normalize_date(value: Any) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value

    text = clean_text(value)
    if re.match(r"^\d{4}-\d{2}-\d{2}", text):
        parsed = pd.to_datetime(value, dayfirst=False, errors="coerce")
    else:
        parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")

    if pd.isna(parsed):
        return None
    return parsed.date()


def normalize_amount(value: Any) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, str):
        value = value.replace(",", "").strip()
        if not value:
            return 0.0
    try:
        return round(float(value), 2)
    except (TypeError, ValueError):
        return 0.0


def amounts_equal(a: float, b: float, tolerance: float = 1.0) -> bool:
    return abs(a - b) <= tolerance
