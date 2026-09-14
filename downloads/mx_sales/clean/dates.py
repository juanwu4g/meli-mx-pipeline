"""Date parsing for the Spanish long-form timestamps used across ML reports.

Three shapes occur in the same workbook::

    "3 de agosto de 2026 09:24 hs."   full date and time
    "3 de agosto | 23:51"             day, month and time, no year
    "11 de agosto"                    day and month only

The year-less variants are resolved against a per-row reference timestamp (the sale
date), which keeps December/January order correct across a year boundary.
"""

from __future__ import annotations

import re
from datetime import datetime

import pandas as pd

from mx_sales.clean.text import blank_to_na

MONTHS = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}

_PATTERN = re.compile(
    r"^(?P<day>\d{1,2})\s+de\s+(?P<month>[a-zñáéíóú]+)"
    r"(?:\s+de\s+(?P<year>\d{4}))?"
    r"(?:\s*[|,]?\s*(?P<hour>\d{1,2}):(?P<minute>\d{2}))?",
    re.IGNORECASE,
)


def parse_spanish_datetime(value: object, reference: pd.Timestamp | None = None) -> pd.Timestamp:
    """Parse one Spanish long-form date, inferring the year from ``reference`` if absent."""
    if value is None or value is pd.NA or (isinstance(value, float) and pd.isna(value)):
        return pd.NaT
    if isinstance(value, (datetime, pd.Timestamp)):
        return pd.Timestamp(value)

    match = _PATTERN.match(str(value).strip().lower().replace(" ", " "))
    if match is None:
        return pd.NaT

    month = MONTHS.get(match.group("month"))
    if month is None:
        return pd.NaT

    day = int(match.group("day"))
    hour = int(match.group("hour") or 0)
    minute = int(match.group("minute") or 0)

    if match.group("year"):
        return _build(int(match.group("year")), month, day, hour, minute)

    return _infer_year(month, day, hour, minute, reference)


def _build(year: int, month: int, day: int, hour: int, minute: int) -> pd.Timestamp:
    try:
        return pd.Timestamp(year=year, month=month, day=day, hour=hour, minute=minute)
    except ValueError:  # e.g. 29 February on a non-leap year
        return pd.NaT


def _infer_year(
    month: int, day: int, hour: int, minute: int, reference: pd.Timestamp | None
) -> pd.Timestamp:
    """Pick the year that places the date nearest to the reference.

    Nearest wins outright. Preferring a candidate at or after the sale looks reasonable
    -- deliveries follow sales -- but it breaks on exchange rows, where the shipment ML
    lists predates the replacement sale: "11 de agosto" against a 14 August sale would
    be pushed to the following year. Distance alone gets both that and the December to
    January rollover right.
    """
    if reference is None or pd.isna(reference):
        return pd.NaT

    reference = pd.Timestamp(reference)
    candidates = [
        ts
        for year in (reference.year - 1, reference.year, reference.year + 1)
        if not pd.isna(ts := _build(year, month, day, hour, minute))
    ]
    if not candidates:
        return pd.NaT

    return min(candidates, key=lambda ts: (abs(ts - reference), ts < reference))


def to_datetime(
    series: pd.Series, reference: pd.Series | None = None, *, spanish: bool = True
) -> pd.Series:
    """Vectorized wrapper: parse a column of dates, optionally against a reference column."""
    if not spanish:
        return pd.to_datetime(blank_to_na(series), errors="coerce", format="mixed")

    if reference is None:
        parsed = [parse_spanish_datetime(v) for v in series]
    else:
        parsed = [parse_spanish_datetime(v, r) for v, r in zip(series, reference)]
    return pd.Series(parsed, index=series.index, dtype="datetime64[ns]")
