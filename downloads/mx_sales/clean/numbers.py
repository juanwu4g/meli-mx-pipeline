"""Numeric coercion for values that arrive as formatted strings.

CSV exports render money as ``"31,350.43"`` (comma thousands, dot decimal) and use a
leading minus for debits; XLSX exports usually arrive already typed.
"""

from __future__ import annotations

import pandas as pd

from mx_sales.clean.text import blank_to_na


def to_numeric(series: pd.Series, *, dtype: str = "Float64") -> pd.Series:
    """Coerce a column to a nullable numeric dtype, stripping currency formatting."""
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(dtype)

    cleaned = (
        blank_to_na(series)
        .str.replace(r"[^\d,.\-]", "", regex=True)
        .str.replace(",", "", regex=False)
    )
    return pd.to_numeric(cleaned, errors="coerce").astype(dtype)
