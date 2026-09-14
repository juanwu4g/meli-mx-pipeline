"""Boolean coercion for the Sí/No and true/false flags used across reports."""

from __future__ import annotations

import pandas as pd

from mx_sales.clean.text import blank_to_na

TRUE_VALUES = {"si", "sí", "s", "yes", "y", "true", "1"}
FALSE_VALUES = {"no", "n", "false", "0", "no aplica"}


def to_boolean(series: pd.Series) -> pd.Series:
    """Map Sí/No style values to a nullable ``boolean`` column; unknowns become ``pd.NA``."""
    if pd.api.types.is_bool_dtype(series):
        return series.astype("boolean")

    normalized = blank_to_na(series).str.lower()
    return normalized.map(
        lambda v: True if v in TRUE_VALUES else False if v in FALSE_VALUES else pd.NA
    ).astype("boolean")
