"""String normalization.

MercadoLibre exports use a single space (not an empty cell) for "no value", which
makes every column look fully populated until it is normalized away.
"""

from __future__ import annotations

import re
import unicodedata

import pandas as pd

_WHITESPACE = re.compile(r"\s+")
_NON_ALNUM = re.compile(r"[^a-z0-9]+")


def collapse_whitespace(series: pd.Series) -> pd.Series:
    """Trim and collapse internal runs of whitespace, including non-breaking spaces."""
    return (
        series.astype("string")
        .str.replace(" ", " ", regex=False)
        .str.replace(_WHITESPACE, " ", regex=True)
        .str.strip()
    )


def blank_to_na(series: pd.Series) -> pd.Series:
    """Normalize whitespace and turn blank / placeholder values into ``pd.NA``."""
    cleaned = collapse_whitespace(series)
    return cleaned.mask(cleaned.isin(["", "-", "--", "N/A", "n/a", "null", "None"]))


def slugify(value: object, sep: str = "_") -> str:
    """Lowercase ASCII slug, e.g. ``"Ingresos por envío (MXN)"`` -> ``ingresos_por_envio_mxn``."""
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return _NON_ALNUM.sub(sep, text.lower()).strip(sep)


def split_code_label(series: pd.Series, pattern: str) -> tuple[pd.Series, pd.Series]:
    """Split a combined ``"<code> <label>"`` column into two series.

    ``pattern`` must define exactly two capture groups: the code and the label.
    Rows that do not match yield ``pd.NA`` for both outputs.
    """
    extracted = blank_to_na(series).str.extract(pattern, expand=True)
    code = extracted[0].str.strip()
    label = extracted[1].str.strip().str.rstrip(".")
    return code, label


def strip_prefix(series: pd.Series, prefix: str) -> pd.Series:
    """Drop a leading literal prefix such as ``"RFC: "`` when present."""
    cleaned = blank_to_na(series)
    return cleaned.str.replace(rf"^\s*{re.escape(prefix)}\s*", "", regex=True).str.strip()
