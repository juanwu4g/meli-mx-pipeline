"""Excel readers that cope with ML's decorated worksheets.

ML workbooks put a title block, an "updated at" stamp and help links above the real
table, so the header row sits at a different offset in almost every report. Several
reports also use a two-level header where the upper row groups columns and the lower
row names them -- and the lower names repeat (``Unidades`` appears under Ventas,
Devoluciones and Reclamos), so the group is what disambiguates them.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from mx_sales.clean.text import slugify


def _row_filled(row: pd.Series) -> int:
    return int(row.notna().sum())


def find_header_row(frame: pd.DataFrame, *, min_columns: int = 3, search_rows: int = 20) -> int:
    """Return the index of the densest row in the top of the sheet: the header.

    Title and note rows occupy one or two cells; the header spans the full table, so
    the first row reaching ``min_columns`` and holding the maximum fill is the header.
    """
    window = frame.head(search_rows)
    counts = [(_row_filled(row), -idx) for idx, row in window.iterrows()]
    best_count, negative_index = max(counts)
    if best_count < min_columns:
        raise ValueError(f"no header row with >= {min_columns} populated cells in first {search_rows} rows")
    return -negative_index


def _drop_empty_columns(frame: pd.DataFrame) -> pd.DataFrame:
    """Drop fully empty columns (several reports start at column B).

    Only columns are trimmed here: empty rows above the table are load-bearing,
    because callers address the group and header rows by position.
    """
    trimmed = frame.dropna(axis=1, how="all").reset_index(drop=True)
    trimmed.columns = range(trimmed.shape[1])
    return trimmed


def read_sheet(
    path: Path | str,
    sheet_name: str | int = 0,
    *,
    header_row: int | None = None,
    trim_empty: bool = True,
) -> pd.DataFrame:
    """Read one sheet with a single header row, auto-detecting the offset when unset."""
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
    if trim_empty:
        raw = _drop_empty_columns(raw)

    if header_row is None:
        header_row = find_header_row(raw)

    columns = [slugify(v) for v in raw.iloc[header_row]]
    data = raw.iloc[header_row + 1 :].reset_index(drop=True)
    data.columns = _deduplicate(columns)
    return data.dropna(axis=0, how="all").reset_index(drop=True)


def read_grouped_sheet(
    path: Path | str,
    sheet_name: str | int = 0,
    *,
    group_row: int,
    header_row: int,
    trim_empty: bool = True,
) -> pd.DataFrame:
    """Read a sheet whose header spans two rows, naming columns ``<group>__<header>``.

    The group row is forward-filled because ML writes each group label only above its
    first column.
    """
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
    if trim_empty:
        raw = _drop_empty_columns(raw)

    groups = raw.iloc[group_row].ffill()
    headers = raw.iloc[header_row]
    columns = [
        f"{slugify(g)}__{slugify(h)}" if pd.notna(g) and slugify(g) else slugify(h)
        for g, h in zip(groups, headers)
    ]

    data = raw.iloc[header_row + 1 :].reset_index(drop=True)
    data.columns = _deduplicate(columns)
    return data.dropna(axis=0, how="all").reset_index(drop=True)


def _deduplicate(columns: list[str]) -> list[str]:
    """Suffix any remaining duplicate names so the frame keeps unique columns."""
    seen: dict[str, int] = {}
    result: list[str] = []
    for name in columns:
        name = name or "unnamed"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        result.append(name)
    return result
