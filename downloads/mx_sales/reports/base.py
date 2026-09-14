"""The contract every report type implements.

A report type answers three questions: which files are mine, how do I become a frame,
and what does a clean row look like. Adding a new ML report means adding one module
with a ``ReportSpec`` and registering it -- the pipeline itself never changes.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol

import pandas as pd


class Loader(Protocol):
    """Reads a raw file into an untyped frame with slugified column names."""

    def __call__(self, path: Path) -> pd.DataFrame: ...


class Transformer(Protocol):
    """Cleans and types a raw frame, returning analysis-ready rows."""

    def __call__(self, frame: pd.DataFrame) -> pd.DataFrame: ...


@dataclass(frozen=True)
class ReportSpec:
    """Everything the pipeline needs to know about one ML report type."""

    name: str
    """Table name used for the Parquet dataset and the DuckDB view."""

    filename_pattern: str
    """Regex matched against the file name to claim a download."""

    loader: Loader
    transformer: Transformer
    description: str = ""
    primary_key: tuple[str, ...] = ()
    partition_by: tuple[str, ...] = field(default=())

    def matches(self, path: Path) -> bool:
        return re.search(self.filename_pattern, path.name, re.IGNORECASE) is not None

    def load(self, path: Path) -> pd.DataFrame:
        return self.loader(path)

    def transform(self, frame: pd.DataFrame) -> pd.DataFrame:
        return self.transformer(frame)


#: Seller account id as it appears at the end of an export file name.
SELLER_ID_PATTERN = re.compile(r"_(\d{6,})\.[a-z]+$", re.IGNORECASE)


def add_lineage(frame: pd.DataFrame, path: Path, *, raw_dir: Path) -> pd.DataFrame:
    """Stamp each row with where it came from, so any value can be traced back.

    ``store`` and ``seller_id`` matter as soon as more than one account is downloaded
    into the same tree: order ids are unique per account, not across accounts.
    """
    try:
        relative = path.relative_to(raw_dir)
        source = relative.as_posix()
        store = relative.parts[0] if len(relative.parts) > 1 else ""
    except ValueError:
        source = path.as_posix()
        store = ""

    match = SELLER_ID_PATTERN.search(path.name)

    frame = frame.copy()
    frame["store"] = store
    frame["seller_id"] = match.group(1) if match else pd.NA
    frame["source_file"] = source
    frame["source_modified_at"] = pd.Timestamp(path.stat().st_mtime, unit="s")
    frame["ingested_at"] = pd.Timestamp.now().floor("s")
    return frame


def require_columns(frame: pd.DataFrame, columns: list[str], *, report: str) -> None:
    """Fail loudly when an export changes shape, instead of silently emitting nulls."""
    missing = [c for c in columns if c not in frame.columns]
    if missing:
        raise ValueError(f"{report}: expected columns missing from export: {missing}")
