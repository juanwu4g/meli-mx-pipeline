"""Discover raw downloads, clean them, and write Parquet.

The pipeline is deliberately thin: it knows how to walk the raw tree and how to write
output, and delegates everything report-specific to the registered ``ReportSpec``.
Re-running it is safe -- output is rewritten from the raw files each time.
"""

from __future__ import annotations

import json
import logging
import shutil
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from mx_sales.config import EXCLUDED_DIRS, SETTINGS, Settings
from mx_sales.registry import SPECS, get, spec_for
from mx_sales.reports.base import ReportSpec, add_lineage

logger = logging.getLogger(__name__)

RAW_SUFFIXES = {".xlsx", ".xls", ".csv"}


@dataclass(frozen=True)
class Discovered:
    """A raw file matched to the report type that knows how to read it."""

    path: Path
    spec: ReportSpec


@dataclass
class RunResult:
    """What one pipeline run produced, per report type."""

    report: str
    files: int = 0
    rows: int = 0
    output: Path | None = None
    errors: list[str] = None  # type: ignore[assignment]

    def __post_init__(self) -> None:
        if self.errors is None:
            self.errors = []


def discover(settings: Settings = SETTINGS) -> list[Discovered]:
    """Walk the raw tree and match every recognized download to its spec."""
    found: list[Discovered] = []
    for path in sorted(settings.raw_dir.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in RAW_SUFFIXES:
            continue
        if path.name.startswith("~$") or EXCLUDED_DIRS & set(path.relative_to(settings.raw_dir).parts):
            continue
        spec = spec_for(path)
        if spec is not None:
            found.append(Discovered(path=path, spec=spec))
    return found


def matches_file(path: Path, only: Path | str, settings: Settings) -> bool:
    """Match ``only`` against a discovered path by full path, relative path or name."""
    candidate = Path(only)
    if candidate.is_absolute():
        return path == candidate.resolve()

    text = str(only).replace("\\", "/")
    relative = path.relative_to(settings.raw_dir).as_posix()
    return relative == text or path.name == candidate.name or path.stem == candidate.stem


def process_file(item: Discovered, settings: Settings = SETTINGS) -> pd.DataFrame:
    """Load, clean and stamp one raw file."""
    raw = item.spec.load(item.path)
    cleaned = item.spec.transform(raw)
    return add_lineage(cleaned, item.path, raw_dir=settings.raw_dir)


def run(
    report: str | None = None,
    settings: Settings = SETTINGS,
    *,
    write: bool = True,
    only: Path | str | None = None,
) -> list[RunResult]:
    """Process the discovered files, optionally limited to one report type or one file.

    ``only`` restricts the run to a single download, which is the normal way to work: a
    snapshot tree holds the same report many times over, and processing one export is
    both faster and easier to check against the workbook in front of you.

    Returns one :class:`RunResult` per report type that had matching files.
    """
    wanted = {spec.name for spec in SPECS} if report is None else {report}
    items = [item for item in discover(settings) if item.spec.name in wanted]

    if only is not None:
        items = [item for item in items if matches_file(item.path, only, settings)]
        if not items:
            raise FileNotFoundError(f"no registered report matched {only!r}")

    results: dict[str, RunResult] = {}
    frames: dict[str, list[pd.DataFrame]] = {}

    for item in items:
        result = results.setdefault(item.spec.name, RunResult(report=item.spec.name))
        try:
            frame = process_file(item, settings)
        except Exception as exc:  # one bad export must not sink the run
            logger.warning("failed to process %s: %s", item.path.name, exc)
            result.errors.append(f"{item.path.name}: {exc}")
            continue
        frames.setdefault(item.spec.name, []).append(frame)
        result.files += 1
        result.rows += len(frame)

    for name, parts in frames.items():
        spec = get(name)
        combined = deduplicate(pd.concat(parts, ignore_index=True), spec)
        results[name].rows = len(combined)
        if write:
            results[name].output = write_parquet(combined, spec, settings)

    return list(results.values())


def deduplicate(frame: pd.DataFrame, spec: ReportSpec) -> pd.DataFrame:
    """Collapse the same record appearing in several snapshots down to its newest copy.

    The download pipeline re-fetches overlapping date ranges, so one order shows up in
    every snapshot taken after it. Without a key there is nothing safe to collapse on,
    so the frame is returned untouched.
    """
    if not spec.primary_key or not set(spec.primary_key) <= set(frame.columns):
        return frame

    ordered = frame.sort_values("source_modified_at", kind="stable")
    kept = ordered.drop_duplicates(subset=list(spec.primary_key), keep="last")
    return kept.sort_index().reset_index(drop=True)


def write_parquet(frame: pd.DataFrame, spec: ReportSpec, settings: Settings = SETTINGS) -> Path:
    """Write one report's rows to ``data/processed/<report>/``, partitioned if configured.

    The target is cleared first: Parquet writes add files rather than replacing them, so
    without this a changed partition layout or a shrunken source would leave stale rows
    behind and silently double-count.
    """
    target = settings.processed_dir / spec.name
    if target.exists():
        shutil.rmtree(target)
    target.mkdir(parents=True, exist_ok=True)

    prepared = _prepare_for_parquet(frame)
    partitions = [c for c in spec.partition_by if c in prepared.columns]

    if partitions:
        prepared.to_parquet(target, engine="pyarrow", index=False, partition_cols=partitions)
    else:
        prepared.to_parquet(target / f"{spec.name}.parquet", engine="pyarrow", index=False)
    return target


def _prepare_for_parquet(frame: pd.DataFrame) -> pd.DataFrame:
    """Make columns Arrow-writable.

    Dict-valued columns (parsed variant attributes) and all-null object columns have no
    stable Arrow type, so they are serialized as JSON text.
    """
    out = frame.copy()
    for column in out.columns:
        series = out[column]
        if series.dtype != object:
            continue
        non_null = series.dropna()
        if not non_null.empty and isinstance(non_null.iloc[0], dict):
            out[column] = series.map(
                lambda v: json.dumps(v, ensure_ascii=False) if isinstance(v, dict) else None
            ).astype("string")
        else:
            out[column] = series.astype("string")
    return out
