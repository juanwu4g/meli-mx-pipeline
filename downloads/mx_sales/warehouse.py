"""DuckDB access over the Parquet outputs.

The Parquet files stay the source of truth; DuckDB just puts a queryable view on top,
so the warehouse can be deleted and rebuilt at any time without losing data.
"""

from __future__ import annotations

import duckdb
import pandas as pd

from mx_sales.config import SETTINGS, Settings
from mx_sales.registry import SPECS


def connect(settings: Settings = SETTINGS, *, read_only: bool = False) -> duckdb.DuckDBPyConnection:
    """Open the warehouse, creating its parent directory on first use."""
    settings.warehouse_path.parent.mkdir(parents=True, exist_ok=True)
    return duckdb.connect(str(settings.warehouse_path), read_only=read_only)


def refresh_views(settings: Settings = SETTINGS) -> list[str]:
    """(Re)create one view per report that has Parquet output on disk."""
    created: list[str] = []
    with connect(settings) as con:
        for spec in SPECS:
            location = settings.processed_dir / spec.name
            if not any(location.rglob("*.parquet")):
                continue
            pattern = (location / "**" / "*.parquet").as_posix()
            con.execute(
                f"CREATE OR REPLACE VIEW {spec.name} AS "
                f"SELECT * FROM read_parquet('{pattern}', hive_partitioning = true)"
            )
            created.append(spec.name)
    return created


def query(sql: str, settings: Settings = SETTINGS) -> pd.DataFrame:
    """Run a query against the warehouse and return the rows as a DataFrame."""
    with connect(settings, read_only=True) as con:
        return con.execute(sql).fetch_df()


def read(report: str, settings: Settings = SETTINGS) -> pd.DataFrame:
    """Read one report's Parquet output straight into pandas, bypassing DuckDB.

    The directory is read as a single dataset so hive partition columns (``sale_month``)
    come back as real columns rather than being lost with the directory names.
    """
    location = settings.processed_dir / report
    if not any(location.rglob("*.parquet")):
        raise FileNotFoundError(f"no Parquet output for {report!r} under {location}")
    return pd.read_parquet(location, engine="pyarrow")


def table_summary(settings: Settings = SETTINGS) -> pd.DataFrame:
    """List the report tables currently available in the warehouse, with row counts."""
    rows = []
    with connect(settings, read_only=True) as con:
        for (name,) in con.execute("SELECT view_name FROM duckdb_views()").fetchall():
            count = con.execute(f"SELECT count(*) FROM {name}").fetchone()[0]
            rows.append({"table": name, "rows": count})
    return pd.DataFrame(rows)
